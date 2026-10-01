"""
GYN Gestão — visita de gestão em enfermaria de Ginecologia Cirúrgica.

Versão 2:
- registro por visita (append-only), mas painel baseado no CENSO ATUAL
  (última visita de cada paciente ainda internada);
- DI, DPO e atrasos recalculados a cada abertura (não ficam "congelados");
- pré-preenchimento a partir da última visita da paciente;
- campos condicionais funcionando (sem st.form);
- registro de alta e indicadores de internações encerradas;
- gravidade das complicações pela classificação de Clavien-Dindo;
- fuso horário de São Paulo (o servidor do Streamlit Cloud roda em UTC);
- login individual por usuário (seção [usuarios] dos secrets, senha com hash);
- interface pensada para celular (cartões, cadastro recolhido, campos empilhados).
"""

import hashlib
import hmac
import io
import os
from datetime import date, datetime, timedelta
from zoneinfo import ZoneInfo

import pandas as pd
import streamlit as st

st.set_page_config(page_title="GYN Gestão", page_icon="🏥", layout="wide")

TZ = ZoneInfo("America/Sao_Paulo")


def hoje() -> date:
    return datetime.now(TZ).date()


def agora_str() -> str:
    return datetime.now(TZ).strftime("%Y-%m-%d %H:%M:%S")


# ---------------------------------------------------------------------------
# Esquema de dados
# ---------------------------------------------------------------------------
# Campos da v1 (ordem preservada para não desalinhar a planilha existente)
FIELDS_V1 = [
    "timestamp", "preceptor", "unidade", "leito", "id_paciente",
    "data_internacao", "dia_internacao",
    "cirurgia_status", "procedimento", "data_cirurgia", "dpo",
    "dias_previstos_internacao", "data_prevista_alta",
    "complicacao", "tipo_complicacao", "complicacao_impede_alta",
    "especialista_necessario", "especialidade", "parecer_solicitado",
    "parecer_realizado", "parecer_atrasa_alta",
    "demanda_enfermagem", "tipo_demanda_enfermagem",
    "pendencia_principal", "responsavel_pendencia", "prazo_pendencia",
    "status_gestao", "barreira_alta",
    "meta_1", "meta_2", "meta_3",
    "alta_previsao", "observacoes",
]
# Campos novos: sempre acrescentados AO FINAL
FIELDS_V2 = ["clavien_dindo", "alta_realizada", "data_alta", "usuario"]
FIELDS = FIELDS_V1 + FIELDS_V2

DATE_FIELDS = [
    "data_internacao", "data_cirurgia", "data_prevista_alta",
    "prazo_pendencia", "data_alta",
]

STATUS = [
    "Verde — dentro do previsto",
    "Amarelo — risco de atraso",
    "Vermelho — atraso/complicação",
]
STATUS_ICON = {"Verde": "🟢", "Amarelo": "🟡", "Vermelho": "🔴"}

TIPOS_COMPLICACAO = [
    "Hemorrágica", "Infecciosa", "Urinária", "Gastrointestinal/íleo",
    "Tromboembólica", "Respiratória", "Cardiovascular",
    "Ferida operatória", "Dor persistente", "Outra",
]
CLAVIEN = ["I", "II", "IIIa", "IIIb", "IVa", "IVb", "V"]
CLAVIEN_HELP = (
    "Clavien-Dindo (Dindo et al., Ann Surg 2004):\n\n"
    "I — desvio do curso pós-operatório sem necessidade de tratamento "
    "farmacológico (exceto antieméticos, antipiréticos, analgésicos, "
    "diuréticos, eletrólitos e fisioterapia), cirúrgico, endoscópico ou "
    "radiológico.\n\n"
    "II — necessita tratamento farmacológico além do grau I, transfusão "
    "ou nutrição parenteral.\n\n"
    "IIIa — intervenção cirúrgica, endoscópica ou radiológica sem anestesia "
    "geral. IIIb — sob anestesia geral.\n\n"
    "IVa — disfunção de um órgão (inclui UTI). IVb — disfunção de "
    "múltiplos órgãos.\n\n"
    "V — óbito."
)
ESPECIALIDADES = [
    "Clínica Médica", "Cardiologia", "Anestesiologia", "Urologia",
    "Cirurgia Geral", "Infectologia", "Hematologia", "Nefrologia",
    "Pneumologia", "Nutrição", "Fisioterapia",
    "Psicologia/Psiquiatria", "Serviço Social", "Outro",
]
DEMANDAS_ENF = [
    "Dor", "Náusea/vômito", "Acesso venoso", "SVD", "Dreno",
    "Ferida operatória", "Sangramento vaginal", "Mobilização",
    "Dieta", "Eliminações", "Curativo", "Medicação",
    "Educação para alta", "Acompanhante/cuidador", "Outra",
]
SIM_NAO = ["Não", "Sim"]

LOCAL_FILE = "gestao_ginecologia.csv"

# ---------------------------------------------------------------------------
# Persistência: Google Sheets (se configurado) ou CSV local
# ---------------------------------------------------------------------------
USE_GOOGLE_SHEETS = False
try:
    import gspread
    from google.oauth2.service_account import Credentials

    if "gcp_service_account" in st.secrets and "google_sheet_id" in st.secrets:
        USE_GOOGLE_SHEETS = True
except Exception:
    USE_GOOGLE_SHEETS = False


@st.cache_resource
def get_sheet():
    creds = Credentials.from_service_account_info(
        dict(st.secrets["gcp_service_account"]),
        # Escopo mínimo: só planilhas (o acesso ao Drive não é necessário)
        scopes=["https://www.googleapis.com/auth/spreadsheets"],
    )
    sh = gspread.authorize(creds).open_by_key(st.secrets["google_sheet_id"])
    try:
        ws = sh.worksheet("visitas")
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title="visitas", rows=2000, cols=len(FIELDS) + 5)

    # Garante/migra o cabeçalho (acrescenta colunas novas ao final)
    header = ws.row_values(1)
    if ws.col_count < len(FIELDS):
        ws.add_cols(len(FIELDS) - ws.col_count)
    if not header or (header != FIELDS and FIELDS[: len(header)] == header):
        ws.update(range_name="A1", values=[FIELDS])
    return ws


@st.cache_data(ttl=60, show_spinner="Carregando registros…")
def _load_raw() -> pd.DataFrame:
    if USE_GOOGLE_SHEETS:
        # numericise_ignore evita perder zeros à esquerda em IDs/leitos
        values = get_sheet().get_all_records(numericise_ignore=["all"])
        df = pd.DataFrame(values)
    elif os.path.exists(LOCAL_FILE):
        df = pd.read_csv(LOCAL_FILE, dtype=str, keep_default_na=False)
    else:
        df = pd.DataFrame(columns=FIELDS)
    return df


def save_record(rec: dict) -> None:
    row = [str(rec.get(c, "")) for c in FIELDS]
    if USE_GOOGLE_SHEETS:
        # RAW: grava exatamente o texto (USER_ENTERED convertia datas para o
        # formato da planilha e IDs numéricos perdiam zeros à esquerda)
        get_sheet().append_row(row, value_input_option="RAW")
    else:
        df = _load_raw()
        for c in FIELDS:
            if c not in df.columns:
                df[c] = ""
        df = pd.concat([df[FIELDS], pd.DataFrame([row], columns=FIELDS)],
                       ignore_index=True)
        df.to_csv(LOCAL_FILE, index=False, encoding="utf-8-sig")
    _load_raw.clear()


# ---------------------------------------------------------------------------
# Preparação e campos derivados
# ---------------------------------------------------------------------------
def _parse_dates(s: pd.Series) -> pd.Series:
    s = s.astype(str).str.strip().str.slice(0, 10)
    iso = pd.to_datetime(s, format="%Y-%m-%d", errors="coerce")
    br = pd.to_datetime(s, format="%d/%m/%Y", errors="coerce")
    return iso.fillna(br)


def _parse_ts(s: pd.Series) -> pd.Series:
    s = s.astype(str).str.strip().str.replace("T", " ", regex=False).str.slice(0, 19)
    iso = pd.to_datetime(s, format="%Y-%m-%d %H:%M:%S", errors="coerce")
    br = pd.to_datetime(s, format="%d/%m/%Y %H:%M:%S", errors="coerce")
    return iso.fillna(br)


def status_cor(s: str) -> str:
    s = str(s)
    for cor in ("Vermelho", "Amarelo", "Verde"):
        if s.startswith(cor):
            return cor
    return ""


def load_data() -> pd.DataFrame:
    df = _load_raw().copy()
    for c in FIELDS:
        if c not in df.columns:
            df[c] = ""
    df = df[FIELDS].fillna("").astype(str)
    df["id_paciente"] = df["id_paciente"].str.strip()
    df["_ts"] = _parse_ts(df["timestamp"])
    for c in DATE_FIELDS:
        df["_" + c] = _parse_dates(df[c])

    t = pd.Timestamp(hoje())
    alta = df["alta_realizada"].eq("Sim")
    fim = df["_data_alta"].where(alta & df["_data_alta"].notna(), t)

    # Recalculados sempre em relação a hoje (ou à data da alta)
    df["DI"] = (fim - df["_data_internacao"]).dt.days + 1
    realizada = df["cirurgia_status"].eq("Realizada")
    df["DPO"] = (fim - df["_data_cirurgia"]).dt.days.where(realizada)

    previstos = pd.to_numeric(df["dias_previstos_internacao"], errors="coerce")
    atraso_edd = (t - df["_data_prevista_alta"]).dt.days
    df["dias_alem_previsto"] = atraso_edd.where(~alta).clip(lower=0)
    df["alta_atrasada"] = (~alta) & (
        (atraso_edd > 0)
        | (df["_data_prevista_alta"].isna() & (df["DI"] > previstos))
    )
    df["pendencia_vencida"] = (
        (~alta)
        & df["pendencia_principal"].str.strip().ne("")
        & (df["_prazo_pendencia"] < t)
    )
    df["cor"] = df["status_gestao"].map(status_cor)
    df["ordem_cor"] = df["cor"].map({"Vermelho": 0, "Amarelo": 1, "Verde": 2}).fillna(3)
    return df


def censo(df: pd.DataFrame) -> pd.DataFrame:
    """Última visita de cada paciente; ativas = sem alta registrada."""
    d = df[df["id_paciente"].ne("")].sort_values("_ts", kind="stable")
    if d.empty:
        return d.assign(n_visitas=pd.Series(dtype=int), visitada_hoje=pd.Series(dtype=bool))
    n = d.groupby("id_paciente").size().rename("n_visitas")
    last = d.groupby("id_paciente").tail(1).join(n, on="id_paciente")
    last["visitada_hoje"] = last["_ts"].dt.date.eq(hoje())
    return last[last["alta_realizada"].ne("Sim")].sort_values(["ordem_cor", "leito"])


def status_sugerido(r: dict) -> str:
    """Regra explícita, inspirada na lógica 'Red2Green' (dia que agrega ou não
    valor ao plano de alta). O preceptor pode sobrescrever."""
    vermelho = (
        r["complicacao_impede_alta"] == "Sim"
        or r["parecer_atrasa_alta"] == "Sim"
        or r["data_prevista_alta"] < hoje()
    )
    amarelo = (
        r["complicacao"] == "Sim"
        or (r["parecer_solicitado"] == "Sim" and r["parecer_realizado"] == "Não")
        or bool(r["barreira_alta"].strip())
        or (r["prazo_pendencia"] is not None and r["prazo_pendencia"] < hoje())
    )
    return STATUS[2] if vermelho else STATUS[1] if amarelo else STATUS[0]


def previsao_relativa(edd: date) -> str:
    d = (edd - hoje()).days
    if d < 0:
        return "Atrasada"
    return {0: "Hoje", 1: "Amanhã", 2: "48 horas"}.get(d, ">48 horas")




# ---------------------------------------------------------------------------
# Autenticação individual
# ---------------------------------------------------------------------------
PBKDF2_ITER = 200_000


def get_secret(key, default=None):
    try:
        return st.secrets.get(key, default)
    except Exception:
        return default


def hash_senha(senha: str) -> str:
    salt = os.urandom(16)
    h = hashlib.pbkdf2_hmac("sha256", senha.encode(), salt, PBKDF2_ITER)
    return f"pbkdf2${PBKDF2_ITER}${salt.hex()}${h.hex()}"


def confere_senha(senha: str, armazenada: str) -> bool:
    armazenada = str(armazenada)
    if armazenada.startswith("pbkdf2$"):
        try:
            _, it, salt, h = armazenada.split("$")
            calc = hashlib.pbkdf2_hmac("sha256", senha.encode(), bytes.fromhex(salt), int(it))
            return hmac.compare_digest(calc.hex(), h)
        except Exception:
            return False
    return hmac.compare_digest(senha.encode(), armazenada.encode())


def carregar_usuarios() -> dict:
    """Lê a tabela [usuarios] dos secrets: login -> nome, perfil, senha."""
    bruto = get_secret("usuarios") or {}
    out = {}
    for login, d in dict(bruto).items():
        d = dict(d)
        if d.get("senha") and d.get("perfil") in ("Preceptor", "Chefia"):
            out[str(login).strip().lower()] = {
                "nome": d.get("nome", login), "perfil": d["perfil"], "senha": d["senha"],
            }
    return out


USUARIOS = carregar_usuarios()
MODO_DEMO = not USUARIOS
if MODO_DEMO and not USE_GOOGLE_SHEETS:
    USUARIOS = {
        "preceptor": {"nome": "Preceptor Demo", "perfil": "Preceptor", "senha": "demo123"},
        "chefia": {"nome": "Chefia Demo", "perfil": "Chefia", "senha": "chefia123"},
    }

ss = st.session_state
ss.setdefault("usuario", None)
ss.setdefault("ver", 0)
ss.setdefault("falhas", 0)
ss.setdefault("bloqueio_ate", None)


def tela_login():
    st.markdown("## 🏥 GYN Gestão")
    st.caption("Enfermaria de Ginecologia Cirúrgica — visita de gestão")
    if MODO_DEMO and USE_GOOGLE_SHEETS:
        st.error("Nenhum usuário configurado. Cadastre os usuários na seção "
                 "`[usuarios]` dos *secrets* do app (veja o README).")
        st.stop()
    if MODO_DEMO:
        st.info("Modo demonstração: usuário `preceptor` / senha `demo123`, "
                "ou `chefia` / `chefia123`.")
    agora = datetime.now(TZ)
    if ss.bloqueio_ate and agora < ss.bloqueio_ate:
        st.error(f"Muitas tentativas. Tente novamente às {ss.bloqueio_ate:%H:%M}.")
        st.stop()
    with st.form("login"):
        login = st.text_input("Usuário", autocomplete="username")
        senha = st.text_input("Senha", type="password", autocomplete="current-password")
        entrar = st.form_submit_button("Entrar", type="primary", width="stretch")
    if entrar:
        u = USUARIOS.get(login.strip().lower())
        if u and confere_senha(senha, u["senha"]):
            ss.usuario = {"login": login.strip().lower(), "nome": u["nome"], "perfil": u["perfil"]}
            ss.falhas = 0
            st.rerun()
        ss.falhas += 1
        if ss.falhas >= 5:
            ss.bloqueio_ate = agora + timedelta(minutes=5)
            ss.falhas = 0
        st.error("Usuário ou senha incorretos.")
    st.stop()


if ss.usuario is None:
    tela_login()

usuario = ss.usuario

# ---------------------------------------------------------------------------
# Cabeçalho compacto (pensado para celular)
# ---------------------------------------------------------------------------
st.markdown(f"**🏥 GYN Gestão** · {usuario['nome']} · {hoje():%d/%m}")
if not USE_GOOGLE_SHEETS:
    st.warning("Modo local (CSV): no Streamlit Cloud os dados são **apagados quando o app "
               "reinicia**. Configure o Google Sheets para uso real.")

try:
    df = load_data()
except Exception as e:  # falha de rede/credenciais
    st.error(f"Não foi possível carregar os registros: {e}")
    st.stop()

ativas = censo(df)

if usuario["perfil"] == "Chefia":
    telas = ["Painel", "Registrar visita"]
else:
    telas = ["Registrar visita", "Censo"]
tela = st.segmented_control("Tela", telas, default=telas[0], key="tela",
                            label_visibility="collapsed", width="stretch") or telas[0]


def label_paciente(r) -> str:
    di = "" if pd.isna(r["DI"]) else f" · DI {int(r['DI'])}"
    return f"Leito {r['leito'] or '?'} — {r['id_paciente']}{di}"


def fmt_data(x) -> str:
    return "—" if pd.isna(x) else f"{x:%d/%m}"


def cartao(r, motivo: str = ""):
    """Resumo de uma paciente em formato de cartão (legível no celular)."""
    icon = STATUS_ICON.get(r["cor"], "⚪")
    dpo = "" if pd.isna(r["DPO"]) else f" · DPO {int(r['DPO'])}"
    di = "" if pd.isna(r["DI"]) else f" · DI {int(r['DI'])}"
    with st.container(border=True):
        st.markdown(f"{icon} **Leito {r['leito'] or '?'}** · {r['id_paciente']}{di}{dpo}")
        linhas = []
        atraso = r["dias_alem_previsto"]
        alta = f"Alta prevista {fmt_data(r['_data_prevista_alta'])}"
        if r["alta_atrasada"] and not pd.isna(atraso) and atraso > 0:
            alta += f" (**+{int(atraso)} d**)"
        linhas.append(alta + (f" · {r['procedimento']}" if r["procedimento"] else ""))
        if motivo:
            linhas.append(f"⚠️ {motivo}")
        if r["complicacao"] == "Sim":
            cd = f" (Clavien {r['clavien_dindo']})" if r["clavien_dindo"] else ""
            linhas.append(f"Complicação: {r['tipo_complicacao']}{cd}"
                          + (" — **impede alta**" if r["complicacao_impede_alta"] == "Sim" else ""))
        if r["especialidade"]:
            est = "realizado" if r["parecer_realizado"] == "Sim" else "pendente"
            linhas.append(f"Parecer {r['especialidade']}: {est}"
                          + (" — **atrasando alta**" if r["parecer_atrasa_alta"] == "Sim" else ""))
        if r["pendencia_principal"]:
            prazo = "" if pd.isna(r["_prazo_pendencia"]) else f", até {fmt_data(r['_prazo_pendencia'])}"
            venc = " — **vencida**" if r["pendencia_vencida"] else ""
            resp = f" ({r['responsavel_pendencia']}{prazo})" if r["responsavel_pendencia"] or prazo else ""
            linhas.append(f"Pendência: {r['pendencia_principal']}{resp}{venc}")
        if r["barreira_alta"]:
            linhas.append(f"Barreira: {r['barreira_alta']}")
        visita = "✅ visitada hoje" if r.get("visitada_hoje") else "⏳ sem visita hoje"
        quem = f" por {r['preceptor']}" if r["preceptor"] else ""
        linhas.append(f"<small>{visita} · última {fmt_data(r['_ts'])} {r['_ts']:%H:%M}{quem}</small>"
                      if not pd.isna(r["_ts"]) else visita)
        st.markdown("  \n".join(linhas), unsafe_allow_html=True)


def motivos_criticos(r) -> str:
    m = []
    if r["complicacao_impede_alta"] == "Sim":
        m.append("complicação impede alta")
    if r["parecer_atrasa_alta"] == "Sim":
        m.append(f"parecer {r['especialidade']} atrasando")
    if r["alta_atrasada"]:
        m.append("além da alta prevista")
    if r["pendencia_vencida"]:
        m.append("pendência vencida")
    if not m and r["cor"] == "Vermelho":
        m.append("status vermelho")
    return "; ".join(m)


def resumo_censo(a: pd.DataFrame):
    """Resumo em etiquetas compactas (cabe em 2 linhas no celular)."""
    itens = [
        ("gray", "Internadas", len(a)),
        ("red", "Vermelho", int(a["cor"].eq("Vermelho").sum())),
        ("orange", "Complicação", int(a["complicacao"].eq("Sim").sum())),
        ("orange", "Parecer atrasando", int(a["parecer_atrasa_alta"].eq("Sim").sum())),
        ("red", "Além da alta prevista", int(a["alta_atrasada"].sum())),
        ("blue", "Sem visita hoje", int((~a["visitada_hoje"]).sum())),
    ]
    st.markdown(" ".join(f":{cor if n else 'gray'}-badge[{rot}: **{n}**]" for cor, rot, n in itens))


# ---------------------------------------------------------------------------
# TELA: Registrar visita
# ---------------------------------------------------------------------------
def tela_visita():
    if msg := ss.pop("msg", None):
        st.success(msg)

    pend_hoje = ativas[~ativas["visitada_hoje"]]
    if len(ativas):
        st.caption(
            f"{len(ativas)} internadas · {len(ativas) - len(pend_hoje)} visitadas hoje"
            + (f" · **faltam: {', '.join(pend_hoje['leito'].replace('', '?'))}**" if len(pend_hoje) else " ✅")
        )

    NOVA = "➕ Nova internação"
    # Pacientes ainda sem visita hoje aparecem primeiro
    ordem = ativas.sort_values(["visitada_hoje", "leito"])
    opcoes = {NOVA: None}
    opcoes.update({("⏳ " if not r["visitada_hoje"] else "✅ ") + label_paciente(r): r
                   for _, r in ordem.iterrows()})
    escolha = st.selectbox("Paciente", list(opcoes), key=f"sel_{ss.ver}",
                           index=1 if len(opcoes) > 1 else 0)
    prev = opcoes[escolha]
    pid = "nova" if prev is None else prev["id_paciente"]

    def k(nome):  # chave única por paciente e por envio
        return f"{nome}_{ss.ver}_{pid}"

    def pv(campo, default=""):
        return default if prev is None else (prev[campo] or default)

    def pdate(campo, default):
        if prev is None or pd.isna(prev["_" + campo]):
            return default
        return prev["_" + campo].date()

    def idx(options, value):
        return options.index(value) if value in options else 0

    def idx_none(options, value):  # sem valor prévio = sem seleção
        return options.index(value) if value in options else None

    if prev is not None:
        cartao(prev)

    # --- Cadastro (recolhido para pacientes já internadas) --------------------
    with st.expander("Cadastro e cirurgia" + ("" if prev is None else " (editar)"),
                     expanded=prev is None):
        c1, c2 = st.columns(2)
        leito = c1.text_input("Leito *", value=pv("leito"), key=k("leito"))
        id_paciente = c2.text_input(
            "ID institucional / código *", value=pv("id_paciente"), key=k("id"),
            disabled=prev is not None,
            help="Use código, nunca nome. Para corrigir o ID, registre como nova internação.",
        )
        unidade = st.text_input("Unidade", value=pv("unidade", "Ginecologia Cirúrgica"), key=k("unidade"))
        data_internacao = st.date_input("Data da internação", value=pdate("data_internacao", hoje()),
                                        max_value=hoje(), format="DD/MM/YYYY", key=k("dint"))
        op_cir = ["Programada", "Realizada", "Aguardando definição"]
        cirurgia_status = st.segmented_control(
            "Situação cirúrgica", op_cir, default=pv("cirurgia_status", "Programada")
            if pv("cirurgia_status", "Programada") in op_cir else "Programada",
            key=k("cir"), width="stretch") or "Programada"
        procedimento = st.text_input("Procedimento", value=pv("procedimento"), key=k("proc"))
        data_cirurgia = None
        if cirurgia_status != "Aguardando definição":
            data_cirurgia = st.date_input(
                "Data da cirurgia" + (" (prevista)" if cirurgia_status == "Programada" else ""),
                value=pdate("data_cirurgia", hoje()), format="DD/MM/YYYY", key=k("dcir"))

    di = (hoje() - data_internacao).days + 1
    dpo = max((hoje() - data_cirurgia).days, 0) if cirurgia_status == "Realizada" and data_cirurgia else ""

    # --- Gestão do dia ---------------------------------------------------------
    st.markdown("##### Plano de alta")
    data_prevista_alta = st.date_input(
        "Data prevista de alta", value=pdate("data_prevista_alta", hoje() + timedelta(days=2)),
        format="DD/MM/YYYY", key=k("edd"),
        help="Definida na admissão e revista a cada visita.")
    st.caption(f"DI {di}" + (f" · DPO {dpo}" if dpo != "" else "")
               + f" · alta prevista: **{previsao_relativa(data_prevista_alta)}**")
    barreira_alta = st.text_input("Barreira para alta (vazio = nenhuma)",
                                  value=pv("barreira_alta"), key=k("barr"))
    pendencia_principal = st.text_input("Pendência principal", value=pv("pendencia_principal"),
                                        key=k("pend"))
    c1, c2 = st.columns(2)
    responsavel_pendencia = c1.text_input("Responsável", value=pv("responsavel_pendencia"), key=k("resp"))
    prazo_pendencia = c2.date_input("Prazo", value=pdate("prazo_pendencia", None),
                                    format="DD/MM/YYYY", key=k("prazo"))

    st.markdown("##### Metas das próximas 24 h")
    if prev is not None and any(prev[m] for m in ("meta_1", "meta_2", "meta_3")):
        st.caption("Anteriores: " + " · ".join(prev[m] for m in ("meta_1", "meta_2", "meta_3") if prev[m]))
    meta_1 = st.text_input("Meta 1", key=k("m1"), label_visibility="collapsed", placeholder="Meta 1")
    meta_2 = st.text_input("Meta 2", key=k("m2"), label_visibility="collapsed", placeholder="Meta 2")
    meta_3 = st.text_input("Meta 3", key=k("m3"), label_visibility="collapsed", placeholder="Meta 3")

    # --- Intercorrências -------------------------------------------------------
    st.markdown("##### Intercorrências")
    complicacao = st.radio("Complicação ativa?", SIM_NAO, index=idx(SIM_NAO, pv("complicacao", "Não")),
                           horizontal=True, key=k("comp"))
    tipo_complicacao = clavien = None
    complicacao_impede_alta = "Não"
    if complicacao == "Sim":
        c1, c2 = st.columns(2)
        tipo_complicacao = c1.selectbox("Tipo", TIPOS_COMPLICACAO, placeholder="Selecione",
                                        index=idx_none(TIPOS_COMPLICACAO, pv("tipo_complicacao")),
                                        key=k("tcomp"))
        clavien = c2.selectbox("Clavien-Dindo", CLAVIEN, placeholder="Selecione", help=CLAVIEN_HELP,
                               index=idx_none(CLAVIEN, pv("clavien_dindo")), key=k("cd"))
        complicacao_impede_alta = st.radio("Complicação impede a alta?", SIM_NAO, horizontal=True,
                                           index=idx(SIM_NAO, pv("complicacao_impede_alta", "Não")),
                                           key=k("cimp"))

    especialista_necessario = st.radio("Parecer de especialista?", SIM_NAO, horizontal=True,
                                       index=idx(SIM_NAO, pv("especialista_necessario", "Não")),
                                       key=k("esp"))
    especialidade = None
    parecer_solicitado = parecer_realizado = parecer_atrasa_alta = "Não"
    if especialista_necessario == "Sim":
        especialidade = st.selectbox("Especialidade", ESPECIALIDADES, placeholder="Selecione",
                                     index=idx_none(ESPECIALIDADES, pv("especialidade")), key=k("espc"))
        c1, c2, c3 = st.columns(3)
        parecer_solicitado = c1.radio("Solicitado?", SIM_NAO, horizontal=True,
                                      index=idx(SIM_NAO, pv("parecer_solicitado", "Não")), key=k("psol"))
        parecer_realizado = c2.radio("Realizado?", SIM_NAO, horizontal=True,
                                     index=idx(SIM_NAO, pv("parecer_realizado", "Não")), key=k("preal"))
        parecer_atrasa_alta = c3.radio("Atrasa a alta?", SIM_NAO, horizontal=True,
                                       index=idx(SIM_NAO, pv("parecer_atrasa_alta", "Não")), key=k("patr"))

    demanda_enfermagem = st.radio("Demanda de enfermagem?", SIM_NAO, horizontal=True,
                                  index=idx(SIM_NAO, pv("demanda_enfermagem", "Não")), key=k("denf"))
    tipo_demanda_enfermagem = ""
    if demanda_enfermagem == "Sim":
        prev_dem = [x.strip() for x in pv("tipo_demanda_enfermagem").split(";") if x.strip() in DEMANDAS_ENF]
        tipo_demanda_enfermagem = "; ".join(
            st.multiselect("Tipo(s) de demanda", DEMANDAS_ENF, default=prev_dem, key=k("tenf")))

    observacoes = st.text_area("Observações de gestão", key=k("obs"), height=80)

    # --- Status e desfecho -----------------------------------------------------
    st.markdown("##### Status do dia")
    sugerido = status_sugerido(dict(
        complicacao_impede_alta=complicacao_impede_alta, parecer_atrasa_alta=parecer_atrasa_alta,
        data_prevista_alta=data_prevista_alta, complicacao=complicacao,
        parecer_solicitado=parecer_solicitado, parecer_realizado=parecer_realizado,
        barreira_alta=barreira_alta, prazo_pendencia=prazo_pendencia,
    ))
    rotulos = {s: f"{STATUS_ICON[status_cor(s)]} {status_cor(s)}" for s in STATUS}
    rotulos["auto"] = "Sugerido"
    st.caption(f"Sugerido pelas regras: **{rotulos[sugerido]}**")
    escolha_st = st.segmented_control(
        "Status", ["auto"] + STATUS, default="auto", format_func=rotulos.get,
        key=k("st"), width="stretch", label_visibility="collapsed") or "auto"
    status_gestao = sugerido if escolha_st == "auto" else escolha_st
    alta_realizada = st.toggle("Alta realizada", key=k("alta"))
    data_alta = st.date_input("Data da alta", value=hoje(), format="DD/MM/YYYY",
                              key=k("dalta")) if alta_realizada else None

    if st.button("Salvar visita", type="primary", width="stretch"):
        erros = []
        if not leito.strip():
            erros.append("Informe o leito (em *Cadastro e cirurgia*).")
        if not id_paciente.strip():
            erros.append("Informe o ID institucional/código (em *Cadastro e cirurgia*).")
        elif prev is None and id_paciente.strip() in set(ativas["id_paciente"]):
            erros.append("Já existe internação ativa com este ID — selecione-a na lista.")
        if data_cirurgia and cirurgia_status == "Realizada" and data_cirurgia < data_internacao:
            erros.append("Data da cirurgia anterior à internação.")
        if complicacao == "Sim" and not (tipo_complicacao and clavien):
            erros.append("Complicação ativa: informe o tipo e o grau de Clavien-Dindo.")
        if especialista_necessario == "Sim" and not especialidade:
            erros.append("Informe a especialidade do parecer.")
        if data_prevista_alta < data_internacao:
            erros.append("Data prevista de alta anterior à internação.")
        if data_alta and data_alta < data_internacao:
            erros.append("Data da alta anterior à internação.")
        if erros:
            for e in erros:
                st.error(e)
            return
        rec = {
            "timestamp": agora_str(),
            "preceptor": usuario["nome"],
            "unidade": unidade.strip(),
            "leito": leito.strip(),
            "id_paciente": id_paciente.strip(),
            "data_internacao": data_internacao.isoformat(),
            "dia_internacao": di,
            "cirurgia_status": cirurgia_status,
            "procedimento": procedimento.strip(),
            "data_cirurgia": data_cirurgia.isoformat() if data_cirurgia else "",
            "dpo": dpo,
            "dias_previstos_internacao": (data_prevista_alta - data_internacao).days,
            "data_prevista_alta": data_prevista_alta.isoformat(),
            "complicacao": complicacao,
            "tipo_complicacao": tipo_complicacao or "",
            "complicacao_impede_alta": complicacao_impede_alta,
            "especialista_necessario": especialista_necessario,
            "especialidade": especialidade or "",
            "parecer_solicitado": parecer_solicitado,
            "parecer_realizado": parecer_realizado,
            "parecer_atrasa_alta": parecer_atrasa_alta,
            "demanda_enfermagem": demanda_enfermagem,
            "tipo_demanda_enfermagem": tipo_demanda_enfermagem,
            "pendencia_principal": pendencia_principal.strip(),
            "responsavel_pendencia": responsavel_pendencia.strip(),
            "prazo_pendencia": prazo_pendencia.isoformat() if prazo_pendencia else "",
            "status_gestao": status_gestao,
            "barreira_alta": barreira_alta.strip(),
            "meta_1": meta_1.strip(),
            "meta_2": meta_2.strip(),
            "meta_3": meta_3.strip(),
            "alta_previsao": "Alta realizada" if data_alta else previsao_relativa(data_prevista_alta),
            "observacoes": observacoes.strip(),
            "clavien_dindo": clavien or "",
            "alta_realizada": "Sim" if data_alta else "Não",
            "data_alta": data_alta.isoformat() if data_alta else "",
            "usuario": usuario["login"],
        }
        try:
            save_record(rec)
        except Exception as e:
            st.error(f"Falha ao salvar (nada foi gravado): {e}")
            return
        ss.msg = f"Leito {rec['leito']} salvo" + (" — alta registrada." if data_alta else ".")
        ss.ver += 1
        st.rerun()


# ---------------------------------------------------------------------------
# TELA: Censo (preceptor) — somente leitura, em cartões
# ---------------------------------------------------------------------------
def tela_censo():
    if ativas.empty:
        st.info("Nenhuma paciente internada.")
        return
    resumo_censo(ativas)
    for _, r in ativas.iterrows():
        cartao(r)


# ---------------------------------------------------------------------------
# TELA: Painel da chefia
# ---------------------------------------------------------------------------
def tabela(d: pd.DataFrame, cols: list):
    d = d.copy()
    d["Status"] = d["cor"].map(STATUS_ICON).fillna("⚪") + " " + d["cor"]
    st.dataframe(
        d[cols], hide_index=True, width="stretch",
        column_config={
            "leito": "Leito", "id_paciente": "ID",
            "DI": st.column_config.NumberColumn("DI", format="%d"),
            "DPO": st.column_config.NumberColumn("DPO", format="%d"),
            "_data_prevista_alta": st.column_config.DateColumn("Alta prevista", format="DD/MM"),
            "dias_alem_previsto": st.column_config.NumberColumn("Dias além do previsto", format="%d"),
            "_prazo_pendencia": st.column_config.DateColumn("Prazo", format="DD/MM"),
            "_ts": st.column_config.DatetimeColumn("Última visita", format="DD/MM HH:mm"),
            "procedimento": "Procedimento", "tipo_complicacao": "Complicação",
            "clavien_dindo": "Clavien", "especialidade": "Parecer",
            "tipo_demanda_enfermagem": "Enfermagem",
            "pendencia_principal": "Pendência", "responsavel_pendencia": "Responsável",
            "barreira_alta": "Barreira p/ alta", "visitada_hoje": "Visitada hoje",
            "preceptor": "Preceptor", "Motivo": "Motivo",
        },
    )


def tela_painel():
    global df, ativas
    if df.empty:
        st.info("Ainda não há registros.")
        return

    unidades = sorted(u for u in df["unidade"].unique() if u)
    if len(unidades) > 1:
        filtro = st.multiselect("Unidade", unidades, placeholder="Todas as unidades")
        if filtro:
            df = df[df["unidade"].isin(filtro)]
            ativas = ativas[ativas["unidade"].isin(filtro)]

    como_tabela = st.toggle("Ver como tabela", value=False,
                            help="Cartões são mais legíveis no celular; tabela, no computador.")
    tab_censo, tab_pend, tab_ind, tab_prec, tab_hist = st.tabs(
        ["Censo", "Críticas", "Indicadores", "Preceptores", "Histórico"])

    with tab_censo:
        if ativas.empty:
            st.info("Nenhuma paciente internada no momento.")
        else:
            resumo_censo(ativas)
            if como_tabela:
                tabela(ativas, ["Status", "leito", "id_paciente", "DI", "DPO", "procedimento",
                                "_data_prevista_alta", "dias_alem_previsto", "tipo_complicacao",
                                "clavien_dindo", "especialidade", "pendencia_principal",
                                "responsavel_pendencia", "barreira_alta", "visitada_hoje",
                                "preceptor", "_ts"])
            else:
                for _, r in ativas.iterrows():
                    cartao(r)

    with tab_pend:
        crit = ativas[
            ativas["cor"].eq("Vermelho")
            | ativas["parecer_atrasa_alta"].eq("Sim")
            | ativas["complicacao_impede_alta"].eq("Sim")
            | ativas["alta_atrasada"]
            | ativas["pendencia_vencida"]
        ].copy()
        if crit.empty:
            st.success("Sem pendências críticas.")
        else:
            crit["Motivo"] = [motivos_criticos(r) for _, r in crit.iterrows()]
            if como_tabela:
                tabela(crit, ["Status", "leito", "id_paciente", "DI", "DPO", "Motivo",
                              "pendencia_principal", "responsavel_pendencia", "_prazo_pendencia",
                              "barreira_alta"])
            else:
                for _, r in crit.iterrows():
                    cartao(r, r["Motivo"])
        par = ativas[ativas["parecer_solicitado"].eq("Sim") & ativas["parecer_realizado"].eq("Não")]
        if len(par):
            st.markdown("**Pareceres solicitados e não realizados**")
            st.dataframe(par.groupby("especialidade").size().rename("Pacientes"), width="content")

    with tab_ind:
        d = df[df["id_paciente"].ne("")].copy()
        d["internacao"] = d["id_paciente"] + "|" + d["data_internacao"]
        encerradas = d[d["alta_realizada"].eq("Sim")].sort_values("_ts").groupby("internacao").tail(1)
        if encerradas.empty:
            st.info("Os indicadores aparecem quando houver altas registradas pelo app.")
        else:
            min_d = encerradas["_data_alta"].min().date()
            periodo = st.date_input("Período (data da alta)", value=(min_d, hoje()), format="DD/MM/YYYY")
            ini, fim_ = (periodo[0], periodo[-1]) if periodo else (min_d, hoje())
            e = encerradas[(encerradas["_data_alta"].dt.date >= ini)
                           & (encerradas["_data_alta"].dt.date <= fim_)]
            comp_any = d[d["complicacao"].eq("Sim")]["internacao"].unique()
            e = e.assign(teve_comp=e["internacao"].isin(comp_any),
                         no_prazo=e["_data_alta"] <= e["_data_prevista_alta"])
            n = len(e)
            with st.container(horizontal=True, gap="medium"):
                st.metric("Altas", n, width="content")
                st.metric("Permanência média", f"{e['DI'].mean():.1f} d" if n else "—", width="content")
                st.metric("Mediana", f"{e['DI'].median():.0f} d" if n else "—", width="content")
                st.metric("Com complicação", f"{100 * e['teve_comp'].mean():.0f}%" if n else "—",
                          width="content")
                st.metric("Alta até a data prevista", f"{100 * e['no_prazo'].mean():.0f}%" if n else "—",
                          width="content",
                          help="Compara a data da alta com a última data prevista registrada.")
            comp = d[d["internacao"].isin(e["internacao"]) & d["complicacao"].eq("Sim")]
            comp = comp.drop_duplicates(["internacao", "tipo_complicacao"])
            if len(comp):
                st.markdown("**Complicações por tipo**")
                st.bar_chart(comp["tipo_complicacao"].value_counts(), horizontal=True, height=220)
                cd = comp[comp["clavien_dindo"].isin(CLAVIEN)].groupby("internacao")["clavien_dindo"] \
                    .agg(lambda s: max(s, key=CLAVIEN.index))
                if len(cd):
                    st.markdown("**Maior grau Clavien-Dindo por internação**")
                    st.bar_chart(cd.value_counts().reindex(CLAVIEN, fill_value=0), height=220)
            proc = e[e["procedimento"].ne("")].groupby("procedimento")["DI"] \
                .agg(["count", "mean", "median"]).sort_values("count", ascending=False)
            if len(proc):
                st.markdown("**Permanência por procedimento**")
                st.dataframe(proc.rename(columns={"count": "Altas", "mean": "Média", "median": "Mediana"})
                             .round(1), width="stretch")

    with tab_prec:
        st.caption("Adesão à visita de gestão por preceptor.")
        v = df[df["preceptor"].ne("")].copy()
        if v.empty:
            st.info("Sem visitas identificadas.")
        else:
            v["dia"] = v["_ts"].dt.date
            janela = st.segmented_control("Período", ["7 dias", "30 dias", "Tudo"], default="7 dias") or "7 dias"
            if janela != "Tudo":
                v = v[v["dia"] >= hoje() - timedelta(days=int(janela.split()[0]) - 1)]
            res = v.groupby("preceptor").agg(Visitas=("id_paciente", "size"),
                                             Pacientes=("id_paciente", "nunique"),
                                             Dias=("dia", "nunique"),
                                             Última=("_ts", "max")).sort_values("Visitas", ascending=False)
            st.dataframe(res, width="stretch", column_config={
                "Última": st.column_config.DatetimeColumn(format="DD/MM HH:mm")})
            hoje_v = ativas[ativas["visitada_hoje"]]
            st.caption(f"Hoje: {len(hoje_v)} de {len(ativas)} pacientes internadas visitadas.")

    with tab_hist:
        ids = sorted(i for i in df["id_paciente"].unique() if i)
        sel = st.multiselect("Paciente (ID)", ids, placeholder="Todas")
        h = df[df["id_paciente"].isin(sel)] if sel else df
        h = h.sort_values("_ts", ascending=False)
        st.dataframe(h[FIELDS], hide_index=True, width="stretch")
        c1, c2 = st.columns(2)
        c1.download_button("Baixar CSV", h[FIELDS].to_csv(index=False).encode("utf-8-sig"),
                           file_name=f"gyn_gestao_{hoje():%Y%m%d}.csv", mime="text/csv",
                           width="stretch")
        buf = io.BytesIO()
        h[FIELDS].to_excel(buf, index=False, engine="openpyxl")
        c2.download_button("Baixar Excel", buf.getvalue(),
                           file_name=f"gyn_gestao_{hoje():%Y%m%d}.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           width="stretch")

    with st.expander("Administração: gerar senha de usuário"):
        st.caption("Gera a linha de senha protegida (hash) para colar nos *secrets* do app, "
                   "na seção `[usuarios.<login>]`. A senha digitada não é gravada.")
        nova = st.text_input("Nova senha", type="password", key="nova_senha")
        if nova:
            if len(nova) < 8:
                st.warning("Use ao menos 8 caracteres.")
            else:
                st.code(f'senha = "{hash_senha(nova)}"', language="toml")


# ---------------------------------------------------------------------------
if tela == "Registrar visita":
    tela_visita()
elif tela == "Censo":
    tela_censo()
else:
    tela_painel()

st.divider()
if st.button(f"Sair ({usuario['login']})", type="tertiary"):
    ss.usuario = None
    st.rerun()
