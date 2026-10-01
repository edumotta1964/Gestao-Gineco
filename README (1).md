
# GYN Gestão — v2

## Novidades da v2
- **Censo atual**: o painel da chefia mostra uma linha por paciente (última visita), não o histórico inteiro.
- **DI, DPO e atraso recalculados diariamente** a partir das datas (antes ficavam congelados no dia do registro).
- **Pré-preenchimento**: ao selecionar uma paciente internada, o formulário vem com os dados da última visita.
- **Campos condicionais funcionando** (tipo de complicação, especialidade, pareceres, demandas de enfermagem).
- **Registro de alta** e aba de **Indicadores** (permanência média/mediana, % complicação, % alta até a data prevista, Clavien-Dindo, permanência por procedimento).
- **Status sugerido automaticamente** a partir de regras explícitas (pode ser sobrescrito).
- Classificação de **Clavien-Dindo** para complicações.
- **Fuso de São Paulo** (o servidor roda em UTC; após as 21h a data saía errada).
- Google Sheets: gravação `RAW` (sem perder zeros à esquerda nem converter datas), escopo mínimo, cache de 60 s e migração automática do cabeçalho.
- Exportação CSV e Excel.
- **Login individual** por usuário (senha protegida por hash); a visita fica assinada com o nome de quem registrou e a chefia vê a adesão por preceptor.
- **Interface para celular**: cartões por paciente, cadastro recolhido, pacientes sem visita hoje primeiro na lista.

Compatível com os registros da v1: as colunas novas (`clavien_dindo`, `alta_realizada`, `data_alta`, `usuario`) são acrescentadas ao final da planilha.

Aplicativo para visita de gestão em enfermaria de Ginecologia Cirúrgica.

## Perfis

### Preceptor
Registra:
- dia de internação;
- situação cirúrgica;
- procedimento e data da cirurgia;
- dia pós-operatório;
- permanência prevista;
- complicações;
- necessidade de especialista;
- status do parecer;
- demandas de enfermagem;
- pendência principal;
- responsável;
- prazo;
- barreira à alta;
- metas das próximas 24 horas;
- previsão de alta.

### Chefia
Visualiza:
- tabela consolidada;
- complicações;
- pareceres atrasando alta;
- pacientes acima do tempo previsto;
- pendências críticas;
- filtros;
- exportação CSV.

## Rodar localmente

1. Instale Python 3.11 ou superior.
2. Abra o terminal na pasta do projeto.
3. Execute:

```bash
pip install -r requirements.txt
streamlit run app.py
```

Sem usuários configurados e sem Google Sheets, o app entra em modo demonstração
(usuário `preceptor` / `demo123`, ou `chefia` / `chefia123`).
Com Google Sheets configurado, o modo demonstração é bloqueado: é obrigatório cadastrar usuários.

## Usuários (login individual)

Nos *secrets* do app (Streamlit Cloud → app → Settings → Secrets), um bloco por pessoa:

```toml
[usuarios.emotta]
nome = "Dr. Eduardo Motta"
perfil = "Chefia"          # Chefia ou Preceptor
senha = "pbkdf2$200000$..." # gere em Painel → Administração: gerar senha

[usuarios.ana]
nome = "Dra. Ana"
perfil = "Preceptor"
senha = "pbkdf2$200000$..."
```

- O login é o nome após `usuarios.` (sem diferenciar maiúsculas).
- A senha pode ser texto simples, mas prefira o *hash* gerado no próprio app.
- Para remover alguém, apague o bloco e salve os *secrets*.
- Após 5 tentativas erradas, o login fica bloqueado por 5 minutos naquela sessão.

## Google Sheets compartilhado

O app já está preparado para usar uma planilha do Google Sheets.

Crie uma planilha e uma conta de serviço no Google Cloud. Depois crie:

`.streamlit/secrets.toml`

com:

```toml
google_sheet_id = "ID_DA_PLANILHA"

[gcp_service_account]
type = "service_account"
project_id = "..."
private_key_id = "..."
private_key = """-----BEGIN PRIVATE KEY-----
...
-----END PRIVATE KEY-----
"""
client_email = "..."
client_id = "..."
auth_uri = "https://accounts.google.com/o/oauth2/auth"
token_uri = "https://oauth2.googleapis.com/token"
auth_provider_x509_cert_url = "https://www.googleapis.com/oauth2/v1/certs"
client_x509_cert_url = "..."
```

Compartilhe a planilha com o `client_email` da conta de serviço.

O app criará automaticamente uma aba chamada `visitas`.

## LGPD e uso institucional

Este protótipo foi desenhado para usar `ID institucional / código` e leito em vez de nome completo.

Para produção:
- usar autenticação institucional/SSO;
- hospedar em ambiente aprovado pela instituição;
- limitar acesso à chefia e equipe assistencial;
- evitar dados identificáveis além do mínimo necessário;
- definir política de retenção;
- validar o fluxo com TI, Segurança da Informação e encarregado de dados/LGPD.

Não assumir que hospedagem pública ou planilha pessoal sejam adequadas para dados clínicos identificáveis sem aprovação institucional.
