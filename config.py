import os
from dotenv import load_dotenv

load_dotenv()


def _s(key: str, default: str = "") -> str:
    """Lê de Streamlit secrets (cloud) ou variáveis de ambiente (local)."""
    try:
        import streamlit as st
        v = st.secrets.get(key, "")
        if v:
            return str(v)
    except Exception:
        pass
    return os.getenv(key, default)


# --- Gerador de documentos (existente) ---
GOOGLE_SHEET_ID      = _s("GOOGLE_SHEET_ID")
GOOGLE_CREDENTIALS_FILE = _s("GOOGLE_CREDENTIALS_FILE", "absolute-disk-428101-n6-2b4d13cfdeb6.json")
ABA_INPUT            = _s("ABA_INPUT", "Dados_para_REQ")
TEMPLATE_PATH        = _s("TEMPLATE_PATH", "templates/requisicao_empenho_fiel_placeholders.docx")
OUTPUT_DIR           = _s("OUTPUT_DIR", "output")

# --- Controle de NCs ---
ANTHROPIC_API_KEY    = _s("ANTHROPIC_API_KEY")
SHEET_ID_NC          = _s("SHEET_ID_NC") or _s("GOOGLE_SHEET_ID", "1AnOjrKRqCD4Y3lqjlfWPVytYDeY6xLVQXil9rCkCFio")
ABA_REQS             = _s("ABA_REQS", "REQUISIÇÕES DE EMPENHOS")
ABA_SSAC_REQS        = _s("ABA_SSAC_REQS", "SSAC_REQS")
ABA_FORNECEDORES     = _s("ABA_FORNECEDORES", "AGENDA")
UG_PADRAO            = _s("UG",  "160482")
OM_PADRAO            = _s("OM",  "10º GAC Sl")
DRIVE_FOLDER_ID      = _s("DRIVE_FOLDER_ID", "")
EMAIL_SENDER         = _s("EMAIL_SENDER",  "")   # ex: sac10gacsl@gmail.com
EMAIL_PASSWORD       = _s("EMAIL_PASSWORD", "")  # senha de app do Gmail

# --- NCs: planilha de controle de crédito da 1ª Bda Inf Sl (somente leitura) ---
# O SSAC não edita NCs — elas são lançadas e mantidas pela SALC da Bda.
# Lemos as abas por UG e filtramos pela coluna RESP (responsável) = RESP_PADRAO.
SHEET_ID_NC_ORIGEM   = _s("SHEET_ID_NC_ORIGEM", "1PhxwM7BLauZQEYNyR7P-T2i2huwyGLaUAy3YVJpD-fs")
ABAS_NC_ORIGEM       = [a.strip() for a in _s("ABAS_NC_ORIGEM", "160482,160487").split(",") if a.strip()]
RESP_PADRAO          = _s("RESP_PADRAO", "10º GAC Sl")

# Planilha "Consolidado" — controle de liquidação/pagamento por NE (empenho),
# com situação, valores por estágio e datas de envio ao fornecedor.
# Se ABA_CONSOLIDADO ficar vazio, usa a primeira aba da planilha.
SHEET_ID_CONSOLIDADO = _s("SHEET_ID_CONSOLIDADO", "1iR8aqsMNgW_SjtRGDYUETftil-8u0K3NQNi_BzdMc4U")
ABA_CONSOLIDADO      = _s("ABA_CONSOLIDADO", "")

# Código de Gestão usado na URL do Portal da Transparência (padrão SIAFI: 00001)
GESTAO_PADRAO        = _s("GESTAO_PADRAO", "00001")
