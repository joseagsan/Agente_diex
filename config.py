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
ABA_NCS              = _s("ABA_NCS",  "NC 2026")
ABA_REQS             = _s("ABA_REQS", "REQUISIÇÕES DE EMPENHOS")
ABA_FORNECEDORES     = _s("ABA_FORNECEDORES", "AGENDA")
UG_PADRAO            = _s("UG",  "160482")
OM_PADRAO            = _s("OM",  "10º GAC Sl")
