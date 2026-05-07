"""
Leitura da aba Dados_para_REQ via Google Sheets API (Service Account).
"""
import logging

import gspread
from google.oauth2.service_account import Credentials

from config import GOOGLE_SHEET_ID, GOOGLE_CREDENTIALS_FILE, ABA_INPUT

logger = logging.getLogger(__name__)

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets.readonly",
    "https://www.googleapis.com/auth/drive.readonly",
]


def ler_dados() -> list[dict]:
    """
    Retorna todas as linhas da aba ABA_INPUT como lista de dicionários.
    A primeira linha da planilha deve conter os cabeçalhos (nomes dos campos/placeholders).
    """
    if not GOOGLE_SHEET_ID:
        raise ValueError("GOOGLE_SHEET_ID não configurado no .env")

    creds = Credentials.from_service_account_file(GOOGLE_CREDENTIALS_FILE, scopes=SCOPES)
    client = gspread.authorize(creds)
    planilha = client.open_by_key(GOOGLE_SHEET_ID)
    ws = planilha.worksheet(ABA_INPUT)
    registros = ws.get_all_records(empty2zero=False, head=1)
    logger.info("%d linha(s) lida(s) da aba '%s'.", len(registros), ABA_INPUT)
    return registros
