"""
Lista de UGs (UASGs gerenciadoras de ARP) que o usuário monitora na
página "Atas" — persistida numa aba própria da planilha do SSAC, para
sobreviver a reinícios do app.
"""
import logging

from sheets_nc import _conectar
from config import SHEET_ID_NC

logger = logging.getLogger(__name__)

ABA_UGS = "SSAC_UGS_MONITORADAS"
COLUNAS = ["UG", "NOME"]


def _ws():
    client = _conectar()
    planilha = client.open_by_key(SHEET_ID_NC)
    try:
        ws = planilha.worksheet(ABA_UGS)
    except Exception:
        ws = planilha.add_worksheet(title=ABA_UGS, rows=200, cols=len(COLUNAS))
        ws.update("A1", [COLUNAS], value_input_option="RAW")
        logger.info("Aba %s criada.", ABA_UGS)
        return ws
    return ws


def ler_ugs_monitoradas() -> list[dict]:
    ws = _ws()
    rows = ws.get_all_values()
    if len(rows) < 2:
        return []
    headers = rows[0]
    return [dict(zip(headers, row)) for row in rows[1:] if row and row[0].strip()]


def adicionar_ug(ug: str, nome: str = "") -> None:
    ws = _ws()
    if any(r["UG"] == ug for r in ler_ugs_monitoradas()):
        return
    ws.append_row([ug, nome], value_input_option="RAW")
    logger.info("UG %s adicionada às monitoradas.", ug)


def remover_ug(ug: str) -> None:
    ws = _ws()
    rows = ws.get_all_values()
    for i, row in enumerate(rows[1:], start=2):
        if row and row[0].strip() == ug:
            ws.delete_rows(i)
            logger.info("UG %s removida das monitoradas.", ug)
            return
