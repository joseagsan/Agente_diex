"""
Lista de UASGs usadas na busca de item por prioridade geográfica (página
"Buscar Item") — persistida numa aba própria da planilha do SSAC.

PRIORIDADE segue a lógica operacional do usuário:
  1 = nossas UASGs em Boa Vista (1ª Bda e subordinadas/apoiadas)
  2 = Boa Vista, outros órgãos
  3 = Manaus, Exército
  4 = Manaus, demais órgãos (inclusive Marinha e FAB)

A API pública de atas só permite buscar por UASG gerenciadora específica —
não existe filtro por UF/município. Por isso as prioridades 2 a 4 dependem
do usuário cadastrar aqui as UASGs que conhece; a prioridade 1 já vem
pré-cadastrada com as 5 UASGs do apoio direto à 1ª Bda em Boa Vista.
"""
import logging

from sheets_nc import _conectar
from config import SHEET_ID_NC

logger = logging.getLogger(__name__)

ABA_UASGS = "SSAC_UASGS_BUSCA"
COLUNAS = ["UASG", "NOME", "PRIORIDADE"]

_PADRAO_P1 = [
    ("160482", "Comando da 1ª Bda Inf Sl"),
    ("160352", "Comando de Fronteira Roraima / 7º BIS"),
    ("160353", "6º Batalhão de Engenharia de Construção (6º BEC)"),
    ("110794", "SECAAE (Operação Acolhida)"),
    ("160907", "1º Batalhão Logístico de Selva (BLog)"),
]


def _ws():
    client = _conectar()
    planilha = client.open_by_key(SHEET_ID_NC)
    try:
        ws = planilha.worksheet(ABA_UASGS)
    except Exception:
        ws = planilha.add_worksheet(title=ABA_UASGS, rows=200, cols=len(COLUNAS))
        linhas = [COLUNAS] + [[uasg, nome, "1"] for uasg, nome in _PADRAO_P1]
        ws.update("A1", linhas, value_input_option="RAW")
        logger.info("Aba %s criada com as UASGs padrão da Prioridade 1.", ABA_UASGS)
        return ws
    return ws


def ler_uasgs_busca() -> list[dict]:
    ws = _ws()
    rows = ws.get_all_values()
    if len(rows) < 2:
        return []
    headers = rows[0]
    return [dict(zip(headers, row)) for row in rows[1:] if row and row[0].strip()]


def adicionar_uasg_busca(uasg: str, nome: str, prioridade: int) -> None:
    ws = _ws()
    if any(r["UASG"] == uasg for r in ler_uasgs_busca()):
        return
    ws.append_row([uasg, nome, str(prioridade)], value_input_option="RAW")
    logger.info("UASG %s adicionada à prioridade %s.", uasg, prioridade)


def remover_uasg_busca(uasg: str) -> None:
    ws = _ws()
    rows = ws.get_all_values()
    for i, row in enumerate(rows[1:], start=2):
        if row and row[0].strip() == uasg:
            ws.delete_rows(i)
            logger.info("UASG %s removida.", uasg)
            return
