"""
Acompanhamento interno (SSAC) da etapa de cada NC.

As NCs em si são somente leitura — vêm da planilha de controle de crédito
da 1ª Bda Inf Sl e são mantidas pela SALC de lá. Este módulo guarda, por
fora, em que etapa o GAC está com cada NC (principalmente as EM TELA), para
ajudar a enxergar onde agir e reduzir o valor total parado.

Fica em uma aba própria (ABA_NC_STATUS) na planilha do SSAC — a mesma onde
já vivem as Requisições (SHEET_ID_NC).
"""
import logging
from datetime import datetime

from sheets_nc import _conectar
from config import SHEET_ID_NC

logger = logging.getLogger(__name__)

ABA_NC_STATUS = "SSAC_NC_STATUS"
COLUNAS = ["NC", "ETAPA", "ATUALIZADO_EM"]

# Etapas fixas — cobrem o percurso típico de uma NC até deixar de estar EM TELA.
ETAPAS_NC = [
    "Não iniciada",
    "Em pesquisa de preço/cotação",
    "Aguardando pregão/ata",
    "Documentação pendente",
    "Pronta para requisitar",
    "Requisição em andamento",
]


def _ws():
    client = _conectar()
    planilha = client.open_by_key(SHEET_ID_NC)
    try:
        return planilha.worksheet(ABA_NC_STATUS)
    except Exception:
        ws = planilha.add_worksheet(title=ABA_NC_STATUS, rows=500, cols=len(COLUNAS))
        ws.update("A1", [COLUNAS], value_input_option="RAW")
        logger.info("Aba %s criada.", ABA_NC_STATUS)
        return ws


def ler_status_ncs() -> dict[str, str]:
    """Retorna {NC: ETAPA} para todas as NCs com etapa registrada."""
    ws = _ws()
    valores = ws.get_all_values()
    if len(valores) < 2:
        return {}
    headers = valores[0]
    idx_nc    = headers.index("NC")    if "NC"    in headers else 0
    idx_etapa = headers.index("ETAPA") if "ETAPA" in headers else 1
    resultado = {}
    for row in valores[1:]:
        if len(row) > max(idx_nc, idx_etapa) and row[idx_nc].strip():
            resultado[row[idx_nc].strip()] = row[idx_etapa]
    return resultado


def definir_status_ncs(mudancas: dict[str, str]) -> None:
    """Atualiza (ou cria) a etapa de uma ou mais NCs. `mudancas` = {NC: ETAPA}."""
    if not mudancas:
        return
    from gspread.utils import rowcol_to_a1

    ws      = _ws()
    valores = ws.get_all_values()
    if not valores:
        ws.update("A1", [COLUNAS], value_input_option="RAW")
        valores = [COLUNAS]
    headers = valores[0]
    idx_nc  = headers.index("NC") if "NC" in headers else 0

    linha_por_nc = {
        row[idx_nc].strip(): i
        for i, row in enumerate(valores[1:], start=2)
        if row and len(row) > idx_nc and row[idx_nc].strip()
    }

    agora   = datetime.today().strftime("%d/%m/%Y %H:%M")
    novas   = []
    batch   = []
    for nc_num, etapa in mudancas.items():
        nc_num = str(nc_num).strip()
        if not nc_num:
            continue
        if nc_num in linha_por_nc:
            i = linha_por_nc[nc_num]
            batch.append({"range": rowcol_to_a1(i, 2), "values": [[etapa]]})
            batch.append({"range": rowcol_to_a1(i, 3), "values": [[agora]]})
        else:
            novas.append([nc_num, etapa, agora])

    if batch:
        ws.batch_update(batch, value_input_option="RAW")
    if novas:
        proxima = len(valores) + 1
        if proxima + len(novas) > ws.row_count:
            ws.add_rows(len(novas) + 50)
        end_col = rowcol_to_a1(1, len(COLUNAS)).replace("1", "")
        ws.update(
            range_name=f"A{proxima}:{end_col}{proxima + len(novas) - 1}",
            values=novas,
            value_input_option="RAW",
        )
    logger.info("Etapa atualizada para %d NC(s).", len(mudancas))
