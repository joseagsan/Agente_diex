"""
CRUD limpo para a aba SSAC_REQS — sem fórmulas, sem conflito.
Colunas: REQ | DATA | NC | NE | PI | ND | EMPRESA | CNPJ | PREGAO
         TIPO | VALOR | SITUACAO | ENTRADA_SALC | OBS | ITENS_JSON
"""
import json
import logging
from datetime import datetime

from sheets_nc import _conectar, parse_moeda, format_moeda
from config import SHEET_ID_NC, ABA_SSAC_REQS

logger = logging.getLogger(__name__)

COLUNAS = [
    "REQ", "DATA", "NC", "NE", "PI", "ND",
    "EMPRESA", "CNPJ", "PREGAO", "TIPO",
    "VALOR", "SITUACAO", "ENTRADA_SALC", "OBS", "ITENS_JSON",
]


def _ws():
    client   = _conectar()
    planilha = client.open_by_key(SHEET_ID_NC)
    try:
        return planilha.worksheet(ABA_SSAC_REQS)
    except Exception:
        ws = planilha.add_worksheet(title=ABA_SSAC_REQS, rows=1000, cols=len(COLUNAS))
        ws.update("A1", [COLUNAS], value_input_option="RAW")
        logger.info("Aba %s criada.", ABA_SSAC_REQS)
        return ws


def ler_reqs() -> list[dict]:
    ws   = _ws()
    rows = ws.get_all_values()
    if len(rows) < 2:
        return []
    headers = rows[0]
    result  = []
    for row in rows[1:]:
        row += [""] * (len(headers) - len(row))
        result.append(dict(zip(headers, row)))
    return [r for r in result if r.get("REQ") or r.get("EMPRESA")]


def adicionar_req(dados: dict) -> None:
    ws      = _ws()
    todas   = ws.get_all_values()
    proxima = len(todas) + 1

    if proxima > ws.row_count:
        ws.add_rows(200)

    valor = dados.get("VALOR", 0.0)
    valor_fmt = format_moeda(float(valor)) if isinstance(valor, (int, float)) else str(valor)

    itens = dados.get("ITENS", [])
    itens_json = json.dumps(itens, ensure_ascii=False) if itens else ""

    linha = {
        "REQ":          str(dados.get("REQ", "")),
        "DATA":         dados.get("DATA", datetime.today().strftime("%d/%m/%Y")),
        "NC":           dados.get("NC", ""),
        "NE":           dados.get("NE", ""),
        "PI":           dados.get("PI", ""),
        "ND":           dados.get("ND", ""),
        "EMPRESA":      dados.get("EMPRESA", ""),
        "CNPJ":         dados.get("CNPJ", ""),
        "PREGAO":       dados.get("PREGAO", ""),
        "TIPO":         dados.get("TIPO", "Ordinário"),
        "VALOR":        valor_fmt,
        "SITUACAO":     dados.get("SITUACAO", "Pendente"),
        "ENTRADA_SALC": dados.get("ENTRADA_SALC", ""),
        "OBS":          dados.get("OBS", ""),
        "ITENS_JSON":   itens_json,
    }
    row = [linha[c] for c in COLUNAS]

    from gspread.utils import rowcol_to_a1
    end_col = rowcol_to_a1(1, len(COLUNAS)).replace("1", "")
    ws.update(
        range_name=f"A{proxima}:{end_col}{proxima}",
        values=[row],
        value_input_option="RAW",
    )
    logger.info("REQ %s adicionada na linha %d", linha["REQ"], proxima)


def atualizar_req(req_num: str, situacao: str = None, entrada_salc: str = None,
                  ne: str = None) -> None:
    ws    = _ws()
    todas = ws.get_all_values()
    if not todas:
        return
    headers = todas[0]

    def col(name):
        return headers.index(name) + 1 if name in headers else None

    col_req   = col("REQ")
    col_sit   = col("SITUACAO")
    col_ent   = col("ENTRADA_SALC")
    col_ne    = col("NE")

    from gspread.utils import rowcol_to_a1
    for i, row in enumerate(todas[1:], start=2):
        val = row[col_req - 1] if col_req and len(row) >= col_req else ""
        if str(val).strip() == str(req_num).strip():
            batch = []
            if situacao   is not None and col_sit: batch.append({"range": rowcol_to_a1(i, col_sit), "values": [[situacao]]})
            if entrada_salc is not None and col_ent: batch.append({"range": rowcol_to_a1(i, col_ent), "values": [[entrada_salc]]})
            if ne         is not None and col_ne:  batch.append({"range": rowcol_to_a1(i, col_ne),  "values": [[ne]]})
            if batch:
                ws.batch_update(batch, value_input_option="RAW")
            logger.info("REQ %s atualizada", req_num)
            return


def itens_da_req(req_num: str) -> list[dict]:
    """Retorna os itens de uma REQ (armazenados como JSON)."""
    reqs = ler_reqs()
    for r in reqs:
        if str(r.get("REQ", "")).strip() == str(req_num).strip():
            raw = r.get("ITENS_JSON", "")
            if raw:
                try:
                    return json.loads(raw)
                except Exception:
                    pass
    return []
