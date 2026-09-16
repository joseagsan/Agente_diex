"""
CRUD limpo para a aba SSAC_REQS — sem fórmulas, sem conflito.
Colunas: REQ | DATA | NC | NE | PI | ND | EMPRESA | CNPJ | PREGAO
         TIPO | VALOR | SITUACAO | ENTRADA_SALC | OBS | ITENS_JSON
         + colunas da Fase 2 (pós-empenho, ver FASE2_COLUNAS)
"""
import json
import logging
from datetime import datetime

from sheets_nc import _conectar, parse_moeda, format_moeda
from config import SHEET_ID_NC, ABA_SSAC_REQS

logger = logging.getLogger(__name__)

COLUNAS_BASE = [
    "REQ", "DATA", "NC", "NE", "PI", "ND",
    "EMPRESA", "CNPJ", "PREGAO", "TIPO",
    "VALOR", "SITUACAO", "ENTRADA_SALC", "OBS", "ITENS_JSON",
]

# Fase 2 — prosseguimento da contratação pelo requisitante (após empenho),
# conforme fluxograma SPED: envio da NE, recebimento da NF, ateste,
# espelho SISCOFIS (material) e minutas de despacho (Cmt/Fisc Adm/OD).
FASE2_COLUNAS = [
    "DATA_ENVIO_NE", "ANEXO_COMPROVANTE",
    "TIPO_NF", "NUM_NF", "DATA_NF", "ANEXO_NF",
    "DATA_ATESTE", "ANEXO_ESPELHO",
    "DESPACHOS_JSON",
]

# Acompanhamento do processo enquanto a REQ ainda está Pendente (antes do
# empenho) — onde o papel está parado na cadeia de despachos do SPED.
# Depois de Empenhada, o acompanhamento passa a ser a própria Fase 2 acima.
ETAPAS_REQ_PROCESSO = [
    "Requisição redigida (a autuar)",
    "Aguardando despacho do Cmt",
    "Aguardando despacho do Fiscal Administrativo",
    "Aguardando despacho do OD",
    "Na SALC (documentação/certidões)",
    "Aguardando emissão da NE",
]
EXTRA_COLUNAS = ["ETAPA_PROCESSO"]

COLUNAS = COLUNAS_BASE + FASE2_COLUNAS + EXTRA_COLUNAS
_COLUNAS_PRESERVADAS = FASE2_COLUNAS + EXTRA_COLUNAS


def _garantir_colunas(ws) -> list[str]:
    """Garante que o cabeçalho da aba tem todas as COLUNAS (adiciona as que faltarem
    ao final, sem tocar nas existentes). Retorna o cabeçalho atualizado."""
    headers = ws.row_values(1)
    faltantes = [c for c in COLUNAS if c not in headers]
    if faltantes:
        if ws.col_count < len(headers) + len(faltantes):
            ws.add_cols(len(headers) + len(faltantes) - ws.col_count)
        from gspread.utils import rowcol_to_a1
        inicio = len(headers) + 1
        fim = len(headers) + len(faltantes)
        ws.update(
            range_name=f"{rowcol_to_a1(1, inicio)}:{rowcol_to_a1(1, fim)}",
            values=[faltantes],
            value_input_option="RAW",
        )
        headers = headers + faltantes
        logger.info("Colunas adicionadas à aba %s: %s", ABA_SSAC_REQS, faltantes)
    return headers


def _ws():
    client   = _conectar()
    planilha = client.open_by_key(SHEET_ID_NC)
    try:
        ws = planilha.worksheet(ABA_SSAC_REQS)
    except Exception:
        ws = planilha.add_worksheet(title=ABA_SSAC_REQS, rows=1000, cols=len(COLUNAS))
        ws.update("A1", [COLUNAS], value_input_option="RAW")
        logger.info("Aba %s criada.", ABA_SSAC_REQS)
        return ws
    _garantir_colunas(ws)
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
    row = [linha.get(c, "") for c in COLUNAS]

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


def excluir_req(req_num: str) -> None:
    """Remove a linha da REQ na aba SSAC_REQS."""
    ws    = _ws()
    todos = ws.get_all_values()
    if not todos:
        return
    headers  = todos[0]
    col_req  = headers.index("REQ") + 1 if "REQ" in headers else 1
    for i, row in enumerate(todos[1:], start=2):
        val = row[col_req - 1] if len(row) >= col_req else ""
        if str(val).strip() == str(req_num).strip():
            ws.delete_rows(i)
            logger.info("REQ %s excluída.", req_num)
            return


def editar_req(req_num: str, dados: dict) -> None:
    """Atualiza os campos básicos de uma REQ existente.
    Preserva os campos da Fase 2 (FASE2_COLUNAS) já gravados, a menos que
    `dados` os informe explicitamente."""
    ws    = _ws()
    todos = ws.get_all_values()
    if not todos:
        return
    headers = todos[0]

    def col(name):
        return headers.index(name) + 1 if name in headers else None

    col_req = col("REQ")
    if not col_req:
        return

    from gspread.utils import rowcol_to_a1
    for i, row in enumerate(todos[1:], start=2):
        val = row[col_req - 1] if len(row) >= col_req else ""
        if str(val).strip() == str(req_num).strip():
            row_atual = row + [""] * (len(headers) - len(row))
            atual     = dict(zip(headers, row_atual))

            valor = dados.get("VALOR", 0.0)
            valor_fmt = format_moeda(float(valor)) if isinstance(valor, (int, float)) else str(valor)
            import json as _json
            itens     = dados.get("ITENS", [])
            linha = {
                "REQ":          str(req_num),
                "DATA":         dados.get("DATA", ""),
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
                "ITENS_JSON":   _json.dumps(itens, ensure_ascii=False) if itens else "",
            }
            for campo in _COLUNAS_PRESERVADAS:
                linha[campo] = dados[campo] if campo in dados else atual.get(campo, "")

            row_vals = [linha.get(c, "") for c in COLUNAS]
            end_col  = rowcol_to_a1(1, len(COLUNAS)).replace("1", "")
            ws.update(range_name=f"A{i}:{end_col}{i}", values=[row_vals],
                      value_input_option="RAW")
            logger.info("REQ %s editada.", req_num)
            return


def atualizar_fase2(req_num: str, campos: dict) -> None:
    """Atualiza campos de acompanhamento de uma REQ (Fase 2 pós-empenho ou
    ETAPA_PROCESSO pré-empenho), sem tocar nos demais. `campos` deve conter
    apenas chaves de FASE2_COLUNAS/EXTRA_COLUNAS."""
    ws    = _ws()
    todos = ws.get_all_values()
    if not todos:
        raise ValueError("Aba SSAC_REQS vazia.")
    headers = todos[0]

    def col(name):
        return headers.index(name) + 1 if name in headers else None

    col_req = col("REQ")
    if not col_req:
        raise ValueError("Coluna REQ não encontrada.")

    from gspread.utils import rowcol_to_a1
    for i, row in enumerate(todos[1:], start=2):
        val = row[col_req - 1] if len(row) >= col_req else ""
        if str(val).strip() == str(req_num).strip():
            batch = []
            for campo, valor in campos.items():
                if campo not in _COLUNAS_PRESERVADAS:
                    continue
                c = col(campo)
                if c:
                    batch.append({"range": rowcol_to_a1(i, c), "values": [[valor]]})
            if batch:
                ws.batch_update(batch, value_input_option="RAW")
            logger.info("REQ %s: acompanhamento atualizado (%s)", req_num, list(campos.keys()))
            return
    raise ValueError(f"REQ '{req_num}' não encontrada.")


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
