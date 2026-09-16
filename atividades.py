"""
Feed de atividades do Dashboard — compara o estado atual (NCs, empenhos,
REQs) com o último "retrato" salvo e gera eventos para o que mudou desde
então (NC nova, empenho realizado, REQ protocolada). O retrato e o
histórico de eventos ficam salvos em abas da própria planilha do SSAC,
para sobreviver a reinícios do app.
"""
import json
import logging
from datetime import datetime

from sheets_nc import _conectar, parse_moeda, format_moeda
from config import SHEET_ID_NC

logger = logging.getLogger(__name__)

ABA_SNAPSHOT   = "SSAC_SNAPSHOT"
ABA_ATIVIDADES = "SSAC_ATIVIDADES"
COLUNAS_ATIV   = ["DATA_HORA", "TIPO", "DESCRICAO", "REF", "VALOR"]
MAX_EVENTOS    = 200  # limite de linhas mantidas no histórico


def _ws(nome: str, colunas: list[str]):
    client   = _conectar()
    planilha = client.open_by_key(SHEET_ID_NC)
    try:
        ws = planilha.worksheet(nome)
    except Exception:
        ws = planilha.add_worksheet(title=nome, rows=1000, cols=len(colunas))
        ws.update("A1", [colunas], value_input_option="RAW")
        logger.info("Aba %s criada.", nome)
    return ws


def _ler_snapshot() -> dict:
    ws   = _ws(ABA_SNAPSHOT, ["CAMPO", "VALOR"])
    rows = ws.get_all_values()
    dados = {}
    for row in rows[1:]:
        if len(row) >= 2 and row[0]:
            try:
                dados[row[0]] = json.loads(row[1]) if row[1] else {}
            except Exception:
                dados[row[0]] = {}
    return dados


def _salvar_snapshot(snapshot: dict) -> None:
    ws = _ws(ABA_SNAPSHOT, ["CAMPO", "VALOR"])
    linhas = [["CAMPO", "VALOR"]] + [
        [campo, json.dumps(valor, ensure_ascii=False)] for campo, valor in snapshot.items()
    ]
    ws.clear()
    ws.update("A1", linhas, value_input_option="RAW")


def _registrar_eventos(eventos: list[dict]) -> None:
    ws = _ws(ABA_ATIVIDADES, COLUNAS_ATIV)
    linhas = [[e["DATA_HORA"], e["TIPO"], e["DESCRICAO"], e.get("REF", ""), e.get("VALOR", "")]
              for e in eventos]
    ws.append_rows(linhas, value_input_option="RAW")

    todas = ws.get_all_values()
    if len(todas) - 1 > MAX_EVENTOS:
        excesso = len(todas) - 1 - MAX_EVENTOS
        ws.delete_rows(2, 1 + excesso)


def ler_atividades(limite: int = 30) -> list[dict]:
    """Retorna os eventos mais recentes primeiro."""
    ws   = _ws(ABA_ATIVIDADES, COLUNAS_ATIV)
    rows = ws.get_all_values()
    if len(rows) < 2:
        return []
    headers = rows[0]
    eventos = [dict(zip(headers, row)) for row in rows[1:]]
    eventos.reverse()
    return eventos[:limite]


def _snapshot_nc(nc: dict) -> dict:
    return {
        "RECEBIDO": nc.get("RECEBIDO", ""),
        "SALDO NC": nc.get("SALDO NC", ""),
        "SITU":     nc.get("SITU", ""),
    }


def _snapshot_empenho(e: dict) -> dict:
    return {
        "SITUAÇÃO":  e.get("SITUAÇÃO", ""),
        "PAGO":      e.get("PAGO", ""),
        "LIQUIDADO": e.get("LIQUIDADO", ""),
    }


def _snapshot_req(r: dict) -> dict:
    return {"SITUACAO": r.get("SITUACAO", "")}


def _chave_req_salc(r: dict) -> str:
    return f"{r.get('Num Doc', '')}_{r.get('DATA', '')}"


def _snapshot_req_salc(r: dict) -> dict:
    return {"SITUAÇÃO": r.get("SITUAÇÃO", "")}


def sincronizar(ncs: list[dict], empenhos: list[dict], reqs: list[dict],
                 reqs_salc: list[dict] | None = None) -> list[dict]:
    """Compara o estado atual com o snapshot salvo, registra eventos para
    o que apareceu de novo desde a última sincronização e atualiza o
    snapshot. Retorna os eventos novos desta chamada (pode ser [])."""
    reqs_salc = reqs_salc or []
    anterior     = _ler_snapshot()
    primeira_vez = not anterior
    nc_ant   = anterior.get("NCS", {})
    emp_ant  = anterior.get("EMPENHOS", {})
    req_ant  = anterior.get("REQS", {})
    rsalc_ant = anterior.get("REQS_SALC", {})

    agora   = datetime.now().strftime("%d/%m/%Y %H:%M")
    eventos = []

    nc_atual = {}
    for nc in ncs:
        num = nc.get("NC", "")
        if not num:
            continue
        nc_atual[num] = _snapshot_nc(nc)
        if not primeira_vez and num not in nc_ant:
            eventos.append({
                "DATA_HORA": agora, "TIPO": "Nova NC",
                "DESCRICAO": f"NC {num} detectada — {nc.get('ORGÃO','')} · {nc.get('FINALIDADE','')[:60]}",
                "REF": num, "VALOR": nc.get("RECEBIDO", ""),
            })

    emp_atual = {}
    for e in empenhos:
        ne = e.get("NE", "")
        if not ne:
            continue
        emp_atual[ne] = _snapshot_empenho(e)
        if not primeira_vez and ne not in emp_ant:
            nc_ref = e.get("_NC_DETECTADA", "")
            valor  = (parse_moeda(e.get("A LIQUIDAR", 0)) + parse_moeda(e.get("EM LIQUIDAÇÃO", 0)) +
                      parse_moeda(e.get("LIQUIDADO", 0)) + parse_moeda(e.get("PAGO", 0)))
            desc = f"NE {ne} empenhada"
            if nc_ref:
                desc += f" — NC {nc_ref}"
            if e.get("NOME_FAV"):
                desc += f" — {e.get('NOME_FAV','')[:40]}"
            eventos.append({
                "DATA_HORA": agora, "TIPO": "Empenho realizado",
                "DESCRICAO": desc, "REF": nc_ref or ne,
                "VALOR": format_moeda(valor) if valor else "",
            })

    req_atual = {}
    for r in reqs:
        num = r.get("REQ", "")
        if not num:
            continue
        req_atual[num] = _snapshot_req(r)
        if not primeira_vez and num not in req_ant:
            eventos.append({
                "DATA_HORA": agora, "TIPO": "REQ registrada (SSAC)",
                "DESCRICAO": f"Requisição {num} registrada no SSAC — {r.get('EMPRESA','')[:40]} · NC {r.get('NC','')}",
                "REF": num, "VALOR": r.get("VALOR", ""),
            })

    rsalc_atual = {}
    for r in reqs_salc:
        chave = _chave_req_salc(r)
        if not r.get("Num Doc") and not r.get("DATA"):
            continue
        rsalc_atual[chave] = _snapshot_req_salc(r)
        if not primeira_vez and chave not in rsalc_ant:
            eventos.append({
                "DATA_HORA": agora, "TIPO": "REQ protocolada",
                "DESCRICAO": f"Requisição {r.get('Num Doc','')} protocolada pela SALC — "
                             f"{r.get('EMPRESA','')[:40]} · {r.get('SITUAÇÃO','')}",
                "REF": r.get("Num Doc", ""), "VALOR": r.get("VALOR", ""),
            })

    mudou = (nc_atual != nc_ant or emp_atual != emp_ant or req_atual != req_ant or
             rsalc_atual != rsalc_ant)
    if mudou:
        _salvar_snapshot({"NCS": nc_atual, "EMPENHOS": emp_atual, "REQS": req_atual,
                           "REQS_SALC": rsalc_atual})
    if eventos:
        _registrar_eventos(eventos)

    return eventos
