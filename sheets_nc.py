"""
Operações de leitura/escrita na planilha de controle de NCs e Requisições.
"""
import logging
from datetime import datetime
from typing import Optional

import gspread
from google.oauth2.service_account import Credentials

from config import SHEET_ID_NC, GOOGLE_CREDENTIALS_FILE, ABA_NCS, ABA_REQS, ABA_FORNECEDORES, DRIVE_FOLDER_ID

logger = logging.getLogger(__name__)

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

COLUNAS_NC = [
    "ORD", "UG", "SITU", "ORGÃO", "DATA NC", "NC", "PI", "ND", "PTRES",
    "OM", "PRAZO", "FINALIDADE", "OP", "DIAS", "RECEBIDO", "RECOLHIDO",
    "SALDO NC", "EMPENHADO", "EMP %", "EM TELA", "EM TELA %",
    "TIPO SALDO", "SOL RECOL", "DOC SOL RCLH", "SITUAÇÃO",
]

COLUNAS_REQ = [
    "REQ", "NE", "DATA REQ", "NC", "PI", "FINALIDADE", "TIPO",
    "EMPRESA", "DESCRIÇÃO", "VALOR", "SITUAÇÃO", "ENTRADA NA BDA", "OBS", "ARQUIVO REQ",
]


def _get_credenciais() -> Credentials:
    """Retorna credenciais Google (service account) — usada por Sheets e Drive."""
    import os, json

    try:
        import streamlit as st
        if "gcp_service_account" in st.secrets:
            return Credentials.from_service_account_info(
                dict(st.secrets["gcp_service_account"]), scopes=SCOPES
            )
    except Exception:
        pass

    raw = os.getenv("GCP_CREDENTIALS_JSON", "").strip()
    if raw:
        try:
            return Credentials.from_service_account_info(json.loads(raw), scopes=SCOPES)
        except Exception as e:
            logger.error("Erro GCP_CREDENTIALS_JSON: %s", e)

    pk = os.getenv("private_key", "").strip()
    ce = os.getenv("client_email", "").strip()
    if pk and ce:
        try:
            info = {
                "type": "service_account",
                "project_id":                  os.getenv("project_id", ""),
                "private_key_id":              os.getenv("private_key_id", ""),
                "private_key":                 pk.replace("\\n", "\n"),
                "client_email":                ce,
                "client_id":                   os.getenv("client_id", ""),
                "auth_uri":                    "https://accounts.google.com/o/oauth2/auth",
                "token_uri":                   "https://oauth2.googleapis.com/token",
                "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",
                "client_x509_cert_url":        os.getenv("client_x509_cert_url", ""),
            }
            return Credentials.from_service_account_info(info, scopes=SCOPES)
        except Exception as e:
            logger.error("Erro campos individuais: %s", e)

    return Credentials.from_service_account_file(GOOGLE_CREDENTIALS_FILE, scopes=SCOPES)


def _conectar() -> gspread.Client:
    creds = _get_credenciais()
    logger.info("GCP autenticado")
    return gspread.authorize(creds)


def parse_moeda(valor_str) -> float:
    try:
        s = str(valor_str).replace("R$", "").replace(" ", "").replace(".", "").replace(",", ".")
        return float(s)
    except (ValueError, TypeError):
        return 0.0


def format_moeda(valor: float) -> str:
    s = f"{valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return f"R$ {s}"


def _calcular_dias(data_nc_str: str) -> int:
    try:
        data_nc = datetime.strptime(data_nc_str, "%d/%m/%Y")
        return (datetime.today() - data_nc).days
    except Exception:
        return 0


def listar_abas() -> list[str]:
    try:
        client = _conectar()
        planilha = client.open_by_key(SHEET_ID_NC)
        return [ws.title for ws in planilha.worksheets()]
    except Exception as e:
        logger.error("Erro ao listar abas: %s", e)
        return []


def _ws_para_dicts(ws) -> list[dict]:
    """Converte worksheet em lista de dicts, tolerando cabeçalhos duplicados."""
    valores = ws.get_all_values()
    if not valores:
        return []
    headers = valores[0]
    # Desambigua duplicatas: col, col_2, col_3, ...
    vistos: dict[str, int] = {}
    headers_unicos = []
    for h in headers:
        if h in vistos:
            vistos[h] += 1
            headers_unicos.append(f"{h}_{vistos[h]}")
        else:
            vistos[h] = 1
            headers_unicos.append(h)

    registros = []
    for row in valores[1:]:
        row_padded = row + [""] * (len(headers_unicos) - len(row))
        registros.append(dict(zip(headers_unicos, row_padded)))
    return registros


def ler_ncs() -> list[dict]:
    client = _conectar()
    planilha = client.open_by_key(SHEET_ID_NC)
    ws = planilha.worksheet(ABA_NCS)
    registros = _ws_para_dicts(ws)
    logger.info("%d NCs lidas.", len(registros))
    return [r for r in registros if r.get("NC")]


def ler_reqs() -> list[dict]:
    client = _conectar()
    planilha = client.open_by_key(SHEET_ID_NC)
    ws = planilha.worksheet(ABA_REQS)
    registros = _ws_para_dicts(ws)
    logger.info("%d registros brutos na aba REQs.", len(registros))
    if registros:
        logger.info("Colunas encontradas: %s", list(registros[0].keys()))
        logger.info("Primeira linha: %s", registros[0])
    resultado = [r for r in registros if any(v.strip() for v in r.values() if v)]
    logger.info("%d REQs não-vazias retornadas.", len(resultado))
    return resultado


def recalcular_empenhados(reqs: list[dict]) -> dict[str, str]:
    """
    Recalcula EMPENHADO de TODAS as NCs somando REQs com situação Empenhada.
    Retorna dict {nc_num: mensagem} com o resultado de cada NC.
    """
    from collections import defaultdict
    totais: dict[str, float] = defaultdict(float)
    for r in reqs:
        nc    = r.get("NC", "")
        valor = parse_moeda(r.get("VALOR", 0))
        sit   = r.get("SITUAÇÃO", "")
        if sit == "Empenhada" and nc:
            totais[nc] += valor
        elif sit == "Anulado" and nc:
            totais[nc] -= valor  # anulação devolve ao saldo

    resultados = {}
    for nc_num, total in totais.items():
        try:
            _set_nc_empenhado(nc_num, total)
            resultados[nc_num] = f"✅ {format_moeda(total)}"
        except Exception as e:
            resultados[nc_num] = f"⚠️ {e}"
    return resultados


def _set_nc_empenhado(nc_num: str, total_empenhado: float) -> None:
    """Define (sobrescreve) o EMPENHADO de uma NC com o valor informado."""
    client = _conectar()
    planilha = client.open_by_key(SHEET_ID_NC)
    ws = planilha.worksheet(ABA_NCS)
    todas = ws.get_all_values()
    if not todas:
        return
    headers = todas[0]

    def _col(name):
        return headers.index(name) + 1 if name in headers else None

    col_nc   = _col("NC")
    col_emp  = _col("EMPENHADO")
    col_pemp = _col("EMP %")
    col_rec  = _col("RECEBIDO")

    if not col_nc or not col_emp:
        raise ValueError("Colunas NC ou EMPENHADO não encontradas.")

    formula_cols: set[str] = set()
    if len(todas) > 1:
        try:
            formula_row = ws.row_values(2, value_render_option="FORMULA")
            formula_cols = {headers[i] for i, v in enumerate(formula_row)
                            if i < len(headers) and str(v).startswith("=")}
        except Exception:
            pass

    if "EMPENHADO" in formula_cols:
        raise ValueError("Coluna EMPENHADO é fórmula — será atualizada automaticamente.")

    col_situ     = _col("SITU")
    col_saldo    = _col("SALDO NC")
    col_emtela   = _col("EM TELA")
    col_emtelap  = _col("EM TELA %")

    from gspread.utils import rowcol_to_a1
    for i, row in enumerate(todas[1:], start=2):
        val_nc = row[col_nc - 1] if len(row) >= col_nc else ""
        if str(val_nc).strip() == str(nc_num).strip():
            rec        = parse_moeda(row[col_rec - 1]) if col_rec and len(row) >= col_rec else 0.0
            pct_num    = round((total_empenhado / rec * 100), 2) if rec else 0.0
            saldo      = round(rec - total_empenhado, 2)
            totalmente = rec > 0 and round(pct_num) >= 100

            # Escreve EMPENHADO como NÚMERO (não texto) para fórmulas funcionarem
            batch = [{"range": rowcol_to_a1(i, col_emp), "values": [[total_empenhado]]}]

            # EMP % como número 0-100 (sobrescreve fórmula que falha com texto)
            if col_pemp:
                batch.append({"range": rowcol_to_a1(i, col_pemp),
                               "values": [[f"{pct_num:.2f}%".replace(".", ",")]]})

            if totalmente:
                # SALDO NC = 0
                if col_saldo and "SALDO NC" not in formula_cols:
                    batch.append({"range": rowcol_to_a1(i, col_saldo),
                                   "values": [[format_moeda(saldo)]]})
                # EM TELA e EM TELA % = 0
                if col_emtela and "EM TELA" not in formula_cols:
                    batch.append({"range": rowcol_to_a1(i, col_emtela),
                                   "values": [[format_moeda(0)]]})
                if col_emtelap and "EM TELA %" not in formula_cols:
                    batch.append({"range": rowcol_to_a1(i, col_emtelap),
                                   "values": [["0,00%"]]})
                # SITU → OK
                if col_situ and "SITU" not in formula_cols:
                    batch.append({"range": rowcol_to_a1(i, col_situ), "values": [["OK"]]})

            ws.batch_update(batch, value_input_option="USER_ENTERED")
            logger.info("NC %s: EMPENHADO=%s EMP%%=%s SITU=%s",
                        nc_num, total_empenhado, pct_num, "OK" if totalmente else "-")
            return
    raise ValueError(f"NC '{nc_num}' não encontrada.")


def atualizar_nc_campos(nc_num: str, finalidade: str, data_nc: str, prazo: str) -> None:
    """Atualiza FINALIDADE, DATA NC e PRAZO de uma NC."""
    client = _conectar()
    planilha = client.open_by_key(SHEET_ID_NC)
    ws = planilha.worksheet(ABA_NCS)
    todas = ws.get_all_values()
    if not todas:
        raise ValueError("Aba NCs vazia.")
    headers = todas[0]

    def _col(name):
        return headers.index(name) + 1 if name in headers else None

    col_nc   = _col("NC")
    col_fin  = _col("FINALIDADE")
    col_data = _col("DATA NC")
    col_prz  = _col("PRAZO")

    if not col_nc:
        raise ValueError("Coluna NC não encontrada.")

    from gspread.utils import rowcol_to_a1
    for i, row in enumerate(todas[1:], start=2):
        val = row[col_nc - 1] if len(row) >= col_nc else ""
        if str(val).strip() == str(nc_num).strip():
            batch = []
            if col_fin:  batch.append({"range": rowcol_to_a1(i, col_fin),  "values": [[finalidade]]})
            if col_data: batch.append({"range": rowcol_to_a1(i, col_data), "values": [[data_nc]]})
            if col_prz:  batch.append({"range": rowcol_to_a1(i, col_prz),  "values": [[prazo]]})
            if batch:
                ws.batch_update(batch, value_input_option="USER_ENTERED")
            logger.info("NC %s atualizada: finalidade/data/prazo", nc_num)
            return
    raise ValueError(f"NC '{nc_num}' não encontrada.")


def atualizar_nc_empenhado(nc_num: str, valor_req: float) -> None:
    """Soma valor_req ao EMPENHADO da NC. SALDO NC é atualizado pela fórmula da planilha."""
    client = _conectar()
    planilha = client.open_by_key(SHEET_ID_NC)
    ws = planilha.worksheet(ABA_NCS)
    todas = ws.get_all_values()
    if not todas:
        return
    headers = todas[0]

    def _col(name):
        return headers.index(name) + 1 if name in headers else None

    col_nc   = _col("NC")
    col_emp  = _col("EMPENHADO")
    col_pemp = _col("EMP %")
    col_rec  = _col("RECEBIDO")

    if not col_nc or not col_emp:
        raise ValueError("Colunas NC ou EMPENHADO não encontradas.")

    # Detecta se EMPENHADO é fórmula (não pode sobrescrever)
    formula_cols: set[str] = set()
    if len(todas) > 1:
        try:
            formula_row = ws.row_values(2, value_render_option="FORMULA")
            formula_cols = {headers[i] for i, v in enumerate(formula_row)
                            if i < len(headers) and str(v).startswith("=")}
        except Exception:
            pass

    if "EMPENHADO" in formula_cols:
        raise ValueError("Coluna EMPENHADO é calculada por fórmula — saldo será atualizado automaticamente.")

    from gspread.utils import rowcol_to_a1
    for i, row in enumerate(todas[1:], start=2):
        val_nc = row[col_nc - 1] if len(row) >= col_nc else ""
        if str(val_nc).strip() == str(nc_num).strip():
            emp_atual = parse_moeda(row[col_emp - 1]) if len(row) >= col_emp else 0.0
            novo_emp  = max(0.0, round(emp_atual + valor_req, 2))  # negativo = anulação
            rec       = parse_moeda(row[col_rec - 1]) if col_rec and len(row) >= col_rec else 0.0
            pct       = f"{(novo_emp / rec * 100):.2f}%" if rec else "0,00%"

            batch = [{"range": rowcol_to_a1(i, col_emp), "values": [[format_moeda(novo_emp)]]}]
            if col_pemp and "EMP %" not in formula_cols:
                batch.append({"range": rowcol_to_a1(i, col_pemp), "values": [[pct]]})
            ws.batch_update(batch, value_input_option="USER_ENTERED")
            logger.info("NC %s: EMPENHADO atualizado de %s para %s", nc_num, emp_atual, novo_emp)
            return
    raise ValueError(f"NC '{nc_num}' não encontrada na aba NCs.")


def atualizar_req(req_num: str, nova_situacao: str, nova_entrada: str, novo_ne: str = "") -> None:
    """Atualiza SITUAÇÃO, ENTRADA NA BDA e NE de uma REQ pelo número."""
    client = _conectar()
    planilha = client.open_by_key(SHEET_ID_NC)
    ws = planilha.worksheet(ABA_REQS)
    todas = ws.get_all_values()
    if not todas:
        raise ValueError("Aba REQs vazia.")
    headers = todas[0]

    def _col(name):
        return headers.index(name) + 1 if name in headers else None

    col_req = _col("REQ")
    col_sit = _col("SITUAÇÃO")
    col_ent = _col("ENTRADA NA BDA")
    col_ne  = _col("NE")

    if col_req is None:
        raise ValueError("Coluna REQ não encontrada na aba.")

    from gspread.utils import rowcol_to_a1
    for i, row in enumerate(todas[1:], start=2):
        val_req = row[col_req - 1] if len(row) >= col_req else ""
        if str(val_req).strip() == str(req_num).strip():
            batch = []
            if col_sit: batch.append({"range": rowcol_to_a1(i, col_sit), "values": [[nova_situacao]]})
            if col_ent: batch.append({"range": rowcol_to_a1(i, col_ent), "values": [[nova_entrada]]})
            if col_ne:  batch.append({"range": rowcol_to_a1(i, col_ne),  "values": [[novo_ne]]})
            if batch:
                ws.batch_update(batch, value_input_option="USER_ENTERED")
            logger.info("REQ %s atualizada: sit=%s entrada=%s ne=%s", req_num, nova_situacao, nova_entrada, novo_ne)
            return
    raise ValueError(f"REQ '{req_num}' não encontrada na planilha.")


def ler_fornecedores() -> list[dict]:
    try:
        client = _conectar()
        planilha = client.open_by_key(SHEET_ID_NC)
        ws = planilha.worksheet(ABA_FORNECEDORES)
        registros = _ws_para_dicts(ws)
        return [r for r in registros if any(r.values())]
    except Exception as e:
        logger.warning("Aba fornecedores não encontrada: %s", e)
        return []


def adicionar_nc(dados: dict) -> None:
    client = _conectar()
    planilha = client.open_by_key(SHEET_ID_NC)
    ws = planilha.worksheet(ABA_NCS)

    ncs = _ws_para_dicts(ws)
    proximo_ord = max(
        (int(str(nc.get("ORD", 0))) for nc in ncs if str(nc.get("ORD", "")).isdigit()),
        default=0,
    ) + 1

    valor = dados.get("RECEBIDO", 0.0)
    valor_fmt = format_moeda(float(valor)) if isinstance(valor, (int, float)) else str(valor)

    linha = {
        "ORD": proximo_ord,
        "UG": dados.get("UG", "160482"),
        "SITU": dados.get("SITU", "OK"),
        "ORGÃO": dados.get("ORGÃO", ""),
        "DATA NC": dados.get("DATA NC", ""),
        "NC": dados.get("NC", ""),
        "PI": dados.get("PI", ""),
        "ND": dados.get("ND", ""),
        "PTRES": dados.get("PTRES", ""),
        "OM": dados.get("OM", "10º GAC Sl"),
        "PRAZO": dados.get("PRAZO", ""),
        "FINALIDADE": dados.get("FINALIDADE", ""),
        "OP": dados.get("OP", ""),
        "DIAS": _calcular_dias(dados.get("DATA NC", "")),
        "RECEBIDO": valor_fmt,
        "RECOLHIDO": "",
        "SALDO NC": valor_fmt,
        "EMPENHADO": "R$ 0,00",
        "EMP %": "0,00%",
        "EM TELA": "R$ 0,00",
        "EM TELA %": "0,00%",
        "TIPO SALDO": "",
        "SOL RECOL": "FALSE",
        "DOC SOL RCLH": "",
        "SITUAÇÃO": dados.get("SITUAÇÃO", ""),
    }

    row_values = [str(linha.get(col, "")) for col in COLUNAS_NC]
    ws.append_row(row_values, value_input_option="USER_ENTERED")
    logger.info("NC adicionada: %s", dados.get("NC", ""))


def adicionar_req(dados: dict) -> None:
    client = _conectar()
    planilha = client.open_by_key(SHEET_ID_NC)
    ws = planilha.worksheet(ABA_REQS)

    reqs = _ws_para_dicts(ws)
    proximo_req = max(
        (int(str(r.get("REQ", 0))) for r in reqs if str(r.get("REQ", "")).isdigit()),
        default=0,
    ) + 1

    valor = dados.get("VALOR", 0.0)
    valor_fmt = format_moeda(float(valor)) if isinstance(valor, (int, float)) else str(valor)

    linha = {
        "REQ": dados.get("REQ") or proximo_req,
        "NE": dados.get("NE", ""),
        "DATA REQ": dados.get("DATA REQ", datetime.today().strftime("%d/%m/%Y")),
        "NC": dados.get("NC", ""),
        "PI": dados.get("PI", ""),
        "FINALIDADE": dados.get("FINALIDADE", ""),
        "TIPO": dados.get("TIPO", "Ordinário"),
        "EMPRESA": dados.get("EMPRESA", ""),
        "DESCRIÇÃO": dados.get("DESCRIÇÃO", ""),
        "VALOR": valor_fmt,
        "SITUAÇÃO": dados.get("SITUAÇÃO", "Pendente"),
        "ENTRADA NA BDA": dados.get("ENTRADA NA BDA", ""),
        "OBS": dados.get("OBS", ""),
        "ARQUIVO REQ": dados.get("ARQUIVO REQ", ""),
    }

    row_values = [str(linha.get(col, "")) for col in COLUNAS_REQ]

    # Detecta colunas com fórmula para não sobrescrevê-las (causaria erro 500)
    headers = ws.row_values(1)
    formula_cols: set[str] = set()
    todas = ws.get_all_values()
    if len(todas) > 1:
        try:
            formula_row = ws.row_values(2, value_render_option="FORMULA")
            formula_cols = {
                headers[i]
                for i, v in enumerate(formula_row)
                if i < len(headers) and str(v).startswith("=")
            }
            logger.info("Colunas com fórmula (ignoradas): %s", formula_cols)
        except Exception as e:
            logger.warning("Não detectou fórmulas: %s", e)

    # Próxima linha vazia
    proxima_linha = len(todas) + 1
    if proxima_linha > ws.row_count:
        ws.add_rows(max(100, proxima_linha - ws.row_count + 10))

    # Escreve somente nas colunas sem fórmula
    from gspread.utils import rowcol_to_a1
    batch = []
    for i, col_name in enumerate(COLUNAS_REQ):
        if col_name in formula_cols:
            continue
        col_idx = (headers.index(col_name) + 1) if col_name in headers else (i + 1)
        cell = rowcol_to_a1(proxima_linha, col_idx)
        batch.append({"range": cell, "values": [[row_values[i]]]})

    if batch:
        ws.batch_update(batch, value_input_option="USER_ENTERED")
    logger.info("REQ adicionada na linha %d (pulou %d colunas com fórmula)",
                proxima_linha, len(formula_cols))


# ── Frases padrão ─────────────────────────────────────────────────────────────
ABA_FRASES = "Frases"


def ler_frases(tipo: str) -> list[str]:
    """Retorna lista de textos de frases cadastradas para o tipo (INTRO / JUST)."""
    try:
        client = _conectar()
        planilha = client.open_by_key(SHEET_ID_NC)
        ws = planilha.worksheet(ABA_FRASES)
        registros = _ws_para_dicts(ws)
        return [r["TEXTO"] for r in registros
                if r.get("TEXTO") and r.get("TIPO", "").upper() == tipo.upper()]
    except Exception:
        return []  # aba Frases ainda não existe — sem aviso


def adicionar_frase(tipo: str, texto: str) -> None:
    """Adiciona uma frase na aba Frases, criando a aba se não existir."""
    client = _conectar()
    planilha = client.open_by_key(SHEET_ID_NC)
    try:
        ws = planilha.worksheet(ABA_FRASES)
    except Exception:
        ws = planilha.add_worksheet(title=ABA_FRASES, rows=200, cols=2)
        ws.append_row(["TIPO", "TEXTO"])
    ws.append_row([tipo.upper(), texto])
    logger.info("Frase adicionada: %s", tipo)


def salvar_html_req(req_num: str, html: str) -> None:
    """Salva o HTML da REQ na aba ARQUIVO_HTML da planilha."""
    client   = _conectar()
    planilha = client.open_by_key(SHEET_ID_NC)
    try:
        ws = planilha.worksheet("ARQUIVO_HTML")
    except Exception:
        ws = planilha.add_worksheet(title="ARQUIVO_HTML", rows=500, cols=2)
        ws.update("A1:B1", [["REQ", "HTML"]], value_input_option="RAW")

    todos = ws.get_all_values()

    # Remove versão anterior se existir
    for i, row in enumerate(todos[1:], start=2):
        if row and str(row[0]).strip() == str(req_num).strip():
            ws.delete_rows(i)
            todos = ws.get_all_values()  # recarrega após deleção
            break

    # Limita a 40k chars para não exceder limite de célula do Sheets
    html_salvo = html[:40000] if len(html) > 40000 else html

    proxima = len(todos) + 1
    if proxima > ws.row_count:
        ws.add_rows(100)

    from gspread.utils import rowcol_to_a1
    ws.update(
        range_name=f"A{proxima}:B{proxima}",
        values=[[str(req_num), html_salvo]],
        value_input_option="RAW",
    )
    logger.info("HTML REQ %s salvo (linha %d, %d chars)", req_num, proxima, len(html_salvo))


def ler_html_req(req_num: str) -> str:
    """Lê o HTML de uma REQ salvo na planilha."""
    client   = _conectar()
    planilha = client.open_by_key(SHEET_ID_NC)
    try:
        ws   = planilha.worksheet("ARQUIVO_HTML")
        todos = ws.get_all_values()
        for row in todos[1:]:
            if row and str(row[0]).strip() == str(req_num).strip():
                return row[1] if len(row) > 1 else ""
    except Exception as e:
        logger.warning("Erro ao ler HTML da REQ %s: %s", req_num, e)
    return ""


def listar_reqs_com_html() -> list[str]:
    """Retorna lista de números de REQ que têm HTML salvo."""
    client   = _conectar()
    planilha = client.open_by_key(SHEET_ID_NC)
    try:
        ws   = planilha.worksheet("ARQUIVO_HTML")
        todos = ws.get_all_values()
        return [row[0] for row in todos[1:] if row and row[0]]
    except Exception:
        return []


def excluir_frase(tipo: str, texto: str) -> None:
    """Remove a frase que corresponde exatamente ao tipo e texto."""
    try:
        client = _conectar()
        planilha = client.open_by_key(SHEET_ID_NC)
        ws = planilha.worksheet(ABA_FRASES)
        todos = ws.get_all_values()
        for i, row in enumerate(todos):
            if len(row) >= 2 and row[0].upper() == tipo.upper() and row[1] == texto:
                ws.delete_rows(i + 1)
                return
    except Exception as e:
        logger.warning("Erro ao excluir frase: %s", e)


# ── Google Drive ───────────────────────────────────────────────────────────────
_DRIVE_FOLDER_NAME  = "SSAC - Requisições"
_DRIVE_SHARE_EMAILS = ["sac10gacsl@gmail.com", "augustumourasantos@gmail.com"]


def _get_drive_service():
    """Retorna cliente Drive usando as mesmas credenciais do Sheets."""
    from googleapiclient.discovery import build
    return build("drive", "v3", credentials=_get_credenciais())


def _compartilhar(service, file_id: str) -> None:
    """Compartilha arquivo/pasta com os emails configurados."""
    # Acesso público por link
    try:
        service.permissions().create(
            fileId=file_id,
            body={"type": "anyone", "role": "reader"},
        ).execute()
    except Exception:
        pass
    # Acesso individual (editor) para os emails
    for email in _DRIVE_SHARE_EMAILS:
        try:
            service.permissions().create(
                fileId=file_id,
                body={"type": "user", "role": "writer", "emailAddress": email},
                sendNotificationEmail=False,
            ).execute()
            logger.info("Compartilhado com %s", email)
        except Exception as e:
            logger.warning("Não foi possível compartilhar com %s: %s", email, e)


def _get_ou_criar_pasta(service) -> str:
    """
    Retorna o ID da pasta no Drive onde salvar as REQs.
    Usa DRIVE_FOLDER_ID da config se definido (pasta no Drive do usuário).
    Caso contrário, tenta criar — mas service accounts precisam de pasta do usuário.
    """
    if DRIVE_FOLDER_ID:
        logger.info("Usando pasta Drive configurada: %s", DRIVE_FOLDER_ID)
        return DRIVE_FOLDER_ID
    raise ValueError(
        "Configure DRIVE_FOLDER_ID no Railway: crie uma pasta no seu Google Drive, "
        f"compartilhe com agentes@absolute-disk-428101-n6.iam.gserviceaccount.com (Editor) "
        "e cole o ID da pasta na variável DRIVE_FOLDER_ID."
    )


def salvar_arquivo_no_drive(conteudo: bytes, nome: str, mimetype: str) -> str:
    """Salva qualquer arquivo no Drive e retorna o link de visualização."""
    from googleapiclient.http import MediaInMemoryUpload
    service   = _get_drive_service()
    folder_id = _get_ou_criar_pasta(service)

    # Remove versão anterior
    q = f"name='{nome}' and '{folder_id}' in parents and trashed=false"
    for arq in service.files().list(q=q, fields="files(id)").execute().get("files", []):
        service.files().delete(fileId=arq["id"]).execute()

    meta  = {"name": nome, "parents": [folder_id], "mimeType": mimetype}
    media = MediaInMemoryUpload(conteudo, mimetype=mimetype)
    arq   = service.files().create(body=meta, media_body=media, fields="id").execute()
    _compartilhar(service, arq["id"])
    link = f"https://drive.google.com/file/d/{arq['id']}/view"
    logger.info("Arquivo %s salvo no Drive: %s", nome, link)
    return link


def salvar_req_no_drive(html_content: str, req_num: str) -> str:
    """
    Salva o HTML da REQ no Google Drive e retorna o link de visualização.
    Cria a pasta 'SSAC - Requisições' se não existir.
    """
    from googleapiclient.http import MediaInMemoryUpload
    try:
        service   = _get_drive_service()
        folder_id = _get_ou_criar_pasta(service)
        nome      = f"REQ_{req_num}.html"

        # Remove versão anterior se existir
        q = f"name='{nome}' and '{folder_id}' in parents and trashed=false"
        antigos = service.files().list(q=q, fields="files(id)").execute().get("files", [])
        for arq in antigos:
            service.files().delete(fileId=arq["id"]).execute()

        # Cria o arquivo
        meta  = {"name": nome, "parents": [folder_id], "mimeType": "text/html"}
        media = MediaInMemoryUpload(html_content.encode("utf-8"), mimetype="text/html")
        arq  = service.files().create(body=meta, media_body=media, fields="id").execute()
        _compartilhar(service, arq["id"])
        link = f"https://drive.google.com/file/d/{arq['id']}/view"
        logger.info("REQ %s salva no Drive: %s", req_num, link)
        return link
    except Exception as e:
        logger.error("Erro ao salvar no Drive: %s", e)
        raise
