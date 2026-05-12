"""
Operações de leitura/escrita na planilha de controle de NCs e Requisições.
"""
import logging
from datetime import datetime
from typing import Optional

import gspread
from google.oauth2.service_account import Credentials

from config import SHEET_ID_NC, GOOGLE_CREDENTIALS_FILE, ABA_NCS, ABA_REQS, ABA_FORNECEDORES

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


def _conectar() -> gspread.Client:
    import os, json

    # 1. Streamlit Cloud: secret [gcp_service_account]
    try:
        import streamlit as st
        if "gcp_service_account" in st.secrets:
            logger.info("GCP via st.secrets")
            creds = Credentials.from_service_account_info(
                dict(st.secrets["gcp_service_account"]), scopes=SCOPES
            )
            return gspread.authorize(creds)
    except Exception:
        pass

    # 2. Variável GCP_CREDENTIALS_JSON (JSON completo)
    raw = os.getenv("GCP_CREDENTIALS_JSON", "").strip()
    logger.info("GCP_CREDENTIALS_JSON: %d chars", len(raw))
    if raw:
        try:
            info = json.loads(raw)
            creds = Credentials.from_service_account_info(info, scopes=SCOPES)
            logger.info("GCP autenticado via GCP_CREDENTIALS_JSON")
            return gspread.authorize(creds)
        except Exception as e:
            logger.error("Erro GCP_CREDENTIALS_JSON: %s", e)

    # 3. Campos individuais (private_key, client_email, ...)
    pk = os.getenv("private_key", "").strip()
    ce = os.getenv("client_email", "").strip()
    logger.info("Campos individuais: pk=%s ce=%s", bool(pk), bool(ce))
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
            creds = Credentials.from_service_account_info(info, scopes=SCOPES)
            logger.info("GCP autenticado via campos individuais")
            return gspread.authorize(creds)
        except Exception as e:
            logger.error("Erro campos individuais: %s", e)

    # 4. Arquivo JSON local
    logger.warning("GCP: fallback para arquivo %s", GOOGLE_CREDENTIALS_FILE)
    creds = Credentials.from_service_account_file(GOOGLE_CREDENTIALS_FILE, scopes=SCOPES)
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
