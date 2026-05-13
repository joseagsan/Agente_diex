"""
Pesquisa de itens via contratos.comprasnet.gov.br
Fluxo:
  1. GET  /empenho/buscacompra                    → CSRF token
  2. POST /empenho/buscacompra                    → redirect a /empenho/fornecedor/{compra_id}
  3. GET  /empenho/fornecedor/{compra_id}?draw=1  → lista fornecedores (JSON DataTables)
  4. GET  /empenho/item/{compra_id}/{forn_id}?draw=1 → itens do fornecedor (JSON DataTables)

Requer cookies de contratos.comprasnet.gov.br no session_json.
"""
import json
import logging
import os
import re
import requests
from bs4 import BeautifulSoup

logger = logging.getLogger(__name__)
BASE = "https://contratos.comprasnet.gov.br"


# ── Sessão autenticada ─────────────────────────────────────────────────────────

def _carregar_cookies() -> list:
    cands = [
        os.path.join(os.path.dirname(__file__), "session.json"),
        "session.json",
    ]
    for c in cands:
        if os.path.exists(c):
            with open(c) as f:
                return json.load(f).get("cookies", [])

    try:
        import streamlit as st
        def _find(node) -> str:
            if hasattr(node, "keys"):
                for k in node.keys():
                    if k == "session_json":
                        return str(node[k])
                    sub = _find(node[k])
                    if sub:
                        return sub
            return ""
        raw = _find(st.secrets)
        if raw:
            return json.loads(raw).get("cookies", [])
    except Exception:
        pass

    raw = os.getenv("session_json") or os.getenv("SESSION_JSON", "")
    if raw:
        return json.loads(raw).get("cookies", [])

    logger.error("Nenhuma sessão encontrada.")
    return []


def _sessao() -> requests.Session:
    s = requests.Session()
    s.headers.update({
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) Chrome/124.0",
    })
    for c in _carregar_cookies():
        domain = c.get("domain", "")
        if any(d in domain for d in ["comprasnet", "sso.acesso", "sistema.gov"]):
            s.cookies.set(c["name"], c["value"], domain=domain)
    return s


def _csrf_da_pagina(s: requests.Session, url: str) -> str:
    r = s.get(url, timeout=(5, 20), allow_redirects=False)
    if r.status_code in (301, 302, 303, 307):
        raise ValueError(
            "Sessão expirada — faça login em contratos.comprasnet.gov.br "
            "e atualize o session_json."
        )
    tag = BeautifulSoup(r.text, "lxml").find("meta", {"name": "csrf-token"})
    return tag["content"] if tag else ""


# ── Passo 1: encontrar compra_id ───────────────────────────────────────────────

def _buscar_uasg_id(s: requests.Session, uasg_code: str, csrf: str) -> str:
    """Busca o ID interno da UASG via endpoint Select2/Backpack."""
    endpoints = [
        f"{BASE}/empenho/buscacompra/fetch/unidade_origem_id",
        f"{BASE}/empenho/fetch/unidade_origem_id",
    ]
    hdrs = {
        "X-CSRF-TOKEN": csrf,
        "X-Requested-With": "XMLHttpRequest",
        "Accept": "application/json",
        "Content-Type": "application/json",
    }
    for url in endpoints:
        try:
            r = s.post(url, json={"q": uasg_code, "page": 1},
                       headers=hdrs, timeout=(5, 10))
            if r.status_code == 200:
                results = r.json().get("results", r.json().get("data", []))
                for item in results:
                    if uasg_code in str(item.get("text", "")):
                        logger.info("UASG %s → id=%s", uasg_code, item.get("id"))
                        return str(item.get("id", ""))
        except Exception as e:
            logger.debug("Endpoint %s falhou: %s", url, e)
    logger.warning("ID interno para UASG %s não encontrado — tentando sem ele.", uasg_code)
    return ""


def _buscar_compra_id(s: requests.Session, uasg: str, numero_ano: str, csrf: str) -> str:
    """POST a /empenho/buscacompra e extrai compra_id do redirect."""
    url = f"{BASE}/empenho/buscacompra"
    uasg_id = _buscar_uasg_id(s, uasg, csrf)

    payload = {
        "_token":            csrf,
        "tipoEmpenho":       "2",   # Compra
        "modalidade_id":     "76",  # Pregão
        "numero_ano":        numero_ano,
        "unidade_origem_id": uasg_id,
        "http_referrer":     url,
    }
    r = s.post(url, data=payload,
               headers={"Referer": url},
               allow_redirects=False,
               timeout=(5, 30))

    logger.info("POST buscacompra → %s | Location: %s",
                r.status_code, r.headers.get("location", ""))

    if r.status_code in (301, 302, 303):
        loc = r.headers.get("location", "")
        m = re.search(r"/empenho/fornecedor/(\d+)", loc)
        if m:
            return m.group(1)

    return ""


# ── Passo 2: listar fornecedores ───────────────────────────────────────────────

def _listar_fornecedores(s: requests.Session, compra_id: str, csrf: str) -> list[dict]:
    url  = f"{BASE}/empenho/fornecedor/{compra_id}"
    hdrs = {
        "Accept":           "application/json",
        "X-Requested-With": "XMLHttpRequest",
        "X-CSRF-TOKEN":     csrf,
        "Referer":          url,
    }
    total = 0
    forn  = []
    draw  = 1
    while True:
        r = s.get(url, params={"draw": draw, "start": len(forn), "length": 200},
                  headers=hdrs, timeout=30, allow_redirects=False)
        if r.status_code in (301, 302, 303):
            raise ValueError("Sessão expirada ao listar fornecedores.")
        r.raise_for_status()
        payload = r.json()
        rows = payload.get("data", [])
        try:
            total = int(payload.get("recordsFiltered") or payload.get("recordsTotal") or 0)
        except (ValueError, TypeError):
            total = 0
        for row in rows:
            forn.append({
                "fornecedor_id": str(row.get("id", "")),
                "nome":          row.get("nome", ""),
                "cnpj":          row.get("cpf_cnpj_idgener", ""),
                "itens":         [],
            })
        draw += 1
        if not rows or len(rows) < 200 or (total > 0 and len(forn) >= total):
            break
    logger.info("%d fornecedores | compra %s", len(forn), compra_id)
    return forn


# ── Passo 3: listar itens de um fornecedor ─────────────────────────────────────

def buscar_itens_fornecedor(compra_id: str, fornecedor_id: str) -> list[dict]:
    """Retorna itens de um fornecedor específico. Chamado sob demanda."""
    cookies = _carregar_cookies()
    if not cookies:
        return []
    s    = _sessao()
    url  = f"{BASE}/empenho/item/{compra_id}/{fornecedor_id}"
    csrf = s.cookies.get("XSRF-TOKEN", "")
    hdrs = {
        "Accept":           "application/json",
        "X-Requested-With": "XMLHttpRequest",
        "X-CSRF-TOKEN":     csrf,
        "Referer":          url,
    }
    r = s.get(url, params={"draw": 1, "start": 0, "length": 500},
              headers=hdrs, timeout=30)
    if not r.ok:
        logger.error("GET item %s/%s → %s", compra_id, fornecedor_id, r.status_code)
        return []
    itens = []
    for row in r.json().get("data", []):
        try:
            val_unit = float(row.get("valor_unitario") or 0)
        except (ValueError, TypeError):
            val_unit = 0.0
        itens.append({
            "item_id":      str(row.get("id", "")),
            "numero":       row.get("numero", ""),
            "descricao":    row.get("descricaodetalhada") or row.get("catmatser_desc", ""),
            "tipo":         row.get("descricao", "Material"),
            "qtd_saldo":    row.get("quantidade_saldo", ""),
            "valor_unit":   val_unit,
            "codigo_siasg": row.get("codigo_siasg", ""),
        })
    logger.info("%d itens | fornecedor %s", len(itens), fornecedor_id)
    return itens


# ── API pública ────────────────────────────────────────────────────────────────

class ResultadoBusca:
    def __init__(self):
        self.fornecedores: list[dict] = []
        self.compra_id:    str        = ""
        self.log:          list[str]  = []

    def add(self, msg: str):
        self.log.append(msg)
        logger.info(msg)


def buscar_pregao(uasg: str, numero_ano: str,
                  compra_id_direto: str = "") -> ResultadoBusca:
    """
    Busca fornecedores de um pregão.
    - compra_id_direto: se informado, pula a etapa de busca.
    """
    r = ResultadoBusca()
    r.add(f"📥 UASG={uasg} | Pregão={numero_ano}")

    cookies = _carregar_cookies()
    if not cookies:
        raise FileNotFoundError(
            "Sessão não encontrada. Configure session_json com cookies de "
            "contratos.comprasnet.gov.br e atualize no Railway."
        )

    s    = _sessao()
    csrf = _csrf_da_pagina(s, f"{BASE}/empenho/buscacompra")

    if compra_id_direto:
        compra_id = compra_id_direto.strip()
        r.add(f"✅ Compra ID informado: {compra_id}")
    else:
        r.add("🔍 Buscando compra no portal...")
        compra_id = _buscar_compra_id(s, uasg, numero_ano, csrf)
        if not compra_id:
            raise ValueError(
                f"Compra não encontrada para UASG {uasg} / Pregão {numero_ano}.\n"
                "Dica: informe o ID diretamente (número na URL do portal, ex: 7194897)."
            )
        r.add(f"✅ Compra encontrada: {compra_id}")

    r.compra_id = compra_id

    r.add("🏢 Carregando fornecedores...")
    r.fornecedores = _listar_fornecedores(s, compra_id, csrf)
    r.add(f"✅ {len(r.fornecedores)} fornecedor(es). Expanda para ver os itens.")
    return r
