"""
Pesquisa de itens de pregão em contratos.sistema.gov.br.
Fluxo:
  1. POST /transparencia/compras/search  →  acha ID da compra pelo UASG + número pregão
  2. POST /transparencia/compras/{ID}/itens/search  →  lista itens (número, descrição)
  3. GET  /transparencia/compras/{ID}/itens/{ITEM_ID}/show  →  detalhe (preço, fornecedor, CNPJ, vigência)
Requer session.json válido (login via agente-licitacoes).
"""
import logging
import os
import re

import requests
from bs4 import BeautifulSoup

logger = logging.getLogger(__name__)

BASE = "https://contratos.sistema.gov.br"

# Procura session.json em locais possíveis (local e Streamlit Cloud)
def _session_file() -> str:
    candidatos = [
        os.path.join(os.path.dirname(__file__), "session.json"),
        os.path.join(os.path.dirname(__file__), "..", "agente-licitacoes", "session.json"),
        "session.json",
    ]
    for c in candidatos:
        if os.path.exists(c):
            return c
    return candidatos[0]  # retorna o primeiro mesmo que não exista (erro mais claro)

SESSION_FILE = _session_file()


# ── Sessão autenticada ─────────────────────────────────────────────────────────

def _carregar_cookies() -> list:
    """Carrega cookies do session.json (arquivo local ou Streamlit secrets)."""
    import json

    # 1. Arquivo local
    if os.path.exists(SESSION_FILE):
        logger.info("Sessão carregada do arquivo: %s", SESSION_FILE)
        with open(SESSION_FILE) as f:
            return json.load(f).get("cookies", [])

    # 2. Streamlit secrets — tenta raiz e todas as seções
    try:
        import streamlit as st

        def _achar_raw(node) -> str:
            """Busca 'session_json' recursivamente em qualquer nível dos secrets."""
            if hasattr(node, "keys"):
                for k in node.keys():
                    if k == "session_json":
                        return str(node[k])
                    sub = _achar_raw(node[k])
                    if sub:
                        return sub
            return ""

        raw = _achar_raw(st.secrets)
        if raw:
            logger.info("Sessão carregada dos Streamlit secrets.")
            return json.loads(raw).get("cookies", [])
    except Exception as e:
        logger.warning("Erro ao ler Streamlit secrets: %s", e)

    # 3. Variável de ambiente (Railway)
    raw = os.getenv("session_json") or os.getenv("SESSION_JSON", "")
    if raw:
        logger.info("Sessão carregada da variável de ambiente.")
        return json.loads(raw).get("cookies", [])

    logger.error("Nenhuma sessão encontrada (arquivo nem secrets nem env).")
    return []


def _sessao() -> requests.Session:
    s = requests.Session()
    s.headers.update({"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) Chrome/124.0"})
    for c in _carregar_cookies():
        s.cookies.set(c["name"], c["value"], domain=c.get("domain", ""))
    return s


def _csrf(s: requests.Session, url: str) -> str:
    r = s.get(url, timeout=(5, 15), allow_redirects=False)
    if r.status_code in (301, 302, 303, 307, 308):
        raise ValueError("Sessão expirada (redirecionado ao buscar CSRF) — atualize session_json.")
    tag = BeautifulSoup(r.text, "lxml").find("meta", {"name": "csrf-token"})
    return tag["content"] if tag else ""


def _hdrs(csrf_token: str, referer: str) -> dict:
    return {
        "Accept": "application/json",
        "X-Requested-With": "XMLHttpRequest",
        "X-CSRF-TOKEN": csrf_token,
        "Referer": referer,
    }


# ── Helpers ────────────────────────────────────────────────────────────────────

def _txt(html) -> str:
    return BeautifulSoup(str(html), "lxml").get_text(strip=True)


def _parse_br(texto: str) -> float:
    t = re.sub(r"[R$\s\xa0]", "", str(texto))
    if not t:
        return 0.0
    if "," in t:
        t = t.replace(".", "").replace(",", ".")
    try:
        return float(t)
    except ValueError:
        return 0.0


def _fmt(v: float) -> str:
    return "R$ " + f"{v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


# ── Passo 1: achar ID da compra ────────────────────────────────────────────────

def _buscar_id_compra(s: requests.Session, uasg: str, num_pregao: str) -> str:
    """
    Retorna o ID interno da compra (ex: '795827') para a combinação UASG + número pregão.
    """
    url_list = BASE + "/transparencia/compras"
    csrf = _csrf(s, url_list)
    logger.info("CSRF obtido: %s...", csrf[:20] if csrf else "VAZIO")

    hdrs = _hdrs(csrf, url_list)

    r = s.post(url_list + "/search",
               data={"start": 0, "length": 200, "draw": 1,
                     "search[value]": num_pregao},
               headers=hdrs, timeout=30)

    logger.info("POST /compras/search → HTTP %s | %d bytes", r.status_code, len(r.content))

    # Verifica se é JSON ou HTML (redirecionamento para login)
    ct = r.headers.get("Content-Type", "")
    if "text/html" in ct:
        logger.error("Resposta é HTML — sessão provavelmente expirou. CT: %s", ct)
        raise ValueError(
            "Sessão expirada. Faça login novamente no agente-licitacoes e atualize o "
            "session_json nos Secrets do Streamlit Cloud."
        )

    r.raise_for_status()
    payload = r.json()
    rows = payload.get("data", [])
    logger.info("Linhas retornadas: %d | recordsFiltered: %s",
                len(rows), payload.get("recordsFiltered", "?"))

    # Log das primeiras 3 linhas para diagnóstico
    for i, row in enumerate(rows[:3]):
        cols = [_txt(c) for c in row]
        logger.info("  Linha %d: %s", i, " | ".join(cols[:7]))

    for row in rows:
        cols = [_txt(c) for c in row]
        uasg_col   = cols[0] if cols else ""
        numero_col = cols[5] if len(cols) > 5 else ""

        if uasg in uasg_col and num_pregao in numero_col:
            for c in row:
                ids = re.findall(r"/transparencia/compras/(\d+)/", str(c))
                if ids:
                    logger.info("ID encontrado: %s", ids[0])
                    return ids[0]

    raise ValueError(
        f"Pregão '{num_pregao}' não encontrado para UASG {uasg} "
        f"({len(rows)} resultado(s) retornado(s)). "
        f"Verifique o número do pregão — formato esperado: 90009/2025"
    )


# ── Passo 2: listar itens ──────────────────────────────────────────────────────

def _listar_itens(s: requests.Session, compra_id: str) -> list[dict]:
    """
    Retorna lista básica de itens: {numero, descricao, item_id}.
    Pagina automaticamente enquanto o servidor retornar linhas.
    """
    url_itens = f"{BASE}/transparencia/compras/{compra_id}/itens"

    # CSRF buscado UMA vez — válido por toda a sessão
    csrf = _csrf(s, url_itens)
    hdrs = _hdrs(csrf, url_itens)

    PAGE  = 200
    itens = []
    draw  = 1
    total = 0

    while True:
        r = s.post(
            url_itens + "/search",
            data={"start": len(itens), "length": PAGE, "draw": draw},
            headers=hdrs,
            timeout=30,
            allow_redirects=False,
        )

        if r.status_code in (301, 302, 303, 307, 308):
            raise ValueError("Sessão expirada ao paginar itens — atualize session_json.")

        r.raise_for_status()
        payload = r.json()
        rows = payload.get("data", [])

        try:
            total = int(payload.get("recordsFiltered") or payload.get("recordsTotal") or 0)
        except (ValueError, TypeError):
            total = 0

        before = len(itens)
        for row in rows:
            numero    = _txt(row[0]) if row else ""
            descricao = _txt(row[2]) if len(row) > 2 else ""
            item_id   = ""
            for c in row:
                ids = re.findall(r"/itens/(\d+)/show", str(c))
                if ids:
                    item_id = ids[0]
                    break
            if descricao and item_id:
                itens.append({"numero": numero, "descricao": descricao, "item_id": item_id})

        logger.info("Página %d: %d linhas, %d novos, total %d/%d",
                    draw, len(rows), len(itens) - before, len(itens), total)

        draw += 1
        if not rows or len(rows) < PAGE or (total > 0 and len(itens) >= total):
            break

    logger.info("Itens coletados: %d / %d", len(itens), total)
    return itens


# ── Passo 3: detalhe de um item ────────────────────────────────────────────────

def _detalhe_item(s: requests.Session, compra_id: str, item_id: str) -> dict:
    """
    Busca fornecedor, CNPJ, valor unitário e vigência ARP na página de detalhe do item.
    """
    url = f"{BASE}/transparencia/compras/{compra_id}/itens/{item_id}/show"
    r = s.get(url, headers={"User-Agent": "Mozilla/5.0"}, timeout=(5, 15))
    logger.info("DETALHE %s → HTTP %s | URL final: %s", item_id, r.status_code, r.url)
    ct = r.headers.get("Content-Type", "")
    if "login" in r.url or "sso" in r.url:
        raise ValueError("Sessão expirada — faça login novamente e atualize session_json.")
    soup = BeautifulSoup(r.text, "lxml")
    texto = soup.get_text(separator="\n", strip=True)
    logger.info("DETALHE texto (300 chars): %s", texto[:300].replace("\n", " | "))

    result = {
        "und":             "UN",
        "valor_unit_num":  0.0,
        "valor_unit":      "R$ 0,00",
        "fornecedor":      "",
        "cnpj":            "",
        "vigencia_inicio": "",
        "vigencia_fim":    "",
        "descricao_det":   "",
    }

    # Descrição detalhada
    m = re.search(r"Descrição detalhada[:\s\n]+([^\n]{5,300})", texto)
    if m:
        result["descricao_det"] = m.group(1).strip()

    # Vigência ARP
    m = re.search(r"Vig\. Início ARP[:\s\n]+(\d{2}/\d{2}/\d{4})", texto)
    if m:
        result["vigencia_inicio"] = m.group(1)
    m = re.search(r"Vig\. Fim ARP[:\s\n]+(\d{2}/\d{2}/\d{4})", texto)
    if m:
        result["vigencia_fim"] = m.group(1)

    # Fornecedor e CNPJ (seção "Fornecedores Homologados")
    cnpjs = re.findall(r"(\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2})\s*[-–]\s*([^\n]{5,80})", texto)
    if cnpjs:
        result["cnpj"]       = cnpjs[0][0]
        result["fornecedor"] = cnpjs[0][1].strip()

    # Valor unitário — pega o "Val. Unitário" da tabela de fornecedores
    m = re.search(r"Val\. Unit[aá]rio\s*\n([\d.,]+)", texto)
    if m:
        v = _parse_br(m.group(1))
        result["valor_unit_num"] = v
        result["valor_unit"]     = _fmt(v)
    else:
        # Tenta extrair da tabela de ATAs: último valor antes de "Valor Total"
        vals = re.findall(r"\n([\d]+[.,][\d]{2})\n[\d]+[.,]", texto)
        if vals:
            v = _parse_br(vals[-1])
            result["valor_unit_num"] = v
            result["valor_unit"]     = _fmt(v)

    return result


# ── Resultado de busca ─────────────────────────────────────────────────────────

class ResultadoBusca:
    def __init__(self):
        self.itens:     list[dict] = []
        self.url_usada: str = ""
        self.log:       list[str] = []

    def add(self, msg: str):
        self.log.append(msg)
        logger.info(msg)


# ── API pública ────────────────────────────────────────────────────────────────

def buscar_itens_pregao(uasg: str, num_pregao: str) -> ResultadoBusca:
    """
    Busca RÁPIDA — só lista itens (número e descrição).
    Detalhes (preço, fornecedor, CNPJ) são buscados depois com buscar_detalhe_item().
    """
    r = ResultadoBusca()
    r.add(f"📥 UASG={uasg} | Pregão={num_pregao}")

    cookies = _carregar_cookies()
    if not cookies:
        raise FileNotFoundError(
            "Sessão não encontrada. Adicione 'session_json' nos Secrets do Streamlit Cloud "
            "ou faça login via agente-licitacoes (session.json)."
        )

    s = _sessao()

    r.add("🔍 Buscando ID da compra...")
    compra_id = _buscar_id_compra(s, uasg, num_pregao)
    r.add(f"✅ ID encontrado: {compra_id}")
    r.url_usada = f"{BASE}/transparencia/compras/{compra_id}/itens"

    r.add("📋 Listando itens...")
    itens_base = _listar_itens(s, compra_id)
    r.add(f"✅ {len(itens_base)} itens encontrados. Selecione um para ver detalhes.")

    resultados = []
    for it in itens_base:
        resultados.append({
            "numero_item":     it["numero"],
            "item_id":         it["item_id"],
            "compra_id":       compra_id,
            "descricao":       it["descricao"],
            "und":             "",
            "valor_unit_num":  0.0,
            "valor_unit":      "—",
            "fornecedor":      "",
            "cnpj":            "",
            "vigencia_inicio": "",
            "vigencia_fim":    "",
        })

    r.itens = resultados
    r.add(f"🎯 Pronto! Selecione um item e clique 'Usar' para carregar os detalhes.")
    return r


def buscar_detalhe_item(compra_id: str, item_id: str) -> dict:
    """
    Busca detalhe de UM item específico (preço, fornecedor, CNPJ, vigência).
    Chamado apenas quando o usuário clica 'Usar'. v2
    """
    try:
        cookies = _carregar_cookies()
        if not cookies:
            logger.error("Sem cookies — session_json não configurado.")
            return {"_erro": "Sessão não encontrada. Configure session_json."}
        s = _sessao()
        return _detalhe_item(s, compra_id, item_id)
    except Exception as e:
        logger.error("Erro ao buscar detalhe %s/%s: %s", compra_id, item_id, e)
        return {"_erro": str(e)}

