"""
Pesquisa de itens de pregão no portal comprasnet.gov.br.
Usa requests + BeautifulSoup para extrair os dados sem depender de Claude API.
"""
import logging
import re

import requests
import urllib3

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

logger = logging.getLogger(__name__)

_HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                  "(KHTML, like Gecko) Chrome/124.0 Safari/537.36",
    "Accept-Language": "pt-BR,pt;q=0.9",
}


def _normalizar_pregao(num: str) -> tuple[str, str, str]:
    """
    Normaliza o número do pregão para o formato comprasnet: AAAANNNNN.
    Padrão esperado: "SEQUENCIAL/ANO" (ex: "90008/2025").

    Returns:
        (num_norm, seq, ano) — ex: ("202590008", "90008", "2025")
    """
    import datetime
    ano_atual = str(datetime.date.today().year)

    num = re.sub(r"^[A-Za-z\s]+", "", num).strip()

    # Formato principal: "90008/2025"
    m = re.match(r"^(\d+)[/\-](20\d{2})$", num)
    if m:
        seq, ano = m.group(1), m.group(2)
        return ano + seq.zfill(5), seq, ano

    # Formato invertido: "2025/90008"
    m = re.match(r"^(20\d{2})[/\-](\d+)$", num)
    if m:
        ano, seq = m.group(1), m.group(2)
        return ano + seq.zfill(5), seq, ano

    # Só número: assume ano atual
    seq = re.sub(r"\D", "", num)
    return ano_atual + seq.zfill(5), seq, ano_atual


def _fmt(v: float) -> str:
    return "R$ " + f"{v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


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


_BASE = "https://www.comprasnet.gov.br"


def _tentativas_pregao(uasg: str, num_norm: str, num_original: str) -> list[dict]:
    """
    Retorna lista de tentativas com POST e GET, testando o número
    normalizado (AAAANNNNN) e o original (ex: 90008/2025).
    """
    ata_url   = f"{_BASE}/livre/pregao/ata0.asp"
    multi_url = f"{_BASE}/livre/pregao/result_multif0.asp"

    def _post(numprp):
        return {"method": "POST", "url": ata_url, "data": {
            "co_no_uasg": uasg, "numprp": numprp,
            "radio_tipo_pregao": "1", "b": "OK",
        }}

    def _get(url, numprp):
        return {"method": "GET", "url": url,
                "params": {"co_no_uasg": uasg, "numprp": numprp}}

    tentativas = [
        {"label": f"ATA POST  num={num_norm}",    **_post(num_norm)},
        {"label": f"ATA POST  num={num_original}", **_post(num_original)},
        {"label": f"ATA GET   num={num_norm}",    **_get(ata_url, num_norm)},
        {"label": f"ATA GET   num={num_original}", **_get(ata_url, num_original)},
        {"label": f"Multi GET num={num_norm}",    **_get(multi_url, num_norm)},
    ]
    return tentativas


def _extrair_tabela(html_bytes: bytes) -> list[dict]:
    """
    Tenta extrair itens do HTML do comprasnet usando BeautifulSoup.
    Suporta a estrutura de tabela da página ata0.asp.
    """
    from bs4 import BeautifulSoup

    soup = BeautifulSoup(html_bytes, "lxml")

    # Vigência — aparece em texto como "Vigência: DD/MM/YYYY a DD/MM/YYYY"
    vig_inicio = vig_fim = ""
    texto_pagina = soup.get_text()
    m = re.search(r"vig[eê]ncia[:\s]+(\d{2}/\d{2}/\d{4})[^0-9]*(\d{2}/\d{2}/\d{4})",
                  texto_pagina, re.IGNORECASE)
    if m:
        vig_inicio, vig_fim = m.group(1), m.group(2)

    # CNPJ do fornecedor — primeiro que aparecer no texto
    cnpj_global = ""
    m = re.search(r"(\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2})", texto_pagina)
    if m:
        cnpj_global = m.group(1)

    # Procura tabelas com colunas de item
    itens: list[dict] = []
    for tabela in soup.find_all("table"):
        headers = [th.get_text(strip=True).lower() for th in tabela.find_all("th")]
        if not headers:
            # Tenta primeira linha como cabeçalho
            primeira = tabela.find("tr")
            if primeira:
                headers = [td.get_text(strip=True).lower()
                           for td in primeira.find_all(["td", "th"])]

        # Identifica colunas relevantes
        col_map = {}
        for i, h in enumerate(headers):
            if re.search(r"\bitem\b|\bnr\b|\bn[oº]\b", h):
                col_map.setdefault("numero_item", i)
            elif re.search(r"descri[cç]", h):
                col_map.setdefault("descricao", i)
            elif re.search(r"unid|und\b", h):
                col_map.setdefault("und", i)
            elif re.search(r"unit[aá]rio|unit\b|valor\s*unit", h):
                col_map.setdefault("valor_unitario", i)
            elif re.search(r"fornecedor|empresa|razão social", h):
                col_map.setdefault("fornecedor", i)
            elif re.search(r"cnpj", h):
                col_map.setdefault("cnpj", i)

        if "descricao" not in col_map:
            continue

        linhas = tabela.find_all("tr")[1:]
        for linha in linhas:
            cels = linha.find_all("td")
            if len(cels) < 2:
                continue

            def _cel(key):
                idx = col_map.get(key)
                if idx is None or idx >= len(cels):
                    return ""
                return cels[idx].get_text(strip=True)

            desc = _cel("descricao")
            if not desc:
                continue

            v = _parse_br(_cel("valor_unitario"))
            cnpj = _cel("cnpj") or cnpj_global

            itens.append({
                "numero_item":     _cel("numero_item"),
                "descricao":       desc,
                "und":             _cel("und") or "UN",
                "valor_unit_num":  v,
                "valor_unit":      _fmt(v),
                "fornecedor":      _cel("fornecedor"),
                "cnpj":            cnpj,
                "vigencia_inicio": vig_inicio,
                "vigencia_fim":    vig_fim,
            })

    return itens


class ResultadoBusca:
    def __init__(self):
        self.itens: list[dict] = []
        self.url_usada: str = ""
        self.log: list[str] = []   # diagnóstico passo a passo

    def add_log(self, msg: str):
        self.log.append(msg)
        logger.info(msg)


def buscar_itens_pregao(uasg: str, num_pregao: str) -> ResultadoBusca:
    """
    Busca itens de um pregão no comprasnet.gov.br.
    Retorna ResultadoBusca com itens, url_usada e log de diagnóstico.
    """
    r = ResultadoBusca()
    num_norm, seq, ano = _normalizar_pregao(num_pregao)

    r.add_log(f"📥 Entrada: UASG={uasg} | Pregão='{num_pregao}' → seq={seq} ano={ano}")
    r.add_log(f"🔢 Normalizado para comprasnet: '{num_norm}' (formato AAAANNNNN)")

    num_original = re.sub(r"^[A-Za-z\s]+", "", num_pregao).strip()
    for t in _tentativas_pregao(uasg, num_norm, num_original):
        label = t["label"]
        url   = t["url"]
        r.add_log(f"🌐 [{label}]: {url}")
        try:
            if t["method"] == "POST":
                resp = requests.post(url, data=t["data"], headers=_HEADERS, timeout=30, verify=False)
            else:
                resp = requests.get(url, params=t.get("params"), headers=_HEADERS, timeout=30, verify=False)
            r.add_log(f"   → HTTP {resp.status_code} | {len(resp.content)} bytes | "
                      f"encoding: {resp.encoding}")

            # Prévia do conteúdo (para diagnóstico)
            try:
                from bs4 import BeautifulSoup
                soup = BeautifulSoup(resp.content, "lxml")
                texto_preview = soup.get_text(separator=" ", strip=True)[:300]
            except Exception:
                texto_preview = resp.content[:300].decode("latin-1", errors="replace")
            r.add_log(f"   → Prévia: {texto_preview!r}")

            if resp.status_code != 200:
                r.add_log(f"   ❌ Status {resp.status_code} — pulando.")
                continue
            if len(resp.content) < 2000:
                r.add_log(f"   ❌ Conteúdo muito curto ({len(resp.content)}B) — pregão provavelmente não encontrado.")
                continue

            itens = _extrair_tabela(resp.content)
            r.add_log(f"   → {len(itens)} itens extraídos pela tabela.")

            if itens:
                r.itens = itens
                r.url_usada = url
                r.add_log(f"   ✅ Sucesso!")
                return r

            r.add_log(f"   ⚠️ HTML obtido mas nenhuma tabela de itens reconhecida.")

        except Exception as e:
            r.add_log(f"   ❌ Exceção: {e}")

    r.add_log("❌ Nenhuma URL retornou itens.")
    return r
