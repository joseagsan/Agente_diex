"""
Pesquisa de itens de pregão no comprasnet.gov.br via Playwright.
Preenche o formulário real da página, igual ao uso manual.
"""
import logging
import os
import re

import urllib3
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

logger = logging.getLogger(__name__)

_BASE = "https://www.comprasnet.gov.br"
_ATA_URL = f"{_BASE}/livre/pregao/ata0.asp"


# ── Normalização ───────────────────────────────────────────────────────────────

def _normalizar_pregao(num: str) -> tuple[str, str, str]:
    """
    Extrai seq e ano do formato "90008/2025".
    Returns (num_norm_AAAANNNNN, seq, ano)
    """
    import datetime
    ano_atual = str(datetime.date.today().year)
    num = re.sub(r"^[A-Za-z\s]+", "", num).strip()

    # "90008/2025"
    m = re.match(r"^(\d+)[/\-](20\d{2})$", num)
    if m:
        seq, ano = m.group(1), m.group(2)
        return ano + seq.zfill(5), seq, ano

    # "2025/90008"
    m = re.match(r"^(20\d{2})[/\-](\d+)$", num)
    if m:
        ano, seq = m.group(1), m.group(2)
        return ano + seq.zfill(5), seq, ano

    seq = re.sub(r"\D", "", num)
    return ano_atual + seq.zfill(5), seq, ano_atual


# ── Helpers ────────────────────────────────────────────────────────────────────

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


def _extrair_tabela_html(html_bytes: bytes, vig_inicio: str = "", vig_fim: str = "") -> list[dict]:
    """Extrai itens de uma tabela HTML do comprasnet."""
    from bs4 import BeautifulSoup
    soup = BeautifulSoup(html_bytes, "lxml")

    texto = soup.get_text()
    if not vig_inicio:
        m = re.search(r"vig[eê]ncia[:\s]+(\d{2}/\d{2}/\d{4})[^0-9]*(\d{2}/\d{2}/\d{4})",
                      texto, re.IGNORECASE)
        if m:
            vig_inicio, vig_fim = m.group(1), m.group(2)

    cnpj_global = ""
    m = re.search(r"(\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2})", texto)
    if m:
        cnpj_global = m.group(1)

    itens: list[dict] = []
    for tabela in soup.find_all("table"):
        headers_els = tabela.find_all("th") or (tabela.find("tr").find_all("td") if tabela.find("tr") else [])
        headers = [h.get_text(strip=True).lower() for h in headers_els]

        col = {}
        for i, h in enumerate(headers):
            if re.search(r"\bitem\b|\bnr\.?\b|\bn[oº]\.?\b", h): col.setdefault("numero_item", i)
            elif re.search(r"descri[cç]", h):                     col.setdefault("descricao", i)
            elif re.search(r"unid|und\b",  h):                    col.setdefault("und", i)
            elif re.search(r"unit[aá]rio|unit\b|vl.*unit", h):    col.setdefault("valor_unitario", i)
            elif re.search(r"fornecedor|empresa|raz[aã]o",  h):   col.setdefault("fornecedor", i)
            elif re.search(r"cnpj", h):                            col.setdefault("cnpj", i)

        if "descricao" not in col:
            continue

        for linha in tabela.find_all("tr")[1:]:
            cels = linha.find_all("td")
            if len(cels) < 2:
                continue

            def _c(key):
                idx = col.get(key)
                return cels[idx].get_text(strip=True) if idx is not None and idx < len(cels) else ""

            desc = _c("descricao")
            if not desc:
                continue

            v = _parse_br(_c("valor_unitario"))
            itens.append({
                "numero_item":     _c("numero_item"),
                "descricao":       desc,
                "und":             _c("und") or "UN",
                "valor_unit_num":  v,
                "valor_unit":      _fmt(v),
                "fornecedor":      _c("fornecedor"),
                "cnpj":            _c("cnpj") or cnpj_global,
                "vigencia_inicio": vig_inicio,
                "vigencia_fim":    vig_fim,
            })
    return itens


# ── Playwright ─────────────────────────────────────────────────────────────────

def _caminho_chromium() -> str | None:
    import shutil
    for nome in ("chromium", "chromium-browser"):
        p = shutil.which(nome)
        if p:
            return p
    for p in ("/usr/bin/chromium", "/usr/bin/chromium-browser"):
        if os.path.exists(p):
            return p
    return None


def garantir_navegador() -> None:
    import subprocess, sys
    try:
        from playwright.sync_api import sync_playwright
        exe = _caminho_chromium()
        with sync_playwright() as p:
            kw: dict = {"headless": True}
            if exe:
                kw["executable_path"] = exe
            b = p.chromium.launch(**kw)
            b.close()
            return
    except Exception:
        pass
    for cmd in [
        [sys.executable, "-m", "playwright", "install", "chromium", "--with-deps"],
        [sys.executable, "-m", "playwright", "install", "chromium"],
    ]:
        try:
            r = subprocess.run(cmd, capture_output=True, timeout=300)
            if r.returncode == 0:
                return
        except Exception:
            continue
    raise RuntimeError("Chromium não encontrado. Execute: playwright install chromium")


# ── Resultado de busca ─────────────────────────────────────────────────────────

class ResultadoBusca:
    def __init__(self):
        self.itens:     list[dict] = []
        self.url_usada: str = ""
        self.log:       list[str] = []

    def add(self, msg: str):
        self.log.append(msg)
        logger.info(msg)


# ── Busca principal ────────────────────────────────────────────────────────────

def buscar_itens_pregao(uasg: str, num_pregao: str) -> ResultadoBusca:
    """
    Abre comprasnet.gov.br/livre/pregao/ata0.asp com Playwright,
    preenche o formulário com UASG e número do pregão e extrai os itens.
    """
    from playwright.sync_api import sync_playwright, TimeoutError as PTimeout

    r = ResultadoBusca()
    num_norm, seq, ano = _normalizar_pregao(num_pregao)
    r.add(f"📥 Entrada: UASG={uasg} | Pregão='{num_pregao}' → seq={seq} ano={ano} norm={num_norm}")
    r.add(f"🌐 Abrindo: {_ATA_URL}")

    exe = _caminho_chromium()
    launch_kw: dict = {"headless": True}
    if exe:
        launch_kw["executable_path"] = exe
        r.add(f"🖥️ Usando Chromium do sistema: {exe}")

    with sync_playwright() as p:
        try:
            browser = p.chromium.launch(**launch_kw)
        except Exception:
            browser = p.chromium.launch(headless=True)

        page = browser.new_context().new_page()
        try:
            page.goto(_ATA_URL, wait_until="networkidle", timeout=30000)
            r.add(f"✅ Página carregada. Título: {page.title()!r}")

            # ── Inspeciona campos disponíveis ──────────────────────────────
            campos = page.evaluate("""
                () => Array.from(document.querySelectorAll('input, select, textarea'))
                     .map(e => ({tag: e.tagName, type: e.type||'', name: e.name||'', id: e.id||''}))
            """)
            r.add(f"📋 Campos do formulário: {campos}")

            # ── Seleciona "Pregão Eletrônico" ──────────────────────────────
            for sel in ["input[value='0'][type='radio']",
                        "input[name='radio_tipo_pregao'][value='0']",
                        "input[type='radio']"]:
                try:
                    if page.locator(sel).count() > 0:
                        page.locator(sel).first.check()
                        r.add(f"☑️ Selecionado radio: {sel}")
                        break
                except Exception:
                    pass

            # ── Preenche UASG ──────────────────────────────────────────────
            uasg_preenchido = False
            for sel in ["input[name='co_no_uasg']", "input[name='uasg']",
                        "#co_no_uasg", "#uasg", "input[name*='uasg' i]"]:
                try:
                    loc = page.locator(sel)
                    if loc.count() > 0:
                        loc.first.fill(uasg)
                        r.add(f"✏️ UASG preenchido em: {sel}")
                        uasg_preenchido = True
                        break
                except Exception:
                    pass
            if not uasg_preenchido:
                r.add("⚠️ Campo UASG não encontrado pelos seletores conhecidos.")

            # ── Preenche nº do pregão ──────────────────────────────────────
            pregao_preenchido = False
            for valor in [num_norm, seq, num_pregao]:
                for sel in ["input[name='numprp']", "input[name='num_pregao']",
                            "input[name='nu_licitacao']", "#numprp", "#num_pregao",
                            "input[name*='preg' i]", "input[name*='num' i]"]:
                    try:
                        loc = page.locator(sel)
                        if loc.count() > 0 and loc.first.is_visible():
                            loc.first.fill(valor)
                            r.add(f"✏️ Pregão preenchido '{valor}' em: {sel}")
                            pregao_preenchido = True
                            break
                    except Exception:
                        pass
                if pregao_preenchido:
                    break

            if not pregao_preenchido:
                r.add("⚠️ Campo do pregão não encontrado. Campos disponíveis listados acima.")

            # ── Submete o formulário ───────────────────────────────────────
            submetido = False
            for sel in ["input[name='b'][value='OK']", "input[type='submit']",
                        "button[type='submit']", "input[value*='Pesquis' i]",
                        "input[value='OK']"]:
                try:
                    loc = page.locator(sel)
                    if loc.count() > 0:
                        loc.first.click()
                        r.add(f"🖱️ Clicado: {sel}")
                        submetido = True
                        break
                except Exception:
                    pass

            if not submetido:
                r.add("⚠️ Botão de submit não encontrado — tentando Enter.")
                try:
                    page.keyboard.press("Enter")
                except Exception:
                    pass

            page.wait_for_load_state("networkidle", timeout=15000)
            r.add(f"🔄 Após submit — URL: {page.url} | Título: {page.title()!r}")

            # ── Extrai HTML e parse ────────────────────────────────────────
            html = page.content().encode("utf-8")
            preview = page.locator("body").inner_text()[:400].replace("\n", " ")
            r.add(f"📄 Prévia: {preview!r}")

            itens = _extrair_tabela_html(html)
            r.add(f"📦 {len(itens)} itens extraídos.")

            if itens:
                r.itens = itens
                r.url_usada = page.url

        except Exception as e:
            r.add(f"❌ Erro: {e}")
        finally:
            browser.close()

    return r
