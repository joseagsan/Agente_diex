"""
Pesquisa de itens de ARP no portal contratos.sistema.gov.br para preenchimento de REQ.
Seletores validados pelo agente-licitacoes/crawler.py.
"""
import logging
import os
import re

logger = logging.getLogger(__name__)

SESSION_FILE = "session.json"
BASE_URL = "https://contratos.sistema.gov.br/transparencia/arp-item"

# Índices das colunas na tabela #itens
COL_NUMERO_ATA  = 0
COL_UNIDADE     = 1
COL_NUMERO_ITEM = 2
COL_PDM         = 3
COL_DESCRICAO   = 4
COL_UF          = 5
COL_FORNECEDOR  = 6
COL_QTD         = 7
COL_SALDO       = 8
COL_INICIO_VIG  = 9
COL_FIM_VIG     = 10
COL_ACAO        = 11


def playwright_disponivel() -> bool:
    try:
        import playwright  # noqa
        return True
    except ImportError:
        return False


def _caminho_chromium_sistema() -> str | None:
    import shutil
    for nome in ("chromium", "chromium-browser"):
        path = shutil.which(nome)
        if path:
            return path
    for path in ("/usr/bin/chromium", "/usr/bin/chromium-browser"):
        if os.path.exists(path):
            return path
    return None


def garantir_navegador() -> None:
    """Verifica se o Playwright consegue abrir um navegador; instala se necessário."""
    import subprocess
    import sys

    try:
        from playwright.sync_api import sync_playwright
        exe = _caminho_chromium_sistema()
        with sync_playwright() as p:
            kwargs: dict = {"headless": True}
            if exe:
                kwargs["executable_path"] = exe
            b = p.chromium.launch(**kwargs)
            b.close()
            return
    except Exception:
        pass

    logger.info("Instalando Chromium via playwright install...")
    for cmd in [
        [sys.executable, "-m", "playwright", "install", "chromium", "--with-deps"],
        [sys.executable, "-m", "playwright", "install", "chromium"],
    ]:
        try:
            r = subprocess.run(cmd, capture_output=True, timeout=300)
            if r.returncode == 0:
                logger.info("Chromium instalado.")
                return
        except Exception:
            continue

    raise RuntimeError("Chromium não encontrado. Execute: playwright install chromium")


def _parse_br(texto: str) -> float:
    t = str(texto).strip().replace("R$", "").replace("\xa0", "").replace(" ", "")
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


def _extrair_detalhe(context, link: str) -> dict:
    """Abre a página de detalhe e extrai valor unitário, UND, fornecedor e CNPJ."""
    page = context.new_page()
    dados = {"valor_unitario": 0.0, "und": "", "fornecedor": "", "cnpj": ""}
    try:
        page.goto(link, wait_until="networkidle", timeout=20000)
        content = page.content()

        # Valor unitário
        try:
            tabela = page.locator("table:has(th:has-text('Valor unitário'))")
            if tabela.count() > 0:
                td = tabela.first.locator("tbody tr:first-child td").last
                if td.count() > 0:
                    dados["valor_unitario"] = _parse_br(td.inner_text())
        except Exception:
            pass

        # Unidade de medida
        try:
            tabela = page.locator("table:has(th:has-text('Unidade de fornecimento'))")
            if tabela.count() == 0:
                tabela = page.locator("table:has(th:has-text('Unidade'))")
            if tabela.count() > 0:
                td = tabela.first.locator("tbody tr:first-child td").nth(1)
                if td.count() > 0:
                    dados["und"] = td.inner_text().strip()
        except Exception:
            pass

        # CNPJ
        m = re.search(r"(\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2})", content)
        if m:
            dados["cnpj"] = m.group(1)

        # Fornecedor — tenta tabela primeiro
        try:
            tabela = page.locator("table:has(th:has-text('Fornecedor'))")
            if tabela.count() > 0:
                td = tabela.first.locator("tbody tr:first-child td").nth(1)
                if td.count() > 0:
                    dados["fornecedor"] = td.inner_text().strip()
        except Exception:
            pass

        if not dados["fornecedor"]:
            for pattern in [
                r"Fornecedor[^:]*?:\s*([^\n<]{5,100})",
                r"Razão Social[^:]*?:\s*([^\n<]{5,100})",
            ]:
                m = re.search(pattern, content, re.IGNORECASE)
                if m:
                    dados["fornecedor"] = m.group(1).strip()
                    break

        return dados
    except Exception as e:
        logger.warning("Erro ao extrair detalhe de %s: %s", link, e)
        return dados
    finally:
        page.close()


def buscar_itens_arp(uasg: str, pregao: str, max_resultados: int = 10) -> list[dict]:
    """
    Busca itens de ARP no portal contratos.sistema.gov.br.

    Args:
        uasg: código UASG (ex: "160482")
        pregao: número do pregão ou termo de busca (ex: "90005/2024")
        max_resultados: máximo de itens retornados

    Returns:
        Lista de dicts com: numero_ata, numero_item, descricao, und,
        valor_unit, valor_unit_num, fornecedor, cnpj,
        vigencia_inicio, vigencia_fim
    """
    from playwright.sync_api import sync_playwright, TimeoutError as PTimeout

    resultados = []
    exe = _caminho_chromium_sistema()

    with sync_playwright() as p:
        launch_kw: dict = {"headless": True}
        if exe:
            launch_kw["executable_path"] = exe
        try:
            browser = p.chromium.launch(**launch_kw)
        except Exception:
            browser = p.chromium.launch(headless=True)

        ctx_kw: dict = {}
        if os.path.exists(SESSION_FILE):
            ctx_kw["storage_state"] = SESSION_FILE

        ctx = browser.new_context(**ctx_kw)
        page = ctx.new_page()

        try:
            page.goto(BASE_URL, wait_until="networkidle", timeout=30000)

            # ── 1. Garante status Vigente ──────────────────────────────────
            try:
                page.check("#vigente", timeout=3000)
            except Exception:
                pass

            # ── 2. Preenche Palavra-chave (pregão) ────────────────────────
            campo_pc = page.locator("#palavra_chave")
            campo_pc.wait_for(state="visible", timeout=10000)
            campo_pc.fill(pregao)

            # ── 3. Seleciona UASG (Select2) ───────────────────────────────
            if uasg.strip():
                # Abre filtros avançados se necessário
                try:
                    campo_uf = page.locator("[name='uf']")
                    if not campo_uf.is_visible():
                        botao = page.locator("button:has-text('Filtros'), a:has-text('Filtros')")
                        if botao.count() > 0:
                            botao.first.click()
                            campo_uf.wait_for(state="visible", timeout=3000)
                except Exception:
                    pass

                try:
                    container = page.locator(".select2-selection--multiple").first
                    container.wait_for(state="visible", timeout=5000)
                    container.click()
                    page.wait_for_timeout(600)
                    page.keyboard.type(uasg, delay=100)
                    page.wait_for_timeout(1500)
                    opcao = page.locator(".select2-results__option").first
                    opcao.wait_for(state="visible", timeout=5000)
                    opcao.click()
                    page.wait_for_timeout(500)
                    # Dispara change para habilitar botão Pesquisar
                    page.evaluate("""
                        const sel = document.getElementById('unidades_gerenciadoras');
                        if (sel) {
                            sel.dispatchEvent(new Event('change', { bubbles: true }));
                            if (typeof $ !== 'undefined') $(sel).trigger('change');
                        }
                    """)
                except Exception as e:
                    logger.warning("Erro ao selecionar UASG: %s", e)

                # Aplica filtro avançado
                try:
                    botao = page.locator("#btn-aplicar-filtro")
                    if botao.count() > 0:
                        if not botao.is_enabled():
                            page.evaluate(
                                "document.getElementById('btn-aplicar-filtro')"
                                ".removeAttribute('disabled')"
                            )
                            page.wait_for_timeout(200)
                        botao.scroll_into_view_if_needed()
                        botao.click()
                    else:
                        campo_pc.press("Enter")
                except Exception:
                    campo_pc.press("Enter")
            else:
                # Busca simples sem UG
                campo_pc.press("Enter")

            page.wait_for_load_state("networkidle", timeout=20000)

            # ── 4. Aguarda tabela carregar ─────────────────────────────────
            try:
                page.wait_for_selector("#itens_processing", state="hidden", timeout=15000)
            except PTimeout:
                pass
            try:
                page.wait_for_selector("#itens tbody tr td", timeout=15000)
            except PTimeout:
                logger.warning("Tabela não carregou — sem resultados.")
                return resultados

            # ── 5. Expande para 100 por página ────────────────────────────
            try:
                campo = page.locator("select[name='itens_length']")
                if campo.count() > 0:
                    campo.first.select_option("100")
                    page.wait_for_load_state("networkidle", timeout=10000)
                    page.wait_for_selector("#itens_processing", state="hidden", timeout=10000)
            except Exception:
                pass

            # ── 6. Extrai linhas ───────────────────────────────────────────
            linhas = page.locator("#itens tbody tr").all()
            logger.info("%d linhas encontradas para '%s' / UASG %s.", len(linhas), pregao, uasg)

            for linha in linhas[:max_resultados]:
                try:
                    cels = linha.locator("td").all()
                    if len(cels) < COL_ACAO + 1:
                        continue

                    textos = [c.inner_text().strip() for c in cels]

                    # Link de detalhe
                    link = ""
                    try:
                        link_el = cels[COL_ACAO].locator("a:has(i)")
                        if link_el.count() == 0:
                            link_el = cels[COL_ACAO].locator("a[href*='show']")
                        if link_el.count() > 0:
                            href = link_el.first.get_attribute("href") or ""
                            link = href if href.startswith("http") else \
                                   "https://contratos.sistema.gov.br" + href
                    except Exception:
                        pass

                    detalhe = _extrair_detalhe(ctx, link) if link else {}

                    fornecedor = textos[COL_FORNECEDOR] if len(textos) > COL_FORNECEDOR else ""
                    if not fornecedor:
                        fornecedor = detalhe.get("fornecedor", "")

                    valor_num = detalhe.get("valor_unitario", 0.0) or \
                                _parse_br(textos[COL_SALDO] if len(textos) > COL_SALDO else "")

                    resultado = {
                        "numero_ata":      textos[COL_NUMERO_ATA]  if len(textos) > COL_NUMERO_ATA  else "",
                        "numero_item":     textos[COL_NUMERO_ITEM] if len(textos) > COL_NUMERO_ITEM else "",
                        "descricao":       textos[COL_DESCRICAO]   if len(textos) > COL_DESCRICAO   else "",
                        "und":             detalhe.get("und", ""),
                        "valor_unit_num":  valor_num,
                        "valor_unit":      _fmt(valor_num),
                        "fornecedor":      fornecedor,
                        "cnpj":            detalhe.get("cnpj", ""),
                        "vigencia_inicio": textos[COL_INICIO_VIG] if len(textos) > COL_INICIO_VIG else "",
                        "vigencia_fim":    textos[COL_FIM_VIG]    if len(textos) > COL_FIM_VIG    else "",
                    }

                    if resultado["descricao"]:
                        resultados.append(resultado)

                except Exception as e:
                    logger.warning("Erro ao processar linha: %s", e)

        except Exception as e:
            logger.error("Erro na pesquisa: %s", e)
            raise
        finally:
            browser.close()

    return resultados
