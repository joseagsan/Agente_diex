"""
Pesquisa de itens de ARP no portal contratos.sistema.gov.br para preenchimento de REQ.
Adaptado do crawler do agente-licitacoes.
"""
import logging
import os
import re

logger = logging.getLogger(__name__)

SESSION_FILE = "session.json"

# Índices das colunas na tabela de resultados do portal
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


def garantir_navegador() -> None:
    """Instala o navegador Chromium se ainda não estiver presente."""
    import subprocess
    import sys

    # Verifica se já funciona
    try:
        from playwright.sync_api import sync_playwright
        with sync_playwright() as p:
            b = p.chromium.launch(headless=True)
            b.close()
            return
    except Exception:
        pass

    logger.info("Instalando Chromium para Playwright...")
    # Tenta com deps; se falhar (sem permissão apt), tenta sem
    for cmd in [
        [sys.executable, "-m", "playwright", "install", "chromium", "--with-deps"],
        [sys.executable, "-m", "playwright", "install", "chromium"],
    ]:
        try:
            r = subprocess.run(cmd, capture_output=True, timeout=300)
            if r.returncode == 0:
                logger.info("Chromium instalado com sucesso.")
                return
        except Exception:
            continue

    raise RuntimeError(
        "Não foi possível instalar o Chromium. "
        "Execute manualmente: playwright install chromium"
    )


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
    """Abre a página de detalhe do item e extrai valor unitário, UND, fornecedor e CNPJ."""
    page = context.new_page()
    dados = {"valor_unitario": 0.0, "und": "", "fornecedor": "", "cnpj": ""}
    try:
        page.goto(link, wait_until="networkidle", timeout=20000)
        content = page.content()

        # ── Valor unitário
        try:
            tabela = page.locator("table:has(th:has-text('Valor unitário'))")
            if tabela.count() > 0:
                td = tabela.first.locator("tbody tr:first-child td").last
                if td.count() > 0:
                    dados["valor_unitario"] = _parse_br(td.inner_text())
        except Exception:
            pass

        # ── Unidade de medida
        try:
            tabela = page.locator("table:has(th:has-text('Unidade'))")
            if tabela.count() > 0:
                td = tabela.first.locator("tbody tr:first-child td:nth-child(2)")
                if td.count() > 0:
                    dados["und"] = td.first.inner_text().strip()
        except Exception:
            pass

        # ── CNPJ
        m = re.search(r"(\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2})", content)
        if m:
            dados["cnpj"] = m.group(1)

        # ── Nome do fornecedor
        for pattern in [
            r"Fornecedor[^:]*?:\s*([^\n<]{5,100})",
            r"Razão Social[^:]*?:\s*([^\n<]{5,100})",
        ]:
            m = re.search(pattern, content, re.IGNORECASE)
            if m:
                dados["fornecedor"] = m.group(1).strip()
                break

        # Fallback: tabela de fornecedor
        if not dados["fornecedor"]:
            try:
                tabela = page.locator("table:has(th:has-text('Fornecedor'))")
                if tabela.count() > 0:
                    td = tabela.first.locator("tbody tr:first-child td:nth-child(2)")
                    if td.count() > 0:
                        dados["fornecedor"] = td.first.inner_text().strip()
            except Exception:
                pass

        return dados
    except Exception as e:
        logger.warning("Erro ao extrair detalhe de %s: %s", link, e)
        return dados
    finally:
        page.close()


def buscar_itens_arp(ug: str, descricao: str, max_resultados: int = 10) -> list[dict]:
    """
    Busca itens de ARP no portal contratos.sistema.gov.br.

    Args:
        ug: código UASG/UG (ex: "160482")
        descricao: descrição do item ou número do pregão para busca
        max_resultados: máximo de itens retornados

    Returns:
        Lista de dicts com: numero_ata, numero_item, descricao, und,
        valor_unit, valor_unit_num, fornecedor, cnpj,
        vigencia_inicio, vigencia_fim
    """
    from playwright.sync_api import sync_playwright

    resultados = []

    with sync_playwright() as p:
        try:
            browser = p.chromium.launch(headless=True, channel="chrome")
        except Exception:
            browser = p.chromium.launch(headless=True)

        ctx_kwargs = {}
        if os.path.exists(SESSION_FILE):
            ctx_kwargs["storage_state"] = SESSION_FILE

        ctx = browser.new_context(**ctx_kwargs)
        page = ctx.new_page()

        try:
            page.goto(
                "https://contratos.sistema.gov.br/transparencia/arp-item",
                wait_until="networkidle", timeout=30000,
            )

            # Abre filtros avançados
            try:
                campo_uf = page.locator("[name='uf']")
                if not campo_uf.is_visible():
                    botao = page.locator("button:has-text('Filtros'), a:has-text('Filtros')")
                    if botao.count() > 0:
                        botao.first.click()
                        campo_uf.wait_for(state="visible", timeout=3000)
            except Exception:
                pass

            # Seleciona UG
            if ug:
                try:
                    container = page.locator(".select2-selection--multiple").first
                    container.wait_for(state="visible", timeout=5000)
                    container.click()
                    page.wait_for_timeout(600)
                    page.keyboard.type(ug, delay=100)
                    page.wait_for_timeout(1500)
                    opcao = page.locator(".select2-results__option").first
                    opcao.wait_for(state="visible", timeout=5000)
                    opcao.click()
                    page.wait_for_timeout(500)
                    page.evaluate("""
                        const sel = document.getElementById('unidades_gerenciadoras');
                        if (sel) {
                            sel.dispatchEvent(new Event('change', { bubbles: true }));
                            if (typeof $ !== 'undefined') $(sel).trigger('change');
                        }
                    """)
                except Exception as e:
                    logger.warning("Erro ao selecionar UG: %s", e)

            # Preenche descrição
            try:
                page.fill("[name='descricaoItem']", descricao)
            except Exception:
                page.fill("#palavra_chave", descricao)

            # Aplica filtro
            try:
                page.click("#btn-aplicar-filtro")
            except Exception:
                page.click("button:has-text('Pesquisar')")

            page.wait_for_load_state("networkidle", timeout=20000)

            # 100 por página
            try:
                campo = page.locator("select[name='itens_length']")
                if campo.count() > 0:
                    campo.first.select_option("100")
                    page.wait_for_load_state("networkidle", timeout=10000)
            except Exception:
                pass

            # Extrai linhas
            linhas = page.locator("table#itens tbody tr, table tbody tr").all()
            logger.info("%d linhas encontradas.", len(linhas))

            for linha in linhas[:max_resultados]:
                try:
                    cels = linha.locator("td").all()
                    if len(cels) < 8:
                        continue

                    textos = [c.inner_text().strip() for c in cels]

                    # Link de detalhe
                    link = ""
                    try:
                        link_el = cels[-1].locator("a")
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
                                _parse_br(textos[7] if len(textos) > 7 else "")

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
