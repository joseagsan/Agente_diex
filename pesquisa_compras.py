"""
Pesquisa de itens de pregão no portal comprasnet.gov.br.
Usa requests para buscar o HTML público da ATA e Claude API para extrair os dados,
seguindo o mesmo padrão do extrator_pdf.py.
"""
import json
import logging
import re

import requests

logger = logging.getLogger(__name__)

# comprasnet usa latin-1 e às vezes tem SSL velho
_HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                  "(KHTML, like Gecko) Chrome/124.0 Safari/537.36",
    "Accept-Language": "pt-BR,pt;q=0.9",
}

_PROMPT_ITENS = """Analise esta página do portal comprasnet.gov.br que contém os itens de um pregão/ATA
de Registro de Preços. Extraia TODOS os itens e retorne SOMENTE um JSON válido com a estrutura abaixo.
Não inclua texto fora do JSON.

{
  "vigencia_inicio": "DD/MM/YYYY",
  "vigencia_fim": "DD/MM/YYYY",
  "itens": [
    {
      "numero_item": "1",
      "descricao": "descrição completa do item",
      "und": "UN",
      "valor_unitario": 0.00,
      "fornecedor": "Razão Social do fornecedor",
      "cnpj": "00.000.000/0000-00"
    }
  ]
}

Use "" para campos não encontrados e 0.0 para valores não encontrados.
Formate CNPJ como XX.XXX.XXX/XXXX-XX.
"""


def _normalizar_pregao(num: str) -> str:
    """
    Normaliza o número do pregão para o formato do comprasnet: AAAANNNNN
    Aceita: "90005", "90005/2024", "2024/90005", "PE 90005/2024", etc.
    """
    import datetime
    ano_atual = str(datetime.date.today().year)

    # Remove letras e espaços do início ("PE ", "SRP ", etc.)
    num = re.sub(r"^[A-Za-z\s]+", "", num).strip()

    # Extrai ano e sequencial
    m = re.search(r"(\d{4})[/\-](\d+)", num)
    if m:
        ano, seq = m.group(1), m.group(2)
    else:
        m2 = re.search(r"(\d+)[/\-](\d{4})", num)
        if m2:
            seq, ano = m2.group(1), m2.group(2)
        else:
            seq = re.sub(r"\D", "", num)
            ano = ano_atual

    return ano + seq.zfill(5)


def _html_para_texto(html_bytes: bytes) -> str:
    """Converte HTML em texto limpo para enviar ao Claude."""
    try:
        from bs4 import BeautifulSoup
        soup = BeautifulSoup(html_bytes, "lxml")
        for tag in soup(["script", "style", "head"]):
            tag.decompose()
        texto = soup.get_text(separator="\n", strip=True)
        texto = re.sub(r"\n{3,}", "\n\n", texto)
        return texto[:15000]
    except ImportError:
        texto = re.sub(r"<[^>]+>", " ", html_bytes.decode("latin-1", errors="replace"))
        return re.sub(r"\s{2,}", " ", texto)[:15000]


def _extrair_com_claude(texto: str) -> dict:
    """Envia o texto ao Claude API e extrai JSON de itens (padrão extrator_pdf.py)."""
    from config import ANTHROPIC_API_KEY
    import anthropic

    client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)
    resp = client.messages.create(
        model="claude-haiku-4-5-20251001",
        max_tokens=4000,
        messages=[{
            "role": "user",
            "content": f"{_PROMPT_ITENS}\n\n---\nCONTEÚDO DA PÁGINA:\n{texto}",
        }],
    )
    raw = resp.content[0].text.strip()
    raw = re.sub(r"^```(?:json)?\s*", "", raw)
    raw = re.sub(r"\s*```$", "", raw)
    return json.loads(raw.strip())


def _fmt(v: float) -> str:
    return "R$ " + f"{v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


# ── URLs do comprasnet ─────────────────────────────────────────────────────────

def _urls_pregao(uasg: str, num_normalizado: str) -> list[str]:
    """Retorna lista de URLs a tentar para a ATA do pregão."""
    return [
        # ATA pública (mais completa — fornecedor + preço + vigência)
        f"https://www.comprasnet.gov.br/livre/pregao/ata0.asp"
        f"?co_no_uasg={uasg}&numprp={num_normalizado}",
        # Resultado por fornecedor (fallback)
        f"https://www.comprasnet.gov.br/livre/pregao/result_multif0.asp"
        f"?co_no_uasg={uasg}&numprp={num_normalizado}",
    ]


# ── API pública ────────────────────────────────────────────────────────────────

def buscar_itens_pregao(uasg: str, num_pregao: str) -> list[dict]:
    """
    Busca itens de um pregão no comprasnet.gov.br.

    Args:
        uasg: código UASG (ex: "160482")
        num_pregao: número do pregão (ex: "90005/2024", "90005", "PE 90005/2024")

    Returns:
        Lista de dicts com: numero_item, descricao, und,
        valor_unit_num, valor_unit, fornecedor, cnpj,
        vigencia_inicio, vigencia_fim
    """
    num_norm = _normalizar_pregao(num_pregao)
    logger.info("Buscando pregão %s (normalizado: %s) UASG %s", num_pregao, num_norm, uasg)

    html_bytes = None
    url_usada = ""

    for url in _urls_pregao(uasg, num_norm):
        try:
            resp = requests.get(url, headers=_HEADERS, timeout=30, verify=False)
            # comprasnet retorna 200 mesmo para "não encontrado" — verifica conteúdo
            if resp.status_code == 200 and len(resp.content) > 2000:
                html_bytes = resp.content
                url_usada = url
                logger.info("HTML obtido de: %s (%d bytes)", url, len(html_bytes))
                break
        except Exception as e:
            logger.warning("Falha em %s: %s", url, e)

    if not html_bytes:
        raise ValueError(
            f"Pregão {num_pregao} não encontrado para UASG {uasg}. "
            f"Verifique o número e o formato (ex: 90005/2024)."
        )

    texto = _html_para_texto(html_bytes)

    try:
        dados = _extrair_com_claude(texto)
    except Exception as e:
        raise RuntimeError(f"Erro ao extrair dados com Claude: {e}") from e

    itens_brutos = dados.get("itens", [])
    vig_inicio = dados.get("vigencia_inicio", "")
    vig_fim = dados.get("vigencia_fim", "")

    resultados = []
    for it in itens_brutos:
        v = float(it.get("valor_unitario") or 0)
        resultados.append({
            "numero_item":     str(it.get("numero_item", "")),
            "descricao":       str(it.get("descricao", "")),
            "und":             str(it.get("und", "UN")),
            "valor_unit_num":  v,
            "valor_unit":      _fmt(v),
            "fornecedor":      str(it.get("fornecedor", "")),
            "cnpj":            str(it.get("cnpj", "")),
            "vigencia_inicio": vig_inicio,
            "vigencia_fim":    vig_fim,
            "_url":            url_usada,
        })

    logger.info("%d itens extraídos.", len(resultados))
    return resultados
