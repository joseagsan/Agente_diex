"""
Extração automática de dados de NCs e REQs a partir de PDFs e HTMLs usando Claude API.
"""
import base64
import json
import logging
import re

import anthropic

from config import ANTHROPIC_API_KEY

logger = logging.getLogger(__name__)

_PROMPT_NC = """Analise este documento — pode ser uma Nota de Crédito (NC) do SIAFI, uma página HTML
do sistema militar ou qualquer exportação que contenha dados de uma NC.
Extraia os dados e retorne SOMENTE um JSON válido (sem texto adicional) com esta estrutura:

{
  "NC": "2026NC000000",
  "ORGÃO": "COTER",
  "DATA NC": "DD/MM/YYYY",
  "PI": "código PI",
  "ND": "339030",
  "PTRES": "251050",
  "OM": "10º GAC Sl",
  "PRAZO": "DD/MM/YYYY",
  "FINALIDADE": "descrição completa da finalidade",
  "OP": "nome ou código da operação",
  "RECEBIDO": 0.00,
  "SITU": "OK"
}

Use "" para campos não encontrados e 0.0 para valores não encontrados.
SITU = "EM TELA" se a NC ainda não foi creditada/recebida, "OK" caso contrário."""

_PROMPT_REQ = """Analise este documento — pode ser uma Requisição de Material/Serviço em PDF,
HTML exportado do sistema ou qualquer formato que contenha dados de uma requisição.
Extraia os dados e retorne SOMENTE um JSON válido (sem texto adicional) com esta estrutura:

{
  "REQ": "913",
  "DATA REQ": "DD/MM/YYYY",
  "NC": "2026NC000000",
  "PI": "código PI",
  "FINALIDADE": "descrição da finalidade/operação",
  "TIPO": "Ordinário",
  "EMPRESA": "Nome da Empresa LTDA",
  "DESCRIÇÃO": "descrição resumida dos itens ou serviços",
  "VALOR": 0.00,
  "SITUAÇÃO": "Pendente",
  "NE": ""
}

Use "" para campos não encontrados e 0.0 para valores não encontrados.
TIPO pode ser "Ordinário" ou "Especial". SITUAÇÃO padrão: "Pendente"."""


def _client() -> anthropic.Anthropic:
    return anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)


def _parse_json(text: str) -> dict:
    text = text.strip()
    text = re.sub(r"^```(?:json)?\s*", "", text)
    text = re.sub(r"\s*```$", "", text)
    return json.loads(text.strip())


def _html_para_texto(html_bytes: bytes) -> str:
    """Converte HTML em texto limpo para enviar ao Claude."""
    try:
        from bs4 import BeautifulSoup
        soup = BeautifulSoup(html_bytes, "lxml")
        # Remove scripts e estilos
        for tag in soup(["script", "style", "head"]):
            tag.decompose()
        texto = soup.get_text(separator="\n", strip=True)
        # Comprime linhas em branco consecutivas
        texto = re.sub(r"\n{3,}", "\n\n", texto)
        return texto[:12000]  # Limita para não estourar contexto
    except ImportError:
        # Fallback sem beautifulsoup: remove tags via regex
        texto = re.sub(r"<[^>]+>", " ", html_bytes.decode("utf-8", errors="replace"))
        texto = re.sub(r"\s{2,}", " ", texto)
        return texto[:12000]


def _extrair_com_pdf(pdf_bytes: bytes, prompt: str) -> dict:
    pdf_b64 = base64.standard_b64encode(pdf_bytes).decode("utf-8")
    response = _client().messages.create(
        model="claude-sonnet-4-6",
        max_tokens=1000,
        messages=[{
            "role": "user",
            "content": [
                {
                    "type": "document",
                    "source": {"type": "base64", "media_type": "application/pdf", "data": pdf_b64},
                },
                {"type": "text", "text": prompt},
            ],
        }],
    )
    return _parse_json(response.content[0].text)


def _extrair_com_texto(texto: str, prompt: str) -> dict:
    response = _client().messages.create(
        model="claude-sonnet-4-6",
        max_tokens=1000,
        messages=[{
            "role": "user",
            "content": f"{prompt}\n\n---\nCONTEÚDO DO DOCUMENTO:\n{texto}",
        }],
    )
    return _parse_json(response.content[0].text)


# ── API pública ───────────────────────────────────────────────────────────────

def extrair_nc(arquivo_bytes: bytes, nome_arquivo: str = "") -> dict:
    """Extrai campos de uma NC a partir de PDF ou HTML."""
    ext = nome_arquivo.lower().rsplit(".", 1)[-1] if "." in nome_arquivo else ""
    if ext in ("html", "htm"):
        texto = _html_para_texto(arquivo_bytes)
        result = _extrair_com_texto(texto, _PROMPT_NC)
    else:
        result = _extrair_com_pdf(arquivo_bytes, _PROMPT_NC)
    logger.info("NC extraída (%s): %s", ext or "pdf", result.get("NC", "N/A"))
    return result


def extrair_req(arquivo_bytes: bytes, nome_arquivo: str = "") -> dict:
    """Extrai campos de uma REQ a partir de PDF ou HTML."""
    ext = nome_arquivo.lower().rsplit(".", 1)[-1] if "." in nome_arquivo else ""
    if ext in ("html", "htm"):
        texto = _html_para_texto(arquivo_bytes)
        result = _extrair_com_texto(texto, _PROMPT_REQ)
    else:
        result = _extrair_com_pdf(arquivo_bytes, _PROMPT_REQ)
    logger.info("REQ extraída (%s): REQ %s", ext or "pdf", result.get("REQ", "N/A"))
    return result


# Mantém compatibilidade com nomes antigos
def extrair_nc_do_pdf(pdf_bytes: bytes) -> dict:
    return extrair_nc(pdf_bytes, "arquivo.pdf")

def extrair_req_do_pdf(pdf_bytes: bytes) -> dict:
    return extrair_req(pdf_bytes, "arquivo.pdf")
