"""
Limpeza, transformação e agrupamento dos dados brutos da planilha.

Campos de nível de item (por linha): ORD, ITEM, SI, DESCRICAO_ITEM, UND, QTD, VALOR_UNIT, VALOR_TOTAL
Campos de nível de documento (por requisition_id): todo o resto
"""
import logging
import re

logger = logging.getLogger(__name__)

# Campos que pertencem à tabela de itens (não vão para o bloco de campos do documento)
_CAMPOS_ITEM = {"ORD", "ITEM", "SI", "DESCRICAO_ITEM", "UND", "QTD", "VALOR_UNIT", "VALOR_TOTAL"}


def _formatar_cnpj(cnpj: str) -> str:
    digits = re.sub(r"\D", "", str(cnpj))
    if len(digits) == 14:
        return f"{digits[:2]}.{digits[2:5]}.{digits[5:8]}/{digits[8:12]}-{digits[12:]}"
    return cnpj


def _formatar_moeda(valor) -> str:
    try:
        v = float(str(valor).replace(",", ".").strip())
        # 1.234,56 — padrão pt-BR
        formatado = f"{v:,.2f}"          # "1,234.56"
        formatado = formatado.replace(",", "X").replace(".", ",").replace("X", ".")  # "1.234,56"
        return f"R$ {formatado}"
    except (ValueError, TypeError):
        return str(valor)


def _calcular_total(qtd, valor_unit) -> float:
    try:
        q = float(str(qtd).replace(",", ".").strip())
        v = float(str(valor_unit).replace(",", ".").strip())
        return round(q * v, 2)
    except (ValueError, TypeError):
        return 0.0


def processar(dados: list[dict]) -> dict[str, dict]:
    """
    Agrupa as linhas por `requisition_id` e retorna:
    {
        "ID": {
            "campos": { campo_doc: valor, ... },   # substituição no corpo do documento
            "itens":  [ { campo_item: valor }, ... ]  # linhas da tabela de itens
        }
    }
    """
    resultado: dict[str, dict] = {}

    for row in dados:
        req_id = str(row.get("requisition_id", "")).strip()
        if not req_id:
            logger.warning("Linha ignorada — requisition_id vazio: %s", row)
            continue

        if req_id not in resultado:
            campos = {k: str(v).strip() for k, v in row.items() if k not in _CAMPOS_ITEM}

            # Aplica máscara de CNPJ (garante formatação mesmo se vier cru da planilha)
            if "FORNECEDOR_CNPJ" in campos and campos["FORNECEDOR_CNPJ"]:
                campos["FORNECEDOR_CNPJ"] = _formatar_cnpj(campos["FORNECEDOR_CNPJ"])

            resultado[req_id] = {"campos": campos, "itens": [], "total_geral_raw": 0.0}

        # Monta linha da tabela de itens
        qtd = row.get("QTD", 0)
        valor_unit = row.get("VALOR_UNIT", 0)
        total_raw = _calcular_total(qtd, valor_unit)
        resultado[req_id]["total_geral_raw"] += total_raw

        item = {
            "ORD": str(row.get("ORD", "")).strip(),
            "ITEM": str(row.get("ITEM", "")).strip(),
            "SI": str(row.get("SI", "")).strip(),
            "DESCRICAO_ITEM": str(row.get("DESCRICAO_ITEM", "")).strip(),
            "UND": str(row.get("UND", "")).strip(),
            "QTD": str(qtd).replace(".", ","),
            "VALOR_UNIT": _formatar_moeda(valor_unit),
            "VALOR_TOTAL": _formatar_moeda(total_raw),
        }
        resultado[req_id]["itens"].append(item)

    # Preenche o campo TOTAL (documento) com a soma calculada e remove auxiliar
    for req_id, dados_req in resultado.items():
        dados_req["campos"]["TOTAL"] = _formatar_moeda(dados_req.pop("total_geral_raw"))

    logger.info("%d requisição(ões) processada(s).", len(resultado))
    return resultado
