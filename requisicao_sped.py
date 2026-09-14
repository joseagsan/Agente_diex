"""
Gera o texto da Requisição pronto para colar no editor do SPED — segue o
modelo oficial (Requisição Nº ___/OM, itens a/b/c/d) usado desde que o
processo passou a ser feito inteiramente pelo SPED, sem anexar DOCX.
"""
from __future__ import annotations

_MESES = {
    1: "janeiro", 2: "fevereiro", 3: "março", 4: "abril", 5: "maio", 6: "junho",
    7: "julho", 8: "agosto", 9: "setembro", 10: "outubro", 11: "novembro", 12: "dezembro",
}


def _tabela_itens(itens: list[dict]) -> str:
    col_desc = 42
    header = f"{'ORD':<5}{'ITEM':<7}{'SI':<5}{'DESCRIÇÃO':<{col_desc}}{'UND':<6}{'QTD':<9}{'V UNIT':<14}{'V TOTAL':<14}"
    linhas = [header, "-" * len(header)]
    for it in itens:
        desc = (it.get("DESCRICAO_ITEM", "") or "")[:col_desc]
        linhas.append(
            f"{str(it.get('ORD','')):<5}{str(it.get('ITEM','')):<7}{str(it.get('SI','')):<5}"
            f"{desc:<{col_desc}}{str(it.get('UND','')):<6}{str(it.get('QTD','')):<9}"
            f"{str(it.get('VALOR_UNIT','')):<14}{str(it.get('VALOR_TOTAL','')):<14}"
        )
    return "\n".join(linhas)


def gerar_texto_req_sped(campos: dict, itens: list[dict]) -> str:
    """Monta o texto da Requisição no modelo do processo 100% SPED.

    campos esperados: REQ_ID, OM, DATA (objeto date, opcional), ASSUNTO,
    MODALIDADE, PREGAO, UASG, TIPO, UG_NC, ORGAO_NC, NC, DATA_NC, PI, ND,
    FORNECEDOR_NOME, FORNECEDOR_CNPJ, JUSTIFICATIVA, ASSINANTE (opcional),
    CARGO (opcional).
    """
    om     = campos.get("OM", "")
    req_id = str(campos.get("REQ_ID", "") or "").strip()
    titulo = f"Requisição Nº {req_id}/{om}" if req_id else f"Requisição Nº ___/{om}"

    data = campos.get("DATA")
    data_txt = f"Boa Vista, RR, {data.day} de {_MESES[data.month]} de {data.year}." if data else ""

    modalidade_txt = f"{campos.get('MODALIDADE','')} {campos.get('PREGAO','')} - {campos.get('UASG','')}".strip(" -")

    partes = [
        titulo,
        "",
        data_txt,
        "",
        f"Assunto: {campos.get('ASSUNTO', '')}",
        "",
        "1. Solicito providências no sentido de aprovar a emissão de Nota de "
        "Empenho para aquisição do(s) item(ns) abaixo especificado(s):",
        f"    a. Modalidade/Nº/UASG: {modalidade_txt}",
        f"    b. Tipo de Empenho: {campos.get('TIPO', '')}",
        f"    c. Disponibilidade Orçamentária - UG Emitente: {campos.get('UG_NC','')} - "
        f"{campos.get('ORGAO_NC','')}; Nº NC: {campos.get('NC','')}; "
        f"Data: {campos.get('DATA_NC','')}; PI: {campos.get('PI','')}; ND: {campos.get('ND','')}",
        "",
        f"FORNECEDOR {campos.get('FORNECEDOR_NOME','')} - CNPJ: {campos.get('FORNECEDOR_CNPJ','')}",
        _tabela_itens(itens),
        "",
        f"    d. Justificativa: {campos.get('JUSTIFICATIVA', '')}",
        "",
        "",
    ]

    assinante = str(campos.get("ASSINANTE", "") or "").strip()
    cargo     = str(campos.get("CARGO", "") or "").strip()
    partes.append(assinante if assinante else "_" * 40)
    if cargo:
        partes.append(cargo)

    return "\n".join(partes)
