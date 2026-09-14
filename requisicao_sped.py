"""
Gera o conteúdo da Requisição pronto para colar no editor do SPED — segue o
modelo oficial (itens a/b/c/d) usado desde que o processo passou a ser feito
inteiramente pelo SPED, sem anexar DOCX. O SPED já cuida do número e da data
do documento, então o conteúdo começa direto no Assunto.
"""
from __future__ import annotations
import html as _html


def _modalidade_txt(campos: dict) -> str:
    return f"{campos.get('MODALIDADE','')} {campos.get('PREGAO','')} - {campos.get('UASG','')}".strip(" -")


def _disponibilidade_txt(campos: dict) -> str:
    return (
        f"{campos.get('UG_NC','')} - {campos.get('ORGAO_NC','')}; "
        f"Nº NC: {campos.get('NC','')}; Data: {campos.get('DATA_NC','')}; "
        f"PI: {campos.get('PI','')}; ND: {campos.get('ND','')}"
    )


def _tabela_itens_texto(itens: list[dict]) -> str:
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
    """Monta o texto simples (sem formatação) da Requisição — bom para colar
    em campos que não aceitam tabela/rich text.

    campos esperados: ASSUNTO, MODALIDADE, PREGAO, UASG, TIPO, UG_NC,
    ORGAO_NC, NC, DATA_NC, PI, ND, FORNECEDOR_NOME, FORNECEDOR_CNPJ,
    JUSTIFICATIVA, ASSINANTE (opcional), CARGO (opcional).
    """
    partes = [
        f"Assunto: {campos.get('ASSUNTO', '')}",
        "",
        "1. Solicito providências no sentido de aprovar a emissão de Nota de "
        "Empenho para aquisição do(s) item(ns) abaixo especificado(s):",
        f"    a. Modalidade/Nº/UASG: {_modalidade_txt(campos)}",
        f"    b. Tipo de Empenho: {campos.get('TIPO', '')}",
        f"    c. Disponibilidade Orçamentária - UG Emitente: {_disponibilidade_txt(campos)}",
        "",
        f"FORNECEDOR {campos.get('FORNECEDOR_NOME','')} - CNPJ: {campos.get('FORNECEDOR_CNPJ','')}",
        _tabela_itens_texto(itens),
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


_CSS = """
<style>
.req-sped { font-family: Calibri, Arial, sans-serif; font-size: 15px; color: #111; line-height: 1.5; }
.req-sped p { margin: 0 0 12px; }
.req-sped .item-a { margin: 0 0 4px 24px; }
.req-sped table { border-collapse: collapse; width: 100%; margin: 10px 0 14px; }
.req-sped th, .req-sped td { border: 1px solid #444; padding: 5px 8px; font-size: 14px; text-align: left; }
.req-sped th { background: #e8e8e8; font-weight: 700; }
.req-sped td.num { text-align: right; }
.req-sped .fornecedor { font-weight: 700; margin: 14px 0 4px; }
.req-sped .assinatura { margin-top: 40px; text-align: center; }
</style>
"""


def gerar_html_req_sped(campos: dict, itens: list[dict]) -> str:
    """Monta a Requisição como HTML (parágrafos + tabela de verdade). Ao
    selecionar e copiar esse bloco renderizado no navegador, editores de
    texto rico (como o do SPED) recebem a tabela já com colunas de
    verdade — não texto com espaços que só alinha em fonte monoespaçada."""
    e = _html.escape

    linhas_tabela = "".join(
        f"<tr>"
        f"<td>{e(str(it.get('ORD','')))}</td>"
        f"<td>{e(str(it.get('ITEM','')))}</td>"
        f"<td>{e(str(it.get('SI','')))}</td>"
        f"<td>{e(str(it.get('DESCRICAO_ITEM','')))}</td>"
        f"<td>{e(str(it.get('UND','')))}</td>"
        f"<td class='num'>{e(str(it.get('QTD','')))}</td>"
        f"<td class='num'>{e(str(it.get('VALOR_UNIT','')))}</td>"
        f"<td class='num'>{e(str(it.get('VALOR_TOTAL','')))}</td>"
        f"</tr>"
        for it in itens
    )

    return f"""{_CSS}
<div class="req-sped">
<p><b>Assunto:</b> {e(campos.get('ASSUNTO',''))}</p>

<p>1. Solicito providências no sentido de aprovar a emissão de Nota de Empenho
para aquisição do(s) item(ns) abaixo especificado(s):</p>
<p class="item-a">a. Modalidade/Nº/UASG: {e(_modalidade_txt(campos))}</p>
<p class="item-a">b. Tipo de Empenho: {e(campos.get('TIPO',''))}</p>
<p class="item-a">c. Disponibilidade Orçamentária - UG Emitente: {e(_disponibilidade_txt(campos))}</p>

<p class="fornecedor">FORNECEDOR {e(campos.get('FORNECEDOR_NOME',''))} - CNPJ: {e(campos.get('FORNECEDOR_CNPJ',''))}</p>
<table>
<thead><tr>
<th>ORD</th><th>ITEM</th><th>SI</th><th>DESCRIÇÃO</th><th>UND</th><th>QTD</th><th>V UNIT</th><th>V TOTAL</th>
</tr></thead>
<tbody>{linhas_tabela}</tbody>
</table>

<p class="item-a">d. Justificativa: {e(campos.get('JUSTIFICATIVA',''))}</p>

<div class="assinatura">
<p>{e(str(campos.get('ASSINANTE','') or '').strip()) or '_' * 40}</p>
<p>{e(str(campos.get('CARGO','') or '').strip())}</p>
</div>
</div>"""
