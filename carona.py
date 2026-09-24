"""
Gera os documentos do processo de Adesão a Ata de Registro de Preços
("carona") — DFD, Requisição (modelo carona, com FORNECEDOR/MODALIDADE/
VIGÊNCIA DA ATA/DADOS NC) e Solicitação de Manifestação de Aceite ao
Fornecedor — a partir dos mesmos dados e itens preenchidos uma única vez
na página "Carona" do SSAC.
"""
from __future__ import annotations
import html as _html


def _tabela_itens_texto(itens: list[dict]) -> str:
    col_desc = 60
    header = f"{'ORD':<5}{'ITEM':<7}{'SI':<5}{'DESCRIÇÃO':<{col_desc}}{'UND':<6}{'QTD':<7}{'V UNIT':<12}{'V TOTAL':<14}"
    linhas = [header, "-" * len(header)]
    for it in itens:
        desc = (it.get("DESCRICAO_ITEM", "") or "")[:col_desc]
        linhas.append(
            f"{str(it.get('ORD','')):<5}{str(it.get('ITEM','')):<7}{str(it.get('SI','')):<5}"
            f"{desc:<{col_desc}}{str(it.get('UND','')):<6}{str(it.get('QTD','')):<7}"
            f"{str(it.get('VALOR_UNIT','')):<12}{str(it.get('VALOR_TOTAL','')):<14}"
        )
    return "\n".join(linhas)


def _tabela_itens_dfd_texto(itens: list[dict]) -> str:
    col_desc = 70
    header = f"{'ITEM':<7}{'DESCRIÇÃO':<{col_desc}}{'UND':<6}{'QTDE':<7}"
    linhas = [header, "-" * len(header)]
    for it in itens:
        desc = (it.get("DESCRICAO_ITEM", "") or "")[:col_desc]
        linhas.append(
            f"{str(it.get('ITEM','')):<7}{desc:<{col_desc}}{str(it.get('UND','')):<6}{str(it.get('QTD','')):<7}"
        )
    return "\n".join(linhas)


def total_itens(itens: list[dict]) -> float:
    return sum(float(it.get("_total", 0) or 0) for it in itens)


# ── Requisição (modelo carona) ─────────────────────────────────────────────
def gerar_texto_requisicao_carona(campos: dict, itens: list[dict]) -> str:
    """campos esperados: FORNECEDOR_NOME, FORNECEDOR_CNPJ, PREGAO, UASG,
    VIGENCIA_ATA, NC, DATA_NC, ND, PI, TIPO, VALOR_TOTAL_FMT,
    ALINHAMENTO (opcional), JUSTIFICATIVA_REQ (opcional)."""
    dados_nc = campos.get("NC", "")
    if campos.get("DATA_NC"):
        dados_nc += f" {campos['DATA_NC']}"

    partes = [
        "1. Solicito providências no sentido de aprovar a emissão de Nota de "
        "Empenho para aquisição do(s) item(ns) abaixo especificado(s), "
        "conforme documentação constante do processo.",
        "",
        f"FORNECEDOR: {campos.get('FORNECEDOR_CNPJ','')} - {campos.get('FORNECEDOR_NOME','')}",
        f"MODALIDADE: Pregão SRP nº {campos.get('PREGAO','')} - UASG {campos.get('UASG','')}"
        + (f" - {campos.get('ORGAO_GERENCIADOR','')}" if campos.get("ORGAO_GERENCIADOR") else ""),
        f"VIGÊNCIA DA ATA: {campos.get('VIGENCIA_ATA','')}",
        f"DADOS NC: {dados_nc}",
        f"ND: {campos.get('ND','')}    PI: {campos.get('PI','')}",
        f"TIPO: {campos.get('TIPO','Ordinário')}",
        "",
        _tabela_itens_texto(itens),
        "",
        f"TOTAL: {campos.get('VALOR_TOTAL_FMT','')}",
        "",
        f"2. Alinhamento com o plano estratégico: {campos.get('ALINHAMENTO') or 'OE1: Elevar o nível estratégico'}",
        "",
        f"3. Justificativa: {campos.get('JUSTIFICATIVA_REQ') or 'Conforme consta do DFD e ETP'}",
        "",
        "4. Declara-se que:",
        "a. A ata encontra-se vigente e válida",
        "b. Há viabilidade orçamentária para a contratação pretendida",
        "c. A contratação por adesão à referida ARP mostra-se mais vantajosa à Administração, "
        "diante da economicidade, celeridade e conveniência administrativa;",
        "d. Será garantido o cumprimento integral das condições estabelecidas na ata original, "
        "bem como os quantitativos não excederão os limites legais previstos.",
    ]
    return "\n".join(partes)


_CSS_REQ = """
<style>
.req-carona { font-family: Calibri, Arial, sans-serif; font-size: 15px; color: #111; line-height: 1.5; }
.req-carona p { margin: 0 0 10px; }
.req-carona table { border-collapse: collapse; width: 100%; margin: 10px 0 14px; }
.req-carona th, .req-carona td { border: 1px solid #444; padding: 5px 8px; font-size: 13px; text-align: left; }
.req-carona th { background: #e8e8e8; font-weight: 700; }
.req-carona td.num { text-align: right; }
.req-carona .campo { margin: 0 0 4px; }
.req-carona .campo b { display: inline-block; min-width: 150px; }
.req-carona .total-row td { font-weight: 700; background: #f4f4f4; }
</style>
"""


def gerar_html_requisicao_carona(campos: dict, itens: list[dict]) -> str:
    e = _html.escape
    dados_nc = campos.get("NC", "")
    if campos.get("DATA_NC"):
        dados_nc += f" {campos['DATA_NC']}"

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

    modalidade = f"Pregão SRP nº {e(campos.get('PREGAO',''))} - UASG {e(campos.get('UASG',''))}"
    if campos.get("ORGAO_GERENCIADOR"):
        modalidade += f" - {e(campos.get('ORGAO_GERENCIADOR',''))}"

    return f"""{_CSS_REQ}
<div class="req-carona">
<p>1. Solicito providências no sentido de aprovar a emissão de Nota de Empenho para aquisição
do(s) item(ns) abaixo especificado(s), conforme documentação constante do processo.</p>

<p class="campo"><b>FORNECEDOR:</b> {e(campos.get('FORNECEDOR_CNPJ',''))} - {e(campos.get('FORNECEDOR_NOME',''))}</p>
<p class="campo"><b>MODALIDADE:</b> {modalidade}</p>
<p class="campo"><b>VIGÊNCIA DA ATA:</b> {e(campos.get('VIGENCIA_ATA',''))}</p>
<p class="campo"><b>DADOS NC:</b> {e(dados_nc)}</p>
<p class="campo"><b>ND:</b> {e(campos.get('ND',''))} &nbsp;&nbsp; <b>PI:</b> {e(campos.get('PI',''))}</p>
<p class="campo"><b>TIPO:</b> {e(campos.get('TIPO','Ordinário'))}</p>

<table>
<thead><tr><th>ORD</th><th>ITEM</th><th>SI</th><th>DESCRIÇÃO</th><th>UND</th><th>QTD</th><th>V UNIT</th><th>V TOTAL</th></tr></thead>
<tbody>{linhas_tabela}
<tr class="total-row"><td colspan="7">TOTAL</td><td class="num">{e(campos.get('VALOR_TOTAL_FMT',''))}</td></tr>
</tbody>
</table>

<p>2. Alinhamento com o plano estratégico: {e(campos.get('ALINHAMENTO') or 'OE1: Elevar o nível estratégico')}</p>
<p>3. Justificativa: {e(campos.get('JUSTIFICATIVA_REQ') or 'Conforme consta do DFD e ETP')}</p>
<p>4. Declara-se que:</p>
<p class="item-a">a. A ata encontra-se vigente e válida</p>
<p class="item-a">b. Há viabilidade orçamentária para a contratação pretendida</p>
<p class="item-a">c. A contratação por adesão à referida ARP mostra-se mais vantajosa à Administração,
diante da economicidade, celeridade e conveniência administrativa;</p>
<p class="item-a">d. Será garantido o cumprimento integral das condições estabelecidas na ata original,
bem como os quantitativos não excederão os limites legais previstos.</p>
</div>"""


# ── DFD (Documento de Formalização da Demanda) ─────────────────────────────
def gerar_texto_dfd(campos: dict, itens: list[dict]) -> str:
    """campos esperados: JUSTIFICATIVA_DFD, PREVISAO_DATA, RESPONSAVEL
    (nome e posto), UASG, ORGAO_NOME (opcional)."""
    partes = [
        "DOCUMENTO DE FORMALIZAÇÃO DA DEMANDA",
        "",
        f"Órgão: {campos.get('ORGAO_NOME') or '10° Grupo de Artilharia de Campanha de Selva (10° GAC Sl)'}"
        f"    UASG: {campos.get('UASG_ORGAO','')}",
        "Setor Requisitante: Seção de Aquisições 10° GAC Sl",
        f"Responsável pela Demanda: {campos.get('RESPONSAVEL','')}",
        "",
        "1. Justificativa da necessidade da aquisição/contratação, considerando o Planejamento "
        "Estratégico, se for o caso.",
        campos.get("JUSTIFICATIVA_DFD", ""),
        "",
        "2. Quantidade de material/serviço a ser adquirido/contratado",
        _tabela_itens_dfd_texto(itens),
        "",
        "3. Previsão de data em que deve ser iniciada a prestação dos serviços/utilização dos materiais",
        campos.get("PREVISAO_DATA", ""),
        "",
        "4. Indicação do membro da equipe de planejamento e se necessário o responsável pela "
        "fiscalização.",
        f"Equipe de Planejamento: {campos.get('RESPONSAVEL','')} – Chefe da Seção de Aquisições",
        "Gestor de Contratos: a ser publicado posteriormente.",
        "Fiscal de Contrato: a ser publicado posteriormente.",
    ]
    return "\n".join(partes)


# ── Solicitação de manifestação de aceite ao fornecedor ────────────────────
def gerar_texto_solicitacao_fornecedor(campos: dict, itens: list[dict]) -> str:
    """campos esperados: UASG_ORGAO, ENDERECO_ORGAO, PREGAO, UASG,
    ORGAO_GERENCIADOR, FORNECEDOR_CNPJ, FORNECEDOR_NOME,
    FORNECEDOR_ENDERECO/TELEFONE/EMAIL (opcionais), ASSINANTE, CARGO."""
    partes = [
        "Senhor Fornecedor",
        "",
        f"1. Com fulcro no art. 31 do Decreto nº 11.462/23, a {campos.get('ORGAO_SOLICITANTE') or '1ª Brigada de Infantaria de Selva'}, "
        f"UASG {campos.get('UASG_ORGAO','')}, localizada na "
        f"{campos.get('ENDERECO_ORGAO') or 'Rua Marques de Pombal, s/n - Quadra 1 - Bairro Treze de Setembro, Boa Vista - RR, CEP: 69.308-515'}, "
        "manifesta interesse em aderir a seguinte Ata de Registro de Preços.",
        "",
        f"UASG {campos.get('UASG','')} - {campos.get('ORGAO_GERENCIADOR','')} PREGÃO SRP nº {campos.get('PREGAO','')}",
        "",
        f"CNPJ {campos.get('FORNECEDOR_CNPJ','')} - {campos.get('FORNECEDOR_NOME','')}",
        f"Endereço: {campos.get('FORNECEDOR_ENDERECO') or '(preencher)'}",
        f"Telefone: {campos.get('FORNECEDOR_TELEFONE') or '(preencher)'}; "
        f"E-mail: {campos.get('FORNECEDOR_EMAIL') or '(preencher)'}",
        "",
    ]
    for it in itens:
        partes.append(
            f"Do item {it.get('ITEM','')} - {it.get('QTD','')} unidades no valor unitário de "
            f"{it.get('VALOR_UNIT','')} totalizando {it.get('VALOR_TOTAL','')}"
        )
        partes.append(f"Descrição: {it.get('DESCRICAO_ITEM','')}")
        partes.append("")
    partes += [
        "2. Neste sentido, solicitamos manifestação formal quanto a aceitação da adesão a "
        "referida Ata de Registro de Preços, e nos colocamos a disposição para maiores "
        "informações.",
        "Agradecemos desde já a atenção dispensada.",
        "",
        "Atenciosamente",
        "",
        campos.get("ASSINANTE", ""),
        campos.get("CARGO") or "Chefe da Seção de Aquisições do 10º GAC Sl.",
    ]
    return "\n".join(partes)
