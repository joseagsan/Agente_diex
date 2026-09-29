#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
gerar_relatorio.py - Gera o "Documento de Consulta ao Comandante" (10º GAC Sl)
a partir dos CSVs exportados do Google Sheets, substituindo a edição manual
feita até agora no chat.

USO
----
    python3 gerar_relatorio.py \
        --credito 160482.csv 167482.csv \
        --consolidado Corrente_Consolidado.csv \
        --rp RP_Adaptado.csv \
        --material material_permanente.csv \
        --saida /mnt/user-data/outputs/documento_consulta_comandante.html

Todos os argumentos de arquivo são opcionais: se você não tiver uma fonte
nesta rodada (ex.: não atualizou Material Permanente), simplesmente omita a
flag correspondente e a seção respectiva não entra no relatório desta vez.

Onde ajustar as REGRAS de negócio (filtro de RESP, fórmulas de %, o que
conta como NC "residual", quais NCs são duplicatas): tudo isso está
centralizado em `core.py`, na classe `Config`. Não é preciso editar este
script para isso.

Onde ajustar o VISUAL do relatório (cores, layout, cards, boxes): edite o
template em `templates/documento_consulta_comandante.template.html`. Ele
usa marcadores simples `{{ALGO}}` que este script substitui por texto.

Justificativas manuais (o texto livre que o Cmt lê para cada NC/NE) não têm
como vir de uma planilha de forma automática -- continuam sendo um texto
que você mantém à parte. Duas opções, veja `dados/justificativas.csv`
(criado automaticamente na primeira execução, se não existir): preencha a
coluna "justificativa" para as NCs/NEs relevantes e este script vai
casar pelo número da NC/NE automaticamente.
"""

from __future__ import annotations

import argparse
import csv
import datetime as dt
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
import core  # noqa: E402


BASE_DIR = Path(__file__).resolve().parent.parent
TEMPLATE_PATH = BASE_DIR / "templates" / "documento_consulta_comandante.template.html"
JUSTIFICATIVAS_PATH = BASE_DIR / "dados" / "justificativas.csv"


# ---------------------------------------------------------------------------
# Justificativas manuais (arquivo lateral, editado à mão pelo usuário)
# ---------------------------------------------------------------------------

def carregar_justificativas() -> dict[str, str]:
    if not JUSTIFICATIVAS_PATH.exists():
        JUSTIFICATIVAS_PATH.parent.mkdir(parents=True, exist_ok=True)
        with open(JUSTIFICATIVAS_PATH, "w", newline="", encoding="utf-8") as f:
            w = csv.writer(f)
            w.writerow(["numero", "justificativa"])
        return {}
    mapa = {}
    with open(JUSTIFICATIVAS_PATH, newline="", encoding="utf-8-sig") as f:
        for row in csv.DictReader(f):
            numero = (row.get("numero") or "").strip()
            texto = (row.get("justificativa") or "").strip()
            if numero:
                mapa[numero] = texto
    return mapa


# ---------------------------------------------------------------------------
# Construção das seções HTML
# ---------------------------------------------------------------------------

def secao_sac(notas, justificativas: dict[str, str]) -> tuple[str, str]:
    if not notas:
        return "", ""

    relevantes = core.notas_com_saldo_relevante(notas)
    total_em_tela = sum(n.em_tela for n in relevantes)

    linhas = []
    for n in relevantes:
        just = justificativas.get(n.nc, "")
        linhas.append(
            "<tr>"
            f'<td class="mono">{n.nc}</td>'
            f"<td>{n.data_nc}</td>"
            f'<td class="num">{n.dias}</td>'
            f"<td>{n.finalidade}</td>"
            f"<td>{n.op}</td>"
            f'<td class="num">{core.fmt_brl(n.recebido)}</td>'
            f'<td class="valor-emtela">{core.fmt_brl(n.em_tela)}</td>'
            f"<td>{just}</td>"
            "</tr>"
        )

    tabela = f"""
    <div class="tabela-scroll">
    <table>
      <caption>NCs com Saldo Relevante em Tela ({len(relevantes)})</caption>
      <thead><tr>
        <th>NC</th><th>Data NC</th><th>Dias</th><th>Finalidade</th><th>OP</th>
        <th>Recebido</th><th>Em Tela</th><th>Justificativa</th>
      </tr></thead>
      <tbody>{''.join(linhas)}</tbody>
    </table>
    </div>"""

    cards = f"""
    <div class="resumo-cards">
      <div class="card destaque"><span class="valor">{len(relevantes)}</span><span class="rotulo">NCs com saldo relevante</span></div>
      <div class="card"><span class="valor">{core.fmt_brl(total_em_tela)}</span><span class="rotulo">Total em tela</span></div>
      <div class="card"><span class="valor">{len(notas)}</span><span class="rotulo">Total de NCs do 10º GAC Sl</span></div>
    </div>"""

    html = f"""
    <section class="assunto" id="sac">
      <h2>1. SAC - Notas de Crédito em Tela</h2>
      {cards}
      {tabela}
    </section>"""

    menu_item = '<a href="#sac">1. SAC</a>'
    return html, menu_item


def secao_indicadores(notas_credito, notas_ne) -> tuple[str, str]:
    if not notas_credito and not notas_ne:
        return "", ""

    cards = []

    if notas_credito:
        ind = core.indicador_pct_empenho_global(notas_credito)
        cards.append(f"""
      <div class="card destaque">
        <span class="valor">{core.fmt_pct(ind['pct_empenho'])}</span>
        <span class="rotulo">% Empenho (sobre Recebido Líquido)</span>
      </div>
      <div class="card"><span class="valor">{ind['qtd_ncs']}</span><span class="rotulo">NCs consideradas</span></div>
      <div class="card"><span class="valor">{core.fmt_brl(ind['recebido_liquido'])}</span><span class="rotulo">Recebido Líquido (Recebido - Recolhido)</span></div>
      <div class="card"><span class="valor">{core.fmt_brl(ind['total_empenhado'])}</span><span class="rotulo">Total Empenhado</span></div>
      <div class="card"><span class="valor">{core.fmt_brl(ind['saldo_a_empenhar'])}</span><span class="rotulo">Saldo a Empenhar</span></div>
""")

    if notas_ne:
        correntes = [n for n in notas_ne if not n.e_restos_a_pagar]
        rp = [n for n in notas_ne if n.e_restos_a_pagar]
        if correntes:
            ind_c = core.indicador_pct_liquidacao_global(correntes)
            cards.append(f"""
      <div class="card destaque">
        <span class="valor">{core.fmt_pct(ind_c['pct_liquidacao'])}</span>
        <span class="rotulo">% Liquidação Ano Corrente</span>
      </div>""")
        if rp:
            ind_rp = core.indicador_pct_liquidacao_global(rp)
            cards.append(f"""
      <div class="card destaque">
        <span class="valor">{core.fmt_pct(ind_rp['pct_liquidacao'])}</span>
        <span class="rotulo">% Liquidação Restos a Pagar</span>
      </div>""")

    html = f"""
    <section class="assunto" id="indicadores">
      <h2>3. Indicadores de Execução Orçamentária</h2>
      <div class="resumo-cards">{''.join(cards)}</div>
    </section>"""

    menu_item = '<a href="#indicadores">3. Indicadores</a>'
    return html, menu_item


def secao_empenhos_liquidacao(notas_consolidado, notas_rp, justificativas) -> tuple[str, str]:
    if not notas_consolidado and not notas_rp:
        return "", ""

    partes = []

    if notas_rp:
        pendentes_rp = core.nes_pendentes(notas_rp)
        total_rp = sum(n.a_liquidar + n.em_liquidacao for n in pendentes_rp)
        linhas_rp = []
        for n in pendentes_rp:
            just = justificativas.get(n.ne, "")
            linhas_rp.append(
                "<tr>"
                f'<td class="mono">{n.ne}</td>'
                f"<td>{n.nome_fav}</td>"
                f'<td class="num">{core.fmt_brl(n.a_liquidar + n.em_liquidacao)}</td>'
                f"<td>{n.situacao}</td>"
                f"<td>{just}</td>"
                "</tr>"
            )
        partes.append(f"""
      <h3 class="sub">Restos a Pagar (RP)</h3>
      <div class="resumo-cards">
        <div class="card destaque"><span class="valor">{len(pendentes_rp)}</span><span class="rotulo">NEs de RP pendentes</span></div>
        <div class="card"><span class="valor">{core.fmt_brl(total_rp)}</span><span class="rotulo">Valor pendente (RP)</span></div>
      </div>
      <div class="tabela-scroll">
      <table>
        <caption>Restos a Pagar pendentes ({len(pendentes_rp)})</caption>
        <thead><tr><th>NE</th><th>Favorecido</th><th>Valor Pendente</th><th>Situação</th><th>Motivo / Providências</th></tr></thead>
        <tbody>{''.join(linhas_rp)}</tbody>
      </table>
      </div>""")

    if notas_consolidado:
        pendentes = core.nes_pendentes(notas_consolidado)
        anuladas = core.nes_anuladas(notas_consolidado)
        total_pendente = sum(n.a_liquidar + n.em_liquidacao for n in pendentes)
        total_anulado = sum(n.anulado for n in anuladas)
        linhas = []
        for n in pendentes:
            linhas.append(
                "<tr>"
                f'<td class="mono">{n.ne}</td>'
                f"<td>{n.nome_fav}</td>"
                f"<td>{n.finalidade}</td>"
                f'<td class="num">{core.fmt_brl(n.a_liquidar + n.em_liquidacao)}</td>'
                f"<td>{n.situacao}</td>"
                "</tr>"
            )
        partes.append(f"""
      <h3 class="sub">Controle de NE (Ano Corrente)</h3>
      <div class="resumo-cards">
        <div class="card destaque"><span class="valor">{len(pendentes)}</span><span class="rotulo">NEs pendentes</span></div>
        <div class="card"><span class="valor">{core.fmt_brl(total_pendente)}</span><span class="rotulo">Valor pendente</span></div>
        <div class="card"><span class="valor">{len(anuladas)}</span><span class="rotulo">NEs anuladas</span></div>
        <div class="card"><span class="valor">{core.fmt_brl(total_anulado)}</span><span class="rotulo">Valor anulado</span></div>
      </div>
      <div class="tabela-scroll">
      <table>
        <caption>NEs pendentes de liquidação ({len(pendentes)})</caption>
        <thead><tr><th>NE</th><th>Favorecido</th><th>Finalidade</th><th>Valor Pendente</th><th>Situação</th></tr></thead>
        <tbody>{''.join(linhas)}</tbody>
      </table>
      </div>""")

    html = f"""
    <section class="assunto" id="empenhos">
      <h2>2. Empenhos / Liquidação</h2>
      {''.join(partes)}
    </section>"""

    menu_item = '<a href="#empenhos">2. Empenhos/Liquidação</a>'
    return html, menu_item


def secao_material_permanente(itens) -> tuple[str, str]:
    if not itens:
        return "", ""

    total = sum(i.valor_empenhado for i in itens)
    chegaram = [i for i in itens if i.chegou]
    linhas = []
    for i in itens:
        linhas.append(
            "<tr>"
            f"<td>{i.material}</td>"
            f"<td>{i.qtd}</td>"
            f'<td class="mono">{i.ne}</td>'
            f'<td class="mono">{i.nc}</td>'
            f'<td class="num">{core.fmt_brl(i.valor_empenhado)}</td>'
            f"<td>{'Sim' if i.chegou else 'Não'}</td>"
            f"<td>{i.situacao}</td>"
            "</tr>"
        )

    html = f"""
    <section class="assunto" id="material-permanente">
      <h2>4. Aquisição de Material Permanente</h2>
      <div class="resumo-cards">
        <div class="card destaque"><span class="valor">{len(itens)}</span><span class="rotulo">Itens em acompanhamento</span></div>
        <div class="card"><span class="valor">{core.fmt_brl(total)}</span><span class="rotulo">Valor total empenhado</span></div>
        <div class="card"><span class="valor">{len(chegaram)}</span><span class="rotulo">Já entregues</span></div>
      </div>
      <div class="tabela-scroll">
      <table>
        <caption>Material Permanente ({len(itens)})</caption>
        <thead><tr><th>Material</th><th>Qtd</th><th>NE</th><th>NC</th><th>Valor</th><th>Chegou</th><th>Situação</th></tr></thead>
        <tbody>{''.join(linhas)}</tbody>
      </table>
      </div>
    </section>"""

    menu_item = '<a href="#material-permanente">4. Material Permanente</a>'
    return html, menu_item


# ---------------------------------------------------------------------------
# Montagem do documento (usada pela linha de comando e pelo app Streamlit)
# ---------------------------------------------------------------------------

def montar_relatorio(notas_credito, notas_consolidado, notas_rp, itens_material,
                     justificativas: dict[str, str], agora: dt.datetime | None = None) -> str:
    """Monta o HTML final a partir dos dados já carregados. Seções sem dados
    ficam de fora; devolve "" se nenhuma seção tiver dados."""
    secoes = []
    menu = []
    for html, item in (
        secao_sac(notas_credito, justificativas),
        secao_empenhos_liquidacao(notas_consolidado, notas_rp, justificativas),
        secao_indicadores(notas_credito, notas_consolidado + notas_rp),
        secao_material_permanente(itens_material),
    ):
        if html:
            secoes.append(html)
            menu.append(item)

    if not secoes:
        return ""

    template = TEMPLATE_PATH.read_text(encoding="utf-8")
    agora = agora or dt.datetime.now()
    return (
        template
        .replace("{{SECOES}}", "\n".join(secoes))
        .replace("{{MENU_ITEMS}}", "\n    ".join(menu))
        .replace("{{DATA_ATUALIZACAO}}", agora.strftime("%d/%m/%Y"))
        .replace("{{DATA_GERACAO}}", agora.strftime("%d/%m/%Y %H:%M"))
    )


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--credito", nargs="*", default=[], help="CSV(s) de Controle de Crédito (UG 160482, 167482, ...)")
    ap.add_argument("--consolidado", default=None, help="CSV Corrente_Consolidado (Empenhos ano corrente)")
    ap.add_argument("--rp", default=None, help="CSV RP_Adaptado (Restos a Pagar)")
    ap.add_argument("--material", default=None, help="CSV de Aquisição de Material Permanente")
    ap.add_argument("--saida", default=str(BASE_DIR / "dados" / "documento_consulta_comandante.html"))
    args = ap.parse_args()

    justificativas = carregar_justificativas()

    notas_credito = core.carregar_notas_credito(*args.credito) if args.credito else []
    notas_consolidado = core.carregar_notas_empenho(args.consolidado, e_restos_a_pagar=False) if args.consolidado else []
    notas_rp = core.carregar_notas_empenho(args.rp, e_restos_a_pagar=True) if args.rp else []
    itens_material = core.carregar_material_permanente(args.material) if args.material else []

    saida_html = montar_relatorio(notas_credito, notas_consolidado, notas_rp, itens_material, justificativas)
    if not saida_html:
        print("Nenhuma fonte de dados informada -- nada a gerar. Use --credito/--consolidado/--rp/--material.",
              file=sys.stderr)
        sys.exit(1)

    caminho_saida = Path(args.saida)
    caminho_saida.parent.mkdir(parents=True, exist_ok=True)
    caminho_saida.write_text(saida_html, encoding="utf-8")
    print(f"Relatório gerado em: {caminho_saida}")
    if not JUSTIFICATIVAS_PATH.exists() or JUSTIFICATIVAS_PATH.stat().st_size < 40:
        print(f"Dica: edite {JUSTIFICATIVAS_PATH} para incluir justificativas por NC/NE.")


if __name__ == "__main__":
    main()
