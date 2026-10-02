"""
Página "Consulta ao Cmt" — gera o Documento de Consulta ao Comandante
(pacote relatorio_gac) lendo as planilhas de Controle de Crédito e de
Empenhos/Liquidação direto do Google Sheets, sem precisar exportar CSVs à
mão.

As regras de negócio (filtro de RESP, % de empenho/liquidação, NCs
residuais e duplicatas) e o visual continuam em relatorio_gac/: o app só
carrega os dados e chama gerar_relatorio.montar_relatorio().
"""
import sys
from datetime import date, datetime
from pathlib import Path

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components

from config import (
    SHEET_ID_NC_ORIGEM, SHEET_ID_CONSOLIDADO,
    ABAS_NC_ORIGEM, ABA_RELATORIO_CORRENTE, ABA_RELATORIO_RP,
)

sys.path.insert(0, str(Path(__file__).resolve().parent / "relatorio_gac" / "scripts"))
import core  # noqa: E402
import gerar_relatorio  # noqa: E402

# Fontes do relatório: chave -> (rótulo, sheet_id, aba)
_FONTES = {
    **{f"credito_{aba}": (f"Controle de Crédito — {aba}", SHEET_ID_NC_ORIGEM, aba)
       for aba in ABAS_NC_ORIGEM},
    "corrente": (f"Empenhos — {ABA_RELATORIO_CORRENTE}", SHEET_ID_CONSOLIDADO, ABA_RELATORIO_CORRENTE),
    "rp":       (f"Restos a Pagar — {ABA_RELATORIO_RP}", SHEET_ID_CONSOLIDADO, ABA_RELATORIO_RP),
}


@st.cache_data(ttl=300, show_spinner=False)
def _ler_fontes() -> dict:
    """Lê cada fonte isoladamente: uma aba inacessível não impede as outras.
    Retorna {chave: {"linhas": [...], "erro": str}}."""
    from sheets_nc import ler_aba
    resultado = {}
    for chave, (_, sheet_id, aba) in _FONTES.items():
        try:
            resultado[chave] = {"linhas": ler_aba(sheet_id, aba), "erro": ""}
        except Exception as e:
            resultado[chave] = {"linhas": [], "erro": str(e) or type(e).__name__}
    return resultado


@st.cache_data(ttl=60, show_spinner=False)
def _ler_justificativas() -> dict:
    from sheets_nc import ler_justificativas
    return ler_justificativas()


@st.cache_data(ttl=300, show_spinner=False)
def _ler_historico() -> list[dict]:
    from sheets_nc import ler_historico_em_tela
    return ler_historico_em_tela()


def _historico_com_fotografia_de_hoje(notas_credito) -> dict:
    """Histórico {data: {NC: em tela}} da planilha. Na primeira geração do
    dia, grava também a fotografia de hoje (uma por dia; as seguintes do
    mesmo dia não regravam)."""
    try:
        historico = core.historico_de_linhas(_ler_historico())
    except Exception as e:
        st.warning(f"⚠️ Não foi possível ler o histórico do saldo em tela: {e}")
        return {}
    hoje = date.today()
    if notas_credito and hoje not in historico:
        from sheets_nc import registrar_fotografia_em_tela
        foto = core.fotografia_em_tela(notas_credito)
        try:
            registrar_fotografia_em_tela(core.linhas_de_fotografia(hoje, foto))
            _ler_historico.clear()
        except Exception as e:
            st.warning(f"⚠️ Não foi possível gravar a fotografia de hoje do saldo em tela: {e}")
    return historico


def page_relatorio_cmt():
    st.title("📑 Documento de Consulta ao Comandante")
    st.caption("Gerado a partir das planilhas de Controle de Crédito e de Empenhos/Liquidação.")

    if st.button("🔄 Recarregar planilhas"):
        _ler_fontes.clear()
        _ler_justificativas.clear()
        _ler_historico.clear()

    with st.spinner("Lendo planilhas…"):
        fontes = _ler_fontes()

    with st.expander("Fontes de dados", expanded=any(f["erro"] for f in fontes.values())):
        for chave, (rotulo, _, _) in _FONTES.items():
            f = fontes[chave]
            if f["erro"]:
                st.warning(f"⚠️ **{rotulo}** — não foi possível ler: {f['erro']}")
            else:
                st.markdown(f"✅ **{rotulo}** — {len(f['linhas'])} linha(s)")

    linhas_credito = [l for aba in ABAS_NC_ORIGEM for l in fontes[f"credito_{aba}"]["linhas"]]
    notas_credito = core.notas_credito_de_linhas(linhas_credito)
    notas_corrente = core.notas_empenho_de_linhas(fontes["corrente"]["linhas"], e_restos_a_pagar=False)
    notas_rp = core.notas_empenho_de_linhas(fontes["rp"]["linhas"], e_restos_a_pagar=True)

    justificativas = _secao_justificativas(notas_credito, notas_rp)

    historico = _historico_com_fotografia_de_hoje(notas_credito)

    html = gerar_relatorio.montar_relatorio(
        notas_credito, notas_corrente, notas_rp, [], justificativas, historico=historico,
    )
    if not html:
        st.error("Nenhuma fonte de dados disponível — nada a gerar.")
        return

    st.download_button(
        "⬇️ Baixar relatório (HTML)",
        data=html.encode("utf-8"),
        file_name=f"documento_consulta_comandante_{datetime.now():%Y%m%d}.html",
        mime="text/html",
        type="primary",
    )
    components.html(html, height=900, scrolling=True)


def _secao_justificativas(notas_credito, notas_rp) -> dict:
    """Editor das justificativas por NC (SAC) e NE (Restos a Pagar), salvas
    na planilha do SSAC. Retorna o dicionário atual para o relatório."""
    salvas = _ler_justificativas()

    linhas = [
        {"Número": n.nc, "Tipo": "NC em tela", "Descrição": n.finalidade,
         "Justificativa": salvas.get(n.nc, "")}
        for n in core.notas_com_saldo_relevante(notas_credito)
    ] + [
        {"Número": n.ne, "Tipo": "RP pendente", "Descrição": f"{n.nome_fav} — {n.situacao}",
         "Justificativa": salvas.get(n.ne, "")}
        for n in core.nes_pendentes(notas_rp)
    ]
    if not linhas:
        return salvas

    with st.expander(f"✏️ Justificativas ({sum(1 for l in linhas if l['Justificativa'])}/{len(linhas)} preenchidas)"):
        editado = st.data_editor(
            pd.DataFrame(linhas),
            disabled=["Número", "Tipo", "Descrição"],
            hide_index=True,
            use_container_width=True,
            key="rel_cmt_justificativas",
        )
        atuais = {**salvas, **{r["Número"]: str(r["Justificativa"] or "").strip()
                               for r in editado.to_dict("records")}}
        if st.button("💾 Salvar justificativas"):
            from sheets_nc import salvar_justificativas
            try:
                salvar_justificativas(atuais)
                _ler_justificativas.clear()
                st.success("Justificativas salvas na planilha.")
            except Exception as e:
                st.error(f"Erro ao salvar justificativas: {e}")
    return atuais
