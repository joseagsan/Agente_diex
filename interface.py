"""
SSAC — Sistema de Suporte ao Controle de NCs e Requisições.
Execute: streamlit run interface.py
"""
import logging
from datetime import date, datetime

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

st.set_page_config(
    page_title="SSAC",
    page_icon="🪖",
    layout="wide",
    initial_sidebar_state="expanded",
)

logging.basicConfig(level=logging.INFO)

from auth import logout, requer_auth
from config import ANTHROPIC_API_KEY, OM_PADRAO, UG_PADRAO, SHEET_ID_NC
from relatorios import (
    exportar_excel, kpis, ncs_por_operacao, ncs_vencendo,
    relatorio_extrato_nc, relatorio_por_empresa, relatorio_saldo_pi, relatorio_saldo_nd,
)

# ── Tema ─────────────────────────────────────────────────────────────────────
_DARK = dict(
    bg="transparent", card_bg="#0d1829", card_border="#1a2540",
    text="#e2e8f0", subtext="#64748b", accent="#34d399",
    sidebar_bg="#080e1a", sidebar_border="#1a2540",
    nav_color="#4b6080", nav_hover_bg="#0f1f35", nav_hover_color="#c8d8e8",
    active_bg="linear-gradient(135deg,#052918,#073d22)", active_border="#0a4a2840",
    active_color="#34d399",
    plot_bg="rgba(0,0,0,0)", paper_bg="rgba(0,0,0,0)",
    grid="#1a2540", font_color="#64748b",
    bar_recv="#3b82f6", bar_emp="#10b981", bar_saldo="#34d399",
)
_LIGHT = dict(
    bg="transparent", card_bg="#f8fafc", card_border="#cbd5e1",
    text="#0f172a", subtext="#475569", accent="#059669",
    sidebar_bg="#1e293b", sidebar_border="#334155",
    nav_color="#94a3b8", nav_hover_bg="#273549", nav_hover_color="#f1f5f9",
    active_bg="linear-gradient(135deg,#065f46,#047857)", active_border="#06996060",
    active_color="#d1fae5",
    plot_bg="rgba(0,0,0,0)", paper_bg="rgba(0,0,0,0)",
    grid="#cbd5e1", font_color="#475569",
    bar_recv="#1d4ed8", bar_emp="#047857", bar_saldo="#059669",
)


def _t() -> dict:
    return _LIGHT if st.session_state.get("tema_claro") else _DARK


def _inject_css():
    t = _t()
    st.markdown(f"""
<style>
#MainMenu, footer {{visibility:hidden;}}
.block-container {{padding:1.5rem 2rem 3rem !important; max-width:100% !important;}}

section[data-testid="stSidebar"] {{width:230px !important; min-width:230px !important;}}
section[data-testid="stSidebar"] > div:first-child {{
    background:{t['sidebar_bg']} !important;
    border-right:1px solid {t['sidebar_border']} !important;
}}
section[data-testid="stSidebar"] [data-testid="baseButton-secondary"] {{
    text-align:left !important; justify-content:flex-start !important;
    background:transparent !important; border:none !important;
    color:{t['nav_color']} !important; font-size:.875rem !important; font-weight:500 !important;
    padding:9px 12px !important; border-radius:8px !important; margin:1px 0 !important;
}}
section[data-testid="stSidebar"] [data-testid="baseButton-secondary"]:hover {{
    background:{t['nav_hover_bg']} !important; color:{t['nav_hover_color']} !important;
}}
section[data-testid="stSidebar"] [data-testid="baseButton-primary"] {{
    text-align:left !important; justify-content:flex-start !important;
    background:{t['active_bg']} !important; border:1px solid {t['active_border']} !important;
    color:{t['active_color']} !important; font-size:.875rem !important; font-weight:600 !important;
    padding:9px 12px !important; border-radius:8px !important; margin:1px 0 !important;
}}
section[data-testid="stSidebar"] [data-testid="baseButton-primary"]:hover {{
    filter:brightness(1.1);
}}
section[data-testid="stSidebar"] hr {{border-color:{t['sidebar_border']} !important; margin:8px 0 !important;}}

[data-testid="stDataFrame"] {{border:1px solid {t['card_border']} !important; border-radius:10px !important;}}
[data-testid="stForm"] {{background:{t['card_bg']} !important; border:1px solid {t['card_border']} !important; border-radius:12px;}}
[data-testid="stExpander"] {{border:1px solid {t['card_border']} !important; border-radius:10px !important;}}
hr {{border-color:{t['card_border']} !important;}}

/* KPI cards */
.kpi-card {{
    background:{t['card_bg']}; border:1px solid {t['card_border']};
    border-radius:12px; padding:16px 20px; position:relative; overflow:hidden;
}}
.kpi-card .kpi-label {{font-size:.75rem; color:{t['subtext']}; text-transform:uppercase; letter-spacing:.8px; margin-bottom:6px;}}
.kpi-card .kpi-value {{font-size:1.45rem; font-weight:800; color:{t['text']};}}
.kpi-card .kpi-delta {{font-size:.75rem; color:{t['subtext']}; margin-top:4px;}}
.kpi-card .kpi-stripe {{position:absolute; left:0; top:0; bottom:0; width:4px; border-radius:12px 0 0 12px;}}
""", unsafe_allow_html=True)


def _kpi_card(label: str, value: str, delta: str = "", color: str = "#34d399"):
    st.markdown(f"""
<div class="kpi-card">
  <div class="kpi-stripe" style="background:{color};"></div>
  <div style="padding-left:8px;">
    <div class="kpi-label">{label}</div>
    <div class="kpi-value">{value}</div>
    <div class="kpi-delta">{delta}</div>
  </div>
</div>""", unsafe_allow_html=True)


def _cl() -> dict:
    t = _t()
    return dict(
        paper_bgcolor=t["paper_bg"], plot_bgcolor=t["plot_bg"], font_color=t["font_color"],
        margin=dict(l=0, r=0, t=20, b=0),
        xaxis=dict(gridcolor=t["grid"], zerolinecolor=t["grid"]),
        yaxis=dict(gridcolor=t["grid"], zerolinecolor=t["grid"]),
        legend=dict(orientation="h", y=1.15, font=dict(size=11)),
    )


ORGAOS = ["COTER", "COEX", "DGO", "DEC", "DECEX", "12 RM", "Outro"]
SITUACOES_REQ = ["Pendente", "Enviada", "Aprovada", "Empenhada", "Liquidada", "Paga"]
TIPOS_RELATORIO = [
    "Resumo Geral", "NCs Detalhado", "Requisições Detalhado",
    "NCs por Operação", "NCs Próximas do Vencimento (30 dias)", "REQs Pendentes",
    "Extrato por NC", "REQs por Empresa", "Saldo por PI", "Saldo por ND",
]
NAV = [
    ("📊", "Dashboard",        "dashboard"),
    ("📋", "Notas de Crédito", "ncs"),
    ("📝", "Requisições",      "reqs"),
    ("📄", "Lançar Documento", "pdf"),
    ("📃", "Gerar REQ",        "gerar_req"),
    ("📥", "Importar Dados",   "importar"),
    ("📈", "Relatórios",       "relatorios"),
    ("🤖", "Assistente",       "assistente"),
]
LINK_SHEETS = f"https://docs.google.com/spreadsheets/d/{SHEET_ID_NC}/edit"


# ── Helpers ───────────────────────────────────────────────────────────────────
def fmt(v: float) -> str:
    return "R$ " + f"{v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def parse(s) -> float:
    try:
        return float(str(s).replace("R$", "").replace(" ", "").replace(".", "").replace(",", "."))
    except Exception:
        return 0.0


def _cor_prazo(prazo_str: str) -> str:
    try:
        n = (datetime.strptime(prazo_str, "%d/%m/%Y").date() - date.today()).days
        return "🔴" if n < 0 else "🟠" if n <= 7 else "🟡" if n <= 30 else "🟢"
    except Exception:
        return "⚪"


# ── Cache ─────────────────────────────────────────────────────────────────────
@st.cache_data(ttl=120, show_spinner=False)
def _load():
    from sheets_nc import ler_fornecedores, ler_ncs, ler_reqs
    return ler_ncs(), ler_reqs(), ler_fornecedores()


@st.cache_data(ttl=60, show_spinner=False)
def _frases(tipo: str) -> list[str]:
    from sheets_nc import ler_frases
    return ["— escrever manualmente —"] + ler_frases(tipo)


def carregar(forcar: bool = False):
    if forcar:
        _load.clear()
    try:
        return _load()
    except Exception as e:
        abas_hint = ""
        try:
            from sheets_nc import listar_abas
            abas = listar_abas()
            if abas:
                abas_hint = f"\n\nAbas disponíveis: **{', '.join(abas)}**"
        except Exception:
            pass
        st.error(f"Erro ao carregar planilha: {e}{abas_hint}")
        return [], [], []


# ── Sidebar ───────────────────────────────────────────────────────────────────
def _sidebar(ncs, reqs) -> str:
    t = _t()
    with st.sidebar:
        st.markdown(f"""
        <div style="padding:16px 8px 14px;border-bottom:1px solid {t['sidebar_border']};margin-bottom:8px;">
            <div style="font-size:1.3rem;font-weight:900;color:#e2e8f0;letter-spacing:1px;">🪖 SSAC</div>
            <div style="font-size:.72rem;color:{t['nav_color']};margin-top:3px;line-height:1.5">
                {OM_PADRAO}<br>UG {UG_PADRAO}
            </div>
        </div>""", unsafe_allow_html=True)

        current = st.session_state.get("page", "dashboard")
        for icon, label, page_id in NAV:
            if st.button(f"{icon}  {label}", key=f"nav_{page_id}",
                         use_container_width=True,
                         type="primary" if current == page_id else "secondary"):
                st.session_state["page"] = page_id
                st.rerun()

        st.divider()

        if ncs or reqs:
            ind = kpis(ncs, reqs)
            st.markdown(f"""
            <div style="padding:4px 6px 8px;font-size:.78rem;">
                <div style="color:{t['nav_color']};font-size:.63rem;text-transform:uppercase;letter-spacing:.8px;margin-bottom:8px;">Resumo</div>
                <div style="display:flex;justify-content:space-between;margin-bottom:5px;color:#8ba0b8">
                    <span>Recebido</span><span style="color:#3b82f6;font-weight:700">{fmt(ind['recebido_total'])}</span>
                </div>
                <div style="display:flex;justify-content:space-between;margin-bottom:5px;color:#8ba0b8">
                    <span>Saldo</span><span style="color:{t['accent']};font-weight:700">{fmt(ind['saldo_total'])}</span>
                </div>
                <div style="display:flex;justify-content:space-between;margin-bottom:5px;color:#8ba0b8">
                    <span>EM TELA</span><span style="color:#fbbf24;font-weight:700">{ind['ncs_em_tela']} NCs</span>
                </div>
                <div style="display:flex;justify-content:space-between;color:#8ba0b8">
                    <span>REQs pend.</span><span style="color:#a78bfa;font-weight:700">{ind['reqs_pendentes']}</span>
                </div>
            </div>""", unsafe_allow_html=True)
            st.divider()

        # Controles
        c1, c2 = st.columns(2)
        with c1:
            if st.button("🔄 Sync", use_container_width=True, help="Recarregar dados"):
                carregar(forcar=True); st.rerun()
        with c2:
            if st.button("🚪 Sair", use_container_width=True):
                logout()

        # Tema
        tema_claro = st.toggle("☀️ Tema Claro", value=st.session_state.get("tema_claro", False))
        if tema_claro != st.session_state.get("tema_claro", False):
            st.session_state["tema_claro"] = tema_claro
            st.rerun()

        # Link planilha
        st.markdown(
            f'<a href="{LINK_SHEETS}" target="_blank" style="font-size:.75rem;color:{t["nav_color"]};'
            f'text-decoration:none;display:block;padding:4px 0;">📊 Abrir Planilha ↗</a>',
            unsafe_allow_html=True,
        )

        st.caption(f"Sync {datetime.now().strftime('%H:%M:%S')}")

    return st.session_state.get("page", "dashboard")


# ── Dashboard ─────────────────────────────────────────────────────────────────
def page_dashboard(ncs, reqs):
    t = _t()
    st.title("📊 Dashboard")
    st.caption(f"{OM_PADRAO} · {datetime.today().strftime('%d/%m/%Y')}")

    # Filtros do dashboard
    with st.expander("🔍 Filtros", expanded=False):
        f1, f2, f3 = st.columns(3)
        f_org = f1.multiselect("Órgão",    sorted({nc.get("ORGÃO", "") for nc in ncs if nc.get("ORGÃO")}))
        f_op  = f2.multiselect("Operação", sorted({nc.get("OP", "")    for nc in ncs if nc.get("OP")}))
        f_sit = f3.multiselect("Status NC", ["OK", "EM TELA"])

    ncs_f = [nc for nc in ncs
             if (not f_org or nc.get("ORGÃO") in f_org)
             and (not f_op  or nc.get("OP")    in f_op)
             and (not f_sit or nc.get("SITU")  in f_sit)]

    ind = kpis(ncs_f, reqs)
    total_rec = ind["recebido_total"]
    pct_emp = (ind["empenhado_total"] / total_rec * 100) if total_rec else 0
    pct_saldo = (ind["saldo_total"] / total_rec * 100) if total_rec else 0

    # KPI cards (5 colunas)
    k1, k2, k3, k4, k5 = st.columns(5)
    with k1:
        _kpi_card("💰 Total Recebido", fmt(total_rec),
                  f"{ind['total_ncs']} NCs · {pct_saldo:.1f}% em saldo", "#3b82f6")
    with k2:
        _kpi_card("✅ Saldo Disponível", fmt(ind["saldo_total"]),
                  f"{ind['ncs_ok']} NCs OK", t["accent"])
    with k3:
        _kpi_card("📌 Empenhado", fmt(ind["empenhado_total"]),
                  f"{pct_emp:.1f}% do recebido", "#f59e0b")
    with k4:
        _kpi_card("🕐 Em Tela", fmt(ind["em_tela_total"]),
                  f"{ind['ncs_em_tela']} NCs aguardando", "#fbbf24")
    with k5:
        _kpi_card("📝 REQs Pendentes", str(ind["reqs_pendentes"]),
                  fmt(ind["valor_reqs_pendentes"]), "#a78bfa")

    st.markdown("<div style='height:12px'></div>", unsafe_allow_html=True)

    # Alertas
    if ind["vencidas"] > 0:
        st.error(f"⛔ **{ind['vencidas']}** NC(s) com prazo vencido!")
    if ind["vencendo_7d"] > 0:
        st.warning(f"⚠️ **{ind['vencendo_7d']}** NC(s) vencem nos próximos 7 dias.")
    if ind["ncs_em_tela"] > 0:
        st.info(f"🕐 **{ind['ncs_em_tela']}** NC(s) EM TELA — {fmt(ind['em_tela_total'])} aguardando crédito.")

    st.divider()

    # Gráfico 1: Valores por mês + Status NCs
    c1, c2 = st.columns([3, 2])
    with c1:
        st.subheader("📅 Valores Recebidos por Mês")
        rows = []
        for nc in ncs_f:
            try:
                d = datetime.strptime(nc.get("DATA NC", ""), "%d/%m/%Y")
                rows.append({"Mês": d.strftime("%b/%y"), "Ord": d.strftime("%Y-%m"),
                             "Recebido": parse(nc.get("RECEBIDO", 0)),
                             "Empenhado": parse(nc.get("EMPENHADO", 0))})
            except Exception:
                pass
        if rows:
            df_m = pd.DataFrame(rows).groupby(["Mês", "Ord"]).sum().reset_index().sort_values("Ord")
            fig = go.Figure()
            fig.add_bar(x=df_m["Mês"], y=df_m["Recebido"],  name="Recebido",  marker_color=t["bar_recv"])
            fig.add_bar(x=df_m["Mês"], y=df_m["Empenhado"], name="Empenhado", marker_color=t["bar_emp"])
            fig.update_layout(height=300, barmode="group", **_cl())
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("Sem dados de data para gráfico mensal.")

    with c2:
        st.subheader("🔵 Status das NCs")
        if ncs_f:
            cnt = pd.DataFrame(ncs_f)["SITU"].value_counts().reset_index()
            cnt.columns = ["Status", "Qtd"]
            fig = px.pie(cnt, values="Qtd", names="Status", hole=0.55,
                         color_discrete_map={"OK": t["bar_emp"], "EM TELA": "#fbbf24"})
            fig.update_layout(height=300, **{**_cl(), "margin": dict(l=0, r=0, t=20, b=30)})
            fig.update_traces(textfont_color="#e2e8f0", textinfo="percent+label")
            st.plotly_chart(fig, use_container_width=True)

    # Gráfico 2: Saldo por operação + Empenho por órgão
    c3, c4 = st.columns(2)
    with c3:
        st.subheader("🎯 Saldo por Operação")
        if ncs_f:
            df_op = (pd.DataFrame([{"OP": nc.get("OP") or "Sem OP",
                                     "Saldo": parse(nc.get("SALDO NC", 0))} for nc in ncs_f])
                     .groupby("OP")["Saldo"].sum().reset_index()
                     .query("Saldo > 0").sort_values("Saldo"))
            if not df_op.empty:
                fig = px.bar(df_op, x="Saldo", y="OP", orientation="h",
                             color="Saldo", color_continuous_scale=["#1a4a2e", t["bar_emp"]],
                             labels={"Saldo": "R$", "OP": ""})
                fig.update_coloraxes(showscale=False)
                fig.update_layout(height=320, **_cl())
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.success("Sem saldo disponível no momento.")

    with c4:
        st.subheader("🏛️ Recebido vs Empenhado por Órgão")
        if ncs_f:
            df_o = (pd.DataFrame([{"ORGÃO": nc.get("ORGÃO") or "N/A",
                                    "Recebido": parse(nc.get("RECEBIDO", 0)),
                                    "Empenhado": parse(nc.get("EMPENHADO", 0))} for nc in ncs_f])
                    .groupby("ORGÃO").sum().reset_index()
                    .melt(id_vars="ORGÃO", var_name="Tipo", value_name="Valor"))
            fig = px.bar(df_o, x="ORGÃO", y="Valor", color="Tipo", barmode="group",
                         color_discrete_map={"Recebido": t["bar_recv"], "Empenhado": t["bar_emp"]},
                         labels={"Valor": "R$", "ORGÃO": ""})
            fig.update_layout(height=320, **_cl())
            st.plotly_chart(fig, use_container_width=True)

    # Gráfico 3: REQs por situação + NCs EM TELA
    c5, c6 = st.columns(2)
    with c5:
        st.subheader("📝 REQs por Situação")
        if reqs:
            df_s = (pd.DataFrame([{"Situação": r.get("SITUAÇÃO") or "N/A",
                                    "Valor": parse(r.get("VALOR", 0))} for r in reqs])
                    .groupby("Situação")["Valor"].sum().reset_index()
                    .sort_values("Valor", ascending=False))
            cores = {"Pendente": "#fbbf24", "Enviada": "#a78bfa", "Aprovada": "#60a5fa",
                     "Empenhada": t["bar_recv"], "Liquidada": t["bar_emp"], "Paga": t["bar_saldo"]}
            fig = px.bar(df_s, x="Situação", y="Valor", color="Situação",
                         color_discrete_map=cores, labels={"Valor": "R$", "Situação": ""})
            fig.update_layout(height=300, **_cl(), showlegend=False)
            st.plotly_chart(fig, use_container_width=True)

    with c6:
        st.subheader("🕐 NCs EM TELA")
        em_tela = [nc for nc in ncs_f if nc.get("SITU") == "EM TELA"]
        if em_tela:
            total_et = sum(parse(nc.get("RECEBIDO", 0)) for nc in em_tela)
            st.warning(f"⚠️ {len(em_tela)} NCs em tela — {fmt(total_et)}")
            df_et = pd.DataFrame([{
                "NC": nc.get("NC", ""), "ÓRGÃO": nc.get("ORGÃO", ""),
                "Finalidade": nc.get("FINALIDADE", "")[:38],
                "Valor": nc.get("RECEBIDO", ""), "Prazo": nc.get("PRAZO", ""),
            } for nc in em_tela])
            st.dataframe(df_et, use_container_width=True, hide_index=True, height=240)
        else:
            st.success("✅ Nenhuma NC EM TELA no momento.")

    # Gráfico 4: Valores por operação
    st.subheader("💹 Valores por Operação")
    if ncs_f:
        df_vop = (pd.DataFrame([{"OP": nc.get("OP") or "Sem OP",
                                  "Recebido": parse(nc.get("RECEBIDO", 0)),
                                  "Empenhado": parse(nc.get("EMPENHADO", 0)),
                                  "Saldo": parse(nc.get("SALDO NC", 0))} for nc in ncs_f])
                  .groupby("OP").sum().reset_index().sort_values("Recebido", ascending=False))
        fig = go.Figure()
        fig.add_bar(x=df_vop["OP"], y=df_vop["Recebido"],  name="Recebido",  marker_color=t["bar_recv"])
        fig.add_bar(x=df_vop["OP"], y=df_vop["Empenhado"], name="Empenhado", marker_color="#fbbf24")
        fig.add_bar(x=df_vop["OP"], y=df_vop["Saldo"],     name="Saldo",     marker_color=t["bar_emp"])
        fig.update_layout(barmode="group", height=300, **_cl())
        st.plotly_chart(fig, use_container_width=True)

    # Controle de prazos
    st.subheader("⏰ Controle de Prazos")
    col_link, _ = st.columns([1, 5])
    with col_link:
        st.markdown(
            f'<a href="{LINK_SHEETS}" target="_blank">'
            f'<button style="background:#1a2540;color:#94a3b8;border:1px solid #1a2540;'
            f'border-radius:6px;padding:4px 12px;font-size:.8rem;cursor:pointer;">'
            f'📊 Abrir Planilha</button></a>',
            unsafe_allow_html=True,
        )
    if ncs_f:
        hoje = date.today()
        rows = []
        for nc in ncs_f:
            p = nc.get("PRAZO", "")
            try:
                dias = (datetime.strptime(p, "%d/%m/%Y").date() - hoje).days
            except Exception:
                dias = None
            rows.append({"": _cor_prazo(p), "NC": nc.get("NC", ""),
                         "ÓRGÃO": nc.get("ORGÃO", ""),
                         "Finalidade": nc.get("FINALIDADE", "")[:45],
                         "Prazo": p,
                         "Dias": dias if dias is not None else "—",
                         "Saldo": nc.get("SALDO NC", ""),
                         "Status": nc.get("SITU", "")})
        df_p = pd.DataFrame(rows).sort_values("Dias", key=lambda x: pd.to_numeric(x, errors="coerce"))
        st.dataframe(df_p, use_container_width=True, hide_index=True)


# ── NCs ───────────────────────────────────────────────────────────────────────
def page_ncs(ncs):
    st.title("📋 Notas de Crédito")

    with st.expander("🔍 Filtros", expanded=False):
        f1, f2, f3 = st.columns(3)
        f_situ  = f1.multiselect("Status",   ["OK", "EM TELA"], default=["OK", "EM TELA"])
        f_org   = f2.multiselect("Órgão",    sorted({nc.get("ORGÃO", "") for nc in ncs if nc.get("ORGÃO")}))
        f_op    = f3.multiselect("Operação", sorted({nc.get("OP", "")    for nc in ncs if nc.get("OP")}))
        f4, f5, f6 = st.columns(3)
        f_nd    = f4.multiselect("ND",       sorted({nc.get("ND", "")    for nc in ncs if nc.get("ND")}))
        f_pi    = f5.multiselect("PI",       sorted({nc.get("PI", "")    for nc in ncs if nc.get("PI")}))
        f_prazo = f6.selectbox("Prazo",      ["Todos", "Vencidas", "Vencem em 7 dias", "Vencem em 30 dias"])

    hoje = date.today()
    filtradas = []
    for nc in ncs:
        if f_situ  and nc.get("SITU")  not in f_situ:  continue
        if f_org   and nc.get("ORGÃO") not in f_org:   continue
        if f_op    and nc.get("OP")    not in f_op:    continue
        if f_nd    and nc.get("ND")    not in f_nd:    continue
        if f_pi    and nc.get("PI")    not in f_pi:    continue
        if f_prazo != "Todos":
            if nc.get("SITU") != "EM TELA":
                continue
            try:
                dias = (datetime.strptime(nc.get("PRAZO", ""), "%d/%m/%Y").date() - hoje).days
                if f_prazo == "Vencidas"           and dias >= 0:          continue
                if f_prazo == "Vencem em 7 dias"   and not 0 <= dias <= 7:  continue
                if f_prazo == "Vencem em 30 dias"  and not 0 <= dias <= 30: continue
            except Exception:
                continue
        filtradas.append(nc)

    b1, b2, b3 = st.columns([1, 1, 3])
    with b1:
        if st.button("➕ Nova NC", type="primary", use_container_width=True):
            st.session_state["form_nc"] = not st.session_state.get("form_nc", False)
    with b2:
        st.markdown(
            f'<a href="{LINK_SHEETS}" target="_blank">'
            f'<button style="background:transparent;color:#64748b;border:1px solid #1a2540;'
            f'border-radius:6px;padding:6px 12px;font-size:.85rem;cursor:pointer;width:100%;">'
            f'📊 Planilha</button></a>',
            unsafe_allow_html=True,
        )
    with b3:
        saldo_f = sum(parse(nc.get("SALDO NC", 0)) for nc in filtradas)
        receb_f = sum(parse(nc.get("RECEBIDO", 0)) for nc in filtradas)
        st.info(f"📋 **{len(filtradas)}** NCs · Recebido: **{fmt(receb_f)}** · Saldo: **{fmt(saldo_f)}**")

    if st.session_state.get("form_nc"):
        _form_nc()

    if filtradas:
        cols = [c for c in ["ORD", "SITU", "ORGÃO", "DATA NC", "NC", "PI", "ND", "FINALIDADE",
                             "OP", "PRAZO", "DIAS", "RECEBIDO", "SALDO NC", "EMPENHADO", "EMP %", "SITUAÇÃO"]
                if c in pd.DataFrame(filtradas).columns]
        st.dataframe(pd.DataFrame(filtradas)[cols], use_container_width=True, hide_index=True)
    else:
        st.info("Nenhuma NC encontrada.")


def _form_nc():
    st.divider()
    st.subheader("➕ Nova Nota de Crédito")
    with st.form("f_nc", clear_on_submit=True):
        c1, c2 = st.columns(2)
        nc_num  = c1.text_input("Número NC *",   placeholder="2026NC000000")
        orgao   = c1.selectbox("Órgão *",        ORGAOS)
        data_nc = c1.date_input("Data NC *",      value=date.today())
        pi      = c1.text_input("PI")
        nd      = c1.text_input("ND",            placeholder="339030")
        ptres   = c2.text_input("PTRES",         placeholder="251050")
        om      = c2.text_input("OM",            value=OM_PADRAO)
        prazo   = c2.date_input("Prazo *")
        op      = c2.text_input("Operação (OP)")
        situ    = c2.selectbox("Status",         ["OK", "EM TELA"])
        finalidade = st.text_area("Finalidade *")
        valor = st.number_input("Valor (R$) *", min_value=0.0, step=0.01, format="%.2f")
        s1, s2 = st.columns(2)
        salvar   = s1.form_submit_button("💾 Salvar", type="primary", use_container_width=True)
        cancelar = s2.form_submit_button("✖ Cancelar", use_container_width=True)
        if cancelar:
            st.session_state["form_nc"] = False; st.rerun()
        if salvar:
            if not nc_num or not finalidade or valor <= 0:
                st.error("Preencha NC, Finalidade e Valor.")
            else:
                try:
                    from sheets_nc import adicionar_nc
                    adicionar_nc({"NC": nc_num, "ORGÃO": orgao,
                                  "DATA NC": data_nc.strftime("%d/%m/%Y"),
                                  "PI": pi, "ND": nd, "PTRES": ptres, "OM": om,
                                  "PRAZO": prazo.strftime("%d/%m/%Y"), "OP": op,
                                  "SITU": situ, "FINALIDADE": finalidade, "RECEBIDO": valor})
                    st.success(f"✅ NC {nc_num} adicionada!")
                    st.session_state["form_nc"] = False
                    carregar(forcar=True); st.rerun()
                except Exception as e:
                    st.error(f"Erro: {e}")


# ── Requisições ───────────────────────────────────────────────────────────────
def page_reqs(reqs, ncs):
    st.title("📝 Requisições")

    with st.expander("🔍 Filtros", expanded=False):
        f1, f2, f3 = st.columns(3)
        f_sit  = f1.multiselect("Situação", sorted({r.get("SITUAÇÃO", "") for r in reqs if r.get("SITUAÇÃO")}))
        f_emp  = f2.multiselect("Empresa",  sorted({r.get("EMPRESA", "")  for r in reqs if r.get("EMPRESA")}))
        f_nc   = f3.multiselect("NC",       sorted({r.get("NC", "")       for r in reqs if r.get("NC")}))
        f4, f5, f6 = st.columns(3)
        f_tipo = f4.multiselect("Tipo",     ["Ordinário", "Especial"])
        f_pi   = f5.multiselect("PI",       sorted({r.get("PI", "") for r in reqs if r.get("PI")}))
        vals   = [parse(r.get("VALOR", 0)) for r in reqs if parse(r.get("VALOR", 0)) > 0]
        v_max  = max(vals) if vals else 100000.0
        f_val  = f6.slider("Faixa de valor (R$)", 0.0, float(v_max), (0.0, float(v_max)), step=500.0)

    filtradas = [r for r in reqs
                 if (not f_sit  or r.get("SITUAÇÃO") in f_sit)
                 and (not f_emp or r.get("EMPRESA")  in f_emp)
                 and (not f_nc  or r.get("NC")       in f_nc)
                 and (not f_tipo or r.get("TIPO")    in f_tipo)
                 and (not f_pi  or r.get("PI")       in f_pi)
                 and f_val[0] <= parse(r.get("VALOR", 0)) <= f_val[1]]

    b1, b2, b3 = st.columns([1, 1, 3])
    with b1:
        if st.button("➕ Nova REQ", type="primary", use_container_width=True):
            st.session_state["form_req"] = not st.session_state.get("form_req", False)
    with b2:
        st.markdown(
            f'<a href="{LINK_SHEETS}" target="_blank">'
            f'<button style="background:transparent;color:#64748b;border:1px solid #1a2540;'
            f'border-radius:6px;padding:6px 12px;font-size:.85rem;cursor:pointer;width:100%;">'
            f'📊 Planilha</button></a>',
            unsafe_allow_html=True,
        )
    with b3:
        total = sum(parse(r.get("VALOR", 0)) for r in filtradas)
        st.info(f"📝 **{len(filtradas)}** REQs · Total: **{fmt(total)}**")

    if st.session_state.get("form_req"):
        _form_req(ncs)

    if filtradas:
        cols = [c for c in ["REQ", "DATA REQ", "NC", "PI", "EMPRESA", "DESCRIÇÃO",
                             "VALOR", "SITUAÇÃO", "NE", "FINALIDADE", "OBS"]
                if c in pd.DataFrame(filtradas).columns]
        st.dataframe(pd.DataFrame(filtradas)[cols], use_container_width=True, hide_index=True)
    else:
        st.info("Nenhuma REQ encontrada.")


def _form_req(ncs):
    st.divider()
    st.subheader("➕ Nova Requisição")
    nums_nc = [""] + [nc.get("NC", "") for nc in ncs if nc.get("NC")]
    with st.form("f_req", clear_on_submit=True):
        c1, c2 = st.columns(2)
        req_num  = c1.text_input("Número REQ")
        nc_sel   = c1.selectbox("NC Vinculada *", nums_nc)
        data_req = c1.date_input("Data REQ",  value=date.today())
        empresa  = c1.text_input("Empresa *")
        pi       = c2.text_input("PI")
        tipo     = c2.selectbox("Tipo",      ["Ordinário", "Especial"])
        ne       = c2.text_input("NE")
        situacao = c2.selectbox("Situação",  SITUACOES_REQ)
        finalidade = st.text_area("Finalidade")
        descricao  = st.text_area("Descrição *")
        valor = st.number_input("Valor (R$) *", min_value=0.0, step=0.01, format="%.2f")
        obs   = st.text_input("Observações")
        s1, s2 = st.columns(2)
        salvar   = s1.form_submit_button("💾 Salvar", type="primary", use_container_width=True)
        cancelar = s2.form_submit_button("✖ Cancelar", use_container_width=True)
        if cancelar:
            st.session_state["form_req"] = False; st.rerun()
        if salvar:
            if not empresa or not descricao or valor <= 0:
                st.error("Preencha Empresa, Descrição e Valor.")
            else:
                try:
                    pi_f = pi; fin_f = finalidade
                    if nc_sel and not pi:
                        nc_d = next((nc for nc in ncs if nc.get("NC") == nc_sel), {})
                        pi_f = nc_d.get("PI", ""); fin_f = fin_f or nc_d.get("OP", "")
                    from sheets_nc import adicionar_req
                    adicionar_req({"REQ": req_num, "NC": nc_sel,
                                   "DATA REQ": data_req.strftime("%d/%m/%Y"),
                                   "PI": pi_f, "FINALIDADE": fin_f, "TIPO": tipo,
                                   "EMPRESA": empresa, "DESCRIÇÃO": descricao,
                                   "VALOR": valor, "SITUAÇÃO": situacao, "NE": ne, "OBS": obs})
                    st.success("✅ REQ adicionada!")
                    st.session_state["form_req"] = False
                    carregar(forcar=True); st.rerun()
                except Exception as e:
                    st.error(f"Erro: {e}")


# ── Lançar por PDF / HTML ─────────────────────────────────────────────────────
def page_pdf(ncs):
    st.title("📄 Lançar por PDF ou HTML")
    if not ANTHROPIC_API_KEY:
        st.error("⚠️ Configure ANTHROPIC_API_KEY para usar este recurso.")
        return
    tipo = st.radio("Tipo", ["Nota de Crédito (NC)", "Requisição (REQ)"], horizontal=True)
    uploaded = st.file_uploader("Selecione o arquivo", type=["pdf", "html", "htm"])
    if uploaded:
        with st.spinner("🤖 Analisando documento..."):
            try:
                from extrator_pdf import extrair_nc, extrair_req
                b = uploaded.read()
                if "NC" in tipo:
                    dados = extrair_nc(b, uploaded.name)
                    st.success("✅ Dados extraídos — revise e confirme.")
                    _confirmar_nc(dados)
                else:
                    dados = extrair_req(b, uploaded.name)
                    st.success("✅ Dados extraídos — revise e confirme.")
                    _confirmar_req(dados, ncs)
            except Exception as e:
                st.error(f"Erro na extração: {e}")


def _confirmar_nc(dados):
    st.divider(); st.subheader("✅ Confirmar NC")
    with st.form("f_pdf_nc"):
        c1, c2 = st.columns(2)
        nc      = c1.text_input("Número NC",            value=dados.get("NC", ""))
        orgao   = c1.text_input("Órgão",                value=dados.get("ORGÃO", ""))
        data_nc = c1.text_input("Data NC (DD/MM/YYYY)", value=dados.get("DATA NC", ""))
        pi      = c1.text_input("PI",                   value=dados.get("PI", ""))
        nd      = c1.text_input("ND",                   value=dados.get("ND", ""))
        ptres   = c2.text_input("PTRES",                value=dados.get("PTRES", ""))
        om      = c2.text_input("OM",                   value=dados.get("OM", OM_PADRAO))
        prazo   = c2.text_input("Prazo (DD/MM/YYYY)",   value=dados.get("PRAZO", ""))
        op      = c2.text_input("Operação",             value=dados.get("OP", ""))
        situ    = c2.selectbox("Status", ["OK", "EM TELA"], index=1 if dados.get("SITU") == "EM TELA" else 0)
        finalidade = st.text_area("Finalidade", value=dados.get("FINALIDADE", ""))
        valor = st.number_input("Valor (R$)", value=float(dados.get("RECEBIDO", 0) or 0),
                                min_value=0.0, step=0.01, format="%.2f")
        if st.form_submit_button("💾 Confirmar e Salvar", type="primary", use_container_width=True):
            try:
                from sheets_nc import adicionar_nc
                adicionar_nc({"NC": nc, "ORGÃO": orgao, "DATA NC": data_nc, "PI": pi, "ND": nd,
                              "PTRES": ptres, "OM": om, "PRAZO": prazo, "OP": op, "SITU": situ,
                              "FINALIDADE": finalidade, "RECEBIDO": valor})
                st.success(f"✅ NC {nc} lançada!")
                carregar(forcar=True)
            except Exception as e:
                st.error(f"Erro: {e}")


def _confirmar_req(dados, ncs):
    st.divider(); st.subheader("✅ Confirmar REQ")
    nums_nc = [""] + [nc.get("NC", "") for nc in ncs if nc.get("NC")]
    nc_idx = nums_nc.index(dados.get("NC", "")) if dados.get("NC", "") in nums_nc else 0
    with st.form("f_pdf_req"):
        c1, c2 = st.columns(2)
        req      = c1.text_input("Número REQ",            value=str(dados.get("REQ", "")))
        nc_sel   = c1.selectbox("NC Vinculada",           nums_nc, index=nc_idx)
        data_req = c1.text_input("Data REQ (DD/MM/YYYY)", value=dados.get("DATA REQ", ""))
        empresa  = c1.text_input("Empresa",               value=dados.get("EMPRESA", ""))
        pi       = c2.text_input("PI",                    value=dados.get("PI", ""))
        tipo     = c2.selectbox("Tipo", ["Ordinário", "Especial"])
        ne       = c2.text_input("NE",                    value=dados.get("NE", ""))
        situacao = c2.selectbox("Situação", SITUACOES_REQ)
        finalidade = st.text_area("Finalidade", value=dados.get("FINALIDADE", ""))
        descricao  = st.text_area("Descrição",  value=dados.get("DESCRIÇÃO", ""))
        valor = st.number_input("Valor (R$)", value=float(dados.get("VALOR", 0) or 0),
                                min_value=0.0, step=0.01, format="%.2f")
        if st.form_submit_button("💾 Confirmar e Salvar", type="primary", use_container_width=True):
            try:
                from sheets_nc import adicionar_req
                adicionar_req({"REQ": req, "NC": nc_sel, "DATA REQ": data_req, "PI": pi,
                               "FINALIDADE": finalidade, "TIPO": tipo, "EMPRESA": empresa,
                               "DESCRIÇÃO": descricao, "VALOR": valor, "SITUAÇÃO": situacao, "NE": ne})
                st.success("✅ REQ lançada!")
                carregar(forcar=True)
            except Exception as e:
                st.error(f"Erro: {e}")


# ── Importar Dados ────────────────────────────────────────────────────────────
def page_importar(ncs, reqs):
    st.title("📥 Importar Dados")

    k1, k2, k3, k4 = st.columns(4)
    k1.metric("NCs carregadas",  len(ncs))
    k2.metric("REQs carregadas", len(reqs))
    k3.metric("NCs OK",          sum(1 for nc in ncs if nc.get("SITU") == "OK"))
    k4.metric("REQs Pendentes",  sum(1 for r  in reqs if r.get("SITUAÇÃO") == "Pendente"))

    st.info(f"🔗 Planilha: `{SHEET_ID_NC[:30]}...`  ·  Sync automático a cada 2 min.")
    st.markdown(f'<a href="{LINK_SHEETS}" target="_blank">📊 Abrir planilha no Google Sheets ↗</a>',
                unsafe_allow_html=True)

    if st.button("🔄 Forçar Sincronização", type="primary"):
        with st.spinner("Sincronizando..."):
            carregar(forcar=True)
        st.success("✅ Dados atualizados!"); st.rerun()

    st.divider()
    st.subheader("📂 Importar de Arquivo (Excel / CSV)")
    tipo_imp = st.radio("O que importar?", ["Notas de Crédito (NCs)", "Requisições (REQs)"], horizontal=True)
    arq = st.file_uploader("Selecione o arquivo", type=["xlsx", "xls", "csv"])

    if arq:
        try:
            df_imp = pd.read_csv(arq, dtype=str).fillna("") if arq.name.endswith(".csv") \
                     else pd.read_excel(arq, dtype=str).fillna("")
            st.write(f"**{len(df_imp)} linhas encontradas** — prévia:")
            st.dataframe(df_imp.head(10), use_container_width=True, hide_index=True)

            if st.button(f"⬆️ Importar {len(df_imp)} linha(s)", type="primary"):
                from sheets_nc import adicionar_nc, adicionar_req
                erros = 0
                bar = st.progress(0)
                for i, (_, row) in enumerate(df_imp.iterrows()):
                    try:
                        if "NCs" in tipo_imp:
                            adicionar_nc(row.to_dict())
                        else:
                            adicionar_req(row.to_dict())
                    except Exception:
                        erros += 1
                    bar.progress((i + 1) / len(df_imp))
                carregar(forcar=True)
                msg = f"✅ {len(df_imp) - erros} importadas"
                if erros:
                    st.warning(f"{msg}, {erros} com erro.")
                else:
                    st.success(f"{msg} com sucesso!")
        except Exception as e:
            st.error(f"Erro ao ler arquivo: {e}")

    st.divider()
    st.subheader("👁️ Preview dos Dados Atuais")
    aba = st.radio("Ver", ["NCs", "REQs"], horizontal=True)
    dados = ncs if aba == "NCs" else reqs
    if dados:
        st.dataframe(pd.DataFrame(dados), use_container_width=True, hide_index=True)
    else:
        st.info("Sem dados.")


# ── Relatórios ────────────────────────────────────────────────────────────────
def page_relatorios(ncs, reqs):
    st.title("📈 Relatórios")

    c_tipo, c_exp = st.columns([3, 1])
    with c_tipo:
        tipo = st.selectbox("Tipo de Relatório", TIPOS_RELATORIO)
    with c_exp:
        st.markdown("<div style='height:28px'></div>", unsafe_allow_html=True)
        st.download_button(
            "⬇️ Excel Completo",
            data=exportar_excel(ncs, reqs),
            file_name=f"ssac_{datetime.today().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    st.divider()
    _exibir_relatorio(tipo, ncs, reqs)


def _exibir_relatorio(tipo, ncs, reqs):
    if tipo == "Resumo Geral":
        ind = kpis(ncs, reqs)
        c1, c2 = st.columns(2)
        with c1:
            st.subheader("📋 NCs")
            for l, v in [("Total", ind["total_ncs"]), ("OK", ind["ncs_ok"]),
                          ("EM TELA", ind["ncs_em_tela"]), ("Vencendo em 7 dias", ind["vencendo_7d"]),
                          ("Vencidas", ind["vencidas"])]:
                st.metric(l, v)
        with c2:
            st.subheader("💰 Financeiro")
            for l, v in [("Total Recebido", fmt(ind["recebido_total"])),
                          ("Saldo Disponível", fmt(ind["saldo_total"])),
                          ("Empenhado", fmt(ind["empenhado_total"])),
                          ("Em Tela", fmt(ind["em_tela_total"])),
                          ("REQs Pendentes", fmt(ind["valor_reqs_pendentes"]))]:
                st.metric(l, v)

    elif tipo == "NCs Detalhado":
        if ncs:
            st.dataframe(pd.DataFrame(ncs), use_container_width=True, hide_index=True)

    elif tipo == "Requisições Detalhado":
        if reqs:
            st.dataframe(pd.DataFrame(reqs), use_container_width=True, hide_index=True)

    elif tipo == "NCs por Operação":
        if ncs:
            st.dataframe(ncs_por_operacao(ncs), use_container_width=True, hide_index=True)

    elif tipo == "NCs Próximas do Vencimento (30 dias)":
        venc = ncs_vencendo(ncs, 30)
        if venc:
            df = pd.DataFrame(venc)
            cols = [c for c in ["_DIAS_RESTANTES", "NC", "ORGÃO", "FINALIDADE", "PRAZO", "SALDO NC", "SITU"] if c in df.columns]
            st.dataframe(df[cols].rename(columns={"_DIAS_RESTANTES": "Dias Restantes"}),
                         use_container_width=True, hide_index=True)
        else:
            st.success("✅ Nenhuma NC vencendo nos próximos 30 dias.")

    elif tipo == "REQs Pendentes":
        pend = [r for r in reqs if r.get("SITUAÇÃO") == "Pendente"]
        if pend:
            st.dataframe(pd.DataFrame(pend), use_container_width=True, hide_index=True)
            st.warning(f"📝 {len(pend)} pendentes — {fmt(sum(parse(r.get('VALOR', 0)) for r in pend))}")
        else:
            st.success("✅ Nenhuma REQ pendente!")

    elif tipo == "Extrato por NC":
        df = relatorio_extrato_nc(ncs, reqs)
        if not df.empty:
            st.dataframe(df, use_container_width=True, hide_index=True)
        else:
            st.info("Sem dados.")

    elif tipo == "REQs por Empresa":
        df = relatorio_por_empresa(reqs)
        if not df.empty:
            t = _t()
            fig = px.bar(df, x="Empresa", y="Total (R$_num)", color="Qtd REQs",
                         color_continuous_scale=["#1a4a2e", t["bar_emp"]],
                         labels={"Total (R$_num)": "Total R$", "Empresa": ""})
            fig.update_coloraxes(showscale=False)
            fig.update_layout(height=320, **_cl())
            st.plotly_chart(fig, use_container_width=True)
            st.dataframe(df.drop(columns=["Total (R$_num)"], errors="ignore"),
                         use_container_width=True, hide_index=True)
        else:
            st.info("Sem dados.")

    elif tipo == "Saldo por PI":
        df = relatorio_saldo_pi(ncs)
        if not df.empty:
            t = _t()
            fig = px.bar(df, x="Saldo_num", y="PI", orientation="h",
                         color="Saldo_num", color_continuous_scale=["#1a4a2e", t["bar_emp"]],
                         labels={"Saldo_num": "Saldo R$", "PI": ""})
            fig.update_coloraxes(showscale=False)
            fig.update_layout(height=max(280, len(df) * 35), **_cl())
            st.plotly_chart(fig, use_container_width=True)
            st.dataframe(df.drop(columns=["Saldo_num"], errors="ignore"),
                         use_container_width=True, hide_index=True)
        else:
            st.info("Sem dados.")

    elif tipo == "Saldo por ND":
        df = relatorio_saldo_nd(ncs)
        if not df.empty:
            t = _t()
            fig = px.pie(df, values="Saldo_num", names="ND", hole=0.5,
                         color_discrete_sequence=px.colors.sequential.Teal)
            fig.update_layout(height=320, **{**_cl(), "margin": dict(l=0, r=0, t=20, b=30)})
            fig.update_traces(textfont_color="#e2e8f0", textinfo="percent+label")
            st.plotly_chart(fig, use_container_width=True)
            st.dataframe(df.drop(columns=["Saldo_num"], errors="ignore"),
                         use_container_width=True, hide_index=True)
        else:
            st.info("Sem dados.")


# ── Assistente ────────────────────────────────────────────────────────────────
def page_assistente(ncs, reqs):
    st.title("🤖 Assistente SSAC")
    if not ANTHROPIC_API_KEY:
        st.error("⚠️ Configure ANTHROPIC_API_KEY para usar o assistente.")
        return

    if "hist_api" not in st.session_state: st.session_state.hist_api = []
    if "hist_ui"  not in st.session_state: st.session_state.hist_ui  = []

    for msg in st.session_state.hist_ui:
        with st.chat_message(msg["role"]):
            st.markdown(msg["content"])

    if prompt := st.chat_input("Pergunte sobre NCs, saldos, prazos, requisições..."):
        st.session_state.hist_ui.append({"role": "user", "content": prompt})
        with st.chat_message("user"): st.markdown(prompt)
        with st.chat_message("assistant"):
            with st.spinner(""):
                try:
                    from assistente import chat
                    resp = chat(prompt, st.session_state.hist_api, ncs, reqs)
                except Exception as e:
                    resp = f"Erro: {e}"
            st.markdown(resp)
        st.session_state.hist_api.append({"role": "user", "content": prompt})
        st.session_state.hist_api.append({"role": "assistant", "content": resp})
        st.session_state.hist_ui.append({"role": "assistant", "content": resp})
        if len(st.session_state.hist_api) > 20:
            st.session_state.hist_api = st.session_state.hist_api[-20:]

    if st.session_state.hist_ui:
        if st.button("🗑️ Limpar conversa"):
            st.session_state.hist_api = []; st.session_state.hist_ui = []; st.rerun()


# ── Gerar REQ ────────────────────────────────────────────────────────────────
def page_gerar_req(reqs, ncs):
    st.title("📃 Gerar Requisição de Empenho")

    from config import TEMPLATE_PATH
    import os

    if not os.path.exists(TEMPLATE_PATH):
        st.error(f"Template não encontrado: `{TEMPLATE_PATH}`")
        st.info("Coloque `requisicao_empenho_fiel_placeholders.docx` em `templates/`.")
        return

    if "req_itens" not in st.session_state:
        st.session_state.req_itens = []

    # ── 1. Seletores reativos (fora do form)
    st.subheader("1. Dados do Documento")

    nums_req = ["— manual —"] + [
        f"REQ {r.get('REQ','')} · {r.get('EMPRESA','')} · {r.get('VALOR','')}"
        for r in reqs if r.get("EMPRESA")
    ]
    sel = st.selectbox("Pré-preencher a partir de REQ existente (opcional)", nums_req, key="gerar_sel_req")

    req_base = {}
    if sel != "— manual —":
        idx = nums_req.index(sel) - 1
        req_base = reqs[idx] if idx < len(reqs) else {}

    # NC selector — ao mudar, preenche campos da NC automaticamente
    nums_nc = [""] + [nc.get("NC", "") for nc in ncs if nc.get("NC")]
    nc_default = req_base.get("NC", "")
    nc_idx = nums_nc.index(nc_default) if nc_default in nums_nc else 0
    nc_sel = st.selectbox("NC Vinculada *", nums_nc, index=nc_idx, key="gerar_nc_sel")

    # Dados da NC selecionada
    nc_d = next((nc for nc in ncs if nc.get("NC") == nc_sel), {}) if nc_sel else {}
    if nc_d:
        st.info(
            f"📋 **{nc_sel}** · {nc_d.get('ORGÃO','')} · {nc_d.get('FINALIDADE','')[:55]}"
            f" · Saldo: {nc_d.get('SALDO NC','')} · Prazo: {nc_d.get('PRAZO','')}"
        )

    def _v(req_campo, nc_campo=None, default=""):
        v = req_base.get(req_campo, "")
        if not v and nc_campo:
            v = nc_d.get(nc_campo, "")
        return v or default

    # ── Saldo da NC selecionada
    nc_saldo = parse(nc_d.get("SALDO NC", 0)) if nc_d else 0.0
    total_itens = sum(i["_total"] for i in st.session_state.req_itens)
    if nc_d and nc_saldo > 0:
        k1, k2, k3 = st.columns(3)
        k1.metric("💰 Saldo NC", fmt(nc_saldo))
        k2.metric("📋 Total desta REQ", fmt(total_itens))
        saldo_apos = nc_saldo - total_itens
        k3.metric("🔖 Saldo após REQ", fmt(saldo_apos),
                  delta="OK" if saldo_apos >= 0 else "⚠️ Excede saldo",
                  delta_color="normal" if saldo_apos >= 0 else "inverse")

    # ── Frases cadastradas (fora do form para reatividade)
    col_i, col_j = st.columns(2)
    with col_i:
        st.caption("📝 Introdução")
        frases_intro = _frases("INTRO")
        sel_intro = st.selectbox("Frase padrão", frases_intro, key="sel_intro_frase",
                                  label_visibility="collapsed")
        if sel_intro != "— escrever manualmente —":
            st.session_state["_intro_txt"] = sel_intro
        elif "_intro_txt" not in st.session_state:
            st.session_state["_intro_txt"] = ""

        with st.expander("➕ Cadastrar nova frase de Introdução"):
            nova_intro = st.text_area("Texto", key="nova_intro_txt", height=80)
            if st.button("💾 Salvar", key="btn_salvar_intro"):
                if nova_intro.strip():
                    from sheets_nc import adicionar_frase
                    adicionar_frase("INTRO", nova_intro.strip())
                    st.cache_data.clear()
                    st.success("Frase salva!")
                    st.rerun()

        if sel_intro != "— escrever manualmente —":
            with st.expander("🗑️ Excluir esta frase"):
                if st.button("Confirmar exclusão", key="btn_del_intro"):
                    from sheets_nc import excluir_frase
                    excluir_frase("INTRO", sel_intro)
                    st.cache_data.clear()
                    st.rerun()

    with col_j:
        st.caption("📝 Justificativa")
        frases_just = _frases("JUST")
        sel_just = st.selectbox("Frase padrão", frases_just, key="sel_just_frase",
                                 label_visibility="collapsed")
        if sel_just != "— escrever manualmente —":
            st.session_state["_just_txt"] = sel_just
        elif "_just_txt" not in st.session_state:
            st.session_state["_just_txt"] = ""

        with st.expander("➕ Cadastrar nova frase de Justificativa"):
            nova_just = st.text_area("Texto", key="nova_just_txt", height=80)
            if st.button("💾 Salvar", key="btn_salvar_just"):
                if nova_just.strip():
                    from sheets_nc import adicionar_frase
                    adicionar_frase("JUST", nova_just.strip())
                    st.cache_data.clear()
                    st.success("Frase salva!")
                    st.rerun()

        if sel_just != "— escrever manualmente —":
            with st.expander("🗑️ Excluir esta frase"):
                if st.button("Confirmar exclusão", key="btn_del_just"):
                    from sheets_nc import excluir_frase
                    excluir_frase("JUST", sel_just)
                    st.cache_data.clear()
                    st.rerun()

    # ── Bloco 1: Dados manuais (NC, req#, data, assunto, intro, just)
    st.subheader("1. Dados da Requisição")
    with st.form("f_campos_req"):
        c1, c2, c3 = st.columns(3)
        req_id   = c1.text_input("Nº Requisição",        value=_v("REQ"))
        data_req = c2.date_input("Data",                  value=date.today())
        ne       = c3.text_input("NE (Nota de Empenho)", value=_v("NE"))

        c4, c5, c6 = st.columns(3)
        pi    = c4.text_input("PI",    value=_v("PI", "PI"))
        nd    = c5.text_input("ND",    value=nc_d.get("ND", ""))
        ptres = c6.text_input("PTRES", value=nc_d.get("PTRES", ""))

        c7, c8, c9 = st.columns(3)
        tipo = c7.selectbox("Tipo", ["Ordinário", "Especial", "Suprimento de Fundos"],
                             index=0 if _v("TIPO") != "Especial" else 1)
        ug   = c8.text_input("UG", value=nc_d.get("UG", UG_PADRAO))
        om   = c9.text_input("OM", value=nc_d.get("OM", OM_PADRAO))

        assunto      = st.text_input("Assunto", value=_v("FINALIDADE", "FINALIDADE"))
        intro        = st.text_area("Introdução",  value=st.session_state.get("_intro_txt", ""), height=80)
        justificativa = st.text_area("Justificativa", value=st.session_state.get("_just_txt", ""), height=80)
        finalidade   = st.text_area("Finalidade / Objeto", value=_v("FINALIDADE", "FINALIDADE"), height=60)

        salvar_basico = st.form_submit_button("✅ Confirmar Dados Básicos", type="primary", use_container_width=True)

    if salvar_basico:
        _MESES = {1:"janeiro",2:"fevereiro",3:"março",4:"abril",5:"maio",6:"junho",
                  7:"julho",8:"agosto",9:"setembro",10:"outubro",11:"novembro",12:"dezembro"}
        local_data = f"Boa Vista/RR, {data_req.day} de {_MESES[data_req.month]} de {data_req.year}"
        dados_nc   = f"{nc_sel} de {nc_d.get('DATA NC','')}" if nc_sel else ""
        st.session_state.campos_req = {
            "requisition_id": req_id,
            "LOCAL_DATA":     local_data,
            "DADOS_NC":       dados_nc,
            "NE":             ne,
            "PI":             pi,
            "ND":             nd,
            "PTRES":          ptres,
            "TIPO":           tipo,
            "UG":             ug,
            "OM":             om,
            "ASSUNTO":        assunto,
            "INTRO_1":        intro,
            "JUSTIFICATIVA":  justificativa,
            "FINALIDADE":     finalidade,
        }
        st.success("✅ Dados básicos confirmados.")

    # ── Pesquisa no portal
    st.divider()
    st.subheader("🔍 Pesquisar Item no Portal Compras.gov.br")

    p1, p2 = st.columns([1, 2])
    p_uasg   = p1.text_input("UASG", value=UG_PADRAO, key="pesq_uasg",
                              help="Código da unidade gerenciadora (ex: 160482)")
    p_pregao = p2.text_input("Nº do Pregão", placeholder="ex: 90008/2025",
                              key="pesq_pregao",
                              help="Formato: SEQUENCIAL/ANO — ex: 90008/2025")

    if st.button("🔍 Pesquisar no Portal Compras", key="btn_pesquisar", use_container_width=True):
        if not p_pregao:
            st.warning("Digite o número do pregão.")
        else:
            with st.spinner("🌐 Buscando em comprasnet.gov.br..."):
                from pesquisa_compras import buscar_itens_pregao
                resultado = buscar_itens_pregao(p_uasg, p_pregao)
                st.session_state["pesq_resultados"] = resultado.itens
                st.session_state["pesq_url"]        = resultado.url_usada
                st.session_state["pesq_log"]        = resultado.log

            # Exibe log de diagnóstico sempre
            with st.expander("🔍 Diagnóstico da busca", expanded=not resultado.itens):
                for linha in resultado.log:
                    st.text(linha)

            if not resultado.itens:
                st.warning("Nenhum item retornado. Veja o diagnóstico acima.")

        # Exibe resultados
        resultados_pesq = st.session_state.get("pesq_resultados", [])
        if resultados_pesq:
            url_exib = st.session_state.get("pesq_url", "")
            st.success(f"✅ {len(resultados_pesq)} item(ns) encontrado(s).")
            if url_exib:
                st.caption(f"🔗 Fonte: `{url_exib}`")

            # Agrupa por fornecedor
            from collections import OrderedDict
            grupos: dict = OrderedDict()
            for res in resultados_pesq:
                forn = res.get("fornecedor", "") or "Sem fornecedor"
                grupos.setdefault(forn, []).append(res)

            for gi, (forn, itens_forn) in enumerate(grupos.items()):
                cnpj_forn = itens_forn[0].get("cnpj", "")
                vig_fim   = itens_forn[0].get("vigencia_fim", "")

                st.markdown(f"---\n**🏢 {forn}** · {cnpj_forn} · Vigência: {vig_fim}")

                for ii, res in enumerate(itens_forn):
                    ca, cb, cc, cd, ce = st.columns([1, 5, 2, 2, 2])
                    ca.markdown(f"**{res.get('numero_item','')}**")
                    cb.markdown(res.get("descricao", "")[:80])
                    cc.markdown(res.get("und", ""))
                    cd.markdown(res.get("valor_unit", ""))
                    if ce.button("➕ Usar", key=f"u{gi}i{ii}"):
                        try:
                            v = float(res.get("valor_unit_num") or 0)
                            st.session_state.req_itens.append({
                                "ORD":            str(len(st.session_state.req_itens) + 1),
                                "ITEM":           str(res.get("numero_item", "")),
                                "SI":             "",
                                "DESCRICAO_ITEM": str(res.get("descricao", "")),
                                "UND":            str(res.get("und", "UN") or "UN"),
                                "QTD":            "1,000",
                                "VALOR_UNIT":     fmt(v),
                                "VALOR_TOTAL":    fmt(v),
                                "_total":         v,
                            })
                            st.session_state["_b2_forn"]   = str(forn)
                            st.session_state["_b2_cnpj"]   = str(cnpj_forn)
                            st.session_state["_b2_pregao"] = str(st.session_state.get("pesq_pregao", ""))
                            st.session_state["_b2_ug"]     = str(st.session_state.get("pesq_uasg", UG_PADRAO))
                            st.session_state["_b2_vig"]    = str(vig_fim)
                        except Exception as ex:
                            st.error(f"Erro ao usar item: {ex}")

                if st.button(f"✅ Usar TODOS ({len(itens_forn)} itens)", key=f"ut{gi}"):
                    try:
                        for res in itens_forn:
                            v = float(res.get("valor_unit_num") or 0)
                            st.session_state.req_itens.append({
                                "ORD":            str(len(st.session_state.req_itens) + 1),
                                "ITEM":           str(res.get("numero_item", "")),
                                "SI":             "",
                                "DESCRICAO_ITEM": str(res.get("descricao", "")),
                                "UND":            str(res.get("und", "UN") or "UN"),
                                "QTD":            "1,000",
                                "VALOR_UNIT":     fmt(v),
                                "VALOR_TOTAL":    fmt(v),
                                "_total":         v,
                            })
                        st.session_state["_b2_forn"]   = str(forn)
                        st.session_state["_b2_cnpj"]   = str(cnpj_forn)
                        st.session_state["_b2_pregao"] = str(st.session_state.get("pesq_pregao", ""))
                        st.session_state["_b2_ug"]     = str(st.session_state.get("pesq_uasg", UG_PADRAO))
                        st.session_state["_b2_vig"]    = str(vig_fim)
                    except Exception as ex:
                        st.error(f"Erro: {ex}")

    # ── Bloco 2: Dados do Fornecedor / Pregão (preenchido pela pesquisa)
    st.divider()
    st.subheader("2. Fornecedor / Pregão")

    _b2_forn   = st.session_state.get("_b2_forn", "")
    _b2_cnpj   = st.session_state.get("_b2_cnpj", "")
    _b2_pregao = st.session_state.get("_b2_pregao", "")
    _b2_ug     = st.session_state.get("_b2_ug", UG_PADRAO)
    _b2_vig    = st.session_state.get("_b2_vig", "")

    if _b2_forn:
        st.success(f"🏢 **{_b2_forn}** | {_b2_cnpj} | Pregão: {_b2_pregao} | Vigência: {_b2_vig}")
    else:
        st.caption("⬆️ Use ➕ Usar na pesquisa acima para preencher automaticamente.")

    b2_c1, b2_c2 = st.columns(2)
    b2_forn_edit = b2_c1.text_input("Fornecedor (Razão Social)", value=_b2_forn, key="b2_forn_inp")
    b2_cnpj_edit = b2_c2.text_input("CNPJ", value=_b2_cnpj, key="b2_cnpj_inp")

    b2_m1, b2_m2, b2_m3, b2_m4 = st.columns([2, 2, 1, 2])
    b2_modal_edit = b2_m1.selectbox("Modalidade",
                    ["PREGÃO", "CARONA", "DISPENSA", "INEXIGIBILIDADE", "SUPRIMENTO DE FUNDOS"],
                    key="b2_modal_inp")
    b2_pregao_edit = b2_m2.text_input("Nº Pregão/ARP",   value=_b2_pregao, key="b2_pregao_inp")
    b2_ug_edit     = b2_m3.text_input("UG",              value=_b2_ug,     key="b2_ug_inp")
    b2_vig_edit    = b2_m4.text_input("Vigência da ATA", value=_b2_vig,    key="b2_vig_inp")

    # ── 3. Itens
    st.divider()
    st.subheader("3. Itens da Requisição")

    # Inicializa chaves do formulário de item se não existirem
    for k, v in [("_fi_item",""), ("_fi_desc",""), ("_fi_und","UN"), ("_fi_vunit", 0.0)]:
        st.session_state.setdefault(k, v)

    with st.form("f_add_item", clear_on_submit=True):
        ci1, ci2, ci3 = st.columns([1, 1, 5])
        item  = ci1.text_input("Item *",   key="_fi_item",  help="Nº do item no pregão")
        si    = ci2.text_input("Sub-item", placeholder="ex: 1", help="Nº do sub-item (se houver)")
        desc  = ci3.text_input("Descrição do Item *", key="_fi_desc")
        ci4, ci5, ci6 = st.columns([1, 1, 2])
        und   = ci4.text_input("Unid.", key="_fi_und")
        qtd   = ci5.number_input("Qtd", min_value=0.001, value=1.0, step=1.0, format="%.3f")
        vunit = ci6.number_input("Valor Unit. (R$)", key="_fi_vunit",
                                 min_value=0.0, step=0.01, format="%.2f")
        add   = st.form_submit_button("➕ Adicionar Item", use_container_width=True)

    if add:
        if item and desc and vunit > 0:
            total   = round(qtd * vunit, 2)
            ord_num = len(st.session_state.req_itens) + 1
            st.session_state.req_itens.append({
                "ORD":            str(ord_num),
                "ITEM":           item,
                "SI":             si,
                "DESCRICAO_ITEM": desc,
                "UND":            und,
                "QTD":            str(qtd).replace(".", ","),
                "VALOR_UNIT":     fmt(vunit),
                "VALOR_TOTAL":    fmt(total),
                "_total":         total,
            })
            st.rerun()
        else:
            st.warning("Preencha o Nº do item, a descrição e o valor unitário.")

    if st.session_state.req_itens:
        df_it = pd.DataFrame(st.session_state.req_itens)
        st.dataframe(
            df_it[["ORD", "ITEM", "SI", "DESCRICAO_ITEM", "UND", "QTD", "VALOR_UNIT", "VALOR_TOTAL"]],
            use_container_width=True, hide_index=True,
        )
        total_geral = sum(i["_total"] for i in st.session_state.req_itens)
        k1, k2 = st.columns([1, 3])
        k1.metric("💰 Total Geral", fmt(total_geral))
        with k2:
            if st.button("🗑️ Limpar todos os itens"):
                st.session_state.req_itens = []
                st.rerun()
    else:
        st.info("Nenhum item adicionado ainda.")

    # ── 4. Gerar
    st.divider()
    st.subheader("4. Gerar Documento")

    if st.button("📄 Gerar DOCX", type="primary"):
        campos = st.session_state.get("campos_req", {})
        itens  = st.session_state.req_itens

        if not campos:
            st.error("Confirme os dados básicos primeiro (Bloco 1).")
        elif not itens:
            st.error("Adicione pelo menos um item (Bloco 3).")
        else:
            try:
                from gerador import gerar_para_bytes

                modalidade = f"{st.session_state.get('b2_modal_inp','PREGÃO')} - {st.session_state.get('b2_pregao_inp','')} {st.session_state.get('b2_ug_inp','')}".strip(" -")
                total_geral = sum(i["_total"] for i in itens)
                campos_finais = {
                    **campos,
                    "FORNECEDOR_NOME": st.session_state.get("b2_forn_inp", ""),
                    "FORNECEDOR_CNPJ": _formatar_cnpj(st.session_state.get("b2_cnpj_inp", "")),
                    "MODALIDADE":      modalidade,
                    "VIGENCIA_DA_ATA": st.session_state.get("b2_vig_inp", ""),
                    "TOTAL":           fmt(total_geral),
                }
                itens_limpos = [{k: v for k, v in i.items() if not k.startswith("_")} for i in itens]
                doc_bytes    = gerar_para_bytes(TEMPLATE_PATH, campos_finais, itens_limpos)
                nome_arq     = f"REQ_{campos.get('requisition_id', 'doc')}.docx"
                st.download_button(
                    "⬇️ Baixar Documento",
                    data=doc_bytes,
                    file_name=nome_arq,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True,
                )
                st.success(f"✅ {nome_arq} pronto para download!")
            except Exception as e:
                st.error(f"Erro ao gerar documento: {e}")


def _formatar_cnpj(cnpj: str) -> str:
    import re
    digits = re.sub(r"\D", "", cnpj)
    if len(digits) == 14:
        return f"{digits[:2]}.{digits[2:5]}.{digits[5:8]}/{digits[8:12]}-{digits[12:]}"
    return cnpj


# ── Main ──────────────────────────────────────────────────────────────────────
def main():
    requer_auth()
    _inject_css()

    for k in ("form_nc", "form_req", "tema_claro"):
        if k not in st.session_state:
            st.session_state[k] = False
    if "page" not in st.session_state:
        st.session_state["page"] = "dashboard"

    ncs, reqs, _ = carregar()
    pagina = _sidebar(ncs, reqs)

    if   pagina == "dashboard":  page_dashboard(ncs, reqs)
    elif pagina == "ncs":        page_ncs(ncs)
    elif pagina == "reqs":       page_reqs(reqs, ncs)
    elif pagina == "pdf":        page_pdf(ncs)
    elif pagina == "gerar_req":  page_gerar_req(reqs, ncs)
    elif pagina == "importar":   page_importar(ncs, reqs)
    elif pagina == "relatorios": page_relatorios(ncs, reqs)
    elif pagina == "assistente": page_assistente(ncs, reqs)


if __name__ == "__main__":
    main()
