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
from relatorios import exportar_excel, kpis, ncs_por_operacao, ncs_vencendo

# ── CSS ───────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
/* ── Base ─────────────────────────────────── */
#MainMenu, footer, header {visibility:hidden;}
.block-container {padding:1.5rem 2rem 3rem !important; max-width:100% !important;}
* {box-sizing:border-box;}

/* ── Sidebar ──────────────────────────────── */
section[data-testid="stSidebar"] {width:230px !important; min-width:230px !important;}
section[data-testid="stSidebar"] > div:first-child {
    background:#080e1a !important;
    border-right:1px solid #1a2540 !important;
    padding:0 !important;
}
section[data-testid="stSidebar"] .block-container {padding:1rem 0.75rem !important;}

/* Nav buttons */
section[data-testid="stSidebar"] [data-testid="baseButton-secondary"] {
    text-align:left !important; justify-content:flex-start !important;
    background:transparent !important; border:none !important;
    color:#4b6080 !important; font-size:.875rem !important; font-weight:500 !important;
    padding:9px 12px !important; border-radius:8px !important;
    margin:1px 0 !important; transition:all .15s !important;
}
section[data-testid="stSidebar"] [data-testid="baseButton-secondary"]:hover {
    background:#0f1f35 !important; color:#c8d8e8 !important;
}
section[data-testid="stSidebar"] [data-testid="baseButton-primary"] {
    text-align:left !important; justify-content:flex-start !important;
    background:linear-gradient(135deg,#052918,#073d22) !important;
    border:1px solid #0a4a2840 !important; color:#34d399 !important;
    font-size:.875rem !important; font-weight:600 !important;
    padding:9px 12px !important; border-radius:8px !important;
    margin:1px 0 !important; box-shadow:0 2px 8px rgba(16,185,129,.1) !important;
}
section[data-testid="stSidebar"] [data-testid="baseButton-primary"]:hover {
    background:linear-gradient(135deg,#063520,#094d2a) !important;
}
/* Action buttons in sidebar */
section[data-testid="stSidebar"] .action-btn [data-testid="baseButton-secondary"] {
    background:#0f1f35 !important; color:#64748b !important;
    font-size:.78rem !important; padding:7px 10px !important;
    border:1px solid #1a2540 !important;
}
section[data-testid="stSidebar"] hr {border-color:#1a2540 !important; margin:10px 0 !important;}

/* ── KPI Grid ─────────────────────────────── */
.kpi-grid {display:grid; grid-template-columns:repeat(auto-fit,minmax(220px,1fr)); gap:14px; margin-bottom:22px;}
.kpi-card {
    background:linear-gradient(135deg,#0d1829 0%,#111f35 100%);
    border:1px solid #1a2540; border-radius:14px; padding:18px 20px;
    position:relative; overflow:hidden; transition:all .2s; cursor:default;
}
.kpi-card:hover {transform:translateY(-3px); box-shadow:0 10px 30px rgba(0,0,0,.4);}
.kpi-card-top {display:flex; justify-content:space-between; align-items:center; margin-bottom:10px;}
.kpi-icon {width:40px;height:40px;border-radius:10px;display:flex;align-items:center;justify-content:center;font-size:1.1rem;}
.kpi-badge {font-size:.65rem;font-weight:700;padding:3px 8px;border-radius:10px;letter-spacing:.3px;}
.kpi-value {font-size:1.65rem;font-weight:800;line-height:1.1;margin-bottom:3px;}
.kpi-label {font-size:.68rem;color:#4b6080;text-transform:uppercase;letter-spacing:.9px;font-weight:600;}
.kpi-sub {font-size:.72rem;color:#4b6080;margin-top:5px;}

/* ── Alerts ───────────────────────────────── */
.al {padding:10px 15px;border-radius:9px;margin:5px 0;font-size:.85rem;display:flex;align-items:center;gap:8px;font-weight:500;}
.al-red    {background:#12060a;border:1px solid #7f1d1d50;color:#fca5a5;}
.al-yellow {background:#130e05;border:1px solid #78350f50;color:#fcd34d;}
.al-green  {background:#060e08;border:1px solid #14532d50;color:#6ee7b7;}
.al-blue   {background:#060c14;border:1px solid #1e3a5f50;color:#93c5fd;}
.al-purple {background:#0d080f;border:1px solid #4c1d9550;color:#c4b5fd;}

/* ── Section header ───────────────────────── */
.sec {display:flex;align-items:center;gap:8px;margin-bottom:14px;padding-bottom:8px;border-bottom:1px solid #1a2540;}
.sec-txt {font-size:.67rem;font-weight:700;color:#4b6080;text-transform:uppercase;letter-spacing:1.1px;}

/* ── Stat chip ────────────────────────────── */
.chip {display:inline-flex;align-items:center;gap:4px;padding:4px 10px;border-radius:20px;font-size:.72rem;font-weight:600;}
.chip-green  {background:#052918;color:#34d399;border:1px solid #0a4a2840;}
.chip-amber  {background:#130e05;color:#fbbf24;border:1px solid #78350f40;}
.chip-blue   {background:#060c14;color:#60a5fa;border:1px solid #1e3a5f40;}
.chip-purple {background:#0d080f;color:#a78bfa;border:1px solid #4c1d9540;}
.chip-red    {background:#12060a;color:#f87171;border:1px solid #7f1d1d40;}

/* ── Import cards ─────────────────────────── */
.import-grid {display:grid;grid-template-columns:repeat(auto-fit,minmax(160px,1fr));gap:14px;margin:16px 0;}
.import-card {background:#0d1829;border:1px solid #1a2540;border-radius:12px;padding:20px;text-align:center;}
.import-num {font-size:2.2rem;font-weight:800;color:#34d399;}
.import-lbl {font-size:.72rem;color:#4b6080;text-transform:uppercase;letter-spacing:.8px;margin-top:4px;}

/* ── Tables ───────────────────────────────── */
[data-testid="stDataFrame"] {border:1px solid #1a2540 !important;border-radius:10px !important;overflow:hidden;}
[data-testid="stDataFrame"] table {font-size:.81rem !important;}

/* ── Forms ────────────────────────────────── */
[data-testid="stForm"] {background:#0d1829 !important;border:1px solid #1a2540 !important;border-radius:12px;padding:20px !important;}

/* ── Expander ─────────────────────────────── */
[data-testid="stExpander"] {border:1px solid #1a2540 !important;border-radius:10px !important;background:#0d1829 !important;}

hr {border-color:#1a2540 !important;}
</style>
""", unsafe_allow_html=True)

_CL = dict(
    paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", font_color="#64748b",
    margin=dict(l=0, r=0, t=20, b=0),
    xaxis=dict(gridcolor="#1a2540", zerolinecolor="#1a2540"),
    yaxis=dict(gridcolor="#1a2540", zerolinecolor="#1a2540"),
    legend=dict(orientation="h", y=1.12, font=dict(size=11)),
)


# ── Helpers ───────────────────────────────────────────────────────────────────
def fmt(v: float) -> str:
    return "R$ " + f"{v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def parse(s) -> float:
    try:
        return float(str(s).replace("R$", "").replace(" ", "").replace(".", "").replace(",", "."))
    except Exception:
        return 0.0


def _sec(icon, text):
    st.markdown(f'<div class="sec"><span style="font-size:1rem">{icon}</span><span class="sec-txt">{text}</span></div>',
                unsafe_allow_html=True)


def _alert(cls, text):
    st.markdown(f'<div class="al {cls}">{text}</div>', unsafe_allow_html=True)


def _chip(text, cls="chip-green"):
    return f'<span class="chip {cls}">{text}</span>'


def _kpi(icon, label, value, sub, color, icon_bg, badge=None):
    badge_html = f'<span class="kpi-badge" style="background:{icon_bg};color:{color}">{badge}</span>' if badge else ""
    return f"""
    <div class="kpi-card" style="border-top:2px solid {color}20;border-left:2px solid {color}15">
        <div class="kpi-card-top">
            <div class="kpi-icon" style="background:{icon_bg}">{icon}</div>
            {badge_html}
        </div>
        <div class="kpi-value" style="color:{color}">{value}</div>
        <div class="kpi-label">{label}</div>
        <div class="kpi-sub">{sub}</div>
    </div>"""


def _cor_prazo(prazo_str: str) -> str:
    try:
        n = (datetime.strptime(prazo_str, "%d/%m/%Y").date() - date.today()).days
        return "🔴" if n < 0 else "🟠" if n <= 7 else "🟡" if n <= 30 else "🟢"
    except Exception:
        return "⚪"


ORGAOS = ["COTER", "COEX", "DGO", "DEC", "DECEX", "12 RM", "Outro"]
SITUACOES_REQ = ["Pendente", "Enviada", "Aprovada", "Empenhada", "Liquidada", "Paga"]

NAV = [
    ("📊", "Dashboard",        "dashboard"),
    ("📋", "Notas de Crédito", "ncs"),
    ("📝", "Requisições",      "reqs"),
    ("📄", "Lançar Documento", "pdf"),
    ("📥", "Importar Dados",   "importar"),
    ("📈", "Relatórios",       "relatorios"),
    ("🤖", "Assistente",       "assistente"),
]


# ── Cache ─────────────────────────────────────────────────────────────────────
@st.cache_data(ttl=120, show_spinner=False)
def _load():
    from sheets_nc import ler_fornecedores, ler_ncs, ler_reqs
    return ler_ncs(), ler_reqs(), ler_fornecedores()


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
    with st.sidebar:
        st.markdown(f"""
        <div style="padding:16px 8px 14px;border-bottom:1px solid #1a2540;margin-bottom:8px;">
            <div style="font-size:1.3rem;font-weight:900;color:#e2e8f0;letter-spacing:1px;">🪖 SSAC</div>
            <div style="font-size:.72rem;color:#4b6080;margin-top:3px;line-height:1.4">
                {OM_PADRAO}<br>UG {UG_PADRAO}
            </div>
        </div>""", unsafe_allow_html=True)

        # Navegação
        current = st.session_state.get("page", "dashboard")
        for icon, label, page_id in NAV:
            is_active = current == page_id
            if st.button(f"{icon}  {label}", key=f"nav_{page_id}",
                         use_container_width=True,
                         type="primary" if is_active else "secondary"):
                st.session_state["page"] = page_id
                st.rerun()

        st.divider()

        # Mini stats
        if ncs or reqs:
            ind = kpis(ncs, reqs)
            st.markdown(f"""
            <div style="padding:4px 4px 8px;">
                <div style="font-size:.65rem;color:#4b6080;text-transform:uppercase;letter-spacing:.8px;margin-bottom:8px;">Resumo</div>
                <div style="display:flex;flex-direction:column;gap:5px;font-size:.78rem;">
                    <div style="display:flex;justify-content:space-between;color:#8ba0b8">
                        <span>Saldo</span>
                        <span style="color:#34d399;font-weight:700">{fmt(ind['saldo_total'])}</span>
                    </div>
                    <div style="display:flex;justify-content:space-between;color:#8ba0b8">
                        <span>NCs EM TELA</span>
                        <span style="color:#fbbf24;font-weight:700">{ind['ncs_em_tela']}</span>
                    </div>
                    <div style="display:flex;justify-content:space-between;color:#8ba0b8">
                        <span>REQs pendentes</span>
                        <span style="color:#a78bfa;font-weight:700">{ind['reqs_pendentes']}</span>
                    </div>
                </div>
            </div>
            """, unsafe_allow_html=True)
            st.divider()

        # Ações
        c1, c2 = st.columns(2)
        with c1:
            if st.button("🔄", use_container_width=True, help="Recarregar dados"):
                carregar(forcar=True); st.rerun()
        with c2:
            if st.button("🚪", use_container_width=True, help="Sair"):
                logout()

        st.markdown(f'<div style="font-size:.65rem;color:#2a3a50;text-align:center;padding-top:6px">Sync {datetime.now().strftime("%H:%M:%S")}</div>',
                    unsafe_allow_html=True)

    return st.session_state.get("page", "dashboard")


# ── Dashboard ─────────────────────────────────────────────────────────────────
def page_dashboard(ncs, reqs):
    st.markdown("## 📊 Dashboard")
    st.markdown(f'<div style="color:#4b6080;font-size:.85rem;margin:-8px 0 20px">{OM_PADRAO} · {datetime.today().strftime("%d de %B de %Y")}</div>',
                unsafe_allow_html=True)

    ind = kpis(ncs, reqs)
    total_rec = ind["recebido_total"]
    pct_emp = (ind["empenhado_total"] / total_rec * 100) if total_rec else 0

    # KPI
    cards = (
        _kpi("💰", "SALDO DISPONÍVEL",  fmt(ind["saldo_total"]),      f"{ind['ncs_ok']} NCs com saldo",       "#34d399", "#05291820", f"{ind['total_ncs']} NCs") +
        _kpi("📌", "EMPENHADO",          fmt(ind["empenhado_total"]),  f"{pct_emp:.1f}% do recebido",           "#60a5fa", "#06111e20") +
        _kpi("🕐", "EM TELA",            fmt(ind["em_tela_total"]),    f"{ind['ncs_em_tela']} NCs aguardando",  "#fbbf24", "#13100520") +
        _kpi("📝", "REQS PENDENTES",     str(ind["reqs_pendentes"]),   fmt(ind["valor_reqs_pendentes"]),        "#a78bfa", "#0d08100f")
    )
    st.markdown(f'<div class="kpi-grid">{cards}</div>', unsafe_allow_html=True)

    # Alertas
    if ind["vencidas"] > 0:
        _alert("al-red",    f'⛔ <b>{ind["vencidas"]}</b> NC(s) com prazo vencido!')
    if ind["vencendo_7d"] > 0:
        _alert("al-yellow", f'⚠️ <b>{ind["vencendo_7d"]}</b> NC(s) vencem nos próximos 7 dias.')
    if ind["ncs_em_tela"] > 0:
        _alert("al-blue",   f'🕐 <b>{ind["ncs_em_tela"]}</b> NC(s) EM TELA — {fmt(ind["em_tela_total"])} aguardando crédito.')
    if ind["reqs_pendentes"] > 0:
        _alert("al-purple", f'📝 <b>{ind["reqs_pendentes"]}</b> requisições pendentes — {fmt(ind["valor_reqs_pendentes"])}.')

    st.divider()

    # ── Linha 1: Valores por mês + Status NCs
    c1, c2 = st.columns([3, 2])
    with c1:
        _sec("📅", "Valores Recebidos por Mês")
        rows = []
        for nc in ncs:
            try:
                d = datetime.strptime(nc.get("DATA NC",""), "%d/%m/%Y")
                rows.append({"Mês": d.strftime("%b/%y"), "Ord": d.strftime("%Y-%m"),
                             "Recebido": parse(nc.get("RECEBIDO",0)),
                             "Empenhado": parse(nc.get("EMPENHADO",0))})
            except Exception:
                pass
        if rows:
            df_m = (pd.DataFrame(rows).groupby(["Mês","Ord"])
                    .sum().reset_index().sort_values("Ord"))
            fig = go.Figure()
            fig.add_bar(x=df_m["Mês"], y=df_m["Recebido"], name="Recebido",
                        marker_color="#3b82f6", marker_line_width=0)
            fig.add_bar(x=df_m["Mês"], y=df_m["Empenhado"], name="Empenhado",
                        marker_color="#10b981", marker_line_width=0)
            fig.update_layout(height=300, barmode="group", **_CL)
            st.plotly_chart(fig, use_container_width=True)
        else:
            _alert("al-blue", "ℹ️ Sem dados de data para gráfico mensal.")

    with c2:
        _sec("🔵", "Status das NCs")
        if ncs:
            cnt = pd.DataFrame(ncs)["SITU"].value_counts().reset_index()
            cnt.columns = ["Status", "Qtd"]
            fig = px.pie(cnt, values="Qtd", names="Status", hole=0.6,
                         color_discrete_map={"OK":"#10b981","EM TELA":"#fbbf24"})
            fig.update_layout(height=300, **{**_CL, "margin": dict(l=0,r=0,t=20,b=30)})
            fig.update_traces(textfont_color="#e2e8f0", textinfo="percent+label")
            st.plotly_chart(fig, use_container_width=True)

    # ── Linha 2: Saldo por operação + Empenho por órgão
    c3, c4 = st.columns(2)
    with c3:
        _sec("🎯", "Saldo Disponível por Operação")
        if ncs:
            df_op = (pd.DataFrame([{"OP": nc.get("OP") or "Sem OP",
                                     "Saldo": parse(nc.get("SALDO NC",0))} for nc in ncs])
                     .groupby("OP")["Saldo"].sum().reset_index()
                     .query("Saldo > 0").sort_values("Saldo"))
            if not df_op.empty:
                fig = px.bar(df_op, x="Saldo", y="OP", orientation="h",
                             color="Saldo", color_continuous_scale=["#1a4a2e","#10b981"],
                             labels={"Saldo":"R$","OP":""})
                fig.update_coloraxes(showscale=False)
                fig.update_layout(height=320, **_CL)
                st.plotly_chart(fig, use_container_width=True)
            else:
                _alert("al-green", "✅ Sem saldo disponível no momento.")

    with c4:
        _sec("🏛️", "Recebido vs Empenhado por Órgão")
        if ncs:
            df_o = (pd.DataFrame([{"ORGÃO": nc.get("ORGÃO") or "N/A",
                                    "Recebido": parse(nc.get("RECEBIDO",0)),
                                    "Empenhado": parse(nc.get("EMPENHADO",0))} for nc in ncs])
                    .groupby("ORGÃO").sum().reset_index()
                    .melt(id_vars="ORGÃO", var_name="Tipo", value_name="Valor"))
            fig = px.bar(df_o, x="ORGÃO", y="Valor", color="Tipo", barmode="group",
                         color_discrete_map={"Recebido":"#3b82f6","Empenhado":"#10b981"},
                         labels={"Valor":"R$","ORGÃO":""})
            fig.update_layout(height=320, **_CL)
            st.plotly_chart(fig, use_container_width=True)

    # ── Linha 3: REQs por situação + NCs EM TELA
    c5, c6 = st.columns(2)
    with c5:
        _sec("📝", "REQs por Situação")
        if reqs:
            df_s = (pd.DataFrame([{"Situação": r.get("SITUAÇÃO") or "N/A",
                                    "Valor": parse(r.get("VALOR",0))} for r in reqs])
                    .groupby("Situação")["Valor"].sum().reset_index()
                    .sort_values("Valor", ascending=False))
            cores = {"Pendente":"#fbbf24","Enviada":"#a78bfa","Aprovada":"#60a5fa",
                     "Empenhada":"#3b82f6","Liquidada":"#10b981","Paga":"#34d399"}
            fig = px.bar(df_s, x="Situação", y="Valor",
                         color="Situação", color_discrete_map=cores,
                         labels={"Valor":"R$","Situação":""})
            fig.update_layout(height=300, **_CL, showlegend=False)
            st.plotly_chart(fig, use_container_width=True)

    with c6:
        _sec("🕐", "NCs EM TELA")
        em_tela = [nc for nc in ncs if nc.get("SITU") == "EM TELA"]
        if em_tela:
            total_et = sum(parse(nc.get("RECEBIDO",0)) for nc in em_tela)
            _alert("al-yellow", f'⚠️ {len(em_tela)} NCs · {fmt(total_et)}')
            df_et = pd.DataFrame([{
                "NC": nc.get("NC",""), "ÓRGÃO": nc.get("ORGÃO",""),
                "Finalidade": nc.get("FINALIDADE","")[:38],
                "Valor": nc.get("RECEBIDO",""), "Prazo": nc.get("PRAZO",""),
            } for nc in em_tela])
            st.dataframe(df_et, use_container_width=True, hide_index=True, height=250)
        else:
            _alert("al-green", "✅ Nenhuma NC EM TELA no momento.")

    # ── Valores por operação (acumulado)
    _sec("💹", "Valor Total por Operação")
    if ncs:
        df_vop = (pd.DataFrame([{"OP": nc.get("OP") or "Sem OP",
                                  "Recebido": parse(nc.get("RECEBIDO",0)),
                                  "Empenhado": parse(nc.get("EMPENHADO",0)),
                                  "Saldo": parse(nc.get("SALDO NC",0))} for nc in ncs])
                  .groupby("OP").sum().reset_index().sort_values("Recebido", ascending=False))
        fig = go.Figure()
        fig.add_bar(x=df_vop["OP"], y=df_vop["Recebido"],  name="Recebido",  marker_color="#3b82f6", marker_line_width=0)
        fig.add_bar(x=df_vop["OP"], y=df_vop["Empenhado"], name="Empenhado", marker_color="#fbbf24", marker_line_width=0)
        fig.add_bar(x=df_vop["OP"], y=df_vop["Saldo"],     name="Saldo",     marker_color="#10b981", marker_line_width=0)
        fig.update_layout(barmode="group", height=300, **_CL)
        st.plotly_chart(fig, use_container_width=True)

    # ── Controle de prazos
    _sec("⏰", "Controle de Prazos")
    if ncs:
        hoje = date.today()
        rows = []
        for nc in ncs:
            p = nc.get("PRAZO","")
            try:
                dias = (datetime.strptime(p, "%d/%m/%Y").date() - hoje).days
            except Exception:
                dias = None
            rows.append({"": _cor_prazo(p), "NC": nc.get("NC",""), "ÓRGÃO": nc.get("ORGÃO",""),
                         "Finalidade": nc.get("FINALIDADE","")[:45], "Prazo": p,
                         "Dias": dias if dias is not None else "—",
                         "Saldo": nc.get("SALDO NC",""), "Status": nc.get("SITU","")})
        df_p = pd.DataFrame(rows).sort_values("Dias", key=lambda x: pd.to_numeric(x, errors="coerce"))
        st.dataframe(df_p, use_container_width=True, hide_index=True)


# ── NCs ───────────────────────────────────────────────────────────────────────
def page_ncs(ncs):
    st.markdown("## 📋 Notas de Crédito")

    with st.expander("🔍 Filtros", expanded=False):
        f1, f2, f3 = st.columns(3)
        f_situ  = f1.multiselect("Status",    ["OK","EM TELA"],   default=["OK","EM TELA"])
        f_org   = f2.multiselect("Órgão",     sorted({nc.get("ORGÃO","") for nc in ncs if nc.get("ORGÃO")}))
        f_op    = f3.multiselect("Operação",  sorted({nc.get("OP","")    for nc in ncs if nc.get("OP")}))
        f4, f5, f6 = st.columns(3)
        f_nd    = f4.multiselect("ND",        sorted({nc.get("ND","")    for nc in ncs if nc.get("ND")}))
        f_pi    = f5.multiselect("PI",        sorted({nc.get("PI","")    for nc in ncs if nc.get("PI")}))
        f_prazo = f6.selectbox("Prazo",       ["Todos","Vencidas","Vencem em 7 dias","Vencem em 30 dias"])

    hoje = date.today()
    filtradas = []
    for nc in ncs:
        if f_situ  and nc.get("SITU")  not in f_situ:  continue
        if f_org   and nc.get("ORGÃO") not in f_org:   continue
        if f_op    and nc.get("OP")    not in f_op:    continue
        if f_nd    and nc.get("ND")    not in f_nd:    continue
        if f_pi    and nc.get("PI")    not in f_pi:    continue
        if f_prazo != "Todos":
            try:
                dias = (datetime.strptime(nc.get("PRAZO",""), "%d/%m/%Y").date() - hoje).days
                if f_prazo == "Vencidas"          and dias >= 0:         continue
                if f_prazo == "Vencem em 7 dias"  and not 0 <= dias <= 7: continue
                if f_prazo == "Vencem em 30 dias" and not 0 <= dias <= 30: continue
            except Exception:
                continue
        filtradas.append(nc)

    b1, b2 = st.columns([1, 3])
    with b1:
        if st.button("➕ Nova NC", type="primary", use_container_width=True):
            st.session_state["form_nc"] = not st.session_state.get("form_nc", False)
    with b2:
        saldo_f = sum(parse(nc.get("SALDO NC",0)) for nc in filtradas)
        _alert("al-green", f'📋 <b>{len(filtradas)}</b> NCs · Saldo filtrado: <b>{fmt(saldo_f)}</b>')

    if st.session_state.get("form_nc"):
        _form_nc()

    if filtradas:
        cols = [c for c in ["ORD","SITU","ORGÃO","DATA NC","NC","PI","ND","FINALIDADE","OP",
                             "PRAZO","DIAS","RECEBIDO","SALDO NC","EMPENHADO","EMP %","SITUAÇÃO"]
                if c in pd.DataFrame(filtradas).columns]
        st.dataframe(pd.DataFrame(filtradas)[cols], use_container_width=True, hide_index=True)
    else:
        st.info("Nenhuma NC encontrada.")


def _form_nc():
    st.divider()
    _sec("➕", "Nova Nota de Crédito")
    with st.form("f_nc", clear_on_submit=True):
        c1, c2 = st.columns(2)
        nc_num  = c1.text_input("Número NC *",    placeholder="2026NC000000")
        orgao   = c1.selectbox("Órgão *",         ORGAOS)
        data_nc = c1.date_input("Data NC *",       value=date.today())
        pi      = c1.text_input("PI")
        nd      = c1.text_input("ND",             placeholder="339030")
        ptres   = c2.text_input("PTRES",          placeholder="251050")
        om      = c2.text_input("OM",             value=OM_PADRAO)
        prazo   = c2.date_input("Prazo *")
        op      = c2.text_input("Operação (OP)")
        situ    = c2.selectbox("Status",          ["OK","EM TELA"])
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
                    adicionar_nc({"NC":nc_num,"ORGÃO":orgao,"DATA NC":data_nc.strftime("%d/%m/%Y"),
                                  "PI":pi,"ND":nd,"PTRES":ptres,"OM":om,
                                  "PRAZO":prazo.strftime("%d/%m/%Y"),"OP":op,"SITU":situ,
                                  "FINALIDADE":finalidade,"RECEBIDO":valor})
                    st.success(f"✅ NC {nc_num} adicionada!")
                    st.session_state["form_nc"] = False
                    carregar(forcar=True); st.rerun()
                except Exception as e:
                    st.error(f"Erro: {e}")


# ── Requisições ───────────────────────────────────────────────────────────────
def page_reqs(reqs, ncs):
    st.markdown("## 📝 Requisições")

    with st.expander("🔍 Filtros", expanded=False):
        f1, f2, f3 = st.columns(3)
        f_sit  = f1.multiselect("Situação", sorted({r.get("SITUAÇÃO","") for r in reqs if r.get("SITUAÇÃO")}))
        f_emp  = f2.multiselect("Empresa",  sorted({r.get("EMPRESA","")  for r in reqs if r.get("EMPRESA")}))
        f_nc   = f3.multiselect("NC",       sorted({r.get("NC","")       for r in reqs if r.get("NC")}))
        f4, f5, f6 = st.columns(3)
        f_tipo = f4.multiselect("Tipo",     ["Ordinário","Especial"])
        f_pi   = f5.multiselect("PI",       sorted({r.get("PI","") for r in reqs if r.get("PI")}))
        vals   = [parse(r.get("VALOR",0)) for r in reqs if parse(r.get("VALOR",0)) > 0]
        v_max  = max(vals) if vals else 100000.0
        f_val  = f6.slider("Faixa de valor (R$)", 0.0, float(v_max), (0.0, float(v_max)), step=500.0)

    filtradas = [r for r in reqs
                 if (not f_sit  or r.get("SITUAÇÃO") in f_sit)
                 and (not f_emp or r.get("EMPRESA")  in f_emp)
                 and (not f_nc  or r.get("NC")       in f_nc)
                 and (not f_tipo or r.get("TIPO")    in f_tipo)
                 and (not f_pi  or r.get("PI")       in f_pi)
                 and f_val[0] <= parse(r.get("VALOR",0)) <= f_val[1]]

    b1, b2 = st.columns([1, 3])
    with b1:
        if st.button("➕ Nova REQ", type="primary", use_container_width=True):
            st.session_state["form_req"] = not st.session_state.get("form_req", False)
    with b2:
        total = sum(parse(r.get("VALOR",0)) for r in filtradas)
        _alert("al-green", f'📝 <b>{len(filtradas)}</b> REQs · Total: <b>{fmt(total)}</b>')

    if st.session_state.get("form_req"):
        _form_req(ncs)

    if filtradas:
        cols = [c for c in ["REQ","DATA REQ","NC","PI","EMPRESA","DESCRIÇÃO","VALOR","SITUAÇÃO","NE","FINALIDADE","OBS"]
                if c in pd.DataFrame(filtradas).columns]
        st.dataframe(pd.DataFrame(filtradas)[cols], use_container_width=True, hide_index=True)
    else:
        st.info("Nenhuma REQ encontrada.")


def _form_req(ncs):
    st.divider()
    _sec("➕", "Nova Requisição")
    nums_nc = [""] + [nc.get("NC","") for nc in ncs if nc.get("NC")]
    with st.form("f_req", clear_on_submit=True):
        c1, c2 = st.columns(2)
        req_num  = c1.text_input("Número REQ")
        nc_sel   = c1.selectbox("NC Vinculada *", nums_nc)
        data_req = c1.date_input("Data REQ",  value=date.today())
        empresa  = c1.text_input("Empresa *")
        pi       = c2.text_input("PI")
        tipo     = c2.selectbox("Tipo",      ["Ordinário","Especial"])
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
                        pi_f = nc_d.get("PI",""); fin_f = fin_f or nc_d.get("OP","")
                    from sheets_nc import adicionar_req
                    adicionar_req({"REQ":req_num,"NC":nc_sel,"DATA REQ":data_req.strftime("%d/%m/%Y"),
                                   "PI":pi_f,"FINALIDADE":fin_f,"TIPO":tipo,"EMPRESA":empresa,
                                   "DESCRIÇÃO":descricao,"VALOR":valor,"SITUAÇÃO":situacao,"NE":ne,"OBS":obs})
                    st.success("✅ REQ adicionada!")
                    st.session_state["form_req"] = False
                    carregar(forcar=True); st.rerun()
                except Exception as e:
                    st.error(f"Erro: {e}")


# ── Lançar por PDF / HTML ─────────────────────────────────────────────────────
def page_pdf(ncs):
    st.markdown("## 📄 Lançar por PDF ou HTML")
    if not ANTHROPIC_API_KEY:
        _alert("al-red", "⚠️ Configure <b>ANTHROPIC_API_KEY</b> para usar este recurso.")
        return
    tipo = st.radio("Tipo", ["Nota de Crédito (NC)","Requisição (REQ)"], horizontal=True)
    uploaded = st.file_uploader("Selecione o arquivo", type=["pdf","html","htm"],
                                help="PDF ou HTML exportado do SIAFI.")
    if uploaded:
        with st.spinner("🤖 Analisando documento..."):
            try:
                from extrator_pdf import extrair_nc, extrair_req
                b = uploaded.read()
                if "NC" in tipo:
                    dados = extrair_nc(b, uploaded.name)
                    _alert("al-green", "✅ Dados extraídos — revise e confirme.")
                    _confirmar_nc(dados)
                else:
                    dados = extrair_req(b, uploaded.name)
                    _alert("al-green", "✅ Dados extraídos — revise e confirme.")
                    _confirmar_req(dados, ncs)
            except Exception as e:
                st.error(f"Erro na extração: {e}")


def _confirmar_nc(dados):
    st.divider(); _sec("✅", "Confirmar Nota de Crédito")
    with st.form("f_pdf_nc"):
        c1, c2 = st.columns(2)
        nc      = c1.text_input("Número NC",            value=dados.get("NC",""))
        orgao   = c1.text_input("Órgão",                value=dados.get("ORGÃO",""))
        data_nc = c1.text_input("Data NC (DD/MM/YYYY)", value=dados.get("DATA NC",""))
        pi      = c1.text_input("PI",                   value=dados.get("PI",""))
        nd      = c1.text_input("ND",                   value=dados.get("ND",""))
        ptres   = c2.text_input("PTRES",                value=dados.get("PTRES",""))
        om      = c2.text_input("OM",                   value=dados.get("OM", OM_PADRAO))
        prazo   = c2.text_input("Prazo (DD/MM/YYYY)",   value=dados.get("PRAZO",""))
        op      = c2.text_input("Operação",             value=dados.get("OP",""))
        situ    = c2.selectbox("Status", ["OK","EM TELA"], index=1 if dados.get("SITU")=="EM TELA" else 0)
        finalidade = st.text_area("Finalidade", value=dados.get("FINALIDADE",""))
        valor = st.number_input("Valor (R$)", value=float(dados.get("RECEBIDO",0) or 0), min_value=0.0, step=0.01, format="%.2f")
        if st.form_submit_button("💾 Confirmar e Salvar", type="primary", use_container_width=True):
            try:
                from sheets_nc import adicionar_nc
                adicionar_nc({"NC":nc,"ORGÃO":orgao,"DATA NC":data_nc,"PI":pi,"ND":nd,
                              "PTRES":ptres,"OM":om,"PRAZO":prazo,"OP":op,"SITU":situ,
                              "FINALIDADE":finalidade,"RECEBIDO":valor})
                st.success(f"✅ NC {nc} lançada!")
                carregar(forcar=True)
            except Exception as e:
                st.error(f"Erro: {e}")


def _confirmar_req(dados, ncs):
    st.divider(); _sec("✅", "Confirmar Requisição")
    nums_nc = [""] + [nc.get("NC","") for nc in ncs if nc.get("NC")]
    nc_idx = nums_nc.index(dados.get("NC","")) if dados.get("NC","") in nums_nc else 0
    with st.form("f_pdf_req"):
        c1, c2 = st.columns(2)
        req      = c1.text_input("Número REQ",           value=str(dados.get("REQ","")))
        nc_sel   = c1.selectbox("NC Vinculada",          nums_nc, index=nc_idx)
        data_req = c1.text_input("Data REQ (DD/MM/YYYY)",value=dados.get("DATA REQ",""))
        empresa  = c1.text_input("Empresa",              value=dados.get("EMPRESA",""))
        pi       = c2.text_input("PI",                   value=dados.get("PI",""))
        tipo     = c2.selectbox("Tipo", ["Ordinário","Especial"])
        ne       = c2.text_input("NE",                   value=dados.get("NE",""))
        situacao = c2.selectbox("Situação", SITUACOES_REQ)
        finalidade = st.text_area("Finalidade", value=dados.get("FINALIDADE",""))
        descricao  = st.text_area("Descrição",  value=dados.get("DESCRIÇÃO",""))
        valor = st.number_input("Valor (R$)", value=float(dados.get("VALOR",0) or 0), min_value=0.0, step=0.01, format="%.2f")
        if st.form_submit_button("💾 Confirmar e Salvar", type="primary", use_container_width=True):
            try:
                from sheets_nc import adicionar_req
                adicionar_req({"REQ":req,"NC":nc_sel,"DATA REQ":data_req,"PI":pi,
                               "FINALIDADE":finalidade,"TIPO":tipo,"EMPRESA":empresa,
                               "DESCRIÇÃO":descricao,"VALOR":valor,"SITUAÇÃO":situacao,"NE":ne})
                st.success("✅ REQ lançada!")
                carregar(forcar=True)
            except Exception as e:
                st.error(f"Erro: {e}")


# ── Importar Dados ────────────────────────────────────────────────────────────
def page_importar(ncs, reqs):
    st.markdown("## 📥 Importar Dados")

    # Status da conexão
    _sec("🔗", "Conexão com Google Sheets")
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        st.markdown(f'<div class="import-card"><div class="import-num">{len(ncs)}</div><div class="import-lbl">NCs carregadas</div></div>',
                    unsafe_allow_html=True)
    with c2:
        st.markdown(f'<div class="import-card"><div class="import-num">{len(reqs)}</div><div class="import-lbl">REQs carregadas</div></div>',
                    unsafe_allow_html=True)
    with c3:
        ncs_ok = sum(1 for nc in ncs if nc.get("SITU") == "OK")
        st.markdown(f'<div class="import-card"><div class="import-num">{ncs_ok}</div><div class="import-lbl">NCs OK</div></div>',
                    unsafe_allow_html=True)
    with c4:
        reqs_p = sum(1 for r in reqs if r.get("SITUAÇÃO") == "Pendente")
        st.markdown(f'<div class="import-card"><div class="import-num" style="color:#fbbf24">{reqs_p}</div><div class="import-lbl">REQs Pendentes</div></div>',
                    unsafe_allow_html=True)

    _alert("al-blue", f'🔗 Planilha: <b>{SHEET_ID_NC[:20]}...</b> · Dados sincronizados a cada 2 minutos.')

    col_sync, _ = st.columns([1, 3])
    with col_sync:
        if st.button("🔄 Forçar Sincronização", use_container_width=True, type="primary"):
            with st.spinner("Sincronizando..."):
                carregar(forcar=True)
            st.success("✅ Dados atualizados!")
            st.rerun()

    st.divider()

    # Importar de arquivo
    _sec("📂", "Importar de Arquivo Excel ou CSV")
    _alert("al-blue", "ℹ️ Importe NCs ou REQs de uma planilha Excel/CSV local. Os dados serão adicionados à planilha oficial.")

    tipo_imp = st.radio("O que importar?", ["Notas de Crédito (NCs)", "Requisições (REQs)"], horizontal=True)
    arq = st.file_uploader("Selecione o arquivo", type=["xlsx","xls","csv"])

    if arq:
        try:
            if arq.name.endswith(".csv"):
                df_imp = pd.read_csv(arq, dtype=str).fillna("")
            else:
                df_imp = pd.read_excel(arq, dtype=str).fillna("")

            st.markdown(f"**{len(df_imp)} linhas encontradas** — prévia:")
            st.dataframe(df_imp.head(10), use_container_width=True, hide_index=True)

            if st.button(f"⬆️ Importar {len(df_imp)} linha(s) para a planilha", type="primary"):
                from sheets_nc import adicionar_nc, adicionar_req
                erros = 0
                progress = st.progress(0)
                for i, row in df_imp.iterrows():
                    try:
                        d = row.to_dict()
                        if "NCs" in tipo_imp:
                            adicionar_nc(d)
                        else:
                            adicionar_req(d)
                    except Exception:
                        erros += 1
                    progress.progress((i + 1) / len(df_imp))

                carregar(forcar=True)
                if erros == 0:
                    st.success(f"✅ {len(df_imp)} linhas importadas com sucesso!")
                else:
                    st.warning(f"⚠️ {len(df_imp) - erros} importadas, {erros} com erro.")
        except Exception as e:
            st.error(f"Erro ao ler arquivo: {e}")

    st.divider()

    # Preview dos dados atuais
    _sec("👁️", "Preview dos Dados Atuais")
    aba = st.radio("Ver", ["NCs", "REQs"], horizontal=True)
    if aba == "NCs" and ncs:
        st.dataframe(pd.DataFrame(ncs), use_container_width=True, hide_index=True)
    elif aba == "REQs" and reqs:
        st.dataframe(pd.DataFrame(reqs), use_container_width=True, hide_index=True)
    else:
        st.info("Sem dados.")


# ── Relatórios ────────────────────────────────────────────────────────────────
def page_relatorios(ncs, reqs):
    st.markdown("## 📈 Relatórios")

    tipo = st.selectbox("Tipo", [
        "Resumo Geral","NCs Detalhado","Requisições Detalhado",
        "NCs por Operação","NCs Próximas do Vencimento (30 dias)","REQs Pendentes",
    ])
    _exibir_relatorio(tipo, ncs, reqs)

    st.divider()
    _sec("📥", "Exportar")
    c1, c2 = st.columns([1, 3])
    with c1:
        st.download_button("⬇️ Excel Completo",
                           data=exportar_excel(ncs, reqs),
                           file_name=f"ssac_{datetime.today().strftime('%Y%m%d_%H%M')}.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           use_container_width=True)
    with c2:
        st.caption("NCs · Requisições · Resumo · Por Operação — em abas separadas.")


def _exibir_relatorio(tipo, ncs, reqs):
    if tipo == "Resumo Geral":
        ind = kpis(ncs, reqs)
        c1, c2 = st.columns(2)
        with c1:
            _sec("📋", "NCs")
            for l, v in [("Total",ind["total_ncs"]),("OK",ind["ncs_ok"]),("EM TELA",ind["ncs_em_tela"]),
                          ("Vencendo em 7 dias",ind["vencendo_7d"]),("Vencidas",ind["vencidas"])]:
                st.metric(l, v)
        with c2:
            _sec("💰", "Financeiro")
            for l, v in [("Recebido",fmt(ind["recebido_total"])),("Saldo",fmt(ind["saldo_total"])),
                          ("Empenhado",fmt(ind["empenhado_total"])),("Em Tela",fmt(ind["em_tela_total"])),
                          ("REQs Pendentes",fmt(ind["valor_reqs_pendentes"]))]:
                st.metric(l, v)
    elif tipo == "NCs Detalhado":
        if ncs: st.dataframe(pd.DataFrame(ncs), use_container_width=True, hide_index=True)
    elif tipo == "Requisições Detalhado":
        if reqs: st.dataframe(pd.DataFrame(reqs), use_container_width=True, hide_index=True)
    elif tipo == "NCs por Operação":
        if ncs: st.dataframe(ncs_por_operacao(ncs), use_container_width=True, hide_index=True)
    elif tipo == "NCs Próximas do Vencimento (30 dias)":
        venc = ncs_vencendo(ncs, 30)
        if venc:
            df = pd.DataFrame(venc)
            cols = [c for c in ["_DIAS_RESTANTES","NC","ORGÃO","FINALIDADE","PRAZO","SALDO NC","SITU"] if c in df.columns]
            st.dataframe(df[cols].rename(columns={"_DIAS_RESTANTES":"Dias Restantes"}),
                         use_container_width=True, hide_index=True)
        else:
            _alert("al-green", "✅ Nenhuma NC vencendo nos próximos 30 dias.")
    elif tipo == "REQs Pendentes":
        pend = [r for r in reqs if r.get("SITUAÇÃO") == "Pendente"]
        if pend:
            st.dataframe(pd.DataFrame(pend), use_container_width=True, hide_index=True)
            _alert("al-yellow", f'📝 {len(pend)} pendentes · {fmt(sum(parse(r.get("VALOR",0)) for r in pend))}')
        else:
            _alert("al-green", "✅ Nenhuma REQ pendente!")


# ── Assistente ────────────────────────────────────────────────────────────────
def page_assistente(ncs, reqs):
    st.markdown("## 🤖 Assistente SSAC")

    if not ANTHROPIC_API_KEY:
        _alert("al-red", "⚠️ Configure <b>ANTHROPIC_API_KEY</b> para usar o assistente.")
        return

    if "hist_api" not in st.session_state: st.session_state.hist_api = []
    if "hist_ui"  not in st.session_state: st.session_state.hist_ui  = []

    for msg in st.session_state.hist_ui:
        with st.chat_message(msg["role"]): st.markdown(msg["content"])

    if prompt := st.chat_input("Pergunte sobre NCs, saldos, prazos, requisições..."):
        st.session_state.hist_ui.append({"role":"user","content":prompt})
        with st.chat_message("user"): st.markdown(prompt)
        with st.chat_message("assistant"):
            with st.spinner(""):
                try:
                    from assistente import chat
                    resp = chat(prompt, st.session_state.hist_api, ncs, reqs)
                except Exception as e:
                    resp = f"Erro: {e}"
            st.markdown(resp)
        st.session_state.hist_api.append({"role":"user","content":prompt})
        st.session_state.hist_api.append({"role":"assistant","content":resp})
        st.session_state.hist_ui.append({"role":"assistant","content":resp})
        if len(st.session_state.hist_api) > 20:
            st.session_state.hist_api = st.session_state.hist_api[-20:]

    if st.session_state.hist_ui:
        if st.button("🗑️ Limpar"):
            st.session_state.hist_api = []; st.session_state.hist_ui = []; st.rerun()


# ── Main ──────────────────────────────────────────────────────────────────────
def main():
    requer_auth()

    for k in ("form_nc","form_req","page"):
        if k not in st.session_state:
            st.session_state[k] = False if k != "page" else "dashboard"

    ncs, reqs, _ = carregar()
    pagina = _sidebar(ncs, reqs)

    if   pagina == "dashboard":  page_dashboard(ncs, reqs)
    elif pagina == "ncs":        page_ncs(ncs)
    elif pagina == "reqs":       page_reqs(reqs, ncs)
    elif pagina == "pdf":        page_pdf(ncs)
    elif pagina == "importar":   page_importar(ncs, reqs)
    elif pagina == "relatorios": page_relatorios(ncs, reqs)
    elif pagina == "assistente": page_assistente(ncs, reqs)


if __name__ == "__main__":
    main()
