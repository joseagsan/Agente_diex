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

# ── CSS (apenas layout e sidebar) ────────────────────────────────────────────
st.markdown("""
<style>
#MainMenu, footer, header {visibility:hidden;}
.block-container {padding:1.5rem 2rem 3rem !important; max-width:100% !important;}

section[data-testid="stSidebar"] {width:230px !important; min-width:230px !important;}
section[data-testid="stSidebar"] > div:first-child {
    background:#080e1a !important;
    border-right:1px solid #1a2540 !important;
}

/* Nav buttons */
section[data-testid="stSidebar"] [data-testid="baseButton-secondary"] {
    text-align:left !important; justify-content:flex-start !important;
    background:transparent !important; border:none !important;
    color:#4b6080 !important; font-size:.875rem !important; font-weight:500 !important;
    padding:9px 12px !important; border-radius:8px !important; margin:1px 0 !important;
}
section[data-testid="stSidebar"] [data-testid="baseButton-secondary"]:hover {
    background:#0f1f35 !important; color:#c8d8e8 !important;
}
section[data-testid="stSidebar"] [data-testid="baseButton-primary"] {
    text-align:left !important; justify-content:flex-start !important;
    background:linear-gradient(135deg,#052918,#073d22) !important;
    border:1px solid #0a4a2840 !important; color:#34d399 !important;
    font-size:.875rem !important; font-weight:600 !important;
    padding:9px 12px !important; border-radius:8px !important; margin:1px 0 !important;
}
section[data-testid="stSidebar"] [data-testid="baseButton-primary"]:hover {
    background:linear-gradient(135deg,#063520,#094d2a) !important;
}
section[data-testid="stSidebar"] hr {border-color:#1a2540 !important; margin:8px 0 !important;}

[data-testid="stDataFrame"] {border:1px solid #1a2540 !important; border-radius:10px !important;}
[data-testid="stForm"] {background:#0d1829 !important; border:1px solid #1a2540 !important; border-radius:12px;}
[data-testid="stExpander"] {border:1px solid #1a2540 !important; border-radius:10px !important;}
hr {border-color:#1a2540 !important;}
</style>
""", unsafe_allow_html=True)

_CL = dict(
    paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", font_color="#64748b",
    margin=dict(l=0, r=0, t=20, b=0),
    xaxis=dict(gridcolor="#1a2540", zerolinecolor="#1a2540"),
    yaxis=dict(gridcolor="#1a2540", zerolinecolor="#1a2540"),
    legend=dict(orientation="h", y=1.15, font=dict(size=11)),
)

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
            <div style="font-size:.72rem;color:#4b6080;margin-top:3px;line-height:1.5">
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
                <div style="color:#4b6080;font-size:.63rem;text-transform:uppercase;letter-spacing:.8px;margin-bottom:8px;">Resumo</div>
                <div style="display:flex;justify-content:space-between;margin-bottom:5px;color:#8ba0b8">
                    <span>Saldo</span><span style="color:#34d399;font-weight:700">{fmt(ind['saldo_total'])}</span>
                </div>
                <div style="display:flex;justify-content:space-between;margin-bottom:5px;color:#8ba0b8">
                    <span>EM TELA</span><span style="color:#fbbf24;font-weight:700">{ind['ncs_em_tela']} NCs</span>
                </div>
                <div style="display:flex;justify-content:space-between;color:#8ba0b8">
                    <span>REQs pend.</span><span style="color:#a78bfa;font-weight:700">{ind['reqs_pendentes']}</span>
                </div>
            </div>""", unsafe_allow_html=True)
            st.divider()

        c1, c2 = st.columns(2)
        with c1:
            if st.button("🔄 Sync", use_container_width=True, help="Recarregar dados"):
                carregar(forcar=True); st.rerun()
        with c2:
            if st.button("🚪 Sair", use_container_width=True):
                logout()

        st.caption(f"Sync {datetime.now().strftime('%H:%M:%S')}")

    return st.session_state.get("page", "dashboard")


# ── Dashboard ─────────────────────────────────────────────────────────────────
def page_dashboard(ncs, reqs):
    st.title("📊 Dashboard")
    st.caption(f"{OM_PADRAO} · {datetime.today().strftime('%d/%m/%Y')}")

    ind = kpis(ncs, reqs)
    total_rec = ind["recebido_total"]
    pct_emp = (ind["empenhado_total"] / total_rec * 100) if total_rec else 0

    # KPIs nativos
    k1, k2, k3, k4 = st.columns(4)
    k1.metric("💰 Saldo Disponível",  fmt(ind["saldo_total"]),
              delta=f"{ind['ncs_ok']} NCs OK", delta_color="off")
    k2.metric("📌 Empenhado",         fmt(ind["empenhado_total"]),
              delta=f"{pct_emp:.1f}% do recebido", delta_color="off")
    k3.metric("🕐 Em Tela",           fmt(ind["em_tela_total"]),
              delta=f"{ind['ncs_em_tela']} NCs aguardando", delta_color="off")
    k4.metric("📝 REQs Pendentes",    str(ind["reqs_pendentes"]),
              delta=fmt(ind["valor_reqs_pendentes"]), delta_color="off")

    # Alertas nativos
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
        for nc in ncs:
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
            fig.add_bar(x=df_m["Mês"], y=df_m["Recebido"],  name="Recebido",  marker_color="#3b82f6")
            fig.add_bar(x=df_m["Mês"], y=df_m["Empenhado"], name="Empenhado", marker_color="#10b981")
            fig.update_layout(height=320, barmode="group", **_CL)
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("Sem dados de data para gráfico mensal.")

    with c2:
        st.subheader("🔵 Status das NCs")
        if ncs:
            cnt = pd.DataFrame(ncs)["SITU"].value_counts().reset_index()
            cnt.columns = ["Status", "Qtd"]
            fig = px.pie(cnt, values="Qtd", names="Status", hole=0.55,
                         color_discrete_map={"OK": "#10b981", "EM TELA": "#fbbf24"})
            fig.update_layout(height=320, **{**_CL, "margin": dict(l=0, r=0, t=20, b=30)})
            fig.update_traces(textfont_color="#e2e8f0", textinfo="percent+label")
            st.plotly_chart(fig, use_container_width=True)

    # Gráfico 2: Saldo por operação + Empenho por órgão
    c3, c4 = st.columns(2)
    with c3:
        st.subheader("🎯 Saldo por Operação")
        if ncs:
            df_op = (pd.DataFrame([{"OP": nc.get("OP") or "Sem OP",
                                     "Saldo": parse(nc.get("SALDO NC", 0))} for nc in ncs])
                     .groupby("OP")["Saldo"].sum().reset_index()
                     .query("Saldo > 0").sort_values("Saldo"))
            if not df_op.empty:
                fig = px.bar(df_op, x="Saldo", y="OP", orientation="h",
                             color="Saldo",
                             color_continuous_scale=["#1a4a2e", "#10b981"],
                             labels={"Saldo": "R$", "OP": ""})
                fig.update_coloraxes(showscale=False)
                fig.update_layout(height=340, **_CL)
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.success("Sem saldo disponível no momento.")

    with c4:
        st.subheader("🏛️ Recebido vs Empenhado por Órgão")
        if ncs:
            df_o = (pd.DataFrame([{"ORGÃO": nc.get("ORGÃO") or "N/A",
                                    "Recebido": parse(nc.get("RECEBIDO", 0)),
                                    "Empenhado": parse(nc.get("EMPENHADO", 0))} for nc in ncs])
                    .groupby("ORGÃO").sum().reset_index()
                    .melt(id_vars="ORGÃO", var_name="Tipo", value_name="Valor"))
            fig = px.bar(df_o, x="ORGÃO", y="Valor", color="Tipo", barmode="group",
                         color_discrete_map={"Recebido": "#3b82f6", "Empenhado": "#10b981"},
                         labels={"Valor": "R$", "ORGÃO": ""})
            fig.update_layout(height=340, **_CL)
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
                     "Empenhada": "#3b82f6", "Liquidada": "#10b981", "Paga": "#34d399"}
            fig = px.bar(df_s, x="Situação", y="Valor", color="Situação",
                         color_discrete_map=cores, labels={"Valor": "R$", "Situação": ""})
            fig.update_layout(height=320, **_CL, showlegend=False)
            st.plotly_chart(fig, use_container_width=True)

    with c6:
        st.subheader("🕐 NCs EM TELA")
        em_tela = [nc for nc in ncs if nc.get("SITU") == "EM TELA"]
        if em_tela:
            total_et = sum(parse(nc.get("RECEBIDO", 0)) for nc in em_tela)
            st.warning(f"⚠️ {len(em_tela)} NCs em tela — {fmt(total_et)}")
            df_et = pd.DataFrame([{
                "NC": nc.get("NC", ""), "ÓRGÃO": nc.get("ORGÃO", ""),
                "Finalidade": nc.get("FINALIDADE", "")[:38],
                "Valor": nc.get("RECEBIDO", ""), "Prazo": nc.get("PRAZO", ""),
            } for nc in em_tela])
            st.dataframe(df_et, use_container_width=True, hide_index=True, height=260)
        else:
            st.success("✅ Nenhuma NC EM TELA no momento.")

    # Gráfico 4: Valores totais por operação
    st.subheader("💹 Valores por Operação")
    if ncs:
        df_vop = (pd.DataFrame([{"OP": nc.get("OP") or "Sem OP",
                                  "Recebido": parse(nc.get("RECEBIDO", 0)),
                                  "Empenhado": parse(nc.get("EMPENHADO", 0)),
                                  "Saldo": parse(nc.get("SALDO NC", 0))} for nc in ncs])
                  .groupby("OP").sum().reset_index().sort_values("Recebido", ascending=False))
        fig = go.Figure()
        fig.add_bar(x=df_vop["OP"], y=df_vop["Recebido"],  name="Recebido",  marker_color="#3b82f6")
        fig.add_bar(x=df_vop["OP"], y=df_vop["Empenhado"], name="Empenhado", marker_color="#fbbf24")
        fig.add_bar(x=df_vop["OP"], y=df_vop["Saldo"],     name="Saldo",     marker_color="#10b981")
        fig.update_layout(barmode="group", height=320, **_CL)
        st.plotly_chart(fig, use_container_width=True)

    # Tabela de prazos
    st.subheader("⏰ Controle de Prazos")
    if ncs:
        hoje = date.today()
        rows = []
        for nc in ncs:
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
            try:
                dias = (datetime.strptime(nc.get("PRAZO", ""), "%d/%m/%Y").date() - hoje).days
                if f_prazo == "Vencidas"           and dias >= 0:          continue
                if f_prazo == "Vencem em 7 dias"   and not 0 <= dias <= 7:  continue
                if f_prazo == "Vencem em 30 dias"  and not 0 <= dias <= 30: continue
            except Exception:
                continue
        filtradas.append(nc)

    b1, b2 = st.columns([1, 3])
    with b1:
        if st.button("➕ Nova NC", type="primary", use_container_width=True):
            st.session_state["form_nc"] = not st.session_state.get("form_nc", False)
    with b2:
        saldo_f = sum(parse(nc.get("SALDO NC", 0)) for nc in filtradas)
        st.info(f"📋 **{len(filtradas)}** NCs · Saldo filtrado: **{fmt(saldo_f)}**")

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

    b1, b2 = st.columns([1, 3])
    with b1:
        if st.button("➕ Nova REQ", type="primary", use_container_width=True):
            st.session_state["form_req"] = not st.session_state.get("form_req", False)
    with b2:
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

    st.subheader("🔗 Status da Conexão")
    k1, k2, k3, k4 = st.columns(4)
    k1.metric("NCs carregadas",    len(ncs))
    k2.metric("REQs carregadas",   len(reqs))
    k3.metric("NCs OK",            sum(1 for nc in ncs if nc.get("SITU") == "OK"))
    k4.metric("REQs Pendentes",    sum(1 for r  in reqs if r.get("SITUAÇÃO") == "Pendente"))

    st.info(f"🔗 Planilha conectada: `{SHEET_ID_NC[:30]}...`  ·  Sync automático a cada 2 min.")

    if st.button("🔄 Forçar Sincronização", type="primary"):
        with st.spinner("Sincronizando..."):
            carregar(forcar=True)
        st.success("✅ Dados atualizados!"); st.rerun()

    st.divider()
    st.subheader("📂 Importar de Arquivo (Excel / CSV)")
    st.info("ℹ️ Importe NCs ou REQs de uma planilha local. Os dados serão adicionados à planilha oficial.")

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
    tipo = st.selectbox("Tipo", ["Resumo Geral", "NCs Detalhado", "Requisições Detalhado",
                                  "NCs por Operação", "NCs Próximas do Vencimento (30 dias)", "REQs Pendentes"])
    _exibir_relatorio(tipo, ncs, reqs)
    st.divider()
    st.subheader("📥 Exportar")
    c1, _ = st.columns([1, 3])
    with c1:
        st.download_button("⬇️ Excel Completo",
                           data=exportar_excel(ncs, reqs),
                           file_name=f"ssac_{datetime.today().strftime('%Y%m%d_%H%M')}.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           use_container_width=True)


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
        if ncs: st.dataframe(pd.DataFrame(ncs), use_container_width=True, hide_index=True)
    elif tipo == "Requisições Detalhado":
        if reqs: st.dataframe(pd.DataFrame(reqs), use_container_width=True, hide_index=True)
    elif tipo == "NCs por Operação":
        if ncs: st.dataframe(ncs_por_operacao(ncs), use_container_width=True, hide_index=True)
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


# ── Main ──────────────────────────────────────────────────────────────────────
def main():
    requer_auth()
    for k in ("form_nc", "form_req"):
        if k not in st.session_state: st.session_state[k] = False
    if "page" not in st.session_state: st.session_state["page"] = "dashboard"

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
