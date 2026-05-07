"""
Agente DIEX — Interface principal.
Execute: streamlit run interface.py
"""
import logging
import os
from datetime import date, datetime

import pandas as pd
import plotly.express as px
import streamlit as st

st.set_page_config(
    page_title="Agente DIEX",
    page_icon="🪖",
    layout="wide",
    initial_sidebar_state="expanded",
)

logging.basicConfig(level=logging.INFO)

from auth import logout, requer_auth
from config import ANTHROPIC_API_KEY, OM_PADRAO, UG_PADRAO
from relatorios import exportar_excel, kpis, ncs_por_operacao, ncs_vencendo

# ── CSS ───────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
/* Layout */
#MainMenu, footer {visibility:hidden;}
.block-container {padding:1.5rem 2rem 3rem !important;}

/* Sidebar */
section[data-testid="stSidebar"] > div:first-child {background:#0d1117 !important;}
section[data-testid="stSidebar"] hr {border-color:#21262d !important;}

/* Headings */
h1 {font-size:1.6rem !important; margin-bottom:.5rem !important;}
h2 {font-size:1.2rem !important;}
h3 {font-size:1rem !important;}

/* KPI cards */
.kpi-row {display:flex; gap:14px; margin-bottom:22px; flex-wrap:wrap;}
.kpi-card {
    flex:1; min-width:160px;
    background:#161b22; border:1px solid #21262d;
    border-radius:12px; padding:18px 16px; text-align:center;
    transition:border-color .2s, transform .15s;
}
.kpi-card:hover {border-color:#22c55e40; transform:translateY(-2px);}
.kpi-icon {font-size:1.4rem; margin-bottom:6px;}
.kpi-value {font-size:1.5rem; font-weight:700; margin:4px 0;}
.kpi-label {font-size:.68rem; color:#8b949e; text-transform:uppercase; letter-spacing:.8px;}
.kpi-sub {font-size:.72rem; color:#8b949e; margin-top:3px;}

/* Alerts */
.al {padding:11px 15px; border-radius:8px; margin:6px 0; font-size:.88rem; display:flex; align-items:center; gap:9px;}
.al-red    {background:#160808; border:1px solid #7f1d1d; color:#fca5a5;}
.al-yellow {background:#161008; border:1px solid #78350f; color:#fcd34d;}
.al-green  {background:#081408; border:1px solid #14532d; color:#86efac;}

/* Status badges */
.badge {
    display:inline-block; padding:2px 9px; border-radius:10px;
    font-size:.68rem; font-weight:600; letter-spacing:.3px;
}
.b-ok      {background:#052e16; color:#4ade80;}
.b-emtela  {background:#1c1107; color:#fbbf24;}
.b-pend    {background:#1c1107; color:#fbbf24;}
.b-emp     {background:#0c1a2e; color:#60a5fa;}
.b-pago    {background:#052e16; color:#4ade80;}
.b-envio   {background:#1c1218; color:#c084fc;}

/* Section label */
.sec-label {
    font-size:.67rem; font-weight:600; color:#8b949e;
    text-transform:uppercase; letter-spacing:1px;
    padding-bottom:6px; border-bottom:1px solid #21262d;
    margin-bottom:14px;
}

/* Table container */
[data-testid="stDataFrame"] {border-radius:8px; overflow:hidden;}
[data-testid="stDataFrame"] table {font-size:.82rem !important;}

/* Form */
[data-testid="stForm"] {
    background:#161b22; border:1px solid #21262d;
    border-radius:12px; padding:20px !important;
}
</style>
""", unsafe_allow_html=True)


# ── Helpers ───────────────────────────────────────────────────────────────────
def fmt(v: float) -> str:
    s = f"{v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return f"R$ {s}"


def parse(s) -> float:
    try:
        return float(str(s).replace("R$", "").replace(" ", "").replace(".", "").replace(",", "."))
    except Exception:
        return 0.0


def _badge(text: str) -> str:
    classes = {
        "OK": "b-ok", "EM TELA": "b-emtela",
        "Pendente": "b-pend", "Empenhada": "b-emp",
        "Paga": "b-pago", "Liquidada": "b-pago",
        "Enviada": "b-envio", "Aprovada": "b-envio",
    }
    cls = classes.get(text, "b-pend")
    return f'<span class="badge {cls}">{text}</span>'


def _cor_prazo(prazo_str: str) -> str:
    try:
        d = datetime.strptime(prazo_str, "%d/%m/%Y").date()
        n = (d - date.today()).days
        return "🔴" if n < 0 else "🟠" if n <= 7 else "🟡" if n <= 30 else "🟢"
    except Exception:
        return "⚪"


def _kpi_cards(ind: dict):
    total_rec = ind["recebido_total"]
    pct_emp = (ind["empenhado_total"] / total_rec * 100) if total_rec else 0
    html = f"""
    <div class="kpi-row">
      <div class="kpi-card">
        <div class="kpi-icon">💰</div>
        <div class="kpi-value" style="color:#22c55e">{fmt(ind['saldo_total'])}</div>
        <div class="kpi-label">Saldo Disponível</div>
        <div class="kpi-sub">{ind['ncs_ok']} NCs OK</div>
      </div>
      <div class="kpi-card">
        <div class="kpi-icon">📌</div>
        <div class="kpi-value" style="color:#3b82f6">{fmt(ind['empenhado_total'])}</div>
        <div class="kpi-label">Empenhado</div>
        <div class="kpi-sub">{pct_emp:.1f}% do recebido</div>
      </div>
      <div class="kpi-card">
        <div class="kpi-icon">🕐</div>
        <div class="kpi-value" style="color:#f59e0b">{fmt(ind['em_tela_total'])}</div>
        <div class="kpi-label">Em Tela</div>
        <div class="kpi-sub">{ind['ncs_em_tela']} NCs aguardando</div>
      </div>
      <div class="kpi-card">
        <div class="kpi-icon">📝</div>
        <div class="kpi-value" style="color:#a78bfa">{ind['reqs_pendentes']}</div>
        <div class="kpi-label">REQs Pendentes</div>
        <div class="kpi-sub">{fmt(ind['valor_reqs_pendentes'])}</div>
      </div>
    </div>
    """
    st.markdown(html, unsafe_allow_html=True)


ORGAOS = ["COTER", "COEX", "DGO", "DEC", "DECEX", "12 RM", "Outro"]
SITUACOES_REQ = ["Pendente", "Enviada", "Aprovada", "Empenhada", "Liquidada", "Paga"]


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
def _sidebar() -> str:
    with st.sidebar:
        st.markdown(f"""
        <div style="padding:8px 4px 16px;">
            <div style="font-size:1.5rem; font-weight:700; color:#e6edf3;">🪖 Agente DIEX</div>
            <div style="font-size:.8rem; color:#8b949e; margin-top:2px;">{OM_PADRAO} · UG {UG_PADRAO}</div>
        </div>
        """, unsafe_allow_html=True)

        st.divider()

        pagina = st.radio(
            "nav",
            ["📊  Dashboard", "📋  Notas de Crédito", "📝  Requisições",
             "📄  Lançar Documento", "📈  Relatórios", "🤖  Assistente"],
            label_visibility="collapsed",
        )

        st.divider()

        col_a, col_b = st.columns(2)
        with col_a:
            if st.button("🔄 Sync", use_container_width=True, help="Recarregar dados"):
                carregar(forcar=True)
                st.rerun()
        with col_b:
            if st.button("🚪 Sair", use_container_width=True):
                logout()

        st.caption(f"Sync · {datetime.now().strftime('%H:%M:%S')}")

    return pagina


# ── Dashboard ─────────────────────────────────────────────────────────────────
def page_dashboard(ncs, reqs):
    st.title("📊 Dashboard")
    st.markdown(f'<div class="sec-label">Visão Geral — {datetime.today().strftime("%d/%m/%Y")}</div>',
                unsafe_allow_html=True)

    ind = kpis(ncs, reqs)
    _kpi_cards(ind)

    # Alertas
    if ind["vencidas"] > 0:
        st.markdown(f'<div class="al al-red">⛔ <b>{ind["vencidas"]}</b> NC(s) com prazo vencido!</div>',
                    unsafe_allow_html=True)
    if ind["vencendo_7d"] > 0:
        st.markdown(f'<div class="al al-yellow">⚠️ <b>{ind["vencendo_7d"]}</b> NC(s) vencem nos próximos 7 dias.</div>',
                    unsafe_allow_html=True)
    if ind["reqs_pendentes"] > 0:
        st.markdown(f'<div class="al al-yellow">📝 <b>{ind["reqs_pendentes"]}</b> requisições pendentes — {fmt(ind["valor_reqs_pendentes"])}.</div>',
                    unsafe_allow_html=True)

    st.divider()

    # Gráficos linha 1
    c1, c2 = st.columns(2)
    with c1:
        st.markdown('<div class="sec-label">Saldo por Operação</div>', unsafe_allow_html=True)
        if ncs:
            df = (pd.DataFrame([{"OP": nc.get("OP") or "N/A", "Saldo": parse(nc.get("SALDO NC", 0))} for nc in ncs])
                  .groupby("OP")["Saldo"].sum().reset_index().query("Saldo > 0").sort_values("Saldo"))
            if not df.empty:
                fig = px.bar(df, x="Saldo", y="OP", orientation="h",
                             color_discrete_sequence=["#22c55e"], labels={"Saldo": "R$", "OP": ""})
                fig.update_layout(height=280, margin=dict(l=0,r=10,t=0,b=0),
                                  paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
                                  font_color="#8b949e", xaxis=dict(gridcolor="#21262d"),
                                  yaxis=dict(gridcolor="rgba(0,0,0,0)"))
                st.plotly_chart(fig, use_container_width=True)

    with c2:
        st.markdown('<div class="sec-label">Status das NCs</div>', unsafe_allow_html=True)
        if ncs:
            cnt = pd.DataFrame(ncs)["SITU"].value_counts().reset_index()
            cnt.columns = ["Status", "Qtd"]
            fig = px.pie(cnt, values="Qtd", names="Status", hole=0.5,
                         color_discrete_map={"OK": "#22c55e", "EM TELA": "#f59e0b"})
            fig.update_layout(height=280, margin=dict(l=0,r=0,t=10,b=20),
                              paper_bgcolor="rgba(0,0,0,0)", font_color="#8b949e",
                              legend=dict(orientation="h", y=-0.1))
            fig.update_traces(textfont_color="#e6edf3")
            st.plotly_chart(fig, use_container_width=True)

    # Gráficos linha 2
    c3, c4 = st.columns(2)
    with c3:
        st.markdown('<div class="sec-label">Empenho por Órgão</div>', unsafe_allow_html=True)
        if ncs:
            df_o = (pd.DataFrame([{"ORGÃO": nc.get("ORGÃO") or "N/A",
                                    "Recebido": parse(nc.get("RECEBIDO", 0)),
                                    "Empenhado": parse(nc.get("EMPENHADO", 0)),
                                    "Saldo": parse(nc.get("SALDO NC", 0))} for nc in ncs])
                    .groupby("ORGÃO").sum().reset_index()
                    .melt(id_vars="ORGÃO", var_name="Tipo", value_name="Valor"))
            fig = px.bar(df_o, x="ORGÃO", y="Valor", color="Tipo", barmode="group",
                         color_discrete_map={"Recebido":"#3b82f6","Empenhado":"#f59e0b","Saldo":"#22c55e"},
                         labels={"Valor":"R$"})
            fig.update_layout(height=280, margin=dict(l=0,r=0,t=0,b=0),
                              paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
                              font_color="#8b949e", yaxis=dict(gridcolor="#21262d"))
            st.plotly_chart(fig, use_container_width=True)

    with c4:
        st.markdown('<div class="sec-label">REQs por Situação</div>', unsafe_allow_html=True)
        if reqs:
            df_s = (pd.DataFrame([{"Situação": r.get("SITUAÇÃO") or "N/A", "Valor": parse(r.get("VALOR", 0))} for r in reqs])
                    .groupby("Situação")["Valor"].sum().reset_index())
            fig = px.bar(df_s, x="Situação", y="Valor",
                         color_discrete_sequence=["#a78bfa"],
                         labels={"Valor": "R$"})
            fig.update_layout(height=280, margin=dict(l=0,r=0,t=0,b=0),
                              paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
                              font_color="#8b949e", yaxis=dict(gridcolor="#21262d"), showlegend=False)
            st.plotly_chart(fig, use_container_width=True)

    # Tabela de prazos
    st.markdown('<div class="sec-label" style="margin-top:8px;">Controle de Prazos</div>',
                unsafe_allow_html=True)
    if ncs:
        hoje = date.today()
        rows = []
        for nc in ncs:
            prazo_str = nc.get("PRAZO", "")
            try:
                d = datetime.strptime(prazo_str, "%d/%m/%Y").date()
                dias = (d - hoje).days
            except Exception:
                dias = None
            rows.append({
                "": _cor_prazo(prazo_str),
                "NC": nc.get("NC", ""),
                "ÓRGÃO": nc.get("ORGÃO", ""),
                "FINALIDADE": nc.get("FINALIDADE", "")[:50],
                "PRAZO": prazo_str,
                "Dias": dias if dias is not None else "—",
                "Saldo": nc.get("SALDO NC", ""),
                "Status": nc.get("SITU", ""),
            })
        df_p = pd.DataFrame(rows).sort_values("Dias", key=lambda x: pd.to_numeric(x, errors="coerce"))
        st.dataframe(df_p, use_container_width=True, hide_index=True)


# ── NCs ───────────────────────────────────────────────────────────────────────
def page_ncs(ncs):
    st.title("📋 Notas de Crédito")

    with st.expander("🔍 Filtros", expanded=False):
        f1, f2, f3 = st.columns(3)
        f_situ = f1.multiselect("Status", ["OK", "EM TELA"], default=["OK", "EM TELA"])
        f_org  = f2.multiselect("Órgão", sorted({nc.get("ORGÃO","") for nc in ncs if nc.get("ORGÃO")}))
        f_op   = f3.multiselect("Operação", sorted({nc.get("OP","") for nc in ncs if nc.get("OP")}))

    filtradas = [
        nc for nc in ncs
        if (not f_situ or nc.get("SITU") in f_situ)
        and (not f_org  or nc.get("ORGÃO") in f_org)
        and (not f_op   or nc.get("OP") in f_op)
    ]

    b1, b2 = st.columns([1, 3])
    with b1:
        if st.button("➕ Nova NC", type="primary", use_container_width=True):
            st.session_state["form_nc"] = not st.session_state.get("form_nc", False)
    with b2:
        saldo_f = sum(parse(nc.get("SALDO NC", 0)) for nc in filtradas)
        st.markdown(f'<div class="al al-green">📋 <b>{len(filtradas)}</b> NCs · Saldo filtrado: <b>{fmt(saldo_f)}</b></div>',
                    unsafe_allow_html=True)

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
    st.markdown("---")
    st.markdown('<div class="sec-label">Nova Nota de Crédito</div>', unsafe_allow_html=True)
    with st.form("f_nc", clear_on_submit=True):
        c1, c2 = st.columns(2)
        nc_num  = c1.text_input("Número NC *", placeholder="2026NC000000")
        orgao   = c1.selectbox("Órgão *", ORGAOS)
        data_nc = c1.date_input("Data NC *", value=date.today())
        pi      = c1.text_input("PI")
        nd      = c1.text_input("ND", placeholder="339030")

        ptres   = c2.text_input("PTRES", placeholder="251050")
        om      = c2.text_input("OM", value=OM_PADRAO)
        prazo   = c2.date_input("Prazo *")
        op      = c2.text_input("Operação (OP)")
        situ    = c2.selectbox("Status", ["OK", "EM TELA"])

        finalidade = st.text_area("Finalidade *")
        valor = st.number_input("Valor Recebido (R$) *", min_value=0.0, step=0.01, format="%.2f")

        s1, s2 = st.columns(2)
        salvar  = s1.form_submit_button("💾 Salvar", type="primary", use_container_width=True)
        cancelar = s2.form_submit_button("✖ Cancelar", use_container_width=True)

        if cancelar:
            st.session_state["form_nc"] = False
            st.rerun()
        if salvar:
            if not nc_num or not finalidade or valor <= 0:
                st.error("Preencha NC, Finalidade e Valor.")
            else:
                try:
                    from sheets_nc import adicionar_nc
                    adicionar_nc({"NC": nc_num, "ORGÃO": orgao,
                                  "DATA NC": data_nc.strftime("%d/%m/%Y"),
                                  "PI": pi, "ND": nd, "PTRES": ptres, "OM": om,
                                  "PRAZO": prazo.strftime("%d/%m/%Y"),
                                  "OP": op, "SITU": situ,
                                  "FINALIDADE": finalidade, "RECEBIDO": valor})
                    st.success(f"✅ NC {nc_num} adicionada!")
                    st.session_state["form_nc"] = False
                    carregar(forcar=True)
                    st.rerun()
                except Exception as e:
                    st.error(f"Erro: {e}")


# ── Requisições ───────────────────────────────────────────────────────────────
def page_reqs(reqs, ncs):
    st.title("📝 Requisições")

    with st.expander("🔍 Filtros", expanded=False):
        f1, f2, f3 = st.columns(3)
        f_sit = f1.multiselect("Situação", sorted({r.get("SITUAÇÃO","") for r in reqs if r.get("SITUAÇÃO")}))
        f_emp = f2.multiselect("Empresa",  sorted({r.get("EMPRESA","")  for r in reqs if r.get("EMPRESA")}))
        f_nc  = f3.multiselect("NC",       sorted({r.get("NC","")       for r in reqs if r.get("NC")}))

    filtradas = [
        r for r in reqs
        if (not f_sit or r.get("SITUAÇÃO") in f_sit)
        and (not f_emp or r.get("EMPRESA") in f_emp)
        and (not f_nc  or r.get("NC") in f_nc)
    ]

    b1, b2 = st.columns([1, 3])
    with b1:
        if st.button("➕ Nova REQ", type="primary", use_container_width=True):
            st.session_state["form_req"] = not st.session_state.get("form_req", False)
    with b2:
        total = sum(parse(r.get("VALOR", 0)) for r in filtradas)
        st.markdown(f'<div class="al al-green">📝 <b>{len(filtradas)}</b> REQs · Total: <b>{fmt(total)}</b></div>',
                    unsafe_allow_html=True)

    if st.session_state.get("form_req"):
        _form_req(ncs)

    if filtradas:
        cols = [c for c in ["REQ","DATA REQ","NC","PI","EMPRESA","DESCRIÇÃO","VALOR","SITUAÇÃO","NE","FINALIDADE","OBS"]
                if c in pd.DataFrame(filtradas).columns]
        st.dataframe(pd.DataFrame(filtradas)[cols], use_container_width=True, hide_index=True)
    else:
        st.info("Nenhuma REQ encontrada.")


def _form_req(ncs):
    st.markdown("---")
    st.markdown('<div class="sec-label">Nova Requisição</div>', unsafe_allow_html=True)
    nums_nc = [""] + [nc.get("NC","") for nc in ncs if nc.get("NC")]

    with st.form("f_req", clear_on_submit=True):
        c1, c2 = st.columns(2)
        req_num  = c1.text_input("Número REQ")
        nc_sel   = c1.selectbox("NC Vinculada *", nums_nc)
        data_req = c1.date_input("Data REQ", value=date.today())
        empresa  = c1.text_input("Empresa *")

        pi      = c2.text_input("PI")
        tipo    = c2.selectbox("Tipo", ["Ordinário", "Especial"])
        ne      = c2.text_input("NE")
        situacao = c2.selectbox("Situação", SITUACOES_REQ)

        finalidade = st.text_area("Finalidade")
        descricao  = st.text_area("Descrição dos itens/serviços *")
        valor = st.number_input("Valor (R$) *", min_value=0.0, step=0.01, format="%.2f")
        obs   = st.text_input("Observações")

        s1, s2 = st.columns(2)
        salvar   = s1.form_submit_button("💾 Salvar", type="primary", use_container_width=True)
        cancelar = s2.form_submit_button("✖ Cancelar", use_container_width=True)

        if cancelar:
            st.session_state["form_req"] = False
            st.rerun()
        if salvar:
            if not empresa or not descricao or valor <= 0:
                st.error("Preencha Empresa, Descrição e Valor.")
            else:
                try:
                    pi_f = pi
                    fin_f = finalidade
                    if nc_sel and not pi:
                        nc_d = next((nc for nc in ncs if nc.get("NC") == nc_sel), {})
                        pi_f = nc_d.get("PI", "")
                        fin_f = fin_f or nc_d.get("OP", "")
                    from sheets_nc import adicionar_req
                    adicionar_req({"REQ": req_num, "NC": nc_sel,
                                   "DATA REQ": data_req.strftime("%d/%m/%Y"),
                                   "PI": pi_f, "FINALIDADE": fin_f, "TIPO": tipo,
                                   "EMPRESA": empresa, "DESCRIÇÃO": descricao,
                                   "VALOR": valor, "SITUAÇÃO": situacao, "NE": ne, "OBS": obs})
                    st.success("✅ REQ adicionada!")
                    st.session_state["form_req"] = False
                    carregar(forcar=True)
                    st.rerun()
                except Exception as e:
                    st.error(f"Erro: {e}")


# ── Lançar por PDF / HTML ─────────────────────────────────────────────────────
def page_pdf(ncs):
    st.title("📄 Lançar por PDF ou HTML")

    if not ANTHROPIC_API_KEY:
        st.markdown('<div class="al al-red">⚠️ Configure <b>ANTHROPIC_API_KEY</b> no .env para usar este recurso.</div>',
                    unsafe_allow_html=True)
        return

    tipo = st.radio("Tipo de documento", ["Nota de Crédito (NC)", "Requisição (REQ)"], horizontal=True)

    uploaded = st.file_uploader(
        "Selecione o arquivo",
        type=["pdf", "html", "htm"],
        help="PDF ou HTML exportado do SIAFI. O Claude extrai os dados automaticamente.",
    )

    if uploaded:
        with st.spinner("🤖 Analisando documento..."):
            try:
                from extrator_pdf import extrair_nc, extrair_req
                b = uploaded.read()
                if "NC" in tipo:
                    dados = extrair_nc(b, uploaded.name)
                    st.markdown('<div class="al al-green">✅ Dados extraídos — revise e confirme abaixo.</div>',
                                unsafe_allow_html=True)
                    _confirmar_nc(dados)
                else:
                    dados = extrair_req(b, uploaded.name)
                    st.markdown('<div class="al al-green">✅ Dados extraídos — revise e confirme abaixo.</div>',
                                unsafe_allow_html=True)
                    _confirmar_req(dados, ncs)
            except Exception as e:
                st.error(f"Erro na extração: {e}")


def _confirmar_nc(dados):
    st.markdown("---")
    st.markdown('<div class="sec-label">Confirmar Nota de Crédito</div>', unsafe_allow_html=True)
    with st.form("f_pdf_nc"):
        c1, c2 = st.columns(2)
        nc     = c1.text_input("Número NC",            value=dados.get("NC",""))
        orgao  = c1.text_input("Órgão",                value=dados.get("ORGÃO",""))
        data_nc = c1.text_input("Data NC (DD/MM/YYYY)", value=dados.get("DATA NC",""))
        pi     = c1.text_input("PI",                   value=dados.get("PI",""))
        nd     = c1.text_input("ND",                   value=dados.get("ND",""))

        ptres  = c2.text_input("PTRES",                value=dados.get("PTRES",""))
        om     = c2.text_input("OM",                   value=dados.get("OM", OM_PADRAO))
        prazo  = c2.text_input("Prazo (DD/MM/YYYY)",   value=dados.get("PRAZO",""))
        op     = c2.text_input("Operação",             value=dados.get("OP",""))
        situ   = c2.selectbox("Status", ["OK","EM TELA"],
                               index=1 if dados.get("SITU")=="EM TELA" else 0)

        finalidade = st.text_area("Finalidade", value=dados.get("FINALIDADE",""))
        valor = st.number_input("Valor (R$)", value=float(dados.get("RECEBIDO",0) or 0),
                                min_value=0.0, step=0.01, format="%.2f")

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
    st.markdown("---")
    st.markdown('<div class="sec-label">Confirmar Requisição</div>', unsafe_allow_html=True)
    nums_nc = [""] + [nc.get("NC","") for nc in ncs if nc.get("NC")]
    nc_idx = nums_nc.index(dados.get("NC","")) if dados.get("NC","") in nums_nc else 0

    with st.form("f_pdf_req"):
        c1, c2 = st.columns(2)
        req      = c1.text_input("Número REQ",          value=str(dados.get("REQ","")))
        nc_sel   = c1.selectbox("NC Vinculada",         nums_nc, index=nc_idx)
        data_req = c1.text_input("Data REQ (DD/MM/YYYY)", value=dados.get("DATA REQ",""))
        empresa  = c1.text_input("Empresa",             value=dados.get("EMPRESA",""))

        pi     = c2.text_input("PI",      value=dados.get("PI",""))
        tipo   = c2.selectbox("Tipo", ["Ordinário","Especial"])
        ne     = c2.text_input("NE",      value=dados.get("NE",""))
        situacao = c2.selectbox("Situação", SITUACOES_REQ)

        finalidade = st.text_area("Finalidade", value=dados.get("FINALIDADE",""))
        descricao  = st.text_area("Descrição",  value=dados.get("DESCRIÇÃO",""))
        valor = st.number_input("Valor (R$)", value=float(dados.get("VALOR",0) or 0),
                                min_value=0.0, step=0.01, format="%.2f")

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


# ── Relatórios ────────────────────────────────────────────────────────────────
def page_relatorios(ncs, reqs):
    st.title("📈 Relatórios")

    tipo = st.selectbox("Tipo", [
        "Resumo Geral", "NCs Detalhado", "Requisições Detalhado",
        "NCs por Operação", "NCs Próximas do Vencimento (30 dias)", "REQs Pendentes",
    ])

    _exibir_relatorio(tipo, ncs, reqs)

    st.divider()
    st.markdown('<div class="sec-label">Exportar</div>', unsafe_allow_html=True)
    c1, c2 = st.columns([1, 3])
    with c1:
        st.download_button(
            "⬇️ Excel Completo",
            data=exportar_excel(ncs, reqs),
            file_name=f"diex_{datetime.today().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    with c2:
        st.caption("NCs · Requisições · Resumo · Por Operação em abas separadas.")


def _exibir_relatorio(tipo, ncs, reqs):
    if tipo == "Resumo Geral":
        ind = kpis(ncs, reqs)
        c1, c2 = st.columns(2)
        with c1:
            st.markdown('<div class="sec-label">NCs</div>', unsafe_allow_html=True)
            st.metric("Total", ind["total_ncs"])
            st.metric("OK", ind["ncs_ok"])
            st.metric("EM TELA", ind["ncs_em_tela"])
            st.metric("Vencendo em 7 dias", ind["vencendo_7d"])
            st.metric("Vencidas", ind["vencidas"])
        with c2:
            st.markdown('<div class="sec-label">Financeiro</div>', unsafe_allow_html=True)
            st.metric("Total Recebido", fmt(ind["recebido_total"]))
            st.metric("Saldo Disponível", fmt(ind["saldo_total"]))
            st.metric("Empenhado", fmt(ind["empenhado_total"]))
            st.metric("Em Tela", fmt(ind["em_tela_total"]))
            st.metric("REQs Pendentes", fmt(ind["valor_reqs_pendentes"]))

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
            cols = [c for c in ["_DIAS_RESTANTES","NC","ORGÃO","FINALIDADE","PRAZO","SALDO NC","SITU"] if c in df.columns]
            st.dataframe(df[cols].rename(columns={"_DIAS_RESTANTES":"Dias Restantes"}),
                         use_container_width=True, hide_index=True)
        else:
            st.markdown('<div class="al al-green">✅ Nenhuma NC vencendo nos próximos 30 dias.</div>',
                        unsafe_allow_html=True)

    elif tipo == "REQs Pendentes":
        pend = [r for r in reqs if r.get("SITUAÇÃO") == "Pendente"]
        if pend:
            st.dataframe(pd.DataFrame(pend), use_container_width=True, hide_index=True)
            st.markdown(f'<div class="al al-yellow">📝 <b>{len(pend)}</b> pendentes · {fmt(sum(parse(r.get("VALOR",0)) for r in pend))}</div>',
                        unsafe_allow_html=True)
        else:
            st.markdown('<div class="al al-green">✅ Nenhuma REQ pendente!</div>', unsafe_allow_html=True)


# ── Assistente ────────────────────────────────────────────────────────────────
def page_assistente(ncs, reqs):
    st.title("🤖 Assistente DIEX")

    if not ANTHROPIC_API_KEY:
        st.markdown('<div class="al al-red">⚠️ Configure <b>ANTHROPIC_API_KEY</b> no .env para usar o assistente.</div>',
                    unsafe_allow_html=True)
        return

    ind = kpis(ncs, reqs)
    with st.sidebar:
        st.divider()
        st.caption(f"**{len(ncs)}** NCs · **{len(reqs)}** REQs")
        st.caption(f"Saldo: {fmt(ind['saldo_total'])}")

    if "hist_api" not in st.session_state:
        st.session_state.hist_api = []
    if "hist_ui" not in st.session_state:
        st.session_state.hist_ui = []

    for msg in st.session_state.hist_ui:
        with st.chat_message(msg["role"]):
            st.markdown(msg["content"])

    if prompt := st.chat_input("Pergunte sobre NCs, saldos, prazos, requisições..."):
        st.session_state.hist_ui.append({"role": "user", "content": prompt})
        with st.chat_message("user"):
            st.markdown(prompt)

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
            st.session_state.hist_api = []
            st.session_state.hist_ui = []
            st.rerun()


# ── Main ──────────────────────────────────────────────────────────────────────
def main():
    requer_auth()

    for k in ("form_nc", "form_req"):
        if k not in st.session_state:
            st.session_state[k] = False

    pagina = _sidebar()
    ncs, reqs, _ = carregar()

    if "Dashboard"   in pagina: page_dashboard(ncs, reqs)
    elif "Notas"     in pagina: page_ncs(ncs)
    elif "Requisições" in pagina: page_reqs(reqs, ncs)
    elif "Documento" in pagina: page_pdf(ncs)
    elif "Relatórios" in pagina: page_relatorios(ncs, reqs)
    elif "Assistente" in pagina: page_assistente(ncs, reqs)


if __name__ == "__main__":
    main()
