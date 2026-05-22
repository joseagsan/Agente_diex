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
    bg="transparent", card_bg="#ffffff", card_border="#e0e0e0",
    text="#1a1a2e", subtext="#555555", accent="#449D44",
    sidebar_bg="#2d4a2d", sidebar_border="#3a5c3a",
    nav_color="#8fac8f", nav_hover_bg="#3a5c3a", nav_hover_color="#f0faf0",
    active_bg="linear-gradient(135deg,#3b8c3b,#449D44)", active_border="#449D4460",
    active_color="#f0faf0",
    plot_bg="rgba(255,255,255,0)", paper_bg="rgba(255,255,255,0)",
    grid="#dee2e6", font_color="#555555",
    bar_recv="#337ab7", bar_emp="#449D44", bar_saldo="#27ae60",
)


def _t() -> dict:
    return _DARK if st.session_state.get("tema_escuro") else _LIGHT


def _inject_css():
    t   = _t()
    esc = st.session_state.get("tema_escuro", False)
    dark_override = """
[data-testid="stAppViewContainer"],[data-testid="stMain"]{background-color:#0d1117!important}
[data-testid="stHeader"]{background:#0d1117!important}
.block-container,[data-testid="stVerticalBlockBorderWrapper"]{background-color:#0d1117!important}
p,label,h1,h2,h3,h4,span,[data-testid="stMarkdownContainer"]{color:#e6edf3!important}
[data-testid="stMetricValue"],[data-testid="stMetricLabel"],[data-testid="stMetricDelta"]{color:#e6edf3!important}
[data-testid="stTextInput"] input,[data-testid="stNumberInput"] input,[data-testid="stTextArea"] textarea{background:#161b22!important;color:#e6edf3!important;border-color:#30363d!important}
[data-baseweb="select"] div,[data-baseweb="popover"] ul{background:#161b22!important;color:#e6edf3!important}
[data-testid="stDataFrameResizable"]{background:#0d1117!important}
""" if esc else ""
    st.markdown(f"""
<style>
{dark_override}
#MainMenu, footer {{visibility:hidden;}}
.block-container {{padding:1.5rem 2rem 3rem !important; max-width:100% !important;}}

/* Expande conteúdo quando sidebar está recolhida */
section[data-testid="stSidebar"][aria-expanded="false"] ~ div[data-testid="stMain"],
section[data-testid="stSidebar"][aria-expanded="false"] ~ section[data-testid="stMain"] {{
    margin-left:0 !important;
    width:100% !important;
}}
section[data-testid="stSidebar"] {{ transition: width .3s ease !important; }}

/* ── Sidebar ── */
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
    color:{t['active_color']} !important; font-size:.875rem !important; font-weight:700 !important;
    padding:9px 12px !important; border-radius:8px !important; margin:1px 0 !important;
}}
section[data-testid="stSidebar"] [data-testid="baseButton-primary"]:hover {{ filter:brightness(1.1); }}
section[data-testid="stSidebar"] hr {{border-color:{t['sidebar_border']} !important; margin:8px 0 !important;}}

/* ── Cards / Forms / Expanders ── */
[data-testid="stDataFrame"] {{
    border:1px solid {t['card_border']} !important; border-radius:8px !important;
    box-shadow:0 2px 6px rgba(0,0,0,.06) !important;
}}
[data-testid="stForm"] {{
    background:{t['card_bg']} !important;
    border:1px solid {t['card_border']} !important;
    border-left:3px solid {t['accent']} !important;
    border-radius:8px !important;
    box-shadow:0 2px 6px rgba(0,0,0,.06) !important;
}}
[data-testid="stExpander"] {{
    border:1px solid {t['card_border']} !important;
    border-radius:8px !important;
    background:{t['card_bg']} !important;
    box-shadow:0 2px 6px rgba(0,0,0,.06) !important;
}}
hr {{border-color:{t['card_border']} !important;}}

/* ── Metric cards ── */
[data-testid="stMetric"] {{
    background:{t['card_bg']};
    border:1px solid {t['card_border']};
    border-left:3px solid {t['accent']};
    border-radius:8px;
    padding:12px 16px !important;
    box-shadow:0 2px 6px rgba(0,0,0,.06);
    transition:all .3s cubic-bezier(.25,.46,.45,.94);
}}
[data-testid="stMetric"]:hover {{
    transform:translateY(-2px);
    box-shadow:0 6px 20px rgba(0,0,0,.1);
}}
[data-testid="stMetricValue"] {{
    font-size:1.35rem !important; font-weight:700 !important; color:{t['text']} !important;
}}
[data-testid="stMetricLabel"] {{
    font-size:.72rem !important; font-weight:700 !important;
    text-transform:uppercase !important; letter-spacing:.5px !important;
    color:{t['subtext']} !important;
}}

/* ── KPI cards (custom HTML) ── */
.kpi-card {{
    background:{t['card_bg']};
    border:1px solid {t['card_border']};
    border-left:3px solid {t['accent']};
    border-radius:8px; padding:14px 18px;
    position:relative; overflow:hidden;
    box-shadow:0 2px 6px rgba(0,0,0,.06);
    transition:all .3s cubic-bezier(.25,.46,.45,.94);
}}
.kpi-card:hover {{ transform:translateY(-3px); box-shadow:0 6px 20px rgba(0,0,0,.1); }}
.kpi-card .kpi-label {{
    font-size:.72rem; color:{t['subtext']}; text-transform:uppercase;
    letter-spacing:.5px; margin-bottom:5px; font-weight:700;
}}
.kpi-card .kpi-value {{
    font-size:1.35rem; font-weight:700; color:{t['text']};
}}
.kpi-card .kpi-delta {{
    font-size:.72rem; color:{t['subtext']}; margin-top:3px;
}}
.kpi-card .kpi-stripe {{
    position:absolute; left:0; top:0; bottom:0; width:4px;
}}

/* ── Buttons ── */
[data-testid="baseButton-primary"] {{
    background:{t['active_bg']} !important;
    border:none !important; border-radius:6px !important;
    font-weight:600 !important;
    box-shadow:0 2px 6px rgba(68,157,68,.3) !important;
}}
[data-testid="baseButton-primary"]:hover {{
    filter:brightness(1.08) !important;
    box-shadow:0 4px 12px rgba(68,157,68,.4) !important;
}}
</style>
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

SUBITENS: dict[str, list[str]] = {
    "339030 – Material de Consumo": [
        "SI 01 – Combustíveis e lubrificantes automotivos",
        "SI 02 – Combustíveis e lubrificantes de aviação",
        "SI 03 – Combustíveis e lubrificantes p/ outras finalidades",
        "SI 04 – Gás e outros materiais engarrafados",
        "SI 05 – Explosivos e munições",
        "SI 06 – Alimentos para animais",
        "SI 07 – Gêneros de alimentação",
        "SI 08 – Animais para pesquisa e abate",
        "SI 09 – Material farmacológico",
        "SI 10 – Material odontológico",
        "SI 11 – Material químico",
        "SI 12 – Material de coudelaria / uso zootécnico",
        "SI 13 – Material de caça e pesca",
        "SI 14 – Material educativo e esportivo",
        "SI 15 – Material para festividades e homenagens",
        "SI 16 – Material de expediente",
        "SI 17 – Material de TIC (consumo)",
        "SI 18 – Material/medicamentos veterinários",
        "SI 19 – Material de acondicionamento e embalagem",
        "SI 20 – Material de cama, mesa e banho",
        "SI 21 – Material de copa e cozinha",
        "SI 22 – Material de limpeza e higienização",
        "SI 23 – Uniformes, tecidos e aviamentos",
        "SI 24 – Material p/ manutenção de bens imóveis",
        "SI 25 – Material p/ manutenção de bens móveis",
        "SI 26 – Material elétrico e eletrônico",
        "SI 27 – Material de manobra e patrulhamento",
        "SI 28 – Material de proteção e segurança",
        "SI 29 – Material de áudio, vídeo e foto",
        "SI 30 – Material para comunicações",
        "SI 31 – Sementes, mudas e insumos",
        "SI 32 – Suprimento de aviação",
        "SI 33 – Material para produção industrial",
        "SI 34 – Sobressalentes p/ embarcações",
        "SI 35 – Material laboratorial",
        "SI 36 – Material hospitalar",
        "SI 37 – Sobressalentes de armamento",
        "SI 38 – Suprimento de proteção ao voo",
        "SI 39 – Material p/ manutenção de veículos",
        "SI 40 – Material biológico",
        "SI 41 – Material gráfico",
        "SI 42 – Ferramentas",
        "SI 43 – Material para reabilitação profissional",
        "SI 44 – Material de sinalização visual",
        "SI 45 – Material técnico p/ seleção e treinamento",
        "SI 46 – Material bibliográfico",
        "SI 47 – Software (produto)",
        "SI 48 – Bens móveis não ativáveis",
        "SI 49 – Bilhetes de passagem",
        "SI 50 – Bandeiras, flâmulas e insígnias",
        "SI 51 – Discotecas e filmotecas",
        "SI 52 – Material sigiloso/reservado",
        "SI 53 – Material meteorológico",
        "SI 54 – Material p/ conservação de estradas",
        "SI 55 – Selos de controle fiscal",
        "SI 57 – Marcação de fauna silvestre",
        "SI 58 – Sobressalentes industriais",
        "SI 59 – Material para divulgação",
        "SI 89 – Material de consumo no exterior",
        "SI 91 – Variação cambial negativa",
        "SI 96 – Pagamento antecipado",
    ],
    "339033 – Passagens": [
        "SI 01 – Passagens no país",
        "SI 02 – Passagens no exterior",
    ],
    "339039 – Serviços de Terceiros (PJ)": [
        "SI 01 – Assinaturas e anuidades",
        "SI 02 – Condomínios",
        "SI 03 – Comissões e corretagens",
        "SI 05 – Serviços técnicos profissionais",
        "SI 10 – Locação de imóveis",
        "SI 11 – Locação de software",
        "SI 12 – Locação de máquinas e equipamentos",
        "SI 16 – Manutenção de imóveis",
        "SI 17 – Manutenção de máquinas",
        "SI 18 – Estacionamento",
        "SI 19 – Manutenção de veículos",
        "SI 22 – Eventos e congressos",
        "SI 23 – Festividades",
        "SI 25 – Taxa de administração",
        "SI 36 – Multas",
        "SI 37 – Juros e mora",
        "SI 40 – Alimentação do trabalhador",
        "SI 41 – Fornecimento de alimentação",
        "SI 43 – Energia elétrica",
        "SI 44 – Água e esgoto",
        "SI 45 – Gás",
        "SI 47 – Comunicação",
        "SI 50 – Serviços médico-hospitalares",
        "SI 58 – Telecomunicações",
        "SI 59 – Áudio, vídeo e foto",
        "SI 63 – Serviços gráficos",
        "SI 69 – Seguros",
        "SI 74 – Fretes",
        "SI 77 – Vigilância",
        "SI 78 – Limpeza e conservação",
        "SI 79 – Apoio administrativo",
        "SI 80 – Hospedagem",
        "SI 83 – Cópias e reprografia",
        "SI 86 – Patrocínios",
        "SI 90 – Publicidade legal",
        "SI 91 – Publicidade institucional",
        "SI 99 – Outros serviços",
    ],
    "339040 – Serviços de TIC": [
        "SI 01 – Locação de ativos de rede",
        "SI 02 – Locação de computadores",
        "SI 03 – Locação de servidores/storage",
        "SI 04 – Locação de impressoras",
        "SI 05 – Locação de telefonia",
        "SI 06 – Licença de software",
        "SI 07 – Manutenção de software",
        "SI 09 – Hospedagem (datacenter)",
        "SI 10 – Suporte a usuários",
        "SI 11 – Infraestrutura de TIC",
        "SI 12 – Manutenção de equipamentos",
        "SI 13 – Comunicação de dados",
        "SI 14 – Telefonia e dados",
        "SI 15 – Digitalização de documentos",
        "SI 16 – Outsourcing de impressão",
        "SI 17 – Computação em nuvem (IaaS)",
        "SI 18 – Computação em nuvem (PaaS)",
        "SI 19 – Computação em nuvem (SaaS)",
        "SI 20 – Treinamento em TIC",
        "SI 21 – Serviços técnicos TIC",
        "SI 22 – Instalação",
        "SI 23 – Certificados digitais",
    ],
    "449052 – Material Permanente": [
        "SI 02 – Aeronaves",
        "SI 04 – Equipamentos de medição",
        "SI 06 – Equipamentos de comunicação",
        "SI 08 – Equipamentos médicos",
        "SI 10 – Equipamentos esportivos",
        "SI 12 – Eletrodomésticos",
        "SI 14 – Armamentos",
        "SI 18 – Material bibliográfico",
        "SI 20 – Embarcações",
        "SI 22 – Equip. de manobra e patrulhamento",
        "SI 24 – Equip. de proteção e segurança",
        "SI 26 – Instrumentos musicais",
        "SI 28 – Máquinas industriais",
        "SI 30 – Equipamentos energéticos",
        "SI 32 – Equipamentos gráficos",
        "SI 33 – Equipamentos de áudio e vídeo",
        "SI 34 – Equipamentos diversos",
        "SI 35 – TIC (permanente)",
        "SI 36 – Equipamentos de escritório",
        "SI 37 – TIC de rede",
        "SI 38 – Ferramentas de oficina",
        "SI 39 – Equip. hidráulicos e elétricos",
        "SI 40 – Máquinas agrícolas e rodoviárias",
        "SI 41 – Computadores",
        "SI 42 – Mobiliário em geral",
        "SI 43 – Servidores e storage",
        "SI 44 – Obras de arte",
        "SI 45 – Impressoras",
        "SI 46 – Semoventes",
        "SI 47 – Telefonia",
        "SI 48 – Veículos diversos",
        "SI 50 – Veículos ferroviários",
        "SI 52 – Veículos automotores",
        "SI 53 – Carros de combate",
        "SI 54 – Equipamentos aeronáuticos",
        "SI 56 – Equip. de proteção ao voo",
        "SI 57 – Acessórios de veículos",
        "SI 58 – Equipamentos de mergulho",
        "SI 60 – Equipamentos marítimos",
        "SI 83 – Vigilância ambiental",
        "SI 99 – Outros permanentes",
    ],
}
SITUACOES_REQ = ["Pendente", "Enviada", "Aprovada", "Empenhada", "Anulado", "Liquidada", "Paga"]
TIPOS_RELATORIO = [
    "Resumo Geral", "NCs Detalhado", "Requisições Detalhado",
    "NCs por Operação", "NCs Próximas do Vencimento (30 dias)", "REQs Pendentes",
    "Extrato por NC", "REQs por Empresa", "Saldo por PI", "Saldo por ND",
]
NAV = [
    ("📊", "Dashboard",        "dashboard"),
    ("📋", "Notas de Crédito", "ncs"),
    ("📝", "Requisições",      "reqs"),
    ("📃", "Gerar REQ",        "gerar_req"),
    ("📥", "Importar Dados",   "importar"),
    ("📄", "Lançar Documento", "pdf"),
    ("🤖", "Assistente",       "assistente"),
    ("📈", "Relatórios",       "relatorios"),
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


@st.cache_data(ttl=1800, show_spinner=False)
def _pesquisar_pregao(uasg: str, num_pregao: str, compra_id_direto: str = "") -> tuple:
    """Cache da pesquisa — persiste mesmo com reset de sessão (30 min TTL)."""
    from pesquisa_compras import buscar_pregao
    try:
        r = buscar_pregao(uasg, num_pregao, compra_id_direto)
        return r.fornecedores, r.compra_id, r.log
    except Exception as e:
        return [], "", [f"Erro: {e}"]


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
        escuro = st.toggle("🌙 Modo escuro", value=st.session_state.get("tema_escuro", False))
        if escuro != st.session_state.get("tema_escuro", False):
            st.session_state["tema_escuro"] = escuro
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
            df_m["Recebido_fmt"]  = df_m["Recebido"].apply(fmt)
            df_m["Empenhado_fmt"] = df_m["Empenhado"].apply(fmt)
            fig = go.Figure()
            fig.add_bar(x=df_m["Mês"], y=df_m["Recebido"],  name="Recebido",  marker_color=t["bar_recv"],
                        customdata=df_m["Recebido_fmt"],
                        hovertemplate="<b>%{x}</b><br>Recebido: <b>%{customdata}</b><extra></extra>")
            fig.add_bar(x=df_m["Mês"], y=df_m["Empenhado"], name="Empenhado", marker_color=t["bar_emp"],
                        customdata=df_m["Empenhado_fmt"],
                        hovertemplate="<b>%{x}</b><br>Empenhado: <b>%{customdata}</b><extra></extra>")
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
                df_op["Saldo_fmt"] = df_op["Saldo"].apply(fmt)
                fig = px.bar(df_op, x="Saldo", y="OP", orientation="h",
                             color="Saldo", color_continuous_scale=["#1a4a2e", t["bar_emp"]],
                             custom_data=["Saldo_fmt"],
                             labels={"Saldo": "R$", "OP": ""})
                fig.update_coloraxes(showscale=False)
                fig.update_traces(hovertemplate="<b>%{y}</b><br>Saldo: <b>%{customdata[0]}</b><extra></extra>")
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
            df_o["Valor_fmt"] = df_o["Valor"].apply(fmt)
            fig = px.bar(df_o, x="ORGÃO", y="Valor", color="Tipo", barmode="group",
                         color_discrete_map={"Recebido": t["bar_recv"], "Empenhado": t["bar_emp"]},
                         custom_data=["Valor_fmt", "Tipo"],
                         labels={"Valor": "R$", "ORGÃO": ""})
            fig.update_traces(hovertemplate="<b>%{x}</b><br>%{customdata[1]}: <b>%{customdata[0]}</b><extra></extra>")
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
            df_s["Valor_fmt"] = df_s["Valor"].apply(fmt)
            fig = px.bar(df_s, x="Situação", y="Valor", color="Situação",
                         color_discrete_map=cores, custom_data=["Valor_fmt"],
                         labels={"Valor": "R$", "Situação": ""})
            fig.update_traces(hovertemplate="<b>%{x}</b><br>Total: <b>%{customdata[0]}</b><extra></extra>")
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
        df_vop["Recebido_fmt"]  = df_vop["Recebido"].apply(fmt)
        df_vop["Empenhado_fmt"] = df_vop["Empenhado"].apply(fmt)
        df_vop["Saldo_fmt"]     = df_vop["Saldo"].apply(fmt)
        fig.add_bar(x=df_vop["OP"], y=df_vop["Recebido"],  name="Recebido",  marker_color=t["bar_recv"],
                    customdata=df_vop["Recebido_fmt"],
                    hovertemplate="<b>%{x}</b><br>Recebido: <b>%{customdata}</b><extra></extra>")
        fig.add_bar(x=df_vop["OP"], y=df_vop["Empenhado"], name="Empenhado", marker_color="#fbbf24",
                    customdata=df_vop["Empenhado_fmt"],
                    hovertemplate="<b>%{x}</b><br>Empenhado: <b>%{customdata}</b><extra></extra>")
        fig.add_bar(x=df_vop["OP"], y=df_vop["Saldo"],     name="Saldo",     marker_color=t["bar_emp"],
                    customdata=df_vop["Saldo_fmt"],
                    hovertemplate="<b>%{x}</b><br>Saldo: <b>%{customdata}</b><extra></extra>")
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
def page_ncs(ncs, reqs=None):
    st.title("📋 Notas de Crédito")
    hoje = date.today()
    reqs = reqs or []

    # Calcula empenhado por NC a partir das requisições (fonte de verdade)
    from collections import defaultdict
    emp_por_nc: dict[str, float] = defaultdict(float)
    for r in reqs:
        nc_r  = r.get("NC", "")
        val_r = parse(r.get("VALOR", 0))
        sit_r = r.get("SITUAÇÃO", "")
        if sit_r == "Empenhada" and nc_r:
            emp_por_nc[nc_r] += val_r
        elif sit_r == "Anulado" and nc_r:
            emp_por_nc[nc_r] -= val_r
    # Garante não negativo
    emp_por_nc = {k: max(0.0, v) for k, v in emp_por_nc.items()}

    def _dias_prazo(nc):
        try:
            return (datetime.strptime(nc["PRAZO"], "%d/%m/%Y").date() - hoje).days
        except Exception:
            return None

    def _parse_pct(s):
        try:
            return float(str(s).replace("%", "").replace(",", ".").strip() or 0)
        except Exception:
            return 0.0

    def _badge(nc):
        if nc.get("SITU") != "EM TELA": return "🟢"
        d = _dias_prazo(nc)
        if d is None: return "🔵"
        if d < 0:     return "🔴"
        if d <= 7:    return "🟠"
        if d <= 30:   return "🟡"
        return "🔵"

    # ── KPI cards ─────────────────────────────────────────────────────
    em_tela  = [nc for nc in ncs if nc.get("SITU") == "EM TELA"]
    vencidas = sum(1 for nc in em_tela if (d := _dias_prazo(nc)) is not None and d < 0)
    venc7    = sum(1 for nc in em_tela if (d := _dias_prazo(nc)) is not None and 0 <= d <= 7)
    t_receb  = sum(parse(nc.get("RECEBIDO", 0)) for nc in ncs)
    t_saldo  = sum(parse(nc.get("SALDO NC",  0)) for nc in ncs)

    k1, k2, k3, k4, k5 = st.columns(5)
    k1.metric("📋 Total NCs",   len(ncs))
    k2.metric("💰 Recebido",    fmt(t_receb))
    k3.metric("🏦 Saldo",       fmt(t_saldo))
    k4.metric("📺 Em Tela",     len(em_tela))
    k5.metric("⚠️ Urgentes",    f"{venc7} vencendo",
              delta=f"{vencidas} vencidas" if vencidas else None,
              delta_color="inverse")

    st.divider()

    # ── Filtros rápidos ────────────────────────────────────────────────
    fc1, fc2 = st.columns([3, 2])
    busca = fc1.text_input("Busca", placeholder="🔍 NC, Órgão, Finalidade...",
                           key="nc_busca", label_visibility="collapsed")
    status_opt = fc2.radio("Status", ["Todos", "🟢 OK", "📺 EM TELA", "🔴 Vencidas", "🟠 Vence 7d"],
                           horizontal=True, key="nc_status", label_visibility="collapsed")

    with st.expander("🔧 Filtros avançados"):
        fa1, fa2, fa3, fa4 = st.columns(4)
        f_org = fa1.multiselect("Órgão",    sorted({nc.get("ORGÃO","") for nc in ncs if nc.get("ORGÃO")}))
        f_op  = fa2.multiselect("Operação", sorted({nc.get("OP","")    for nc in ncs if nc.get("OP")}))
        f_nd  = fa3.multiselect("ND",       sorted({nc.get("ND","")    for nc in ncs if nc.get("ND")}))
        f_pi  = fa4.multiselect("PI",       sorted({nc.get("PI","")    for nc in ncs if nc.get("PI")}))

    # ── Aplicar filtros ────────────────────────────────────────────────
    filtradas = []
    for nc in ncs:
        d = _dias_prazo(nc)
        em = nc.get("SITU") == "EM TELA"
        if status_opt == "🟢 OK"       and nc.get("SITU") != "OK":                          continue
        if status_opt == "📺 EM TELA"  and not em:                                           continue
        if status_opt == "🔴 Vencidas" and not (em and d is not None and d < 0):             continue
        if status_opt == "🟠 Vence 7d" and not (em and d is not None and 0 <= d <= 7):       continue
        if f_org and nc.get("ORGÃO") not in f_org: continue
        if f_op  and nc.get("OP")    not in f_op:  continue
        if f_nd  and nc.get("ND")    not in f_nd:  continue
        if f_pi  and nc.get("PI")    not in f_pi:  continue
        if busca:
            haystack = " ".join(str(v) for v in nc.values()).lower()
            if busca.lower() not in haystack: continue
        filtradas.append(nc)

    # ── Barra de ações ─────────────────────────────────────────────────
    ac1, ac2, ac3 = st.columns([1, 1, 4])
    if ac1.button("➕ Nova NC", type="primary", use_container_width=True):
        st.session_state["form_nc"] = not st.session_state.get("form_nc", False)
    ac2.markdown(
        f'<a href="{LINK_SHEETS}" target="_blank">'
        f'<button style="background:transparent;color:#64748b;border:1px solid #334155;'
        f'border-radius:6px;padding:8px 12px;font-size:.85rem;cursor:pointer;width:100%;">'
        f'📊 Planilha</button></a>', unsafe_allow_html=True)
    saldo_f = sum(parse(nc.get("SALDO NC", 0)) for nc in filtradas)
    receb_f = sum(parse(nc.get("RECEBIDO", 0)) for nc in filtradas)
    ac3.info(f"**{len(filtradas)}** de {len(ncs)} NCs · Recebido: **{fmt(receb_f)}** · Saldo: **{fmt(saldo_f)}**")

    if st.session_state.get("form_nc"):
        _form_nc()

    if not filtradas:
        st.info("Nenhuma NC encontrada com os filtros aplicados.")
        return

    # ── Tabela enriquecida ─────────────────────────────────────────────
    rows = []
    for nc in filtradas:
        d         = _dias_prazo(nc)
        nc_num    = nc.get("NC", "")
        recebido  = parse(nc.get("RECEBIDO", 0))
        empenhado = parse(nc.get("EMPENHADO", 0))
        saldo     = max(0.0, recebido - empenhado)
        pct_emp   = round(empenhado / recebido * 100, 1) if recebido else 0.0
        rows.append({
            "":           _badge(nc),
            "NC":         nc_num,
            "ORGÃO":      nc.get("ORGÃO", ""),
            "OP":         nc.get("OP", ""),
            "FINALIDADE": nc.get("FINALIDADE", ""),
            "DATA NC":    nc.get("DATA NC", ""),
            "PRAZO":      nc.get("PRAZO", ""),
            "RESTAM":     d if d is not None else 9999,
            "RECEBIDO":   fmt(recebido),
            "EMPENHADO":  fmt(empenhado),
            "SALDO":      fmt(saldo),
            "EMP %":      pct_emp,
            "SITUAÇÃO":   nc.get("SITUAÇÃO", ""),
        })

    df = pd.DataFrame(rows)
    edited_nc = st.data_editor(
        df,
        use_container_width=True,
        hide_index=True,
        column_config={
            "":           st.column_config.TextColumn("",           width=35,  disabled=True),
            "NC":         st.column_config.TextColumn("NC",         width=130, disabled=True),
            "ORGÃO":      st.column_config.TextColumn("Órgão",      width=110, disabled=True),
            "OP":         st.column_config.TextColumn("Operação",   width=80,  disabled=True),
            "FINALIDADE": st.column_config.TextColumn("Finalidade", width=200),
            "DATA NC":    st.column_config.TextColumn("Data NC",    width=90),
            "PRAZO":      st.column_config.TextColumn("Prazo",      width=90),
            "RESTAM":     st.column_config.NumberColumn("Restam",   width=75,  disabled=True),
            "RECEBIDO":   st.column_config.TextColumn("Recebido",   width=130, disabled=True),
            "EMPENHADO":  st.column_config.TextColumn("Empenhado",  width=130, disabled=True),
            "SALDO":      st.column_config.TextColumn("Saldo",      width=130, disabled=True),
            "EMP %":      st.column_config.ProgressColumn("Emp %",
                              format="%.1f%%", min_value=0, max_value=100, width=90),
            "SITUAÇÃO":   st.column_config.TextColumn("Situação",   width=120, disabled=True),
        },
        key="nc_editor",
    )

    # Detecta alterações em FINALIDADE, DATA NC ou PRAZO
    nc_changes = []
    for i, (orig, novo) in enumerate(zip(rows, edited_nc.to_dict("records"))):
        if (orig["FINALIDADE"] != novo["FINALIDADE"] or
            orig["DATA NC"]    != novo["DATA NC"] or
            orig["PRAZO"]      != novo["PRAZO"]):
            nc_changes.append({
                "nc":        orig["NC"],
                "finalidade": novo["FINALIDADE"],
                "data_nc":   novo["DATA NC"],
                "prazo":     novo["PRAZO"],
            })

    if nc_changes:
        st.info(f"✏️ {len(nc_changes)} NC(s) alterada(s).")
        if st.button("💾 Salvar alterações NCs", type="primary", key="btn_salvar_ncs"):
            try:
                from sheets_nc import atualizar_nc_campos
                for c in nc_changes:
                    atualizar_nc_campos(c["nc"], c["finalidade"], c["data_nc"], c["prazo"])
                carregar(forcar=True)
                st.success("✅ Alterações salvas!")
                st.rerun()
            except Exception as e:
                st.error(f"Erro: {e}")


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
def page_reqs(reqs_legado, ncs):
    """Página de Requisições — usa SSAC_REQS (nova aba limpa)."""
    import streamlit.components.v1 as components
    from reqs_crud import ler_reqs, atualizar_req, itens_da_req
    from sheets_nc import atualizar_nc_empenhado, recalcular_empenhados

    reqs = ler_reqs()

    st.title("📝 Requisições")

    def _badge(sit):
        return {"Pendente": "🟡", "Empenhada": "🟢", "Anulado": "🔴",
                "Enviada": "🔵", "Aprovada": "🟠"}.get(sit, "⚪")

    def _val(r): return parse(r.get("VALOR", "0"))

    # ── KPIs ──────────────────────────────────────────────────────────
    total_geral  = sum(_val(r) for r in reqs)
    n_pend       = sum(1 for r in reqs if r.get("SITUACAO") == "Pendente")
    n_emp        = sum(1 for r in reqs if r.get("SITUACAO") == "Empenhada")
    n_anul       = sum(1 for r in reqs if r.get("SITUACAO") == "Anulado")
    val_emp      = sum(_val(r) for r in reqs if r.get("SITUACAO") == "Empenhada")

    k1, k2, k3, k4, k5 = st.columns(5)
    k1.metric("📋 Total REQs",  len(reqs))
    k2.metric("💰 Valor Total", fmt(total_geral))
    k3.metric("⏳ Pendentes",   n_pend)
    k4.metric("🟢 Empenhadas",  n_emp, delta=fmt(val_emp), delta_color="off")
    k5.metric("🔴 Anuladas",    n_anul)

    st.divider()

    # ── Filtros rápidos ────────────────────────────────────────────────
    fc1, fc2 = st.columns([3, 2])
    busca = fc1.text_input("Busca", placeholder="🔍 REQ, NC, Empresa...",
                           key="req_busca", label_visibility="collapsed")
    status_opt = fc2.radio("Status", ["Todos", "🟡 Pendente", "🟢 Empenhada", "🔴 Anulado"],
                           horizontal=True, key="req_status", label_visibility="collapsed")

    with st.expander("🔧 Filtros avançados"):
        fa1, fa2, fa3 = st.columns(3)
        f_emp = fa1.multiselect("Empresa", sorted({r.get("EMPRESA","") for r in reqs if r.get("EMPRESA")}))
        f_nc  = fa2.multiselect("NC",      sorted({r.get("NC","")      for r in reqs if r.get("NC")}))
        f_pi  = fa3.multiselect("PI",      sorted({r.get("PI","")      for r in reqs if r.get("PI")}))

    filtradas = []
    for r in reqs:
        sit = r.get("SITUACAO", "")
        if status_opt == "🟡 Pendente"  and sit != "Pendente":  continue
        if status_opt == "🟢 Empenhada" and sit != "Empenhada": continue
        if status_opt == "🔴 Anulado"   and sit != "Anulado":   continue
        if f_emp and r.get("EMPRESA") not in f_emp: continue
        if f_nc  and r.get("NC")      not in f_nc:  continue
        if f_pi  and r.get("PI")      not in f_pi:  continue
        if busca:
            hay = " ".join(str(v) for v in r.values()).lower()
            if busca.lower() not in hay: continue
        filtradas.append(r)

    # ── Ações ──────────────────────────────────────────────────────────
    ac1, ac2, ac3 = st.columns([1, 1, 4])
    if ac1.button("➕ Nova REQ", type="primary", use_container_width=True):
        st.session_state["form_req"] = not st.session_state.get("form_req", False)
    if ac2.button("🔄 Recalcular NCs", use_container_width=True):
        with st.spinner("Recalculando..."):
            try:
                resultado = recalcular_empenhados(reqs)
                carregar(forcar=True)
                msgs = [f"**{k}**: {v}" for k, v in resultado.items()]
                st.success("✅ " + " | ".join(msgs) if msgs else "Sem REQs empenhadas.")
            except Exception as e:
                st.error(f"Erro: {e}")
    total_f = sum(_val(r) for r in filtradas)
    ac3.info(f"**{len(filtradas)}** de {len(reqs)} REQs · **{fmt(total_f)}**")

    if st.session_state.get("form_req"):
        _form_req_novo(ncs)

    if not filtradas:
        st.info("Nenhuma REQ encontrada.")
        return

    # ── Tabela ────────────────────────────────────────────────────────
    rows = []
    for r in filtradas:
        rows.append({
            "":           _badge(r.get("SITUACAO", "")),
            "REQ":        r.get("REQ", ""),
            "DATA":       r.get("DATA", ""),
            "NC":         r.get("NC", ""),
            "NE":         r.get("NE", ""),
            "PI":         r.get("PI", ""),
            "EMPRESA":    r.get("EMPRESA", ""),
            "VALOR":      fmt(_val(r)),
            "SITUACAO":   r.get("SITUACAO", ""),
            "ENTRADA":    r.get("ENTRADA_SALC", ""),
            "OBS":        r.get("OBS", ""),
        })

    # ── Lista com ações por linha ──────────────────────────────────────
    from reqs_crud import excluir_req as _excluir

    for ridx, (row, r_orig) in enumerate(zip(rows, filtradas)):
        ca, cb, cc, cd, ce, cf = st.columns([1, 2, 5, 1, 1, 1])
        ca.markdown(row[""])
        cb.markdown(f"**{row['REQ']}** {row['DATA']}")
        cc.markdown(f"{row['EMPRESA']} · {row['NC']} · {row['VALOR']} · _{row['SITUACAO']}_")
        if cd.button("✏️", key=f"edit_{ridx}"):
            st.session_state["edit_req"] = row["REQ"]
            st.rerun()
        if ce.button("🗑️", key=f"del_{ridx}"):
            _excluir(row["REQ"])
            st.session_state.pop("edit_req", None)
            st.rerun()
        if cf.button("📧", key=f"email_{ridx}"):
            st.session_state["email_req"] = row["REQ"]
            st.rerun()

    # ── Painel de edição / email (fora do loop) ───────────────────────
    req_acao = st.session_state.get("edit_req") or st.session_state.get("email_req")
    if req_acao:
        r_edit = next((r for r in filtradas if r.get("REQ") == req_acao), None)
        if not r_edit:
            st.session_state.pop("edit_req", None)
            st.session_state.pop("email_req", None)
        elif st.session_state.get("edit_req"):
            st.divider()
            st.subheader(f"✏️ Editar REQ {req_acao}")
            from reqs_crud import editar_req as _editar
            # Guarda valores editáveis no session state para persistir
            _pfx = f"_ed_{req_acao}_"
            for k, default in [("EMPRESA", r_edit.get("EMPRESA","")),
                                ("NE",      r_edit.get("NE","")),
                                ("CNPJ",    r_edit.get("CNPJ","")),
                                ("OBS",     r_edit.get("OBS","")),
                                ("ENTRADA_SALC", r_edit.get("ENTRADA_SALC","")),
                                ("SITUACAO",r_edit.get("SITUACAO","Pendente"))]:
                st.session_state.setdefault(_pfx + k, default)

            ec1, ec2 = st.columns(2)
            emp  = ec1.text_input("Empresa",      key=_pfx+"EMPRESA")
            ne_e = ec1.text_input("NE",           key=_pfx+"NE")
            cnpj = ec1.text_input("CNPJ",         key=_pfx+"CNPJ")
            obs  = ec1.text_input("Obs",          key=_pfx+"OBS")
            sit_ops = ["Pendente","Empenhada","Anulado"]
            sit_e   = ec2.selectbox("Situação", sit_ops,
                                    index=sit_ops.index(st.session_state[_pfx+"SITUACAO"])
                                    if st.session_state[_pfx+"SITUACAO"] in sit_ops else 0,
                                    key=_pfx+"SITUACAO")
            ent_e   = ec2.text_input("Entrada SALC", key=_pfx+"ENTRADA_SALC")
            try:
                val_def = parse(r_edit.get("VALOR","0"))
            except Exception:
                val_def = 0.0
            val_e = ec2.number_input("Valor (R$)", value=val_def,
                                     min_value=0.0, step=0.01, format="%.2f",
                                     key=_pfx+"VALOR")

            sb1, sb2 = st.columns(2)
            if sb1.button("💾 Salvar edição", type="primary", use_container_width=True, key="btn_sv_edit"):
                try:
                    _editar(req_acao, {**r_edit,
                                       "EMPRESA": emp, "NE": ne_e, "CNPJ": cnpj,
                                       "OBS": obs, "SITUACAO": sit_e,
                                       "ENTRADA_SALC": ent_e, "VALOR": val_e})
                    # Limpa prefixo
                    for k in list(st.session_state.keys()):
                        if k.startswith(_pfx):
                            del st.session_state[k]
                    st.session_state.pop("edit_req", None)
                    st.success("✅ REQ atualizada!")
                    st.rerun()
                except Exception as e:
                    st.error(f"Erro ao salvar: {e}")
            if sb2.button("✖ Cancelar", use_container_width=True, key="btn_cancel_edit"):
                st.session_state.pop("edit_req", None)
                st.rerun()

        elif st.session_state.get("email_req"):
            st.divider()
            st.subheader(f"📧 Enviar REQ {req_acao} por email")
            dest = st.text_input("Email do destinatário", key="email_dest_inp")
            se1, se2 = st.columns(2)
            if se1.button("📧 Enviar", type="primary", use_container_width=True, key="btn_send_email"):
                if dest:
                    try:
                        from gerador_req_html import gerar_html_req
                        itens_e = itens_da_req(req_acao)
                        campos_e = {"requisition_id": r_edit.get("REQ",""),
                                    "LOCAL_DATA": r_edit.get("DATA",""), "UG": UG_PADRAO,
                                    "OM": OM_PADRAO,
                                    "FORNECEDOR_NOME": r_edit.get("EMPRESA",""),
                                    "FORNECEDOR_CNPJ": r_edit.get("CNPJ",""),
                                    "MODALIDADE": r_edit.get("PREGAO",""),
                                    "DADOS_NC": r_edit.get("NC",""), "NE": r_edit.get("NE",""),
                                    "PI": r_edit.get("PI",""), "ND": r_edit.get("ND",""),
                                    "TIPO": r_edit.get("TIPO","Ordinário"),
                                    "TOTAL": fmt(_val(r_edit)),
                                    "ASSUNTO":"","INTRO_1":"","JUSTIFICATIVA":"",
                                    "FINALIDADE":"","PTRES":""}
                        html_e = gerar_html_req(campos_e, itens_e)
                        _enviar_email(dest, f"REQ {req_acao}", html_e)
                        st.success(f"✅ Email enviado para {dest}!")
                        st.session_state.pop("email_req", None)
                        st.rerun()
                    except Exception as e:
                        st.error(f"Erro: {e}")
            if se2.button("✖ Cancelar", use_container_width=True, key="btn_cancel_email"):
                st.session_state.pop("email_req", None)
                st.rerun()

    st.divider()
    st.caption("Tabela completa (edite Situação, Entrada SALC e NE diretamente):")
    df_r   = pd.DataFrame(rows)
    edited = st.data_editor(
        df_r, use_container_width=True, hide_index=True,
        column_config={
            "":        st.column_config.TextColumn("",        width=35,  disabled=True),
            "REQ":     st.column_config.TextColumn("REQ",     width=65,  disabled=True),
            "DATA":    st.column_config.TextColumn("Data",    width=90,  disabled=True),
            "NC":      st.column_config.TextColumn("NC",      width=130, disabled=True),
            "NE":      st.column_config.TextColumn("NE",      width=90),
            "PI":      st.column_config.TextColumn("PI",      width=70,  disabled=True),
            "EMPRESA": st.column_config.TextColumn("Empresa", width=170, disabled=True),
            "VALOR":   st.column_config.TextColumn("Valor",   width=120, disabled=True),
            "SITUACAO":st.column_config.SelectboxColumn("Situação", width=110,
                           options=["Pendente", "Empenhada", "Anulado"]),
            "ENTRADA": st.column_config.TextColumn("Entrada SALC", width=110),
            "OBS":     st.column_config.TextColumn("Obs",     width=120, disabled=True),
        },
        key="req_editor_novo",
    )

    # Detecta e salva alterações
    changed = []
    for i, (orig, novo) in enumerate(zip(rows, edited.to_dict("records"))):
        if orig["SITUACAO"] != novo["SITUACAO"] or orig["ENTRADA"] != novo["ENTRADA"] or orig["NE"] != novo["NE"]:
            req_orig = filtradas[i] if i < len(filtradas) else {}
            changed.append({"req": orig["REQ"], "sit": novo["SITUACAO"],
                             "sit_ant": orig["SITUACAO"], "entrada": novo["ENTRADA"],
                             "ne": novo["NE"], "nc": req_orig.get("NC",""),
                             "valor": _val(req_orig)})

    if changed:
        st.info(f"✏️ {len(changed)} linha(s) alterada(s).")
        if st.button("💾 Salvar alterações", type="primary", key="btn_salvar_reqs"):
            try:
                for c in changed:
                    atualizar_req(c["req"], c["sit"], c["entrada"], c["ne"])
                    nc, val = c["nc"], c["valor"]
                    if nc and val:
                        if c["sit"] == "Empenhada" and c["sit_ant"] != "Empenhada":
                            try: atualizar_nc_empenhado(nc, val)
                            except Exception as e: st.warning(f"NC não atualizada: {e}")
                        elif c["sit"] == "Anulado" and c["sit_ant"] != "Anulado":
                            try: atualizar_nc_empenhado(nc, -val)
                            except Exception as e: st.warning(f"NC não atualizada: {e}")
                carregar(forcar=True)
                st.success("✅ Salvo!")
                st.rerun()
            except Exception as e:
                st.error(f"Erro: {e}")

    # ── Consulta de saldo por NC ──────────────────────────────────────
    st.divider()
    with st.expander("🔎 Consultar saldo de uma NC", expanded=False):
        ncs_nc = sorted({r.get("NC","") for r in reqs if r.get("NC")})
        nc_q   = st.selectbox("NC", [""] + ncs_nc, key="req_nc_q", label_visibility="collapsed")
        if nc_q:
            nc_d    = next((n for n in ncs if n.get("NC") == nc_q), {})
            receb   = parse(nc_d.get("RECEBIDO", 0))
            reqs_nc = [r for r in reqs if r.get("NC") == nc_q and r.get("SITUACAO") != "Anulado"]
            t_reqs  = sum(_val(r) for r in reqs_nc)
            saldo   = max(0.0, receb - t_reqs)
            c1,c2,c3,c4 = st.columns(4)
            c1.metric("💰 Recebido", fmt(receb))
            c2.metric("📋 Total REQs", fmt(t_reqs), delta=f"{len(reqs_nc)} req(s)", delta_color="off")
            c3.metric("🟢 Saldo", fmt(saldo), delta="OK" if saldo >= 0 else "⚠️",
                      delta_color="normal" if saldo >= 0 else "inverse")
            c4.metric("⏳ Pendentes", fmt(sum(_val(r) for r in reqs_nc if r.get("SITUACAO")=="Pendente")))
            if receb > 0:
                st.progress(min(1.0, t_reqs/receb), text=f"{t_reqs/receb*100:.1f}% comprometido")

    # ── Visualizador de HTML ──────────────────────────────────────────
    st.divider()
    st.subheader("📄 Visualizar REQ")
    reqs_com_html = [r.get("REQ","") for r in reqs if r.get("REQ")]
    if not reqs_com_html:
        st.caption("Nenhuma REQ cadastrada ainda.")
    else:
        req_vis = st.selectbox("REQ para visualizar", reqs_com_html, key="vis_req_sel")
        if st.button("📄 Gerar visualização", type="primary",
                     use_container_width=True, key="btn_vis_req"):
            r_data = next((r for r in reqs if r.get("REQ") == req_vis), {})
            try:
                from gerador_req_html import gerar_html_req
                campos_v = {
                    "requisition_id": r_data.get("REQ",""),
                    "LOCAL_DATA": r_data.get("DATA",""),
                    "UG": UG_PADRAO, "OM": OM_PADRAO,
                    "FORNECEDOR_NOME": r_data.get("EMPRESA",""),
                    "FORNECEDOR_CNPJ": r_data.get("CNPJ",""),
                    "MODALIDADE": r_data.get("PREGAO",""),
                    "DADOS_NC": r_data.get("NC",""),
                    "NE": r_data.get("NE",""), "PI": r_data.get("PI",""),
                    "ND": r_data.get("ND",""), "TIPO": r_data.get("TIPO","Ordinário"),
                    "TOTAL": fmt(_val(r_data)),
                    "ASSUNTO":"","INTRO_1":"","JUSTIFICATIVA":"","FINALIDADE":"","PTRES":"",
                }
                html_g = gerar_html_req(campos_v, itens_da_req(req_vis))
                st.session_state["_vis_html"]  = html_g
                st.session_state["_vis_bytes"] = html_g.encode("utf-8")
                st.session_state["_vis_req"]   = req_vis
            except Exception as e:
                st.error(f"Erro ao gerar: {e}")

        if st.session_state.get("_vis_req") == req_vis and st.session_state.get("_vis_html"):
            st.download_button("⬇️ Baixar HTML", data=st.session_state["_vis_bytes"],
                               file_name=f"REQ_{req_vis}.html", mime="text/html",
                               use_container_width=True)
            components.html(st.session_state["_vis_html"], height=860, scrolling=True)


def _form_editar_req(ncs, reqs):
    from reqs_crud import editar_req
    req_num = st.session_state.get("edit_req_num", "")
    r_data  = next((r for r in reqs if r.get("REQ") == req_num), {})
    if not r_data:
        st.session_state.pop("edit_req_num", None)
        return

    st.divider()
    st.subheader(f"✏️ Editar REQ {req_num}")

    nums_nc = [""] + [nc.get("NC","") for nc in ncs if nc.get("NC")]
    nc_idx  = nums_nc.index(r_data.get("NC","")) if r_data.get("NC","") in nums_nc else 0
    nc_sel  = st.selectbox("NC Vinculada", nums_nc, index=nc_idx, key="_edit_nc")
    nc_d    = next((nc for nc in ncs if nc.get("NC") == nc_sel), {}) if nc_sel else {}

    with st.form("f_editar_req", clear_on_submit=False):
        c1, c2 = st.columns(2)
        empresa = c1.text_input("Empresa", value=r_data.get("EMPRESA",""))
        cnpj    = c1.text_input("CNPJ",    value=r_data.get("CNPJ",""))
        ne      = c1.text_input("NE",      value=r_data.get("NE",""))
        obs     = c1.text_input("Obs",     value=r_data.get("OBS",""))

        tipo    = c2.selectbox("Tipo", ["Ordinário","Especial","Anulação"],
                               index=["Ordinário","Especial","Anulação"].index(r_data.get("TIPO","Ordinário"))
                               if r_data.get("TIPO","Ordinário") in ["Ordinário","Especial","Anulação"] else 0)
        sit_ops = ["Pendente","Empenhada","Anulado"]
        sit_idx = sit_ops.index(r_data.get("SITUACAO","Pendente")) if r_data.get("SITUACAO","Pendente") in sit_ops else 0
        situacao = c2.selectbox("Situação", sit_ops, index=sit_idx)
        entrada  = c2.text_input("Entrada SALC", value=r_data.get("ENTRADA_SALC",""))

        try:
            valor_atual = float(str(r_data.get("VALOR","0")).replace("R$","").replace(".","").replace(",",".").strip() or 0)
        except Exception:
            valor_atual = 0.0
        valor = st.number_input("Valor (R$)", value=valor_atual, min_value=0.0, step=0.01, format="%.2f")

        s1, s2 = st.columns(2)
        salvar   = s1.form_submit_button("💾 Salvar", type="primary", use_container_width=True)
        cancelar = s2.form_submit_button("✖ Cancelar", use_container_width=True)

        if cancelar:
            st.session_state.pop("edit_req_num", None)
            st.rerun()
        if salvar:
            try:
                editar_req(req_num, {
                    "DATA":         r_data.get("DATA",""),
                    "NC":           nc_sel,
                    "NE":           ne,
                    "PI":           nc_d.get("PI", r_data.get("PI","")),
                    "ND":           nc_d.get("ND", r_data.get("ND","")),
                    "EMPRESA":      empresa,
                    "CNPJ":         cnpj,
                    "PREGAO":       r_data.get("PREGAO",""),
                    "TIPO":         tipo,
                    "VALOR":        valor,
                    "SITUACAO":     situacao,
                    "ENTRADA_SALC": entrada,
                    "OBS":          obs,
                })
                st.success(f"✅ REQ {req_num} atualizada!")
                st.session_state.pop("edit_req_num", None)
                st.rerun()
            except Exception as e:
                st.error(f"Erro: {e}")


def _form_req(ncs):
    _form_req_novo(ncs)


def _form_req_novo(ncs):
    from reqs_crud import adicionar_req as adicionar_req_novo
    st.divider()

    tipo_req    = st.radio("Tipo", ["📋 Empenho", "🔴 Anulação"], horizontal=True,
                           key="_frq_tipo2", label_visibility="collapsed")
    eh_anulacao = "Anulação" in tipo_req
    if eh_anulacao:
        st.error("⚠️ **Anulação** — saldo da NC será devolvido.")

    nums_nc = [""] + [nc.get("NC","") for nc in ncs if nc.get("NC")]
    nc_sel  = st.selectbox("NC Vinculada", nums_nc, key="_frq_nc2")
    nc_d    = next((nc for nc in ncs if nc.get("NC") == nc_sel), {}) if nc_sel else {}
    if nc_d:
        st.info(f"📋 **{nc_sel}** · PI: {nc_d.get('PI','')} · ND: {nc_d.get('ND','')}")

    with st.form("f_req_novo", clear_on_submit=True):
        c1, c2 = st.columns(2)
        req_num  = c1.text_input("Nº REQ")
        data_req = c1.date_input("Data", value=date.today())
        empresa  = c1.text_input("Empresa *")
        cnpj     = c1.text_input("CNPJ")

        ne      = c2.text_input("NE")
        tipo    = c2.selectbox("Tipo", ["Anulação"] if eh_anulacao else ["Ordinário","Especial"])
        sit_def = "Anulado" if eh_anulacao else "Pendente"
        sit_idx = 0
        situacao = c2.selectbox("Situação", ["Pendente","Empenhada","Anulado"], index=["Pendente","Empenhada","Anulado"].index(sit_def))
        entrada  = c2.date_input("Entrada SALC", value=None)
        obs      = c2.text_input("Obs")

        valor  = st.number_input("Valor (R$) *", min_value=0.0, step=0.01, format="%.2f")
        desc   = st.text_area("Descrição *")

        s1, s2 = st.columns(2)
        salvar   = s1.form_submit_button("💾 Salvar", type="primary", use_container_width=True)
        cancelar = s2.form_submit_button("✖ Cancelar", use_container_width=True)

        if cancelar:
            st.session_state["form_req"] = False
            st.rerun()

        if salvar:
            if not empresa or not desc or valor <= 0:
                st.error("Preencha Empresa, Descrição e Valor.")
            else:
                try:
                    adicionar_req_novo({
                        "REQ":          req_num,
                        "DATA":         data_req.strftime("%d/%m/%Y"),
                        "NC":           nc_sel,
                        "NE":           ne,
                        "PI":           nc_d.get("PI",""),
                        "ND":           nc_d.get("ND",""),
                        "EMPRESA":      empresa,
                        "CNPJ":         cnpj,
                        "PREGAO":       "",
                        "TIPO":         tipo,
                        "VALOR":        valor,
                        "SITUACAO":     situacao,
                        "ENTRADA_SALC": entrada.strftime("%d/%m/%Y") if entrada else "",
                        "OBS":          obs,
                        "ITENS":        [],
                    })
                    st.success("✅ REQ cadastrada!")
                    st.session_state["form_req"] = False
                    st.rerun()
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
        tipo     = c2.selectbox("Tipo", ["Ordinário", "Especial", "Anulação"])
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
    col_i, col_j, col_f = st.columns(3)
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

    with col_f:
        st.caption("📝 Assunto")
        frases_assunto = _frases("FINAL")
        sel_assunto = st.selectbox("Frase padrão", frases_assunto, key="sel_assunto_frase",
                                    label_visibility="collapsed")
        if sel_assunto != "— escrever manualmente —":
            st.session_state["_assunto_txt"] = sel_assunto
        elif "_assunto_txt" not in st.session_state:
            st.session_state["_assunto_txt"] = ""

        with st.expander("➕ Cadastrar nova frase de Assunto"):
            nova_assunto = st.text_area("Texto", key="nova_assunto_txt", height=80)
            if st.button("💾 Salvar", key="btn_salvar_assunto"):
                if nova_assunto.strip():
                    from sheets_nc import adicionar_frase
                    adicionar_frase("FINAL", nova_assunto.strip())
                    st.cache_data.clear()
                    st.success("Frase salva!")
                    st.rerun()

        if sel_assunto != "— escrever manualmente —":
            with st.expander("🗑️ Excluir esta frase"):
                if st.button("Confirmar exclusão", key="btn_del_assunto"):
                    from sheets_nc import excluir_frase
                    excluir_frase("FINAL", sel_assunto)
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

        assunto      = st.text_input("Assunto", value=st.session_state.get("_assunto_txt") or _v("FINALIDADE", "FINALIDADE"))
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

    # ── Pesquisa no portal (contratos.comprasnet.gov.br)
    st.divider()
    st.subheader("🔍 Pesquisar Pregão — contratos.comprasnet.gov.br")

    st.info(
        "💡 **Como encontrar o ID da compra:** acesse "
        "[contratos.comprasnet.gov.br/empenho/buscacompra](https://contratos.comprasnet.gov.br/empenho/buscacompra), "
        "pesquise o pregão e copie o número da URL: `.../fornecedor/**7194897**`"
    )
    p1, p2 = st.columns([1, 2])
    p_uasg   = p1.text_input("UASG", value=UG_PADRAO, key="pesq_uasg")
    p_pregao = p2.text_input("Nº Pregão/Ano", placeholder="ex: 90009/2025", key="pesq_pregao")
    p_cid    = st.text_input("🔑 ID da Compra (da URL do portal)",
                              placeholder="ex: 7194897 — encontrado na URL após pesquisar no portal",
                              key="pesq_cid")

    col_pesq, col_limpar = st.columns([4, 1])
    btn_pesquisar = col_pesq.button("🔍 Pesquisar", key="btn_pesquisar", use_container_width=True)
    if col_limpar.button("🗑️ Limpar", key="btn_limpar_cache", use_container_width=True):
        _pesquisar_pregao.clear()
        for k in ("pesq_fornecedores", "pesq_compra_id", "pesq_log", "pesq_itens_cache"):
            st.session_state.pop(k, None)
        st.rerun()

    if btn_pesquisar:
        if not p_pregao and not p_cid:
            st.warning("Digite o número do pregão ou o ID da compra.")
        else:
            with st.spinner("🌐 Buscando no portal..."):
                fornecedores, compra_id, log = _pesquisar_pregao(p_uasg, p_pregao, p_cid)
            st.session_state["pesq_fornecedores"] = fornecedores
            st.session_state["pesq_compra_id"]    = compra_id
            st.session_state["pesq_log"]          = log
            st.session_state["pesq_itens_cache"]  = {}
            if not fornecedores:
                erro = next((l for l in log if "Erro:" in str(l)), None)
                st.error(erro or "Nenhum fornecedor encontrado.")
                with st.expander("🔎 Log"):
                    for linha in log: st.text(linha)

    # ── Exibe fornecedores e itens ─────────────────────────────────────────────
    fornecedores_pesq = st.session_state.get("pesq_fornecedores", [])
    compra_id_pesq    = st.session_state.get("pesq_compra_id", "")

    if fornecedores_pesq:
        from pesquisa_compras import buscar_itens_fornecedor
        forn_atual           = st.session_state.get("_b2_forn", "")
        itens_ja_adicionados = {it["ITEM"] for it in st.session_state.req_itens}

        # Mostra ND/SI que será aplicado ao clicar ➕
        _nd_atual = st.session_state.get("_fi_nd", "— sem ND —")
        _si_atual = st.session_state.get("_fi_si", "— sem SI —")
        _si_num   = _num_si(_si_atual)
        if _nd_atual != "— sem ND —":
            st.caption(f"📋 Ao clicar ➕: **ND** = {_nd_atual.split('–')[0].strip()} | **SI** = {_si_num or '—'} *(ajuste abaixo se necessário)*")
        itens_cache          = st.session_state.get("pesq_itens_cache", {})
        pregao_atual         = st.session_state.get("pesq_pregao", "")

        st.success(f"✅ {len(fornecedores_pesq)} fornecedor(es) | Compra ID: {compra_id_pesq}")
        if forn_atual:
            st.info(f"🔒 Vinculado a **{forn_atual}**")

        for fidx, forn in enumerate(fornecedores_pesq):
            forn_id   = forn["fornecedor_id"]
            forn_nome = forn["nome"]
            cnpj_grp  = forn["cnpj"]
            label_exp = f"🏢 **{forn_nome}** — {cnpj_grp}"

            with st.expander(label_exp, expanded=False):
                if forn_id not in itens_cache:
                    if st.button("📦 Carregar itens", key=f"load_{forn_id}"):
                        with st.spinner("Carregando..."):
                            itens = buscar_itens_fornecedor(compra_id_pesq, forn_id)
                        itens_cache[forn_id] = itens
                        st.session_state["pesq_itens_cache"] = itens_cache
                        st.rerun()
                else:
                    itens = itens_cache[forn_id]
                    st.caption(f"{len(itens)} item(ns)")
                    for gidx, item in enumerate(itens):
                        num   = item.get("numero", "")
                        ja    = num in itens_ja_adicionados
                        saldo = item.get("qtd_saldo", "")
                        v     = item.get("valor_unit", 0.0)
                        ca, cb, cc = st.columns([1, 8, 2])
                        ca.markdown(f"**{num}**")
                        cb.markdown(
                            f"{item.get('descricao','')[:80]}"
                            f"&nbsp;&nbsp;<span style='color:#449D44;font-weight:700'>{fmt(v)}</span>"
                            f"&nbsp;&nbsp;<span style='color:#888;font-size:.85em'>Saldo: {saldo}</span>",
                            unsafe_allow_html=True,
                        )
                        if ja:
                            cc.markdown("✓")
                        elif cc.button("➕", key=f"add_{forn_id}_{gidx}"):
                            if forn_atual and forn_nome != forn_atual:
                                st.error(f"❌ Req vinculada a **{forn_atual}**.")
                            else:
                                _si = _num_si(st.session_state.get("_fi_si", ""))
                                st.session_state.req_itens.append({
                                    "ORD":            str(len(st.session_state.req_itens) + 1),
                                    "ITEM":           str(num),
                                    "SI":             _si,
                                    "DESCRICAO_ITEM": item.get("descricao", ""),
                                    "UND":            "UN",
                                    "QTD":            "1,000",
                                    "VALOR_UNIT":     fmt(v),
                                    "VALOR_TOTAL":    fmt(v),
                                    "_total":         v,
                                    "_vunit":         v,
                                })
                                st.session_state["_b2_forn"]   = forn_nome
                                st.session_state["_b2_cnpj"]   = cnpj_grp
                                st.session_state["_b2_pregao"] = pregao_atual
                                st.session_state["_b2_ug"]     = str(st.session_state.get("pesq_uasg", UG_PADRAO))
                                st.session_state["_b2_vig"]    = ""
                                st.session_state["_b2_ver"]    = st.session_state.get("_b2_ver", 0) + 1

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

    # Chave versionada garante que value= é respeitado após auto-fill
    _b2v = st.session_state.get("_b2_ver", 0)
    b2_c1, b2_c2 = st.columns(2)
    b2_forn_edit = b2_c1.text_input("Fornecedor (Razão Social)", value=_b2_forn, key=f"b2_forn_{_b2v}")
    b2_cnpj_edit = b2_c2.text_input("CNPJ",                     value=_b2_cnpj, key=f"b2_cnpj_{_b2v}")

    b2_m1, b2_m2, b2_m3, b2_m4 = st.columns([2, 2, 1, 2])
    b2_modal_edit  = b2_m1.selectbox("Modalidade",
                     ["PREGÃO", "CARONA", "DISPENSA", "INEXIGIBILIDADE", "SUPRIMENTO DE FUNDOS"],
                     key=f"b2_modal_{_b2v}")
    b2_pregao_edit = b2_m2.text_input("Nº Pregão/ARP",   value=_b2_pregao, key=f"b2_pregao_{_b2v}")
    b2_ug_edit     = b2_m3.text_input("UG",              value=_b2_ug,     key=f"b2_ug_{_b2v}")
    b2_vig_edit    = b2_m4.text_input("Vigência da ATA", value=_b2_vig,    key=f"b2_vig_{_b2v}")

    # ── 3. Itens
    st.divider()
    st.subheader("3. Itens da Requisição")

    # Inicializa chaves do formulário de item se não existirem
    for k, v in [("_fi_item",""), ("_fi_desc",""), ("_fi_und","UN"), ("_fi_vunit", 0.0)]:
        st.session_state.setdefault(k, v)

    # ── Seletor ND / SI (fora do form para cascata) ───────────────────
    # Pré-seleciona ND a partir da NC selecionada (ex: NC com ND=339030)
    nc_nd_raw = nc_d.get("ND", "").strip()
    if nc_nd_raw:
        nd_match = next((k for k in SUBITENS if k.startswith(nc_nd_raw)), None)
        nd_prev_key = f"_fi_nd_from_nc_{nc_sel}"
        if nd_match and st.session_state.get("_fi_nd_prev_nc") != nd_prev_key:
            st.session_state["_fi_nd"] = nd_match
            st.session_state["_fi_nd_prev_nc"] = nd_prev_key
            # Auto-seleciona o primeiro SI disponível para o ND
            primeiros_si = SUBITENS.get(nd_match, [])
            if primeiros_si and st.session_state.get("_fi_si", "— sem SI —") == "— sem SI —":
                st.session_state["_fi_si"] = primeiros_si[0]

    nd_keys = ["— sem ND —"] + list(SUBITENS.keys())
    snd1, snd2 = st.columns(2)
    nd_sel = snd1.selectbox("ND (Natureza de Despesa)", nd_keys, key="_fi_nd")
    si_opts = ["— sem SI —"] + (SUBITENS.get(nd_sel, []) if nd_sel != "— sem ND —" else [])
    si_sel  = snd2.selectbox("Sub-item (SI)", si_opts, key="_fi_si")

    with st.form("f_add_item", clear_on_submit=True):
        ci1, ci2 = st.columns([1, 5])
        item  = ci1.text_input("Item *",  key="_fi_item", help="Nº do item no pregão")
        desc  = ci2.text_input("Descrição do Item *", key="_fi_desc")
        ci3, ci4, ci5 = st.columns([1, 1, 2])
        und   = ci3.text_input("Unid.", key="_fi_und")
        qtd   = ci4.number_input("Qtd", min_value=0.001, value=1.0, step=1.0, format="%.3f")
        vunit = ci5.number_input("Valor Unit. (R$)", key="_fi_vunit",
                                 min_value=0.0, step=0.01, format="%.2f")
        add   = st.form_submit_button("➕ Adicionar Item", use_container_width=True)

    if add:
        if item and desc and vunit > 0:
            si_val  = _num_si(si_sel)
            total   = round(qtd * vunit, 2)
            ord_num = len(st.session_state.req_itens) + 1
            st.session_state.req_itens.append({
                "ORD":            str(ord_num),
                "ITEM":           item,
                "SI":             si_val,
                "DESCRICAO_ITEM": desc,
                "UND":            und,
                "QTD":            str(qtd).replace(".", ","),
                "VALOR_UNIT":     fmt(vunit),
                "VALOR_TOTAL":    fmt(total),
                "_total":         total,
                "_vunit":         vunit,
            })
            st.rerun()
        else:
            st.warning("Preencha o Nº do item, a descrição e o valor unitário.")

    if st.session_state.req_itens:
        df_it = pd.DataFrame(st.session_state.req_itens)
        df_it.insert(0, "🗑️", False)
        edited = st.data_editor(
            df_it[["🗑️", "ORD", "ITEM", "SI", "DESCRICAO_ITEM", "UND", "QTD", "VALOR_UNIT", "VALOR_TOTAL"]],
            use_container_width=True, hide_index=True,
            column_config={
                "🗑️":           st.column_config.CheckboxColumn("", width=35),
                "ORD":           st.column_config.TextColumn("Ord", disabled=True, width="small"),
                "ITEM":          st.column_config.TextColumn("Item", disabled=True, width="small"),
                "SI":            st.column_config.TextColumn("SI", width="small"),
                "DESCRICAO_ITEM":st.column_config.TextColumn("Descrição", disabled=True),
                "UND":           st.column_config.TextColumn("Und", disabled=True, width="small"),
                "QTD":           st.column_config.TextColumn("Qtd"),
                "VALOR_UNIT":    st.column_config.TextColumn("Valor Unit.", disabled=True),
                "VALOR_TOTAL":   st.column_config.TextColumn("Valor Total", disabled=True),
            },
            key="editor_itens",
        )

        # Apagar selecionados
        selecionados = [i for i, row in edited.iterrows() if row.get("🗑️")]
        if selecionados:
            if st.button(f"🗑️ Apagar {len(selecionados)} item(ns) selecionado(s)",
                         type="primary", key="btn_apagar_sel"):
                st.session_state.req_itens = [
                    it for i, it in enumerate(st.session_state.req_itens)
                    if i not in selecionados
                ]
                # Renumera ORD
                for i, it in enumerate(st.session_state.req_itens):
                    it["ORD"] = str(i + 1)
                st.rerun()
        # Aplica edições de SI e QTD — detecta mudança e recalcula total
        changed = False
        for i, row in edited.iterrows():
            if i >= len(st.session_state.req_itens):
                break
            it = st.session_state.req_itens[i]

            # SI
            si_novo = "" if pd.isna(row.get("SI")) else str(row.get("SI", "") or "")
            if si_novo != str(it.get("SI", "") or ""):
                it["SI"] = si_novo
                changed = True

            # QTD — recalcula total
            qtd_raw = row.get("QTD")
            if pd.isna(qtd_raw):
                continue
            qtd_str = str(qtd_raw).strip()
            qtd_num = parse(qtd_str)
            if qtd_num <= 0:
                continue
            vunit      = float(it.get("_vunit") or 0)
            novo_total = round(qtd_num * vunit, 2)
            if abs(novo_total - float(it.get("_total", 0))) > 0.001:
                it["QTD"]         = qtd_str
                it["_total"]      = novo_total
                it["VALOR_TOTAL"] = fmt(novo_total)
                changed = True

        if changed:
            st.rerun()

        total_geral = sum(i["_total"] for i in st.session_state.req_itens)
        k1, k2 = st.columns([1, 3])
        k1.metric("💰 Total Geral", fmt(total_geral))
        with k2:
            if st.button("🗑️ Limpar todos os itens"):
                st.session_state.req_itens = []
                for k in ("_b2_forn", "_b2_cnpj", "_b2_pregao", "_b2_ug", "_b2_vig"):
                    st.session_state.pop(k, None)
                st.session_state["_b2_ver"] = st.session_state.get("_b2_ver", 0) + 1
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

                _b2v      = st.session_state.get("_b2_ver", 0)
                b2_modal  = st.session_state.get(f"b2_modal_{_b2v}", "PREGÃO")
                b2_pregao = st.session_state.get("_b2_pregao") or st.session_state.get(f"b2_pregao_{_b2v}", "")
                b2_ug     = st.session_state.get("_b2_ug") or st.session_state.get(f"b2_ug_{_b2v}", "")
                b2_forn   = st.session_state.get("_b2_forn") or st.session_state.get(f"b2_forn_{_b2v}", "")
                b2_cnpj   = st.session_state.get("_b2_cnpj") or st.session_state.get(f"b2_cnpj_{_b2v}", "")
                b2_vig    = st.session_state.get("_b2_vig") or st.session_state.get(f"b2_vig_{_b2v}", "")
                modalidade  = f"{b2_modal} - {b2_pregao} {b2_ug}".strip(" -")
                total_geral = sum(i["_total"] for i in itens)
                campos_finais = {
                    **campos,
                    "FORNECEDOR_NOME": b2_forn,
                    "FORNECEDOR_CNPJ": _formatar_cnpj(b2_cnpj),
                    "MODALIDADE":      modalidade,
                    "VIGENCIA_DA_ATA": b2_vig,
                    "TOTAL":           fmt(total_geral),
                }
                itens_limpos = [{k: v for k, v in i.items() if not k.startswith("_")} for i in itens]
                doc_bytes    = gerar_para_bytes(TEMPLATE_PATH, campos_finais, itens_limpos)
                nome_arq     = f"REQ_{campos.get('requisition_id', 'doc')}.docx"

                st.session_state["_doc_bytes"]        = doc_bytes
                st.session_state["_doc_nome"]         = nome_arq
                st.session_state["_doc_campos"]       = campos_finais
                st.session_state["_doc_itens_limpos"] = itens_limpos
                st.session_state["_doc_total"]        = total_geral
                st.session_state["_doc_req_cadastrada"] = False
            except Exception as e:
                st.error(f"Erro ao gerar documento: {e}")

    # ── Downloads ─────────────────────────────────────────────────────
    if st.session_state.get("_doc_bytes"):
        nome_arq  = st.session_state.get("_doc_nome", "REQ.docx")
        doc_bytes = st.session_state["_doc_bytes"]

        st.download_button(
            "⬇️ Baixar DOCX",
            data=doc_bytes,
            file_name=nome_arq,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
        )

        # HTML para download e visualização na aba Requisições
        if not st.session_state.get("_doc_html_bytes"):
            try:
                from gerador_req_html import gerar_html_req
                html = gerar_html_req(st.session_state.get("_doc_campos", {}),
                                      st.session_state.get("_doc_itens_limpos", []))
                st.session_state["_doc_html_bytes"] = html.encode("utf-8")
            except Exception:
                pass
        if st.session_state.get("_doc_html_bytes"):
            nome_html = nome_arq.replace(".docx", ".html")
            st.download_button("⬇️ Baixar HTML", data=st.session_state["_doc_html_bytes"],
                               file_name=nome_html, mime="text/html",
                               use_container_width=True)

        if st.session_state.get("_doc_req_cadastrada"):
            st.success("✅ Requisição cadastrada!")
        else:
            st.info("📋 Deseja cadastrar esta requisição na planilha?")
            cad1, cad2 = st.columns(2)
            if cad1.button("✅ Sim, cadastrar", type="primary", use_container_width=True, key="btn_cad_sim"):
                try:
                    from reqs_crud import adicionar_req as adicionar_req_novo
                    cf        = st.session_state["_doc_campos"]
                    itens_lim = st.session_state.get("_doc_itens_limpos", [])
                    req_num   = cf.get("requisition_id", "")
                    adicionar_req_novo({
                        "REQ":     req_num,
                        "DATA":    cf.get("LOCAL_DATA", date.today().strftime("%d/%m/%Y")),
                        "NC":      st.session_state.get("gerar_nc_sel", ""),
                        "NE":      cf.get("NE", ""),
                        "PI":      cf.get("PI", ""),
                        "ND":      cf.get("ND", ""),
                        "EMPRESA": cf.get("FORNECEDOR_NOME", ""),
                        "CNPJ":    cf.get("FORNECEDOR_CNPJ", ""),
                        "PREGAO":  cf.get("MODALIDADE", ""),
                        "TIPO":    cf.get("TIPO", "Ordinário"),
                        "VALOR":   st.session_state.get("_doc_total", 0.0),
                        "SITUACAO": "Pendente",
                        "ITENS":   itens_lim,
                    })
                    st.session_state["_doc_req_cadastrada"] = True
                    st.rerun()
                except Exception as e:
                    st.error(f"Erro ao cadastrar: {e}")
            if cad2.button("✖ Não", use_container_width=True, key="btn_cad_nao"):
                st.session_state["_doc_req_cadastrada"] = True
                st.rerun()


def _num_si(si_str: str) -> str:
    """Extrai só o número de 'SI 16 – Material de expediente' → '16'."""
    import re
    if not si_str or si_str == "— sem SI —":
        return ""
    m = re.search(r"SI\s+(\d+)", si_str)
    return m.group(1) if m else si_str


def _enviar_email(destinatario: str, assunto: str, html_body: str) -> None:
    """Envia email com o HTML da REQ como corpo."""
    import smtplib
    from email.mime.multipart import MIMEMultipart
    from email.mime.text import MIMEText
    from config import EMAIL_SENDER, EMAIL_PASSWORD

    if not EMAIL_SENDER or not EMAIL_PASSWORD:
        raise ValueError(
            "Configure EMAIL_SENDER e EMAIL_PASSWORD no Railway.\n"
            "Use uma senha de app do Gmail: "
            "myaccount.google.com/security → Senhas de app"
        )

    msg = MIMEMultipart("alternative")
    msg["Subject"] = assunto
    msg["From"]    = EMAIL_SENDER
    msg["To"]      = destinatario
    msg.attach(MIMEText(html_body, "html", "utf-8"))

    with smtplib.SMTP("smtp.gmail.com", 587) as smtp:
        smtp.starttls()
        smtp.login(EMAIL_SENDER, EMAIL_PASSWORD)
        smtp.sendmail(EMAIL_SENDER, destinatario, msg.as_string())


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
    elif pagina == "ncs":        page_ncs(ncs, reqs)
    elif pagina == "reqs":       page_reqs(reqs, ncs)
    elif pagina == "pdf":        page_pdf(ncs)
    elif pagina == "gerar_req":  page_gerar_req(reqs, ncs)
    elif pagina == "importar":   page_importar(ncs, reqs)
    elif pagina == "relatorios": page_relatorios(ncs, reqs)
    elif pagina == "assistente": page_assistente(ncs, reqs)


if __name__ == "__main__":
    main()
