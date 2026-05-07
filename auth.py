"""
Autenticação por senha. Hash gerado com: python -c "from auth import gerar_hash; print(gerar_hash('sua_senha'))"
"""
import hashlib
import os
import streamlit as st


def _hash(pwd: str) -> str:
    return hashlib.sha256(pwd.encode("utf-8")).hexdigest()


def _hash_armazenado() -> str:
    try:
        h = st.secrets.get("AUTH_PASSWORD_HASH", "")
        if h:
            return str(h)
    except Exception:
        pass
    return os.getenv("AUTH_PASSWORD_HASH", "")


def verificar(senha: str) -> bool:
    h = _hash_armazenado()
    return (not h) or (_hash(senha) == h)


def gerar_hash(senha: str) -> str:
    """Gera o hash para colocar no .env ou secrets do Streamlit Cloud."""
    return _hash(senha)


def requer_auth():
    """Bloqueia o app se não autenticado. Chame no início de main()."""
    if st.session_state.get("_auth"):
        return
    _pagina_login()
    st.stop()


def logout():
    st.session_state["_auth"] = False
    st.rerun()


def _pagina_login():
    from config import OM_PADRAO, UG_PADRAO

    st.markdown("""
    <style>
    #MainMenu, header, footer {visibility:hidden;}
    .block-container {padding-top:0 !important;}
    .login-wrap {
        max-width: 360px; margin: 80px auto 0;
        background: #161b22; border: 1px solid #30363d;
        border-radius: 16px; padding: 40px 36px;
        box-shadow: 0 8px 32px rgba(0,0,0,0.4);
    }
    .login-title { text-align:center; margin-bottom:28px; }
    .login-title .icon { font-size:2.8rem; }
    .login-title h1 { color:#e6edf3; font-size:1.5rem; margin:8px 0 4px; }
    .login-title p  { color:#8b949e; font-size:0.85rem; margin:0; }
    </style>
    """, unsafe_allow_html=True)

    _, col, _ = st.columns([1, 2, 1])
    with col:
        st.markdown(f"""
        <div class="login-wrap">
            <div class="login-title">
                <div class="icon">🪖</div>
                <h1>SSAC</h1>
                <p>{OM_PADRAO} &nbsp;·&nbsp; UG {UG_PADRAO}</p>
            </div>
        </div>
        """, unsafe_allow_html=True)

        with st.form("login", clear_on_submit=True):
            senha = st.text_input("Senha", type="password",
                                  label_visibility="collapsed",
                                  placeholder="Senha de acesso")
            entrar = st.form_submit_button("Entrar →", use_container_width=True, type="primary")

        if entrar:
            if verificar(senha):
                st.session_state["_auth"] = True
                st.rerun()
            else:
                st.error("Senha incorreta.")
