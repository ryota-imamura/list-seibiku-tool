import streamlit as st

st.set_page_config(
    page_title="リスト整備ツール",
    page_icon="📋",
    layout="centered",
)

# ── 認証（全ページ共通） ────────────────────────────────────────────────
if not st.session_state.get("authenticated", False):
    st.markdown("""
    <style>
        .main-title { font-size: 2rem; font-weight: 700; color: #1F4E79; margin-bottom: 0.2rem; }
        .sub-title { font-size: 0.95rem; color: #555; margin-bottom: 1.5rem; }
    </style>
    <div class="main-title">📋 リスト整備ツール</div>
    <div class="sub-title">認証が必要です</div>
    """, unsafe_allow_html=True)
    with st.form(key="auth_form"):
        password = st.text_input("パスワードを入力", type="password")
        submit = st.form_submit_button("ログイン", use_container_width=True, type="primary")
        if submit:
            if password == "seibi0000":
                st.session_state.authenticated = True
                st.rerun()
            else:
                st.error("❌ パスワードが間違っています")
    st.stop()

# ── ナビゲーション ──────────────────────────────────────────────────────
pg = st.navigation([
    st.Page("pages/1_リスト整備ツール.py",    title="リスト整備ツール",    icon="📋"),
    st.Page("pages/2_リバブル重複チェック.py", title="リバブル重複チェック", icon="🔍"),
])
pg.run()
