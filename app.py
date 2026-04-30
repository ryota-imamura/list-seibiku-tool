import streamlit as st
import pandas as pd
from datetime import datetime
from list_processor import process

st.set_page_config(
    page_title="リスト整備ツール",
    page_icon="📋",
    layout="centered",
)

# ── 認証 ────────────────────────────────────────────────────────────────
if "authenticated" not in st.session_state:
    st.session_state.authenticated = False

if not st.session_state.authenticated:
    st.markdown("""
    <style>
        .main-title { font-size: 2rem; font-weight: 700; color: #1F4E79; margin-bottom: 0.2rem; }
        .sub-title { font-size: 0.95rem; color: #555; margin-bottom: 1.5rem; }
    </style>
    <div class="main-title">📋 リスト整備ツール</div>
    <div class="sub-title">認証が必要です</div>
    """, unsafe_allow_html=True)

    with st.form(key='auth_form'):
        password = st.text_input("パスワードを入力", type="password")
        submit = st.form_submit_button("ログイン", use_container_width=True, type="primary")
        if submit:
            if password == "seibi0000":
                st.session_state.authenticated = True
                st.rerun()
            else:
                st.error("❌ パスワードが間違っています")
    st.stop()

# ── スタイル ──────────────────────────────────────────────────────────
st.markdown("""
<style>
    .main-title {
        font-size: 2rem;
        font-weight: 700;
        color: #1F4E79;
        margin-bottom: 0.2rem;
    }
    .sub-title {
        font-size: 0.95rem;
        color: #555;
        margin-bottom: 1.5rem;
    }
    .rule-box {
        background: #f8f9fa;
        border-left: 4px solid #1F4E79;
        padding: 0.8rem 1rem;
        border-radius: 0 8px 8px 0;
        font-size: 0.88rem;
        color: #333;
        line-height: 1.8;
    }
</style>
""", unsafe_allow_html=True)

# ── ヘッダー ──────────────────────────────────────────────────────────
st.markdown('<div class="main-title">📋 リスト整備ツール</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-title">DM送付用オーナーリストを自動整備します</div>', unsafe_allow_html=True)

# ── 整備ルール説明 ──────────────────────────────────────────────────
with st.expander("📌 整備ルールを確認する"):
    st.markdown("""
<div class="rule-box">
<b>【自動整備の内容】</b><br>
✅ 全角→半角変換・前後スペース除去<br>
✅ 郵便番号の正規化（xxx-xxxx形式に統一）<br>
✅ 都道府県が抜けている場合、郵便番号または市区町村名から補完<br>
✅ オーナー住所が空欄の場合、郵便番号から補完（物件住所との混同は禁止）<br>
✅ 郵便番号が不正な場合、オーナー住所から逆引き補完<br>
✅ 重複行の削除（オーナー名＋住所＋郵便番号が一致）<br>
✅ 同一住所の名義人を連名としてまとめる（表記揺れも対応）<br>
✅ 文字化けデータの検出<br><br>
<b>【エラーとして除外する条件】</b><br>
❌ オーナー名が未入力<br>
❌ オーナー住所が未入力（補完不可）<br>
❌ 郵便番号が未入力または不正（逆引き不可）<br>
❌ 文字化けの疑いがある<br><br>
<b>【処理時間の目安】</b><br>
郵便番号の逆引き補完が多い場合は1〜2分かかることがあります。
</div>
""", unsafe_allow_html=True)

st.divider()

# ── ファイルアップロード ──────────────────────────────────────────────
st.subheader("① Excelファイルをアップロード")
uploaded = st.file_uploader(
    "クライアントから受け取ったリストをそのまま貼り付けてください",
    type=["xlsx"],
    label_visibility="collapsed",
)

if uploaded:
    df_preview = pd.read_excel(uploaded, header=0, nrows=5)
    st.caption(f"ファイル名: `{uploaded.name}`　/ プレビュー（先頭5行）")
    st.dataframe(df_preview, use_container_width=True)

    st.divider()
    st.subheader("② 整備を実行")

    if st.button("🚀 整備スタート", type="primary", use_container_width=True):
        uploaded.seek(0)
        file_bytes = uploaded.read()

        status_area = st.empty()
        progress_bar = st.progress(0)

        def on_progress(msg, progress=None):
            status_area.info("⏳ " + msg)
            if progress is not None:
                progress_bar.progress(min(progress, 1.0))

        try:
            excel_bytes, summary, error_list = process(file_bytes, on_progress)
        except Exception as e:
            st.error(f"処理中にエラーが発生しました: {e}")
            st.stop()

        progress_bar.progress(1.0)
        status_area.empty()
        progress_bar.empty()
        st.success("✅ 整備が完了しました！")
        st.divider()

        # ── サマリー ──
        st.subheader("③ 処理結果サマリー")

        col1, col2 = st.columns(2)
        with col1:
            st.metric("📬 発注可能件数（整備済み）", f"{summary['発注可能件数']} 件",
                      help="エラーなしの整備済みリスト件数")
        with col2:
            st.metric("⚠️ エラー件数", f"{summary['エラー件数']} 件",
                      delta=f"-{summary['エラー件数']}" if summary['エラー件数'] else None,
                      delta_color="inverse",
                      help="住所不明・名前なし等でDM送付不可の件数")

        c1, c2, c3, c4, c5 = st.columns(5)
        c1.metric("🗑️ 重複削除",    f"{summary['重複削除件数']} 件")
        c2.metric("🏠 住所補完",    f"{summary['住所補完件数']} 件")
        c3.metric("📮 郵便番号補完", f"{summary['郵便番号補完件数']} 件")
        c4.metric("👥 連名統合",    f"{summary['連名統合件数']} 件")
        c5.metric("🔤 文字化け",    f"{summary['文字化け検出件数']} 件")

        # ── エラー一覧 ──
        if error_list:
            st.divider()
            st.subheader("⚠️ エラー行一覧")
            df_err = pd.DataFrame(error_list).rename(columns={
                '行番号': '元行番号', 'オーナー名': '名前', 'エラー理由': 'エラー理由'
            })
            st.dataframe(df_err, use_container_width=True, hide_index=True)

        # ── ダウンロード ──
        st.divider()
        st.subheader("④ 整備済みファイルをダウンロード")
        now = datetime.now().strftime("%Y%m%d_%H%M")
        dl_name = f"整備済みリスト_{now}.xlsx"
        st.download_button(
            label="📥 Excelをダウンロード（整備済み・エラー・ログの3シート）",
            data=excel_bytes,
            file_name=dl_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True,
        )
        st.caption(f"出力ファイル: `{dl_name}`　（整備済みリスト / エラーリスト / 整備ログ の3シート構成）")

else:
    st.info("👆 Excelファイル（.xlsx）をアップロードしてください")

# ── フッター ──────────────────────────────────────────────────────────
st.divider()
st.markdown(
    '<div style="text-align:center; color:#aaa; font-size:0.8rem;">リスト整備ツール (β版)</div>',
    unsafe_allow_html=True
)
