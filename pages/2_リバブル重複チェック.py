"""
リバブル重複チェック - Streamlit ページ
"""

import io
import streamlit as st
import pandas as pd
from livable_check import check_duplicates

# ── スタイル ──────────────────────────────────────────────────────────
st.markdown("""
<style>
    .main-title { font-size: 2rem; font-weight: 700; color: #1F4E79; margin-bottom: 0.2rem; }
    .sub-title { font-size: 0.95rem; color: #555; margin-bottom: 1.5rem; }
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

st.markdown('<div class="main-title">🔍 リバブル重複チェック</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-title">作業リスト × リバブル一括リストを突合し、過去の送付履歴を付記します</div>', unsafe_allow_html=True)

# ── 説明 ──────────────────────────────────────────────────────────────
with st.expander("📌 チェック内容を確認する"):
    st.markdown("""
<div class="rule-box">
<b>【マッチングロジック】</b><br>
✅ マッチキー：（オーナー住所の先頭5文字 ＋ オーナー名）の先頭10文字<br>
✅ リバブル一括リストの「列3」と同一のキーで照合<br>
✅ 1件のオーナーが複数のリストにヒットした場合も全て表示<br><br>
<b>【出力Excelの構成】</b><br>
📄 <b>重複チェック済みリスト</b>：元の作業リスト全行 ＋ 重複フラグ・センター名・リスト名・発送日を追記（重複行はオレンジ背景）<br>
📄 <b>重複詳細</b>：1マッチ1行で展開（複数センターにヒットした場合も全表示）<br>
📄 <b>センター別集計</b>：センターごとの重複件数<br>
</div>
""", unsafe_allow_html=True)

st.divider()

# ── アップロード ──────────────────────────────────────────────────────
col1, col2 = st.columns(2)

with col1:
    st.subheader("① 作業リスト")
    work_file = st.file_uploader(
        "整備対象の作業リスト (.xlsx)",
        type=["xlsx"],
        key="work",
        label_visibility="collapsed",
    )
    if work_file:
        xl = pd.ExcelFile(io.BytesIO(work_file.read()))
        work_file.seek(0)
        work_sheet = st.selectbox("シートを選択", xl.sheet_names)
    else:
        work_sheet = 0

with col2:
    st.subheader("② リバブル一括リスト")
    livable_file = st.file_uploader(
        "提供されたリバブル一括リスト (.xlsx)",
        type=["xlsx"],
        key="livable",
        label_visibility="collapsed",
    )

st.divider()

# ── 実行 ──────────────────────────────────────────────────────────────
if work_file and livable_file:
    if st.button("🚀 重複チェックを実行", type="primary", use_container_width=True):
        progress_bar = st.progress(0.0)
        status_area = st.empty()

        def on_progress(msg, pct):
            if pct is not None:
                progress_bar.progress(float(pct))
            status_area.info("⏳ " + msg)

        try:
            work_bytes = work_file.read()
            livable_bytes = livable_file.read()

            excel_bytes, summary = check_duplicates(
                work_bytes,
                livable_bytes,
                work_sheet=work_sheet,
                progress_callback=on_progress,
            )

            progress_bar.progress(1.0)
            status_area.empty()
            st.success("✅ チェックが完了しました！")
            st.divider()

            # サマリー
            st.subheader("③ 処理結果サマリー")
            c1, c2, c3 = st.columns(3)
            c1.metric("総行数", f"{summary['total']:,} 件")
            c2.metric("重複あり", f"{summary['dup_count']:,} 件",
                      delta=f"{summary['dup_count']/summary['total']*100:.1f}%",
                      delta_color="off")
            c3.metric("重複なし", f"{summary['clean_count']:,} 件")

            # センター別集計
            if summary["center_stats"]:
                st.divider()
                st.subheader("センター別 重複件数")
                center_df = pd.DataFrame(
                    sorted(summary["center_stats"].items(), key=lambda x: -x[1]),
                    columns=["センター名", "重複件数"],
                )
                st.dataframe(center_df, use_container_width=True, hide_index=True)

            # ダウンロード
            st.divider()
            st.subheader("④ 結果をダウンロード")
            from datetime import datetime
            now = datetime.now().strftime("%Y%m%d_%H%M")
            st.download_button(
                label="📥 Excelをダウンロード（重複チェック済みリスト・重複詳細・センター別集計）",
                data=excel_bytes,
                file_name=f"リバブル重複チェック結果_{now}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True,
            )

        except Exception as e:
            st.error(f"エラーが発生しました: {e}")
            raise
else:
    st.info("👆 ① 作業リストと ② リバブル一括リストの両方をアップロードしてください")
