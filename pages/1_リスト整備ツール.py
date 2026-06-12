import streamlit as st
import pandas as pd
import subprocess
import os
from datetime import datetime
from list_processor import process

@st.cache_data(ttl=300)
def _get_version():
    try:
        cwd = os.path.dirname(os.path.abspath(__file__))
        sha = subprocess.check_output(
            ['git', 'log', '-1', '--format=%h'], cwd=cwd,
            stderr=subprocess.DEVNULL, timeout=2).decode().strip()
        dt = subprocess.check_output(
            ['git', 'log', '-1', '--format=%cd', '--date=format:%Y-%m-%d %H:%M'],
            cwd=cwd, stderr=subprocess.DEVNULL, timeout=2).decode().strip()
        return f"{sha} ({dt})"
    except Exception:
        return "unknown"

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

st.markdown('<div class="main-title">📋 リスト整備ツール</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-title">DM送付用オーナーリストを自動整備します</div>', unsafe_allow_html=True)

with st.expander("📌 整備ルールを確認する"):
    st.markdown("""
<div class="rule-box">
<b>【自動整備の内容】</b><br>
✅ 全角→半角変換・前後スペース除去<br>
✅ ダッシュ類（−・―・–など）の統一。住所・地番は出力時に全角「－」で印字<br>
✅ 番地表記の統一（1丁目2番3号 → 1－2－3、漢数字丁目も対応。「一番町」等の町名や号室・号棟は保持）<br>
✅ 郵便番号の正規化（xxx-xxxx形式に統一）<br>
✅ 都道府県が抜けている場合、郵便番号または市区町村名から補完<br>
✅ オーナー住所が空欄の場合、郵便番号から補完（物件住所との混同は禁止）<br>
✅ 郵便番号が不正な場合、オーナー住所から逆引き補完<br>
✅ 郵便番号が住所と別地域の場合、住所から逆引きして自動修正<br>
✅ 重複行の削除（オーナー名＋住所＋郵便番号が一致。異体字「髙/高」等やスペース揺れも同一視）<br>
✅ 同一住所の名義人を連名としてまとめる（表記揺れも対応）<br>
✅ 文字化けデータの検出<br><br>
<b>【エラーとして除外する条件】</b><br>
❌ オーナー名が未入力<br>
❌ オーナー住所が未入力（補完不可）<br>
❌ 郵便番号が未入力または不正（逆引き不可）<br>
❌ 文字化けの疑いがある<br><br>
<b>【要確認（エラーにはせず黄色表示）】</b><br>
⚠️ 住所に番地がない（不着リスク）<br>
⚠️ 郵便番号と住所の地域が不一致で自動修正もできない<br>
→ 提供リスト対応表シートで「採用（要確認）」として表示されます<br><br>
<b>【出力ファイル】</b><br>
整備済みリスト／エラーリスト／提供リスト対応表／整備ログ の4シート構成<br><br>
<b>【処理時間の目安】</b><br>
郵便番号の逆引き補完が多い場合は1〜2分かかることがあります。
</div>
""", unsafe_allow_html=True)

st.divider()

st.subheader("① Excelファイルをアップロード")
uploaded = st.file_uploader(
    "クライアントから受け取ったリストをそのまま貼り付けてください",
    type=["xlsx"],
    label_visibility="collapsed",
)

if uploaded:
    xl = pd.ExcelFile(uploaded)
    all_sheets = xl.sheet_names
    uploaded.seek(0)

    if len(all_sheets) > 1:
        valid_sheets = []
        for sheet in all_sheets:
            try:
                df_tmp = pd.read_excel(uploaded, sheet_name=sheet, header=0, nrows=10)
                uploaded.seek(0)
                if 'オーナー名' in df_tmp.columns:
                    real = df_tmp[~df_tmp.iloc[:, 0].astype(str).str.startswith('例')]
                    if real['オーナー名'].notna().any():
                        valid_sheets.append(sheet)
            except Exception:
                uploaded.seek(0)
        if not valid_sheets:
            valid_sheets = all_sheets
        selected_sheet = st.selectbox("📄 処理するシートを選択してください", valid_sheets,
                                      help="データが入っているシートのみ表示しています")
        uploaded.seek(0)
    else:
        selected_sheet = all_sheets[0]

    df_preview = pd.read_excel(uploaded, sheet_name=selected_sheet, header=0, nrows=5)
    uploaded.seek(0)
    st.caption(f"ファイル名: `{uploaded.name}`　シート: `{selected_sheet}`　/ プレビュー（先頭5行）")
    st.dataframe(df_preview, use_container_width=True)

    # ── 空ヘッダー列の確認 ──────────────────────────────────────────────
    from list_processor import detect_columns
    df_full = pd.read_excel(uploaded, sheet_name=selected_sheet, header=0)
    uploaded.seek(0)
    cmap = detect_columns(df_full)
    col_to_field = {c: f for f, c in cmap.items() if c is not None}
    field_options = ['使わない', '地番', '物件名', '物件住所', 'オーナー名', 'オーナー住所',
                     '郵便番号', '連名①', '連名②', '連名③', '連名④', '連名⑤', '備考']
    unnamed_cols = [c for c in df_full.columns
                    if str(c).startswith('Unnamed') and df_full[c].notna().any()]
    manual_map = {}
    if unnamed_cols:
        st.divider()
        st.subheader("⚠️ 列名（ヘッダー）が空の列の確認")
        st.caption("ヘッダーが空の列が見つかりました。それぞれどの項目に入れるか確認してください。"
                   "（自動推測がある場合は初期選択済みです）")
        for c in unnamed_cols:
            samples = [str(v) for v in df_full[c].dropna().head(3).tolist()]
            default_field = col_to_field.get(c, '使わない')
            default_idx = field_options.index(default_field) if default_field in field_options else 0
            sel = st.selectbox(
                f"空ヘッダー列（例: {', '.join(samples) or '（データ例なし）'}）の割り当て",
                field_options, index=default_idx, key=f"map_{selected_sheet}_{c}",
            )
            manual_map[c] = sel

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
            excel_bytes, summary, error_list = process(file_bytes, on_progress, sheet_name=selected_sheet, manual_map=manual_map)
        except Exception as e:
            st.error(f"処理中にエラーが発生しました: {e}")
            st.stop()

        progress_bar.progress(1.0)
        status_area.empty()
        progress_bar.empty()
        st.success("✅ 整備が完了しました！")
        st.divider()

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

        c1, c2, c3, c4, c5, c6 = st.columns(6)
        c1.metric("🗑️ 重複削除",    f"{summary['重複削除件数']} 件")
        c2.metric("🏠 住所補完",    f"{summary['住所補完件数']} 件")
        c3.metric("📮 郵便番号補完", f"{summary['郵便番号補完件数'] + summary.get('郵便番号修正件数', 0)} 件")
        c4.metric("👥 連名統合",    f"{summary['連名統合件数']} 件")
        c5.metric("🔤 文字化け",    f"{summary['文字化け検出件数']} 件")
        c6.metric("⚠️ 要確認",     f"{summary.get('要確認件数', 0)} 件",
                  help="番地なし・郵便番号と住所の地域不一致など、発送前に目視確認を推奨する行（提供リスト対応表シートで黄色表示）")

        if error_list:
            st.divider()
            st.subheader("⚠️ エラー行一覧")
            df_err = pd.DataFrame(error_list).rename(columns={
                '行番号': '元行番号', 'オーナー名': '名前', 'エラー理由': 'エラー理由'
            })
            st.dataframe(df_err, use_container_width=True, hide_index=True)

        st.divider()
        st.subheader("④ 整備済みファイルをダウンロード")
        now = datetime.now().strftime("%Y%m%d_%H%M")
        dl_name = f"整備済みリスト_{now}.xlsx"
        st.download_button(
            label="📥 Excelをダウンロード（整備済み・エラー・対応表・ログの4シート）",
            data=excel_bytes,
            file_name=dl_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True,
        )
        st.caption(f"出力ファイル: `{dl_name}`　（整備済みリスト / エラーリスト / 提供リスト対応表 / 整備ログ の4シート構成）")
else:
    st.info("👆 Excelファイル（.xlsx）をアップロードしてください")

st.divider()
st.markdown(
    f'<div style="text-align:center; color:#aaa; font-size:0.8rem;">'
    f'リスト整備ツール (β版) ｜ ver. {_get_version()}'
    f'</div>',
    unsafe_allow_html=True
)
