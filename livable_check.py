"""
リバブル重複チェックモジュール

作業リスト × リバブル一括リスト の突合を行い、
過去にどのセンターがいつ・どのリストで送ったかをExcelに出力する。

マッチキー: (オーナー住所[:5] + オーナー名)[:10]  ← リバブル一括リストの「列3」と同一ロジック
"""

import io
import re
import unicodedata
from collections import defaultdict

import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment


# ── 正規化 ────────────────────────────────────────────────────────────

def _to_halfwidth(s: str) -> str:
    return unicodedata.normalize("NFKC", s).strip() if isinstance(s, str) else s


def _build_key(address, owner_name):
    """リバブル一括リストの「列3」と同じキーを生成"""
    addr = _to_halfwidth(str(address)).strip() if pd.notna(address) else ""
    name = _to_halfwidth(str(owner_name)).strip() if pd.notna(owner_name) else ""
    if not name:
        return None
    return (addr[:5] + name)[:10]


# ── リバブル一括リスト インデックス構築 ─────────────────────────────────

def build_livable_index(livable_bytes):
    """
    リバブル一括リストをロードし、列3キー → 送付履歴リスト のインデックスを返す。

    各エントリ:
      {'リスト名': str, 'センター名': str, '発送日': str}

    ※ リバブル一括リストでは「発注日」列にセンター名、「発送日」列に日付が入っている。
    """
    df = pd.read_excel(io.BytesIO(livable_bytes))

    # 列3が既存ならそれを使用、なければ再計算
    if "列3" in df.columns:
        key_col = df["列3"].astype(str)
    else:
        key_col = (
            df["オーナー住所"].astype(str).str[:5] + df["オーナー名"].astype(str)
        ).str[:10]

    index = defaultdict(list)
    for key, row in zip(key_col, df.itertuples(index=False)):
        if key and key != "nan":
            # 発注日列にセンター名が入っている（列名と内容が逆転）
            center = str(getattr(row, "発注日", "")) if pd.notna(getattr(row, "発注日", None)) else ""
            list_name = str(getattr(row, "リスト名", "")) if pd.notna(getattr(row, "リスト名", None)) else ""
            send_date = str(getattr(row, "発送日", "")) if pd.notna(getattr(row, "発送日", None)) else ""
            index[key].append({
                "リスト名": list_name,
                "センター名": center,
                "発送日": send_date,
            })

    return dict(index)


# ── メイン処理 ────────────────────────────────────────────────────────

def check_duplicates(
    work_bytes: bytes,
    livable_bytes: bytes,
    work_sheet=0,
    progress_callback=None,
) -> tuple[bytes, dict]:
    """
    作業リスト × リバブル一括リスト の重複チェック。

    Returns:
        (excel_bytes, summary_dict)

    summary_dict キー:
        total       - 作業リストの総行数
        dup_count   - 重複あり件数
        clean_count - 重複なし件数
        center_stats - {センター名: 件数} の集計
    """

    def notify(msg, pct=None):
        if progress_callback:
            progress_callback(msg, pct)

    notify("リバブル一括リストを読み込み中...", 0.05)
    livable_index = build_livable_index(livable_bytes)
    notify(f"インデックス構築完了: {len(livable_index):,}件のユニークキー", 0.20)

    df_work = pd.read_excel(io.BytesIO(work_bytes), sheet_name=work_sheet)
    total = len(df_work)
    notify(f"作業リスト読み込み完了: {total}行", 0.25)

    # 各行に対してマッチングを実行
    keys = []
    match_results: list[list[dict]] = []  # 行ごとのマッチリスト

    for i, row in df_work.iterrows():
        address = row.get("オーナー住所", "")
        owner = row.get("オーナー名", "")
        key = _build_key(address, owner)
        keys.append(key)
        matches = livable_index.get(key, []) if key else []
        match_results.append(matches)

        if (i + 1) % 50 == 0 or i + 1 == total:
            notify(f"マッチング中... ({i+1}/{total})", 0.25 + 0.55 * (i + 1) / total)

    # ── 付加列を生成 ──
    dup_flags = []
    center_names = []
    list_names = []
    send_dates = []
    match_counts = []

    for matches in match_results:
        if matches:
            # 重複: 同じリスト名が複数行ある場合は1件に集約
            seen_lists = {}
            for m in matches:
                ln = m["リスト名"]
                if ln not in seen_lists:
                    seen_lists[ln] = m
            deduped = list(seen_lists.values())
            # 発送日の昇順で並べる
            deduped.sort(key=lambda x: x["発送日"])

            dup_flags.append("有")
            center_names.append(" / ".join(d["センター名"] for d in deduped))
            list_names.append(" / ".join(d["リスト名"] for d in deduped))
            send_dates.append(" / ".join(d["発送日"] for d in deduped))
            match_counts.append(len(deduped))
        else:
            dup_flags.append("無")
            center_names.append("")
            list_names.append("")
            send_dates.append("")
            match_counts.append(0)

    df_out = df_work.copy()
    df_out["リバブル重複"] = dup_flags
    df_out["重複件数"] = match_counts
    df_out["重複センター名"] = center_names
    df_out["重複リスト名"] = list_names
    df_out["重複発送日"] = send_dates

    notify("Excel出力を作成中...", 0.85)

    # ── 重複詳細（1マッチ1行）──
    detail_rows = []
    for i, (row, matches) in enumerate(zip(df_work.itertuples(index=False), match_results)):
        owner = getattr(row, "オーナー名", "")
        postal = getattr(row, "郵便番号", "")
        address = getattr(row, "オーナー住所", "")
        if matches:
            seen_lists = {}
            for m in matches:
                if m["リスト名"] not in seen_lists:
                    seen_lists[m["リスト名"]] = m
            for m in sorted(seen_lists.values(), key=lambda x: x["発送日"]):
                detail_rows.append({
                    "オーナー名": owner,
                    "郵便番号": postal,
                    "オーナー住所": address,
                    "センター名": m["センター名"],
                    "リスト名": m["リスト名"],
                    "発送日": m["発送日"],
                })

    df_detail = pd.DataFrame(detail_rows) if detail_rows else pd.DataFrame(
        columns=["オーナー名", "郵便番号", "オーナー住所", "センター名", "リスト名", "発送日"]
    )

    # ── センター別集計 ──
    center_stats: dict[str, int] = defaultdict(int)
    if detail_rows:
        for r in detail_rows:
            center_stats[r["センター名"]] += 1

    # ── Excel作成 ──
    wb = openpyxl.Workbook()

    H_COLOR = {
        "blue":   "1F4E79",
        "red":    "C00000",
        "purple": "7030A0",
        "green":  "375623",
    }

    def style_header(cell, color_key):
        cell.font = Font(bold=True, color="FFFFFF", name="Arial", size=10)
        cell.fill = PatternFill("solid", start_color=H_COLOR[color_key])
        cell.alignment = Alignment(horizontal="center", vertical="center")

    def write_df(ws, df, color_key, col_widths=None):
        for ci, col in enumerate(df.columns, 1):
            style_header(ws.cell(row=1, column=ci, value=col), color_key)
        for ri, row in enumerate(df.itertuples(index=False), 2):
            for ci, val in enumerate(row, 1):
                ws.cell(row=ri, column=ci, value="" if pd.isna(val) else val)
        if col_widths:
            for i, w in enumerate(col_widths, 1):
                ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
        ws.row_dimensions[1].height = 20

    # シート1: 重複チェック済み作業リスト
    ws1 = wb.active
    ws1.title = "重複チェック済みリスト"

    base_cols = list(df_work.columns)
    added_cols = ["リバブル重複", "重複件数", "重複センター名", "重複リスト名", "重複発送日"]
    all_cols = base_cols + added_cols

    for ci, col in enumerate(all_cols, 1):
        style_header(ws1.cell(row=1, column=ci, value=col), "blue")

    # 重複行は背景色で強調
    FILL_DUP = PatternFill("solid", start_color="FCE4D6")

    for ri, (row_vals, flag) in enumerate(
        zip(df_out[all_cols].itertuples(index=False), dup_flags), 2
    ):
        for ci, val in enumerate(row_vals, 1):
            cell = ws1.cell(row=ri, column=ci, value="" if pd.isna(val) else val)
            if flag == "有":
                cell.fill = FILL_DUP

    # 列幅: 元列は元のまま、追加列を広く
    base_widths = [max(12, len(str(c)) + 2) for c in base_cols]
    added_widths = [10, 8, 35, 55, 20]
    for i, w in enumerate(base_widths + added_widths, 1):
        ws1.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
    ws1.row_dimensions[1].height = 20

    # シート2: 重複詳細（1マッチ1行）
    ws2 = wb.create_sheet("重複詳細")
    write_df(ws2, df_detail, "red", col_widths=[18, 13, 40, 35, 60, 15])

    # シート3: センター別集計
    ws3 = wb.create_sheet("センター別集計")
    center_df = pd.DataFrame(
        sorted(center_stats.items(), key=lambda x: -x[1]),
        columns=["センター名", "重複件数"],
    ) if center_stats else pd.DataFrame(columns=["センター名", "重複件数"])
    write_df(ws3, center_df, "purple", col_widths=[40, 12])

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)

    dup_count = sum(1 for f in dup_flags if f == "有")
    summary = {
        "total": total,
        "dup_count": dup_count,
        "clean_count": total - dup_count,
        "center_stats": dict(center_stats),
    }

    notify("完了", 1.0)
    return buf.read(), summary


# ── CLI 実行 ──────────────────────────────────────────────────────────

if __name__ == "__main__":
    import sys
    import argparse

    parser = argparse.ArgumentParser(description="リバブル重複チェックツール")
    parser.add_argument("work_list", help="作業リスト Excelファイルパス")
    parser.add_argument("livable_list", help="リバブル一括リスト Excelファイルパス")
    parser.add_argument("-o", "--output", default="重複チェック結果.xlsx", help="出力ファイルパス")
    parser.add_argument("-s", "--sheet", default=0, help="作業リストのシート名またはインデックス")
    args = parser.parse_args()

    sheet = int(args.sheet) if str(args.sheet).isdigit() else args.sheet

    with open(args.work_list, "rb") as f:
        work_bytes = f.read()
    with open(args.livable_list, "rb") as f:
        livable_bytes = f.read()

    def progress(msg, pct):
        bar = int((pct or 0) * 30)
        print(f"\r[{'#'*bar}{'.'*(30-bar)}] {msg}      ", end="", flush=True)

    excel_bytes, summary = check_duplicates(work_bytes, livable_bytes, sheet, progress)
    print()

    with open(args.output, "wb") as f:
        f.write(excel_bytes)

    print(f"\n=== 結果サマリー ===")
    print(f"  総行数       : {summary['total']:,}")
    print(f"  重複あり     : {summary['dup_count']:,}")
    print(f"  重複なし     : {summary['clean_count']:,}")
    if summary["center_stats"]:
        print(f"\n  センター別重複件数:")
        for center, cnt in sorted(summary["center_stats"].items(), key=lambda x: -x[1]):
            print(f"    {center}: {cnt}件")
    print(f"\n出力: {args.output}")
