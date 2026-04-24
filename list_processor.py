import io
import re
import time
import json
import unicodedata
import urllib.request
import urllib.parse

import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment

# ── 正規化ユーティリティ ──────────────────────────────────────────────

def to_halfwidth(s):
    if not isinstance(s, str):
        return s
    return unicodedata.normalize('NFKC', s).strip()

def normalize_postal(s):
    if not isinstance(s, str):
        s = str(s) if pd.notna(s) else ""
    s = to_halfwidth(s)
    for ch in ["-", "ー", "−", "‐", "－"]:
        s = s.replace(ch, "")
    digits = re.sub(r'\D', '', s)
    if len(digits) == 7:
        return f"{digits[:3]}-{digits[3:]}"
    return None

def is_valid_postal(s):
    return bool(s and re.match(r'^\d{3}-\d{4}$', str(s)))

def is_garbled(s):
    if not isinstance(s, str):
        return False
    for p in [r'[\?]{3,}', r'[□]{3,}', r'[〓]{3,}', r'[■]{3,}', r'\?{3,}']:
        if re.search(p, s):
            return True
    return False

PREF_PATTERN = re.compile(r'^(東京都|北海道|(?:大阪|京都)府|.{2,3}県)')

def has_prefecture(address):
    return bool(address and PREF_PATTERN.match(address))

def normalize_address_for_compare(addr):
    if not addr:
        return ""
    s = to_halfwidth(addr)
    s = re.sub(r'(\d+)丁目', r'\1-', s)
    s = re.sub(r'(\d+)番地', r'\1-', s)
    s = re.sub(r'(\d+)番',  r'\1-', s)
    s = re.sub(r'(\d+)号',  r'\1',  s)
    s = re.sub(r'-+$', '', s)
    s = re.sub(r'-{2,}', '-', s)
    return s.strip()

# ── API呼び出し ───────────────────────────────────────────────────────

def _get_json(url, timeout=5):
    try:
        with urllib.request.urlopen(url, timeout=timeout) as r:
            return json.loads(r.read().decode('utf-8'))
    except Exception:
        return None

def lookup_address_from_postal(postal):
    code = postal.replace("-", "")
    data = _get_json(f"https://zipcloud.ibsnet.co.jp/api/search?zipcode={code}")
    if data and data.get('results'):
        res = data['results'][0]
        addr = res.get('address1','') + res.get('address2','') + res.get('address3','')
        return addr.strip() or None
    return None

def lookup_prefecture_from_postal(postal):
    if not is_valid_postal(postal):
        return None
    code = postal.replace("-", "")
    data = _get_json(f"https://zipcloud.ibsnet.co.jp/api/search?zipcode={code}")
    if data and data.get('results'):
        return data['results'][0].get('address1')
    return None

# 都道府県リスト
_ALL_PREFS = [
    '北海道','青森県','岩手県','宮城県','秋田県','山形県','福島県',
    '茨城県','栃木県','群馬県','埼玉県','千葉県','東京都','神奈川県',
    '新潟県','富山県','石川県','福井県','山梨県','長野県','岐阜県',
    '静岡県','愛知県','三重県','滋賀県','京都府','大阪府','兵庫県',
    '奈良県','和歌山県','鳥取県','島根県','岡山県','広島県','山口県',
    '徳島県','香川県','愛媛県','高知県','福岡県','佐賀県','長崎県',
    '熊本県','大分県','宮崎県','鹿児島県','沖縄県',
]
_CITY_PREF_CACHE = {}  # {city_name: pref}

def _build_city_pref_cache():
    """全都道府県の市区町村→都道府県マッピングをHeartRailsから構築（初回のみ）"""
    global _CITY_PREF_CACHE
    if _CITY_PREF_CACHE:
        return
    for pref in _ALL_PREFS:
        url = f"https://geoapi.heartrails.com/api/json?method=getCities&prefecture={urllib.parse.quote(pref)}"
        data = _get_json(url)
        if not data:
            continue
        for loc in data.get('response', {}).get('location', []):
            city_full = loc.get('city', '')
            if not city_full:
                continue
            _CITY_PREF_CACHE[city_full] = pref
            # 政令指定都市: 「北九州市小倉南区」→「北九州市」も登録
            m = re.match(r'^(.+市)', city_full)
            if m:
                _CITY_PREF_CACHE.setdefault(m.group(1), pref)
        time.sleep(0.3)

def lookup_prefecture_from_city(address):
    """住所の先頭から市区町村名を抽出して都道府県を逆引き"""
    _build_city_pref_cache()
    # 非欲張りで最短一致: 鳥栖市本鳥栖町→鳥栖市、北九州市小倉南区→北九州市
    for pat in [r'^(.{2,12}?[市区町村])', r'^(.{2,6}?[市区町村])']:
        m = re.match(pat, address)
        if m:
            city = m.group(1)
            if city in _CITY_PREF_CACHE:
                return _CITY_PREF_CACHE[city], city
    return None, None

def _extract_city_and_town(address):
    """住所から都道府県除去済み文字列の市区町村と町域を抽出"""
    city_m = re.match(r'^(.{2,10}?[市区町村])', address)
    if not city_m:
        return None, None
    city = city_m.group(1)
    town_rest = address[len(city):]
    # 政令指定都市: 市の後に「○○区」が続く場合は区まで含める
    if city.endswith('市'):
        ku_m = re.match(r'^([^\d一二三四五六七八九十]+区)', town_rest)
        if ku_m:
            ku = ku_m.group(1)
            city = city + ku
            town_rest = town_rest[len(ku):]
    # 大字・字 プレフィックスを除去
    town_rest = re.sub(r'^[大小]?字', '', town_rest)
    # 丁目/番地/番/号 の番号部分（算用数字・漢数字）以降を除去して町域名を取り出す
    # 例: 五郎丸四丁目3番1→五郎丸, 水ヶ江二丁目→水ヶ江, 曽根崎町1208番地→曽根崎町
    town = re.sub(r'[\d一二三四五六七八九十百千万]+(?:丁目|番地|番|号).*', '', town_rest).strip()
    return city, town

def lookup_postal_from_address(address):
    if not isinstance(address, str) or not address.strip():
        return None
    pref_m = PREF_PATTERN.match(address)
    if not pref_m:
        return None
    pref = pref_m.group(1)
    rest = address[len(pref):]
    city, town = _extract_city_and_town(rest)
    if not city or not town:
        return None
    url = (
        "https://geoapi.heartrails.com/api/json?method=getTowns"
        f"&prefecture={urllib.parse.quote(pref)}"
        f"&city={urllib.parse.quote(city)}"
    )
    data = _get_json(url)
    if data:
        def _kana_norm(s):
            return s.replace('ヶ', 'ケ').replace('ヵ', 'カ')
        town_clean = _kana_norm(re.sub(r'^[大小]?字', '', town))
        for loc in data.get('response', {}).get('location', []):
            loc_town = _kana_norm(re.sub(r'^[大小]?字', '', loc.get('town', '')))
            if loc_town.startswith(town_clean) or town_clean.startswith(loc_town):
                p = loc['postal']
                return f"{p[:3]}-{p[3:]}"
    return None

# ── 列マッピング ──────────────────────────────────────────────────────

def detect_columns(df):
    col_map = {k: None for k in
               ['オーナー名','連名①','連名②','連名③','連名④','連名⑤',
                '郵便番号','オーナー住所','物件名','物件住所']}
    alias_cols = []
    for col in df.columns:
        c = str(col)
        if any(k in c for k in ['所有者','オーナー','名義','氏名','代表者']) and col_map['オーナー名'] is None:
            col_map['オーナー名'] = col
        elif any(k in c for k in ['共有者','連名']):
            alias_cols.append(col)
        elif any(k in c for k in ['〒','zip','postal','郵便']):
            col_map['郵便番号'] = col
        elif any(k in c for k in ['居住地','オーナー住所','自宅','住所']) and '物件' not in c:
            col_map['オーナー住所'] = col
        elif any(k in c for k in ['物件名','建物名','マンション名','物件名称']):
            col_map['物件名'] = col
        elif any(k in c for k in ['物件所在地','物件住所']):
            col_map['物件住所'] = col
    for i, ac in enumerate(alias_cols[:5]):
        col_map[f'連名{["①","②","③","④","⑤"][i]}'] = ac
    return col_map

def get_val(row, col_map, key):
    col = col_map.get(key)
    if col is None:
        return ""
    v = row.get(col, "")
    if pd.isna(v):
        return ""
    return to_halfwidth(str(v)).strip()

# ── メイン処理 ────────────────────────────────────────────────────────

def process(file_bytes, progress_callback=None):
    """
    file_bytes: bytes (Excelファイル)
    progress_callback: callable(message: str) | None  進捗通知用
    戻り値: (excel_bytes, summary_dict, error_list)
    """
    def notify(msg, progress=None):
        if progress_callback:
            progress_callback(msg, progress)

    df_raw = pd.read_excel(io.BytesIO(file_bytes), header=0)
    col_map = detect_columns(df_raw)

    # 空行除去：オーナー名・オーナー住所・郵便番号がすべて空の行はスキップ
    key_cols = [c for c in [col_map.get('オーナー名'), col_map.get('オーナー住所'), col_map.get('郵便番号')] if c]
    if key_cols:
        df_raw = df_raw[df_raw[key_cols].notna().any(axis=1)].reset_index(drop=True)

    logs, errors, raw_rows = [], [], []
    seen_keys = set()
    dup_count = addr_fill_count = postal_fill_count = merge_count = garbled_count = 0

    # ── 行データ抽出 ──
    for idx, row in df_raw.iterrows():
        orig_no = idx + 2
        r = {
            'orig_no': orig_no,
            **{k: get_val(row, col_map, k)
               for k in ['オーナー名','連名①','連名②','連名③','連名④','連名⑤',
                         'オーナー住所','物件名','物件住所']},
            '郵便番号_raw': get_val(row, col_map, '郵便番号'),
            '郵便番号': '',
        }
        raw_rows.append(r)

    total = len(raw_rows)
    notify(f"住所・郵便番号を補完中... (全{total}件)", 0.05)

    # ── 郵便番号正規化 & 各種補完 ──
    for i, r in enumerate(raw_rows):
        if (i + 1) % 5 == 0 or i == total - 1:
            notify(f"住所・郵便番号を補完中... ({i+1}/{total}件)", 0.05 + 0.75 * (i + 1) / total)
        no = r['orig_no']
        postal_raw = r['郵便番号_raw']
        postal_norm = normalize_postal(postal_raw)

        if postal_norm and postal_norm != postal_raw:
            logs.append((no, f"郵便番号を正規化: 「{postal_raw}」→「{postal_norm}」"))
        r['郵便番号'] = postal_norm or ""

        # 都道府県補完（郵便番号優先 → 市区町村名から逆引き）
        if r['オーナー住所'] and not has_prefecture(r['オーナー住所']):
            orig = r['オーナー住所']
            pref = lookup_prefecture_from_postal(r['郵便番号'])
            method = "郵便番号から"
            if not pref:
                pref, _ = lookup_prefecture_from_city(orig)
                method = "市区町村名から"
                time.sleep(0.2)
            if pref:
                r['オーナー住所'] = pref + orig
                logs.append((no, f"都道府県を補完（{method}）: 「{orig}」→「{r['オーナー住所']}」"))
            else:
                logs.append((no, f"都道府県を特定できず: 「{orig}」"))

        # 郵便番号→住所補完（物件住所と一致する場合は禁止）
        if is_valid_postal(r['郵便番号']) and not r['オーナー住所']:
            filled = lookup_address_from_postal(r['郵便番号'])
            if filled:
                prop = r['物件住所']
                if prop and (prop.startswith(filled) or filled in prop):
                    logs.append((no, f"郵便番号({r['郵便番号']})の補完結果が物件住所と一致 → 物件の郵便番号と判断し補完を中止"))
                else:
                    logs.append((no, f"郵便番号からオーナー住所を補完: {r['郵便番号']}→「{filled}」"))
                    r['オーナー住所'] = filled
                    addr_fill_count += 1
            time.sleep(0.2)

        # 住所→郵便番号補完
        if r['オーナー住所'] and not is_valid_postal(r['郵便番号']):
            filled_postal = lookup_postal_from_address(r['オーナー住所'])
            if filled_postal:
                logs.append((no, f"オーナー住所から郵便番号を補完: 「{r['オーナー住所']}」→「{filled_postal}」"))
                r['郵便番号'] = filled_postal
                postal_fill_count += 1
            else:
                logs.append((no, f"オーナー住所から郵便番号を逆引きできず"))
            time.sleep(0.5)  # API レート制限対策

    notify("重複削除・連名統合を処理中...", 0.82)

    # ── 重複削除 ──
    dedup_rows = []
    for r in raw_rows:
        key = (r['オーナー名'], normalize_address_for_compare(r['オーナー住所']), r['郵便番号'])
        if key in seen_keys and r['オーナー名']:
            logs.append((r['orig_no'], f"重複行として除外 (オーナー名: {r['オーナー名']}, 住所: {r['オーナー住所']})"))
            dup_count += 1
            continue
        seen_keys.add(key)
        dedup_rows.append(r)

    # ── 同一住所の連名統合（表記揺れ正規化後に比較）──
    addr_groups = {}
    for r in dedup_rows:
        if r['オーナー住所']:
            k = normalize_address_for_compare(r['オーナー住所'])
            addr_groups.setdefault(k, []).append(r)

    merged_ids = set()
    final_rows = []
    for r in dedup_rows:
        if id(r) in merged_ids:
            continue
        norm_key = normalize_address_for_compare(r['オーナー住所']) if r['オーナー住所'] else ""
        group = addr_groups.get(norm_key, [])
        if r['オーナー住所'] and len(group) > 1 and id(group[0]) == id(r):
            all_names = []
            for g in group:
                if g['オーナー名']:
                    all_names.append(g['オーナー名'])
                for k in ['連名①','連名②','連名③','連名④','連名⑤']:
                    if g[k]:
                        all_names.append(g[k])
                merged_ids.add(id(g))
            merged = dict(r)
            merged['オーナー名'] = all_names[0] if all_names else ""
            for i, ak in enumerate(['連名①','連名②','連名③','連名④','連名⑤']):
                merged[ak] = all_names[i+1] if i+1 < len(all_names) else ""
            final_rows.append(merged)
            nos = ','.join(str(g['orig_no']) for g in group)
            logs.append((nos, f"同一住所({r['オーナー住所']})の行を統合: {', '.join(all_names)}"))
            merge_count += 1
        elif r['オーナー住所'] and len(group) > 1:
            merged_ids.add(id(r))
        else:
            final_rows.append(r)

    notify("エラー判定・出力ファイル作成中...", 0.92)

    # ── エラー判定 ──
    ok_rows = []
    for r in final_rows:
        reasons = []
        for f in [r['オーナー名'], r['オーナー住所'], r['物件名'], r['物件住所']]:
            if is_garbled(f):
                reasons.append("文字化けの疑い")
                garbled_count += 1
                break
        if not r['オーナー名']:
            reasons.append("オーナー名が未入力")
        if not r['オーナー住所']:
            reasons.append("オーナー住所が未入力（補完不可）")
        if not is_valid_postal(r['郵便番号']):
            reasons.append("郵便番号が未入力または形式不正（逆引き不可）")
        if reasons:
            r['エラー理由'] = ' / '.join(reasons)
            errors.append(r)
            logs.append((r['orig_no'], f"エラーリストへ: {r['エラー理由']}"))
        else:
            ok_rows.append(r)

    # ── Excel出力 ──
    wb = openpyxl.Workbook()

    def style_header(cell, bg):
        cell.font = Font(bold=True, color='FFFFFF', name='Arial', size=10)
        cell.fill = PatternFill('solid', start_color=bg)
        cell.alignment = Alignment(horizontal='center', vertical='center')

    def set_col_widths(ws, widths):
        for i, w in enumerate(widths, 1):
            ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w

    # シート1: 整備済みリスト
    ws1 = wb.active
    ws1.title = "整備済みリスト"
    h1 = ['No','オーナー名','連名①','連名②','連名③','連名④','連名⑤','郵便番号','オーナー住所','物件名','物件住所']
    for ci, h in enumerate(h1, 1):
        style_header(ws1.cell(row=1, column=ci, value=h), '1F4E79')
    for ri, r in enumerate(ok_rows, 2):
        ws1.cell(row=ri, column=1, value=ri-1)
        for ci, k in enumerate(h1[1:], 2):
            ws1.cell(row=ri, column=ci, value=r.get(k,'') or '')
    set_col_widths(ws1, [5,15,12,12,12,12,12,13,35,18,35])
    ws1.row_dimensions[1].height = 20

    # シート2: エラーリスト
    ws2 = wb.create_sheet("エラーリスト")
    h2 = ['元行番号','オーナー名','連名①','連名②','連名③','連名④','連名⑤','郵便番号','オーナー住所','物件名','物件住所','エラー理由']
    for ci, h in enumerate(h2, 1):
        style_header(ws2.cell(row=1, column=ci, value=h), 'C00000')
    for ri, r in enumerate(errors, 2):
        ws2.cell(row=ri, column=1, value=r.get('orig_no',''))
        for ci, k in enumerate(h2[1:], 2):
            ws2.cell(row=ri, column=ci, value=r.get(k,'') or '')
    set_col_widths(ws2, [10,15,12,12,12,12,12,13,35,18,35,45])
    ws2.row_dimensions[1].height = 20

    # シート3: 整備ログ
    ws3 = wb.create_sheet("整備ログ")
    for ci, h in enumerate(['行番号','処理内容'], 1):
        style_header(ws3.cell(row=1, column=ci, value=h), '375623')
    for ri, (no, msg) in enumerate(logs, 2):
        ws3.cell(row=ri, column=1, value=str(no))
        ws3.cell(row=ri, column=2, value=msg)
    set_col_widths(ws3, [12, 90])

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    excel_bytes = buf.read()

    summary = {
        '発注可能件数': len(ok_rows),
        'エラー件数': len(errors),
        '重複削除件数': dup_count,
        '住所補完件数': addr_fill_count,
        '郵便番号補完件数': postal_fill_count,
        '連名統合件数': merge_count,
        '文字化け検出件数': garbled_count,
    }
    error_list = [
        {'行番号': e['orig_no'], 'オーナー名': e.get('オーナー名',''), 'エラー理由': e.get('エラー理由','')}
        for e in errors
    ]
    return excel_bytes, summary, error_list
