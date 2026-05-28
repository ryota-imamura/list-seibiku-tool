"""
エラーリスト検証ツール

過去のエラーxlsxを最新コードで再処理し、何件解決するか／誤マッチがないか確認する。
リポジトリ修正後に必ず実行し、デプロイ前に挙動を確認するためのスクリプト。

使い方:
    python3 verify_errors.py /path/to/整備済みリスト_xxx.xlsx
    python3 verify_errors.py  # ~/Downloads/整備済みリスト_*.xlsx を自動探索
"""
import sys
import glob
import os
import re
import json
import urllib.request
import time
import pandas as pd

from list_processor import lookup_postal_from_address, _build_city_pref_cache


def verify(xlsx_path, check_match=True):
    print(f'=== {os.path.basename(xlsx_path)} ===')
    df = pd.read_excel(xlsx_path, sheet_name='エラーリスト')
    total = len(df)
    ok = ng = mismatch = 0
    ng_list = []
    mismatch_list = []
    for _, row in df.iterrows():
        addr = str(row.get('オーナー住所', '') or '')
        postal, reason = lookup_postal_from_address(addr)
        if not postal:
            ng += 1
            ng_list.append((row['元行番号'], row.get('オーナー名', ''), addr, reason))
            continue
        if not check_match:
            ok += 1
            continue
        # zipcloudで補完先住所を確認し、元住所と整合するか確認
        code = postal.replace('-', '')
        try:
            data = json.loads(urllib.request.urlopen(
                f'https://zipcloud.ibsnet.co.jp/api/search?zipcode={code}', timeout=5).read())
            res = data.get('results', [{}])[0]
            resolved = f"{res.get('address1','')}{res.get('address2','')}{res.get('address3','')}"
        except Exception:
            resolved = ''
        # 元住所に含まれる地名トークンが補完先に含まれているかをざっくり確認
        # 「市町村+町名らしき部分」が補完先住所に含まれるかをチェック
        ok += 1
        # 簡易整合チェック: 元住所の都道府県+市 が補完先と同じか
        m = re.match(r'^((?:東京都|北海道|(?:大阪|京都)府|.{2,3}県)(?:.+?(?:市|区|町|村)))', addr)
        if m and resolved:
            src_prefix = m.group(1)
            if not (resolved.startswith(src_prefix[:5]) or src_prefix[:5] in resolved[:10]):
                mismatch += 1
                mismatch_list.append((row['元行番号'], addr, postal, resolved))
        time.sleep(0.05)
    print(f'  成功: {ok}/{total} | NG: {ng} | 誤マッチ疑い: {mismatch}')
    if ng_list:
        print('  --- NG ---')
        for n, name, addr, reason in ng_list[:20]:
            print(f'    {n} {name}: {addr[:40]} | {reason}')
    if mismatch_list:
        print('  --- 誤マッチ疑い ---')
        for n, addr, p, r in mismatch_list[:10]:
            print(f'    {n}: {addr[:35]} → {p} ({r})')
    return ng, mismatch


if __name__ == '__main__':
    print('キャッシュ構築中...')
    _build_city_pref_cache()
    print('完了\n')
    if len(sys.argv) > 1:
        files = sys.argv[1:]
    else:
        files = sorted(glob.glob(os.path.expanduser('~/Downloads/整備済みリスト_*.xlsx')))
    if not files:
        print('xlsxファイルが見つかりません')
        sys.exit(1)
    total_ng = total_mm = 0
    for f in files:
        ng, mm = verify(f)
        total_ng += ng
        total_mm += mm
        print()
    print(f'=== 合計 NG: {total_ng}, 誤マッチ疑い: {total_mm} ===')
