#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PT事業部 ダッシュボード更新サーバー
====================================
エンドポイント：
  POST /update  : 4ファイルを受け取り集計・pt_data.jsonを返す
  GET  /health  : サーバー稼働確認用
"""

import io
import json
import os
import traceback
from datetime import date, timedelta

import pandas as pd
import requests
from flask import Flask, jsonify, request
from flask_cors import CORS
from openpyxl import load_workbook

app = Flask(__name__)
CORS(app)

# ============================================================
# 定数
# ============================================================
SITE_LABEL    = {'新宿SC': '新宿SC', '在宅G': 'リモートSC', 'AI': 'AI'}
B_TO_D        = {'B0000106': 'D0000295', 'B0000107': 'D0000326'}
KONO          = '幸野有希子CRM'
EXCLUDE_OPS   = ['堀川璃歩']

# GitHubからスタッフマスターを自動読み込み
MASTER_URL = 'https://raw.githubusercontent.com/minakawa-star/digross-server/main/%E3%82%B9%E3%82%BF%E3%83%83%E3%83%95%E3%83%9E%E3%82%B9%E3%82%BF%E3%83%BC.xlsx'
_master_cache = None

def get_master_from_github():
    global _master_cache
    try:
        res = requests.get(MASTER_URL, timeout=30)
        res.raise_for_status()
        _master_cache = res.content
        return res.content
    except Exception as e:
        if _master_cache:
            return _master_cache
        raise ValueError(f'スタッフマスターの取得に失敗しました: {e}')


# ============================================================
# ヘルスチェック
# ============================================================
@app.route('/health', methods=['GET'])
def health():
    return jsonify({'status': 'ok', 'message': 'PT事業部ダッシュボードサーバー稼働中'})


# ============================================================
# メイン更新エンドポイント
# ============================================================
@app.route('/update', methods=['POST'])
def update():
    try:
        # --- ファイル受け取り ---
        if 'apo' not in request.files:
            return jsonify({'error': 'アポイントリストが見つかりません'}), 400
        if 'prod' not in request.files:
            return jsonify({'error': '生産性レポートが見つかりません'}), 400
        if 'work' not in request.files:
            return jsonify({'error': '勤務データが見つかりません'}), 400
        # スタッフマスター：アップロードされていればそちらを優先、なければGitHubから取得
        if 'prev' not in request.files:
            return jsonify({'error': '前回のpt_data.jsonが見つかりません'}), 400

        apo_file    = request.files['apo'].read()
        prod_file   = request.files['prod'].read()
        work_file   = request.files['work'].read()
        master_file = (request.files['master'].read()
                       if 'master' in request.files
                       else get_master_from_github())
        prev_json   = json.loads(request.files['prev'].read().decode('utf-8'))
        inc_map     = prev_json.get('incentive', {})

        # インセンティブファイル（任意）
        if 'inc' in request.files:
            inc_map = _load_incentive(request.files['inc'].read())

        # --- マスター読み込み ---
        df_master, df_wage = _load_master(master_file)
        site_map  = dict(zip(df_master['スタッフ名'], df_master['サイト']))
        rank_map  = dict(zip(df_master['スタッフ名'], df_master['ランク']))
        id_map    = dict(zip(df_master['スタッフ名'], df_master['社員番号']))
        master_ids = set(df_master['社員番号'].tolist())

        # --- アポイントリスト ---
        df_apo = _load_apo(apo_file)

        # --- 生産性レポート ---
        df_prod = _load_csv(prod_file)

        # --- 勤務データ ---
        df_work = _load_work(work_file)

        # --- 設定値 ---
        elapsed = prev_json['meta']['elapsedDays']
        working = prev_json['meta']['workingDays']

        # --- 新規営業日の特定 ---
        existing_dates = set(
            row['date']
            for rows in prev_json['daily'].values()
            for row in rows
        )
        all_dates = set(df_apo['取得日'].dropna().unique())
        new_dates = sorted(all_dates - existing_dates)

        if not new_dates:
            return jsonify({'error': '新しい営業日データが見つかりません'}), 400

        target_dates = sorted(existing_dates | set(new_dates))

        # --- 人件費計算 ---
        site_labor, work_by_id, labor_by_id = _calc_labor(
            df_work, df_master, df_wage, master_ids, working)

        # --- コール・稼働人数（日次）---
        calls_by_date = {}
        ops_by_date   = {}
        for d in new_dates:
            d_prod = d.replace('/', '-')
            mask = df_prod['日付'] == d_prod
            calls_by_date[d] = int(df_prod[mask]['コール数'].sum())
            ops_by_date[d]   = int(len(df_prod[mask]))

        # --- 日次明細追加 ---
        for date_str in new_dates:
            elapsed += 1
            daily = _calc_daily(df_apo, date_str, site_map,
                                calls_by_date, ops_by_date, elapsed)
            for key in ['all', 'shinjuku', 'remote', 'ai']:
                prev_json['daily'][key].append(daily[key])

        prev_json['meta']['elapsedDays'] = elapsed

        # --- サイト別累計更新 ---
        inc_site = {'新宿SC': 0, 'リモートSC': 0, 'AI': 0}
        for name, inc_total in inc_map.items():
            if inc_total <= 0:
                continue
            site = SITE_LABEL.get(site_map.get(name, ''), '')
            if site in inc_site:
                inc_site[site] += round(inc_total / working * elapsed)
        inc_all = sum(inc_site.values())

        for k in ['all', 'shinjuku', 'remote', 'ai']:
            rows   = prev_json['daily'][k]
            sales  = sum(r['sales']  for r in rows)
            apo    = sum(r['apo']    for r in rows)
            cancel = sum(r['cancel'] for r in rows)
            calls  = sum(r['calls']  for r in rows)
            valid  = apo - cancel
            cr     = round(cancel / apo * 100, 1) if apo > 0 else 0
            site_jp = {'all': None, 'shinjuku': '新宿SC',
                       'remote': 'リモートSC', 'ai': 'AI'}[k]
            jinjer = {
                'all':      sum(site_labor.values()),
                'shinjuku': site_labor.get('新宿SC', 0),
                'remote':   site_labor.get('リモートSC', 0),
                'ai':       site_labor.get('AI', 0),
            }[k]
            labor  = jinjer + (inc_all if k == 'all'
                               else inc_site.get(site_jp, 0))
            cost   = round(labor / sales * 100, 1) if sales > 0 else 0
            ar     = round(apo / calls * 100, 2)   if calls > 0 else 0
            last   = rows[-1] if rows else {}
            unit_last = (round(last['sales'] / last['ops'])
                         if last and last.get('ops', 0) > 0 else 0)
            prev_json['sites'][k] = {
                'sales': sales, 'apo': apo, 'cancel': cancel, 'valid': valid,
                'cancelRate': cr, 'labor': labor, 'costRate': cost,
                'gross': sales - labor, 'ops': last.get('ops', 0),
                'unit': unit_last, 'calls': calls, 'apoRate': ar,
            }

        # --- チャート更新 ---
        chart_labels = prev_json['chart']['labels']
        chart_data   = [0.0] * len(chart_labels)
        for row in prev_json['daily']['all']:
            d_label = row['date'].replace('2026/', '').lstrip('0').replace('/0', '/')
            if d_label in chart_labels:
                idx = chart_labels.index(d_label)
                chart_data[idx] = round(row['sales'] / 10000, 1)
        prev_json['chart']['data'] = chart_data

        # --- ヒートマップ更新 ---
        prev_json['heatmap'] = _calc_heatmap(
            df_apo, list(target_dates),
            prev_json['sites']['all']['sales'])

        # --- OP個人実績更新 ---
        prev_calls = prev_json.get('prev_calls', {})
        prev_json['operators'] = _calc_operators(
            df_apo, list(target_dates), df_master, df_prod,
            work_by_id, labor_by_id, id_map, site_map, rank_map,
            inc_map, elapsed, working, prev_calls)

        # --- メタ更新 ---
        today = date.today().strftime('%Y/%m/%d')
        last_new = new_dates[-1].replace('2026/', '')
        prev_json['meta']['lastUpdate']      = today
        prev_json['meta']['lastUpdateLabel'] = \
            f"{date.today().strftime('%m/%d')}（{last_new}分反映済）"
        prev_json['meta']['alertText'] = \
            f"{last_new}のデータを反映済みです（累計{elapsed}営業日）。最終更新: {today}"

        if 'inc' in request.files:
            prev_json['incentive'] = inc_map

        # --- 検証 ---
        j = json.dumps(prev_json, ensure_ascii=False)
        assert 'NaN'      not in j
        assert 'Infinity' not in j

        return jsonify({'status': 'ok', 'data': prev_json})

    except Exception as e:
        return jsonify({'error': str(e), 'detail': traceback.format_exc()}), 500


# ============================================================
# ヘルパー関数
# ============================================================
def _load_master(file_bytes):
    wb = load_workbook(io.BytesIO(file_bytes), read_only=True)
    ws_m = wb['スタッフマスター']
    df_m = pd.DataFrame(
        list(ws_m.iter_rows(values_only=True))[1:],
        columns=['社員番号', 'スタッフ名', 'サイト', 'ランク']
    ).dropna(subset=['社員番号']).fillna('')
    df_m['社員番号'] = df_m['社員番号'].astype(str).str.strip()

    ws_w = wb['時給マスター']
    df_w = pd.DataFrame(
        list(ws_w.iter_rows(values_only=True))[1:],
        columns=['社員番号', 'スタッフ名', '時給', '備考']
    ).dropna(subset=['社員番号']).fillna('')
    df_w['社員番号'] = df_w['社員番号'].astype(str).str.strip()
    df_w['時給'] = pd.to_numeric(df_w['時給'], errors='coerce').fillna(0).astype(int)
    return df_m, df_w


def _load_apo(file_bytes):
    wb = load_workbook(io.BytesIO(file_bytes), read_only=True)
    ws = wb['Sheet1']
    rows = list(ws.iter_rows(values_only=True))
    df = pd.DataFrame(rows[1:], columns=rows[0])
    df = df[~df['スタッフ名'].isin(EXCLUDE_OPS)].copy()
    df['スタッフ名'] = df['スタッフ名'].replace('君塚綾子', '君塚綾子1104')
    df['cancel_date_str'] = df['キャンセル受付日'].astype(str).str.strip()
    df['sales'] = pd.to_numeric(df['案件金額'], errors='coerce').fillna(0)
    return df


def _load_csv(file_bytes):
    for enc in ['utf-8-sig', 'cp932']:
        try:
            return pd.read_csv(io.BytesIO(file_bytes), encoding=enc)
        except Exception:
            continue
    raise ValueError('CSVの読み込みに失敗しました')


def _load_work(file_bytes):
    for enc in ['cp932', 'utf-8-sig']:
        try:
            df = pd.read_csv(io.BytesIO(file_bytes), encoding=enc)
            df = df.dropna(subset=['従業員ID'])
            df['総労働時間'] = pd.to_numeric(df['総労働時間'], errors='coerce').fillna(0)
            df['出勤日数']   = pd.to_numeric(df['出勤日数'],   errors='coerce').fillna(0)
            return df.sort_values('総労働時間', ascending=False).drop_duplicates(subset='従業員ID')
        except Exception:
            continue
    raise ValueError('勤務データの読み込みに失敗しました')


def _load_incentive(file_bytes):
    wb = load_workbook(io.BytesIO(file_bytes), read_only=True)
    ws = wb.active
    result = {}
    for row in list(ws.iter_rows(values_only=True))[1:]:
        if len(row) >= 3 and row[2]:
            name   = str(row[1]).strip() if row[1] else ''
            amount = int(pd.to_numeric(row[2], errors='coerce') or 0)
            if name and amount > 0:
                result[name] = amount
    return result


def _calc_labor(df_work, df_master, df_wage, master_ids, working):
    site_map_id = dict(zip(df_master['社員番号'], df_master['サイト']))
    wage_by_id  = dict(zip(df_wage['社員番号'], df_wage['時給']))
    note_by_id  = dict(zip(df_wage['社員番号'], df_wage['備考']))
    df_active   = df_work[(df_work['出勤日数'] >= 1) & (df_work['総労働時間'] > 0)]

    site_labor  = {'新宿SC': 0, 'リモートSC': 0, 'AI': 0}
    work_by_id  = {}
    labor_by_id = {}

    for _, row in df_active.iterrows():
        emp_id = str(row['従業員ID']).strip()
        hours  = float(row['総労働時間'])
        days   = float(row['出勤日数'])
        lookup = B_TO_D.get(emp_id, emp_id)
        if lookup not in master_ids and emp_id not in master_ids:
            continue
        site_r = site_map_id.get(lookup) or site_map_id.get(emp_id, '')
        site_d = SITE_LABEL.get(site_r, 'その他')
        wage   = wage_by_id.get(lookup) or wage_by_id.get(emp_id)
        note   = str(note_by_id.get(lookup) or note_by_id.get(emp_id, ''))
        if not wage:
            continue
        wage = float(wage)
        cost = (round(wage * 1.15 / working * days)
                if '月給' in note else round(wage * hours))
        site_labor[site_d] = site_labor.get(site_d, 0) + cost
        work_by_id[lookup]  = work_by_id[emp_id]  = hours
        labor_by_id[lookup] = labor_by_id[emp_id] = cost

    return site_labor, work_by_id, labor_by_id


def _calc_daily(df_apo, date_str, site_map,
                calls_by_date, ops_by_date, day_num):
    df_g = df_apo[
        (df_apo['取得日'] == date_str) &
        (df_apo['再送当否'].astype(str).str.strip() != '再送') &
        (df_apo['スタッフ名'] != KONO)
    ].copy()
    df_g['site_raw'] = df_g['スタッフ名'].map(site_map)

    df_c = df_apo[
        (df_apo['cancel_date_str'] == date_str) &
        (df_apo['スタッフ名'] != KONO)
    ].copy()
    df_c['site_raw'] = df_c['スタッフ名'].map(site_map)

    result = {}
    for key, raw in [('all', None), ('shinjuku', '新宿SC'),
                     ('remote', '在宅G'), ('ai', 'AI')]:
        g = df_g if raw is None else df_g[df_g['site_raw'] == raw]
        c = df_c if raw is None else df_c[df_c['site_raw'] == raw]
        result[key] = {
            'day':    f'{day_num}営業日',
            'date':   date_str,
            'sales':  int(g['sales'].sum()) - int(c['sales'].sum()),
            'apo':    len(g),
            'cancel': len(c),
            'valid':  len(g) - len(c),
            'ops':    ops_by_date.get(date_str, 0),
            'calls':  calls_by_date.get(date_str, 0) if key == 'all' else 0,
        }
    return result


def _calc_heatmap(df_apo, target_dates, total_sales):
    df_g = df_apo[
        df_apo['取得日'].isin(target_dates) &
        (df_apo['再送当否'].astype(str).str.strip() != '再送') &
        (df_apo['スタッフ名'] != KONO)
    ].copy()
    df_c = df_apo[
        df_apo['cancel_date_str'].isin(target_dates) &
        (df_apo['スタッフ名'] != KONO)
    ].copy()
    pg = df_g.groupby('登録案件名').agg(
        apo=('アポイントID', 'count'), sg=('sales', 'sum')).reset_index()
    pc = df_c.groupby('登録案件名').agg(
        cxl=('アポイントID', 'count'), sc=('sales', 'sum')).reset_index()
    proj = pg.merge(pc, on='登録案件名', how='left').fillna(0)
    proj['valid'] = (proj['apo'] - proj['cxl']).astype(int)
    proj['net']   = (proj['sg']  - proj['sc']).astype(int)
    proj = proj[proj['valid'] > 0].sort_values('net', ascending=False).head(10)
    return [
        {'rank': i + 1, 'name': str(r['登録案件名']),
         'valid': int(r['valid']), 'sales': int(r['net']),
         'pct': round(r['net'] / total_sales * 100, 1) if total_sales > 0 else 0}
        for i, (_, r) in enumerate(proj.iterrows())
    ]


def _calc_operators(df_apo, target_dates, df_master, df_prod,
                    work_by_id, labor_by_id, id_map,
                    site_map, rank_map, inc_map,
                    elapsed, working, prev_calls):
    df_get = df_apo[
        df_apo['取得日'].isin(target_dates) &
        (df_apo['再送当否'].astype(str).str.strip() != '再送')
    ].copy()
    kouryo = (
        df_apo['cancel_date_str'].str.contains('考慮', na=False) &
        df_apo['cancel_date_str'].str.extract(
            r'(\d{4}/\d+/\d+)', expand=False).isin(target_dates)
    )
    df_cxl_op   = df_apo[df_apo['cancel_date_str'].isin(target_dates) | kouryo].copy()
    df_cxl_dash = df_apo[df_apo['cancel_date_str'].isin(target_dates)].copy()

    op_get    = df_get.groupby('スタッフ名').agg(
        apo=('アポイントID', 'count'), sg=('sales', 'sum')).reset_index()
    op_cxl_op = df_cxl_op.groupby('スタッフ名').agg(
        cxl_op=('アポイントID', 'count'), sc_op=('sales', 'sum')).reset_index()
    op_cxl_d  = df_cxl_dash.groupby('スタッフ名').agg(
        cxl_d=('アポイントID', 'count'), sc_d=('sales', 'sum')).reset_index()

    calls_new  = df_prod.groupby('エージェント')['コール数'].sum().to_dict()
    all_names  = set(list(df_master['スタッフ名']) + list(op_get['スタッフ名']))
    operators  = []

    for name in all_names:
        g  = op_get[op_get['スタッフ名'] == name]
        co = op_cxl_op[op_cxl_op['スタッフ名'] == name]
        cd = op_cxl_d[op_cxl_d['スタッフ名'] == name]
        apo    = int(g['apo'].iloc[0])      if len(g)  > 0 else 0
        sg     = int(g['sg'].iloc[0])       if len(g)  > 0 else 0
        cxl_op = int(co['cxl_op'].iloc[0]) if len(co) > 0 else 0
        sc_op  = int(co['sc_op'].iloc[0])  if len(co) > 0 else 0
        cxl_d  = int(cd['cxl_d'].iloc[0]) if len(cd) > 0 else 0
        sc_d   = int(cd['sc_d'].iloc[0])   if len(cd) > 0 else 0
        net_op   = sg - sc_op
        net_dash = sg - sc_d
        calls    = prev_calls.get(name, 0) + int(calls_new.get(name, 0))
        emp_id   = id_map.get(name, '')
        lookup   = B_TO_D.get(emp_id, emp_id)
        hours    = work_by_id.get(lookup)  or work_by_id.get(emp_id, 0)
        l_base   = labor_by_id.get(lookup) or labor_by_id.get(emp_id, 0)
        inc_t    = inc_map.get(name, 0)
        inc_day  = round(inc_t / working * elapsed) if inc_t > 0 else 0
        labor    = l_base + inc_day
        ar       = round(apo / calls * 100, 1)     if calls > 0 else None
        cost_r   = round(labor / net_op * 100, 1)  if net_op > 0 and labor > 0 else None
        site_d   = SITE_LABEL.get(site_map.get(name, ''), '')
        rank_d   = rank_map.get(name, '')
        # 稼働単価（売上÷出勤日数）
        days_work = float(hours) / 8 if hours > 0 else 0  # 簡易換算
        unit_pd   = round(net_op / days_work) if days_work > 0 and net_op > 0 else 0

        operators.append({
            'name': name, 'site': site_d, 'rank': rank_d,
            'sales': net_op, 'sales_dash': net_dash,
            'apo': apo, 'cancel': cxl_op, 'valid': apo - cxl_op,
            'calls': calls, 'apoRate': ar,
            'workH': round(float(hours), 1), 'labor': labor,
            'labor_base': l_base, 'incentive_daily': inc_day,
            'unitPerDay': unit_pd, 'costRate': cost_r,
        })

    operators.sort(key=lambda x: (-x['sales'] if x['sales'] > 0
                                  else (0 if x['sales'] == 0 else 1)))
    return operators


# ============================================================
# 起動
# ============================================================
if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
