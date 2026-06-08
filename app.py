#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PT事業部 ダッシュボード更新サーバー v2
========================================
設計方針：
  毎回アポイントリストから当月分を全て集計し直す（差分加算なし）
  前回のpt_data.jsonからは「目標・設定値・チャートラベル」のみ参照

エンドポイント：
  POST /update  : 4ファイルを受け取り集計・pt_data.jsonを返す
  GET  /health  : サーバー稼働確認用
"""

import io
import json
import os
import re
import traceback
import zipfile
from datetime import date, timedelta

import pandas as pd
import requests
from flask import Flask, jsonify, request
from flask_cors import CORS
from openpyxl import load_workbook
from supabase import create_client, Client

app = Flask(__name__)
CORS(app)

# ============================================================
# 定数
# ============================================================
SITE_LABEL  = {'新宿SC': '新宿SC', '在宅G': 'リモートSC', 'AI': 'AI'}
B_TO_D      = {'B0000106': 'D0000295', 'B0000107': 'D0000326', 'D0001318': 'B0000095'}
KONO        = '幸野有希子CRM'   # 全体/サイト集計から除外
EXCLUDE_OPS = ['堀川璃歩']      # 全集計から除外
# 在籍カウント除外スタッフ（名前が空・ランクなし・退職者等）
EXCLUDE_ENROLL = {'堀川璃歩', '堀川', ' 坂本杏奈1208', '藤公誉1212', '幸野有希子', '加藤隆治'}
VALID_RANKS    = {'トップ', 'ミドル', 'ルーキープラス', 'ルーキー', '管理者', '社員'}

# Supabaseクライアント
SUPABASE_URL = os.environ.get('SUPABASE_URL', '')
SUPABASE_KEY = os.environ.get('SUPABASE_KEY', '')
_supabase_client = None

def get_supabase() -> Client:
    global _supabase_client
    if not _supabase_client and SUPABASE_URL and SUPABASE_KEY:
        _supabase_client = create_client(SUPABASE_URL, SUPABASE_KEY)
    return _supabase_client

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
# スタッフ管理エンドポイント
# ============================================================
@app.route('/staff', methods=['GET'])
def get_staff():
    try:
        sb = get_supabase()
        if not sb:
            return jsonify({'error': 'Supabase未設定'}), 500
        res = sb.table('staff').select('*').order('site').order('rank').order('name').execute()
        return jsonify({'status': 'ok', 'data': res.data})
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/staff', methods=['POST'])
def add_staff():
    try:
        sb = get_supabase()
        if not sb:
            return jsonify({'error': 'Supabase未設定'}), 500
        body = request.get_json()
        res = sb.table('staff').insert(body).execute()
        return jsonify({'status': 'ok', 'data': res.data})
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/staff/<int:staff_id>', methods=['PUT'])
def update_staff(staff_id):
    try:
        sb = get_supabase()
        if not sb:
            return jsonify({'error': 'Supabase未設定'}), 500
        body = request.get_json()
        res = sb.table('staff').update(body).eq('id', staff_id).execute()
        return jsonify({'status': 'ok', 'data': res.data})
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/staff/<int:staff_id>', methods=['DELETE'])
def delete_staff(staff_id):
    try:
        sb = get_supabase()
        if not sb:
            return jsonify({'error': 'Supabase未設定'}), 500
        res = sb.table('staff').delete().eq('id', staff_id).execute()
        return jsonify({'status': 'ok'})
    except Exception as e:
        return jsonify({'error': str(e)}), 500

# インセンティブ管理
@app.route('/incentives/<month>', methods=['GET'])
def get_incentives(month):
    try:
        sb = get_supabase()
        if not sb:
            return jsonify({'error': 'Supabase未設定'}), 500
        res = sb.table('incentives').select('*').eq('month', month).execute()
        return jsonify({'status': 'ok', 'data': res.data})
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/incentives', methods=['POST'])
def upsert_incentive():
    try:
        sb = get_supabase()
        if not sb:
            return jsonify({'error': 'Supabase未設定'}), 500
        body = request.get_json()
        res = sb.table('incentives').upsert(body, on_conflict='employee_id,month').execute()
        return jsonify({'status': 'ok', 'data': res.data})
    except Exception as e:
        return jsonify({'error': str(e)}), 500

# ============================================================
# メイン更新エンドポイント
# ============================================================
@app.route('/update', methods=['POST'])
def update():
    try:
        # --- ファイル受け取り ---
        for key, label in [('apo','アポイントリスト'),('prod','生産性レポート'),
                            ('work','勤務データ'),('report','レポート')]:
            if key not in request.files:
                return jsonify({'error': f'{label}が見つかりません'}), 400

        apo_file  = request.files['apo'].read()
        prod_file = request.files['prod'].read()
        work_raw  = request.files['work'].read()
        work_file, work_end_date = _extract_work_from_zip(work_raw)
        prev_json = json.loads(request.files['prev'].read().decode('utf-8'))

        # スタッフマスター：アップロード優先、なければGitHubから取得
        master_file = (request.files['master'].read()
                       if 'master' in request.files
                       else get_master_from_github())

        # インセンティブ（任意）
        inc_map = prev_json.get('incentive', {})
        if 'inc' in request.files:
            inc_map = _load_incentive(request.files['inc'].read())

        # --- マスター読み込み ---
        df_master, df_wage = _load_master(master_file)
        site_map   = dict(zip(df_master['スタッフ名'], df_master['サイト']))
        rank_map   = dict(zip(df_master['スタッフ名'], df_master['ランク']))
        id_map     = dict(zip(df_master['スタッフ名'], df_master['社員番号']))
        master_ids = set(df_master['社員番号'].tolist())

        # --- アポイントリスト読み込み ---
        df_apo = _load_apo(apo_file)

        # --- 生産性レポート読み込み ---
        df_prod = _load_csv(prod_file)

        # --- 勤務データ読み込み ---
        df_work = _load_work(work_file)

        # --- 設定値（前回pt_dataから引き継ぐもの）---
        working  = prev_json['meta']['workingDays']   # 当月営業日数
        targets  = prev_json['targets']               # 目標値
        chart_labels = prev_json['chart']['labels']   # チャートラベル

        # ============================================================
        # 当月営業日の特定
        # ============================================================
        all_get_dates = sorted(df_apo['取得日'].dropna().unique())

        month_str = prev_json['meta']['month']  # 例：2026年5月
        m = re.search(r'(\d{4})年(\d+)月', month_str)
        target_year  = int(m.group(1))
        target_month = int(m.group(2))

        biz_dates = sorted([
            d for d in all_get_dates
            if str(d).startswith(f'{target_year}/{str(target_month).zfill(2)}')
        ])

        if not biz_dates:
            return jsonify({'error': '当月のアポイントデータが見つかりません'}), 400

        elapsed = len(biz_dates)

        # 6月以降は幸野有希子CRMを通常スタッフとして全集計に含める
        kono_excluded = target_month < 6  # True=除外（5月まで）, False=含める（6月以降）

        # ============================================================
        # jinjer勤務データの期間検証
        # ファイル名の終了日とアポリストの最終営業日を比較
        # 差が2営業日以上あれば警告をmetaに記録
        # ============================================================
        work_period_warning = None
        if work_end_date:
            last_biz = biz_dates[-1]  # 例：2026/05/28
            # 終了日を yyyy/mm/dd に正規化
            work_end_norm = '/'.join(p.zfill(2) for p in work_end_date.split('/'))
            if work_end_norm < last_biz:
                work_period_warning = (
                    f'jinjer勤務データの終了日（{work_end_date}）が'
                    f'アポリストの最終営業日（{last_biz}）より前です。'
                    f'月給制スタッフの人件費が過小計上になる可能性があります。'
                    f'jinjerは当月1日〜最終更新日の期間で出力してください。'
                )

        # ============================================================
        # 人件費計算（勤務データから）
        # 【修正】days_by_id も返すよう変更
        # ============================================================
        site_labor, work_by_id, labor_by_id, days_by_id = _calc_labor(
            df_work, df_master, df_wage, master_ids, working)

        # ============================================================
        # コール・稼働人数（生産性レポートから日次集計）
        # 【修正】日付フォーマットを正規化して突合
        # ============================================================
        calls_by_date = {}
        ops_by_date   = {}

        # 生産性レポートの日付列を正規化
        # 例：「2026/5/1」→「2026/05/01」、「2026-05-01」→「2026/05/01」
        def normalize_date(s):
            s = str(s).replace('-', '/').strip()
            parts = s.split('/')
            if len(parts) == 3:
                return f"{parts[0]}/{parts[1].zfill(2)}/{parts[2].zfill(2)}"
            return s

        if '日付' in df_prod.columns:
            df_prod['日付_norm'] = df_prod['日付'].apply(normalize_date)
        else:
            df_prod['日付_norm'] = ''

        # エージェント列名を特定
        agent_col = next((c for c in ['エージェント', 'エージェント名', 'Agent']
                          if c in df_prod.columns), None)

        for d in biz_dates:
            mask = df_prod['日付_norm'] == d
            calls_by_date[d] = int(df_prod[mask]['コール数'].sum()) if mask.any() else 0
            ops_by_date[d]   = int(df_prod[mask][agent_col].nunique()) if (mask.any() and agent_col) else 0

        # ============================================================
        # 日次明細
        # ============================================================
        daily = {'all': [], 'shinjuku': [], 'remote': [], 'ai': []}
        for i, date_str in enumerate(biz_dates, 1):
            d = _calc_daily(df_apo, date_str, site_map,
                            calls_by_date, ops_by_date, i, kono_excluded)
            for key in ['all', 'shinjuku', 'remote', 'ai']:
                daily[key].append(d[key])

        # ============================================================
        # サイト別累計
        # ============================================================
        inc_site = {'新宿SC': 0, 'リモートSC': 0, 'AI': 0}
        for name, inc_total in inc_map.items():
            if inc_total <= 0:
                continue
            site = SITE_LABEL.get(site_map.get(name, ''), '')
            if site in inc_site:
                inc_site[site] += round(inc_total / working * elapsed)
        inc_all = sum(inc_site.values())

        sites = {}
        for k in ['all', 'shinjuku', 'remote', 'ai']:
            rows   = daily[k]
            sales  = sum(r['sales']  for r in rows)
            apo    = sum(r['apo']    for r in rows)
            cancel = sum(r['cancel'] for r in rows)
            # 【修正】calls は daily['all'] の合計を使う（サイト別callsは0のため）
            calls  = sum(r['calls']  for r in daily['all']) if k == 'all' else 0
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
            all_ops = [op for op in _calc_operators(
                df_apo, biz_dates, df_master, df_prod,
                work_by_id, labor_by_id, days_by_id, id_map, site_map, rank_map,
                inc_map, elapsed, working, kono_excluded)
                if (k == 'all' or op['site'] == site_jp)
                and op['sales'] > 0 and op.get('days', 0) > 0]
            ts = sum(o['sales'] for o in all_ops)
            td = sum(o['days']  for o in all_ops)
            unit = round(ts / td) if td > 0 else 0

            sites[k] = {
                'sales': sales, 'apo': apo, 'cancel': cancel, 'valid': valid,
                'cancelRate': cr, 'labor': labor, 'costRate': cost,
                'gross': sales - labor, 'ops': last.get('ops', 0),
                'unit': unit, 'calls': calls, 'apoRate': ar,
            }

        # ============================================================
        # チャートデータ
        # ============================================================
        chart_data = [0.0] * len(chart_labels)
        for row in daily['all']:
            d_label = row['date'].replace(f'{target_year}/', '').lstrip('0').replace('/0', '/')
            if d_label in chart_labels:
                idx = chart_labels.index(d_label)
                chart_data[idx] = round(row['sales'] / 10000, 1)

        # ============================================================
        # ヒートマップ
        # ============================================================
        heatmap = _calc_heatmap(df_apo, biz_dates, sites['all']['sales'], kono_excluded)

        # ============================================================
        # OP個人実績
        # ============================================================
        operators = _calc_operators(
            df_apo, biz_dates, df_master, df_prod,
            work_by_id, labor_by_id, days_by_id, id_map, site_map, rank_map,
            inc_map, elapsed, working, kono_excluded)

        # ============================================================
        # 【修正】enrollCount・activeCount を計算
        # ============================================================
        # ============================================================
        # enrollCount・activeCount を計算
        # 在籍 = マスターに存在する有効スタッフ（除外対象・ランクなし除く）
        # アクティブ = 在籍スタッフ かつ 当月生産性レポートに名前がある
        # ============================================================
        master_names = set(df_master['スタッフ名'].tolist())
        # 6月以降（kono_excluded=False）はAIサイトも除外
        ai_excluded = not kono_excluded
        enroll_ops = [op for op in operators
                      if op['name'].strip() not in EXCLUDE_ENROLL
                      and op['rank'] in VALID_RANKS
                      and op['name'] in master_names
                      and not (ai_excluded and op.get('site') == 'AI')]
        enroll_count = len(enroll_ops)

        # 生産性レポートに登場するスタッフ名のセット
        if agent_col:
            prod_names = set(df_prod[agent_col].astype(str).str.strip().tolist())
        else:
            prod_names = set()
        active_count = sum(
            1 for op in enroll_ops
            if op['name'] in prod_names
        )

        # ============================================================
        # PT_DATA組み立て
        # ============================================================
        today = date.today().strftime('%Y/%m/%d')
        last_date = biz_dates[-1].replace(f'{target_year}/', '')
        month_label = f'{target_year}年{target_month}月'

        PT_DATA = {
            'meta': {
                'month':           month_label,
                'lastUpdate':      today,
                'lastUpdateLabel': f"{date.today().strftime('%m/%d')}（{last_date}分反映済）",
                'elapsedDays':     elapsed,
                'workingDays':     working,
                'alertText':       f"{last_date}のデータを反映済みです（累計{elapsed}営業日）。最終更新: {today}",
                'enrollCount':     enroll_count,   # 【追加】在籍数
                'activeCount':     active_count,   # 【追加】稼働数
                'workPeriodWarning': work_period_warning,  # jinjer期間ズレ警告
            },
            'targets':   targets,
            'sites':     sites,
            'daily':     daily,
            'chart':     {'labels': chart_labels, 'data': chart_data},
            'heatmap':   heatmap,
            'operators': operators,
            'incentive': inc_map,
            'prev_calls': {op['name']: op['calls'] for op in operators},
        }

        # 検証
        j = json.dumps(PT_DATA, ensure_ascii=False)
        assert 'NaN'      not in j, 'NaN混入'
        assert 'Infinity' not in j, 'Infinity混入'

        return jsonify({'status': 'ok', 'data': PT_DATA})

    except Exception as e:
        return jsonify({'error': str(e), 'detail': traceback.format_exc()}), 500


# ============================================================
# ヘルパー関数
# ============================================================
def _extract_work_from_zip(file_bytes):
    """
    zip圧縮の勤務データを展開してcsvバイトを返す。
    【追加】ファイル名から期間（終了日）を抽出して返す。
    戻り値: (csv_bytes, end_date_str or None)
    """
    try:
        with zipfile.ZipFile(io.BytesIO(file_bytes)) as zf:
            csv_files = [n for n in zf.namelist() if n.lower().endswith('.csv')]
            if not csv_files:
                raise ValueError('zip内にcsvファイルが見つかりません')
            target = max(csv_files, key=lambda n: zf.getinfo(n).file_size)
            # ファイル名から終了日を抽出
            # 例：勤務データ_打刻グループ_2026_05_01_2026_05_28.csv
            import re
            try:
                fname = target.encode('cp437').decode('cp932')
            except:
                fname = target
            end_date = None
            m = re.search(r'\d{4}_\d{2}_\d{2}_(\d{4}_\d{2}_\d{2})', fname)
            if m:
                end_date = m.group(1).replace('_', '/')
            return zf.read(target), end_date
    except zipfile.BadZipFile:
        return file_bytes, None


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
    """
    人件費計算（勤務データ全期間から）
    【修正】
    - 月給制スタッフは総労働時間=0でも出勤日数>0なら稼働扱い
    - days_by_id を返す（_calc_operatorsで出勤日数に使用）
    """
    site_map_id = dict(zip(df_master['社員番号'], df_master['サイト']))
    wage_by_id  = dict(zip(df_wage['社員番号'], df_wage['時給']))
    note_by_id  = dict(zip(df_wage['社員番号'], df_wage['備考']))

    # 【修正】月給制は出勤日数>0のみで稼働判定、時給制は従来通り
    # まず全行をループして月給/時給で条件分岐
    site_labor  = {'新宿SC': 0, 'リモートSC': 0, 'AI': 0}
    work_by_id  = {}
    labor_by_id = {}
    days_by_id  = {}

    for _, row in df_work.iterrows():
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
        is_monthly = '月給' in note

        # 【修正】稼働判定：月給制は出勤日数>0、時給制は労働時間>0
        if is_monthly:
            if days <= 0:
                continue
            cost = round(wage * 1.15 / working * days)
        else:
            if hours <= 0:
                continue
            cost = round(wage * hours)

        site_labor[site_d] = site_labor.get(site_d, 0) + cost
        work_by_id[lookup]  = work_by_id[emp_id]  = hours
        labor_by_id[lookup] = labor_by_id[emp_id] = cost
        days_by_id[lookup]  = days_by_id[emp_id]  = days  # 【追加】出勤日数を保存

    return site_labor, work_by_id, labor_by_id, days_by_id  # 【修正】days_by_idを追加


def _calc_daily(df_apo, date_str, site_map, calls_by_date, ops_by_date, day_num, kono_excluded=True):
    """1営業日分の集計"""
    kono_filter = (df_apo['スタッフ名'] != KONO) if kono_excluded else pd.Series([True]*len(df_apo), index=df_apo.index)
    df_g = df_apo[
        (df_apo['取得日'] == date_str) &
        kono_filter
    ].copy()
    df_g['site_raw'] = df_g['スタッフ名'].map(site_map)

    df_c = df_apo[
        (df_apo['cancel_date_str'] == date_str) &
        kono_filter
    ].copy()
    df_c['site_raw'] = df_c['スタッフ名'].map(site_map)

    # 当日の案件別集計（全体・TOP5）
    pg = df_g.groupby('登録案件名').agg(apo=('アポイントID','count'), sg=('sales','sum')).reset_index()
    pc = df_c.groupby('登録案件名').agg(cxl=('アポイントID','count'), sc=('sales','sum')).reset_index()
    proj = pg.merge(pc, on='登録案件名', how='left').fillna(0)
    proj['valid'] = (proj['apo'] - proj['cxl']).astype(int)
    proj['net']   = (proj['sg']  - proj['sc']).astype(int)
    proj = proj[proj['valid'] > 0].sort_values('valid', ascending=False).head(5)
    top_projects = [
        {'name': str(r['登録案件名']), 'valid': int(r['valid']), 'sales': int(r['net'])}
        for _, r in proj.iterrows()
    ]
    max_valid = top_projects[0]['valid'] if top_projects else 1

    result = {}
    for key, raw in [('all', None), ('shinjuku', '新宿SC'),
                     ('remote', '在宅G'), ('ai', 'AI')]:
        g = df_g if raw is None else df_g[df_g['site_raw'] == raw]
        c = df_c if raw is None else df_c[df_c['site_raw'] == raw]
        row = {
            'day':    f'{day_num}営業日',
            'date':   date_str,
            'sales':  int(g['sales'].sum()) - int(c['sales'].sum()),
            'apo':    len(g),
            'cancel': len(c),
            'valid':  len(g) - len(c),
            'ops':    ops_by_date.get(date_str, 0),
            'calls':  calls_by_date.get(date_str, 0) if key == 'all' else 0,
        }
        # 案件内訳はallのみ付与
        if key == 'all':
            row['top_projects'] = top_projects
            row['max_valid']    = max_valid
        result[key] = row
    return result


def _calc_heatmap(df_apo, biz_dates, total_sales, kono_excluded=True):
    kono_filter = (df_apo['スタッフ名'] != KONO) if kono_excluded else pd.Series([True]*len(df_apo), index=df_apo.index)
    df_g = df_apo[
        df_apo['取得日'].isin(biz_dates) &
        kono_filter
    ].copy()
    df_c = df_apo[
        df_apo['cancel_date_str'].isin(biz_dates) &
        kono_filter
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


def _calc_operators(df_apo, biz_dates, df_master, df_prod,
                    work_by_id, labor_by_id, days_by_id, id_map,
                    site_map, rank_map, inc_map, elapsed, working, kono_excluded=True):
    """OP個人実績"""
    kono_filter_get = (df_apo['スタッフ名'] != KONO) if kono_excluded else pd.Series([True]*len(df_apo), index=df_apo.index)
    df_get = df_apo[
        df_apo['取得日'].isin(biz_dates) &
        kono_filter_get
    ].copy()

    kouryo = (
        df_apo['cancel_date_str'].str.contains('考慮', na=False) &
        df_apo['cancel_date_str'].str.extract(
            r'(\d{4}/\d+/\d+)', expand=False).isin(biz_dates)
    )
    df_cxl_op   = df_apo[df_apo['cancel_date_str'].isin(biz_dates) | kouryo].copy()
    df_cxl_dash = df_apo[df_apo['cancel_date_str'].isin(biz_dates)].copy()

    op_get    = df_get.groupby('スタッフ名').agg(
        apo=('アポイントID', 'count'), sg=('sales', 'sum')).reset_index()
    op_cxl_op = df_cxl_op.groupby('スタッフ名').agg(
        cxl_op=('アポイントID', 'count'), sc_op=('sales', 'sum')).reset_index()
    op_cxl_d  = df_cxl_dash.groupby('スタッフ名').agg(
        cxl_d=('アポイントID', 'count'), sc_d=('sales', 'sum')).reset_index()

    # 生産性レポートのコール集計（日付正規化済み列を使用）
    if '日付_norm' in df_prod.columns and biz_dates:
        month_prefix = biz_dates[0][:7]  # 例：2026/05
        prod_month = df_prod[df_prod['日付_norm'].str.startswith(month_prefix)]
    else:
        prod_month = df_prod
    # コール集計（エージェント列名を柔軟に対応）
    agent_col = None
    for col in ['エージェント', 'エージェント名', 'Agent']:
        if col in prod_month.columns:
            agent_col = col
            break
    if agent_col:
        calls_total = prod_month.groupby(agent_col)['コール数'].sum().to_dict()
    else:
        calls_total = {}

    all_names = set(list(df_master['スタッフ名']) + list(op_get['スタッフ名']))
    operators = []

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
        calls    = int(calls_total.get(name, 0))

        emp_id = id_map.get(name, '')
        lookup = B_TO_D.get(emp_id, emp_id)
        hours  = work_by_id.get(lookup)  or work_by_id.get(emp_id, 0)
        l_base = labor_by_id.get(lookup) or labor_by_id.get(emp_id, 0)
        inc_t  = inc_map.get(name, 0)
        inc_day= round(inc_t / working * elapsed) if inc_t > 0 else 0
        labor  = l_base + inc_day

        # 【修正】出勤日数：days_by_idから取得（なければhours÷8で推定）
        days = days_by_id.get(lookup) or days_by_id.get(emp_id)
        if days is None:
            days = round(float(hours) / 8, 1) if hours and float(hours) > 0 else 0
        else:
            days = float(days)

        ar     = round(apo / calls * 100, 1)    if calls > 0 else None
        cost_r = round(labor / net_op * 100, 1) if net_op > 0 and labor > 0 else None
        site_d = SITE_LABEL.get(site_map.get(name, ''), '')
        rank_d = rank_map.get(name, '')
        unit_pd= round(net_op / days) if days > 0 and net_op > 0 else 0

        # 案件別内訳（有効アポ数TOP10）
        g_proj  = df_get[df_get['スタッフ名'] == name].groupby('登録案件名').agg(
            apo=('アポイントID','count'), sales=('sales','sum')).reset_index()
        c_proj  = df_cxl_dash[df_cxl_dash['スタッフ名'] == name].groupby('登録案件名').agg(
            cxl=('アポイントID','count'), sc=('sales','sum')).reset_index()
        proj_m  = g_proj.merge(c_proj, on='登録案件名', how='left').fillna(0)
        proj_m['valid'] = (proj_m['apo'] - proj_m['cxl']).astype(int)
        proj_m['net']   = (proj_m['sales'] - proj_m['sc']).astype(int)
        proj_m = proj_m[proj_m['valid'] > 0].sort_values('valid', ascending=False).head(10)
        projects = [
            {'name': str(r['登録案件名']), 'valid': int(r['valid']), 'sales': int(r['net'])}
            for _, r in proj_m.iterrows()
        ]

        operators.append({
            'name': name, 'site': site_d, 'rank': rank_d,
            'sales': net_op, 'sales_dash': net_dash,
            'apo': apo, 'cancel': cxl_op, 'valid': apo - cxl_op,
            'calls': calls, 'apoRate': ar,
            'workH': round(float(hours), 1) if hours else None,
            'labor': labor,
            'labor_base': l_base, 'incentive_daily': inc_day,
            'days': days, 'unitPerDay': unit_pd, 'costRate': cost_r,
            'projects': projects,
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
