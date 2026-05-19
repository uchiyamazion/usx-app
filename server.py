"""
USX 単体運転記録表 ローカルサーバー
ダブルクリックで起動 → ブラウザが自動で開く → ボタン1つでExcel出力
"""
import sys, os, json, shutil, threading, webbrowser, sqlite3, uuid
from datetime import datetime
from flask import Flask, request, send_file, send_from_directory, jsonify
from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell

# PyInstallerでexe化した場合のパス解決
BASE = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))

app = Flask(__name__, static_folder=os.path.join(BASE, 'static'))

# ===== データベース設定 =====

DB_PATH = os.path.join(os.path.expanduser('~'), '.usx_app', 'records.db')
os.makedirs(os.path.dirname(DB_PATH), exist_ok=True)

def init_db():
    """データベース初期化"""
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute('''
        CREATE TABLE IF NOT EXISTS records (
            id TEXT PRIMARY KEY,
            name TEXT NOT NULL,
            created_at TEXT NOT NULL,
            data TEXT NOT NULL
        )
    ''')
    conn.commit()
    conn.close()

init_db()

def get_db():
    """DB接続"""
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn

# ===== Excel生成 =====

UNIT_COLS = [
    {"no":"G",  "A":"G",  "B":"H",  "C":"I",  "D":"J"},
    {"no":"K",  "A":"K",  "B":"L",  "C":"M",  "D":"N"},
    {"no":"O",  "A":"O",  "B":"P",  "C":"Q",  "D":"R"},
    {"no":"S",  "A":"S",  "B":"T",  "C":"U",  "D":"V"},
    {"no":"W",  "A":"W",  "B":"X",  "C":"Y",  "D":"Z"},
    {"no":"AA", "A":"AA", "B":"AB", "C":"AC", "D":"AD"},
    {"no":"AE", "A":"AE", "B":"AF", "C":"AG", "D":"AH"},
    {"no":"AI", "A":"AI", "B":"AJ", "C":"AK", "D":"AL"},
    {"no":"AM", "A":"AM", "B":"AN", "C":"AO", "D":"AP"},
    {"no":"AQ", "A":"AQ", "B":"AR", "C":"AS", "D":"AT"},
    {"no":"AU", "A":"AU", "B":"AV", "C":"AW", "D":"AX"},
    {"no":"AY", "A":"AY", "B":"AZ", "C":"BA", "D":"BB"},
]

FIELD_ROWS = [
    ("運転時間",            16, True),
    ("運転回数",            17, True),
    ("冷温水入口圧力",      18, False),
    ("冷温水出口圧力",      19, False),
    ("換算流量",            20, False),
    ("圧縮機電流_R",        21, True),
    ("圧縮機電流_S",        22, True),
    ("圧縮機電流_T",        23, True),
    ("圧縮機電圧_RS",       24, False),
    ("圧縮機電圧_ST",       25, False),
    ("圧縮機電圧_TR",       26, False),
    ("冷温水入口温度",      27, False),
    ("冷温水中間温度",      28, False),
    ("冷温水出口温度",      29, False),
    ("外気温度",            30, False),
    ("冷媒高圧圧力",        31, True),
    ("冷媒低圧圧力",        32, True),
    ("冷媒吐出ガス温度",    33, True),
    ("冷媒吸入ガス温度",    34, True),
    ("冷媒コイルガス温度1", 35, True),
    ("冷媒コイルガス温度2", 36, True),
    ("ファン回転数",        37, True),
    ("圧縮機運転周波数",    38, True),
    ("膨脹弁1開度",         39, True),
    ("膨脹弁2開度",         40, True),
    ("高圧圧力異常",        42, True),
    ("絶縁抵抗",            44, True),
    ("出荷時充填量",        46, True),
    ("初期充填量",          47, True),
    ("総充填量",            48, True),
]

def safe_set(ws, coord, value):
    cell = ws[coord]
    if isinstance(cell, MergedCell):
        for mc in ws.merged_cells.ranges:
            if coord in mc:
                ws.cell(mc.min_row, mc.min_col).value = value
                return
    else:
        cell.value = value

def generate_excel(data: dict) -> str:
    template = os.path.join(BASE, 'template.xlsx')
    out_dir  = os.path.join(os.path.expanduser('~'), 'Desktop', 'USX運転記録')
    os.makedirs(out_dir, exist_ok=True)
    ts = datetime.now().strftime('%Y%m%d_%H%M%S')
    site = data.get('現場名', '').replace('/', '_').replace(' ', '_')[:20]
    out_path = os.path.join(out_dir, f'運転記録表_{site}_{ts}.xlsx')

    shutil.copy2(template, out_path)
    wb = load_workbook(out_path)
    
    # 基本情報（全シートで同じ）
    basic_info = {
        '現場名': 'D2',
        '住所': 'D3',
        '作業日時': 'T3',
        '作業者': 'T4',
        '系統名': 'E5',
        '型式_RUA': 'E6',
        '型式_UP': 'F6',
        '型式_HKZGM': 'I6',
        '運転状態': 'E7',
        '設定温度': 'F7',
        '連結台数': 'T7',
        '設置年月日': 'E8',
        '冷媒種類': 'K8',
        '作成者': 'T70',
    }
    
    # チェック結果セル
    check_cells = {
        '圧縮機関係':       {'異常音':'F50','異常振動':'H50','圧力不良':'K50','オイル量':'N50','異常過熱':'Q50'},
        '凝縮器関係':       {'汚れ':'F51','流量不足':'H51','Yスト詰り':'K51','風量不足':'N51'},
        '蒸発器関係':       {'汚れ':'F52','流量不足':'H52','Yスト詰り':'K52','風量不足':'N52'},
        '送風機関係':       {'異常音':'F53','異常振動':'H53','ショートサイクル':'K53'},
        '冷媒配管系統部品関係': {'ストレーナ':'F54','四方弁':'H54','電磁弁':'K54','膨張弁':'N54'},
        '電装品関係':       {'センサー':'F55','リレー':'H55','トランス':'K55','制御基板':'N55','通信不具合':'Q55'},
        '補機関係':         {'ポンプ':'F56','圧力計':'H56','アキュムレータ':'K56'},
        'その他':           {'その他':'F57'},
    }
    
    # 事前点検セル
    precheck_map = {
        '電装品結線状態確認':    ('F59','H59'),
        '通信線アドレス設定':    ('F60','H60'),
        '熱源機外観点検':        ('F61','H61'),
        '内蔵ポンプ動作確認':    ('F62','H62'),
        '水配管フラッシング':    ('Q59','T59'),
        '12時間前通電':          ('Q60','T60'),
        '冷媒漏洩点検':          ('Q61','T61'),
        'インターロック動作確認':('Q62','T62'),
    }

    units = data.get('熱源機', [])
    
    # 各シートに対応: 1-4号機, 5-8号機, 9-12号機
    sheet_ranges = [
        (0, 4),   # Sheet1: 0-3 (1-4号機)
        (4, 8),   # Sheet2: 4-7 (5-8号機)
        (8, 12),  # Sheet3: 8-11 (9-12号機)
    ]

    for sheet_idx, (start_unit, end_unit) in enumerate(sheet_ranges):
        # シートが存在するかチェック
        if sheet_idx >= len(wb.sheetnames):
            break
        
        ws = wb.worksheets[sheet_idx]

        # ===== 基本情報を書き込み =====
        if data.get('現場名'):     safe_set(ws, 'D2', data['現場名'])
        if data.get('住所'):       safe_set(ws, 'D3', data['住所'])
        if data.get('作業日時'):   safe_set(ws, 'T3', data['作業日時'])
        if data.get('作業者'):     safe_set(ws, 'T4', data['作業者'])
        if data.get('系統名'):     safe_set(ws, 'E5', data['系統名'])
        if data.get('型式_RUA'):   safe_set(ws, 'E6', 'RUA-' + data['型式_RUA'])
        if data.get('型式_UP'):    safe_set(ws, 'F6', data['型式_UP'])
        if data.get('型式_HKZGM'): safe_set(ws, 'I6', data['型式_HKZGM'])
        if data.get('運転状態'):   safe_set(ws, 'E7', data['運転状態'])
        if data.get('設定温度'):   safe_set(ws, 'F7', str(data['設定温度']) + '℃')
        if data.get('連結台数'):   safe_set(ws, 'T7', data['連結台数'])
        if data.get('設置年月日'): safe_set(ws, 'E8', data['設置年月日'])
        if data.get('冷媒種類'):   safe_set(ws, 'K8', data['冷媒種類'])
        if data.get('作成者'):     safe_set(ws, 'T70', data['作成者'])

        for k, cell in {'試運転':'D4','定期点検':'F4','簡易点検':'H4','修理':'J4','整備':'L4','故障判定':'N4'}.items():
            safe_set(ws, cell, ('■' if data.get('作業区分') == k else '□') + k)
        for k, cell in {'定期点検':'Q8','簡易点検':'T8'}.items():
            safe_set(ws, cell, ('■' if data.get('フロン区分') == k else '□') + k)

        # ===== 熱源機データを書き込み =====
        for local_idx, unit_idx in enumerate(range(start_unit, end_unit)):
            if unit_idx >= len(units):
                break
            
            u = units[unit_idx]
            cols = UNIT_COLS[local_idx]
            nc = cols['no']
            
            if u.get('No'):         safe_set(ws, f'{nc}11', u['No'])
            if u.get('UC製造番号'): safe_set(ws, f'{nc}12', u['UC製造番号'])
            if u.get('冷却加熱'):   safe_set(ws, f'{nc}13', u['冷却加熱'])
            if u.get('製造年'):     safe_set(ws, f'{nc}14', u['製造年'])
            if u.get('記録時間'):   safe_set(ws, f'{nc}49', u['記録時間'])

            for field, row, is_circuit in FIELD_ROWS:
                if is_circuit:
                    for c in ['A','B','C','D']:
                        val = u.get(f'{field}_{c}')
                        if val is not None: safe_set(ws, f'{cols[c]}{row}', val)
                else:
                    val = u.get(f'{field}_A') or u.get(field)
                    if val is not None: safe_set(ws, f'{cols["A"]}{row}', val)

        # ===== チェック結果を書き込み =====
        checks = data.get('チェック結果', {})
        for group, items in check_cells.items():
            selected = checks.get(group, [])
            for item, cell in items.items():
                safe_set(ws, cell, ('■' if item in selected else '□') + item)

        # ===== 事前点検を書き込み =====
        prechecks = data.get('試運転事前点検', {})
        for item, (ok_cell, ng_cell) in precheck_map.items():
            result = prechecks.get(item, '')
            safe_set(ws, ok_cell, '■ＯＫ' if result == 'OK' else '□ＯＫ')
            safe_set(ws, ng_cell, '■NG'   if result == 'NG' else '□NG')

        if data.get('備考'): safe_set(ws, 'A71', data['備考'])

    wb.save(out_path)
    return out_path

# ===== ルーティング =====

@app.route('/')
def index():
    return send_from_directory(app.static_folder, 'index.html')

@app.route('/api/records', methods=['GET'])
def get_records():
    """保存済み記録一覧を取得"""
    try:
        conn = get_db()
        c = conn.cursor()
        c.execute('SELECT id, name, created_at FROM records ORDER BY created_at DESC')
        records = [dict(row) for row in c.fetchall()]
        conn.close()
        return jsonify(records)
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/records/<record_id>', methods=['GET'])
def get_record(record_id):
    """特定の記録を取得"""
    try:
        conn = get_db()
        c = conn.cursor()
        c.execute('SELECT id, name, created_at, data FROM records WHERE id = ?', (record_id,))
        row = c.fetchone()
        conn.close()
        if not row:
            return jsonify({'error': 'Record not found'}), 404
        record = dict(row)
        record['data'] = json.loads(record['data'])
        return jsonify(record)
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/records', methods=['POST'])
def save_record():
    """新規記録を保存"""
    try:
        data = request.get_json(force=True)
        name = data.get('name', '').strip()
        record_data = data.get('data', {})
        
        if not name:
            return jsonify({'error': 'Name is required'}), 400
        
        record_id = str(uuid.uuid4())
        created_at = datetime.now().isoformat()
        
        conn = get_db()
        c = conn.cursor()
        c.execute(
            'INSERT INTO records (id, name, created_at, data) VALUES (?, ?, ?, ?)',
            (record_id, name, created_at, json.dumps(record_data))
        )
        conn.commit()
        conn.close()
        
        return jsonify({'id': record_id, 'name': name, 'created_at': created_at}), 201
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/records/<record_id>', methods=['DELETE'])
def delete_record(record_id):
    """記録を削除"""
    try:
        conn = get_db()
        c = conn.cursor()
        c.execute('DELETE FROM records WHERE id = ?', (record_id,))
        conn.commit()
        conn.close()
        return jsonify({'success': True})
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/generate', methods=['POST'])
def generate():
    try:
        data = request.get_json(force=True)
        out_path = generate_excel(data)
        return send_file(
            out_path,
            as_attachment=True,
            download_name=os.path.basename(out_path),
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
    except Exception as e:
        return jsonify({'error': str(e)}), 500

# ===== 起動 =====

def open_browser(port):
    import time; time.sleep(1.2)
    webbrowser.open(f'http://localhost:{port}')

if __name__ == '__main__':
    import socket, os
    port = int(os.environ.get('PORT', 5199))
    is_render = 'RENDER' in os.environ

    print('=' * 50)
    print(' USX 運転記録表サーバー 起動中...')
    print('=' * 50)

    if not is_render:
        try:
            ip = socket.gethostbyname(socket.gethostname())
        except:
            ip = '127.0.0.1'
        print(f' PC用URL:      http://localhost:{port}')
        print(f' スマホ用URL:  http://{ip}:{port}  (同じWi-Fi内)')
        threading.Thread(target=open_browser, args=(port,), daemon=True).start()

    app.run(host='0.0.0.0', port=port, debug=False)
