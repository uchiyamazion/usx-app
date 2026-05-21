import sys, os, json, shutil, threading, webbrowser, uuid, base64
from datetime import datetime
from flask import Flask, request, send_file, send_from_directory, jsonify
from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell
from openpyxl.drawing.image import Image as XLImage
from PIL import Image as PILImage
import io

BASE = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
app = Flask(__name__, static_folder=os.path.join(BASE, 'static'))

# ===== DB (PostgreSQL on Render / SQLite locally) =====
DATABASE_URL = os.environ.get('DATABASE_URL', '')

def get_db():
    if DATABASE_URL:
        import psycopg2
        import psycopg2.extras
        conn = psycopg2.connect(DATABASE_URL)
        return conn, 'pg'
    else:
        import sqlite3
        DB_PATH = os.path.join(os.path.expanduser('~'), '.usx_app', 'records.db')
        os.makedirs(os.path.dirname(DB_PATH), exist_ok=True)
        conn = sqlite3.connect(DB_PATH)
        conn.row_factory = sqlite3.Row
        return conn, 'sqlite'

def init_db():
    conn, db_type = get_db()
    cur = conn.cursor()
    if db_type == 'pg':
        cur.execute('''CREATE TABLE IF NOT EXISTS records (
            id TEXT PRIMARY KEY, name TEXT NOT NULL,
            created_at TEXT NOT NULL, data TEXT NOT NULL)''')
        cur.execute('''CREATE TABLE IF NOT EXISTS photos (
            id TEXT PRIMARY KEY, record_id TEXT,
            filename TEXT NOT NULL, caption TEXT,
            file_data BYTEA,
            created_at TEXT NOT NULL)''')
    else:
        cur.execute('''CREATE TABLE IF NOT EXISTS records (
            id TEXT PRIMARY KEY, name TEXT NOT NULL,
            created_at TEXT NOT NULL, data TEXT NOT NULL)''')
        cur.execute('''CREATE TABLE IF NOT EXISTS photos (
            id TEXT PRIMARY KEY, record_id TEXT,
            filename TEXT NOT NULL, caption TEXT,
            file_data BLOB,
            created_at TEXT NOT NULL)''')
    conn.commit()
    conn.close()

init_db()

def rows_to_dicts(rows, db_type):
    if db_type == 'pg':
        return [dict(zip([d[0] for d in rows.description], row)) for row in rows]
    else:
        return [dict(row) for row in rows]

# 写真リサイズ
def resize_image(img_bytes):
    img = PILImage.open(io.BytesIO(img_bytes))
    img = img.convert('RGB')
    try:
        import PIL.ExifTags
        exif = img._getexif()
        if exif:
            for tag, val in exif.items():
                if PIL.ExifTags.TAGS.get(tag) == 'Orientation':
                    if val == 3:   img = img.rotate(180, expand=True)
                    elif val == 6: img = img.rotate(270, expand=True)
                    elif val == 8: img = img.rotate(90,  expand=True)
    except Exception:
        pass
    img.thumbnail((501, 408), PILImage.LANCZOS)
    for quality in [60, 40, 25]:
        buf = io.BytesIO()
        img.save(buf, format='JPEG', quality=quality, optimize=True)
        if buf.tell() <= 200 * 1024:
            return buf.getvalue()
    return buf.getvalue()

# ===== Excel生成 =====
UNIT_COLS = [
    {"info":"G",  "no_val":"H",  "A":"G",  "B":"H",  "C":"I",  "D":"J"},
    {"info":"K",  "no_val":"L",  "A":"K",  "B":"L",  "C":"M",  "D":"N"},
    {"info":"O",  "no_val":"P",  "A":"O",  "B":"P",  "C":"Q",  "D":"R"},
    {"info":"S",  "no_val":"T",  "A":"S",  "B":"T",  "C":"U",  "D":"V"},
    {"info":"W",  "no_val":"X",  "A":"W",  "B":"X",  "C":"Y",  "D":"Z"},
    {"info":"AA", "no_val":"AB", "A":"AA", "B":"AB", "C":"AC", "D":"AD"},
    {"info":"AE", "no_val":"AF", "A":"AE", "B":"AF", "C":"AG", "D":"AH"},
    {"info":"AI", "no_val":"AJ", "A":"AI", "B":"AJ", "C":"AK", "D":"AL"},
    {"info":"AM", "no_val":"AN", "A":"AM", "B":"AN", "C":"AO", "D":"AP"},
    {"info":"AQ", "no_val":"AR", "A":"AQ", "B":"AR", "C":"AS", "D":"AT"},
    {"info":"AU", "no_val":"AV", "A":"AU", "B":"AV", "C":"AW", "D":"AX"},
    {"info":"AY", "no_val":"AZ", "A":"AY", "B":"AZ", "C":"BA", "D":"BB"},
]
FIELD_ROWS = [
    ("運転時間",16,True),("運転回数",17,True),
    ("冷温水入口圧力",18,False),("冷温水出口圧力",19,False),("換算流量",20,False),
    ("圧縮機電流_R",21,True),("圧縮機電流_S",22,True),("圧縮機電流_T",23,True),
    ("圧縮機電圧_RS",24,False),("圧縮機電圧_ST",25,False),("圧縮機電圧_TR",26,False),
    ("冷温水入口温度",27,False),("冷温水中間温度",28,False),("冷温水出口温度",29,False),
    ("外気温度",30,False),
    ("冷媒高圧圧力",31,True),("冷媒低圧圧力",32,True),
    ("冷媒吐出ガス温度",33,True),("冷媒吸入ガス温度",34,True),
    ("冷媒コイルガス温度1",35,True),("冷媒コイルガス温度2",36,True),
    ("ファン回転数",37,True),("圧縮機運転周波数",38,True),
    ("膨脹弁1開度",39,True),("膨脹弁2開度",40,True),
    ("高圧圧力異常",42,True),("絶縁抵抗",44,True),
    ("出荷時充填量",46,True),("初期充填量",47,True),("総充填量",48,True),
]
CHECK_CELLS = {
    '圧縮機関係':{'異常音':'F50','異常振動':'H50','圧力不良':'K50','オイル量':'N50','異常過熱':'Q50'},
    '凝縮器関係':{'汚れ':'F51','流量不足':'H51','Yスト詰り':'K51','風量不足':'N51'},
    '蒸発器関係':{'汚れ':'F52','流量不足':'H52','Yスト詰り':'K52','風量不足':'N52'},
    '送風機関係':{'異常音':'F53','異常振動':'H53','ショートサイクル':'K53'},
    '冷媒配管系統部品関係':{'ストレーナ':'F54','四方弁':'H54','電磁弁':'K54','膨張弁':'N54'},
    '電装品関係':{'センサー':'F55','リレー':'H55','トランス':'K55','制御基板':'N55','通信不具合':'Q55'},
    '補機関係':{'ポンプ':'F56','圧力計':'H56','アキュムレータ':'K56'},
    'その他':{'その他':'F57'},
}
PRECHECK_MAP = {
    '電装品結線状態確認':('F59','H59'),'通信線アドレス設定':('F60','H60'),
    '熱源機外観点検':('F61','H61'),'内蔵ポンプ動作確認':('F62','H62'),
    '水配管フラッシング':('Q59','T59'),'12時間前通電':('Q60','T60'),
    '冷媒漏洩点検':('Q61','T61'),'インターロック動作確認':('Q62','T62'),
}

def safe_set(ws, coord, value):
    cell = ws[coord]
    if isinstance(cell, MergedCell):
        for mc in ws.merged_cells.ranges:
            if coord in mc:
                ws.cell(mc.min_row, mc.min_col).value = value
                return
    else:
        cell.value = value

def write_sheet(ws, data, unit_start, unit_end):
    if data.get('現場名'):     safe_set(ws,'D2',data['現場名'])
    if data.get('住所'):       safe_set(ws,'D3',data['住所'])
    if data.get('作業日時'):   safe_set(ws,'T3',data['作業日時'].replace('T',' '))
    if data.get('作業者'):     safe_set(ws,'T4',data['作業者'])
    if data.get('系統名'):     safe_set(ws,'E5',data['系統名'])
    if data.get('型式'):       safe_set(ws,'F6',data['型式'])
    if data.get('運転状態'):   safe_set(ws,'E7',data['運転状態'])
    if data.get('設定温度'):   safe_set(ws,'F7',str(data['設定温度'])+'℃')
    if data.get('連結台数'):   safe_set(ws,'T7',data['連結台数'])
    if data.get('設置年月日'): safe_set(ws,'E8',data['設置年月日'])
    if data.get('冷媒種類'):   safe_set(ws,'I8',data['冷媒種類'])
    if data.get('作成者'):     safe_set(ws,'T70',data['作成者'])
    if data.get('備考'):       safe_set(ws,'B71',data['備考'])
    for k,cell in {'試運転':'D4','定期点検':'F4','簡易点検':'H4','修理':'J4','整備':'L4','故障判定':'N4'}.items():
        safe_set(ws,cell,('■' if data.get('作業区分')==k else '□')+k)
    for k,cell in {'定期点検':'Q8','簡易点検':'T8'}.items():
        safe_set(ws,cell,('■' if data.get('フロン区分')==k else '□')+k)
    units=data.get('熱源機',[])
    for li,ui in enumerate(range(unit_start,unit_end)):
        if ui>=len(units): break
        u=units[ui]; cols=UNIT_COLS[li]
        ic=cols['info']; nc=cols['no_val']
        if u.get('No'):         safe_set(ws,f'{nc}11',u['No'])
        if u.get('UC製造番号'): safe_set(ws,f'{ic}12',u['UC製造番号'])
        if u.get('冷却加熱'):   safe_set(ws,f'{ic}13',u['冷却加熱'])
        if u.get('製造年'):     safe_set(ws,f'{ic}14',u['製造年'])
        if u.get('記録時間'):   safe_set(ws,f'{ic}49',u['記録時間'])
        for field,row,is_c in FIELD_ROWS:
            if is_c:
                for c in ['A','B','C','D']:
                    val=u.get(f'{field}_{c}')
                    if val is not None: safe_set(ws,f'{cols[c]}{row}',val)
            else:
                val=u.get(f'{field}_A') or u.get(field)
                if val is not None: safe_set(ws,f'{cols["A"]}{row}',val)
    checks=data.get('チェック結果',{})
    for group,items in CHECK_CELLS.items():
        sel=checks.get(group,[])
        for item,cell in items.items():
            safe_set(ws,cell,('■' if item in sel else '□')+item)
    pre=data.get('試運転事前点検',{})
    for item,(ok,ng) in PRECHECK_MAP.items():
        r=pre.get(item,'')
        safe_set(ws,ok,'■ＯＫ' if r=='OK' else '□ＯＫ')
        safe_set(ws,ng,'■NG' if r=='NG' else '□NG')

def make_excel(data, photos=None):
    template=os.path.join(BASE,'template.xlsx')
    out_dir=os.path.join(os.path.expanduser('~'),'Desktop','USX運転記録')
    os.makedirs(out_dir,exist_ok=True)
    ts=datetime.now().strftime('%Y%m%d_%H%M%S')
    site=data.get('現場名','').replace('/','_').replace(' ','_')[:20]
    out_path=os.path.join(out_dir,f'運転記録表_{site}_{ts}.xlsx')
    shutil.copy2(template,out_path)
    wb=load_workbook(out_path)
    while len(wb.worksheets)<3:
        wb.copy_worksheet(wb.worksheets[0])
    for si,(s,e) in enumerate([(0,4),(4,8),(8,12)]):
        if si<len(wb.worksheets): write_sheet(wb.worksheets[si],data,s,e)

    # 写真シート追加
    if photos:
        photo_tmpl=os.path.join(BASE,'photo_template.xlsx')
        if os.path.exists(photo_tmpl):
            wb_p=load_workbook(photo_tmpl)
            ppr_sheet=[(3,23),(25,45),(47,67)]
            info_rows=[
                {'date':4,'content':7},
                {'date':26,'content':29},
                {'date':48,'content':51},
            ]
            sheet_idx=0
            for i,photo in enumerate(photos):
                li=i%3
                if li==0 and i>0: sheet_idx+=1
                if sheet_idx>=len(wb_p.worksheets): break
                ws_p=wb_p.worksheets[sheet_idx]
                pr=ppr_sheet[li]; ir=info_rows[li]
                if photo.get('file_data'):
                    try:
                        img_bytes=bytes(photo['file_data']) if not isinstance(photo['file_data'],bytes) else photo['file_data']
                        xl_img=XLImage(io.BytesIO(img_bytes))
                        xl_img.width=501; xl_img.height=408
                        ws_p.add_image(xl_img,f'B{pr[0]}')
                    except Exception as ex:
                        print('img error:',ex)
                dt=photo.get('created_at','')[:10]
                cap=photo.get('caption','')
                ws_p.cell(ir['date'],8,dt)
                ws_p.cell(ir['content'],8,cap)
            for ws_p in wb_p.worksheets:
                ws_new=wb.create_sheet(title=ws_p.title[:31])
                for row in ws_p.iter_rows():
                    for cell in row:
                        nc=ws_new.cell(cell.row,cell.column,cell.value)
                        if cell.has_style:
                            nc.font=cell.font.copy()
                            nc.border=cell.border.copy()
                            nc.fill=cell.fill.copy()
                            nc.alignment=cell.alignment.copy()
                for mc in ws_p.merged_cells.ranges:
                    ws_new.merge_cells(str(mc))
                for img in ws_p._images:
                    ws_new.add_image(img)
                for col in ws_p.column_dimensions:
                    ws_new.column_dimensions[col].width=ws_p.column_dimensions[col].width
                for row in ws_p.row_dimensions:
                    ws_new.row_dimensions[row].height=ws_p.row_dimensions[row].height

    wb.save(out_path)
    return out_path

# ===== API =====
@app.route('/')
def index():
    return send_from_directory(app.static_folder,'index.html')

@app.route('/api/records',methods=['GET'])
def get_records():
    try:
        conn,db_type=get_db()
        cur=conn.cursor()
        cur.execute('SELECT id,name,created_at FROM records ORDER BY created_at DESC')
        rows=cur.fetchall()
        conn.close()
        if db_type=='pg':
            cols=[d[0] for d in cur.description]
            return jsonify([dict(zip(cols,r)) for r in rows])
        return jsonify([dict(r) for r in rows])
    except Exception as e:
        return jsonify({'error':str(e)}),500

@app.route('/api/records/<rid>',methods=['GET'])
def get_record(rid):
    try:
        conn,db_type=get_db()
        cur=conn.cursor()
        cur.execute('SELECT id,name,created_at,data FROM records WHERE id=%s' if db_type=='pg' else
                    'SELECT id,name,created_at,data FROM records WHERE id=?',(rid,))
        row=cur.fetchone(); conn.close()
        if not row: return jsonify({'error':'Not found'}),404
        if db_type=='pg':
            cols=[d[0] for d in cur.description]
            r=dict(zip(cols,row))
        else:
            r=dict(row)
        r['data']=json.loads(r['data'])
        return jsonify(r)
    except Exception as e:
        return jsonify({'error':str(e)}),500

@app.route('/api/records',methods=['POST'])
def save_record():
    try:
        body=request.get_json(force=True)
        name=body.get('name','').strip()
        if not name: return jsonify({'error':'Name required'}),400
        rid=str(uuid.uuid4())
        now=datetime.now().isoformat()
        conn,db_type=get_db()
        ph='%s' if db_type=='pg' else '?'
        conn.cursor().execute(
            f'INSERT INTO records (id,name,created_at,data) VALUES ({ph},{ph},{ph},{ph})',
            (rid,name,now,json.dumps(body.get('data',{}))))
        conn.commit(); conn.close()
        return jsonify({'id':rid,'name':name,'created_at':now}),201
    except Exception as e:
        return jsonify({'error':str(e)}),500

@app.route('/api/records/<rid>',methods=['DELETE'])
def delete_record(rid):
    try:
        conn,db_type=get_db()
        ph='%s' if db_type=='pg' else '?'
        conn.cursor().execute(f'DELETE FROM records WHERE id={ph}',(rid,))
        conn.cursor().execute(f'DELETE FROM photos WHERE record_id={ph}',(rid,))
        conn.commit(); conn.close()
        return jsonify({'success':True})
    except Exception as e:
        return jsonify({'error':str(e)}),500

@app.route('/api/photos/<rid>',methods=['GET'])
def get_photos(rid):
    try:
        conn,db_type=get_db()
        ph='%s' if db_type=='pg' else '?'
        cur=conn.cursor()
        cur.execute(f'SELECT id,caption,created_at FROM photos WHERE record_id={ph} ORDER BY created_at',(rid,))
        rows=cur.fetchall(); conn.close()
        if db_type=='pg':
            cols=[d[0] for d in cur.description]
            return jsonify([dict(zip(cols,r)) for r in rows])
        return jsonify([dict(r) for r in rows])
    except Exception as e:
        return jsonify({'error':str(e)}),500

@app.route('/api/photos/<rid>',methods=['POST'])
def upload_photo(rid):
    try:
        # record存在確認
        conn,db_type=get_db()
        ph='%s' if db_type=='pg' else '?'
        cur=conn.cursor()
        cur.execute(f'SELECT id FROM records WHERE id={ph}',(rid,))
        if not cur.fetchone():
            conn.close()
            return jsonify({'error':'Record not found'}),404
        body=request.get_json(force=True)
        img_b64=body.get('image','')
        caption=body.get('caption','')
        if not img_b64: return jsonify({'error':'No image'}),400
        if ',' in img_b64: img_b64=img_b64.split(',',1)[1]
        img_bytes=resize_image(base64.b64decode(img_b64))
        pid=str(uuid.uuid4())
        now=datetime.now().isoformat()
        if db_type=='pg':
            import psycopg2.extras
            cur.execute(
                'INSERT INTO photos (id,record_id,filename,caption,file_data,created_at) VALUES (%s,%s,%s,%s,%s,%s)',
                (pid,rid,f'{pid}.jpg',caption,psycopg2.Binary(img_bytes),now))
        else:
            cur.execute(
                'INSERT INTO photos (id,record_id,filename,caption,file_data,created_at) VALUES (?,?,?,?,?,?)',
                (pid,rid,f'{pid}.jpg',caption,img_bytes,now))
        conn.commit(); conn.close()
        return jsonify({'id':pid,'caption':caption,'created_at':now}),201
    except Exception as e:
        return jsonify({'error':str(e)}),500

@app.route('/api/photos/item/<pid>',methods=['DELETE'])
def delete_photo(pid):
    try:
        conn,db_type=get_db()
        ph='%s' if db_type=='pg' else '?'
        conn.cursor().execute(f'DELETE FROM photos WHERE id={ph}',(pid,))
        conn.commit(); conn.close()
        return jsonify({'success':True})
    except Exception as e:
        return jsonify({'error':str(e)}),500

@app.route('/api/photos/file/<pid>',methods=['GET'])
def get_photo_file(pid):
    try:
        conn,db_type=get_db()
        ph='%s' if db_type=='pg' else '?'
        cur=conn.cursor()
        cur.execute(f'SELECT file_data FROM photos WHERE id={ph}',(pid,))
        row=cur.fetchone(); conn.close()
        if not row: return jsonify({'error':'Not found'}),404
        data=row[0] if db_type=='pg' else row['file_data']
        if data is None: return jsonify({'error':'No data'}),404
        return send_file(io.BytesIO(bytes(data)),mimetype='image/jpeg')
    except Exception as e:
        return jsonify({'error':str(e)}),500

@app.route('/generate',methods=['POST'])
def generate():
    try:
        data=request.get_json(force=True)
        out=make_excel(data)
        return send_file(out,as_attachment=True,download_name=os.path.basename(out),
                         mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    except Exception as e:
        return jsonify({'error':str(e)}),500

@app.route('/generate_with_photos',methods=['POST'])
def generate_with_photos():
    try:
        body=request.get_json(force=True)
        data=body.get('data',{})
        rid=body.get('record_id','')
        photos=[]
        if rid:
            conn,db_type=get_db()
            ph='%s' if db_type=='pg' else '?'
            cur=conn.cursor()
            cur.execute(f'SELECT id,caption,created_at,file_data FROM photos WHERE record_id={ph} ORDER BY created_at',(rid,))
            rows=cur.fetchall(); conn.close()
            if db_type=='pg':
                cols=[d[0] for d in cur.description]
                photos=[dict(zip(cols,r)) for r in rows]
            else:
                photos=[dict(r) for r in rows]
        out=make_excel(data,photos)
        return send_file(out,as_attachment=True,download_name=os.path.basename(out),
                         mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    except Exception as e:
        return jsonify({'error':str(e)}),500

# ===== 起動 =====
def open_browser(port):
    import time; time.sleep(1.2)
    webbrowser.open(f'http://localhost:{port}')

if __name__=='__main__':
    port=int(os.environ.get('PORT',5199))
    is_render='RENDER' in os.environ
    print('='*50)
    print(' USX 運転記録表サーバー 起動中...')
    print('='*50)
    if not is_render:
        import socket
        try: ip=socket.gethostbyname(socket.gethostname())
        except: ip='127.0.0.1'
        print(f' PC用:     http://localhost:{port}')
        print(f' スマホ用: http://{ip}:{port}')
        threading.Thread(target=open_browser,args=(port,),daemon=True).start()
    app.run(host='0.0.0.0',port=port,debug=False)
