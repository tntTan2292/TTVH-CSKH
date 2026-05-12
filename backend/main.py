import os
import sqlite3
import pandas as pd
import unicodedata
import json
import traceback
import sys
from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import HTMLResponse, JSONResponse
from datetime import datetime, timedelta
import uvicorn

# CONFIGURATION
BASE_DIR = r"d:\Antigravity - Project - TTVH\CSKH"
DB_PATH = os.path.join(BASE_DIR, "cskh_vip.db")
UPLOAD_DIR = os.path.join(BASE_DIR, "backend", "temp_uploads")
TEMPLATE_PATH = os.path.join(BASE_DIR, "backend", "report_template.md")
os.makedirs(UPLOAD_DIR, exist_ok=True)

# GEOGRAPHIC MAPPING
NORTH_PROVINCES = ["Hà Nội", "Hải Phòng", "Bắc Ninh", "Thái Nguyên", "Vĩnh Phúc", "Hải Dương", "Quảng Ninh", "Ninh Bình", "Nam Định", "Hà Nam", "Hòa Bình", "Sơn La", "Điện Biên", "Lai Châu", "Lào Cai", "Yên Bái", "Phú Thọ", "Bắc Giang", "Lạng Sơn", "Tuyên Quang", "Hà Giang", "Cao Bằng", "Bắc Kạn"]

def log_terminal(msg):
    print(f"[{datetime.now().strftime('%H:%M:%S')}] {msg}")
    sys.stdout.flush()

def clean_text_nfc(text):
    if not text or pd.isna(text): return ""
    return unicodedata.normalize('NFC', str(text)).strip()

def clean_text_lower(text):
    if not text or pd.isna(text): return ""
    return unicodedata.normalize('NFC', str(text)).lower().strip()

def parse_dt(val):
    if pd.isna(val) or val == "": return None
    if isinstance(val, datetime): return val
    if isinstance(val, (int, float)):
        try: return datetime(1899, 12, 30) + timedelta(days=val)
        except: return None
    if isinstance(val, str):
        val = val.strip()
        for fmt in ('%Y-%m-%d %H:%M:%S', '%d/%m/%Y %H:%M:%S', '%Y-%m-%d', '%d/%m/%Y'):
            try: return datetime.strptime(val, fmt)
            except: continue
    return None

def find_best_col(headers, target_keywords, exclude_keywords=[]):
    headers_lower = [clean_text_lower(h) for h in headers]
    for i, h in enumerate(headers_lower):
        if h in target_keywords: return i
    for i, h in enumerate(headers_lower):
        if any(tk in h for tk in target_keywords):
            if not any(ek in h for ek in exclude_keywords): return i
    return None

app = FastAPI()
app.add_middleware(CORSMiddleware, allow_origins=["*"], allow_methods=["*"], allow_headers=["*"])

def get_db_conn():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    conn = get_db_conn()
    cursor = conn.cursor()
    # CLEAN START FOR V3.2
    cursor.execute("DROP TABLE IF EXISTS orders")
    cursor.execute("DROP TABLE IF EXISTS import_sessions")
    cursor.execute('''CREATE TABLE import_sessions (session_id INTEGER PRIMARY KEY AUTOINCREMENT, filename TEXT, imported_at DATETIME, total_rows INTEGER, status TEXT)''')
    cursor.execute('''CREATE TABLE orders (
        tracking_id TEXT, province TEXT, post_office_name TEXT, acceptance_date TEXT, 
        ttp_first_date TEXT, result_first TEXT, result_final TEXT, status TEXT, 
        aging INTEGER, is_sla_violation BOOLEAN, session_id INTEGER)''')
    conn.commit()
    conn.close()

init_db()

@app.post("/upload")
async def upload_excel(file: UploadFile = File(...)):
    log_terminal(f"--- [DATABASE RECOVERY] START: {file.filename} ---")
    file_path = os.path.join(UPLOAD_DIR, f"recov_{datetime.now().strftime('%H%M%S')}_{file.filename}")
    try:
        content = await file.read()
        with open(file_path, "wb") as f: f.write(content)
        with pd.ExcelFile(file_path, engine='openpyxl') as xls:
            target_sheet = 'DanhSach' if 'DanhSach' in xls.sheet_names else xls.sheet_names[0]
            df_raw_10 = pd.read_excel(xls, sheet_name=target_sheet, header=None, nrows=10)
            header_idx = 0
            for i, row in df_raw_10.iterrows():
                row_str = " ".join([clean_text_lower(val) for val in row if not pd.isna(val)])
                if any(k in row_str for k in ["số hiệu", "mã bưu gửi", "kết quả"]):
                    header_idx = i
                    break
            df = pd.read_excel(xls, sheet_name=target_sheet, header=header_idx)
            df.columns = [clean_text_nfc(h) for h in df.columns]

        mapping = {
            "tracking_id": find_best_col(df.columns, ["số hiệu", "mã bưu gửi"]),
            "acceptance_date": find_best_col(df.columns, ["ngày chấp nhận", "ngày gửi"]),
            "ttp_first_date": find_best_col(df.columns, ["thời gian nhập ttp lần đầu", "thời gian nhập lần đầu"]),
            "result_first": find_best_col(df.columns, ["kết quả phát lần đầu", "kết quả lần đầu"]),
            "result_final": find_best_col(df.columns, ["kết quả phát cuối cùng", "kết quả phát lần cuối"]),
            "province": find_best_col(df.columns, ["tỉnh"], ["mã"]),
            "post_office": find_best_col(df.columns, ["bcvh", "tên bcvh", "bưu cục vận hành"])
        }

        df = df.drop_duplicates(subset=[df.columns[mapping["tracking_id"]]], keep='last')
        now = datetime.now()
        conn = get_db_conn()
        cursor = conn.cursor()
        try:
            cursor.execute("BEGIN TRANSACTION")
            cursor.execute('INSERT INTO import_sessions (filename, imported_at, total_rows, status) VALUES (?, ?, ?, ?)', (file.filename, now.isoformat(), len(df), "SUCCESS"))
            sid = cursor.lastrowid
            
            processed = []
            for _, row in df.iterrows():
                tid = str(row.iloc[mapping["tracking_id"]]).strip()
                if not tid: continue
                dt_acc = parse_dt(row.iloc[mapping["acceptance_date"]])
                dt_ttp = parse_dt(row.iloc[mapping["ttp_first_date"]]) if mapping["ttp_first_date"] is not None else None
                res_first = clean_text_lower(row.iloc[mapping["result_first"]])
                res_final = clean_text_lower(row.iloc[mapping["result_final"]])
                
                is_sla = False
                if dt_acc and dt_ttp and (dt_ttp - dt_acc).days > 3: is_sla = True
                if not is_sla and dt_acc and "chưa có tt phát" in res_first and (now - dt_acc).days > 3: is_sla = True
                
                is_success = "đã phát thành công" in res_final
                aging = (now - dt_acc).days if dt_acc else 0
                
                processed.append((
                    tid, clean_text_nfc(row.iloc[mapping["province"]]), clean_text_nfc(row.iloc[mapping["post_office"]]),
                    str(dt_acc) if dt_acc else "", str(dt_ttp) if dt_ttp else "",
                    res_first, res_final, 'Thành công' if is_success else 'Chưa thành công',
                    int(aging), is_sla, sid
                ))

            cursor.executemany('INSERT INTO orders VALUES (?,?,?,?,?,?,?,?,?,?,?)', processed)
            cursor.execute("COMMIT")
            log_terminal(f"RECOVERY SUCCESS: {len(processed)} orders saved with SID {sid}")
            return {"message": "Success", "sid": sid}
        finally: conn.close()
    except Exception as e:
        log_terminal(f"ERR: {str(e)}")
        return JSONResponse(status_code=500, content={"detail": str(e)})
    finally:
        if os.path.exists(file_path): os.remove(file_path)

@app.get("/api/dashboard/stats")
async def get_stats():
    conn = get_db_conn()
    cursor = conn.cursor()
    session = cursor.execute("SELECT session_id FROM import_sessions ORDER BY session_id DESC LIMIT 1").fetchone()
    if not session: return {"error": "No Data"}
    sid = session['session_id']
    
    kpis = cursor.execute('SELECT COUNT(*) as t, SUM(CASE WHEN status="Thành công" THEN 1 ELSE 0 END) as s, SUM(CASE WHEN is_sla_violation=1 THEN 1 ELSE 0 END) as sla FROM orders WHERE session_id=?', (sid,)).fetchone()
    intra = cursor.execute("SELECT COUNT(*) as t, SUM(CASE WHEN status='Thành công' THEN 1 ELSE 0 END) as s FROM orders WHERE session_id=? AND (province LIKE '%Huế%' OR province LIKE '%Hue%')", (sid,)).fetchone()
    north = cursor.execute('''
        SELECT COUNT(*) as t, SUM(CASE WHEN status="Thành công" THEN 1 ELSE 0 END) as s 
        FROM orders WHERE session_id=? AND NOT (province LIKE '%Huế%' OR province LIKE '%Hue%') 
        AND (province LIKE '%Hà Nội%' OR province LIKE '%Bắc Ninh%' OR province LIKE '%Hải Phòng%' OR province LIKE '%Thái Nguyên%' OR province LIKE '%Quảng Ninh%' OR province LIKE '%Hải Dương%' OR province LIKE '%Bắc Giang%' OR province LIKE '%Lạng Sơn%')
    ''', (sid,)).fetchone()
    
    conn.close()
    t, s = kpis['t'], kpis['s']
    it, isuc = intra['t'] or 0, intra['s'] or 0
    nt, nsuc = north['t'] or 0, north['s'] or 0
    
    return {
        "kpis": {"total": t, "success": s, "pending": t - s, "sla": kpis['sla']},
        "directions": {
            "intra": {"total": it, "success": isuc},
            "north": {"total": nt, "success": nsuc},
            "south": {"total": t - it - nt, "success": s - isuc - nsuc}
        }
    }

@app.get("/api/dashboard/province-performance")
async def get_province():
    conn = get_db_conn()
    cursor = conn.cursor()
    session = cursor.execute("SELECT session_id FROM import_sessions ORDER BY session_id DESC LIMIT 1").fetchone()
    if not session: return []
    sid = session['session_id']
    rows = cursor.execute('SELECT province, COUNT(*) as ct, SUM(CASE WHEN status="Thành công" THEN 1 ELSE 0 END) as su, SUM(CASE WHEN is_sla_violation=1 THEN 1 ELSE 0 END) as sla FROM orders WHERE session_id=? AND province!="" GROUP BY province ORDER BY ct DESC LIMIT 10', (sid,)).fetchall()
    conn.close()
    return [{"name": r['province'], "direction": "Nội tỉnh" if "Huế" in r['province'] else ("Bắc" if any(n in r['province'] for n in NORTH_PROVINCES) else "Nam"), "total": r['ct'], "success": r['su'], "sla": r['sla'], "success_rate": round(r['su']*100/r['ct']) if r['ct']>0 else 0, "sla_rate": round(r['sla']*100/r['ct']) if r['ct']>0 else 0} for r in rows]

@app.get("/api/dashboard/bcvh-summary")
async def get_bcvh():
    conn = get_db_conn()
    cursor = conn.cursor()
    session = cursor.execute("SELECT session_id FROM import_sessions ORDER BY session_id DESC LIMIT 1").fetchone()
    if not session: return []
    sid = session['session_id']
    rows = cursor.execute('SELECT post_office_name, province, COUNT(*) as ct, SUM(CASE WHEN status="Thành công" THEN 1 ELSE 0 END) as su, SUM(CASE WHEN is_sla_violation=1 THEN 1 ELSE 0 END) as sla FROM orders WHERE session_id=? GROUP BY post_office_name, province ORDER BY ct DESC', (sid,)).fetchall()
    conn.close()
    return [{"name": r['post_office_name'], "province": r['province'], "total": r['ct'], "success": r['su'], "sla": r['sla'], "rate": round(r['su']*100/r['ct']) if r['ct']>0 else 0} for r in rows]

@app.get("/api/dashboard/sla-risk")
async def get_sla():
    conn = get_db_conn()
    cursor = conn.cursor()
    session = cursor.execute("SELECT session_id FROM import_sessions ORDER BY session_id DESC LIMIT 1").fetchone()
    if not session: return []
    sid = session['session_id']
    rows = cursor.execute('SELECT tracking_id, aging, province, post_office_name FROM orders WHERE session_id=? AND is_sla_violation=1 ORDER BY aging DESC LIMIT 50', (sid,)).fetchall()
    conn.close()
    return [dict(r) for r in rows]

@app.get("/api/dashboard/generate-report")
async def generate_report():
    # Reuse stats logic for report
    conn = get_db_conn()
    cursor = conn.cursor()
    session = cursor.execute("SELECT session_id FROM import_sessions ORDER BY session_id DESC LIMIT 1").fetchone()
    if not session: return {"error": "No Data"}
    sid = session['session_id']
    
    kpis = cursor.execute('SELECT COUNT(*) as t, SUM(CASE WHEN status="Thành công" THEN 1 ELSE 0 END) as s, SUM(CASE WHEN result_first LIKE "%chưa có tt phát%" THEN 1 ELSE 0 END) as p FROM orders WHERE session_id=?', (sid,)).fetchone()
    trends = cursor.execute("SELECT SUBSTR(acceptance_date, 1, 10) as d, COUNT(*) as ct, SUM(CASE WHEN status='Thành công' THEN 1 ELSE 0 END) as su FROM orders WHERE session_id=? GROUP BY d ORDER BY d ASC", (sid,)).fetchall()
    provinces = cursor.execute("SELECT province, COUNT(*) as ct, SUM(CASE WHEN status='Thành công' THEN 1 ELSE 0 END) as su FROM orders WHERE session_id=? GROUP BY province ORDER BY su*1.0/ct DESC LIMIT 3", (sid,)).fetchall()
    
    t, s, p = kpis['t'], kpis['s'], kpis['p']
    failed = t - s - p
    trend_text = ", ".join([f"{r['d'][-5:]}: {round(r['su']*100/r['ct'])}%" for r in trends])
    top_provinces = " và ".join([r['province'] for r in provinces])
    
    ctx = {
        "date_now": datetime.now().strftime("%Hh%M ngày %d/%m/%Y"),
        "total": t, "success_count": s, "success_rate": round(s*100/t),
        "pending_count": p, "pending_rate": round(p*100/t),
        "failed_count": failed, "failed_rate": round(failed*100/t),
        "trend_text": trend_text, "top_provinces": top_provinces,
        "n_rate": 0, "s_rate": 0, "i_rate": 0, "s_ct": 0, "risk_0705": 0 # Simpler for now
    }
    
    with open(TEMPLATE_PATH, "r", encoding="utf-8") as f:
        return {"report": f.read().format(**ctx)}

@app.get("/", response_class=HTMLResponse)
async def get_index():
    with open(os.path.join(BASE_DIR, "backend", "index.html"), "r", encoding="utf-8") as f: return f.read()

if __name__ == "__main__":
    uvicorn.run(app, host="0.0.0.0", port=8010)
