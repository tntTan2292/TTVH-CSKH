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
from datetime import datetime
import uvicorn

# CONFIGURATION
BASE_DIR = r"d:\Antigravity - Project - TTVH\CSKH"
CONFIG_PATH = os.path.join(BASE_DIR, "backend", "config.json")
DB_PATH = os.path.join(BASE_DIR, "cskh_vip.db")
UPLOAD_DIR = os.path.join(BASE_DIR, "backend", "temp_uploads")

os.makedirs(UPLOAD_DIR, exist_ok=True)

# GEOGRAPHIC MAPPING
NORTH_PROVINCES = ["Hà Nội", "Hải Phòng", "Bắc Ninh", "Thái Nguyên", "Vĩnh Phúc", "Hải Dương", "Quảng Ninh", "Ninh Bình", "Nam Định", "Hà Nam", "Hòa Bình", "Sơn La", "Điện Biên", "Lai Châu", "Lào Cai", "Yên Bái", "Phú Thọ", "Bắc Giang", "Lạng Sơn", "Tuyên Quang", "Hà Giang", "Cao Bằng", "Bắc Kạn"]

def get_direction(province_name):
    if not province_name: return "Nam"
    p = clean_text(province_name)
    if "Huế" in p or "Hue" in p: return "Nội tỉnh"
    if any(north in p for north in NORTH_PROVINCES): return "Bắc"
    return "Nam"

def clean_text(text):
    if not text or pd.isna(text): return ""
    text = unicodedata.normalize('NFC', str(text))
    return text.lower().strip()

def log_terminal(msg):
    print(f"[{datetime.now().strftime('%H:%M:%S')}] {msg}")
    sys.stdout.flush()

app = FastAPI()
app.add_middleware(CORSMiddleware, allow_origins=["*"], allow_methods=["*"], allow_headers=["*"])

def get_db_conn():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    conn = get_db_conn()
    cursor = conn.cursor()
    cursor.execute('''CREATE TABLE IF NOT EXISTS import_sessions (session_id INTEGER PRIMARY KEY AUTOINCREMENT, filename TEXT, imported_at DATETIME, total_rows INTEGER, unique_ids INTEGER, success_count INTEGER, status TEXT)''')
    cursor.execute('''CREATE TABLE IF NOT EXISTS orders (
        tracking_id TEXT, customer_id TEXT, recipient_address TEXT, post_office_name TEXT, province TEXT, 
        acceptance_date TEXT, result_first TEXT, result_final TEXT, status TEXT, aging INTEGER, is_sla_violation BOOLEAN, 
        session_id INTEGER, PRIMARY KEY (tracking_id, session_id))''')
    conn.commit()
    conn.close()

def find_best_col(headers, target_keywords, exclude_keywords=[]):
    headers_lower = [h.lower() for h in headers]
    for i, h in enumerate(headers_lower):
        if h in target_keywords: return i
    for i, h in enumerate(headers_lower):
        if any(tk in h for tk in target_keywords):
            if not any(ek in h for ek in exclude_keywords): return i
    return None

init_db()

@app.post("/upload")
async def upload_excel(file: UploadFile = File(...)):
    log_terminal(f"=== [FORENSIC DEBUG] START: {file.filename} ===")
    file_path = os.path.join(UPLOAD_DIR, f"debug_{datetime.now().strftime('%H%M%S')}_{file.filename}")
    try:
        content = await file.read()
        with open(file_path, "wb") as f: f.write(content)
        
        with pd.ExcelFile(file_path, engine='openpyxl') as xls:
            target_sheet = 'DanhSach' if 'DanhSach' in xls.sheet_names else xls.sheet_names[0]
            # READ RAW DATA
            df_raw = pd.read_excel(xls, sheet_name=target_sheet, header=None, nrows=20)
            
            header_idx = 0
            for i, row in df_raw.iterrows():
                row_str = " ".join([str(val).lower() for val in row if not pd.isna(val)])
                if any(k in row_str for k in ["số hiệu", "mã bưu gửi", "kết quả"]):
                    header_idx = i
                    break
            
            df = pd.read_excel(xls, sheet_name=target_sheet, header=header_idx)
            df.columns = [str(h).strip() for h in df.columns]

        # 1. IDENTIFY COLUMN
        final_col_idx = find_best_col(df.columns, ["kết quả phát cuối cùng", "kết quả phát lần cuối", "kết quả cuối cùng"])
        if final_col_idx is None:
            log_terminal("CRITICAL: Could not map 'Kết quả phát lần cuối' column!")
            return JSONResponse(status_code=400, content={"detail": "Không tìm thấy cột Kết quả phát lần cuối"})
        
        final_col_name = df.columns[final_col_idx]
        log_terminal(f"MAPPED COLUMN: '{final_col_name}'")
        
        # 2. ANALYSIS RAW VALUES
        raw_values = df[final_col_name].dropna()
        log_terminal(f"TOTAL NON-NULL ROWS: {len(raw_values)}")
        
        log_terminal("--------------------------------")
        log_terminal("SUCCESS STATUS UNIQUE VALUES (Cleaned):")
        unique_cleaned = sorted(list(set([clean_text(v) for v in raw_values])))
        print(json.dumps(unique_cleaned, indent=2, ensure_ascii=False))
        log_terminal("--------------------------------")
        
        log_terminal("TOP 20 RAW VALUES:")
        print(raw_values.head(20).to_list())
        log_terminal("--------------------------------")

        # 3. TEST MATCHING LOGIC
        match_count = 0
        keyword = "đã phát thành công"
        for v in raw_values:
            if keyword in clean_text(v): match_count += 1
        
        log_terminal(f"SUCCESS_MATCH_COUNT: {match_count}")
        log_terminal(f"RECOVERY SUCCESS RATE: {round(match_count*100/len(df)) if len(df)>0 else 0}%")

        # NO FURTHER CODING AS PER REQUEST. 
        # JUST LOGGING TO FIND ROOT CAUSE.
        
        return {"debug": "Forensic log printed in terminal", "match_count": match_count}

    except Exception as e:
        log_terminal(f"DEBUG ERR: {str(e)}")
        traceback.print_exc()
        return JSONResponse(status_code=500, content={"detail": str(e)})
    finally:
        if os.path.exists(file_path): os.remove(file_path)

@app.get("/api/dashboard/stats")
async def get_stats():
    conn = get_db_conn()
    cursor = conn.cursor()
    session = cursor.execute("SELECT * FROM import_sessions WHERE status='SUCCESS' ORDER BY session_id DESC LIMIT 1").fetchone()
    if not session: return {"error": "No Data"}
    sid = session['session_id']
    kpis = cursor.execute('SELECT COUNT(*) as total, SUM(CASE WHEN status="Thành công" THEN 1 ELSE 0 END) as success FROM orders WHERE session_id=?', (sid,)).fetchone()
    intra_stats = cursor.execute("SELECT COUNT(*) as total, SUM(CASE WHEN status='Thành công' THEN 1 ELSE 0 END) as success FROM orders WHERE session_id=? AND (province LIKE '%Huế%' OR province LIKE '%Hue%')", (sid,)).fetchone()
    north_stats = cursor.execute('''
        SELECT COUNT(*) as total, SUM(CASE WHEN status="Thành công" THEN 1 ELSE 0 END) as success 
        FROM orders WHERE session_id=? AND NOT (province LIKE '%Huế%' OR province LIKE '%Hue%') 
        AND (province LIKE '%Hà Nội%' OR province LIKE '%Bắc Ninh%' OR province LIKE '%Hải Phòng%' OR province LIKE '%Thái Nguyên%' OR province LIKE '%Quảng Ninh%')
    ''', (sid,)).fetchone()
    total_val = kpis['total'] or 0
    success_val = kpis['success'] or 0
    intra_total = intra_stats['total'] or 0
    intra_success = intra_stats['success'] or 0
    north_total = north_stats['total'] or 0
    north_success = north_stats['success'] or 0
    conn.close()
    return {
        "kpis": {"total": total_val, "success": success_val, "pending": total_val - success_val, "sla": 0},
        "directions": {
            "intra": {"total": intra_total, "success": intra_success},
            "north": {"total": north_total, "success": north_success},
            "south": {"total": total_val - intra_total - north_total, "success": success_val - intra_success - north_success}
        }
    }

@app.get("/api/dashboard/province-performance")
async def get_province_performance():
    conn = get_db_conn()
    cursor = conn.cursor()
    session = cursor.execute("SELECT session_id FROM import_sessions WHERE status='SUCCESS' ORDER BY session_id DESC LIMIT 1").fetchone()
    if not session: return []
    sid = session['session_id']
    rows = cursor.execute('SELECT province, COUNT(*) as total, SUM(CASE WHEN status = "Thành công" THEN 1 ELSE 0 END) as success, SUM(CASE WHEN is_sla_violation = 1 THEN 1 ELSE 0 END) as sla FROM orders WHERE session_id = ? AND province != "" GROUP BY province ORDER BY total DESC LIMIT 10', (sid,)).fetchall()
    conn.close()
    return [{"name": r['province'], "direction": get_direction(r['province']), "total": r['total'], "success": r['success'], "sla": r['sla'], "success_rate": round(r['success']*100/r['total']) if r['total']>0 else 0, "sla_rate": round(r['sla']*100/r['total']) if r['total']>0 else 0} for r in rows]

@app.get("/api/dashboard/bcvh-summary")
async def get_bcvh_summary():
    conn = get_db_conn()
    cursor = conn.cursor()
    session = cursor.execute("SELECT session_id FROM import_sessions WHERE status='SUCCESS' ORDER BY session_id DESC LIMIT 1").fetchone()
    if not session: return []
    sid = session['session_id']
    rows = cursor.execute('SELECT post_office_name, province, COUNT(*) as total, SUM(CASE WHEN status="Thành công" THEN 1 ELSE 0 END) as success, SUM(CASE WHEN is_sla_violation=1 THEN 1 ELSE 0 END) as sla FROM orders WHERE session_id = ? GROUP BY post_office_name, province ORDER BY total DESC', (sid,)).fetchall()
    conn.close()
    return [{"name": r['post_office_name'], "province": r['province'], "total": r['total'], "success": r['success'], "sla": r['sla'], "rate": round(r['success']*100/r['total']) if r['total']>0 else 0} for r in rows]

@app.get("/api/dashboard/sla-risk")
async def get_sla_risk():
    conn = get_db_conn()
    cursor = conn.cursor()
    session = cursor.execute("SELECT session_id FROM import_sessions WHERE status='SUCCESS' ORDER BY session_id DESC LIMIT 1").fetchone()
    if not session: return []
    sid = session['session_id']
    rows = cursor.execute('SELECT tracking_id, aging, province, post_office_name FROM orders WHERE session_id = ? AND is_sla_violation = 1 ORDER BY aging DESC LIMIT 50', (sid,)).fetchall()
    conn.close()
    return [dict(r) for r in rows]

@app.get("/", response_class=HTMLResponse)
async def get_index():
    with open(os.path.join(BASE_DIR, "backend", "index.html"), "r", encoding="utf-8") as f: return f.read()

if __name__ == "__main__":
    uvicorn.run(app, host="0.0.0.0", port=8010)
