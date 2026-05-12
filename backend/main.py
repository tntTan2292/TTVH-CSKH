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
os.makedirs(UPLOAD_DIR, exist_ok=True)

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
    cursor.execute('''CREATE TABLE IF NOT EXISTS import_sessions (session_id INTEGER PRIMARY KEY AUTOINCREMENT, filename TEXT, imported_at DATETIME, total_rows INTEGER, status TEXT)''')
    cursor.execute('''CREATE TABLE IF NOT EXISTS orders (
        tracking_id TEXT, acceptance_date TEXT, ttp_first_date TEXT, result_first TEXT, result_final TEXT, status TEXT, aging INTEGER, is_sla_violation BOOLEAN, 
        session_id INTEGER, PRIMARY KEY (tracking_id, session_id))''')
    conn.commit()
    conn.close()

init_db()

@app.post("/upload")
async def upload_excel(file: UploadFile = File(...)):
    log_terminal(f"--- [DEEP FORENSIC] START: {file.filename} ---")
    file_path = os.path.join(UPLOAD_DIR, f"debug_{datetime.now().strftime('%H%M%S')}_{file.filename}")
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
            "result_final": find_best_col(df.columns, ["kết quả phát cuối cùng", "kết quả phát lần cuối"])
        }

        df = df.drop_duplicates(subset=[df.columns[mapping["tracking_id"]]], keep='last')
        
        # DEBUG COUNTERS
        d_parsed_accept = 0
        d_parsed_ttp = 0
        d_first_result_found = 0
        d_case1 = 0
        d_case2 = 0
        d_success = 0
        samples = []
        
        now = datetime.now()
        processed_orders = []
        
        for _, row in df.iterrows():
            tid = str(row.iloc[mapping["tracking_id"]]).strip()
            if not tid: continue
            
            # PARSING
            dt_accept = parse_dt(row.iloc[mapping["acceptance_date"]])
            dt_ttp = parse_dt(row.iloc[mapping["ttp_first_date"]]) if mapping["ttp_first_date"] is not None else None
            res_first = clean_text_lower(row.iloc[mapping["result_first"]])
            res_final = clean_text_lower(row.iloc[mapping["result_final"]])
            
            if dt_accept: d_parsed_accept += 1
            if dt_ttp: d_parsed_ttp += 1
            if res_first: d_first_result_found += 1
            
            # SLA LOGIC
            is_sla = False
            # CASE 1
            if dt_accept and dt_ttp and (dt_ttp - dt_accept).days > 3:
                is_sla = True
                d_case1 += 1
            # CASE 2
            if not is_sla and dt_accept and "chưa có tt phát" in res_first and (now - dt_accept).days > 3:
                is_sla = True
                d_case2 += 1
            
            is_success = "đã phát thành công" in res_final
            if is_success: d_success += 1
            
            aging = (now - dt_accept).days if dt_accept else 0
            
            if len(samples) < 5:
                samples.append({
                    "tracking_id": tid,
                    "accept_date": str(dt_accept) if dt_accept else "NULL",
                    "first_ttp": str(dt_ttp) if dt_ttp else "NULL",
                    "first_result": res_first or "NULL",
                    "aging_days": int(aging),
                    "is_sla": is_sla
                })
            
            processed_orders.append((
                tid, str(dt_accept) if dt_accept else "", str(dt_ttp) if dt_ttp else "",
                res_first, res_final, 'Thành công' if is_success else 'Chưa thành công', 
                int(aging), is_sla, 0 # session_id will be added
            ))

        # LOG BLOCK
        print("\n================ SLA DEBUG ================")
        print(f"TOTAL_ROWS:          {len(df)}")
        print(f"PARSED_ACCEPT_DATE:  {d_parsed_accept}")
        print(f"PARSED_FIRST_TTP:    {d_parsed_ttp}")
        print(f"FIRST_RESULT_FOUND:  {d_first_result_found}")
        print(f"CASE_1_SLA_COUNT:    {d_case1}")
        print(f"CASE_2_SLA_COUNT:    {d_case2}")
        print(f"FINAL_SLA_COUNT:     {d_case1 + d_case2}")
        print(f"SUCCESS_COUNT:       {d_success}")
        print(f"PENDING_COUNT:       {len(df) - d_success}")
        print("===========================================\n")
        
        print("SAMPLE ROWS (JSON):")
        print(json.dumps(samples, indent=2, ensure_ascii=False))
        print("-------------------------------------------\n")

        conn = get_db_conn()
        cursor = conn.cursor()
        try:
            cursor.execute("BEGIN TRANSACTION")
            cursor.execute('INSERT INTO import_sessions (filename, imported_at, total_rows, status) VALUES (?, ?, ?, ?)', (file.filename, datetime.now().isoformat(), len(df), "SUCCESS"))
            sid = cursor.lastrowid
            final_data = [list(r)[:-1] + [sid] for r in processed_orders]
            cursor.executemany('INSERT OR REPLACE INTO orders VALUES (?,?,?,?,?,?,?,?,?)', final_data)
            cursor.execute("COMMIT")
            return {"message": "Forensic Logged", "sid": sid}
        finally: conn.close()

    except Exception as e:
        log_terminal(f"ERR: {str(e)}")
        traceback.print_exc()
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
    kpis = cursor.execute('SELECT COUNT(*) as total, SUM(CASE WHEN status="Thành công" THEN 1 ELSE 0 END) as success, SUM(CASE WHEN is_sla_violation=1 THEN 1 ELSE 0 END) as sla FROM orders WHERE session_id=?', (sid,)).fetchone()
    conn.close()
    return {"kpis": {"total": kpis['total'], "success": kpis['success'], "pending": kpis['total'] - kpis['success'], "sla": kpis['sla']}}

@app.get("/", response_class=HTMLResponse)
async def get_index():
    with open(os.path.join(BASE_DIR, "backend", "index.html"), "r", encoding="utf-8") as f: return f.read()

if __name__ == "__main__":
    uvicorn.run(app, host="0.0.0.0", port=8010)
