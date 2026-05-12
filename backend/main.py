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

def clean_text_lower(text):
    if not text or pd.isna(text): return ""
    text = unicodedata.normalize('NFC', str(text))
    return text.lower().strip()

def parse_dt(val):
    """Defensive Datetime Parsing for Excel (Strings, Objects, Serial Numbers)"""
    if pd.isna(val) or val == "": return None
    if isinstance(val, datetime): return val
    if isinstance(val, (int, float)):
        # Handle Excel Serial Date
        try: return datetime(1899, 12, 30) + timedelta(days=val)
        except: return None
    if isinstance(val, str):
        for fmt in ('%Y-%m-%d %H:%M:%S', '%d/%m/%Y %H:%M:%S', '%Y-%m-%d', '%d/%m/%Y'):
            try: return datetime.strptime(val, fmt)
            except: continue
    return None

def find_best_col(headers, target_keywords, exclude_keywords=[]):
    headers_lower = [h.lower() for h in headers]
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
    cursor.execute('''CREATE TABLE IF NOT EXISTS import_sessions (session_id INTEGER PRIMARY KEY AUTOINCREMENT, filename TEXT, imported_at DATETIME, total_rows INTEGER, success_count INTEGER, status TEXT)''')
    cursor.execute('''CREATE TABLE IF NOT EXISTS orders (
        tracking_id TEXT, recipient_address TEXT, post_office_name TEXT, province TEXT, 
        acceptance_date TEXT, ttp_first_date TEXT, result_first TEXT, result_final TEXT, status TEXT, aging INTEGER, is_sla_violation BOOLEAN, 
        session_id INTEGER, PRIMARY KEY (tracking_id, session_id))''')
    conn.commit()
    conn.close()

init_db()

@app.post("/upload")
async def upload_excel(file: UploadFile = File(...)):
    log_terminal(f"=== [SLA HARDENING] START: {file.filename} ===")
    file_path = os.path.join(UPLOAD_DIR, f"sla_hard_{datetime.now().strftime('%H%M%S')}_{file.filename}")
    try:
        content = await file.read()
        with open(file_path, "wb") as f: f.write(content)
        
        with pd.ExcelFile(file_path, engine='openpyxl') as xls:
            target_sheet = 'DanhSach' if 'DanhSach' in xls.sheet_names else xls.sheet_names[0]
            df_raw_10 = pd.read_excel(xls, sheet_name=target_sheet, header=None, nrows=10)
            header_idx = 0
            for i, row in df_raw_10.iterrows():
                row_str = " ".join([str(val).lower() for val in row if not pd.isna(val)])
                if any(k in row_str for k in ["số hiệu", "mã bưu gửi", "kết quả"]):
                    header_idx = i
                    break
            df = pd.read_excel(xls, sheet_name=target_sheet, header=header_idx)
            df.columns = [str(h).strip() for h in df.columns]

        mapping = {
            "tracking_id": find_best_col(df.columns, ["số hiệu", "mã bưu gửi"]),
            "acceptance_date": find_best_col(df.columns, ["ngày chấp nhận", "ngày gửi"]),
            "ttp_first_date": find_best_col(df.columns, ["thời gian nhập ttp lần đầu", "thời gian nhập lần đầu"]),
            "result_first": find_best_col(df.columns, ["kết quả phát lần đầu", "kết quả lần đầu"]),
            "result_final": find_best_col(df.columns, ["kết quả phát cuối cùng", "kết quả phát lần cuối"]),
            "province": find_best_col(df.columns, ["tỉnh"], ["mã"]),
            "post_office": find_best_col(df.columns, ["bcvh", "bưu cục vận hành"]),
            "address": find_best_col(df.columns, ["địa chỉ"])
        }

        df = df.drop_duplicates(subset=[df.columns[mapping["tracking_id"]]], keep='last')
        
        conn = get_db_conn()
        cursor = conn.cursor()
        try:
            cursor.execute("BEGIN TRANSACTION")
            cursor.execute('INSERT INTO import_sessions (filename, imported_at, total_rows, status) VALUES (?, ?, ?, ?)', (file.filename, datetime.now().isoformat(), len(df), "PROCESSING"))
            session_id = cursor.lastrowid
            
            processed_orders = []
            now = datetime.now()
            
            # DEBUG COUNTERS
            d_accept_count = 0
            d_ttp_first_count = 0
            d_case1_count = 0
            d_case2_count = 0
            
            for _, row in df.iterrows():
                tid = str(row.iloc[mapping["tracking_id"]]).strip()
                if not tid: continue
                
                # DT PARSING
                dt_accept = parse_dt(row.iloc[mapping["acceptance_date"]])
                dt_ttp_first = parse_dt(row.iloc[mapping["ttp_first_date"]]) if mapping["ttp_first_date"] is not None else None
                
                if dt_accept: d_accept_count += 1
                if dt_ttp_first: d_ttp_first_count += 1
                
                # SLA LOGIC
                is_sla_violation = False
                res_first = clean_text_lower(row.iloc[mapping["result_first"]])
                
                # CASE 1: (TTP Lần đầu - Chấp nhận) > 3 days
                if dt_accept and dt_ttp_first:
                    if (dt_ttp_first - dt_accept).days > 3:
                        is_sla_violation = True
                        d_case1_count += 1
                
                # CASE 2: "Chưa có TT phát" AND (NOW - Chấp nhận) > 3 days
                if not is_sla_violation and dt_accept and "chưa có tt phát" in res_first:
                    if (now - dt_accept).days > 3:
                        is_sla_violation = True
                        d_case2_count += 1
                
                # GENERAL STATS
                res_final = clean_text_lower(row.iloc[mapping["result_final"]])
                is_success = "đã phát thành công" in res_final
                aging = (now - dt_accept).days if dt_accept else 0
                
                processed_orders.append((
                    tid, str(row.iloc[mapping["address"]]) if mapping["address"] is not None else "", 
                    str(row.iloc[mapping["post_office"]]), str(row.iloc[mapping["province"]]) if mapping["province"] is not None else "", 
                    str(dt_accept) if dt_accept else "", str(dt_ttp_first) if dt_ttp_first else "",
                    res_first, res_final, 'Thành công' if is_success else 'Chưa thành công', 
                    max(0, int(aging)), is_sla_violation, session_id
                ))

            log_terminal(f"DEBUG -> parsed_accept_date_count: {d_accept_count}")
            log_terminal(f"DEBUG -> parsed_ttp_first_count: {d_ttp_first_count}")
            log_terminal(f"DEBUG -> sla_case_1_count: {d_case1_count}")
            log_terminal(f"DEBUG -> sla_case_2_count: {d_case2_count}")
            log_terminal(f"DEBUG -> final_sla_count: {d_case1_count + d_case2_count}")

            cursor.execute('UPDATE import_sessions SET success_count = ?, status = ? WHERE session_id = ?', (0, "SUCCESS", session_id))
            cursor.executemany('INSERT OR REPLACE INTO orders VALUES (?,?,?,?,?,?,?,?,?,?,?,?)', processed_orders)
            cursor.execute("COMMIT")
            return {"message": "SLA Engine Hardened", "session_id": session_id}
        except Exception as e:
            cursor.execute("ROLLBACK")
            raise e
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
    session = cursor.execute("SELECT * FROM import_sessions WHERE status='SUCCESS' ORDER BY session_id DESC LIMIT 1").fetchone()
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
