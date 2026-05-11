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

def get_config():
    with open(CONFIG_PATH, "r", encoding="utf-8") as f:
        return json.load(f)

def clean_text(text):
    if not text or pd.isna(text): return ""
    text = unicodedata.normalize('NFC', str(text))
    return text.replace('\xa0', ' ').strip()

def log_terminal(msg):
    print(f"[{datetime.now().strftime('%H:%M:%S')}] {msg}")
    sys.stdout.flush()

app = FastAPI(title="VNPost Hue VIP Governance System")
app.add_middleware(CORSMiddleware, allow_origins=["*"], allow_methods=["*"], allow_headers=["*"])

def get_db_conn():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    conn = get_db_conn()
    cursor = conn.cursor()
    cursor.execute('''CREATE TABLE IF NOT EXISTS import_sessions (session_id INTEGER PRIMARY KEY AUTOINCREMENT, filename TEXT, imported_at DATETIME, total_rows INTEGER, unique_ids INTEGER, success_count INTEGER, status TEXT)''')
    cursor.execute('CREATE TABLE IF NOT EXISTS customers (customer_id TEXT PRIMARY KEY, customer_name TEXT, last_import DATETIME)')
    cursor.execute('''CREATE TABLE IF NOT EXISTS orders (
        tracking_id TEXT, customer_id TEXT, recipient_address TEXT, post_office_name TEXT, province TEXT, 
        acceptance_date TEXT, final_delivery_result TEXT, status TEXT, aging INTEGER, is_sla_violation BOOLEAN, 
        session_id INTEGER, PRIMARY KEY (tracking_id, session_id))''')
    conn.commit()
    conn.close()

init_db()

@app.post("/upload")
async def upload_excel(file: UploadFile = File(...)):
    log_terminal(f"--- [FINAL FIX] START IMPORT: {file.filename} ---")
    session_id = None
    file_path = os.path.join(UPLOAD_DIR, f"sess_{datetime.now().strftime('%Y%m%d%H%M%S')}_{file.filename}")
    try:
        CONFIG = get_config()
        content = await file.read()
        with open(file_path, "wb") as buffer: buffer.write(content)
        
        with pd.ExcelFile(file_path, engine='openpyxl') as xls:
            target_sheet = 'DanhSach' if 'DanhSach' in xls.sheet_names else xls.sheet_names[0]
            df_raw_10 = pd.read_excel(xls, sheet_name=target_sheet, header=None, nrows=10)
            
            header_row_idx = 1 
            for i, row in df_raw_10.iterrows():
                row_str = " ".join([clean_text(val).lower() for val in row if not pd.isna(val)])
                if "số hiệu" in row_str or "mã bưu gửi" in row_str:
                    header_row_idx = i
                    break
            
            df = pd.read_excel(xls, sheet_name=target_sheet, header=header_row_idx)
            df.columns = [clean_text(h) for h in df.columns]

        # 3. CHỐT HẠ MAPPING LOGIC
        headers_lower = [h.lower() for h in df.columns]
        
        def find_best_col(target_keywords, exclude_keywords=[]):
            # Ưu tiên khớp chính xác hoặc có từ khóa quan trọng nhất (như "cuối cùng")
            for i, h in enumerate(headers_lower):
                if any(tk in h for tk in target_keywords):
                    if not any(ek in h for ek in exclude_keywords):
                        return i
            return None

        mapping = {
            "tracking_id": find_best_col(["số hiệu", "mã bưu gửi"]),
            "result_final": find_best_col(["kết quả phát cuối cùng", "kết quả cuối cùng"], []) or find_best_col(["kết quả phát", "trạng thái phát"], ["lần đầu"]),
            "province": find_best_col(["tỉnh phát", "tỉnh"], ["mã"]),
            "post_office": find_best_col(["bcvh", "tên bcvh", "bưu cục vận hành"], ["mã"]),
            "acceptance_date": find_best_col(["ngày chấp nhận", "ngày gửi"]),
            "address": find_best_col(["địa chỉ"], ["mã"])
        }

        log_terminal(f"FINAL MAPPING DETECTED: {mapping}")
        
        # Validation
        for k in ["tracking_id", "result_final", "post_office"]:
            if mapping[k] is None: raise Exception(f"Thiếu cột bắt buộc: {k}")

        df = df.drop_duplicates(subset=[df.columns[mapping["tracking_id"]]], keep='last')
        customer_info = clean_text(df_raw_10.iloc[0, 0])
        customer_id = customer_info.split('-')[0].strip() if '-' in customer_info else "UNKNOWN"
        customer_name = customer_info.split('-')[1].strip() if '-' in customer_info else customer_info

        conn = get_db_conn()
        cursor = conn.cursor()
        try:
            cursor.execute("BEGIN TRANSACTION")
            cursor.execute('INSERT INTO import_sessions (filename, imported_at, total_rows, status) VALUES (?, ?, ?, ?)', (file.filename, datetime.now().isoformat(), len(df), "PROCESSING"))
            session_id = cursor.lastrowid
            
            processed_orders = []
            success_count = 0
            now = datetime.now()
            
            for _, row in df.iterrows():
                tid = clean_text(row.iloc[mapping["tracking_id"]])
                if not tid: continue
                
                res_val = clean_text(row.iloc[mapping["result_final"]]).lower()
                is_success = any(k in res_val for k in CONFIG["SUCCESS_KEYWORDS"]) and not any(k in res_val for k in CONFIG["FAIL_KEYWORDS"])
                if is_success: success_count += 1
                
                aging = 0
                try:
                    acc_val = row.iloc[mapping["acceptance_date"]]
                    if not pd.isna(acc_val): aging = (now - pd.to_datetime(acc_val)).days
                except: pass

                processed_orders.append((
                    tid, customer_id, 
                    clean_text(row.iloc[mapping["address"]]) if mapping["address"] is not None else "", 
                    clean_text(row.iloc[mapping["post_office"]]), 
                    clean_text(row.iloc[mapping["province"]]) if mapping["province"] is not None else "", 
                    str(row.iloc[mapping["acceptance_date"]]), 
                    clean_text(row.iloc[mapping["result_final"]]), 
                    'Thành công' if is_success else 'Chưa thành công', 
                    max(0, int(aging)), (aging > CONFIG["SLA_THRESHOLD_DAYS"] and not is_success), 
                    session_id
                ))

            cursor.execute('UPDATE import_sessions SET unique_ids = ?, success_count = ?, status = ? WHERE session_id = ?', (len(processed_orders), success_count, "SUCCESS", session_id))
            cursor.execute("INSERT OR REPLACE INTO customers (customer_id, customer_name, last_import) VALUES (?, ?, ?)", (customer_id, customer_name, datetime.now().isoformat()))
            cursor.executemany('INSERT OR REPLACE INTO orders VALUES (?,?,?,?,?,?,?,?,?,?,?)', processed_orders)
            cursor.execute("COMMIT")
            
            log_terminal(f"SUCCESS! Session {session_id}, Count = {success_count} (Should be 256)")
            return {"message": "Success", "session_id": session_id}
        except Exception as e:
            cursor.execute("ROLLBACK")
            raise e
        finally: conn.close()
    except Exception as e:
        log_terminal(f"ERR: {str(e)}")
        return JSONResponse(status_code=500, content={"detail": str(e)})
    finally:
        if os.path.exists(file_path): os.remove(file_path)

@app.get("/api/dashboard/radar")
async def get_radar_data():
    conn = get_db_conn()
    cursor = conn.cursor()
    session = cursor.execute("SELECT session_id FROM import_sessions WHERE status='SUCCESS' ORDER BY session_id DESC LIMIT 1").fetchone()
    if not session: return {"labels": [], "datasets": []}
    sid = session['session_id']
    rows = cursor.execute('''
        SELECT province as grp, COUNT(*) as total, 
        SUM(CASE WHEN status = "Thành công" THEN 1 ELSE 0 END) as success 
        FROM orders WHERE session_id = ? AND province != '' AND province IS NOT NULL
        GROUP BY province ORDER BY total DESC LIMIT 8
    ''', (sid,)).fetchall()
    if not rows or len(rows) < 2:
        rows = cursor.execute('''
            SELECT post_office_name as grp, COUNT(*) as total, 
            SUM(CASE WHEN status = "Thành công" THEN 1 ELSE 0 END) as success 
            FROM orders WHERE session_id = ? 
            GROUP BY post_office_name ORDER BY total DESC LIMIT 8
        ''', (sid,)).fetchall()
    labels = [r['grp'][:15] for r in rows]
    data = [round(r['success']*100/r['total']) if r['total']>0 else 0 for r in rows]
    conn.close()
    return {"labels": labels, "datasets": [{"label": "Hiệu suất (%)", "data": data}]}

@app.get("/api/dashboard/stats")
async def get_stats():
    conn = get_db_conn()
    cursor = conn.cursor()
    session = cursor.execute("SELECT * FROM import_sessions WHERE status='SUCCESS' ORDER BY session_id DESC LIMIT 1").fetchone()
    if not session: return {"error": "No Data"}
    sid = session['session_id']
    kpis = cursor.execute('SELECT COUNT(*) as total, SUM(CASE WHEN status="Thành công" THEN 1 ELSE 0 END) as success, SUM(CASE WHEN is_sla_violation=1 THEN 1 ELSE 0 END) as sla FROM orders WHERE session_id=?', (sid,)).fetchone()
    conn.close()
    return {
        "session_info": {"id": sid, "timestamp": session['imported_at'], "filename": session['filename']},
        "kpis": {"total": kpis['total'], "success": kpis['success'], "pending": kpis['total'] - kpis['success'], "sla": kpis['sla']}
    }

@app.get("/api/dashboard/bcvh-summary")
async def get_bcvh_summary():
    conn = get_db_conn()
    cursor = conn.cursor()
    session = cursor.execute("SELECT session_id FROM import_sessions WHERE status='SUCCESS' ORDER BY session_id DESC LIMIT 1").fetchone()
    if not session: return []
    sid = session['session_id']
    rows = cursor.execute('SELECT post_office_name, COUNT(*) as total, SUM(CASE WHEN status="Thành công" THEN 1 ELSE 0 END) as success, SUM(CASE WHEN is_sla_violation=1 THEN 1 ELSE 0 END) as sla FROM orders WHERE session_id = ? GROUP BY post_office_name ORDER BY total DESC', (sid,)).fetchall()
    conn.close()
    return [{"name": r['post_office_name'], "total": r['total'], "success": r['success'], "sla": r['sla'], "rate": round(r['success']*100/r['total']) if r['total']>0 else 0} for r in rows]

@app.get("/api/dashboard/bcvh-bottleneck")
async def get_bottlenecks():
    conn = get_db_conn()
    cursor = conn.cursor()
    session = cursor.execute("SELECT session_id FROM import_sessions WHERE status='SUCCESS' ORDER BY session_id DESC LIMIT 1").fetchone()
    if not session: return []
    sid = session['session_id']
    rows = cursor.execute('SELECT post_office_name, COUNT(*) as backlog, MAX(aging) as max_aging FROM orders WHERE session_id = ? AND status != "Thành công" GROUP BY post_office_name ORDER BY backlog DESC LIMIT 10', (sid,)).fetchall()
    conn.close()
    return [{"name": r['post_office_name'], "backlog": r['backlog'], "max_aging": r['max_aging']} for r in rows]

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
