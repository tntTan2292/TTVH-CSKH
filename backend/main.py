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
RULES_PATH = os.path.join(BASE_DIR, "backend", "report_rules.json")
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

@app.post("/upload")
async def upload_excel(file: UploadFile = File(...)):
    log_terminal(f"--- [GOVERNANCE IMPORT] START: {file.filename} ---")
    file_path = os.path.join(UPLOAD_DIR, f"gov_{datetime.now().strftime('%H%M%S')}_{file.filename}")
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
            "post_office": find_best_col(df.columns, ["bcvh", "tên bcvh", "bưu cục vận hành"]),
            "address": find_best_col(df.columns, ["địa chỉ"])
        }

        df = df.drop_duplicates(subset=[df.columns[mapping["tracking_id"]]], keep='last')
        now = datetime.now()
        processed_orders = []
        
        for _, row in df.iterrows():
            tid = str(row.iloc[mapping["tracking_id"]]).strip()
            if not tid: continue
            dt_accept = parse_dt(row.iloc[mapping["acceptance_date"]])
            dt_ttp = parse_dt(row.iloc[mapping["ttp_first_date"]]) if mapping["ttp_first_date"] is not None else None
            res_first = clean_text_lower(row.iloc[mapping["result_first"]])
            res_final = clean_text_lower(row.iloc[mapping["result_final"]])
            is_sla = False
            if dt_accept and dt_ttp and (dt_ttp - dt_accept).days > 3: is_sla = True
            if not is_sla and dt_accept and "chưa có tt phát" in res_first and (now - dt_accept).days > 3: is_sla = True
            is_success = "đã phát thành công" in res_final
            aging = (now - dt_accept).days if dt_accept else 0
            processed_orders.append((tid, "UNKNOWN", clean_text_nfc(row.iloc[mapping["address"]]), clean_text_nfc(row.iloc[mapping["post_office"]]), clean_text_nfc(row.iloc[mapping["province"]]), str(dt_accept) if dt_accept else "", str(dt_ttp) if dt_ttp else "", res_first, res_final, 'Thành công' if is_success else 'Chưa thành công', int(aging), is_sla, 0))

        conn = get_db_conn()
        cursor = conn.cursor()
        try:
            cursor.execute("BEGIN TRANSACTION")
            cursor.execute('INSERT INTO import_sessions (filename, imported_at, total_rows, status) VALUES (?, ?, ?, ?)', (file.filename, datetime.now().isoformat(), len(df), "SUCCESS"))
            sid = cursor.lastrowid
            final_data = [list(r)[:-1] + [sid] for r in processed_orders]
            cursor.executemany('INSERT OR REPLACE INTO orders VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)', final_data)
            cursor.execute("COMMIT")
            return {"message": "Success", "sid": sid}
        finally: conn.close()
    except Exception as e:
        log_terminal(f"ERR: {str(e)}")
        return JSONResponse(status_code=500, content={"detail": str(e)})
    finally:
        if os.path.exists(file_path): os.remove(file_path)

@app.get("/api/dashboard/generate-report")
async def generate_report():
    conn = get_db_conn()
    cursor = conn.cursor()
    session = cursor.execute("SELECT session_id FROM import_sessions ORDER BY session_id DESC LIMIT 1").fetchone()
    if not session: return {"error": "No Data"}
    sid = session['session_id']
    kpis = cursor.execute('SELECT COUNT(*) as total, SUM(CASE WHEN status="Thành công" THEN 1 ELSE 0 END) as success, SUM(CASE WHEN is_sla_violation=1 THEN 1 ELSE 0 END) as sla FROM orders WHERE session_id=?', (sid,)).fetchone()
    intra = cursor.execute("SELECT COUNT(*) as total, SUM(CASE WHEN status='Thành công' THEN 1 ELSE 0 END) as success FROM orders WHERE session_id=? AND (province LIKE '%Huế%' OR province LIKE '%Hue%')", (sid,)).fetchone()
    north = cursor.execute('''
        SELECT COUNT(*) as total, SUM(CASE WHEN status="Thành công" THEN 1 ELSE 0 END) as success 
        FROM orders WHERE session_id=? AND NOT (province LIKE '%Huế%' OR province LIKE '%Hue%') 
        AND (province LIKE '%Hà Nội%' OR province LIKE '%Bắc Ninh%' OR province LIKE '%Hải Phòng%' OR province LIKE '%Thái Nguyên%' OR province LIKE '%Quảng Ninh%')
    ''', (sid,)).fetchone()
    conn.close()

    # DATA BINDING
    t, s, sla = kpis['total'], kpis['success'], kpis['sla']
    s_rate = round(s * 100 / t) if t > 0 else 0
    p_rate = 100 - s_rate
    sla_rate = round(sla * 100 / t) if t > 0 else 0
    i_rate = round((intra['success'] or 0) * 100 / (intra['total'] or 1))
    n_rate = round((north['success'] or 0) * 100 / (north['total'] or 1))
    south_total = t - (intra['total'] or 0) - (north['total'] or 0)
    south_success = s - (intra['success'] or 0) - (north['success'] or 0)
    so_rate = round(south_success * 100 / (south_total or 1))

    ctx = {
        "date": datetime.now().strftime("%d/%m/%Y %H:%M"),
        "total": t, "success_count": s, "success_rate": s_rate,
        "pending_count": t - s, "pending_rate": p_rate,
        "sla_count": sla, "sla_rate": sla_rate,
        "intra_rate": i_rate, "north_rate": n_rate, "south_rate": so_rate
    }

    # RULE ENGINE
    commentaries = []
    actions = []
    try:
        with open(RULES_PATH, "r", encoding="utf-8") as rf:
            rules_cfg = json.load(rf)
            for rule in rules_cfg['rules']:
                # DYNAMIC EVALUATION
                try:
                    if eval(rule['condition'], {}, ctx):
                        commentaries.append(rule['commentary'].format(**ctx))
                        actions.append(rule['action'].format(**ctx))
                except: pass
    except: pass

    # TEMPLATE BINDING
    try:
        with open(TEMPLATE_PATH, "r", encoding="utf-8") as tf:
            template = tf.read()
            report = template.format(
                commentaries="\n".join(commentaries) if commentaries else "- Hệ thống chưa ghi nhận điểm nóng bất thường.",
                actions="\n".join([f"- {a}" for a in actions]) if actions else "- Tiếp tục duy trì quy trình kiểm soát hiện tại.",
                **ctx
            )
            return {"report": report}
    except Exception as e:
        return {"error": str(e)}

@app.get("/api/dashboard/stats")
async def get_stats():
    conn = get_db_conn()
    cursor = conn.cursor()
    session = cursor.execute("SELECT session_id FROM import_sessions ORDER BY session_id DESC LIMIT 1").fetchone()
    if not session: return {"error": "No Data"}
    sid = session['session_id']
    kpis = cursor.execute('SELECT COUNT(*) as total, SUM(CASE WHEN status="Thành công" THEN 1 ELSE 0 END) as success, SUM(CASE WHEN is_sla_violation=1 THEN 1 ELSE 0 END) as sla FROM orders WHERE session_id=?', (sid,)).fetchone()
    intra = cursor.execute("SELECT COUNT(*) as total, SUM(CASE WHEN status='Thành công' THEN 1 ELSE 0 END) as success FROM orders WHERE session_id=? AND (province LIKE '%Huế%' OR province LIKE '%Hue%')", (sid,)).fetchone()
    north = cursor.execute('''
        SELECT COUNT(*) as total, SUM(CASE WHEN status="Thành công" THEN 1 ELSE 0 END) as success 
        FROM orders WHERE session_id=? AND NOT (province LIKE '%Huế%' OR province LIKE '%Hue%') 
        AND (province LIKE '%Hà Nội%' OR province LIKE '%Bắc Ninh%' OR province LIKE '%Hải Phòng%' OR province LIKE '%Thái Nguyên%' OR province LIKE '%Quảng Ninh%')
    ''', (sid,)).fetchone()
    conn.close()
    t, s = kpis['total'], kpis['success']
    it, isuc = intra['total'] or 0, intra['success'] or 0
    nt, nsuc = north['total'] or 0, north['success'] or 0
    return {"kpis": {"total": t, "success": s, "pending": t - s, "sla": kpis['sla']}, "directions": {"intra": {"total": it, "success": isuc}, "north": {"total": nt, "success": nsuc}, "south": {"total": t - it - nt, "success": s - isuc - nsuc}}}

@app.get("/api/dashboard/province-performance")
async def get_province_performance():
    conn = get_db_conn()
    cursor = conn.cursor()
    session = cursor.execute("SELECT session_id FROM import_sessions ORDER BY session_id DESC LIMIT 1").fetchone()
    if not session: return []
    sid = session['session_id']
    rows = cursor.execute('SELECT province, COUNT(*) as total, SUM(CASE WHEN status = "Thành công" THEN 1 ELSE 0 END) as success, SUM(CASE WHEN is_sla_violation = 1 THEN 1 ELSE 0 END) as sla FROM orders WHERE session_id = ? AND province != "" GROUP BY province ORDER BY total DESC LIMIT 10', (sid,)).fetchall()
    conn.close()
    return [{"name": r['province'], "direction": "Nội tỉnh" if "Huế" in r['province'] else ("Bắc" if any(n in r['province'] for n in NORTH_PROVINCES) else "Nam"), "total": r['total'], "success": r['success'], "sla": r['sla'], "success_rate": round(r['success']*100/r['total']) if r['total']>0 else 0, "sla_rate": round(r['sla']*100/r['total']) if r['total']>0 else 0} for r in rows]

@app.get("/api/dashboard/bcvh-summary")
async def get_bcvh_summary():
    conn = get_db_conn()
    cursor = conn.cursor()
    session = cursor.execute("SELECT session_id FROM import_sessions ORDER BY session_id DESC LIMIT 1").fetchone()
    if not session: return []
    sid = session['session_id']
    rows = cursor.execute('SELECT post_office_name, province, COUNT(*) as total, SUM(CASE WHEN status="Thành công" THEN 1 ELSE 0 END) as success, SUM(CASE WHEN is_sla_violation=1 THEN 1 ELSE 0 END) as sla FROM orders WHERE session_id = ? GROUP BY post_office_name, province ORDER BY total DESC', (sid,)).fetchall()
    conn.close()
    return [{"name": r['post_office_name'], "province": r['province'], "total": r['total'], "success": r['success'], "sla": r['sla'], "rate": round(r['success']*100/r['total']) if r['total']>0 else 0} for r in rows]

@app.get("/api/dashboard/sla-risk")
async def get_sla_risk():
    conn = get_db_conn()
    cursor = conn.cursor()
    session = cursor.execute("SELECT session_id FROM import_sessions ORDER BY session_id DESC LIMIT 1").fetchone()
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
