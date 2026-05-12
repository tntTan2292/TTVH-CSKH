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

app = FastAPI()
app.add_middleware(CORSMiddleware, allow_origins=["*"], allow_methods=["*"], allow_headers=["*"])

def get_db_conn():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn

@app.post("/upload")
async def upload_excel(file: UploadFile = File(...)):
    log_terminal(f"--- [REALITY IMPORT] START: {file.filename} ---")
    file_path = os.path.join(UPLOAD_DIR, f"final_{datetime.now().strftime('%H%M%S')}_{file.filename}")
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
            "tracking_id": next((i for i, h in enumerate(df.columns) if any(k in h.lower() for k in ["số hiệu", "mã bưu gửi"])), 0),
            "acceptance_date": next((i for i, h in enumerate(df.columns) if any(k in h.lower() for k in ["ngày chấp nhận", "ngày gửi"])), 1),
            "result_first": next((i for i, h in enumerate(df.columns) if "lần đầu" in h.lower()), 2),
            "result_final": next((i for i, h in enumerate(df.columns) if "cuối cùng" in h.lower() or "lần cuối" in h.lower()), 3),
            "province": next((i for i, h in enumerate(df.columns) if "tỉnh" in h.lower() and "mã" not in h.lower()), 4),
            "post_office": next((i for i, h in enumerate(df.columns) if "bcvh" in h.lower() or "bưu cục vận hành" in h.lower()), 5)
        }

        df = df.drop_duplicates(subset=[df.columns[mapping["tracking_id"]]], keep='last')
        now = datetime.now()
        processed = []
        for _, row in df.iterrows():
            tid = str(row.iloc[mapping["tracking_id"]]).strip()
            dt_acc = parse_dt(row.iloc[mapping["acceptance_date"]])
            res_first = clean_text_lower(row.iloc[mapping["result_first"]])
            res_final = clean_text_lower(row.iloc[mapping["result_final"]])
            province = clean_text_nfc(row.iloc[mapping["province"]])
            
            is_success = "đã phát thành công" in res_final
            is_sla = (now - dt_acc).days > 3 if dt_acc else False
            aging = (now - dt_acc).days if dt_acc else 0
            
            processed.append((tid, province, clean_text_nfc(row.iloc[mapping["post_office"]]), str(dt_acc) if dt_acc else "", res_first, res_final, 'Thành công' if is_success else 'Chưa thành công', int(aging), is_sla, 0))

        conn = get_db_conn()
        cursor = conn.cursor()
        try:
            cursor.execute("BEGIN TRANSACTION")
            cursor.execute("DELETE FROM orders") # Refresh for report accuracy
            cursor.execute('INSERT INTO import_sessions (filename, imported_at, total_rows, status) VALUES (?, ?, ?, ?)', (file.filename, now.isoformat(), len(df), "SUCCESS"))
            sid = cursor.lastrowid
            final_data = [list(r)[:-1] + [sid] for r in processed]
            cursor.executemany('INSERT OR REPLACE INTO orders (tracking_id, province, post_office_name, acceptance_date, result_first, result_final, status, aging, is_sla_violation, session_id) VALUES (?,?,?,?,?,?,?,?,?,?)', final_data)
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
    # 1. TỔNG QUÁT
    kpis = cursor.execute('SELECT COUNT(*) as t, SUM(CASE WHEN status="Thành công" THEN 1 ELSE 0 END) as s, SUM(CASE WHEN result_first LIKE "%chưa có tt phát%" THEN 1 ELSE 0 END) as p FROM orders').fetchone()
    if not kpis or kpis['t'] == 0: return {"error": "Chưa có dữ liệu"}
    t, s, p = kpis['t'], kpis['s'], kpis['p']
    failed = t - s - p
    
    # 2. XU HƯỚNG THEO NGÀY (SUCCESS RATE)
    trends = cursor.execute("SELECT SUBSTR(acceptance_date, 1, 10) as d, COUNT(*) as ct, SUM(CASE WHEN status='Thành công' THEN 1 ELSE 0 END) as su FROM orders GROUP BY d ORDER BY d ASC").fetchall()
    trend_text = ", ".join([f"{r['d'][-5:]}: {round(r['su']*100/r['ct'])}%" for r in trends])

    # 3. TOP TỈNH
    provinces = cursor.execute("SELECT province, COUNT(*) as ct, SUM(CASE WHEN status='Thành công' THEN 1 ELSE 0 END) as su FROM orders GROUP BY province ORDER BY su*1.0/ct DESC LIMIT 3").fetchall()
    top_provinces = " và ".join([r['province'] for r in provinces])

    # 4. HƯỚNG
    def get_dir_stats(keyword_list, exclude_hue=False):
        where = " OR ".join([f"province LIKE '%{k}%'" for k in keyword_list])
        if exclude_hue: where = f"({where}) AND province NOT LIKE '%Huế%'"
        res = cursor.execute(f"SELECT COUNT(*) as ct, SUM(CASE WHEN status='Thành công' THEN 1 ELSE 0 END) as su FROM orders WHERE {where}").fetchone()
        return res['ct'] or 0, round((res['su'] or 0)*100/(res['ct'] or 1))

    i_ct, i_rate = get_dir_stats(["Huế"])
    n_ct, n_rate = get_dir_stats(NORTH_PROVINCES, exclude_hue=True)
    s_ct, s_rate = get_dir_stats(["Hồ Chí Minh", "Bình Dương", "Đồng Nai", "Long An"], exclude_hue=True) # Sample South Keywords

    # 5. ĐIỂM NÓNG (RISK 07/05)
    risk_0705 = cursor.execute("SELECT COUNT(*) as ct FROM orders WHERE acceptance_date LIKE '%05-07%' AND result_first LIKE '%chưa có tt phát%'").fetchone()['ct'] or 0

    ctx = {
        "date_now": datetime.now().strftime("%Hh%M ngày %d/%m/%Y"),
        "total": t, "success_count": s, "success_rate": round(s*100/t),
        "pending_count": p, "pending_rate": round(p*100/t),
        "failed_count": failed, "failed_rate": round(failed*100/t),
        "trend_text": trend_text,
        "top_provinces": top_provinces,
        "n_rate": n_rate, "s_rate": s_rate, "i_rate": i_rate, "s_ct": s_ct,
        "risk_0705": risk_0705
    }

    try:
        with open(TEMPLATE_PATH, "r", encoding="utf-8") as f:
            template = f.read()
            return {"report": template.format(**ctx)}
    except: return {"error": "Template Error"}

@app.get("/api/dashboard/stats")
async def get_stats():
    conn = get_db_conn()
    cursor = conn.cursor()
    kpis = cursor.execute('SELECT COUNT(*) as total, SUM(CASE WHEN status="Thành công" THEN 1 ELSE 0 END) as success, SUM(CASE WHEN is_sla_violation=1 THEN 1 ELSE 0 END) as sla FROM orders').fetchone()
    conn.close()
    if not kpis: return {"error": "No Data"}
    t, s = kpis['total'], kpis['success']
    return {"kpis": {"total": t, "success": s, "pending": t - s, "sla": kpis['sla']}, "directions": {"intra": {"total": 0, "success": 0}, "north": {"total": 0, "success": 0}, "south": {"total": 0, "success": 0}}}

@app.get("/", response_class=HTMLResponse)
async def get_index():
    with open(os.path.join(BASE_DIR, "backend", "index.html"), "r", encoding="utf-8") as f: return f.read()

if __name__ == "__main__":
    uvicorn.run(app, host="0.0.0.0", port=8010)
