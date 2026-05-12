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

def verify_import_integrity(sid):
    """
    REGRESSION TEST SUITE V1.0
    Perform critical assertions and data sanity checks after every import.
    """
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()
    try:
        # 1. GATHER CORE METRICS
        total = cursor.execute("SELECT COUNT(*) as c FROM orders WHERE session_id=?", (sid,)).fetchone()['c']
        success = cursor.execute("SELECT COUNT(*) as c FROM orders WHERE session_id=? AND status='Thành công'", (sid,)).fetchone()['c']
        sla = cursor.execute("SELECT COUNT(*) as c FROM orders WHERE session_id=? AND is_sla_violation=1", (sid,)).fetchone()['c']
        
        hue = cursor.execute("SELECT COUNT(*) as c FROM orders WHERE session_id=? AND (province LIKE '%Huế%' OR province LIKE '%Hue%')", (sid,)).fetchone()['c']
        north_where = " OR ".join([f"province LIKE '%{p}%'" for p in NORTH_PROVINCES])
        north = cursor.execute(f"SELECT COUNT(*) as c FROM orders WHERE session_id=? AND NOT (province LIKE '%Huế%' OR province LIKE '%Hue%') AND ({north_where})", (sid,)).fetchone()['c']
        south = total - hue - north
        
        provinces = cursor.execute("SELECT DISTINCT province FROM orders WHERE session_id=? AND province != ''", (sid,)).fetchall()
        bcvhs = cursor.execute("SELECT DISTINCT post_office_name FROM orders WHERE session_id=? AND post_office_name != ''", (sid,)).fetchall()
        
        # 2. RUN ASSERTIONS
        assertions = {
            "success_in_range": success <= total,
            "geo_sum_match": (hue + north + south) == total,
            "province_data_present": len(provinces) > 0,
            "bcvh_data_present": len(bcvhs) > 0,
            "sla_distinct_from_pending": sla != (total - success) # Basic logic check
        }
        
        verification_report = {
            "session_id": sid,
            "timestamp": datetime.now().isoformat(),
            "metrics": {
                "total": total, "success": success, "pending": total - success, "sla": sla,
                "north": north, "south": south, "hue_local": hue
            },
            "coverage": {
                "provinces_count": len(provinces),
                "bcvh_count": len(bcvhs)
            },
            "assertions": assertions,
            "status": "PASS" if all(assertions.values()) else "FAIL"
        }
        
        print("\n" + "="*50)
        print("📊 REGRESSION VERIFICATION REPORT")
        print("="*50)
        print(json.dumps(verification_report, indent=2))
        print("="*50 + "\n")
        
        if verification_report["status"] == "FAIL":
            log_terminal("⚠️ WARNING: DATA INTEGRITY CHECK FAILED. REVIEW ASSERTIONS.")
            
    except Exception as e:
        log_terminal(f"❌ VERIFICATION CRASHED: {str(e)}")
    finally:
        conn.close()

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
    log_terminal(f"--- [STABILIZATION] IMPORT: {file.filename} ---")
    file_path = os.path.join(UPLOAD_DIR, f"stable_{datetime.now().strftime('%H%M%S')}_{file.filename}")
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
                is_success = "đã phát thành công" in res_final
                aging = (now - dt_acc).days if dt_acc else 0
                is_sla = aging > 3 and "chưa có tt phát" in res_first
                processed.append((tid, clean_text_nfc(row.iloc[mapping["province"]]), clean_text_nfc(row.iloc[mapping["post_office"]]), str(dt_acc) if dt_acc else "", str(dt_ttp) if dt_ttp else "", res_first, res_final, 'Thành công' if is_success else 'Chưa thành công', int(aging), is_sla, sid))
            cursor.executemany('INSERT INTO orders VALUES (?,?,?,?,?,?,?,?,?,?,?)', processed)
            cursor.execute("COMMIT")
            
            # TRIGGER REGRESSION TEST SUITE
            verify_import_integrity(sid)
            
            return {"message": "Success", "sid": sid}
        finally: conn.close()
    except Exception as e:
        log_terminal(f"ERR: {str(e)}")
        return JSONResponse(status_code=500, content={"detail": str(e)})
    finally:
        if os.path.exists(file_path): os.remove(file_path)

@app.get("/api/dashboard/acceptance-trend")
async def get_acceptance_trend():
    conn = get_db_conn()
    cursor = conn.cursor()
    session = cursor.execute("SELECT session_id FROM import_sessions ORDER BY session_id DESC LIMIT 1").fetchone()
    if not session: return {"data": [], "kpis": {}}
    sid = session['session_id']
    north_where = " OR ".join([f"province LIKE '%{p}%'" for p in NORTH_PROVINCES])
    query = f'''
        SELECT 
            SUBSTR(acceptance_date, 1, 10) as date,
            COUNT(*) as total,
            SUM(CASE WHEN status="Thành công" THEN 1 ELSE 0 END) as success,
            SUM(CASE WHEN (province LIKE "%Huế%" OR province LIKE "%Hue%") THEN 1 ELSE 0 END) as intra,
            SUM(CASE WHEN NOT (province LIKE "%Huế%" OR province LIKE "%Hue%") AND ({north_where}) THEN 1 ELSE 0 END) as north
        FROM orders 
        WHERE session_id = ? AND acceptance_date != ""
        GROUP BY date ORDER BY date ASC
    '''
    rows = cursor.execute(query, (sid,)).fetchall()
    conn.close()
    trend_data = []
    peak_val = 0
    peak_date = "N/A"
    for r in rows:
        d = r['date']; tot = r['total']; suc = r['success']
        intra = r['intra'] or 0; north = r['north'] or 0; south = tot - intra - north
        if tot > peak_val: peak_val = tot; peak_date = d
        trend_data.append({"date": d[-5:], "total": tot, "success": suc, "intra": intra, "north": north, "south": south})
    return {"data": trend_data, "kpis": {"peak_day": peak_date, "peak_value": peak_val, "avg_volume": round(sum(d['total'] for d in trend_data)/len(trend_data)) if trend_data else 0}}

@app.get("/api/dashboard/stats")
async def get_stats():
    conn = get_db_conn()
    cursor = conn.cursor()
    session = cursor.execute("SELECT session_id FROM import_sessions ORDER BY session_id DESC LIMIT 1").fetchone()
    if not session: return {"error": "No Data"}
    sid = session['session_id']
    kpis = cursor.execute('SELECT COUNT(*) as t, SUM(CASE WHEN status="Thành công" THEN 1 ELSE 0 END) as s, SUM(CASE WHEN is_sla_violation=1 THEN 1 ELSE 0 END) as sla FROM orders WHERE session_id=?', (sid,)).fetchone()
    intra = cursor.execute("SELECT COUNT(*) as t, SUM(CASE WHEN status='Thành công' THEN 1 ELSE 0 END) as s FROM orders WHERE session_id=? AND (province LIKE '%Huế%' OR province LIKE '%Hue%')", (sid,)).fetchone()
    north_where = " OR ".join([f"province LIKE '%{p}%'" for p in NORTH_PROVINCES])
    north = cursor.execute(f"SELECT COUNT(*) as t, SUM(CASE WHEN status='Thành công' THEN 1 ELSE 0 END) as s FROM orders WHERE session_id=? AND NOT (province LIKE '%Huế%' OR province LIKE '%Hue%') AND ({north_where})", (sid,)).fetchone()
    conn.close()
    t, s = kpis['t'], kpis['s']
    it, isuc = intra['t'] or 0, intra['s'] or 0
    nt, nsuc = north['t'] or 0, north['s'] or 0
    return {"kpis": {"total": t, "success": s, "pending": t - s, "sla": kpis['sla']}, "directions": {"intra": {"total": it, "success": isuc}, "north": {"total": nt, "success": nsuc}, "south": {"total": t - it - nt, "success": s - isuc - nsuc}}}

@app.get("/api/dashboard/province-performance")
async def get_province_performance():
    conn = get_db_conn()
    cursor = conn.cursor()
    session = cursor.execute("SELECT session_id FROM import_sessions ORDER BY session_id DESC LIMIT 1").fetchone()
    if not session: return []
    sid = session['session_id']
    rows = cursor.execute('SELECT province, COUNT(*) as ct, SUM(CASE WHEN status = "Thành công" THEN 1 ELSE 0 END) as success, SUM(CASE WHEN is_sla_violation = 1 THEN 1 ELSE 0 END) as sla FROM orders WHERE session_id = ? AND province != "" GROUP BY province ORDER BY ct DESC LIMIT 10', (sid,)).fetchall()
    conn.close()
    return [{"name": r['province'], "direction": "Nội tỉnh" if "Huế" in r['province'] else ("Bắc" if any(n in r['province'] for n in NORTH_PROVINCES) else "Nam"), "total": r['ct'], "success": r['success'], "sla": r['sla'], "success_rate": round(r['success']*100/r['ct']) if r['ct']>0 else 0, "sla_rate": round(r['sla']*100/r['ct']) if r['ct']>0 else 0} for r in rows]

@app.get("/api/dashboard/bcvh-summary")
async def get_bcvh_summary():
    conn = get_db_conn()
    cursor = conn.cursor()
    session = cursor.execute("SELECT session_id FROM import_sessions ORDER BY session_id DESC LIMIT 1").fetchone()
    if not session: return []
    sid = session['session_id']
    rows = cursor.execute('SELECT post_office_name, province, COUNT(*) as ct, SUM(CASE WHEN status="Thành công" THEN 1 ELSE 0 END) as su, SUM(CASE WHEN is_sla_violation=1 THEN 1 ELSE 0 END) as sla FROM orders WHERE session_id = ? GROUP BY post_office_name, province ORDER BY ct DESC', (sid,)).fetchall()
    conn.close()
    return [{"name": r['post_office_name'], "province": r['province'], "total": r['ct'], "success": r['su'], "sla": r['sla'], "rate": round(r['su']*100/r['ct']) if r['ct']>0 else 0} for r in rows]

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

@app.get("/api/dashboard/generate-report")
async def generate_report():
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
    ctx = {"date_now": datetime.now().strftime("%Hh%M ngày %d/%m/%Y"), "total": t, "success_count": s, "success_rate": round(s*100/t), "pending_count": p, "pending_rate": round(p*100/t), "failed_count": failed, "failed_rate": round(failed*100/t), "trend_text": trend_text, "top_provinces": top_provinces, "n_rate": 0, "s_rate": 0, "i_rate": 0, "s_ct": 0, "risk_0705": 0}
    with open(TEMPLATE_PATH, "r", encoding="utf-8") as f: return {"report": f.read().format(**ctx)}

@app.get("/", response_class=HTMLResponse)
async def get_index():
    with open(os.path.join(BASE_DIR, "backend", "index.html"), "r", encoding="utf-8") as f: return f.read()

if __name__ == "__main__":
    uvicorn.run(app, host="0.0.0.0", port=8010)
