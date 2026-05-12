import os
import sqlite3
import pandas as pd
import unicodedata
import json
import traceback
import sys
import re
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

# GEOGRAPHIC MAPPING (GOVERNED V4.3)
# Adding Quang Binh, Quang Tri to North as they are "Hướng Bắc" from Hue
NORTH_PROVINCES_RAW = [
    "Hà Nội", "Hải Phòng", "Bắc Ninh", "Thái Nguyên", "Vĩnh Phúc", "Hải Dương", "Quảng Ninh", "Ninh Bình", 
    "Nam Định", "Hà Nam", "Hòa Bình", "Sơn La", "Điện Biên", "Lai Châu", "Lào Cai", "Yên Bái", "Phú Thọ", 
    "Bắc Giang", "Lạng Sơn", "Tuyên Quang", "Hà Giang", "Cao Bằng", "Bắc Kạn", "Hưng Yên", "Thái Bình",
    "Thanh Hóa", "Nghệ An", "Hà Tĩnh", "Quảng Bình", "Quảng Trị"
]

def log_terminal(msg):
    print(f"[{datetime.now().strftime('%H:%M:%S')}] {msg}")
    sys.stdout.flush()

def normalize_province_forensic(name):
    if not name or pd.isna(name): return ""
    name = str(name).lower()
    name = unicodedata.normalize('NFC', name)
    # Remove TP., Tỉnh, spaces
    name = name.replace("tp.", "").replace("tỉnh", "").replace("thành phố", "").strip()
    # Remove extra internal spaces
    name = " ".join(name.split())
    return name

NORTH_PROVINCES_NORM = [normalize_province_forensic(p) for p in NORTH_PROVINCES_RAW]

def run_directional_forensic(sid):
    """
    DIRECTIONAL FORENSIC ENGINE V4.3
    Dumps classification logic to terminal for root-cause analysis.
    """
    conn = sqlite3.connect(DB_PATH); conn.row_factory = sqlite3.Row; cursor = conn.cursor()
    try:
        rows = cursor.execute("SELECT tracking_id, province FROM orders WHERE session_id=?", (sid,)).fetchall()
        
        n_count, s_count, i_count = 0, 0, 0
        mismatches = []
        
        for r in rows:
            raw_p = r['province']
            norm_p = normalize_province_forensic(raw_p)
            
            region = "SOUTH"
            if any(k in norm_p for k in ["huế", "hue"]):
                region = "INTRA"
                i_count += 1
            elif norm_p in NORTH_PROVINCES_NORM or any(np in norm_p for np in NORTH_PROVINCES_NORM):
                region = "NORTH"
                n_count += 1
            else:
                s_count += 1
                # Potential Mismatch Detection (If it looks North but was classified South)
                if any(k in norm_p for k in ["bình", "trị", "hà", "hải", "bắc", "nam", "thái"]):
                    mismatches.append({
                        "tracking": r['tracking_id'],
                        "raw": raw_p,
                        "norm": norm_p,
                        "actual": region
                    })

            # Sample Logging
            if n_count + s_count + i_count <= 10:
                print(f"[FORENSIC_DIRECTION] raw='{raw_p}' -> norm='{norm_p}' -> class={region}")

        print("\n" + "!"*60)
        print("🔥 FORENSIC DIRECTION REPORT (V4.3)")
        print(f"RESULT -> NORTH: {n_count} | SOUTH: {s_count} | INTRA: {i_count} | TOTAL: {n_count+s_count+i_count}")
        print("!"*60)
        
        if mismatches:
            print("\n[POTENTIAL_MISMATCH_SAMPLES]")
            for m in mismatches[:5]:
                print(f"  TRACKING: {m['tracking']} | PROV: {m['raw']} | NORM: {m['norm']} | CLASS: {m['actual']}")
        
    except Exception as e:
        log_terminal(f"❌ FORENSIC FAILED: {str(e)}")
    finally:
        conn.close()

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
    headers_lower = [h.lower().strip() for h in headers]
    for i, h in enumerate(headers_lower):
        if h in target_keywords: return i
    for i, h in enumerate(headers_lower):
        if any(tk in h for tk in target_keywords):
            if not any(ek in h for ek in exclude_keywords): return i
    return None

app = FastAPI()
app.add_middleware(CORSMiddleware, allow_origins=["*"], allow_methods=["*"], allow_headers=["*"])

def get_db_conn():
    conn = sqlite3.connect(DB_PATH); conn.row_factory = sqlite3.Row
    return conn

@app.post("/upload")
async def upload_excel(file: UploadFile = File(...)):
    log_terminal(f"--- [DIRECTIONAL FORENSIC] IMPORT: {file.filename} ---")
    file_path = os.path.join(UPLOAD_DIR, f"dir_{datetime.now().strftime('%H%M%S')}_{file.filename}")
    try:
        content = await file.read()
        with open(file_path, "wb") as f: f.write(content)
        with pd.ExcelFile(file_path, engine='openpyxl') as xls:
            target_sheet = 'DanhSach' if 'DanhSach' in xls.sheet_names else xls.sheet_names[0]
            df = pd.read_excel(xls, sheet_name=target_sheet)
            df.columns = [str(h).strip() for h in df.columns]

        mapping = {
            "tracking_id": find_best_col(df.columns, ["số hiệu", "mã bưu gửi"]),
            "acceptance_date": find_best_col(df.columns, ["ngày chấp nhận", "ngày gửi"]),
            "ttp_first_date": find_best_col(df.columns, ["thời gian nhập ttp lần đầu", "thời gian nhập lần đầu"]),
            "result_first": find_best_col(df.columns, ["kết quả phát lần đầu", "kết quả lần đầu"]),
            "result_final": find_best_col(df.columns, ["kết quả phát cuối cùng", "kết quả phát lần cuối"]),
            "province": find_best_col(df.columns, ["tỉnh"], ["mã"]),
            "post_office": find_best_col(df.columns, ["bcvh", "tên bcvh", "bưu cục vận hành"])
        }

        now = datetime.now(); conn = get_db_conn(); cursor = conn.cursor()
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
                res_first = str(row.iloc[mapping["result_first"]]).lower()
                res_final = str(row.iloc[mapping["result_final"]]).lower()
                is_success = "đã phát thành công" in res_final
                aging = (now - dt_acc).days if dt_acc else 0
                is_sla = aging > 3 and "chưa có tt phát" in res_first
                processed.append((tid, str(row.iloc[mapping["province"]]), str(row.iloc[mapping["post_office"]]), str(dt_acc) if dt_acc else "", str(dt_ttp) if dt_ttp else "", res_first, res_final, 'Thành công' if is_success else 'Chưa thành công', int(aging), is_sla, sid))
            cursor.executemany('INSERT INTO orders VALUES (?,?,?,?,?,?,?,?,?,?,?)', processed)
            cursor.execute("COMMIT")
            run_directional_forensic(sid)
            return {"message": "Success", "sid": sid}
        finally: conn.close()
    except Exception as e:
        log_terminal(f"ERR: {str(e)}")
        return JSONResponse(status_code=500, content={"detail": str(e)})
    finally:
        if os.path.exists(file_path): os.remove(file_path)

@app.get("/api/dashboard/acceptance-trend")
async def get_acceptance_trend():
    conn = get_db_conn(); cursor = conn.cursor()
    try:
        session = cursor.execute("SELECT session_id FROM import_sessions ORDER BY session_id DESC LIMIT 1").fetchone()
        if not session: return {"data": []}
        sid = session['session_id']
        rows = cursor.execute("SELECT SUBSTR(acceptance_date, 1, 10) as date, COUNT(*) as total, SUM(CASE WHEN status='Thành công' THEN 1 ELSE 0 END) as success, SUM(CASE WHEN is_sla_violation=1 THEN 1 ELSE 0 END) as sla FROM orders WHERE session_id = ? AND acceptance_date != '' GROUP BY date ORDER BY date ASC", (sid,)).fetchall()
        result_data = [{"date": r['date'], "total": r['total'], "success": r['success'], "success_rate": round(r['success']*100/r['total']) if r['total']>0 else 0, "sla_rate": round(r['sla']*100/r['total']) if r['total']>0 else 0} for r in rows]
        log_terminal(f"[ACCEPTANCE_TREND_OK] sid={sid} rows={len(result_data)}")
        return {"data": result_data}
    except Exception as e: return {"data": [], "error": str(e)}
    finally: conn.close()

@app.get("/api/dashboard/stats")
async def get_stats():
    conn = get_db_conn(); cursor = conn.cursor()
    session = cursor.execute("SELECT session_id FROM import_sessions ORDER BY session_id DESC LIMIT 1").fetchone()
    if not session: return {"error": "No Data"}
    sid = session['session_id']
    kpis = cursor.execute('SELECT COUNT(*) as t, SUM(CASE WHEN status="Thành công" THEN 1 ELSE 0 END) as s, SUM(CASE WHEN is_sla_violation=1 THEN 1 ELSE 0 END) as sla FROM orders WHERE session_id=?', (sid,)).fetchone()
    
    # RE-CLASSIFY DIRECTIONS FOR STATS
    rows = cursor.execute("SELECT province, status FROM orders WHERE session_id=?", (sid,)).fetchall()
    n_tot, n_suc, s_tot, s_suc, i_tot, i_suc = 0, 0, 0, 0, 0, 0
    
    for r in rows:
        p = normalize_province_forensic(r['province'])
        is_suc = 1 if r['status'] == 'Thành công' else 0
        if any(k in p for k in ["huế", "hue"]): i_tot += 1; i_suc += is_suc
        elif p in NORTH_PROVINCES_NORM or any(np in p for np in NORTH_PROVINCES_NORM): n_tot += 1; n_suc += is_suc
        else: s_tot += 1; s_suc += is_suc
    
    conn.close()
    return {"kpis": {"total": kpis['t'], "success": kpis['s'], "pending": kpis['t'] - kpis['s'], "sla": kpis['sla']}, "directions": {"intra": {"total": i_tot, "success": i_suc}, "north": {"total": n_tot, "success": n_suc}, "south": {"total": s_tot, "success": s_suc}}}

@app.get("/api/dashboard/province-performance")
async def get_province_performance():
    conn = get_db_conn(); cursor = conn.cursor()
    session = cursor.execute("SELECT session_id FROM import_sessions ORDER BY session_id DESC LIMIT 1").fetchone()
    if not session: return []
    sid = session['session_id']
    rows = cursor.execute('SELECT province, COUNT(*) as ct, SUM(CASE WHEN status = "Thành công" THEN 1 ELSE 0 END) as success, SUM(CASE WHEN is_sla_violation = 1 THEN 1 ELSE 0 END) as sla FROM orders WHERE session_id = ? AND province != "" GROUP BY province ORDER BY ct DESC LIMIT 15', (sid,)).fetchall()
    conn.close()
    return [{"name": r['province'], "direction": "Nội tỉnh" if any(k in normalize_province_forensic(r['province']) for k in ["huế", "hue"]) else ("Bắc" if normalize_province_forensic(r['province']) in NORTH_PROVINCES_NORM or any(np in normalize_province_forensic(r['province']) for np in NORTH_PROVINCES_NORM) else "Nam"), "total": r['ct'], "success": r['success'], "sla": r['sla'], "success_rate": round(r['success']*100/r['ct']) if r['ct']>0 else 0, "sla_rate": round(r['sla']*100/r['ct']) if r['ct']>0 else 0} for r in rows]

@app.get("/api/dashboard/bcvh-summary")
async def get_bcvh_summary():
    conn = get_db_conn(); cursor = conn.cursor()
    session = cursor.execute("SELECT session_id FROM import_sessions ORDER BY session_id DESC LIMIT 1").fetchone()
    if not session: return []
    sid = session['session_id']
    rows = cursor.execute('SELECT post_office_name, province, COUNT(*) as ct, SUM(CASE WHEN status="Thành công" THEN 1 ELSE 0 END) as su, SUM(CASE WHEN is_sla_violation=1 THEN 1 ELSE 0 END) as sla FROM orders WHERE session_id = ? GROUP BY post_office_name, province ORDER BY ct DESC', (sid,)).fetchall()
    conn.close()
    return [{"name": r['post_office_name'], "province": r['province'], "total": r['ct'], "success": r['su'], "sla": r['sla'], "rate": round(r['su']*100/r['ct']) if r['ct']>0 else 0} for r in rows]

@app.get("/api/dashboard/sla-risk")
async def get_sla_risk():
    conn = get_db_conn(); cursor = conn.cursor()
    session = cursor.execute("SELECT session_id FROM import_sessions ORDER BY session_id DESC LIMIT 1").fetchone()
    if not session: return []
    sid = session['session_id']
    rows = cursor.execute('SELECT tracking_id, aging, province, post_office_name FROM orders WHERE session_id = ? AND is_sla_violation = 1 ORDER BY aging DESC LIMIT 50', (sid,)).fetchall()
    conn.close()
    return [dict(r) for r in rows]

@app.get("/api/dashboard/generate-report")
async def generate_report():
    conn = get_db_conn(); cursor = conn.cursor()
    session = cursor.execute("SELECT session_id FROM import_sessions ORDER BY session_id DESC LIMIT 1").fetchone()
    if not session: return {"error": "No Data"}
    sid = session['session_id']
    kpis = cursor.execute('SELECT COUNT(*) as t, SUM(CASE WHEN status="Thành công" THEN 1 ELSE 0 END) as s, SUM(CASE WHEN result_first LIKE "%chưa có tt phát%" THEN 1 ELSE 0 END) as p FROM orders WHERE session_id=?', (sid,)).fetchone()
    t, s, p = kpis['t'], kpis['s'], kpis['p']
    failed = t - s - p
    trends = cursor.execute("SELECT SUBSTR(acceptance_date, 1, 10) as d, COUNT(*) as ct, SUM(CASE WHEN status='Thành công' THEN 1 ELSE 0 END) as su FROM orders WHERE session_id=? AND acceptance_date != '' GROUP BY d ORDER BY d ASC", (sid,)).fetchall()
    trend_text = ", ".join([f"{r['d'][-5:]}: {round(r['su']*100/r['ct'])}%" for r in trends])
    provinces = cursor.execute("SELECT province, COUNT(*) as ct, SUM(CASE WHEN status='Thành công' THEN 1 ELSE 0 END) as su FROM orders WHERE session_id=? GROUP BY province ORDER BY su*1.0/ct DESC LIMIT 3", (sid,)).fetchall()
    top_provinces = " và ".join([r['province'] for r in provinces])
    
    rows = cursor.execute("SELECT province, status FROM orders WHERE session_id=?", (sid,)).fetchall()
    n_tot, n_suc, s_tot, s_suc, i_tot, i_suc = 0, 0, 0, 0, 0, 0
    for r in rows:
        p = normalize_province_forensic(r['province'])
        is_suc = 1 if r['status'] == 'Thành công' else 0
        if any(k in p for k in ["huế", "hue"]): i_tot += 1; i_suc += is_suc
        elif p in NORTH_PROVINCES_NORM or any(np in p for np in NORTH_PROVINCES_NORM): n_tot += 1; n_suc += is_suc
        else: s_tot += 1; s_suc += is_suc
        
    risk_0705 = cursor.execute("SELECT COUNT(*) as ct FROM orders WHERE session_id=? AND acceptance_date LIKE '%05-07%' AND result_first LIKE '%chưa có tt phát%'", (sid,)).fetchone()['ct'] or 0
    conn.close()
    ctx = {"date_now": datetime.now().strftime("%Hh%M ngày %d/%m/%Y"), "total": t, "success_count": s, "success_rate": round(s*100/t), "pending_count": p, "pending_rate": round(p*100/t), "failed_count": failed, "failed_rate": round(failed*100/t), "trend_text": trend_text, "top_provinces": top_provinces, "n_rate": round(n_suc*100/n_tot) if n_tot>0 else 0, "s_rate": round(s_suc*100/s_tot) if s_tot>0 else 0, "i_rate": round(i_suc*100/i_tot) if i_tot>0 else 0, "s_ct": s_tot, "risk_0705": risk_0705}
    with open(TEMPLATE_PATH, "r", encoding="utf-8") as f: return {"report": f.read().format(**ctx)}

@app.get("/", response_class=HTMLResponse)
async def get_index():
    with open(os.path.join(BASE_DIR, "backend", "index.html"), "r", encoding="utf-8") as f: return f.read()

if __name__ == "__main__":
    uvicorn.run(app, host="0.0.0.0", port=8010)
