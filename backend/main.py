import os
import sqlite3
import pandas as pd
from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import HTMLResponse
from datetime import datetime
import uvicorn
import sys

BASE_DIR = r"d:\Antigravity - Project - TTVH\CSKH"
ORIGINAL_FILE = os.path.join(BASE_DIR, "File CSKH.xlsx")
DB_PATH = os.path.join(BASE_DIR, "cskh_vip.db")

app = FastAPI(title="VNPost Hue VIP Dashboard")
app.add_middleware(CORSMiddleware, allow_origins=["*"], allow_methods=["*"], allow_headers=["*"])

@app.post("/upload")
async def upload_excel(file: UploadFile = File(...)):
    try:
        df_dict = pd.read_excel(ORIGINAL_FILE, sheet_name=None, engine='openpyxl')
        sheet_name = 'DanhSach' if 'DanhSach' in df_dict.keys() else list(df_dict.keys())[0]
        df = df_dict[sheet_name]
        
        headers = [str(h).lower() for h in df.columns]
        idx_res_final = 8
        for i, h in enumerate(headers):
            if "kết quả" in h and "cuối" in h:
                idx_res_final = i
                break

        customer_info = str(df.iloc[0, 0])
        customer_id = customer_info.split('-')[0].strip() if '-' in customer_info else "UNKNOWN"
        customer_name = customer_info.split('-')[1].strip() if '-' in customer_info else customer_info

        conn = sqlite3.connect(DB_PATH)
        conn.execute("DELETE FROM orders WHERE customer_id = ?", (customer_id,))
        
        processed_orders = []
        now = datetime.now()

        for idx, row in df.iterrows():
            if idx < 1: continue 
            tid = str(row.iloc[0]).strip()
            if not tid or len(tid) < 5 or tid.lower() == 'nan': continue
            
            res_val = str(row.iloc[idx_res_final]).lower() if len(row) > idx_res_final else ""
            is_success = ("thành công" in res_val or "đã phát" in res_val) and ("không" not in res_val)
            status = 'Thành công' if is_success else 'Chưa thành công'
            
            aging = 0
            try:
                acc_val = row.iloc[5]
                if acc_val and not pd.isna(acc_val):
                    aging = (now - pd.to_datetime(acc_val)).days
            except: pass

            order_data = (tid, customer_id, str(row.iloc[1]), "", str(row.iloc[3]), str(row.iloc[4]), None, "", "", str(row.iloc[idx_res_final]), "", str(row.iloc[10]), "", "", status, max(0, int(aging)), (aging > 3 and not is_success))
            processed_orders.append(order_data)

        conn.execute("INSERT OR REPLACE INTO customers (customer_id, customer_name, last_import) VALUES (?, ?, ?)", (customer_id, customer_name, datetime.now().isoformat()))
        conn.executemany('INSERT OR REPLACE INTO orders VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)', processed_orders)
        conn.commit()
        conn.close()
        return {"message": "Success", "customer_id": customer_id}
    except Exception as e:
        return {"error": str(e)}

@app.get("/stats")
async def get_stats():
    try:
        conn = sqlite3.connect(DB_PATH)
        conn.row_factory = sqlite3.Row
        cursor = conn.cursor()
        cust = cursor.execute("SELECT customer_id, customer_name FROM customers ORDER BY last_import DESC LIMIT 1").fetchone()
        if not cust: return {"error": "No data"}
        cid = cust['customer_id']
        kpis = cursor.execute('SELECT COUNT(*) as total, SUM(CASE WHEN status="Thành công" THEN 1 ELSE 0 END) as success, SUM(CASE WHEN is_sla_violation=1 THEN 1 ELSE 0 END) as sla FROM orders WHERE customer_id=?', (cid,)).fetchone()
        prov_rows = cursor.execute('SELECT province, COUNT(*) as total, SUM(CASE WHEN status = "Thành công" THEN 1 ELSE 0 END) as success FROM orders WHERE customer_id = ? GROUP BY province ORDER BY total DESC', (cid,)).fetchall()
        
        province_summary = []
        radar_data = []
        for r in prov_rows:
            rate = round((r['success'] / r['total']) * 100) if r['total'] > 0 else 0
            name = r['province'] or "N/A"
            province_summary.append({"province": name, "total": r['total'], "success": r['success'], "unsuccessful": r['total']-r['success'], "rate": rate})
            if len(radar_data) < 8:
                radar_data.append({"subject": name, "A": rate, "fullMark": 100})

        po_rows = cursor.execute('SELECT post_office_name, COUNT(*) as backlog, MAX(aging) as max_aging FROM orders WHERE customer_id = ? AND status != "Thành công" GROUP BY post_office_name ORDER BY max_aging DESC LIMIT 15', (cid,)).fetchall()
        bottlenecks = [{"name": r[0], "backlog": r[1], "max_aging": r[2], "sla": 0, "top_reason": "Đang xử lý"} for r in po_rows]
        sla_rows = cursor.execute('SELECT tracking_id as id, aging, province, post_office_name as bcvh FROM orders WHERE customer_id=? AND is_sla_violation=1 ORDER BY aging DESC LIMIT 50', (cid,)).fetchall()
        conn.close()
        return {"customer": {"name": cust['customer_name']}, "kpis": {"total": kpis['total'], "success": kpis['success'], "pending": kpis['total'] - kpis['success'], "sla": kpis['sla']}, "provinceSummary": province_summary, "radarData": radar_data, "bottlenecks": bottlenecks, "slaList": [dict(s) for s in sla_rows]}
    except Exception as e:
        return {"error": str(e)}

@app.get("/", response_class=HTMLResponse)
async def get_index():
    with open(os.path.join(BASE_DIR, "backend", "index.html"), "r", encoding="utf-8") as f: return f.read()

if __name__ == "__main__":
    uvicorn.run(app, host="0.0.0.0", port=8088)
