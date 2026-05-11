import pandas as pd
import sys

# Ép dùng UTF-8
if sys.stdout.encoding != 'UTF-8':
    import codecs
    sys.stdout = codecs.getwriter('utf-8')(sys.stdout.buffer, 'strict')

file_path = r"d:\Antigravity - Project - TTVH\CSKH\File CSKH.xlsx"
print(f"--- ĐANG KIỂM TRA FILE: {file_path} ---")

try:
    xl = pd.ExcelFile(file_path, engine='openpyxl')
    sheet_name = 'DanhSach' if 'DanhSach' in xl.sheet_names else xl.sheet_names[0]
    df = pd.read_excel(xl, sheet_name=sheet_name, header=1)
    
    print(f"Sheet đang đọc: {sheet_name}")
    print(f"Danh sách cột: {df.columns.tolist()}")
    
    # Tìm cột Kết quả phát cuối cùng
    target_col = None
    for c in df.columns:
        if "kết quả phát cuối cùng" in str(c).lower():
            target_col = c
            break
            
    if target_col:
        print(f"==> Đã tìm thấy cột: {target_col}")
        # Đếm số dòng chứa "thành công"
        success_rows = df[df[target_col].astype(str).str.contains("thành công", case=False, na=False)]
        print(f"SỐ DÒNG THÀNH CÔNG ĐẾM ĐƯỢC: {len(success_rows)}")
        
        print("\n--- 20 DÒNG ĐẦU TIÊN CỦA CỘT NÀY ---")
        for i, val in enumerate(df[target_col].head(20)):
            print(f"Dòng {i+2}: {val}")
            
    else:
        print("!!! KHÔNG TÌM THẤY CỘT 'Kết quả phát cuối cùng'")

except Exception as e:
    print(f"LỖI KHI ĐỌC FILE: {str(e)}")
