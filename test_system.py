import pandas as pd
import os
import sys

# Ensure UTF-8 output
if sys.stdout.encoding != 'UTF-8':
    import codecs
    sys.stdout = codecs.getwriter('utf-8')(sys.stdout.buffer, 'strict')

file_path = r"d:\Antigravity - Project - TTVH\CSKH\File CSKH.xlsx"
print(f"--- TESTING FILE: {file_path} ---")

try:
    # Use pandas read_excel
    df_dict = pd.read_excel(file_path, sheet_name=None, engine='openpyxl')
    sheet_name = 'DanhSach' if 'DanhSach' in df_dict.keys() else list(df_dict.keys())[0]
    df = df_dict[sheet_name]
    
    print(f"Total Rows: {len(df)}")
    
    success_count = 0
    fail_sample = []

    for idx, row in df.iterrows():
        if idx < 1: continue 
        
        # Merge all columns into one string
        row_str = " ".join([str(v).lower() for v in row.values if v and not pd.isna(v)])
        
        # Check success keywords
        is_success = any(x in row_str for x in ["thành công", "thanh cong", "đã phát", "da phat"])
        
        if is_success:
            success_count += 1
        else:
            if len(fail_sample) < 10 and len(row_str) > 20:
                fail_sample.append(f"Row {idx+2}: {row_str}")

    print(f"==> FINAL SUCCESS COUNT: {success_count}")
    
    if success_count < 200:
        print("\n--- SAMPLES OF NON-SUCCESS ROWS ---")
        for s in fail_sample:
            print(s)

except Exception as e:
    print(f"ERROR: {str(e)}")
