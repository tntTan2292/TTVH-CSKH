import pandas as pd
import sys

sys.stdout.reconfigure(encoding='utf-8')
file_path = r'd:\Antigravity - Project - TTVH\CSKH\File CSKH.xlsx'
df = pd.read_excel(file_path, sheet_name='DanhSach', header=None)

for i, row in df.iterrows():
    if any('Số hiệu BG' in str(cell) for cell in row):
        print(f'Header found at row {i}')
        print('|'.join([str(x) for x in row]))
        # Print sample data from next few rows
        for j in range(i + 1, min(i + 6, len(df))):
            print('|'.join([str(x) for x in df.iloc[j]]))
        break
