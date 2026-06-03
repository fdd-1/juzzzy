"""检查报表2的表头行"""
import sys
sys.stdout.reconfigure(encoding="utf-8")
import openpyxl

path = r"C:\Users\fengjianyi\Desktop\方法积分核算\01_bi_exports\海外思维学员上课明细.xlsx"
wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
ws = wb.active

row_idx = 0
for row in ws.iter_rows(min_row=8, max_row=12, values_only=False):
    row_idx += 1
    actual_row = 7 + row_idx
    vals = [(c.column, c.value) for c in row if c.value is not None]
    print(f"R{actual_row} ({len(vals)} non-null): {vals[:25]}")

wb.close()
