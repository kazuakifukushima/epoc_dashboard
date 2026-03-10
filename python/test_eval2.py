from openpyxl import load_workbook
from epoc_auto_visualizer import find_header_row, cell_text
import pandas as pd

wb = load_workbook("/Users/fukushimakazuaki/cursor/epoc_dashboard/P000005919_研修医評価票_example.xlsx", data_only=True)
ws = wb.worksheets[0]
header_row = find_header_row(ws)
print("header_row:", header_row)

data = list(ws.values)
header = [cell_text(v) for v in data[header_row - 1]]
rows = data[header_row:]
df = pd.DataFrame(rows, columns=header)
print("Shape before dropna:", df.shape)
df = df.dropna(how="all").copy()
print("Shape after dropna:", df.shape)
print("Head:")
print(df[["研修医氏名", "研修医UMIN ID", "評価者氏名"]].head())
