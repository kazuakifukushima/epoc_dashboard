from epoc_auto_visualizer import read_all_eval_sheets
from pathlib import Path

dfs = read_all_eval_sheets(Path("/Users/fukushimakazuaki/cursor/epoc_dashboard/P000005919_研修医評価票_example.xlsx"))
for k, df in dfs.items():
    print(f"Sheet: {k}")
    print(f"Columns: {list(df.columns)[:10]}")
    print(df.head(2).to_string())
