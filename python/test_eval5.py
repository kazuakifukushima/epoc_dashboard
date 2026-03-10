from epoc_auto_visualizer import read_all_eval_sheets, make_eval_department_summary, to_long_eval
from pathlib import Path
import pandas as pd

dfs = read_all_eval_sheets(Path("/Users/fukushimakazuaki/cursor/epoc_dashboard/P000005919_研修医評価票_example.xlsx"))
for k, df in dfs.items():
    df.rename(columns={
        "研修施設": "施設名",
        "診療科": "診療科名",
        "評価者氏名": "指導医氏名",
        "評価者UMIN ID": "指導医UMIN ID"
    }, inplace=True)
    long_df = to_long_eval(df, k)
    dep_sum = make_eval_department_summary(long_df)
    print("Department Summary:")
    print(dep_sum.head())
