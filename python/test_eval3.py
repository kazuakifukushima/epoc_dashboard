import pandas as pd
excel = pd.ExcelFile("/Users/fukushimakazuaki/cursor/epoc_dashboard/P000005919_研修医評価票_example.xlsx")
df = pd.read_excel(excel, sheet_name=0, header=1) # Row 1 is header
df = df.dropna(subset=['研修医氏名'])
print("Shape:", df.shape)
print("Columns:", list(df.columns)[:15])
print(df[['研修医氏名', '評価者氏名', '指導医／上級医区分']].head(10).to_string())
