# EPOC Auto Visualizer (Python)

## 概要
類似のEPOC Excelファイルを自動判定し、可視化済みExcelを出力します。

対応形式:
- 経験症候・疾患ファイル
- 評価ファイル（研修医評価 / 指導医評価）

## セットアップ
```bash
pip install -r requirements.txt
```

## 実行
```bash
python epoc_auto_visualizer.py input.xlsx
```

## 出力
- 経験症候・疾患ファイル: `*_visualized.xlsx`
- 評価ファイル: `*_evaluation_visualized.xlsx`

## 備考
列名やシート構造に揺れがある場合は、`detect_workbook_type` と各 parser を調整してください。
