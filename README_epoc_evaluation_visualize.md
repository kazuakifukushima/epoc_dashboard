# EPOC評価 可視化スクリプトの使い方

## 概要
「研修医評価」タブと「指導医評価」タブを持つExcelを読み込み、
研修医別・診療科別・評価者別に見やすく整理したExcelを出力します。

## 1. 必要ライブラリ
```bash
pip install pandas openpyxl
```

## 2. 実行
```bash
python epoc_evaluation_visualize.py EPOC評価ファイル.xlsx
```

## 3. 出力
`EPOC評価ファイル_evaluation_visualized.xlsx`

## 4. 出力シート
- `raw_研修医評価`
- `raw_指導医評価`
- `long_研修医評価`
- `long_指導医評価`
- `resident_summary`
  - 研修医別の平均点、A/B/C群平均、低評価件数
- `department_summary`
  - 診療科別・施設別・評価元別の平均点
- `evaluator_summary`
  - 指導医別の平均点傾向
- `radar_source`
  - 研修医×評価元×A/B/C群 の平均値
  - Excel側でレーダーチャートや棒グラフの元データに使えます
- `alerts`
  - 2点未満、0点のアラート一覧
- `dashboard`
  - 管理者向けの簡易ダッシュボード

## 5. この形式で見やすくなるポイント
- 研修医ごとの弱点領域が分かる
- A/B/C群ごとの傾向が分かる
- 診療科や施設ごとの評価傾向が見える
- 研修医評価と指導医評価を並べて比較しやすい
- 低評価項目だけを抜き出して面談資料にしやすい

## 6. 発展案
- 研修医評価と指導医評価の差分分析
- 年度別比較
- ローテーション期間ごとの時系列推移
- Plotly/Streamlit/Next.jsでのWebダッシュボード化
