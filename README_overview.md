# EPOC Auto Visualizer Package

このパッケージには2つ入っています。

## 1. Python統合版
`python/epoc_auto_visualizer.py`

- Excelファイルを自動判定
- 経験症候・疾患ファイルか評価ファイルかを判定
- 可視化済みExcelを出力

## 2. Next.js最小UI
`nextjs-app/`

- Excelアップロード画面
- API route から Python を実行
- 出力されたExcelをダウンロード

## 推奨導入順
1. まず Python 単体で動作確認
2. 次に Next.js UI を使ってブラウザ化
3. 必要に応じて SQLite / Supabase 保存や年度比較を追加
