# EPOC Auto Visualizer UI (Next.js)

## 概要
ExcelアップロードUIの最小サンプルです。
サーバー側で Python の `epoc_auto_visualizer.py` を呼び出し、
可視化済みExcelをダウンロードさせます。

## 前提
- Node.js
- Python
- `../python/epoc_auto_visualizer.py` が存在すること
- Python 側で `pandas`, `openpyxl` がインストールされていること

## セットアップ
```bash
npm install
```

## 開発起動
```bash
npm run dev
```

## 補足
`app/api/visualize/route.ts` 内の `scriptPath` は、
配置場所に応じて調整してください。
