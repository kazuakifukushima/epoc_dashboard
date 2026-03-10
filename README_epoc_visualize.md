# EPOC可視化スクリプトの使い方

## 1. 必要ライブラリ
```bash
pip install pandas openpyxl
```

## 2. 実行
```bash
python epoc_visualize.py EPOC元ファイル.xlsx
```

## 3. 出力されるExcel
`EPOC元ファイル_visualized.xlsx`

### 含まれるシート
- `raw_all`
  - 全研修医のデータを縦持ちで統合した表
- `summary_resident`
  - 研修医別の達成率、承認率、未提出件数、要対応件数
- `alert_list`
  - 未経験、承認待ち、病歴要約未提出、未確認などの一覧
- `heatmap_symptom`
  - 症候の経験数ヒートマップ
- `heatmap_disease`
  - 疾患の経験数ヒートマップ
- `dashboard`
  - 管理者向けの簡易一覧

## 4. このスクリプトの前提
- 各研修医が1シート
- シート上部に `研修医氏名`, `研修医UMIN ID`
- `経験すべき症候` と `経験すべき疾患` ブロックがある
- 見出し名が大きく変わっていない

## 5. カスタマイズしやすい点
- アラート条件の追加
- 達成率の計算式の変更
- シート名や列名の微修正
- 将来的なCSV出力やWebダッシュボード連携

## 6. 次の発展
- 月次比較
- 学年別比較
- ローテーション科別集計
- Streamlit / Next.jsでのWebダッシュボード化
