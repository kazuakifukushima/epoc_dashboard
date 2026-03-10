# EPOC Auto Visualizer

EPOC由来の類似Excelファイルを自動判定し、可視化済みExcelとして出力するツールです。

## 使い方の流れ

### パターンA: サーバー起動（Excel をその場でアップロード）

> Excel アップロードはブラウザ内で変換するため、サーバー不要で動作します（病歴要約等・症例入力状況一覧・研修医評価票に対応）。

1. **準備**
   ```bash
   cd nextjs-app && npm install
   cd ../python && pip install -r requirements.txt
   ```

2. **起動**
   ```bash
   cd nextjs-app && npm run dev
   ```

3. **ブラウザで http://localhost:3001 を開く**

4. **操作**
   - ファイル形式（経験症候・疾患 / 研修医評価票 / 病歴要約等）を選択
   - Excel ファイル（.xlsx）をドラッグ＆ドロップまたは選択
   - 「ダッシュボードを生成する」をクリック
   - ダッシュボードが表示される

---

### パターンB: 静的サイト（Excel をその場で変換）

1. **ビルド＆起動**
   ```bash
   cd nextjs-app
   npm install && npm run build:static
   npm run serve:static
   ```

2. **ブラウザで http://localhost:8080 を開く**

3. **操作**
   - 病歴要約等・症例入力状況一覧・研修医評価票のいずれかの Excel をアップロード
   - ダッシュボードが表示される（ブラウザ内で変換、サーバー不要）

---

### パターンC: Netlify + Render（Excel アップロードでダッシュボード表示）

アプリを Netlify に、Excel 処理 API を Render にデプロイし、ブラウザから Excel をアップロードしてダッシュボードを表示します。

#### 1. Render に API をデプロイ

1. [Render](https://render.com) にログイン
2. New → Web Service
3. リポジトリを接続（epoc_dashboard を push した GitHub 等）
4. 設定:
   - **Root Directory**: `epoc_dashboard`（リポジトリ直下の場合は空）
   - **Build Command**: `pip install -r api/requirements.txt && pip install -r python/requirements.txt`
   - **Start Command**: `uvicorn api.main:app --host 0.0.0.0 --port $PORT`
5. Deploy → デプロイ後に表示される URL（例: `https://epoc-visualize-api.onrender.com`）を控える

#### 2. Netlify にフロントをデプロイ

1. [Netlify](https://netlify.com) にログイン
2. Add new site → Import an existing project
3. リポジトリを接続
4. 設定（netlify.toml が読み込まれる）:
   - **Base directory**: `epoc_dashboard` または `nextjs-app`（リポジトリ構成に応じて）
   - **Build command**: `npm run build:static`
   - **Publish directory**: `out`
5. **Environment variables** を追加:
   - `NEXT_PUBLIC_VISUALIZE_API_URL` = `https://epoc-visualize-api.onrender.com`（上記の Render URL）
6. Deploy

#### 3. 使い方

- Netlify の URL を開く
- ファイル形式を選択し、Excel をアップロード
- ダッシュボードが表示される

> **注意**: Render の無料プランはコールドスタートで初回アクセスが遅くなることがあります。

---

## 対応ファイル
- 経験症候・疾患ファイル
- 評価ファイル（研修医評価 / 指導医評価）

## 機能
- ファイル形式自動判定
- 研修医別サマリー作成
- 未経験 / 低評価アラート抽出
- ヒートマップ用データ生成
- ダッシュボードシート出力
- Next.js経由のアップロードUI

## 使い方
### Python
```bash
cd python
pip install -r requirements.txt
python epoc_auto_visualizer.py input.xlsx
```

### Next.js
```bash
cd nextjs-app
npm install
npm run dev
```

## 静的サイトのデプロイ

`out/` フォルダを任意の静的ホスティングにアップロードできます。

### デプロイ例
- GitHub Pages: `out/` を `gh-pages` ブランチに push
- Netlify / Vercel: ビルドコマンド `npm run build:static`、公開ディレクトリ `out`
- 任意の静的ホスティング: `out/` 内のファイルをアップロード

## 今後の拡張
- 経験症候・疾患と評価の統合分析
- 匿名化ファイルの擬似ID推定
- Webプレビュー
- 年度比較
