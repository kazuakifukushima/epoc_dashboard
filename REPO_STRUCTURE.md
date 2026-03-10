# 推奨リポジトリ構成

```text
epoc-auto-visualizer/
├─ README.md
├─ PRD.md
├─ CLAUDE.md
├─ .gitignore
├─ python/
│  ├─ requirements.txt
│  ├─ epoc_auto_visualizer.py
│  ├─ parsers/
│  │  ├─ __init__.py
│  │  ├─ detector.py
│  │  ├─ symptom_disease.py
│  │  └─ evaluation.py
│  ├─ renderers/
│  │  ├─ __init__.py
│  │  ├─ excel_writer.py
│  │  └─ dashboard_builders.py
│  ├─ utils/
│  │  ├─ __init__.py
│  │  ├─ excel_helpers.py
│  │  └─ pseudo_id.py
│  └─ tests/
│     ├─ test_detector.py
│     ├─ test_symptom_disease.py
│     ├─ test_evaluation.py
│     └─ fixtures/
│        ├─ symptom_example.xlsx
│        └─ evaluation_example.xlsx
├─ nextjs-app/
│  ├─ package.json
│  ├─ tsconfig.json
│  ├─ app/
│  │  ├─ layout.tsx
│  │  ├─ page.tsx
│  │  └─ api/
│  │     └─ visualize/
│  │        └─ route.ts
│  ├─ components/
│  │  ├─ UploadForm.tsx
│  │  ├─ SummaryCard.tsx
│  │  └─ AlertTable.tsx
│  └─ lib/
│     └─ types.ts
└─ docs/
   ├─ sample_files.md
   ├─ deployment.md
   └─ changelog.md
```

## 最小構成からの拡張順
1. `python/epoc_auto_visualizer.py` で単体動作
2. parser を `parsers/` に分離
3. writer を `renderers/` に分離
4. Next.js でアップロードUI提供
5. tests を追加
6. docs を整理
