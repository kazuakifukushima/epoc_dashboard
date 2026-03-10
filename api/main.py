#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
EPOC 可視化 API
Excel ファイルを受け取り、ダッシュボード用 JSON を返す。
Netlify にデプロイしたフロントから呼び出される。
"""

import json
import shutil
import subprocess
import tempfile
from pathlib import Path

from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware

app = FastAPI(title="EPOC Visualize API")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # 本番では Netlify ドメイン等に制限推奨
    allow_credentials=True,
    allow_methods=["POST", "OPTIONS"],
    allow_headers=["*"],
)


def run_visualizer(input_path: Path, file_type: str | None) -> tuple[Path, Path]:
    """Python スクリプトを実行し、出力パスを取得"""
    script_dir = Path(__file__).resolve().parent.parent / "python"
    script_path = script_dir / "epoc_auto_visualizer.py"
    if not script_path.exists():
        raise FileNotFoundError(f"スクリプトが見つかりません: {script_path}")

    python_cmd = "python3" if shutil.which("python3") else "python"
    args = [python_cmd, str(script_path), str(input_path)]
    if file_type:
        args.extend(["--type", file_type])

    result = subprocess.run(
        args,
        capture_output=True,
        text=True,
        cwd=str(script_dir),
        timeout=120,
    )

    if result.returncode != 0:
        raise RuntimeError(result.stderr or result.stdout or "処理に失敗しました")

    output_line = next((l for l in result.stdout.split("\n") if l.startswith("output=")), None)
    json_line = next((l for l in result.stdout.split("\n") if l.startswith("json_output=")), None)

    if not output_line or not json_line:
        raise RuntimeError("出力パスを取得できませんでした")

    output_path = Path(output_line.replace("output=", "").strip())
    json_path = Path(json_line.replace("json_output=", "").strip())

    return output_path, json_path


@app.post("/visualize")
async def visualize(
    file: UploadFile = File(...),
    fileType: str | None = Form(None),
):
    """Excel ファイルをアップロードし、ダッシュボード用 JSON を返す"""
    if not file.filename or not file.filename.lower().endswith(".xlsx"):
        raise HTTPException(status_code=400, detail=".xlsx ファイルをアップロードしてください")

    content = await file.read()
    if len(content) > 20 * 1024 * 1024:  # 20MB 制限
        raise HTTPException(status_code=400, detail="ファイルサイズは 20MB 以下にしてください")

    ft = fileType if fileType in ("symptom_disease", "evaluation", "case_summary") else None

    with tempfile.TemporaryDirectory() as tmpdir:
        input_path = Path(tmpdir) / file.filename
        input_path.write_bytes(content)

        try:
            output_path, json_path = run_visualizer(input_path, ft)
        except FileNotFoundError as e:
            raise HTTPException(status_code=500, detail=str(e))
        except RuntimeError as e:
            raise HTTPException(status_code=500, detail=str(e))

        if not json_path.exists():
            raise HTTPException(status_code=500, detail="JSON 出力が見つかりません")

        dashboard_data = json.loads(json_path.read_text(encoding="utf-8"))

        # Excel も返す（ダウンロード用）
        excel_base64 = None
        filename = None
        if output_path.exists():
            import base64
            excel_base64 = base64.b64encode(output_path.read_bytes()).decode("ascii")
            filename = output_path.name

    return {
        "filename": filename,
        "excelBase64": excel_base64,
        "dashboardData": dashboard_data,
    }


@app.get("/health")
async def health():
    return {"status": "ok"}
