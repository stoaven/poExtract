#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Web 界面: PDF 批量上传 → 解析 → 预览 → 下载 Excel

技术栈: Flask + openpyxl (复用 webapp/scripts/parse_pdf_fields.py 与 webapp/scripts/json_to_excel.py 的函数)

用法:
    python webapp.py
    浏览器打开 http://127.0.0.1:5000

接口:
    GET  /                    上传页面 + 已有文件列表
    POST /upload              多文件上传, 处理并返回 JSON
    GET  /preview/<filename>   把指定 Excel 渲染为 HTML 表格片段
    GET  /download/<filename> 下载 Excel 文件
    GET  /list                列出所有已生成的 Excel
"""

from __future__ import annotations

import json
import re
import shutil
import sys
import uuid
from pathlib import Path
from typing import Any

from flask import Flask, abort, jsonify, render_template, request, send_from_directory
from openpyxl import load_workbook
from werkzeug.utils import secure_filename

# 解析与 Excel 生成模块放在 ./webapp/scripts/, 加入 sys.path 让下面的 import 能工作
BASE_DIR = Path(__file__).resolve().parent
SCRIPTS_DIR = BASE_DIR / "webapp" / "scripts"
sys.path.insert(0, str(BASE_DIR))
sys.path.insert(0, str(SCRIPTS_DIR))

# 复用我们已有的解析与 Excel 生成模块 (位于 ./webapp/scripts/)
from parse_pdf_fields import parse_pdf  # noqa: E402
from json_to_excel import build_rows  # noqa: E402

UPLOAD_DIR = BASE_DIR / "webapp" / "uploads"
OUTPUT_DIR = BASE_DIR / "webapp" / "outputs"
ALLOWED_EXT = {".pdf"}
MAX_CONTENT_LENGTH = 64 * 1024 * 1024  # 64MB

UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

app = Flask(__name__, template_folder=str(BASE_DIR / "webapp" / "templates"))
app.config["MAX_CONTENT_LENGTH"] = MAX_CONTENT_LENGTH


# ---------- 工具函数 ----------

def _safe_name(name: str) -> str:
    """清洗上传文件名, 保留扩展名 (避免路径注入). 只允许 .pdf."""
    safe = secure_filename(name or "")
    if not safe:
        return "upload.pdf"
    # 如果清洗后丢了扩展名或扩展名不对, 强制加上 .pdf
    if Path(safe).suffix.lower() not in ALLOWED_EXT:
        stem = re.sub(r"[^\w.\-]", "_", Path(name).stem) or "file"
        return stem + ".pdf"
    return safe


def _safe_xlsx_name(name: str) -> str:
    """清洗已生成的 xlsx 文件名, 只允许 .xlsx."""
    safe = secure_filename(name or "")
    if not safe:
        return "file.xlsx"
    if not safe.lower().endswith(".xlsx"):
        # 不修改扩展名, 加 .xlsx 后缀
        return (Path(safe).stem or "file") + ".xlsx"
    return safe


def _load_excel_as_rows(xlsx_path: Path) -> list[list[Any]]:
    """把 Excel 读成二维 list 供前端 HTML 表格渲染."""
    wb = load_workbook(xlsx_path, data_only=True)
    ws = wb.active
    rows: list[list[Any]] = []
    for row in ws.iter_rows(values_only=True):
        rows.append(["" if v is None else v for v in row])
    return rows


def _save_upload(uploaded_file) -> tuple[Path, str]:
    """把 Werkzeug FileStorage 保存到 uploads/, 返回 (路径, 安全文件名)."""
    raw_name = uploaded_file.filename or "upload.pdf"
    safe = _safe_name(raw_name)
    # 防御性检查扩展名
    suffix = Path(safe).suffix.lower()
    if suffix not in ALLOWED_EXT:
        raise ValueError(f"不支持的扩展名: {suffix or '(无)'}")
    dest = UPLOAD_DIR / f"{uuid.uuid4().hex[:8]}_{safe}"
    uploaded_file.save(dest)
    return dest, safe


def _process_one(pdf_path: Path, job_id: str) -> dict[str, Any]:
    """调用 parse_pdf + json_to_excel, 生成 JSON 与 Excel, 返回结果 dict."""
    out_dir = OUTPUT_DIR / job_id
    out_dir.mkdir(parents=True, exist_ok=True)

    parsed = parse_pdf(pdf_path)
    json_path = out_dir / (pdf_path.stem + ".json")
    json_path.write_text(
        json.dumps(parsed, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )

    rows = build_rows(parsed)
    xlsx_path = out_dir / (pdf_path.stem + ".xlsx")
    # 直接调 write_excel 的逻辑 (从 json_to_excel 导入)
    from json_to_excel import write_excel
    write_excel(rows, xlsx_path)

    return {
        "name": pdf_path.stem,
        "display_name": pdf_path.name,
        "json_path": str(json_path.relative_to(BASE_DIR)),
        "xlsx_name": xlsx_path.name,
        "rows": len(rows) - 1,
        "page_count": parsed.get("page_count"),
    }


# ---------- 路由 ----------

@app.route("/")
def index():
    return render_template("index.html", max_size_mb=MAX_CONTENT_LENGTH // (1024 * 1024))


@app.route("/upload", methods=["POST"])
def upload():
    if "files" not in request.files:
        return jsonify({"error": "未发现上传文件 (字段名应为 files)"}), 400
    files = request.files.getlist("files")
    if not files:
        return jsonify({"error": "请至少选择一个 PDF"}), 400

    job_id = uuid.uuid4().hex[:12]
    results: list[dict[str, Any]] = []
    errors: list[dict[str, str]] = []

    for f in files:
        try:
            saved, safe = _save_upload(f)
            res = _process_one(saved, job_id)
            res["original"] = f.filename
            results.append(res)
        except Exception as e:  # noqa: BLE001
            errors.append({"file": f.filename or "?", "error": str(e)})

    return jsonify({
        "job_id": job_id,
        "results": results,
        "errors": errors,
    })


@app.route("/preview/<job_id>/<path:xlsx_name>")
def preview(job_id: str, xlsx_name: str):
    """把指定 Excel 渲染为 HTML 表格片段."""
    safe_xlsx = _safe_xlsx_name(xlsx_name)
    xlsx_path = OUTPUT_DIR / job_id / safe_xlsx
    if not xlsx_path.exists():
        abort(404)

    rows = _load_excel_as_rows(xlsx_path)
    return render_template(
        "preview.html",
        name=safe_xlsx.replace(".xlsx", ""),
        job_id=job_id,
        xlsx_name=safe_xlsx,
        rows=rows,
    )


@app.route("/download/<job_id>/<path:xlsx_name>")
def download(job_id: str, xlsx_name: str):
    safe_xlsx = _safe_xlsx_name(xlsx_name)
    directory = OUTPUT_DIR / job_id
    if not (directory / safe_xlsx).exists():
        abort(404)
    return send_from_directory(
        directory, safe_xlsx, as_attachment=True,
        download_name=safe_xlsx,
    )


@app.route("/download-all/<job_id>")
def download_all(job_id: str):
    """批量下载某个 job 的所有 Excel (打成 zip)."""
    import zipfile
    import io

    directory = OUTPUT_DIR / job_id
    if not directory.exists() or not any(directory.glob("*.xlsx")):
        abort(404)

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for xlsx in directory.glob("*.xlsx"):
            zf.write(xlsx, arcname=xlsx.name)
    buf.seek(0)

    from flask import send_file
    return send_file(
        buf,
        mimetype="application/zip",
        as_attachment=True,
        download_name=f"excel_{job_id}.zip",
    )


@app.route("/list")
def list_jobs():
    """列出所有已生成的 job (按修改时间排序)."""
    jobs = []
    if OUTPUT_DIR.exists():
        for d in sorted(OUTPUT_DIR.iterdir(), key=lambda p: p.stat().st_mtime, reverse=True):
            if not d.is_dir():
                continue
            xlsxs = sorted(d.glob("*.xlsx"))
            if not xlsxs:
                continue
            jobs.append({
                "job_id": d.name,
                "mtime": d.stat().st_mtime,
                "files": [x.name for x in xlsxs],
            })
    return jsonify({"jobs": jobs})


# ---------- 启动 ----------

if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser()
    parser.add_argument("--host", default="127.0.0.1")
    parser.add_argument("--port", type=int, default=5000)
    parser.add_argument("--debug", action="store_true")
    args = parser.parse_args()
    print(f"打开浏览器访问: http://{args.host}:{args.port}")
    app.run(host=args.host, port=args.port, debug=args.debug)
