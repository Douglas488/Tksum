# -*- coding: utf-8 -*-
"""
Upseller 月营业额报表 - Flask Web API
接受一个 ZIP（内含多个 .xlsx），汇总后返回「Upseller月营业额报表.xlsx」。
"""
import os
import tempfile
import zipfile
from io import BytesIO
from pathlib import Path

from flask import Flask, request, send_file, jsonify

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 80 * 1024 * 1024  # 80MB


def _cors_headers(resp):
    resp.headers["Access-Control-Allow-Origin"] = "*"
    resp.headers["Access-Control-Allow-Methods"] = "GET, POST, OPTIONS"
    resp.headers["Access-Control-Allow-Headers"] = "Content-Type"
    return resp


@app.after_request
def after_request(resp):
    return _cors_headers(resp)


def _get_run_merge():
    from uptotal_core import run_merge
    return run_merge


@app.route("/api/generate", methods=["POST", "OPTIONS"])
def api_generate():
    if request.method == "OPTIONS":
        return _cors_headers(app.make_default_options_response())

    if "file" not in request.files:
        return jsonify({"error": "请上传 ZIP 文件（字段名: file），ZIP 内包含多个 .xlsx"}), 400

    f = request.files["file"]
    if not f.filename or not f.filename.lower().endswith(".zip"):
        return jsonify({"error": "仅支持 .zip 文件，且 ZIP 内需包含 .xlsx 文件"}), 400

    try:
        run_merge = _get_run_merge()
    except Exception as e:
        return jsonify({"error": "服务未就绪: " + str(e)}), 500

    try:
        with tempfile.TemporaryDirectory() as tmp:
            tmp = Path(tmp)
            zip_path = tmp / "upload.zip"
            f.save(str(zip_path))
            with zipfile.ZipFile(zip_path, "r") as z:
                z.extractall(tmp)
            # 若 ZIP 根目录无 xlsx 但仅有一个子文件夹且其中有 xlsx，则用该子文件夹作为汇总目录
            merge_dir = tmp
            xlsx_in_tmp = list(tmp.glob("*.xlsx"))
            if not xlsx_in_tmp:
                entries = [p for p in tmp.iterdir() if p.is_dir()]
                if len(entries) == 1 and list(entries[0].glob("*.xlsx")):
                    merge_dir = entries[0]
            xlsx_count = len(list(merge_dir.glob("*.xlsx")))
            if xlsx_count == 0:
                return jsonify({"error": "ZIP 内未找到 .xlsx 文件，请将多个店铺的 Excel 放入 ZIP 后上传"}), 400
            out_name = "Upseller月营业额报表.xlsx"
            output_path = run_merge(str(merge_dir), output_filename=out_name)
            buf = BytesIO(Path(output_path).read_bytes())
            buf.seek(0)
            return _cors_headers(
                send_file(
                    buf,
                    as_attachment=True,
                    download_name=out_name,
                    mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            )
    except ValueError as e:
        return jsonify({"error": str(e)}), 400
    except FileNotFoundError as e:
        return jsonify({"error": str(e)}), 400
    except Exception as e:
        return jsonify({"error": "生成失败: " + str(e)}), 500


@app.route("/")
def index():
    info = {
        "service": "Upseller月营业额报表 API",
        "usage": "POST /api/generate，上传一个 ZIP 文件（内含多个 .xlsx），返回汇总 Excel",
    }
    try:
        _get_run_merge()
        info["status"] = "ok"
    except Exception as e:
        info["status"] = "failed"
        info["error"] = str(e)
    return jsonify(info)
