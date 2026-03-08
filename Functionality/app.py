# -*- coding: utf-8 -*-
"""
Functionality 统一 Web API：4 个功能共用一个服务，供 index12–index15 调用。
- POST /api/purchasing   -> 采购 Excel 转 JSON（采购信息）
- POST /api/sku-pescar   -> 库存导 Excel 为 JSON
- POST /api/export-excel-json -> 新品 Nx Excel 转 JSON
- POST /api/empalagem    -> 包裹尺寸 Excel 转 JSON
"""
import os
import tempfile
from pathlib import Path
from flask import Flask, request, jsonify

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024  # 50MB


def _cors(resp):
    resp.headers["Access-Control-Allow-Origin"] = "*"
    resp.headers["Access-Control-Allow-Methods"] = "GET, POST, OPTIONS"
    resp.headers["Access-Control-Allow-Headers"] = "Content-Type"
    return resp


@app.after_request
def after_request(resp):
    return _cors(resp)


def _get_file():
    if "file" not in request.files:
        return None, "请上传文件（字段名: file）"
    f = request.files["file"]
    if not f.filename:
        return None, "未选择文件"
    return f, None


@app.route("/")
def index():
    return _cors(jsonify({
        "service": "Functionality API",
        "endpoints": [
            "POST /api/purchasing — 采购Excel转JSON",
            "POST /api/sku-pescar — 库存导Excel为JSON",
            "POST /api/export-excel-json — 新品Nx",
            "POST /api/empalagem — 包裹尺寸",
        ],
    }))


@app.route("/api/purchasing", methods=["POST", "OPTIONS"])
def api_purchasing():
    if request.method == "OPTIONS":
        return _cors(app.make_default_options_response())
    f, err = _get_file()
    if err:
        return jsonify({"error": err}), 400
    try:
        from purchasing_core import extract_hyperlinks_from_excel
    except Exception as e:
        return jsonify({"error": "服务未就绪: " + str(e)}), 500
    with tempfile.NamedTemporaryFile(suffix=Path(f.filename).suffix or ".xlsx", delete=False) as tmp:
        f.save(tmp.name)
        try:
            data = extract_hyperlinks_from_excel(tmp.name)
            if data is None:
                return jsonify({"error": "解析 Excel 失败或文件格式不正确"}), 400
            return _cors(jsonify(data))
        finally:
            try:
                os.unlink(tmp.name)
            except Exception:
                pass


@app.route("/api/sku-pescar", methods=["POST", "OPTIONS"])
def api_sku_pescar():
    if request.method == "OPTIONS":
        return _cors(app.make_default_options_response())
    f, err = _get_file()
    if err:
        return jsonify({"error": err}), 400
    try:
        from sku_pescar_core import run
    except Exception as e:
        return jsonify({"error": "服务未就绪: " + str(e)}), 500
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
        f.save(tmp.name)
        try:
            data = run(tmp.name)
            if data is None:
                return jsonify({"error": "解析 Excel 失败或缺少 SKU/产品图/变种图列"}), 400
            return _cors(jsonify(data))
        finally:
            try:
                os.unlink(tmp.name)
            except Exception:
                pass


@app.route("/api/export-excel-json", methods=["POST", "OPTIONS"])
def api_export_excel_json():
    if request.method == "OPTIONS":
        return _cors(app.make_default_options_response())
    f, err = _get_file()
    if err:
        return jsonify({"error": err}), 400
    try:
        from excel_export_core import run
    except Exception as e:
        return jsonify({"error": "服务未就绪: " + str(e)}), 500
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
        f.save(tmp.name)
        try:
            data = run(tmp.name)
            if data is None:
                return jsonify({"error": "解析 Excel 失败"}), 400
            return _cors(jsonify(data))
        finally:
            try:
                os.unlink(tmp.name)
            except Exception:
                pass


@app.route("/api/empalagem", methods=["POST", "OPTIONS"])
def api_empalagem():
    if request.method == "OPTIONS":
        return _cors(app.make_default_options_response())
    f, err = _get_file()
    if err:
        return jsonify({"error": err}), 400
    try:
        from excel_export_core import run
    except Exception as e:
        return jsonify({"error": "服务未就绪: " + str(e)}), 500
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
        f.save(tmp.name)
        try:
            data = run(tmp.name)
            if data is None:
                return jsonify({"error": "解析 Excel 失败"}), 400
            return _cors(jsonify(data))
        finally:
            try:
                os.unlink(tmp.name)
            except Exception:
                pass
