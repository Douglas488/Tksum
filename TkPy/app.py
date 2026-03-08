# -*- coding: utf-8 -*-
"""
Tk 总结表生成 - Flask Web API
部署到 Render 后，前端 index9 通过此接口上传 Excel 并下载生成的总结表。
需配置环境变量 TOTP_SECRET（谷歌验证器 Base32 密钥）以启用验证码校验。
"""
import os
import sys
import tempfile
from io import BytesIO
from pathlib import Path

from flask import Flask, request, send_file, jsonify

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024  # 50MB

TOTP_SECRET = os.environ.get("TOTP_SECRET", "").strip()


def _cors_headers(resp):
    resp.headers["Access-Control-Allow-Origin"] = "*"
    resp.headers["Access-Control-Allow-Methods"] = "GET, POST, OPTIONS"
    resp.headers["Access-Control-Allow-Headers"] = "Content-Type"
    return resp


@app.after_request
def after_request(resp):
    return _cors_headers(resp)


@app.route("/api/verify-totp", methods=["POST", "OPTIONS"])
def api_verify_totp():
    """验证谷歌验证器 6 位动态码，防止滥用。需在 Render 配置 TOTP_SECRET。"""
    if request.method == "OPTIONS":
        return _cors_headers(app.make_default_options_response())
    if not TOTP_SECRET:
        return jsonify({"ok": False, "error": "服务未配置验证"}), 503
    try:
        data = request.get_json(force=True, silent=True) or {}
        raw = (data.get("pin") or "").strip()
        pin = "".join(c for c in raw if c.isdigit())[:6]
        if len(pin) != 6:
            return jsonify({"ok": False, "error": "请输入 6 位数字验证码"}), 400
        import pyotp
        totp = pyotp.TOTP(TOTP_SECRET)
        if totp.verify(pin, valid_window=2):
            return jsonify({"ok": True})
        return jsonify({"ok": False, "error": "验证码错误或已过期，请使用验证器中的当前 6 位码重试"})
    except Exception:
        return jsonify({"ok": False}), 500


def _get_generate_report():
    """延迟导入，便于在首页显示导入错误原因。"""
    from generate_summary import generate_report
    return generate_report


@app.route("/api/generate", methods=["POST", "OPTIONS"])
def api_generate():
    if request.method == "OPTIONS":
        return _cors_headers(app.make_default_options_response())

    if "file" not in request.files:
        return jsonify({"error": "请上传 Excel 文件（字段名: file）"}), 400

    f = request.files["file"]
    if not f.filename or not (f.filename.endswith(".xlsx") or f.filename.endswith(".xls") or f.filename.endswith(".xlsm")):
        return jsonify({"error": "仅支持 .xlsx / .xls / .xlsm 文件"}), 400

    try:
        generate_report = _get_generate_report()
    except Exception as e:
        return jsonify({"error": "服务未就绪，请检查依赖与日志: " + str(e)}), 500

    try:
        with tempfile.TemporaryDirectory() as tmp:
            tmp = Path(tmp)
            source_path = tmp / "source.xlsx"
            output_path = tmp / "Tk月总结表.xlsx"
            f.save(str(source_path))
            generate_report(source_path, None, output_path)
            buf = BytesIO(output_path.read_bytes())
            buf.seek(0)
            return _cors_headers(
                send_file(
                    buf,
                    as_attachment=True,
                    download_name="Tk月总结表.xlsx",
                    mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            )
    except FileNotFoundError as e:
        return jsonify({"error": str(e)}), 400
    except Exception as e:
        return jsonify({"error": f"生成失败: {str(e)}"}), 500


@app.route("/")
def index():
    info = {
        "service": "Tk总结表生成 API",
        "usage": "POST /api/generate with multipart form field 'file' (Excel)",
        "totp_verify": "POST /api/verify-totp with JSON body {\"pin\": \"6位动态码\"}，需先验证通过后再调用 /api/generate",
    }
    try:
        _get_generate_report()
        info["generate_summary"] = "ok"
    except Exception as e:
        info["generate_summary"] = "failed"
        info["error"] = str(e)
        info["python"] = sys.version.split()[0]
    return jsonify(info)
