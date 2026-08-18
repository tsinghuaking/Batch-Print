"""
批量打印 Web 服务（本地使用）
访问 http://127.0.0.1:5000
"""
import json
import os
import sys
import threading
import webbrowser
import socket

from flask import Flask, request, jsonify, send_from_directory, Response
import print_core

if getattr(sys, "frozen", False):
    BASE = sys._MEIPASS
else:
    BASE = os.path.dirname(os.path.abspath(__file__))
TEMPLATES = os.path.join(BASE, "templates")

PORT = 5001

CONFIG_DIR = os.path.join(os.path.expanduser("~"), ".batch_print")
CONFIG_FILE = os.path.join(CONFIG_DIR, "config.json")


def load_config():
    try:
        with open(CONFIG_FILE, encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def save_config(cfg):
    os.makedirs(CONFIG_DIR, exist_ok=True)
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)


app = Flask(__name__, template_folder=TEMPLATES)
app.config["MAX_CONTENT_LENGTH"] = 200 * 1024 * 1024  # 单文件上限 200MB

ALLOWED = {".pdf", ".doc", ".docx", ".xls", ".xlsx"}


@app.route("/")
def index():
    return send_from_directory(TEMPLATES, "index.html")


@app.route("/printers")
def printers():
    names = print_core.list_printers()
    return jsonify(names)


@app.route("/config")
def config():
    """返回已保存的默认打印机名。"""
    return jsonify({"default_printer": load_config().get("default_printer", "")})


@app.route("/default_printer", methods=["POST"])
def set_default_printer():
    name = (request.get_json(force=True) or {}).get("printer", "")
    if not name:
        return jsonify({"error": "未提供打印机名"}), 400
    cfg = load_config()
    cfg["default_printer"] = name
    save_config(cfg)
    return jsonify({"ok": True, "default_printer": name})


@app.route("/upload", methods=["POST"])
def upload():
    files = request.files.getlist("files")
    if not files:
        return jsonify({"error": "没有收到文件"}), 400
    saved = []
    for f in files:
        ext = os.path.splitext(f.filename or "")[1].lower()
        if ext not in ALLOWED:
            saved.append({"name": f.filename, "ok": False,
                          "msg": f"不支持的格式 {ext}（仅支持 PDF/Word/Excel）"})
            continue
        try:
            info = print_core.save_upload(f)
            saved.append({"id": info["id"], "name": info["name"],
                          "ext": info["ext"], "ok": True})
        except Exception as e:
            saved.append({"name": f.filename, "ok": False, "msg": str(e)})
    return jsonify(saved)


@app.route("/print", methods=["POST"])
def do_print():
    data = request.get_json(force=True)
    printer = data.get("printer", "")
    items = data.get("items", [])
    if not printer:
        return jsonify({"error": "未选择打印机"}), 400
    if not items:
        return jsonify({"error": "没有要打印的文件"}), 400

    # 先校验打印机名是否真实存在，避免 SumatraPDF 对无效名弹错误框卡死
    known = set(print_core.list_printers())
    if printer not in known:
        return jsonify({"error": f"打印机“{printer}”不存在，请刷新后重选"}), 400

    results = []
    for it in items:
        fid = it.get("id", "")
        path = os.path.join(print_core.UPLOADS, fid)
        name = it.get("name", fid)
        if not os.path.exists(path):
            results.append({"name": name, "id": fid, "ok": False,
                             "msg": "文件不存在，请重新上传"})
            continue
        ext = os.path.splitext(path)[1].lower()
        try:
            pdf = print_core.to_pdf(path, ext)
            settings = print_core.build_settings(
                it.get("duplex", "simplex"),
                it.get("copies", 1),
                it.get("color", "color"),
                it.get("pages", ""),
            )
            ok, msg = print_core.print_pdf(pdf, printer, settings)
            results.append({"name": name, "id": fid, "ok": ok, "msg": msg,
                            "settings": settings})
        except Exception as e:
            results.append({"name": name, "ok": False, "msg": str(e)})
    ok_count = sum(1 for r in results if r["ok"])
    return jsonify({"results": results, "ok": ok_count, "total": len(results)})


if __name__ == "__main__":
    print_core.ensure_uploads()

    def port_in_use(p):
        s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        try:
            return s.connect_ex(("127.0.0.1", p)) == 0
        finally:
            s.close()

    if port_in_use(PORT):
        # 端口已被占用（多半是上次未关闭的实例），直接打开浏览器复用之
        print(f"端口 {PORT} 已被占用，可能已有实例在运行，直接打开浏览器。")
        webbrowser.open(f"http://127.0.0.1:{PORT}")
    else:
        # 等服务真正起来后再自动打开默认浏览器
        threading.Timer(1.2, lambda: webbrowser.open(f"http://127.0.0.1:{PORT}")).start()
        print(f"批量打印服务已启动: http://127.0.0.1:{PORT}")
        app.run(host="127.0.0.1", port=PORT, debug=False, threaded=True)
