"""
批量打印 Web 服务（本地使用，纯标准库实现，无第三方依赖）
访问 http://127.0.0.1:5001
打包：python build_exe.py  ->  dist/批量打印工具.exe
"""
import os
import sys
import json
import uuid
import base64
import socket
import threading
import webbrowser
import urllib.parse

from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
import print_core

if getattr(sys, "frozen", False):
    RES = sys._MEIPASS          # 只读资源（templates/）
    APP_DIR = os.path.dirname(sys.executable)
else:
    RES = os.path.dirname(os.path.abspath(__file__))
    APP_DIR = RES

TEMPLATES = os.path.join(RES, "templates")
PORT = 5001
ALLOWED = {".pdf", ".doc", ".docx", ".xls", ".xlsx"}

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


def safe_name(name):
    base = os.path.basename(name)
    base = "".join(c if (c.isalnum() or c in "._- ") else "_" for c in base)
    return base.strip() or "file"


class Handler(BaseHTTPRequestHandler):
    def log_message(self, *a):  # 静默访问日志
        pass

    def _send_json(self, obj, code=200):
        b = json.dumps(obj, ensure_ascii=False).encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(b)))
        self.end_headers()
        self.wfile.write(b)

    def _static(self):
        try:
            with open(os.path.join(TEMPLATES, "index.html"), "rb") as f:
                b = f.read()
        except Exception as e:
            self.send_response(500)
            self.end_headers()
            self.wfile.write(str(e).encode("utf-8"))
            return
        self.send_response(200)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Content-Length", str(len(b)))
        self.end_headers()
        self.wfile.write(b)

    def do_GET(self):
        p = urllib.parse.urlparse(self.path).path
        if p in ("/", "/index.html"):
            self._static()
        elif p == "/printers":
            self._send_json(print_core.list_printers())
        elif p == "/config":
            self._send_json({"default_printer": load_config().get("default_printer", "")})
        else:
            self._send_json({"error": "not found"}, 404)

    def do_POST(self):
        p = urllib.parse.urlparse(self.path).path
        try:
            length = int(self.headers.get("Content-Length", 0) or 0)
        except ValueError:
            length = 0
        body = self.rfile.read(length) if length else b""
        try:
            data = json.loads(body.decode("utf-8")) if body else {}
        except Exception:
            data = {}

        if p == "/default_printer":
            self._post_default(data)
        elif p == "/upload":
            self._post_upload(data)
        elif p == "/print":
            self._post_print(data)
        else:
            self._send_json({"error": "not found"}, 404)

    def _post_default(self, data):
        name = (data or {}).get("printer", "")
        if not name:
            self._send_json({"error": "未提供打印机名"}, 400)
            return
        cfg = load_config()
        cfg["default_printer"] = name
        save_config(cfg)
        self._send_json({"ok": True, "default_printer": name})

    def _post_upload(self, data):
        name = (data or {}).get("name", "file")
        ext = os.path.splitext(name)[1].lower()
        if ext not in ALLOWED:
            self._send_json({"ok": False, "name": name,
                             "msg": f"不支持的格式 {ext}（仅支持 PDF/Word/Excel）"})
            return
        try:
            raw = base64.b64decode((data or {}).get("data", ""))
        except Exception:
            self._send_json({"ok": False, "name": name, "msg": "文件解码失败"})
            return
        print_core.ensure_uploads()
        fid = uuid.uuid4().hex + "_" + safe_name(name)
        with open(os.path.join(print_core.UPLOADS, fid), "wb") as f:
            f.write(raw)
        self._send_json({"ok": True, "id": fid, "name": safe_name(name), "ext": ext})

    def _post_print(self, data):
        printer = (data or {}).get("printer", "")
        items = (data or {}).get("items", [])
        if not printer:
            self._send_json({"error": "未选择打印机"}, 400)
            return
        if not items:
            self._send_json({"error": "没有要打印的文件"}, 400)
            return
        known = set(print_core.list_printers())
        if printer not in known:
            self._send_json({"error": f"打印机“{printer}”不存在，请刷新后重选"}, 400)
            return

        results = []
        ok_count = 0
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
                if ok:
                    ok_count += 1
            except Exception as e:
                results.append({"name": name, "id": fid, "ok": False, "msg": str(e)})
        self._send_json({"results": results, "ok": ok_count, "total": len(results)})


def port_in_use(p):
    s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    try:
        return s.connect_ex(("127.0.0.1", p)) == 0
    finally:
        s.close()


if __name__ == "__main__":
    print_core.ensure_uploads()
    if port_in_use(PORT):
        print(f"端口 {PORT} 已被占用，可能已有实例在运行，直接打开浏览器。")
        webbrowser.open(f"http://127.0.0.1:{PORT}")
    else:
        server = ThreadingHTTPServer(("127.0.0.1", PORT), Handler)
        threading.Timer(1.2, lambda: webbrowser.open(f"http://127.0.0.1:{PORT}")).start()
        print(f"批量打印服务已启动: http://127.0.0.1:{PORT}")
        print("关闭本窗口（或按 Ctrl+C）即停止服务。")
        try:
            server.serve_forever()
        except KeyboardInterrupt:
            pass
