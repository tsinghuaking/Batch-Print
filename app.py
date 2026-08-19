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
import time
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
ASSETS = os.path.join(RES, "assets")
PORT = 5001
ALLOWED = {".pdf", ".doc", ".docx", ".xls", ".xlsx"}

if getattr(sys, "frozen", False):
    # 无黑窗打包后，日志写入 exe 旁边的「运行日志.txt」（每次启动重新写），
    # 出问题把该文件发给维护者即可排查
    try:
        _log = open(os.path.join(APP_DIR, "运行日志.txt"), "w",
                    encoding="utf-8", buffering=1)
        sys.stdout = sys.stderr = _log
    except Exception:
        pass

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


# 打印机列表缓存：Get-Printer 枚举要跑 PowerShell（约 1.4 秒），
# 之前 /print_one 每个文件都重新枚举一遍，批量打印时白白多等好几秒。
_PRINTER_CACHE = {"t": 0.0, "set": set()}
_PRINTER_CACHE_TTL = 120  # 秒


def printers_cached():
    now = time.time()
    if now - _PRINTER_CACHE["t"] > _PRINTER_CACHE_TTL or not _PRINTER_CACHE["set"]:
        _PRINTER_CACHE["set"] = set(print_core.list_printers())
        _PRINTER_CACHE["t"] = now
    return _PRINTER_CACHE["set"]


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

    def _favicon(self):
        # 图标：exe 内置 assets/Batch Print.ico；浏览器标签页显示
        try:
            with open(os.path.join(ASSETS, "Batch Print.ico"), "rb") as f:
                b = f.read()
        except Exception:
            self.send_response(404)
            self.end_headers()
            return
        self.send_response(200)
        self.send_header("Content-Type", "image/x-icon")
        self.send_header("Content-Length", str(len(b)))
        self.end_headers()
        self.wfile.write(b)

    def _logo(self):
        # 页面顶部 LOGO：assets/logo.png（96px 小图）
        try:
            with open(os.path.join(ASSETS, "logo.png"), "rb") as f:
                b = f.read()
        except Exception:
            self.send_response(404)
            self.end_headers()
            return
        self.send_response(200)
        self.send_header("Content-Type", "image/png")
        self.send_header("Content-Length", str(len(b)))
        self.end_headers()
        self.wfile.write(b)

    def do_GET(self):
        p = urllib.parse.urlparse(self.path).path
        if p in ("/", "/index.html"):
            self._static()
        elif p == "/favicon.ico":
            self._favicon()
        elif p == "/logo.png":
            self._logo()
        elif p == "/printers":
            self._send_json(print_core.list_printers())
        elif p == "/config":
            self._send_json({"default_printer": load_config().get("default_printer", "")})
        elif p == "/pagecount":
            self._get_pagecount(urllib.parse.urlparse(self.path).query)
        else:
            self._send_json({"error": "not found"}, 404)

    def _get_pagecount(self, query):
        qs = urllib.parse.parse_qs(query or "")
        fid = (qs.get("id") or [""])[0]
        if not fid:
            self._send_json({"pages": None, "msg": "缺少 id"}, 400)
            return
        # 安全校验：仅匹配 UPLOADS 目录下的现有文件
        if "/" in fid or "\\" in fid or ".." in fid:
            self._send_json({"pages": None, "msg": "非法 id"}, 400)
            return
        path = os.path.join(print_core.UPLOADS, fid)
        if not os.path.exists(path):
            self._send_json({"pages": None, "msg": "文件不存在"}, 404)
            return
        ext = os.path.splitext(path)[1].lower()
        try:
            n = print_core.get_page_count(path, ext)
            self._send_json({"pages": n})
        except Exception as e:
            self._send_json({"pages": None, "msg": str(e)})

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
        elif p == "/prepare":
            self._post_prepare(data)
        elif p == "/print_one":
            self._post_print_one(data)
        elif p == "/print":
            self._post_print(data)
        elif p == "/shutdown":
            # 前端「退出程序」按钮调用：无黑窗模式下这是唯一的退出方式
            self._send_json({"ok": True})
            threading.Timer(0.5, lambda: os._exit(0)).start()
        else:
            self._send_json({"error": "not found"}, 404)

    def _post_prepare(self, data):
        """预转换端点：Office 现改为本机直打，无需提前转 PDF，故直接返回成功。
        （保留该接口仅为了兼容前端调用；真正需要转 PDF 的场景是统计页数，已在上传时完成。）"""
        self._send_json({"ok": True})

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

    def _post_print_one(self, data):
        """单文件打印端点：实时逐个调用，前端每份完成即更新状态。"""
        printer = (data or {}).get("printer", "")
        fid = (data or {}).get("id", "")
        name = (data or {}).get("name", fid)
        if not printer:
            return self._send_json({"ok": False, "id": fid, "msg": "未选打印机"}, 400)
        if not fid:
            return self._send_json({"ok": False, "id": fid, "msg": "缺 id"}, 400)
        known = printers_cached()
        if printer not in known:
            return self._send_json(
                {"ok": False, "id": fid, "msg": f"打印机「{printer}」不存在，请刷新后重选"}, 400)
        path = os.path.join(print_core.UPLOADS, fid)
        if not os.path.exists(path):
            return self._send_json(
                {"ok": False, "id": fid, "name": name, "msg": "文件不存在，请重新上传"}, 404)
        ext = os.path.splitext(path)[1].lower()
        try:
            settings = {
                "duplex": (data or {}).get("duplex", "simplex"),
                "copies": (data or {}).get("copies", 1),
                "color": (data or {}).get("color", "color"),
                "pages": (data or {}).get("pages", ""),
            }
            # PDF 走系统关联程序「右击打印」；Office 走本机 Office 直打；均不经 SumatraPDF
            ok, msg = print_core.print_file(path, ext, printer, settings)
            desc = self._settings_desc(settings)
            return self._send_json(
                {"ok": ok, "id": fid, "name": name, "msg": msg,
                 "settings": desc})
        except Exception as e:
            return self._send_json(
                {"ok": False, "id": fid, "name": name, "msg": str(e)})

    @staticmethod
    def _settings_desc(s):
        """把设置字典拼成可读串，用于前端日志展示。"""
        d = {"simplex": "单面", "duplex": "双面长边",
             "duplexshort": "双面短边"}.get(s.get("duplex"), "单面")
        col = "彩色" if s.get("color") == "color" else "黑白"
        c = s.get("copies", 1)
        try:
            c = int(c)
        except Exception:
            c = 1
        parts = [d, col, f"{c}份"]
        pages = (s.get("pages") or "").strip()
        if pages:
            parts.append("页码:" + pages)
        return "·".join(parts)

    def _post_print(self, data):
        printer = (data or {}).get("printer", "")
        items = (data or {}).get("items", [])
        if not printer:
            self._send_json({"error": "未选择打印机"}, 400)
            return
        if not items:
            self._send_json({"error": "没有要打印的文件"}, 400)
            return
        known = printers_cached()
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
                settings = {
                    "duplex": it.get("duplex", "simplex"),
                    "copies": it.get("copies", 1),
                    "color": it.get("color", "color"),
                    "pages": it.get("pages", ""),
                }
                ok, msg = print_core.print_file(path, ext, printer, settings)
                desc = self._settings_desc(settings)
                results.append({"name": name, "id": fid, "ok": ok,
                                "msg": msg, "settings": desc})
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


def open_browser(url):
    """用系统默认浏览器打开网页（普通标签页）。"""
    webbrowser.open(url)


if __name__ == "__main__":
    print_core.ensure_uploads()
    if port_in_use(PORT):
        print(f"端口 {PORT} 已被占用，可能已有实例在运行，直接打开浏览器。")
        open_browser(f"http://127.0.0.1:{PORT}")
    else:
        server = ThreadingHTTPServer(("127.0.0.1", PORT), Handler)
        threading.Timer(1.2, lambda: open_browser(f"http://127.0.0.1:{PORT}")).start()
        print(f"批量打印服务已启动: http://127.0.0.1:{PORT}")
        if getattr(sys, "frozen", False):
            print("点击网页右上角「退出程序」按钮即可停止服务。")
        else:
            print("关闭本窗口（或按 Ctrl+C）即停止服务。")
        try:
            server.serve_forever()
        except KeyboardInterrupt:
            pass
