"""
批量打印核心模块（本机专用，不涉及任何网络上传）
- 枚举系统打印机
- PDF：用系统关联的 PDF 阅读器做「右击 -> 打印」（os.startfile print），最快最稳；
      指定页码时用 pypdf 抽取成临时 PDF 再打印，从而精确打第几页
- Word/Excel/PPT：用本机 Microsoft Office COM 直接 PrintOut，不经任何中间转换
- 双面/彩色/份数/页码 通过临时修改打印机默认 DEVMODE 或 Office PrintOut 参数实现（打印后还原）
- 不再依赖 SumatraPDF / LibreOffice
- Office 转 PDF（to_pdf）仅用于后台统计页数，不参与打印
"""
import os
import re
import json
import subprocess
import threading
import uuid
import sys
import tempfile

if os.name == "nt":
    # 打包成无黑窗 exe 后，调用外部程序（SumatraPDF/PowerShell/LibreOffice）
    # 若不加此标记，每次调用都会闪一个黑色控制台窗口
    NO_WINDOW = subprocess.CREATE_NO_WINDOW
else:
    NO_WINDOW = 0

if getattr(sys, "frozen", False):
    # 打包后：BASE 为只读资源目录（bin/、templates/ 已内嵌于此），
    # APP_DIR 为 exe 所在目录（可写、持久），上传文件与配置放这里。
    BASE = sys._MEIPASS
    APP_DIR = os.path.dirname(sys.executable)
    # 若用户把 bin/ 放到 exe 同目录（外置分发），优先用外置版（便于单独更新引擎）
    EXTERNAL_BIN = os.path.join(APP_DIR, "bin")
    BIN = EXTERNAL_BIN if os.path.exists(os.path.join(EXTERNAL_BIN, "SumatraPDF.exe")) else os.path.join(BASE, "bin")
else:
    BASE = os.path.dirname(os.path.abspath(__file__))
    APP_DIR = BASE
    BIN = os.path.join(APP_DIR, "bin")

SUMATRA = os.path.join(BIN, "SumatraPDF.exe")
UPLOADS = os.path.join(APP_DIR, "uploads")

# 常见 LibreOffice 安装位置（soffice.exe）
SOFFICE_CANDIDATES = [
    r"C:\Program Files\LibreOffice\program\soffice.exe",
    r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    r"D:\Program Files\LibreOffice\program\soffice.exe",
]

OFFICE_EXT = {".doc", ".docx", ".xls", ".xlsx", ".ppt", ".pptx"}

# LibreOffice 转换全局锁：soffice 同一用户配置目录不能并发跑多个实例，
# 否则第二个实例会等锁甚至失败。上传时多文件并行请求页数会同时触发转换，
# 必须串行化（ThreadingHTTPServer 每个请求一个线程）。
_OFFICE_LOCK = threading.Lock()


def ensure_uploads():
    os.makedirs(UPLOADS, exist_ok=True)


def find_soffice():
    for p in SOFFICE_CANDIDATES:
        if os.path.exists(p):
            return p
    # 兜底：在 PATH 里找
    import shutil
    return shutil.which("soffice") or shutil.which("soffice.exe")


def list_printers():
    """枚举 Windows 已安装打印机，返回名称列表。用 PowerShell 写入临时文件（UTF-8），避开管道编码问题。"""
    try:
        return _list_printers_file()
    except Exception as e:
        print("PowerShell 文件法失败，回退 API:", e)
        try:
            return _list_printers_api()
        except Exception as e2:
            print("API 也失败:", e2)
            return []


def _list_printers_api():
    import ctypes
    from ctypes import wintypes
    winspool = ctypes.WinDLL("winspool.drv", use_last_error=True)
    PRINTER_ENUM_LOCAL = 0x2
    PRINTER_ENUM_CONNECTIONS = 0x4
    flags = PRINTER_ENUM_LOCAL | PRINTER_ENUM_CONNECTIONS
    pcbNeeded = wintypes.DWORD(0)
    pcReturned = wintypes.DWORD(0)
    # 第一次调用拿所需缓冲区大小
    winspool.EnumPrintersW(flags, None, 2, None, 0,
                           ctypes.byref(pcbNeeded), ctypes.byref(pcReturned))
    buf = ctypes.create_string_buffer(pcbNeeded.value)
    if not winspool.EnumPrintersW(flags, None, 2, buf, pcbNeeded.value,
                                  ctypes.byref(pcbNeeded), ctypes.byref(pcReturned)):
        raise ctypes.WinError(ctypes.get_last_error())
    # PRINTER_INFO_2W：pPrinterName 在偏移 8 字节（第一个指针字段）
    names = []
    size = pcbNeeded.value
    base = ctypes.cast(buf, ctypes.c_void_p).value
    PTR = ctypes.sizeof(ctypes.c_void_p)
    for i in range(pcReturned.value):
        info = base + i * 136  # PRINTER_INFO_2 结构大小（x64）= 13 指针 + 8 DWORD = 136 字节
        pname = ctypes.cast(info + PTR, ctypes.POINTER(ctypes.c_void_p))[0]
        if pname:
            s = ctypes.cast(pname, ctypes.c_wchar_p).value
            if s:
                names.append(s)
    return names


def _list_printers_file():
    import tempfile
    tmp = os.path.join(tempfile.gettempdir(), f"printers_{os.getpid()}.txt")
    ps = (f"(Get-Printer | ForEach-Object {{ $_.Name }}) | "
          f"Out-File -Encoding utf8 '{tmp.replace(chr(92), '/')}'")
    subprocess.run(["powershell.exe", "-NoProfile", "-Command", ps],
                   capture_output=True, timeout=30, check=True,
                   creationflags=NO_WINDOW)
    with open(tmp, "r", encoding="utf-8-sig") as f:
        names = [l.strip() for l in f.read().splitlines() if l.strip()]
    try:
        os.remove(tmp)
    except Exception:
        pass
    return names


def safe_name(name):
    """只保留文件名中的安全字符，避免路径穿越。"""
    base = os.path.basename(name)
    base = re.sub(r"[^\w\u4e00-\u9fff.\- ]", "_", base)
    return base.strip() or "file"


def save_upload(file_storage):
    """保存上传文件，返回 {id, name, ext}。id 即存储文件名。"""
    ensure_uploads()
    orig = safe_name(file_storage.filename or "file")
    ext = os.path.splitext(orig)[1].lower()
    fid = f"{uuid.uuid4().hex}_{orig}"
    path = os.path.join(UPLOADS, fid)
    file_storage.save(path)
    return {"id": fid, "name": orig, "ext": ext, "path": path}


def _remove_if_exists(p):
    try:
        if os.path.exists(p):
            os.remove(p)
    except Exception:
        pass


def _word_to_pdf(src, pdf):
    """用本机 Microsoft Word 导出 PDF（原生、秒级、稳）。"""
    import win32com.client
    _remove_if_exists(pdf)
    app = win32com.client.Dispatch("Word.Application")
    app.Visible = False
    try:
        app.DisplayAlerts = False
        doc = app.Documents.Open(os.path.abspath(src))
        doc.ExportAsFixedFormat(pdf, 17)  # 17 = wdExportFormatPDF
        doc.Close(False)
    finally:
        app.Quit()
    if not os.path.exists(pdf):
        raise RuntimeError("Word 导出 PDF 未生成")
    return pdf


def _excel_to_pdf(src, pdf):
    """用本机 Microsoft Excel 导出 PDF。"""
    import win32com.client
    _remove_if_exists(pdf)
    app = win32com.client.Dispatch("Excel.Application")
    app.Visible = False
    app.DisplayAlerts = False
    try:
        wb = app.Workbooks.Open(os.path.abspath(src))
        wb.ExportAsFixedFormat(0, pdf)  # 0 = xlTypePDF
        wb.Close(False)
    finally:
        app.Quit()
    if not os.path.exists(pdf):
        raise RuntimeError("Excel 导出 PDF 未生成")
    return pdf


def _ppt_to_pdf(src, pdf):
    """用本机 Microsoft PowerPoint 导出 PDF。"""
    import win32com.client
    _remove_if_exists(pdf)
    app = win32com.client.Dispatch("PowerPoint.Application")
    app.Visible = False
    try:
        pres = app.Presentations.Open(os.path.abspath(src), WithWindow=False)
        pres.ExportAsFixedFormat(pdf, 32)  # 32 = ppFixedFormatTypePDF
        pres.Close()
    finally:
        app.Quit()
    if not os.path.exists(pdf):
        raise RuntimeError("PPT 导出 PDF 未生成")
    return pdf


def _office_to_pdf(src, pdf):
    ext = os.path.splitext(src)[1].lower()
    if ext in (".doc", ".docx"):
        return _word_to_pdf(src, pdf)
    if ext in (".xls", ".xlsx"):
        return _excel_to_pdf(src, pdf)
    if ext in (".ppt", ".pptx"):
        return _ppt_to_pdf(src, pdf)
    raise ValueError("不支持的 Office 格式: " + ext)


def _libreoffice_to_pdf(src, pdf):
    """LibreOffice 兜底：转换后改名为与源文件同名的 pdf。"""
    soffice = find_soffice()
    if not soffice:
        raise RuntimeError("未找到 LibreOffice（soffice.exe），无法兜底转换。")
    outdir = os.path.dirname(src)
    cmd = [soffice, "--headless", "--norestore", "--convert-to", "pdf",
           "--outdir", outdir, src]
    r = subprocess.run(cmd, capture_output=True, text=True,
                       errors="replace", timeout=180, creationflags=NO_WINDOW)
    if r.returncode != 0:
        raise RuntimeError(f"LibreOffice 转换失败: {r.stderr[-500:]}")
    base = os.path.splitext(os.path.basename(src))[0]
    out_pdf = os.path.join(outdir, base + ".pdf")
    if not os.path.exists(out_pdf):
        raise RuntimeError("LibreOffice 未输出 PDF: " + out_pdf)
    if out_pdf != pdf:
        _remove_if_exists(pdf)
        os.replace(out_pdf, pdf)
    return pdf


def to_pdf(src_path, ext):
    """Office 文件转 PDF；PDF 直接返回原路径。

    优先用本机 Microsoft Office 原生「导出 PDF」（快、稳、无 LibreOffice 锁冲突）；
    未装 Office 时回退 LibreOffice。失败抛异常（含各路错误信息）。
    """
    if ext == ".pdf":
        return src_path
    if ext not in OFFICE_EXT:
        raise ValueError(f"不支持的格式: {ext}")
    pdf = os.path.splitext(src_path)[0] + ".pdf"
    errors = []
    for exporter in (_office_to_pdf, _libreoffice_to_pdf):
        try:
            with _OFFICE_LOCK:
                return exporter(src_path, pdf)
        except Exception as e:
            errors.append(f"[{exporter.__name__}] {str(e)[:200]}")
    raise RuntimeError("Office 与 LibreOffice 均转换失败:\n" + "\n".join(errors))


def get_page_count(src_path, ext):
    """返回文件总页数。PDF 用 pypdf 快速直读；Office 转 PDF 再统计。"""
    if ext == ".pdf":
        try:
            from pypdf import PdfReader
            return len(PdfReader(src_path).pages)
        except Exception as e:
            raise RuntimeError(f"PDF 页数读取失败: {e}")
    if ext in OFFICE_EXT:
        pdf = to_pdf(src_path, ext)  # 转换副产物，缓存由 OS 文件系统承担
        try:
            from pypdf import PdfReader
            return len(PdfReader(pdf).pages)
        except Exception as e:
            raise RuntimeError(f"页数读取失败: {e}")
    raise ValueError(f"不支持的格式: {ext}")


# ============================================================================
# 本地原生打印（不再经过 SumatraPDF / LibreOffice）
# - PDF：用系统关联的 PDF 阅读器做「右击 -> 打印」（os.startfile 的 print 动作），
#        最快最稳，就像在资源管理器里右键打印一样。
# - Word/Excel/PPT：用本机 Office COM 直接 PrintOut，不经过任何中间转换。
# 双面 / 彩色 / 份数 通过临时修改打印机默认 DEVMODE 实现（打印后还原）。
# ============================================================================

def _int_copies(v, default=1):
    try:
        c = int(v)
        return c if c > 0 else default
    except Exception:
        return default


def _parse_pages(pages):
    """把 '1-3' / '1,3,5-7' / '' 解析成 [(from,to), ...] 列表；空或非法返回 None。"""
    pages = (pages or "").strip()
    if not pages:
        return None
    parts = []
    for seg in pages.replace(" ", "").split(","):
        if not seg:
            continue
        if "-" in seg:
            a, b = seg.split("-", 1)
            try:
                parts.append((int(a), int(b)))
            except Exception:
                pass
        else:
            try:
                n = int(seg)
                parts.append((n, n))
            except Exception:
                pass
    return parts or None


def _apply_printer_mode(printer, duplex, color, copies):
    """临时把打印机的默认双面/彩色/份数改成用户选择，返回原函数便于还原。
    失败返回 None（调用方跳过还原）。win32print 法。
    """
    try:
        import win32print
    except Exception:
        return None
    h = None
    try:
        h = win32print.OpenPrinter(printer)
        info = win32print.GetPrinter(h, 2)
        dm = info["pDevMode"]
        prev = {"Duplex": dm.Duplex, "Color": dm.Color,
                "Copies": dm.Copies, "Fields": dm.Fields}
        # DM_DUPLEX: 1=单面 2=长边 3=短边；DM_COLOR: 1=黑白 2=彩色；DM_COPIES
        dm.Duplex = {"simplex": 1, "duplex": 2, "duplexshort": 3}.get(duplex, 1)
        dm.Color = 1 if color == "monochrome" else 2
        c = _int_copies(copies)
        if c > 1:
            dm.Copies = c
        # 确保 Fields 包含我们修改的位，驱动才可能真正采纳
        try:
            dm.Fields |= (0x1000 | 0x800 | 0x100)
        except Exception:
            pass
        info["pDevMode"] = dm
        win32print.SetPrinter(h, 2, info, 0)
        return prev
    except Exception:
        return None
    finally:
        if h:
            win32print.ClosePrinter(h)


def _restore_printer_mode(printer, prev):
    if not prev:
        return
    try:
        import win32print
    except Exception:
        return
    h = None
    try:
        h = win32print.OpenPrinter(printer)
        info = win32print.GetPrinter(h, 2)
        dm = info["pDevMode"]
        dm.Duplex = prev["Duplex"]
        dm.Color = prev["Color"]
        dm.Copies = prev["Copies"]
        dm.Fields = prev["Fields"]
        info["pDevMode"] = dm
        win32print.SetPrinter(h, 2, info, 0)
    except Exception:
        pass
    finally:
        if h:
            win32print.ClosePrinter(h)


def print_pdf_native(path, printer, settings):
    """打印 PDF：用系统关联的 PDF 阅读器做「右击 -> 打印」（os.startfile 的 print 动作）。

    最快最稳，就像在资源管理器里右键打印一样；打印完阅读器是否保持打开由系统/用户决定，
    本程序不强行关闭它（避免误杀用户已打开的其它 PDF 文档）。
    页码范围：用 pypdf 把指定页抽成临时 PDF 再交给阅读器，从而精确打第几页。
    """
    pages = (settings.get("pages") or "").strip()
    target = os.path.abspath(path)
    tmp = None
    ranges = _parse_pages(pages)
    if ranges:
        try:
            from pypdf import PdfReader, PdfWriter
            reader = PdfReader(os.path.abspath(path))
            total = len(reader.pages)
            writer = PdfWriter()
            wanted = set()
            for a, b in ranges:
                for p in range(a, b + 1):
                    if 1 <= p <= total:
                        wanted.add(p - 1)  # pypdf 是 0 基
            if not wanted:
                return False, f"指定的页码超出范围（文件共 {total} 页）"
            for idx in sorted(wanted):
                writer.add_page(reader.pages[idx])
            tmp = os.path.join(tempfile.gettempdir(), f"bp_{uuid.uuid4().hex}.pdf")
            with open(tmp, "wb") as f:
                writer.write(f)
            target = tmp
        except Exception as e:
            return False, f"提取指定页失败：{e}"

    try:
        import win32print
        old_default = win32print.GetDefaultPrinter()
    except Exception:
        old_default = None
    prev = _apply_printer_mode(
        printer, settings.get("duplex", "simplex"),
        settings.get("color", "color"), settings.get("copies", 1))
    try:
        if old_default and printer != old_default:
            try:
                win32print.SetDefaultPrinter(printer)
            except Exception:
                pass

        # 系统关联程序「右击打印」（os.startfile 的 print 动作）：最快最稳，
        # 就像在资源管理器里右键打印一样。打印完阅读器是否保持打开由系统/用户决定，
        # 本程序不强行关闭它（避免误杀用户已打开的其它 PDF 文档）。
        os.startfile(target, "print")
        msg = "已发送打印任务（本地程序打印）"

        if tmp:
            msg += f"  [已抽取第 {pages} 页]"
        # 阅读器异步读取临时文件，60 秒后再尝试清理
        if tmp:
            try:
                threading.Timer(60.0, lambda: _remove_if_exists(tmp)).start()
            except Exception:
                pass
        return True, msg
    except Exception as e:
        _remove_if_exists(tmp)
        return False, f"本地打印失败：{e}"
    finally:
        if old_default:
            try:
                win32print.SetDefaultPrinter(old_default)
            except Exception:
                pass
        _restore_printer_mode(printer, prev)


def _word_print(src, printer, settings):
    import win32com.client
    import pythoncom
    prev = _apply_printer_mode(
        printer, settings.get("duplex", "simplex"),
        settings.get("color", "color"), settings.get("copies", 1))
    app = None
    doc = None
    try:
        pythoncom.CoInitialize()  # Flask 工作线程必须先初始化 COM 公寓，否则 Dispatch 报 -2147221008
    except Exception:
        pass
    try:
        app = win32com.client.Dispatch("Word.Application")
        app.Visible = False
        app.DisplayAlerts = False
        # Word 的打印机切换走 Application.ActivePrinter 属性，
        # 直接给 Document.PrintOut 传 ActivePrinter 关键字在新版 win32com 里会报
        # `unexpected keyword argument 'ActivePrinter'`。
        app.ActivePrinter = printer
        doc = app.Documents.Open(os.path.abspath(src))
        copies = _int_copies(settings.get("copies", 1))
        base = {"Copies": copies, "Background": False}
        ranges = _parse_pages(settings.get("pages", ""))
        if ranges:
            for a, b in ranges:
                kw = dict(base)
                if a == b:
                    kw["Range"] = 4          # wdPrintRangeOfPages
                    kw["Pages"] = str(a)
                else:
                    kw["Range"] = 3          # wdPrintFromTo
                    kw["From"] = str(a)
                    kw["To"] = str(b)
                doc.PrintOut(**kw)
        else:
            doc.PrintOut(**base)
        return True, "已发送打印任务（Word 直打）"
    except Exception as e:
        return False, f"Word 打印失败：{e}"
    finally:
        try:
            if doc is not None:
                doc.Close(False)
        except Exception:
            pass
        try:
            if app is not None:
                app.Quit()
        except Exception:
            pass
        _restore_printer_mode(printer, prev)
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass


def _excel_print(src, printer, settings):
    import win32com.client
    import pythoncom
    prev = _apply_printer_mode(
        printer, settings.get("duplex", "simplex"),
        settings.get("color", "color"), settings.get("copies", 1))
    app = None
    wb = None
    try:
        pythoncom.CoInitialize()  # Flask 工作线程必须先初始化 COM 公寓，否则 Dispatch 报 -2147221008
    except Exception:
        pass
    try:
        app = win32com.client.Dispatch("Excel.Application")
        app.Visible = False
        app.DisplayAlerts = False
        wb = app.Workbooks.Open(os.path.abspath(src))
        copies = _int_copies(settings.get("copies", 1))
        ranges = _parse_pages(settings.get("pages", ""))
        if ranges:
            for a, b in ranges:
                wb.PrintOut(From=a, To=b, Copies=copies,
                            ActivePrinter=printer, IgnorePrintAreas=True)
        else:
            wb.PrintOut(Copies=copies, ActivePrinter=printer,
                        IgnorePrintAreas=True)
        return True, "已发送打印任务（Excel 直打）"
    except Exception as e:
        return False, f"Excel 打印失败：{e}"
    finally:
        try:
            if wb is not None:
                wb.Close(False)
        except Exception:
            pass
        try:
            if app is not None:
                app.Quit()
        except Exception:
            pass
        _restore_printer_mode(printer, prev)
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass


def _ppt_print(src, printer, settings):
    import win32com.client
    import pythoncom
    prev = _apply_printer_mode(
        printer, settings.get("duplex", "simplex"),
        settings.get("color", "color"), settings.get("copies", 1))
    app = None
    pres = None
    try:
        pythoncom.CoInitialize()  # Flask 工作线程必须先初始化 COM 公寓，否则 Dispatch 报 -2147221008
    except Exception:
        pass
    try:
        app = win32com.client.Dispatch("PowerPoint.Application")
        app.Visible = False
        pres = app.Presentations.Open(os.path.abspath(src), WithWindow=False)
        copies = _int_copies(settings.get("copies", 1))
        pages = (settings.get("pages", "") or "").strip()
        # PrintOut(Range, From, To, PrintToFile, Copies, Collate, ActivePrinter)
        if pages:
            pres.PrintOut(Range=pages, From=0, To=0,
                          Copies=copies, ActivePrinter=printer)
        else:
            pres.PrintOut(Copies=copies, ActivePrinter=printer)
        return True, "已发送打印任务（PPT 直打）"
    except Exception as e:
        return False, f"PPT 打印失败：{e}"
    finally:
        try:
            if pres is not None:
                pres.Close()
        except Exception:
            pass
        try:
            if app is not None:
                app.Quit()
        except Exception:
            pass
        _restore_printer_mode(printer, prev)
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass


def print_office_native(src, ext, printer, settings):
    """Word/Excel/PPT 用本机 Office COM 直接打印（不走 SumatraPDF / 不转 PDF）。"""
    if ext in (".doc", ".docx"):
        return _word_print(src, printer, settings)
    if ext in (".xls", ".xlsx"):
        return _excel_print(src, printer, settings)
    if ext in (".ppt", ".pptx"):
        return _ppt_print(src, printer, settings)
    return False, "不支持的格式：" + ext


def print_file(path, ext, printer, settings):
    """统一打印入口（不经过 SumatraPDF）：
    - PDF：系统关联程序「右击打印」（最快最稳）
    - Word/Excel/PPT：本机 Office COM 直接打印
    settings: {duplex, copies, color, pages}
    """
    if ext == ".pdf":
        return print_pdf_native(path, printer, settings)
    if ext in OFFICE_EXT:
        return print_office_native(path, ext, printer, settings)
    return False, "不支持的格式：" + ext


def build_settings(duplex, copies, color, pages):
    """
    构造 SumatraPDF -print-settings 参数串。
    duplex: simplex / duplex(长边) / duplexshort(短边)
    color : color / monochrome
    copies: 整数
    pages : 形如 '1-3' / '1,3,5-7' / ''（空=全部）
    """
    parts = []
    if duplex == "duplex":
        parts.append("duplex")
    elif duplex == "duplexshort":
        parts.append("duplexshort")
    else:
        parts.append("simplex")
    parts.append("monochrome" if color == "monochrome" else "color")
    try:
        c = int(copies)
        if c < 1:
            c = 1
    except Exception:
        c = 1
    if c > 1:
        parts.append(f"copies={c}")
    p = (pages or "").strip()
    if p:
        # 仅允许数字、逗号、连字符、空格
        if re.fullmatch(r"[\d,\-\s]+", p):
            parts.append(f"pages={p.replace(' ', '')}")
    return ",".join(parts)


def print_pdf(pdf_path, printer, settings, timeout=180):
    """
    调用 SumatraPDF 静默打印，返回 (ok, msg)。

    timeout: 子进程等待秒数。默认 180 秒。
    部分 PCL/GDI 打印机在双面/多页模式下驱动层处理较慢，60 秒会误超时。
    """
    cmd = [SUMATRA, "-silent", "-print-to", printer,
           "-print-settings", settings, pdf_path]
    try:
        r = subprocess.run(cmd, capture_output=True, text=True,
                           errors="replace", timeout=timeout,
                           creationflags=NO_WINDOW)
    except subprocess.TimeoutExpired:
        return False, (f"打印超时（{timeout}s）；某些 PCL 打印机处理双面/多页较慢，"
                       f"可点「重新打印失败的文件」再试，或把任务拆小")
    except Exception as e:
        return False, f"调用失败: {e}"
    # SumatraPDF 打印成功通常返回 0；虚拟 PDF 打印机可能弹窗阻塞，这里只看退出码
    if r.returncode == 0:
        return True, "已发送打印任务"
    # 输出详细原因给前端日志（驱动错误/SumatraPDF 报错）
    err = (r.stderr or r.stdout or "").strip().splitlines()[-3:]
    return False, (f"退出码 {r.returncode}\n  cmd: {' '.join(cmd)}\n  "
                   + "\n  ".join(err) or "无输出")


if __name__ == "__main__":
    print("打印机列表:")
    for p in list_printers():
        print(" -", p)
    print("SumatraPDF:", SUMATRA, os.path.exists(SUMATRA))
    print("soffice:", find_soffice())
