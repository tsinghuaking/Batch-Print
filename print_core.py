"""
批量打印核心模块（本机专用，不涉及任何网络上传）
- 枚举系统打印机
- PDF：用系统关联的 PDF 阅读器做「右击 -> 打印」（os.startfile print），最快最稳；
      指定页码时用 pypdf 抽取成临时 PDF 再打印，从而精确打第几页
- Word/Excel/PPT：用本机 Microsoft Office COM 直接 PrintOut，不经任何中间转换
- 双面/彩色/份数/页码 通过临时修改打印机默认 DEVMODE 或 Office PrintOut 参数实现（打印后还原）
- 不依赖 SumatraPDF / LibreOffice（V1.7 已彻底移除相关代码）
- Office 转 PDF（to_pdf）仅用于后台统计页数，不参与打印
"""
import os
import re
import subprocess
import threading
import uuid
import sys
import tempfile

if os.name == "nt":
    # 打包成无黑窗 exe 后，调用外部程序（PowerShell）若不加此标记，
    # 每次调用都会闪一个黑色控制台窗口
    NO_WINDOW = subprocess.CREATE_NO_WINDOW
else:
    NO_WINDOW = 0

if getattr(sys, "frozen", False):
    # 打包后：BASE 为只读资源目录（templates/ 已内嵌于此），
    # APP_DIR 为 exe 所在目录（可写、持久），上传文件与配置放这里。
    BASE = sys._MEIPASS
    APP_DIR = os.path.dirname(sys.executable)
else:
    BASE = os.path.dirname(os.path.abspath(__file__))
    APP_DIR = BASE

UPLOADS = os.path.join(APP_DIR, "uploads")

OFFICE_EXT = {".doc", ".docx", ".xls", ".xlsx", ".ppt", ".pptx"}

# Office 转换全局锁：Word/Excel/PowerPoint 的 COM 自动化不能并发跑多个实例，
# 否则会互相抢焦点甚至失败。上传时多文件并行请求页数会同时触发转换，
# 必须串行化（ThreadingHTTPServer 每个请求一个线程）。
_OFFICE_LOCK = threading.Lock()


def ensure_uploads():
    os.makedirs(UPLOADS, exist_ok=True)


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


def to_pdf(src_path, ext):
    """Office 文件转 PDF（仅用于统计页数）；PDF 直接返回原路径。

    用本机 Microsoft Office 原生「导出 PDF」，快且稳；未装 Office 时抛异常，
    调用方（页数统计）会显示「读取失败」，不影响打印。
    """
    if ext == ".pdf":
        return src_path
    if ext not in OFFICE_EXT:
        raise ValueError(f"不支持的格式: {ext}")
    pdf = os.path.splitext(src_path)[0] + ".pdf"
    with _OFFICE_LOCK:
        return _office_to_pdf(src_path, pdf)


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
# 本地原生打印（不经任何中间转换）
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
        # PowerPoint 的 PrintOut 用 From/To 指定页码区间（Range 参数是 PrintRange 对象，
        # 传字符串会报错），故按每段区间逐个提交，与 Word/Excel 保持一致。
        ranges = _parse_pages(settings.get("pages", ""))
        if ranges:
            for a, b in ranges:
                pres.PrintOut(From=a, To=b, Copies=copies,
                              ActivePrinter=printer)
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
    """Word/Excel/PPT 用本机 Office COM 直接打印（不经任何中间转换）。"""
    if ext in (".doc", ".docx"):
        return _word_print(src, printer, settings)
    if ext in (".xls", ".xlsx"):
        return _excel_print(src, printer, settings)
    if ext in (".ppt", ".pptx"):
        return _ppt_print(src, printer, settings)
    return False, "不支持的格式：" + ext


def print_file(path, ext, printer, settings):
    """统一打印入口（不经任何中间转换）：
    - PDF：系统关联程序「右击打印」（最快最稳）
    - Word/Excel/PPT：本机 Office COM 直接打印
    settings: {duplex, copies, color, pages}
    """
    if ext == ".pdf":
        return print_pdf_native(path, printer, settings)
    if ext in OFFICE_EXT:
        return print_office_native(path, ext, printer, settings)
    return False, "不支持的格式：" + ext


if __name__ == "__main__":
    print("打印机列表:")
    for p in list_printers():
        print(" -", p)
