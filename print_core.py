"""
批量打印核心模块
- 枚举系统打印机
- Word/Excel 经 LibreOffice 转 PDF
- 用 SumatraPDF 命令行带参数打印（双面/份数/彩色/页码）
本机专用，不涉及任何网络上传。
"""
import os
import re
import json
import subprocess
import uuid
import sys

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
                   capture_output=True, timeout=30, check=True)
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


def to_pdf(src_path, ext):
    """Office 文件转 PDF；PDF 直接返回原路径。失败抛异常。"""
    if ext == ".pdf":
        return src_path
    if ext not in OFFICE_EXT:
        raise ValueError(f"不支持的格式: {ext}")
    soffice = find_soffice()
    if not soffice:
        raise RuntimeError("未找到 LibreOffice（soffice.exe），无法转换 Word/Excel。")
    outdir = os.path.dirname(src_path)
    # LibreOffice 转换是阻塞的；加 --norestore 避免弹恢复窗
    cmd = [soffice, "--headless", "--norestore", "--convert-to", "pdf",
           "--outdir", outdir, src_path]
    r = subprocess.run(cmd, capture_output=True, text=True, timeout=180)
    if r.returncode != 0:
        raise RuntimeError(f"LibreOffice 转换失败: {r.stderr[-500:]}")
    # 输出文件名 = 原 basename + .pdf
    base = os.path.splitext(os.path.basename(src_path))[0]
    pdf = os.path.join(outdir, base + ".pdf")
    if not os.path.exists(pdf):
        raise RuntimeError("转换完成但未找到输出 PDF: " + pdf)
    return pdf


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


def print_pdf(pdf_path, printer, settings):
    """调用 SumatraPDF 静默打印，返回 (ok, msg)。"""
    cmd = [SUMATRA, "-silent", "-print-to", printer,
           "-print-settings", settings, pdf_path]
    try:
        r = subprocess.run(cmd, capture_output=True, text=True, timeout=60)
    except Exception as e:
        return False, f"调用失败或超时: {e}"
    # SumatraPDF 打印成功通常返回 0；虚拟 PDF 打印机可能弹窗阻塞，这里只看退出码
    if r.returncode == 0:
        return True, "已发送打印任务"
    return False, f"退出码 {r.returncode}；{r.stderr[-300:] or r.stdout[-300:]}"


if __name__ == "__main__":
    print("打印机列表:")
    for p in list_printers():
        print(" -", p)
    print("SumatraPDF:", SUMATRA, os.path.exists(SUMATRA))
    print("soffice:", find_soffice())
