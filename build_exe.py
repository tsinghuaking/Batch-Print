"""
把批量打印工具打包成单个 exe（Windows）。

用法：
    pip install pyinstaller pywin32 pypdf
    python build_exe.py

产物：dist/BatchPrint.exe
- 运行时仅用 Python 标准库（不含 Flask）；pypdf 只用于统计页数。
- templates/ 与 assets/ 全部内嵌进单个 exe，双击即用、自包含。
- 打印改用本机 Office 直打 + 系统 PDF 阅读器（右键打印），不再依赖 SumatraPDF / LibreOffice，故不打包任何外部引擎。

说明：
- 无控制台黑窗口（--windowed）。运行日志写入 exe 同目录「运行日志.txt」，
  页面右上角「退出程序」按钮可停止后台服务。
- pypdf 用 --collect-all 整体打包，否则统计页数时会出现「模块缺失」。
"""
import os
import PyInstaller.__main__

HERE = os.path.dirname(os.path.abspath(__file__))


def main():
    PyInstaller.__main__.run([
        os.path.join(HERE, "app.py"),
        "--name", "BatchPrint",
        "--onefile",
        "--noconfirm",
        "--windowed",
        "--clean",
        "--icon", os.path.join(HERE, "assets", "Batch Print.ico"),
        "--add-data", os.path.join(HERE, "templates") + os.pathsep + "templates",
        "--add-data", os.path.join(HERE, "assets") + os.pathsep + "assets",
        "--collect-all", "pypdf",
        # PIL / numpy 只是 pypdf 处理图片型 PDF 时的可选依赖，本工具仅统计页数与抽页，
        # 排除后 exe 体积可减小十几 MB
        "--exclude-module", "PIL",
        "--exclude-module", "numpy",
        "--hidden-import", "win32com",
        "--hidden-import", "win32com.client",
        "--hidden-import", "win32print",
        "--hidden-import", "pythoncom",
        "--hidden-import", "pywintypes",
    ])


if __name__ == "__main__":
    main()
