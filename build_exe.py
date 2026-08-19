"""
把批量打印工具打包成单个 exe（Windows）。

用法：
    pip install pyinstaller
    python build_exe.py

产物：dist/批量打印工具.exe
- 仅用 Python 标准库，不含 Flask。
- templates/ 与 bin/（SumatraPDF 引擎）全部内嵌进单个 exe，双击即用、自包含。
- 若把 exe 同目录放了 bin/（外置 SumatraPDF），会优先用外置版，便于单独更新引擎。

说明：
- 无控制台黑窗口（--windowed）。运行日志写入 exe 同目录「运行日志.txt」，
  页面右上角「退出程序」按钮可停止后台服务。
"""
import os
import PyInstaller.__main__

HERE = os.path.dirname(os.path.abspath(__file__))


def main():
    PyInstaller.__main__.run([
        os.path.join(HERE, "app.py"),
        "--name", "批量打印工具",
        "--onefile",
        "--noconfirm",
        "--windowed",
        "--icon", os.path.join(HERE, "assets", "Batch Print.ico"),
        "--add-data", os.path.join(HERE, "templates") + os.pathsep + "templates",
        "--add-data", os.path.join(HERE, "bin") + os.pathsep + "bin",
        "--add-data", os.path.join(HERE, "assets") + os.pathsep + "assets",
    ])


if __name__ == "__main__":
    main()
