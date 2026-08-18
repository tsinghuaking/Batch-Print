"""
把批量打印工具打包成单个 exe（Windows）。

用法：
    pip install pyinstaller
    python build_exe.py

产物：dist/批量打印工具.exe
- 仅用 Python 标准库，体积小（不含 Flask），约 10MB，可直接作为 GitHub Release 附件。
- templates/ 会打进 exe；SumatraPDF 引擎（bin/）外置，需与 exe 放在同一目录。

说明：
- 保留控制台窗口（方便看日志、关闭即停服务）。如需无控制台窗口，
  在下方 run() 列表里加上 "--windowed"。
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
        "--add-data", os.path.join(HERE, "templates") + os.pathsep + "templates",
    ])


if __name__ == "__main__":
    main()
