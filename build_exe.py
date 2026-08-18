"""
把批量打印工具打包成单个 exe（Windows）。

用法：
    pip install pyinstaller
    python build_exe.py

产物：dist/批量打印工具.exe （自包含，内嵌 SumatraPDF 与网页模板，双击即用）

说明：
- --onefile 会把 bin/ 与 templates/ 全部打进 exe，运行时自动解压，
  因此分发给别人只需这一个 exe 文件。
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
        "--add-data", os.path.join(HERE, "bin") + os.pathsep + "bin",
        "--add-data", os.path.join(HERE, "templates") + os.pathsep + "templates",
    ])


if __name__ == "__main__":
    main()
