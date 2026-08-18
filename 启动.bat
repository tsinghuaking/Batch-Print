@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo 正在启动批量打印工具，请在浏览器打开 http://127.0.0.1:5000
echo 关闭本窗口即停止服务。
"C:\Users\tsing\AppData\Local\Programs\Python\Python314\python.exe" app.py
pause
