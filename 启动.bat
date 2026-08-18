@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo [源码模式] 已打包 exe 的用户请直接双击 dist\批量打印工具.exe
echo 浏览器将打开 http://127.0.0.1:5001 ，关闭本窗口即停止服务。
python app.py
pause
