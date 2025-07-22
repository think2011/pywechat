@echo off
chcp 65001 >nul
echo 在本地机器上运行微信脚本...
echo.

REM 切换到项目目录
cd /d "%~dp0"

REM 激活虚拟环境（如果存在）
if exist ".venv\Scripts\activate.bat" (
    call .venv\Scripts\activate.bat
    echo 虚拟环境已激活
) else (
    echo 未找到虚拟环境，使用系统Python
)

REM 运行微信脚本
echo 开始运行微信脚本...
python wxdump_adapter.py

echo.
echo 脚本执行完成
pause