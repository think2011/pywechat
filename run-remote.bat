@echo off
chcp 65001 >nul
echo 在宿主机器上运行微信脚本...

REM 创建临时任务
schtasks /create /tn "PywechatTemp" /tr "cmd /c start /d C:\Users\think2011\workspace\pywechat run.bat" /sc once /st 00:00 /ru %USERNAME% /f

REM 立即运行任务
schtasks /run /tn "PywechatTemp"

REM 等待一下然后删除任务
timeout /t 2 /nobreak >nul
schtasks /delete /tn "PywechatTemp" /f

echo 脚本已在宿主环境启动
