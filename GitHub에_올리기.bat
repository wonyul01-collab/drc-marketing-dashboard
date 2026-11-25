@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo GitHub에 올리는 중입니다...
git push
echo.
echo 끝났습니다. 아무 키나 누르면 창이 닫힙니다.
pause >nul
