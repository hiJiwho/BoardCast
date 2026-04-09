@echo off
echo Boardcast 개발 모드 실행 중...
taskkill /f /im electron.exe 2>nul
npm start
pause
