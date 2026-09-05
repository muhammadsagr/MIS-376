@echo off
REM تشغيل البرنامج بالواجهة الرسومية على ويندوز
cd /d "%~dp0"
py -3 main.py || python main.py
pause
