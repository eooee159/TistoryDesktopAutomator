@echo off
chcp 65001 >nul
cd /d "%~dp0"
py -3 -m pip install -r requirements.txt
if errorlevel 1 (
 echo [오류] Python 3 설치 여부를 확인하세요.
 pause
 exit /b 1
)
echo 패키지 설치 완료.
pause
