@echo off
chcp 65001 >nul
cd /d "%~dp0"
py -3 -m pip install -r requirements.txt pyinstaller
if errorlevel 1 (
 echo [오류] Python 또는 패키지 설치 실패
 pause
 exit /b 1
)
py -3 -m PyInstaller --onefile --windowed --clean --name tistory_desktop tistory_desktop.py
echo EXE 파일: dist\tistory_desktop.exe
pause
