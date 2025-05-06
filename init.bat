@echo off
REM Main.ps1 스크립트가 있는 폴더로 작업 디렉토리 변경
cd /D "%~dp0\script"

REM PowerShell 실행 및 Main.ps1 스크립트 실행
START /B powershell.exe -ExecutionPolicy Bypass -File "Main.ps1" -InitialPath "%~dp0\script"