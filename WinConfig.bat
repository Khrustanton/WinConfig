@echo off
chcp 65001 >nul
echo Запуск скрипта WinConfig.ps1...

:: Переход в папку, где находится этот bat-файл
cd /d "%~dp0"

:: Запуск PowerShell с изменением политики выполнения и запуском скрипта
powershell -Command "Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope Process -Force; .\WinConfig.ps1"

echo Скрипт выполнен!
pause