@echo off
chcp 65001 >nul
call "%~dp0bin\_env.bat" || exit /b 1
setlocal enabledelayedexpansion
set INPUT=%~1
if "%INPUT%"=="" (
  for /f "delims=" %%f in ('dir /b /o-d "%~dp0input\個案檢查_輸入_*.xlsx" 2^>nul') do (if not defined INPUT set INPUT=%~dp0input\%%f)
)
if not defined INPUT (echo [錯誤] input\ 找不到 個案檢查_輸入_*.xlsx（可從 Apps Script 表單匯出，或先跑 2_建立檢查輸入檔.bat）& pause & exit /b 1)
echo ==== 第3步 產出：%INPUT% → output\檢查表產出\ （檢查表二/三欄式/檢查表一/AML/彙總表/申報站CSV）====
python "%~dp0bin\render_case_docs.py" render --input "%INPUT%"
pause
