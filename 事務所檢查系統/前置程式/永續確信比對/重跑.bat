@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ============================================
echo  EY x ESG assurance - full refresh
echo ============================================
echo.

echo [1/5] Sustainability report filings...
python build_ey_assurance.py --refresh
if errorlevel 1 goto :err

echo.
echo [2/5] Downloading GHG assurance PDFs...
rem fetch_ghg_assurance.py has no --refresh: delete the cache or it reuses last year's snapshot
if exist raw_ghg_2025.json del raw_ghg_2025.json
python fetch_ghg_assurance.py --download --extract
if errorlevel 1 goto :err

echo.
echo [3/5] GHG overlap table...
python build_ghg_overlap.py
if errorlevel 1 goto :err

echo.
echo [4/5] Combined master table...
python build_combined.py
if errorlevel 1 goto :err

echo.
echo [5/5] Detail workbook...
python build_detail_workbook.py
if errorlevel 1 goto :err

echo.
echo ============================================
echo  DONE.
echo.
echo  Open: an-yong-114 master table + detail workbook
echo  NOTE: newly-added companies show "(未判讀)" in the
echo        issuer column - their PDFs still need a human
echo        (or Claude) to read them. See 重跑說明.md
echo ============================================
pause
exit /b 0

:err
echo.
echo *** FAILED - see message above ***
pause
exit /b 1
