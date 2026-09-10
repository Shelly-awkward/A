@echo off
chcp 65001 >nul
call "%~dp0bin\_env.bat" || exit /b 1
echo ==== 第2步 建立個案檢查輸入檔：自 output\ 選案工作表「建議選案」帶入個案 → input\個案檢查_輸入_<年度><事務所>.xlsx ====
echo （若改用 Apps Script 網頁表單，此檔可不用；表單匯出的 xlsx 直接放 input\ 即可）
python "%~dp0bin\render_case_docs.py" init %*
pause
