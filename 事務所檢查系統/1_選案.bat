@echo off
chcp 65001 >nul
call "%~dp0bin\_env.bat" || exit /b 1
echo ==== 第1步 選案：讀 input\選案參數_*.xlsx、權數表、auditor 快照、名冊 → output\<年度>年度_<事務所>_選案工作表.xlsx ====
python "%~dp0bin\select_cases.py" %*
echo.
echo 完成後請開 output\ 的選案工作表，確認「建議選案」與「排除事由」；要改規則只改 input\選案參數 再重跑本檔。
pause
