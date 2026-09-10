@echo off
chcp 65001 >nul
where python >nul 2>nul || (echo [錯誤] 找不到 python，請先安裝 Python 3.10 以上並勾選 Add to PATH。& pause & exit /b 1)
python -c "import openpyxl, docx" 2>nul || (echo 第一次執行，安裝需要的套件... & python -m pip install -q -r "%~dp0requirements.txt")
