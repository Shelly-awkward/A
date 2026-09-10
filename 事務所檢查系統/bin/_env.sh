#!/bin/bash
# macOS / Linux 共用環境檢查：找 python3，第一次自動裝套件
cd "$(dirname "$0")/.."
PY=$(command -v python3 || true)
if [ -z "$PY" ]; then echo "[錯誤] 找不到 python3。請到 https://www.python.org/downloads/macos/ 安裝 Python 3.11（2017 MacBook Air 可用），或 brew install python@3.11"; read -p "按 Enter 關閉"; exit 1; fi
"$PY" -c "import openpyxl, docx" 2>/dev/null || { echo "第一次執行，安裝需要的套件..."; "$PY" -m pip install -q --user -r bin/requirements.txt; }
export PY
