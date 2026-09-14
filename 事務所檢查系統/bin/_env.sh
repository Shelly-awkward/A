#!/bin/bash
# macOS / Linux 共用環境檢查：找 python3，第一次自動在專案內建 .venv 並安裝套件
# 注意：本檔會被 .command 以 source 載入，不可 cd（以免切換呼叫端目錄）
ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
SYS_PY=""
for c in python3.11 python3.12 python3.13 python3 "$HOME/.local/bin/python3.11" /usr/local/bin/python3.11 /opt/homebrew/bin/python3.11; do
  if command -v "$c" >/dev/null 2>&1; then SYS_PY=$(command -v "$c"); break; fi
done
if [ -z "$SYS_PY" ]; then echo "[錯誤] 找不到 python3。請到 https://www.python.org/downloads/macos/ 安裝 Python 3.11（2017 MacBook Air 可用），或 brew install python@3.11"; read -p "按 Enter 關閉"; exit 1; fi
VENV="$ROOT/.venv"
if [ ! -x "$VENV/bin/python" ]; then
  echo "第一次執行，以 $SYS_PY 建立虛擬環境 .venv ..."
  "$SYS_PY" -m venv "$VENV" || { echo "[錯誤] 建立 .venv 失敗"; read -p "按 Enter 關閉"; exit 1; }
fi
PY="$VENV/bin/python"
"$PY" -c "import openpyxl, docx" 2>/dev/null || { echo "安裝需要的套件..."; "$PY" -m pip install -q --disable-pip-version-check -r "$ROOT/bin/requirements.txt" || { echo "[錯誤] 套件安裝失敗"; read -p "按 Enter 關閉"; exit 1; }; }
export PY
