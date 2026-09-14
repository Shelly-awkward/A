#!/bin/bash
# 第一次在 Mac 上使用：讓 .command 可雙擊執行、安裝套件
cd "$(dirname "$0")"
chmod +x *.command bin/*.sh
xattr -dr com.apple.quarantine . 2>/dev/null || true
source bin/_env.sh && "$PY" -c "import openpyxl, docx; print('套件 OK，Python', __import__('sys').version.split()[0])"
echo "設定完成。之後直接雙擊 1_選案.command / 3_產出檢查表.command。"
read -p "按 Enter 關閉"
