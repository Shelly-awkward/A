#!/bin/bash
cd "$(dirname "$0")"; source bin/_env.sh || exit 1
echo "==== 第2步 建立個案檢查輸入檔（無網頁表單時用）：output/ 選案工作表「建議選案」→ input/個案檢查_輸入_<年度><事務所>.xlsx ===="
"$PY" bin/render_case_docs.py init "$@"
read -p "按 Enter 關閉"
