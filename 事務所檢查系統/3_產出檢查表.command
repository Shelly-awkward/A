#!/bin/bash
cd "$(dirname "$0")"; source bin/_env.sh || exit 1
INPUT="$1"
if [ -z "$INPUT" ]; then INPUT=$(ls -t input/個案檢查_輸入_*.xlsx 2>/dev/null | head -1); fi
if [ -z "$INPUT" ]; then echo "[錯誤] input/ 找不到 個案檢查_輸入_*.xlsx（從網頁表單匯出，或先跑 2_建立檢查輸入檔.command）"; read -p "按 Enter 關閉"; exit 1; fi
echo "==== 第3步 產出：$INPUT → output/檢查表產出/ ===="
"$PY" bin/render_case_docs.py render --input "$INPUT"
open output/檢查表產出 2>/dev/null || true
read -p "按 Enter 關閉"
