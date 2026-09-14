#!/bin/bash
cd "$(dirname "$0")"; source bin/_env.sh || exit 1
echo "==== 第1步 選案：input/選案參數_*.xlsx、權數表、auditor 快照、名冊 → output/<年度>年度_<事務所>_選案工作表.xlsx ===="
"$PY" bin/select_cases.py "$@"
echo; echo "完成後開 output/ 的選案工作表確認「建議選案」「排除事由」；改規則只改 input/選案參數 再重跑。"
read -p "按 Enter 關閉"
