# 事務所檢查系統 — 給本機 Claude Code session 的專案說明

## 這是什麼
金管會證期局會計審計組「會計師事務所檢查」的自動化系統。平時由同仁不靠 AI 操作；AI（你）負責維護程式、處理改版與除錯。
115 年度受查事務所：安永（另有立本）。專案控管機：2017 MacBook Air；辦公室另有 Windows 機。

## 先讀
1. `00_操作手冊.md`：三步流程（選案 → 網頁表單輸入 → 產出）、資料夾、部署、改版方式。
2. `進度.md`：模組狀態與待辦。**每次改動後更新它**，這是使用者看進度的唯一地方。
3. `參考/選案規則說明_115安永.md`：選案規則與本次 15 家建議名單的由來。

## 結構
- `bin/select_cases.py` 選案；`bin/render_case_docs.py` 產出；`bin/items_chk2_115.py` 檢查表二題庫；`bin/make_seed.py` 產生表單題庫 `apps_script/Seed.gs`
- `apps_script/` 多人網頁表單（Google Apps Script，部署在共用帳號 accsfb01 的試算表）；`apps_script_缺失申報站/` 缺失申報站
- `input/` 每年更新的資料與 `選案參數_*.xlsx`（所有規則數字都在這裡，不寫死在程式）
- `templates/<年度>/` 原始 Word/Excel 範本；`output/` 產出（不進版控）
- 一鍵執行：Windows `*.bat`，Mac `*.command`（`bin/_env.sh` 共用）

## 工作規則
- 規則調整優先改 `input/選案參數_*.xlsx`，不改程式；只有新增規則型態才動 `select_cases.py`。
- 範本改版：新檔放 `templates/<新年度>/`，欄位結構相同不用改程式；檢查表二換題目改 `items_chk2_<年度>.py` 後跑 `python3 bin/make_seed.py`，再重新部署 Seed.gs。
- 產出格式以「與原檔一致」為準；檢查表二、三欄式對照表目前是重建版面（原檔為 .doc），拿到 .docx 就改為套原檔。
- 改完程式一定實跑：`1_選案.command` 與 `3_產出檢查表.command`（或 `python3 bin/select_cases.py`、`python3 bin/render_case_docs.py render --input ...`），確認 output 有檔且能開啟。
- Git：在分支 `claude/survey-part-i-iv-check-nttsu7` 上工作（或使用者指定的新分支）；commit 訊息用中文說明改了什麼；`output/` 不進版。
- 不要動 `input/` 裡的原始資料（權數表、快照、名冊、確信總表），除非使用者提供新版。
- 沙盒／遠端 session 無法連 Google，Apps Script 的部署與實測只能在本機或瀏覽器做。

## 目前待辦（詳見 進度.md）
- 在 Mac 上跑通 `0_首次設定_mac.command` → `1_選案.command`
- 以 accsfb01 部署網頁表單並實測 setup／匯入個案／填答／匯出 xlsx → `3_產出檢查表.command`
- 取得 114 年度實質審閱名單後填入參數表重跑，目標家數改 8
- 總結會議紀錄範本到位後加入產出；檢查表一底稿版由作答帶入
