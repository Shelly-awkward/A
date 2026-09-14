# 前置程式 — 產生 input/ 五個輸入檔的來源程式

`1_選案` 需要的五個輸入檔，各由下列程式或人工產生。除權數表外都有程式；本資料夾把原本散落在
Google 雲端硬碟「02工作 › 115年度事務所檢查-安永」的程式收進版控（2026-09-13 複製，逐檔以位元組數核對與原檔一致）。

| input/ 檔案 | 產生方式 | 程式位置 | 執行環境 |
|---|---|---|---|
| auditor 快照（母體來源，不直接進 input/） | listdata 的 `執行.bat` → `check_auditor.ps1` → `<年>Q<季>_auditor_<日期>.xlsx` | GitHub `listdata` 倉庫（辦公機 `C:\Users\dinef\AI\listdata`） | Windows |
| `<年季>_選案清單_<事務所>.xlsx`、`_全市場.xlsx` | `python make_selection_list.py --year 114 --season 04 [--firm 安永]` 合併 auditor 快照＋權數表 | `選案清單/make_selection_list.py`（SOP 見同資料夾 `選案SOP.md`） | Windows（路徑寫死 listdata） |
| `<年季>權數表(本國上市上櫃公司)_SFB.xlsx` | 證交所／櫃買中心下載後手動合併成一檔，工作表名含「合併」 | 無程式 | 人工 |
| `四大會計師事務所會計師名冊_<日期>.xlsx` | 自會計師公會全聯會「執業入會查詢」擷取並依分所地址歸類 | **尚未找到原程式**（待辦） | — |
| `安永114客戶_確信總表.csv` | `重跑.bat` 五步驟：build_ey_assurance → fetch_ghg_assurance → build_ghg_overlap → build_combined → build_detail_workbook；新增公司的 PDF 需人工／AI 判讀後填入 `ghg_verdicts.py` | `永續確信比對/`（說明見 `重跑說明.md`） | Windows（`C:\Users\dinef\AI\projects\FR_extract`） |

## 注意
- 這些程式內含 Windows 絕對路徑（`C:\Users\dinef\AI\...`），在 Mac 上不能直接跑；要在 Mac 跑需改成相對路徑或參數化（待辦）。
- 「事務所檢查系統」本體（bin/、1_選案 等）只讀 input/ 內的成品檔，與這裡的前置程式互不相依。
- 改年度時：先跑這裡的程式產生新輸入檔放進 input/，再依 00_操作手冊 執行 1_選案。
