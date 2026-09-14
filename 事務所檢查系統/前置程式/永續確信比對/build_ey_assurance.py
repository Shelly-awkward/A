# -*- coding: utf-8 -*-
"""
安永財簽客戶 × 永續報告書確信 比對

輸入：
  A. 本機財簽名單 C:\\Users\\dinef\\AI\\listdata\\114Q04_auditor_*.xlsx（欄位：市場別/公司代號/公司簡稱/事務所/會計師1/會計師2/核閱日期/核閱意見）
  B. TWSE ESG GenPlus  POST /api/api/MopsSustainReport/data
     （＝MOPS 永續報告書申報專區結構化欄位，含確信/查證申報格）

用法：
  python build_ey_assurance.py            # 用快取 raw_mops.json（沒有才連線）
  python build_ey_assurance.py --refresh   # 強制重抓
  python build_ey_assurance.py --year 2024 # 換報告年度
"""
import sys, os, json, time, argparse, re
import requests
import pandas as pd

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

HERE = os.path.dirname(os.path.abspath(__file__))
AUDITOR_XLSX = r"C:\Users\dinef\AI\listdata\114Q04_auditor_20260716.xlsx"
AUDITOR_SHEET = "114Q04"
BASE = "https://esggenplus.twse.com.tw"
MARKETS = {0: "上市", 1: "上櫃"}

# 安永在台實體名稱變體（事務所 + 顧問公司都算「安永」）
EY_PAT = re.compile(r"安永|Ernst\s*&\s*Young|EY\s*Taiwan|\bEY\b", re.I)
# 判斷出具實體是「事務所」還是「顧問/其他公司」
FIRM_PAT = re.compile(r"會計師事務所")


def fetch_mops(year: int) -> dict:
    """抓 MOPS 永續報告書申報專區（經 GenPlus API）。節制：每次請求間隔 1 秒。"""
    s = requests.Session()
    s.headers.update({
        "User-Agent": "Mozilla/5.0",
        "Referer": f"{BASE}/inquiry/report",
    })
    token = s.get(f"{BASE}/api/api/Antiforgery/token", timeout=30).json()["data"]
    H = {"X-CSRF-TOKEN": token, "Content-Type": "application/json"}
    out = {}
    for mt, label in MARKETS.items():
        body = {"companyCodeList": [], "year": year, "industryNameList": [],
                "marketType": mt, "industryName": "all", "companyCode": "all"}
        r = s.post(f"{BASE}/api/api/MopsSustainReport/data", headers=H, json=body, timeout=60)
        r.raise_for_status()
        out[label] = r.json().get("data") or []
        print(f"  {label}：{len(out[label])} 家已申報 {year} 年度永續報告書")
        time.sleep(1.0)
    return out


def blank(v) -> bool:
    return v is None or str(v).strip() in ("", "無", "N/A", "nan")


def txt(v) -> str:
    return "" if blank(v) else str(v).strip()


def build(year: int, refresh: bool):
    raw_path = os.path.join(HERE, f"raw_mops_{year}.json")
    if refresh or not os.path.exists(raw_path):
        print(f"抓取 {year} 年度永續報告書申報資料…")
        raw = fetch_mops(year)
        json.dump(raw, open(raw_path, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    else:
        print(f"沿用快取 {os.path.basename(raw_path)}（--refresh 可強制重抓）")
        raw = json.load(open(raw_path, encoding="utf-8"))

    mops, mkt_of = {}, {}
    for label, lst in raw.items():
        for x in lst:
            code = str(x["code"]).strip()
            mops[code] = x
            mkt_of[code] = label

    # --- A. 安永上市櫃財簽名單 ---
    df = pd.read_excel(AUDITOR_XLSX, sheet_name=AUDITOR_SHEET)
    df["公司代號"] = df["公司代號"].astype(str).str.strip()
    ey = (df[df["事務所"].astype(str).str.contains("安永")
             & df["市場別"].isin(["上市", "上櫃"])]
          .drop_duplicates("公司代號")
          .sort_values("公司代號"))
    print(f"安永財簽上市櫃客戶：{len(ey)} 家")

    rows = []
    for _, a in ey.iterrows():
        code = a["公司代號"]
        m = mops.get(code)
        cpny_url = f"{BASE}/inquiry/report?market={0 if a['市場別']=='上市' else 1}&year={year}&companyCode={code}&lang=zh-TW"

        r = {
            "公司代號": code,
            "公司名稱": a["公司簡稱"],
            "市場別": a["市場別"],
            "財簽事務所": a["事務所"],
            "財簽會計師": "、".join([x for x in [txt(a.get("會計師1")), txt(a.get("會計師2"))] if x and x != "無"]),
        }

        if m is None:
            r.update({
                "永續報告書狀態": "尚未出版",
                "報告書年度": f"{year+1911-1911}",  # placeholder, 覆寫於下
                "是否有會計師確信": "不適用",
                "確信出具者": "", "確信準則": "", "確信等級": "", "確信範圍": "",
                "是否有第三方查證": "不適用", "查證機構": "", "查證標準": "",
                # 尚未出版者無從判斷，填「不適用」而非「無確信」，避免與「已出版但無確信」混為一談
                "確信者是否為安永": "不適用",
                "財簽與永續確信同一事務所": "不適用",
                "資料來源": cpny_url,
                "佐證摘要": f"截至查詢日，MOPS 永續報告書申報專區（經 TWSE ESG GenPlus）查無 {year} 年度申報紀錄",
                "需人工確認": "否",
            })
            r["報告書年度"] = f"{year-1911}年度（{year}）尚未申報"
        else:
            partner = txt(m.get("partnerOrganization"))   # 確信事務所
            std = txt(m.get("stdCompliance"))             # 確信準則
            level = txt(m.get("accOptionType"))           # 確信等級
            scope = txt(m.get("aConfirmTarget"))          # 確信範圍
            vali = txt(m.get("valiProvider"))             # 查證機構
            tstd = txt(m.get("thirdStd"))                 # 查證標準
            tver = txt(m.get("thirdPartyVerify"))         # 查證內容

            has_assr = bool(partner or std or level or scope)
            has_veri = bool(vali or tstd or tver)
            is_ey = bool(partner and EY_PAT.search(partner))

            interval = txt(m.get("reportingInterval"))
            ryear = f"{year-1911}年度（{year}）" + (f"　{interval}" if interval else "")

            NULL_GUID = "00000000-0000-0000-0000-000000000000"
            assr_file = txt(m.get("aReportUploadId")) not in ("", NULL_GUID)      # 確信報告附件
            veri_file = txt(m.get("thirdPartyUploadId")) not in ("", NULL_GUID)   # 查證聲明書附件

            # 需人工確認規則
            flags = []
            if not has_assr and not has_veri:
                if not assr_file and not veri_file:
                    flags.append("申報表確信與查證欄位全空，且未上傳確信報告或查證聲明書附件"
                                 "——申報面證據一致指向無確信、無查證，惟仍須翻報告書 PDF 附錄複核")
                else:
                    flags.append("申報表確信與查證欄位全空但有附件上傳——申報不一致，須翻 PDF 確認")
            if (has_assr and not assr_file):
                flags.append("填有確信資訊但未上傳確信報告附件")
            if has_assr and not std:
                flags.append("有確信出具者但未填確信準則")
            if has_assr and not level:
                flags.append("有確信出具者但未填確信等級")
            if has_assr and not partner:
                flags.append("填了確信準則/等級但未填出具者")
            if has_veri and not vali:
                flags.append("有查證內容但未填查證機構")
            if has_veri and not tstd:
                flags.append("有查證機構但未填查證標準")
            if is_ey and not FIRM_PAT.search(partner):
                flags.append("確信出具者為安永體系但非會計師事務所（顧問公司）——獨立性意涵不同，須人工判讀")
            if interval and interval != f"{year}/01/01~{year}/12/31":
                flags.append(f"報導期間非曆年（{interval}）")

            ev = []
            if has_assr:
                ev.append("確信：" + "／".join([x for x in [partner or "(未填出具者)", std, level] if x])
                          + ("（附確信報告）" if assr_file else "（未附確信報告檔）"))
                if scope:
                    ev.append("確信範圍：" + scope)
            if has_veri:
                ev.append("查證：" + "／".join([x for x in [vali or "(未填機構)", tstd] if x])
                          + ("（附查證聲明書）" if veri_file else "（未附查證聲明書檔）"))
                if tver:
                    ev.append("查證內容：" + tver)
            if not ev:
                ev.append("申報表確信欄（確信事務所/準則/等級/範圍）與查證欄（查證機構/標準/內容）均空白；"
                          "確信報告與查證聲明書附件亦均未上傳"
                          if not assr_file and not veri_file else
                          "申報表確信與查證文字欄均空白，惟有附件上傳")

            r.update({
                "永續報告書狀態": "已出版",
                "報告書年度": ryear,
                "是否有會計師確信": "是" if has_assr else "否",
                "確信出具者": partner,
                "確信準則": std,
                "確信等級": level,
                "確信範圍": scope,
                "是否有第三方查證": "是" if has_veri else "否",
                "查證機構": vali,
                "查證標準": tstd,
                "確信者是否為安永": ("是" if is_ey else "否") if has_assr else "無確信",
                "財簽與永續確信同一事務所": ("是" if is_ey else "否") if has_assr else "不適用",
                "資料來源": cpny_url + ("　｜報告書：" + txt(m.get("twDocLink")) if txt(m.get("twDocLink")) else ""),
                "佐證摘要": "；".join(ev),
                "需人工確認": ("是：" + "；".join(flags)) if flags else "否",
            })
        rows.append(r)

    cols = ["公司代號", "公司名稱", "市場別", "財簽事務所", "財簽會計師",
            "永續報告書狀態", "報告書年度",
            "是否有會計師確信", "確信出具者", "確信準則", "確信等級", "確信範圍",
            "是否有第三方查證", "查證機構", "查證標準",
            "確信者是否為安永", "財簽與永續確信同一事務所",
            "資料來源", "佐證摘要", "需人工確認"]
    out = pd.DataFrame(rows)[cols]

    # xlsx 不在此產出：明細併入 build_detail_workbook.py 的兩張 sheet
    csv = os.path.join(HERE, f"安永財簽×永續確信比對_{year-1911}年度.csv")
    out.to_csv(csv, index=False, encoding="utf-8-sig")
    print(f"→ {csv}")
    return out, year


def summarize(out: pd.DataFrame, year: int):
    n = len(out)
    pub = (out["永續報告書狀態"] == "已出版").sum()
    assr = (out["是否有會計師確信"] == "是").sum()
    veri = (out["是否有第三方查證"] == "是").sum()
    ey_assr = (out["確信者是否為安永"] == "是").sum()
    overlap = (out["財簽與永續確信同一事務所"] == "是").sum()
    manual = out[out["需人工確認"].str.startswith("是")]

    def pct(a, b):
        return f"{a/b*100:.1f}%" if b else "—"

    L = []
    L.append(f"# 安永財簽客戶 × 永續報告書確信 比對 — {year-1911} 年度（{year} 會計年度）\n")
    L.append(f"查詢日：{time.strftime('%Y-%m-%d')}\n")
    L.append("## 母體\n")
    L.append(f"- 安永聯合會計師事務所簽證之**上市櫃**公司：**{n} 家**"
             f"（上市 {(out['市場別']=='上市').sum()}、上櫃 {(out['市場別']=='上櫃').sum()}）")
    L.append(f"- 來源：`listdata\\114Q04_auditor_20260716.xlsx`（MOPS 會計師查核意見申報，114Q4）\n")
    L.append("## 結果\n")
    L.append("| 指標 | 家數 | 占母體 | 占已出版 |")
    L.append("|---|---:|---:|---:|")
    L.append(f"| 已出版 {year-1911} 年度永續報告書 | {pub} | {pct(pub,n)} | — |")
    L.append(f"| 尚未出版 | {n-pub} | {pct(n-pub,n)} | — |")
    L.append(f"| 有會計師確信 | {assr} | {pct(assr,n)} | {pct(assr,pub)} |")
    L.append(f"| 有第三方查證 | {veri} | {pct(veri,n)} | {pct(veri,pub)} |")
    L.append(f"| **確信者為安永** | **{ey_assr}** | {pct(ey_assr,n)} | {pct(ey_assr,pub)} |")
    L.append(f"| **財簽與永續確信同一事務所（重疊）** | **{overlap}** | {pct(overlap,n)} | {pct(overlap,pub)} |")
    L.append("")

    sub = out[out["是否有會計師確信"] == "是"]
    if len(sub):
        L.append("### 有會計師確信者明細\n")
        L.append("| 代號 | 公司 | 確信出具者 | 準則 | 等級 | 是否安永 | 重疊 |")
        L.append("|---|---|---|---|---|---|---|")
        for _, r in sub.iterrows():
            L.append(f"| {r['公司代號']} | {r['公司名稱']} | {r['確信出具者']} | {r['確信準則']} | "
                     f"{r['確信等級']} | {r['確信者是否為安永']} | {r['財簽與永續確信同一事務所']} |")
        L.append("")

    sub = out[(out["是否有第三方查證"] == "是")]
    if len(sub):
        L.append("### 有第三方查證者明細\n")
        L.append("| 代號 | 公司 | 查證機構 | 查證標準 |")
        L.append("|---|---|---|---|")
        for _, r in sub.iterrows():
            L.append(f"| {r['公司代號']} | {r['公司名稱']} | {r['查證機構']} | {r['查證標準']} |")
        L.append("")

    L.append(f"## 需人工確認清單（{len(manual)} 家）\n")
    if len(manual) == 0:
        L.append("無。\n")
    else:
        L.append("| 代號 | 公司 | 狀態 | 原因 |")
        L.append("|---|---|---|---|")
        for _, r in manual.iterrows():
            L.append(f"| {r['公司代號']} | {r['公司名稱']} | {r['永續報告書狀態']} | {r['需人工確認'][2:]} |")
        L.append("")

    p = os.path.join(HERE, f"summary_{year-1911}年度.md")
    open(p, "w", encoding="utf-8").write("\n".join(L))
    print(f"→ {p}")
    print("\n".join(L[:30]))


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--year", type=int, default=2025)
    ap.add_argument("--refresh", action="store_true")
    a = ap.parse_args()
    out, y = build(a.year, a.refresh)
    summarize(out, y)
