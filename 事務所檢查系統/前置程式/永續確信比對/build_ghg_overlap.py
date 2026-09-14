# -*- coding: utf-8 -*-
"""
安永財簽客戶 × 溫室氣體確信出具者 重疊比對（永續發展路徑圖那條線）

輸入：ghg_targets.json（安永上市櫃客戶中有溫室氣體確信報告者）＋ ghg_verdicts.py（PDF 判讀結果）
輸出：安永財簽×溫室氣體確信_2025.xlsx / .csv / summary_溫室氣體_2025.md
"""
import sys, os, json, re, time
import pandas as pd
from ghg_verdicts import VERDICTS, FLAGS, FAILED

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
HERE = os.path.dirname(os.path.abspath(__file__))
EY_PAT = re.compile(r"安永|Ernst\s*&?\s*Young|EY\s*Taiwan")
TYPE_LABEL = {"cpa": "會計師事務所", "consult": "會計師事務所體系顧問公司",
              "verify": "驗證／查驗機構", "unknown": "未具名"}


def build():
    tg = json.load(open(os.path.join(HERE, "ghg_targets.json"), encoding="utf-8"))
    rows = []
    for t in tg:
        code = t["code"]
        v = VERDICTS.get(code)
        if v:
            issuer, kind, ev = v
        elif code in FAILED:
            issuer, kind, ev = "（下載失敗）", "unknown", "GenPlus 下載端點回 success:false，檔案取不到"
        else:
            issuer, kind, ev = "（未判讀）", "unknown", ""

        is_ey = bool(EY_PAT.search(issuer))
        acct = kind in ("cpa", "consult")
        flag = FLAGS.get(code, "")
        if code in FAILED:
            flag = "檔案下載失敗，出具者無從辨識"

        rows.append({
            "公司代號": code,
            "公司名稱": t["name"],
            "市場別": t["market"],
            "財簽事務所": "安永聯合會計師事務所",
            "財簽會計師": "、".join([x for x in [t.get("auditor1"), t.get("auditor2")]
                                 if x and str(x) not in ("nan", "無", "None")]),
            "確信報告出具者": issuer,
            "出具者類型": TYPE_LABEL[kind],
            "是否會計師體系": "是" if acct else ("不適用" if kind == "unknown" else "否"),
            "確信者是否為安永": "是" if is_ey else ("不適用" if kind == "unknown" else "否"),
            "財簽與溫室氣體確信同一事務所": "是" if is_ey else ("不適用" if kind == "unknown" else "否"),
            "範疇一確信意見": (t.get("op1") or "").strip(),
            "範疇二確信意見": (t.get("op2") or "").strip(),
            "範疇三確信意見": (t.get("op3") or "").strip(),
            "資料來源": f"https://esggenplus.twse.com.tw/inquiry/info/individual?market={0 if t['market']=='上市' else 1}"
                        f"&companyCode={code}&year=2025&lang=zh-TW　｜PDF：ghg_pdf/{code}_{t['name']}.pdf",
            "佐證摘要": ev,
            "需人工確認": ("是：" + flag) if flag else "否",
        })

    df = pd.DataFrame(rows)
    # xlsx 不在此產出：明細併入 build_detail_workbook.py 的兩張 sheet
    csv = os.path.join(HERE, "安永財簽×溫室氣體確信_2025.csv")
    df.to_csv(csv, index=False, encoding="utf-8-sig")
    print(f"→ {csv}")
    return df


def summarize(df):
    n = len(df)
    ey = (df["確信者是否為安永"] == "是").sum()
    cpa = (df["是否會計師體系"] == "是").sum()
    manual = df[df["需人工確認"].str.startswith("是")]
    L = []
    L.append("# 安永財簽客戶 × 溫室氣體排放確信 出具者比對（2025 年度）\n")
    L.append(f"查詢日：{time.strftime('%Y-%m-%d')}　來源：TWSE ESG GenPlus 溫室氣體排放申報＋所附確信報告 PDF 逐份判讀\n")
    L.append("## 母體\n")
    L.append("- 安永簽證之 114 年度上市櫃客戶：221 家")
    L.append("- 其中已申報 2025 年度溫室氣體資料：214 家")
    L.append(f"- 其中已上傳溫室氣體排放確信報告：**{n} 家**（本表母體）\n")
    L.append("## 結果\n")
    L.append("| 出具者類型 | 家數 | 占比 |")
    L.append("|---|---:|---:|")
    for k, lab in [("是", "會計師事務所體系"), ("否", "驗證／查驗機構"), ("不適用", "未具名／取不到檔")]:
        c = (df["是否會計師體系"] == k).sum()
        L.append(f"| {lab} | {c} | {c/n*100:.1f}% |")
    L.append("")
    L.append(f"**其中確信者為安永：{ey} 家（占 {ey/n*100:.1f}%）——即財簽與溫室氣體確信同一事務所的重疊家數。**\n")

    sub = df[df["是否會計師體系"] == "是"].sort_values("公司代號")
    L.append(f"### 會計師體系出具者明細（{len(sub)} 家）\n")
    L.append("| 代號 | 公司 | 出具者 | 類型 | 是否安永 |")
    L.append("|---|---|---|---|---|")
    for _, r in sub.iterrows():
        L.append(f"| {r['公司代號']} | {r['公司名稱']} | {r['確信報告出具者']} | {r['出具者類型']} | **{r['確信者是否為安永']}** |")
    L.append("")

    vc = df[df["是否會計師體系"] == "否"]["確信報告出具者"].value_counts()
    L.append("### 驗證／查驗機構分布\n")
    L.append("| 機構 | 家數 |")
    L.append("|---|---:|")
    for k, v in vc.items():
        L.append(f"| {k} | {v} |")
    L.append("")

    L.append(f"## 需人工確認清單（{len(manual)} 家）\n")
    L.append("| 代號 | 公司 | 事由 |")
    L.append("|---|---|---|")
    for _, r in manual.iterrows():
        L.append(f"| {r['公司代號']} | {r['公司名稱']} | {r['需人工確認'][2:]} |")
    L.append("")

    p = os.path.join(HERE, "summary_溫室氣體_2025.md")
    open(p, "w", encoding="utf-8").write("\n".join(L))
    print(f"→ {p}\n")
    print("\n".join(L))


if __name__ == "__main__":
    summarize(build())
