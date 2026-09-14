# -*- coding: utf-8 -*-
"""
抓取安永財簽客戶的「溫室氣體排放確信報告」PDF，辨識出具者。

來源：TWSE ESG GenPlus
  溫室氣體資料  POST /api/api/mopsEsg/summaryData          （欄位 greenhouseGasEmissionsAssertionReport = 上傳檔 id）
  確信報告下載  GET  /api/api/MopsSustainQuestionnaire/{id}/download

用法：
  python fetch_ghg_assurance.py --download    # 下載 PDF（節制：每檔間隔 1 秒）
  python fetch_ghg_assurance.py --extract     # 抽文字 → extract.json
"""
import sys, os, io, json, time, re, argparse
import requests
import pandas as pd

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

HERE = os.path.dirname(os.path.abspath(__file__))
PDF_DIR = os.path.join(HERE, "ghg_pdf")
AUDITOR_XLSX = r"C:\Users\dinef\AI\listdata\114Q04_auditor_20260716.xlsx"
BASE = "https://esggenplus.twse.com.tw"
INDS = ["水泥工業","食品工業","塑膠工業","紡織纖維","電機機械","電器電纜","化學工業","生技醫療業",
        "化學生技醫療","玻璃陶瓷","造紙工業","鋼鐵工業","橡膠工業","汽車工業","半導體業",
        "電腦及週邊設備業","光電業","通信網路業","電子零組件業","電子通路業","資訊服務業",
        "其他電子業","電子工業","油電燃氣業","建材營造","航運業","觀光餐旅","金融保險業",
        "貿易百貨","綜合企業","綠能環保","數位雲端","運動休閒","居家生活","其他","存託憑證"]


def session():
    s = requests.Session()
    s.headers.update({"User-Agent": "Mozilla/5.0", "Referer": f"{BASE}/inquiry/info/individual"})
    tok = s.get(f"{BASE}/api/api/Antiforgery/token", timeout=30).json()["data"]
    return s, tok


def flatten(x):
    d = {"code": str(x["companyCode"]).strip(), "name": x["companyName"]}
    for sec in x.get("sectionDatas") or []:
        for c in sec.get("controls") or []:
            if c.get("code"):
                d[c["code"]] = c.get("value")
    return d


def ey_targets(year=2025):
    """回傳 [(code, 簡稱, uploadId, 確信意見範疇一)] — 安永上市櫃客戶中有溫室氣體確信報告者"""
    ghg_path = os.path.join(HERE, f"raw_ghg_{year}.json")
    if os.path.exists(ghg_path):
        raw = json.load(open(ghg_path, encoding="utf-8"))
    else:
        s, tok = session()
        raw = {}
        for mt, lab in [(0, "上市"), (1, "上櫃")]:
            r = s.post(f"{BASE}/api/api/mopsEsg/summaryData",
                       headers={"X-CSRF-TOKEN": tok, "Content-Type": "application/json"},
                       json={"marketType": mt, "industryYearId": None, "industryNameList": INDS,
                             "sampleId": "溫室氣體排放", "year": year, "sampleName": "溫室氣體排放"},
                       timeout=120)
            raw[lab] = r.json().get("data") or []
            time.sleep(1)
        json.dump(raw, open(ghg_path, "w", encoding="utf-8"), ensure_ascii=False)

    ghg = {}
    for lst in raw.values():
        for x in lst:
            f = flatten(x)
            ghg[f["code"]] = f

    df = pd.read_excel(AUDITOR_XLSX, sheet_name="114Q04")
    df["公司代號"] = df["公司代號"].astype(str).str.strip()
    ey = (df[df["事務所"].astype(str).str.contains("安永")
             & df["市場別"].isin(["上市", "上櫃"])]
          .drop_duplicates("公司代號").sort_values("公司代號"))

    out = []
    for _, a in ey.iterrows():
        x = ghg.get(a["公司代號"])
        if not x:
            continue
        uid = x.get("greenhouseGasEmissionsAssertionReport")
        if not uid:
            continue
        out.append({
            "code": a["公司代號"], "name": a["公司簡稱"], "market": a["市場別"],
            "auditor1": a.get("會計師1"), "auditor2": a.get("會計師2"),
            "uid": uid,
            "op1": x.get("grossScope1ConfidenceOpinion"),
            "op2": x.get("grossScope2ConfidenceOpinion"),
            "op3": x.get("grossScope3ConfidenceOpinion"),
            "scope1": x.get("grossScope1ConfidenceScope"),
        })
    return out


def download(targets):
    os.makedirs(PDF_DIR, exist_ok=True)
    s, tok = session()
    ok = skip = fail = 0
    for i, t in enumerate(targets, 1):
        path = os.path.join(PDF_DIR, f"{t['code']}_{t['name']}.pdf")
        if os.path.exists(path) and os.path.getsize(path) > 1000:
            skip += 1
            continue
        try:
            r = s.get(f"{BASE}/api/api/MopsSustainQuestionnaire/{t['uid']}/download",
                      headers={"X-CSRF-TOKEN": tok}, timeout=120)
            if r.status_code == 200 and r.content[:4] == b"%PDF":
                open(path, "wb").write(r.content)
                ok += 1
                print(f"  [{i}/{len(targets)}] {t['code']} {t['name']}  {len(r.content)/1024:.0f} KB")
            else:
                fail += 1
                print(f"  [{i}/{len(targets)}] {t['code']} {t['name']}  失敗 {r.status_code} {r.content[:40]}")
        except Exception as e:
            fail += 1
            print(f"  [{i}/{len(targets)}] {t['code']} {t['name']}  例外 {type(e).__name__}")
        time.sleep(1.0)   # 節制，不灌爆對方
    print(f"下載完成：新增 {ok}、既有 {skip}、失敗 {fail}")


def extract(targets):
    import fitz
    res = []
    for t in targets:
        path = os.path.join(PDF_DIR, f"{t['code']}_{t['name']}.pdf")
        rec = dict(t)
        if not os.path.exists(path):
            rec.update({"pages": 0, "chars": 0, "text": "", "note": "檔案不存在"})
            res.append(rec)
            continue
        d = fitz.open(path)
        txt = "\n".join(p.get_text() for p in d)
        rec.update({"pages": d.page_count, "chars": len(txt.strip()),
                    "text": txt, "note": "" if len(txt.strip()) > 80 else "純掃描影像，需視覺判讀"})
        d.close()
        res.append(rec)
    json.dump(res, open(os.path.join(HERE, "ghg_extract.json"), "w", encoding="utf-8"),
              ensure_ascii=False, indent=1)
    n_img = sum(1 for r in res if r["note"])
    print(f"抽取完成：{len(res)} 份，其中 {n_img} 份為純掃描影像需視覺判讀")
    return res


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--download", action="store_true")
    ap.add_argument("--extract", action="store_true")
    a = ap.parse_args()
    tg = ey_targets()
    print(f"安永上市櫃客戶中有溫室氣體確信報告者：{len(tg)} 家")
    json.dump(tg, open(os.path.join(HERE, "ghg_targets.json"), "w", encoding="utf-8"),
              ensure_ascii=False, indent=1)
    if a.download:
        download(tg)
    if a.extract:
        extract(tg)
