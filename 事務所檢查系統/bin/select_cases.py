# -*- coding: utf-8 -*-
"""
select_cases.py — 事務所檢查「選案模組」（v1，115年度安永首用）
=====================================================================
流程：選案表輸入（選案參數.xlsx）→ 事務所選取規則（每2年/每6年）→ 個案公司篩選規則 → 產出選案工作表

用法：
  python select_cases.py                       # 使用同資料夾預設檔名
  python select_cases.py --params 選案參數_115安永.xlsx --out 115年度_安永_選案工作表.xlsx

輸入檔（同資料夾）：
  1. 114Q4_選案清單_<事務所>.xlsx   母體（listdata auditor 快照；make_selection_list.py 產出）
  2. 114Q4_選案清單_全市場.xlsx     全市場母體 → 各事務所簽證家數（事務所選取規則用）
  3. 114Q4權數表(本國上市上櫃公司)_SFB.xlsx  官方權數表（工作表「114Q4上市上櫃合併」，固定索引 65/83/84/85）
  4. 安永114客戶_確信總表.csv        永續報告書／溫室氣體確信比對（可缺）
  5. 選案參數_<年度><事務所>.xlsx     選案表輸入（若不存在會自動產生預設檔，改完再跑一次）
  6. 四大會計師事務所會計師名冊_*.xlsx  會計師公會執業名冊（會計師→所屬分所；可缺，缺時以客戶所在地推定）

規則全部寫在「選案參數」活頁簿，程式不藏數字；每一條規則在母體清單上都有對應欄位可追溯。
"""
import argparse, csv, collections, datetime, os, re, sys, warnings
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
warnings.filterwarnings("ignore")

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
IN_DIR = os.path.join(ROOT, "input"); OUT_DIR = os.path.join(ROOT, "output")
FONT = "Microsoft JhengHei"
IDX_FIN_ITEMS, IDX_FIN_TOTAL = range(56, 65), 65
IDX_NF_ITEMS, IDX_NF_TOTAL, IDX_TOTAL, IDX_MKT = range(66, 83), 83, 84, 85
HDR_ROWS = 5

# ───────────────────────── 參數（預設值；以「選案參數」活頁簿為準） ─────────────────────────
DEFAULT_PARAMS = [
    # (鍵, 值, 說明)
    ("檢查年度", 115, "本次檢查年度（民國）"),
    ("受查事務所", "安永", "事務所名稱關鍵字（比對母體「事務所」欄）"),
    ("資料季別", "114Q4", "權數表／auditor 快照季別"),
    ("目標家數", 15, "本次擬抽選個案數（先列15家，取得實質審閱名單後再篩至8家；113年安永為8家）"),
    ("分數門檻", 60, "權數總計低於此分數者不列入建議名單（僅列備選）"),
    ("台北所上限", 7, "建議名單中台北所簽證個案之上限家數（地域分配）"),
    ("每所至少", 1, "桃園/新竹/台中/台南/高雄所各至少抽選家數（該所有達門檻之候選者才適用）"),
    ("新承接加分", 20, "備註含「新承接」者加分（權數表抓不到的風險）"),
    ("繼續經營加分", 30, "查核意見含「繼續經營」者加分"),
    ("保留意見加分", 30, "查核意見為保留/否定/無法表示意見者加分"),
    ("其他事務所查核加分", 5, "意見含「採用其他會計師查核」（集團查核依賴他人）者加分"),
    ("確信重疊加分", 10, "財簽事務所同時出具永續報告書或溫室氣體確信者加分（新檢查重點）"),
    ("金融業排除", "是", "銀行/證券/保險/金控（另有監理）不列入"),
    ("前次已查排除", "是", "本會前次檢查（前一輪）已抽查之個案排除"),
    ("實質審閱排除", "是", "114年度財報業經證交所/櫃買中心實質審閱者排除（名單填於「實質審閱名單」工作表）"),
    ("每2年門檻", 100, "簽證公開發行公司家數 ≥ 此值 → 每2年檢查1次"),
    ("每6年門檻", 10, "簽證公開發行公司家數 ≥ 此值（未達每2年門檻）→ 每6年檢查1次；未達者視需要"),
]
DEFAULT_PRIOR = [  # (年度, 代號, 簡稱, 備註)
    (113, "2312", "金寶", "113年檢查個案（含永續確信）"), (113, "3284", "太普高", "113年檢查個案"),
    (113, "1591", "駿吉-KY", "113年檢查個案"), (113, "6287", "元隆", "113年檢查個案"),
    (113, "2637", "慧洋-KY", "113年檢查個案"), (113, "6841", "長佳", "113年檢查個案"),
    (113, "3167", "大量", "113年檢查個案"), (113, "6603", "富強鑫", "113年檢查個案"),
    (111, "1446", "宏和", "111年檢查個案"), (111, "8462", "柏文", "111年檢查個案"),
    (111, "5381", "合正", "111年檢查個案"), (111, "1236", "宏亞", "111年檢查個案"),
    (111, "1258", "其祥-KY", "111年檢查個案"), (111, "4503", "金雨", "111年檢查個案"),
    (111, "3064", "泰偉", "111年檢查個案"), (111, "5514", "三豐", "111年檢查個案"),
]
DEFAULT_DESIGNATED = [  # (代號, 簡稱, 處置, 理由)
    ("3017", "奇鋐", "指定納入", "安永同時出具114年度溫室氣體確信報告，擬一併調閱確信工作底稿"),
]
DEFAULT_FIRM_HISTORY = [  # (事務所關鍵字, 前次檢查年度, 備註)
    ("勤業眾信", 114, ""), ("安侯建業", 114, ""), ("資誠", 113, ""), ("安永", 113, ""),
    ("立本台灣", 115, "110年前次；115年9月檢查"), ("國富浩華", None, "請填前次檢查年度"),
    ("正風", None, "請填前次檢查年度"), ("馬施云大華", None, ""), ("德昌", 114, ""), ("建智", None, ""),
]
REGION_OFFICE = {  # 公司/會計師客戶所在縣市 → 安永分所
    "台北": "台北所", "臺北": "台北所", "新北": "台北所", "基隆": "台北所", "桃園": "桃園所", "宜蘭": "台北所", "花蓮": "台北所", "台東": "台北所", "金門": "台北所",
    "新竹": "新竹所", "苗栗": "新竹所",
    "台中": "台中所", "臺中": "台中所", "彰化": "台中所", "南投": "台中所", "雲林": "台中所",
    "嘉義": "台南所", "台南": "台南所", "臺南": "台南所",
    "高雄": "高雄所", "屏東": "高雄所", "澎湖": "高雄所",
}
CPA_ALIAS = {"(二余)清淵": "涂清淵", "?清淵": "涂清淵", "邱婉茹": "邱琬茹"}  # MOPS 申報之罕用字亂碼／異體字
FIN_KW = ["銀行", "金控", "證券", "保險", "人壽", "產險", "票券", "投信", "期貨", "信託", "銀", "金融"]
FIN_RE = re.compile(r"(金|證|銀|壽|保)$")  # 簡稱以金/證/銀/壽/保結尾（台新新光金、宏遠證…）
REASON_MAP = {  # 權數表分項 → 簽呈用語
    "營業收入、營業利益或稅前純益變動較大": "營運情形變動較大",
    "採用權益法之關聯企業及合資損益份額": "採用權益法之投資損失重大",
    "關係企業之交易": "關係人交易比重較高",
    "資金貸於他人": "資金貸與他人金額較高",
    "背書保證": "背書保證金額較高",
    "財務比率不佳": "財務比率不佳",
    "非流動之股權投資占權益比重較高": "非流動之股權投資占權益比重較高",
    "當期非流動之股權投資增減金額占權益比重較高者": "當期股權投資變動較大",
    "每股淨值偏低": "每股淨值偏低",
    "財務主管辭職": "財務主管異動", "會計主管辭職": "會計主管異動", "內稽主管辭職": "內稽主管異動",
    "研發主管辭職": "研發主管異動", "更換會計師": "更換會計師", "董監事持股異動": "董監事持股異動",
    "董事長辭職": "董事長異動", "總經理辭職": "總經理異動", "獨立董事辭職": "獨立董事異動",
    "董事辭職": "董事異動", "監察人辭職": "監察人異動", "董監事報酬不合理": "董監事報酬不合理",
    "董事、監察人質權設定": "董監事質押比率偏高", "董事長質權設定": "董事長質押比率偏高",
    "總經理質權設定": "總經理質押比率偏高", "最近一年內有訴訟事件": "重大訴訟",
    "財務主管或會計主管或內稽主管是否與董事有二等親": "財會主管與董事具二等親關係",
}


def reason_word(item_name):
    for k, v in REASON_MAP.items():
        if item_name.startswith(k) or k in item_name:
            return v
    return item_name[:12]


def city_of(addr):
    if not addr:
        return ""
    m = re.match(r"\s*(台北|臺北|新北|桃園|新竹|苗栗|台中|臺中|彰化|南投|雲林|嘉義|台南|臺南|高雄|屏東|宜蘭|花蓮|台東|基隆|金門|澎湖)", str(addr))
    return m.group(1).replace("臺", "台") if m else "境外/其他"


# ───────────────────────── 選案參數活頁簿 ─────────────────────────
def write_default_params(path):
    wb = openpyxl.Workbook()
    ws = wb.active; ws.title = "參數"
    ws.append(["項目", "值", "說明"])
    for k, v, d in DEFAULT_PARAMS: ws.append([k, v, d])
    ws2 = wb.create_sheet("前次檢查個案"); ws2.append(["檢查年度", "公司代號", "公司簡稱", "備註"])
    for r in DEFAULT_PRIOR: ws2.append(list(r))
    ws3 = wb.create_sheet("指定名單"); ws3.append(["公司代號", "公司簡稱", "處置(指定納入/指定排除)", "理由"])
    for r in DEFAULT_DESIGNATED: ws3.append(list(r))
    ws4 = wb.create_sheet("實質審閱名單"); ws4.append(["公司代號", "公司簡稱", "審閱單位(證交所/櫃買)", "備註"])
    ws5 = wb.create_sheet("內部訊號"); ws5.append(["公司代號", "公司簡稱", "訊號(檢舉/監視/前次缺失/其他)", "加分", "說明"])
    ws6 = wb.create_sheet("事務所檢查紀錄"); ws6.append(["事務所關鍵字", "前次檢查年度", "備註"])
    for r in DEFAULT_FIRM_HISTORY: ws6.append(list(r))
    ws7 = wb.create_sheet("會計師所別(覆寫)"); ws7.append(["會計師", "所別(台北所/新竹所/台中所/台南所/高雄所)", "備註"])
    for w in wb.worksheets:
        for c in w[1]:
            c.font = Font(name=FONT, bold=True, color="FFFFFF"); c.fill = PatternFill("solid", fgColor="305496")
        for col, wd in zip("ABCDE", (22, 18, 40, 30, 30)): w.column_dimensions[col].width = wd
    wb.save(path)


def read_params(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    P = {}
    for r in wb["參數"].iter_rows(min_row=2, values_only=True):
        if r[0]: P[str(r[0]).strip()] = r[1]
    def rows(name):
        return [r for r in wb[name].iter_rows(min_row=2, values_only=True) if r and r[0] not in (None, "")] if name in wb.sheetnames else []
    prior = {str(r[1]).strip(): (r[0], r[3]) for r in rows("前次檢查個案")}
    desig = {str(r[0]).strip(): (str(r[2] or ""), r[3] or "") for r in rows("指定名單")}
    subst = {str(r[0]).strip(): (r[2] or "", r[3] or "") for r in rows("實質審閱名單")}
    signals = collections.defaultdict(list)
    for r in rows("內部訊號"): signals[str(r[0]).strip()].append((r[2] or "", float(r[3] or 0), r[4] or ""))
    firm_hist = [(str(r[0]).strip(), r[1], r[2] or "") for r in rows("事務所檢查紀錄")]
    office_override = {str(r[0]).strip(): r[1] for r in rows("會計師所別(覆寫)") if r[1]}
    return P, prior, desig, subst, signals, firm_hist, office_override


# ───────────────────────── 讀取母體／權數表／確信 ─────────────────────────
def load_list(path):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = [w for w in wb.worksheets if "選案清單" in w.title][0]
    rows = list(ws.iter_rows(values_only=True)); hdr = list(rows[0])
    out, seen, dups = [], set(), set()
    for r in rows[1:]:
        if r[1] in (None, ""): continue
        code = str(r[1]).strip()
        if code in seen: dups.add(code); continue  # 證券期貨商雙申報（兩列）
        seen.add(code); out.append(dict(zip(hdr, r)))
    for d in out: d["_dup"] = d["公司代號"] in dups or str(d["公司代號"]).strip() in dups
    return out


def load_weights(path):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = [w for w in wb.worksheets if "合併" in w.title][0]
    rows = list(ws.iter_rows(values_only=True)); h5 = rows[HDR_ROWS - 1]
    W, bad = {}, 0
    num = lambda v: float(v) if isinstance(v, (int, float)) else 0.0
    for r in rows[HDR_ROWS:]:
        if r[0] in (None, ""): continue
        fs = sum(num(r[c]) for c in IDX_FIN_ITEMS); ns = sum(num(r[c]) for c in IDX_NF_ITEMS)
        if abs(fs - num(r[IDX_FIN_TOTAL])) > .01 or abs(ns - num(r[IDX_NF_TOTAL])) > .01 or abs(fs + ns - num(r[IDX_TOTAL])) > .01:
            bad += 1
        items = sorted([(num(r[c]), str(h5[c]).strip()) for c in list(IDX_FIN_ITEMS) + list(IDX_NF_ITEMS) if num(r[c]) > 0], key=lambda x: -x[0])
        W[str(r[0]).strip().lstrip("0")] = dict(mkt=r[IDX_MKT], fin=num(r[IDX_FIN_TOTAL]), nf=num(r[IDX_NF_TOTAL]), total=num(r[IDX_TOTAL]),
                                                items=items, breakdown="；".join(f"{n[:18]}({int(v)})" for v, n in items))
    if bad: sys.exit(f"權數表勾稽失敗 {bad} 筆（分項加總≠合計），請確認欄位索引 65/83/84 是否仍正確")
    return W


def load_assurance(path):
    A = {}
    if not os.path.exists(path): return A
    with open(path, encoding="utf-8-sig") as f:
        for r in csv.DictReader(f):
            code = r.get("公司代號", "").strip()
            A[code] = dict(sr_ey=(r.get("[永續報告書] 確信者為安永", "") == "是"), ghg_ey=(r.get("[溫室氣體] 確信者為安永", "") == "是"),
                           sr_assurer=r.get("[永續報告書] 確信出具者", ""), ghg_assurer=r.get("[溫室氣體] 確信報告出具者", ""),
                           overlap=r.get("【綜合】重疊態樣", ""))
    return A


def load_roster(here, firm):
    here = IN_DIR
    """會計師公會執業名冊：會計師姓名 → 所別（台北總所→台北所、桃園分所→桃園所…）。"""
    import glob
    files = sorted(glob.glob(os.path.join(here, "四大會計師事務所會計師名冊*.xlsx")))
    if not files: return {}, ""
    wb = openpyxl.load_workbook(files[-1], read_only=True, data_only=True)
    ws = next((w for w in wb.worksheets if firm in w.title), None) or wb["全部名單"]
    rows = list(ws.iter_rows(values_only=True)); h = [str(x).strip() if x else "" for x in rows[0]]
    i_n, i_o, i_f = h.index("會計師姓名"), h.index("所屬分所"), h.index("事務所名稱")
    R = {}
    for r in rows[1:]:
        if not r[i_n] or (firm not in str(r[i_f])): continue
        o = str(r[i_o] or "").strip().replace("總所", "所").replace("分所", "所")
        R[str(r[i_n]).strip()] = o or "未知"
    return R, os.path.basename(files[-1])


def norm_cpa(name):
    n = str(name or "").strip()
    return CPA_ALIAS.get(n, n)


# ───────────────────────── 規則引擎 ─────────────────────────
def infer_offices(pop, override, roster):
    """會計師 → 分所：優先順序 參數覆寫 > 公會名冊 > 以其簽證客戶所在縣市多數決推定。"""
    votes = collections.defaultdict(collections.Counter)
    for r in pop:
        reg = REGION_OFFICE.get(city_of(r.get("地址")))
        if not reg: continue
        for k in ("會計師A", "會計師B"):
            if r.get(k): votes[norm_cpa(r[k])][reg] += 1
    off = {}
    for cpa, c in votes.items():
        top, n = c.most_common(1)[0]
        if cpa in override: off[cpa] = (override[cpa], n, sum(c.values()), "覆寫")
        elif cpa in roster: off[cpa] = (roster[cpa], n, sum(c.values()), "公會名冊")
        else: off[cpa] = (top, n, sum(c.values()), f"推定({n}/{sum(c.values())})")
    for cpa, o in override.items():
        if cpa not in off: off[cpa] = (o, 0, 0, "覆寫")
    return off


def evaluate(pop, W, A, P, prior, desig, subst, signals, offices):
    tgt_year = int(P.get("检查年度", P.get("檢查年度", 115)))
    out = []
    for r in pop:
        code = str(r["公司代號"]).strip(); key = code.lstrip("0") or code
        w = W.get(key); a = A.get(key, {})
        d = dict(r)
        d["市場別"] = w["mkt"] if w else r["市場別"]
        d["財務權數"] = w["fin"] if w else None; d["非財務權數"] = w["nf"] if w else None; d["權數總計"] = w["total"] if w else None
        d["主要扣分項目(分數)"] = w["breakdown"] if w else (r.get("主要扣分項目(分數)") or "")
        cpaA = norm_cpa(r.get("會計師A")); cpaB = norm_cpa(r.get("會計師B"))
        oa = offices.get(cpaA, ("未知", 0, 0, ""))[0]; ob = offices.get(cpaB, ("未知", 0, 0, ""))[0]
        d["推定所別"] = oa if oa != "未知" else ob
        d["所別依據"] = f"A:{cpaA}→{oa}({offices.get(cpaA, ('', 0, 0, '無'))[3]})；B:{cpaB}→{ob}({offices.get(cpaB, ('', 0, 0, '無'))[3]})"
        d["公司所在地"] = city_of(r.get("地址"))
        d["永續確信(安永)"] = "是" if a.get("sr_ey") else ("否" if a else "")
        d["溫室氣體確信(安永)"] = "是" if a.get("ghg_ey") else ("否" if a else "")
        d["確信重疊態樣"] = a.get("overlap", "")
        # ── 加分項
        adj, why = 0.0, []
        rmk = str(r.get("備註") or ""); op = str(r.get("查核意見") or "")
        if "新承接" in rmk: adj += float(P["新承接加分"]); why.append(f"新承接+{int(P['新承接加分'])}")
        if "繼續經營" in op: adj += float(P["繼續經營加分"]); why.append(f"繼續經營+{int(P['繼續經營加分'])}")
        if re.match(r"^\s*(保留意見|否定意見|無法表示意見)", op): adj += float(P["保留意見加分"]); why.append(f"修正式意見+{int(P['保留意見加分'])}")
        if "採用其他會計師查核" in op: adj += float(P["其他事務所查核加分"]); why.append(f"依賴其他會計師+{int(P['其他事務所查核加分'])}")
        if a.get("sr_ey") or a.get("ghg_ey"): adj += float(P["確信重疊加分"]); why.append(f"確信重疊+{int(P['確信重疊加分'])}")
        for sig, pts, note in signals.get(code, []): adj += pts; why.append(f"{sig}+{int(pts)}")
        d["內部訊號加分"] = adj; d["加分說明"] = "；".join(why)
        d["調整後分數"] = (w["total"] + adj) if w else None
        # ── 排除項
        ex = []
        nm = str(r.get("公司簡稱") or "")
        if P.get("金融業排除") == "是" and (any(k in nm for k in FIN_KW) or FIN_RE.search(nm) or r.get("_dup")):
            ex.append("證券期貨商(雙申報)" if r.get("_dup") and not any(k in nm for k in FIN_KW) else "金融業(另有監理)")
        if P.get("前次已查排除") == "是" and code in prior:
            y = prior[code][0]
            if y is not None and int(y) >= tgt_year - 2: ex.append(f"{y}年已抽查")
            else: d["備註2"] = f"{y}年曾抽查"
        if P.get("實質審閱排除") == "是" and code in subst: ex.append(f"114年度財報經{subst[code][0] or '交易所'}實質審閱")
        if code in desig and desig[code][0].startswith("指定排除"): ex.append(f"指定排除：{desig[code][1]}")
        d["排除事由"] = "；".join(ex)
        d["指定納入"] = desig[code][1] if code in desig and desig[code][0].startswith("指定納入") else ""
        d["前次檢查"] = f"{prior[code][0]}年" if code in prior else ""
        out.append(d)
    return out


def pick(cands, P):
    N = int(P["目標家數"]); thr = float(P["分數門檻"]); cap_tpe = int(P["台北所上限"]); min_each = int(P["每所至少"])
    chosen, reasons = [], {}
    elig = [c for c in cands if not c["排除事由"]]
    # 1) 指定納入
    for c in elig:
        if c["指定納入"]: chosen.append(c); reasons[c["公司代號"]] = "指定納入：" + c["指定納入"]
    scored = sorted([c for c in elig if c["調整後分數"] is not None and c["調整後分數"] >= thr and c not in chosen], key=lambda c: (-c["調整後分數"], c["公司代號"]))
    def count(off): return sum(1 for c in chosen if c["推定所別"] == off)
    # 2) 各分所至少 min_each（非台北所）
    for off in sorted({c["推定所別"] for c in scored} - {"台北所", "未知", ""}):
        for c in [c for c in scored if c["推定所別"] == off]:
            if len(chosen) >= N or count(off) >= min_each: break
            if c not in chosen: chosen.append(c); reasons[c["公司代號"]] = f"地域分配：{off}達門檻分數最高者"
    # 3) 依分數補滿（台北所受上限限制）
    for c in scored:
        if len(chosen) >= N: break
        if c in chosen: continue
        if c["推定所別"] == "台北所" and count("台北所") >= cap_tpe: continue
        chosen.append(c); reasons[c["公司代號"]] = "風險排序"
    chosen.sort(key=lambda c: -(c["調整後分數"] or 0))
    backup = [c for c in scored if c not in chosen][:10]
    return chosen, backup, reasons


def memo_text(c, year_prev):
    A = str(c.get("會計師A") or ""); B = str(c.get("會計師B") or "")
    items = [n for n in str(c.get("主要扣分項目(分數)") or "").split("；") if n]
    words = []
    for it in items:
        wd = reason_word(re.sub(r"\(\d+\)$", "", it))
        if wd not in words: words.append(wd)
    extra = []
    rmk = str(c.get("備註") or ""); op = str(c.get("查核意見") or "")
    if "新承接" in rmk: extra.append("本年度新承接（審計準則第510號期初餘額）")
    if "繼續經營" in op: extra.append("查核報告載有繼續經營重大不確定性")
    base = "、".join(words[:3]) if words else "風險權數較高"
    t = f"{c['公司代號']} {c['公司簡稱']}【{c['市場別']}】：該公司由{A}及{B}會計師簽證，因{year_prev}年度{base}" + ("及" + "、".join(extra) if extra else "") + \
        f"等因素（權數總計{int(c['權數總計']) if c['權數總計'] is not None else '—'}分），爰擬選取該公司為本次查核個案。"
    if c.get("溫室氣體確信(安永)") == "是": t += f"另安永亦對該公司{year_prev}年度溫室氣體盤查資訊出具確信報告，擬一併調閱相關工作底稿進行瞭解。"
    elif c.get("永續確信(安永)") == "是": t += f"另安永亦對該公司{year_prev}年度永續報告書出具確信報告，擬一併調閱相關工作底稿進行瞭解。"
    return t


# ───────────────────────── 事務所選取規則 ─────────────────────────
def firm_schedule(all_list_path, P, firm_hist):
    wb = openpyxl.load_workbook(all_list_path, read_only=True, data_only=True)
    ws = [w for w in wb.worksheets if "選案清單" in w.title][0]
    rows = list(ws.iter_rows(values_only=True)); hdr = list(rows[0]); i_f = hdr.index("事務所"); i_m = hdr.index("市場別"); i_c = hdr.index("公司代號")
    seen = set(); cnt = collections.Counter(); bym = collections.defaultdict(collections.Counter)
    for r in rows[1:]:
        f = re.sub(r"\s+", "", str(r[i_f] or "")); code = str(r[i_c]).strip()
        if not f or (f, code) in seen: continue
        seen.add((f, code)); cnt[f] += 1; bym[f][r[i_m]] += 1
    t2, t6 = int(P["每2年門檻"]), int(P["每6年門檻"]); year = int(P["檢查年度"])
    out = []
    for f, n in cnt.most_common():
        cyc = 2 if n >= t2 else (6 if n >= t6 else None)
        last, note = None, ""
        for kw, y, nt in firm_hist:
            if kw and kw in f: last, note = y, nt; break
        due = (int(last) + cyc) if (cyc and last) else None
        status = ("本年度應檢查" if due is not None and due <= year else ("尚未到期" if due else ("請補前次檢查年度" if cyc else "視需要辦理")))
        out.append([f, n, bym[f].get("上市", 0), bym[f].get("上櫃", 0), bym[f].get("興櫃", 0), bym[f].get("公開發行", 0),
                    f"每{cyc}年" if cyc else "視需要", last, due, status, note])
    return out


# ───────────────────────── 輸出 ─────────────────────────
def style_header(ws, ncol):
    for c in range(1, ncol + 1):
        cell = ws.cell(1, c); cell.font = Font(name=FONT, bold=True, color="FFFFFF", size=10)
        cell.fill = PatternFill("solid", fgColor="305496"); cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.freeze_panes = "A2"


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--params", default=None); ap.add_argument("--out", default=None)
    ap.add_argument("--list", default=None); ap.add_argument("--all", default=None); ap.add_argument("--weights", default=None); ap.add_argument("--assurance", default=None)
    a = ap.parse_args()
    params_path = a.params or os.path.join(IN_DIR, "選案參數_115安永.xlsx")
    if not os.path.exists(params_path):
        write_default_params(params_path); print(f"已建立預設選案參數：{params_path}（可修改後重跑）")
    P, prior, desig, subst, signals, firm_hist, override = read_params(params_path)
    firm = str(P["受查事務所"]); q = str(P["資料季別"]); year = int(P["檢查年度"])
    list_path = a.list or os.path.join(IN_DIR, f"{q}_選案清單_{firm}.xlsx")
    all_path = a.all or os.path.join(IN_DIR, f"{q}_選案清單_全市場.xlsx")
    w_path = a.weights or os.path.join(IN_DIR, f"{q}權數表(本國上市上櫃公司)_SFB.xlsx")
    as_path = a.assurance or os.path.join(IN_DIR, f"{firm}{year-1}客戶_確信總表.csv")
    pop = load_list(list_path); W = load_weights(w_path); A = load_assurance(as_path)
    print(f"母體 {len(pop)} 家；權數表 {len(W)} 家；確信比對 {len(A)} 家")
    roster, roster_name = load_roster(HERE, firm)
    print(f"會計師名冊：{roster_name or '無'}（{len(roster)} 位）")
    offices = infer_offices(pop, override, roster)
    ev = evaluate(pop, W, A, P, prior, desig, subst, signals, offices)
    ev.sort(key=lambda d: (-(d["調整後分數"] if d["調整後分數"] is not None else -1), d["公司代號"]))
    for i, d in enumerate(ev, 1): d["排序"] = i
    chosen, backup, why = pick(ev, P)
    sched = firm_schedule(all_path, P, firm_hist) if os.path.exists(all_path) else []

    out = openpyxl.Workbook(); body = Font(name=FONT, size=10)
    # 0 說明
    ws0 = out.active; ws0.title = "說明"
    lines = [
        f"{year}年度事務所檢查（{firm}）— 選案工作表", "",
        "流程：選案表輸入（選案參數活頁簿）→ 事務所選取規則（每2年/每6年）→ 個案公司篩選規則 → 建議選案名單",
        f"母體：{q} auditor 快照篩「{firm}」{len(pop)} 家（含上市/上櫃/興櫃/公發，證券期貨商雙申報已去重）",
        f"分數：證交所／櫃買中心 {q} 平時管理權數統計表（固定索引 65/83/84，全表分項加總＝合計勾稽通過）；匹配 {sum(1 for d in ev if d['權數總計'] is not None)} 家",
        "", "個案公司篩選規則（依「選案參數」）：",
        f"  1. 排除：金融業（另有監理）／前次檢查已抽查（{year-2}年）／{year-1}年度財報經交易所實質審閱（名單由使用者填入）／指定排除",
        f"  2. 風險排序：調整後分數 = 權數總計 + 內部訊號加分（新承接+{P['新承接加分']}、繼續經營+{P['繼續經營加分']}、修正式意見+{P['保留意見加分']}、依賴其他會計師+{P['其他事務所查核加分']}、確信重疊+{P['確信重疊加分']}、內部訊號表）",
        f"  3. 地域分配：以簽證會計師所屬分所分配（來源：{roster_name or '無名冊'} 公會執業名冊；名冊查無者依其客戶所在地推定；可於參數表覆寫）；台北所上限 {P['台北所上限']} 家、其餘各所至少 {P['每所至少']} 家（須達分數門檻 {P['分數門檻']}）",
        f"  4. 指定納入：{', '.join(f'{k} {v[1]}' for k, v in desig.items() if v[0].startswith('指定納入')) or '無'}",
        f"  5. 目標家數 {P['目標家數']}；不足者依分數補足；備選名單另列 10 家",
        "", "各工作表：參數快照／事務所檢查週期／母體評分清單（每條規則有對應欄位）／建議選案／簽呈文字／會計師所別對照",
        "", f"注意：本表先列 {P['目標家數']} 家，待取得交易所 {year-1} 年度財報實質審閱名單填入參數表重跑後再篩至最終家數；所別以 A 會計師為準，A/B 不同所者於「所別依據」欄可見。",
        f"產製：{datetime.date.today().isoformat()}；腳本：select_cases.py",
    ]
    for i, t in enumerate(lines, 1): ws0.cell(i, 1, t).font = Font(name=FONT, size=11, bold=(i == 1))
    ws0.column_dimensions["A"].width = 120
    # 1 參數快照
    ws1 = out.create_sheet("參數快照"); ws1.append(["項目", "值"])
    for k, v in P.items(): ws1.append([k, v])
    style_header(ws1, 2); ws1.column_dimensions["A"].width = 22; ws1.column_dimensions["B"].width = 20
    # 2 事務所檢查週期
    ws2 = out.create_sheet("事務所檢查週期")
    ws2.append(["事務所", f"{q}簽證公開發行公司家數", "上市", "上櫃", "興櫃", "公開發行", "檢查週期", "前次檢查年度", "應檢查年度", f"{year}年度判定", "備註"])
    for r in sched: ws2.append(r)
    style_header(ws2, 11)
    for col, wd in zip("ABCDEFGHIJK", (30, 14, 8, 8, 8, 10, 10, 12, 12, 16, 30)): ws2.column_dimensions[col].width = wd
    # 3 母體評分清單
    cols = ["排序", "市場別", "公司代號", "公司簡稱", "推定所別", "公司所在地", "承辦人", "會計師A", "會計師B", "查核意見", "簽證日期", "每股盈餘(元)", "每股淨值(元)",
            "財務權數", "非財務權數", "權數總計", "內部訊號加分", "調整後分數", "加分說明", "排除事由", "指定納入", "前次檢查", "永續確信(安永)", "溫室氣體確信(安永)", "確信重疊態樣",
            "主要扣分項目(分數)", "備註", "所別依據", "董事長", "地址"]
    ws3 = out.create_sheet(f"{firm}母體評分清單"); ws3.append(cols)
    for d in ev: ws3.append([d.get(c) for c in cols])
    style_header(ws3, len(cols))
    widths = [6, 8, 9, 12, 9, 9, 10, 9, 9, 30, 11, 9, 9, 9, 9, 9, 9, 10, 26, 24, 18, 9, 9, 10, 16, 60, 24, 30, 12, 36]
    for i, wd in enumerate(widths, 1): ws3.column_dimensions[get_column_letter(i)].width = wd
    ws3.auto_filter.ref = f"A1:{get_column_letter(len(cols))}{len(ev)+1}"
    red = PatternFill("solid", fgColor="FCE4D6"); green = PatternFill("solid", fgColor="E2EFDA"); yellow = PatternFill("solid", fgColor="FFF2CC")
    chosen_codes = {c["公司代號"] for c in chosen}
    for i, d in enumerate(ev, 2):
        for c in range(1, len(cols) + 1): ws3.cell(i, c).font = body
        if d["排除事由"]:
            for c in range(1, 5): ws3.cell(i, c).fill = red
        if d["公司代號"] in chosen_codes:
            for c in range(1, 5): ws3.cell(i, c).fill = green
        if d["指定納入"]: ws3.cell(i, 21).fill = yellow
    # 4 建議選案
    ws4 = out.create_sheet("建議選案")
    hc = ["序", "區分", "公司代號", "公司簡稱", "市場別", "推定所別", "公司所在地", "承辦人", "會計師A", "會計師B", "權數總計", "調整後分數", "加分說明", "選入依據", "主要扣分項目(分數)", "確信重疊態樣", "前次檢查", "備註"]
    ws4.append(hc)
    for i, c in enumerate(chosen, 1):
        ws4.append([i, "建議", c["公司代號"], c["公司簡稱"], c["市場別"], c["推定所別"], c["公司所在地"], c.get("承辦人"), c.get("會計師A"), c.get("會計師B"), c["權數總計"], c["調整後分數"], c["加分說明"], why.get(c["公司代號"], ""), c["主要扣分項目(分數)"], c["確信重疊態樣"], c["前次檢查"], c.get("備註")])
    for i, c in enumerate(backup, 1):
        ws4.append([i, "備選", c["公司代號"], c["公司簡稱"], c["市場別"], c["推定所別"], c["公司所在地"], c.get("承辦人"), c.get("會計師A"), c.get("會計師B"), c["權數總計"], c["調整後分數"], c["加分說明"], "備選（依分數）", c["主要扣分項目(分數)"], c["確信重疊態樣"], c["前次檢查"], c.get("備註")])
    style_header(ws4, len(hc))
    for i, wd in enumerate((4, 6, 9, 12, 8, 9, 9, 10, 9, 9, 9, 10, 26, 26, 60, 16, 9, 24), 1): ws4.column_dimensions[get_column_letter(i)].width = wd
    for row in ws4.iter_rows(min_row=2):
        for cell in row: cell.font = body
    r0 = len(chosen) + len(backup) + 3
    ws4.cell(r0, 1, "分所分布（建議名單）").font = Font(name=FONT, bold=True)
    dist = collections.Counter(c["推定所別"] for c in chosen)
    for j, (k, v) in enumerate(sorted(dist.items()), 1): ws4.cell(r0 + j, 1, k); ws4.cell(r0 + j, 2, v)
    # 5 簽呈文字
    ws5 = out.create_sheet("簽呈文字"); ws5.column_dimensions["A"].width = 130
    ws5.cell(1, 1, f"（一）審計個案之抽核：本次檢查擬抽核審查{year-1}年度{len(chosen)}個檢查個案。經檢視{firm}簽證之公開發行公司明細，{year-1}年度計簽證{len(pop)}家公開發行公司"
                   f"【其中{sum(1 for d in ev if d['市場別']=='上市')}家上市、{sum(1 for d in ev if d['市場別']=='上櫃')}家上櫃、{sum(1 for d in ev if d['市場別']=='興櫃')}家興櫃及{sum(1 for d in ev if d['市場別']=='公開發行')}家公開發行公司】；"
                   f"經以證交所及櫃買中心{q}平時管理權數統計表之風險權數（營運情形變動、財務比率不佳、資金貸與及背書保證、關係人交易、權益法投資損失、治理警訊等）為基礎，"
                   f"綜合考量新承接案件、查核意見、事務所同時出具永續確信之情形，並兼顧各分所（台北、新竹、台中、台南、高雄）之分布，"
                   f"排除{year-1}年度財務報告業經證交所或櫃買中心實質審閱、金融業及本會前次檢查（{year-2}年度）已抽查之個案，本次計抽選{len(chosen)}家個案公司，謹說明如下：").alignment = Alignment(wrap_text=True, vertical="top")
    for i, c in enumerate(chosen, 1):
        ws5.cell(i + 1, 1, f"{i}、" + memo_text(c, year - 1)).alignment = Alignment(wrap_text=True, vertical="top")
    for row in ws5.iter_rows():
        for cell in row: cell.font = Font(name=FONT, size=11)
    # 6 會計師所別
    ws6 = out.create_sheet("會計師所別對照"); ws6.append(["會計師", "推定所別", "多數縣市家數", "簽證家數(母體)", "依據"])
    for cpa, (o, n, tot, src) in sorted(offices.items(), key=lambda x: -x[1][2]): ws6.append([cpa, o, n, tot, src])
    style_header(ws6, 5)
    for col, wd in zip("ABCDE", (12, 10, 12, 14, 14)): ws6.column_dimensions[col].width = wd

    os.makedirs(OUT_DIR, exist_ok=True)
    fn = a.out or os.path.join(OUT_DIR, f"{year}年度_{firm}_選案工作表.xlsx")
    out.save(fn)
    print(f"完成：{fn}")
    print("建議名單：")
    for c in chosen: print(f"  {c['公司代號']} {c['公司簡稱']:8s} {c['市場別']} {c['推定所別']} 總計{c['權數總計']} 調整{c['調整後分數']} | {why.get(c['公司代號'])} | {c['加分說明']}")
    print("分所分布：", dict(collections.Counter(c["推定所別"] for c in chosen)))
    print("備選：", [(c["公司代號"], c["公司簡稱"], c["推定所別"], c["調整後分數"]) for c in backup])


if __name__ == "__main__":
    if hasattr(sys.stdout, "reconfigure"): sys.stdout.reconfigure(encoding="utf-8")
    main()
