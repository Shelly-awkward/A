# -*- coding: utf-8 -*-
"""
render_case_docs.py — 個案檢查表「輸入 → 產出」雛型（v0.1）
==============================================================
  python render_case_docs.py init   [--selection 115年度_安永_選案工作表.xlsx] [--year 115] [--firm 安永聯合會計師事務所]
      → 產生 個案檢查_輸入_<年度><事務所>.xlsx（個案主檔／檢查表二作答／檢查表一作答／AML作答／缺失／事務所）
  python render_case_docs.py render --input 個案檢查_輸入_115安永.xlsx [--out out]
      → 檢查表二_<代號><簡稱>.docx（每案）、缺失三欄式對照表_<代號><簡稱>.docx（有缺失之案）、
        檢查表一_<事務所>.docx、AML評量表_<事務所>.docx、缺失分類彙總表_<年度><事務所>.xlsx、缺失申報站匯入.csv
範本（同資料夾 templates/）：檢查表一_115.docx、AML評量表_115_範本.docx、彙總表_空白範本.xlsx
檢查表二與三欄式對照表因原檔為 .doc，改由程式依 115 年版逐列重建（題庫 items_chk2_115.py）。
"""
import argparse, copy, csv, datetime, os, re, sys
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
TPL = os.path.join(ROOT, "templates", os.environ.get("CHK_TPL_YEAR", "115"))
sys.path.insert(0, HERE)
from items_chk2_115 import ITEMS, TITLE_PREFIX, VERSION, TITLE, INTRO4, HEAD, NOTES

CJK = "標楷體"
CASE_CATS = ["會計估計/公允價值衡量", "證實性分析程序", "關鍵查核事項", "底稿記載缺失", "重大性", "內控測試", "財報表達與揭露", "存貨查核", "風險評估",
             "管理階層專家工作之採用", "個案複核及管理", "資金貸與及背書保證", "收入認列", "集團查核", "舞弊查核", "關係人交易", "審計抽樣", "其他"]
QM_CATS = ["一、事務所之風險評估流程", "二、治理及領導階層", "三、攸關職業道德規範(包括獨立性)", "四、客戶關係及案件之承接與續任", "五、案件之執行", "六、資源", "七、資訊及溝通", "八、監督及改正流程"]
OPINIONS = ["V", "✗", "NA"]

# ───────────────────────── docx 小工具 ─────────────────────────
def set_run_font(run, size=12, bold=None, name=CJK):
    run.font.name = name; run.font.size = Pt(size)
    if bold is not None: run.font.bold = bold
    rpr = run._r.get_or_add_rPr(); rf = rpr.find(qn("w:rFonts"))
    if rf is None: rf = OxmlElement("w:rFonts"); rpr.append(rf)
    for k in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"): rf.set(qn(k), name)

def add_par(doc_or_cell, text="", size=12, bold=False, align=None, space_after=4, first_line=None):
    p = doc_or_cell.add_paragraph()
    lines = text.split("\n")
    for i, line in enumerate(lines):
        r = p.add_run(line); set_run_font(r, size, bold)
        if i < len(lines) - 1: r.add_break()
    if align: p.alignment = align
    p.paragraph_format.space_after = Pt(space_after); p.paragraph_format.space_before = Pt(0)
    if first_line is not None: p.paragraph_format.first_line_indent = Cm(first_line)
    return p

def set_cell(cell, text, size=11, bold=False, align=None):
    cell.text = ""
    p = cell.paragraphs[0]
    lines = str(text if text is not None else "").split("\n")
    for i, line in enumerate(lines):
        r = p.add_run(line); set_run_font(r, size, bold)
        if i < len(lines) - 1: r.add_break()
    if align: p.alignment = align
    p.paragraph_format.space_after = Pt(0)

def set_col_widths(table, widths_cm):
    table.autofit = False
    grid = table._tbl.tblGrid
    for gc in grid.findall(qn("w:gridCol")): grid.remove(gc)
    for w in widths_cm:
        gc = OxmlElement("w:gridCol"); gc.set(qn("w:w"), str(int(w * 567))); grid.append(gc)
    for row in table.rows:
        for c, w in zip(row.cells, widths_cm): c.width = Cm(w)

def set_borders(table):
    tblPr = table._tbl.tblPr
    b = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        e = OxmlElement(f"w:{edge}"); e.set(qn("w:val"), "single"); e.set(qn("w:sz"), "6"); e.set(qn("w:color"), "000000"); b.append(e)
    tblPr.append(b)

def repeat_header(row):
    trPr = row._tr.get_or_add_trPr(); h = OxmlElement("w:tblHeader"); h.set(qn("w:val"), "true"); trPr.append(h)

def new_doc(landscape=False, margins_cm=(2.0, 2.0, 2.0, 2.0)):
    doc = Document(); sec = doc.sections[0]
    sec.page_width, sec.page_height = Cm(21.0), Cm(29.7)
    if landscape:
        sec.orientation = WD_ORIENT.LANDSCAPE; sec.page_width, sec.page_height = Cm(29.7), Cm(21.0)
    sec.top_margin, sec.bottom_margin, sec.left_margin, sec.right_margin = [Cm(m) for m in margins_cm]
    st = doc.styles["Normal"]; st.font.name = CJK; st.font.size = Pt(12)
    st.element.rPr.rFonts.set(qn("w:eastAsia"), CJK)
    return doc

# ───────────────────────── 檢查表二（重建） ─────────────────────────
def render_chk2(case, answers, year, out):
    doc = new_doc(False, (1.8, 1.8, 2.0, 1.8))
    add_par(doc, TITLE_PREFIX, 12); add_par(doc, VERSION, 12)
    add_par(doc, TITLE.format(year=year), 16, True, WD_ALIGN_PARAGRAPH.CENTER, 8)
    add_par(doc, f"一、公司代號： {case['公司代號']}　公司名稱：{case.get('公司全名') or case.get('公司簡稱') or ''}", 12)
    add_par(doc, f"二、簽證會計師事務所： {case.get('事務所') or ''}", 12)
    add_par(doc, f"　　簽證會計師： {case.get('會計師A') or ''}會計師、 {case.get('會計師B') or ''}會計師", 12)
    add_par(doc, f"　　工作底稿： {case.get('工作底稿冊數') or '　'}冊(含電子底稿 {case.get('電子底稿冊數') or '　'} 冊)", 12)
    add_par(doc, "三、調閱工作底稿之緣由：會計師事務所檢查", 12)
    add_par(doc, INTRO4, 12, space_after=6)
    t = doc.add_table(rows=1, cols=3); set_borders(t); t.alignment = WD_TABLE_ALIGNMENT.CENTER
    for c, h in zip(t.rows[0].cells, HEAD): set_cell(c, h, 11, True, WD_ALIGN_PARAGRAPH.CENTER)
    repeat_header(t.rows[0])
    for key, text in ITEMS:
        row = t.add_row()
        if key == "S":
            set_cell(row.cells[0], text, 11, True); row.cells[0].merge(row.cells[2])
        else:
            a = answers.get(key, {})
            set_cell(row.cells[0], text, 11)
            set_cell(row.cells[1], a.get("意見") or "", 11, align=WD_ALIGN_PARAGRAPH.CENTER)
            set_cell(row.cells[2], a.get("索引") or "", 11)
    row = t.add_row()
    set_cell(row.cells[0], f"覆核者：{case.get('覆核者') or ''}", 11); set_cell(row.cells[1], f"日期：{case.get('覆核日期') or ''}", 11); row.cells[1].merge(row.cells[2])
    set_col_widths(t, [10.6, 2.2, 4.4])
    for i, n in enumerate(NOTES): add_par(doc, n, 11, i == 0, space_after=2)
    doc.save(out)

# ───────────────────────── 缺失三欄式對照表（重建） ─────────────────────────
def render_3col(case, findings, year, out):
    doc = new_doc(True, (1.8, 1.8, 2.0, 2.0))
    name = case.get("公司全名") or case.get("公司簡稱") or ""
    add_par(doc, f"有關會計師簽證{name}", 14, True, WD_ALIGN_PARAGRAPH.CENTER, 2)
    add_par(doc, f"{case.get('事務所') or ''} {case.get('會計師A') or ''}、{case.get('會計師B') or ''}會計師 查核簽證", 14, True, WD_ALIGN_PARAGRAPH.CENTER, 2)
    add_par(doc, f"{name}{year}年度財務報告疏失彙總表", 14, True, WD_ALIGN_PARAGRAPH.CENTER, 8)
    add_par(doc, f"設立日期： {case.get('設立日期') or ''}", 12)
    add_par(doc, f"{'上市' if case.get('市場別') == '上市' else '上櫃' if case.get('市場別') == '上櫃' else '上市(櫃)'}日期： {case.get('上市櫃日期') or ''}", 12)
    add_par(doc, f"{case.get('公司簡稱') or name}公司主要營業項目為{case.get('主要營業項目') or ''}。", 12, space_after=8)
    t = doc.add_table(rows=1, cols=3); set_borders(t); t.alignment = WD_TABLE_ALIGNMENT.CENTER
    for c, h in zip(t.rows[0].cells, ["會計師所涉\n查核疏失", "會計師說明", "分析意見"]): set_cell(c, h, 12, True, WD_ALIGN_PARAGRAPH.CENTER)
    repeat_header(t.rows[0])
    for i, f in enumerate(findings, 1):
        row = t.add_row()
        set_cell(row.cells[0], f"{i}.{f.get('會計師所涉查核疏失') or ''}", 12)
        set_cell(row.cells[1], f.get("會計師說明") or "", 12)
        set_cell(row.cells[2], f.get("分析意見") or "", 12)
    set_col_widths(t, [8.4, 8.8, 8.5])
    doc.save(out)

# ───────────────────────── 檢查表一（套原檔） ─────────────────────────
def chk1_items(tpl_path):
    """回傳 [(列序, 子項, 章節, 檢查項目)]：每個 □ 一列，供作答表使用。"""
    doc = Document(tpl_path); t = doc.tables[0]; items, sec = [], ""
    for ri, row in enumerate(t.rows):
        c0 = row.cells[0].text.strip()
        boxes = [p for p in row.cells[1].paragraphs if p.text.strip() == "□"]
        if not boxes:
            if ri >= 2 and c0: sec = c0
            continue
        lines = [l.strip() for l in row.cells[0].text.split("\n") if l.strip()]
        if len(boxes) == 1: items.append((ri, 1, sec, "\n".join(lines)))
        else:
            subs = [l for l in lines if re.match(r"^\(\d+\)|^（\d+）", l)]
            head = lines[0] if lines else ""
            for k in range(len(boxes)):
                items.append((ri, k + 1, sec, (head + "\n" if k == 0 else "") + (subs[k] if k < len(subs) else f"(子項{k+1})")))
    return items

def render_chk1(tpl_path, firm_info, answers, out):
    doc = Document(tpl_path)
    def rewrite(p, text):  # 保留第一個 run 的格式，整段換字
        if not p.runs: p.add_run(text); return
        p.runs[0].text = text
        for r in p.runs[1:]: r.text = ""
    for p in doc.paragraphs:
        t = p.text
        if "○○聯合會計師事務所" in t: rewrite(p, t.replace("○○聯合會計師事務所", firm_info.get("事務所名稱") or "○○聯合會計師事務所"))
        elif t.strip().startswith("檢 查 者："): rewrite(p, "檢 查 者：" + str(firm_info.get("檢查者") or ""))
        elif t.strip().startswith("日   期 ：") or t.strip().startswith("日　　期："): rewrite(p, "日   期 ：" + str(firm_info.get("檢查日期") or ""))
    t = doc.tables[0]
    for (ri, k), a in answers.items():
        row = t.rows[ri]
        col = 1 if a.get("是") else 2 if a.get("否") else 3 if a.get("不適用") else None
        if col:
            boxes = [p for p in row.cells[col].paragraphs if p.text.strip() == "□"]
            if k - 1 < len(boxes):
                p = boxes[k - 1]
                for r in p.runs: r.text = ""
                (p.runs[0] if p.runs else p.add_run()).text = "■"
        if a.get("說明及發現"):
            cell = row.cells[4]; old = cell.text.strip()
            set_cell(cell, (old + "\n" if old else "") + (f"({k})" if k > 1 or a.get("_multi") else "") + a["說明及發現"], 10)
        if a.get("附件"):
            cell = row.cells[5]; old = cell.text.strip(); set_cell(cell, (old + "\n" if old else "") + a["附件"], 10)
    doc.save(out)

# ───────────────────────── AML 評量表（套原檔） ─────────────────────────
def aml_items(tpl_path):
    doc = Document(tpl_path); items = []
    for ti, t in enumerate(doc.tables):
        hdr = [c.text.strip() for c in t.rows[0].cells]
        if len(t.rows) > 1 and any("正常" in c.text for c in t.rows[1].cells):
            for ri in range(2, len(t.rows)):
                txt = t.rows[ri].cells[0].text.strip()
                if txt and txt != hdr[0]: items.append((ti, ri, "項目", txt))
        elif hdr and hdr[0] in ("綜合結論:", "評估意見:"):
            items.append((ti, 1, hdr[0].rstrip(":"), t.rows[1].cells[0].text.strip()))
    return items

def render_aml(tpl_path, firm_info, answers, out):
    doc = Document(tpl_path)
    t0 = doc.tables[0]
    for row in t0.rows:
        c = row.cells[0]
        if c.text.strip().startswith("事務所名稱") and firm_info.get("事務所名稱"): set_cell(c, "事務所名稱: " + firm_info["事務所名稱"], 11)
    for (ti, ri), a in answers.items():
        t = doc.tables[ti]; row = t.rows[ri]
        if a.get("kind") == "項目":
            ncol = len(row.cells)
            if ncol >= 4 and row.cells[1]._tc is not row.cells[2]._tc:
                if a.get("正常"): set_cell(row.cells[1], "V", 11, align=WD_ALIGN_PARAGRAPH.CENTER)
                if a.get("異常"): set_cell(row.cells[2], "V", 11, align=WD_ALIGN_PARAGRAPH.CENTER)
                if a.get("說明"): set_cell(row.cells[3], a["說明"], 10)
            else:
                mark = "■正常　□異常" if a.get("正常") else "□正常　■異常" if a.get("異常") else ""
                if mark: set_cell(row.cells[1], mark, 10, align=WD_ALIGN_PARAGRAPH.CENTER)
                if a.get("說明"): set_cell(row.cells[-1], a["說明"], 10)
        elif a.get("說明"):
            set_cell(row.cells[0], a["說明"], 11)
    doc.save(out)

# ───────────────────────── 缺失分類彙總表（套原檔） ─────────────────────────
def render_summary(tpl_path, firm_info, findings, n_cases, year, out):
    wb = openpyxl.load_workbook(tpl_path)
    firm = firm_info.get("事務所名稱") or ""
    for sheet, cats, kind in (("品質缺失", QM_CATS, "品質"), ("個案缺失", CASE_CATS, "個案")):
        ws = wb[sheet]
        rows = [f for f in findings if f.get("類型") == kind]
        ws.cell(1, 2 if ws.cell(1, 2).value else 3).value = f"{year}年度會計師事務所檢查{'品質管理' if kind == '品質' else '個案'}缺失彙總"
        ws.cell(2, 1).value = (f"檢查事務所名稱: {firm}\n檢查發現缺失數:共 {len(rows)} 項" if kind == "品質"
                               else f"檢查事務所名稱: {firm}\n檢查個案數: 共抽檢{n_cases} 家上市(櫃)公司\n檢查發現缺失數:共 {len(rows)} 項")
        # 樣式來源：第 5 列（分類列）與 6 列（明細列）
        sty = {c: copy.copy(ws.cell(5, c)._style) for c in range(1, 5)}
        total_row = next(r for r in range(5, ws.max_row + 1) if str(ws.cell(r, 1).value or "").startswith("合計"))
        sty_total = {c: copy.copy(ws.cell(total_row, c)._style) for c in range(1, 5)}
        ws.delete_rows(5, ws.max_row - 4)
        r = 5; total = 0
        for cat in cats:
            fs = [f for f in rows if f.get("缺失分類") == cat]
            if not fs:
                ws.cell(r, 1, cat); ws.cell(r, 2, 0)
                for c in range(1, 5): ws.cell(r, c)._style = copy.copy(sty[c])
                ws.row_dimensions[r].height = 19.5; r += 1; continue
            for i, f in enumerate(fs, 1):
                ws.cell(r, 1, cat if i == 1 else None); ws.cell(r, 2, total + i)
                content = f.get("會計師所涉查核疏失") or ""
                if kind == "個案" and f.get("公司簡稱"): content = f"【{f['公司簡稱']}】" + content
                ws.cell(r, 3, content); ws.cell(r, 4, f.get("違反或相關之準則或法規") or "")
                for c in range(1, 5): ws.cell(r, c)._style = copy.copy(sty[c])
                ws.row_dimensions[r].height = max(19.5, 15 * (len(content) // 28 + 1)); r += 1
            total += len(fs)
        ws.cell(r, 1, "合計缺失數"); ws.cell(r, 2, total)
        for c in range(1, 5): ws.cell(r, c)._style = copy.copy(sty_total[c])
        others = [f for f in rows if f.get("缺失分類") not in cats]
        if others: ws.cell(r + 2, 1, "※下列缺失之分類不在表列類別，請確認：" + "；".join(f"{f.get('缺失分類')}({f.get('公司簡稱') or ''})" for f in others))
    wb.save(out)

# ───────────────────────── 輸入活頁簿 ─────────────────────────
HDR = {
    "個案主檔": ["公司代號", "公司全名", "公司簡稱", "市場別", "事務所", "會計師A", "會計師B", "工作底稿冊數", "電子底稿冊數", "設立日期", "上市櫃日期", "主要營業項目", "主查", "覆核者", "覆核日期"],
    "檢查表二作答": ["公司代號", "公司簡稱", "題號", "章節", "審閱內容(摘)", "本局審查意見(V/✗/NA)", "會計師工作底稿索引", "備註"],
    "檢查表一作答": ["列序", "子項", "章節", "檢查項目", "是", "否", "不適用", "說明及發現", "附件"],
    "AML作答": ["表序", "列序", "類型", "檢查項目", "正常", "異常", "說明"],
    "缺失": ["類型(個案/品質)", "公司代號", "公司簡稱", "缺失分類", "項次", "會計師所涉查核疏失", "會計師說明", "分析意見", "違反或相關之準則或法規", "對應題號", "IFIAR主題(選填)"],
    "事務所": ["項目", "值"],
}
def style_hdr(ws, n):
    for c in range(1, n + 1):
        cell = ws.cell(1, c); cell.font = Font(name="Microsoft JhengHei", bold=True, color="FFFFFF", size=10)
        cell.fill = PatternFill("solid", fgColor="305496"); cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.freeze_panes = "A2"

def cmd_init(a):
    year = int(a.year); firm = a.firm
    cases = []
    if a.selection and os.path.exists(a.selection):
        wb = openpyxl.load_workbook(a.selection, read_only=True, data_only=True); ws = wb["建議選案"]
        rows = list(ws.iter_rows(values_only=True)); h = list(rows[0])
        for r in rows[1:]:
            if not r or r[1] != "建議": continue
            d = dict(zip(h, r))
            cases.append([str(d["公司代號"]), "", d["公司簡稱"], d["市場別"], firm, d.get("會計師A"), d.get("會計師B"), None, None, None, None, None, d.get("承辦人"), None, None])
    out = openpyxl.Workbook(); ws0 = out.active; ws0.title = "說明"
    for i, t in enumerate([
        f"{year}年度事務所檢查（{firm}）— 個案檢查表輸入活頁簿", "",
        "1. 個案主檔：每案一列（已自選案工作表帶入代號/簡稱/會計師；公司全名、底稿冊數、設立/上市櫃日期、主要營業項目請補）。",
        "2. 檢查表二作答：每案 31 題已展開，只填「本局審查意見」(V／✗／NA) 與「會計師工作底稿索引」。填 ✗ 者請到「缺失」表寫一筆並填對應題號。",
        "3. 檢查表一作答：品質管理制度檢查表，每個 □ 一列，於「是／否／不適用」擇一填 V，說明寫在「說明及發現」。",
        "4. AML作答：115年度AML現地檢查評量表逐項，正常／異常擇一填 V；「綜合結論」「評估意見」列直接寫在說明欄。",
        "5. 缺失：個案缺失（類型=個案，填公司代號）與品質缺失（類型=品質）皆寫在這裡；缺失分類請用下拉選單（與彙總表一致）。",
        "6. 事務所：事務所名稱、檢查者、檢查日期。",
        "填完執行：python render_case_docs.py render --input <本檔>", ""], 1):
        ws0.cell(i, 1, t).font = Font(name="Microsoft JhengHei", size=11, bold=(i == 1))
    ws0.column_dimensions["A"].width = 120
    ws = out.create_sheet("個案主檔"); ws.append(HDR["個案主檔"]); [ws.append(c) for c in cases]; style_hdr(ws, len(HDR["個案主檔"]))
    ws = out.create_sheet("檢查表二作答"); ws.append(HDR["檢查表二作答"])
    for c in cases:
        sec = ""
        for key, text in ITEMS:
            if key == "S": sec = text; continue
            ws.append([c[0], c[2], key, sec, text.split("\n")[0][:40], None, None, None])
    style_hdr(ws, 8)
    from openpyxl.worksheet.datavalidation import DataValidation
    dv = DataValidation(type="list", formula1='"V,✗,NA"', allow_blank=True); ws.add_data_validation(dv); dv.add(f"F2:F{max(ws.max_row, 2)}")
    ws = out.create_sheet("檢查表一作答"); ws.append(HDR["檢查表一作答"])
    for ri, k, sec, txt in chk1_items(os.path.join(TPL, "檢查表一_115.docx")): ws.append([ri, k, sec, txt, None, None, None, None, None])
    style_hdr(ws, 9)
    ws = out.create_sheet("AML作答"); ws.append(HDR["AML作答"])
    for ti, ri, kind, txt in aml_items(os.path.join(TPL, "AML評量表_115_範本.docx")): ws.append([ti, ri, kind, txt, None, None, None])
    style_hdr(ws, 7)
    ws = out.create_sheet("缺失"); ws.append(HDR["缺失"]); style_hdr(ws, 11)
    lst = out.create_sheet("選單"); lst.append(["個案缺失分類", "品質缺失分類"])
    for i in range(max(len(CASE_CATS), len(QM_CATS))): lst.append([CASE_CATS[i] if i < len(CASE_CATS) else None, QM_CATS[i] if i < len(QM_CATS) else None])
    dv1 = DataValidation(type="list", formula1='"個案,品質"', allow_blank=True); ws.add_data_validation(dv1); dv1.add("A2:A500")
    dv2 = DataValidation(type="list", formula1=f"=選單!$A$2:$B${len(CASE_CATS)+1}", allow_blank=True); ws.add_data_validation(dv2); dv2.add("D2:D500")
    ws = out.create_sheet("事務所"); ws.append(HDR["事務所"])
    for k, v in (("事務所名稱", firm), ("檢查年度", year), ("財報年度", year - 1), ("檢查者", None), ("檢查日期", None)): ws.append([k, v])
    style_hdr(ws, 2)
    for name, widths in (("個案主檔", [9, 26, 10, 8, 20, 9, 9, 10, 10, 12, 12, 40, 8, 8, 14]), ("檢查表二作答", [9, 9, 6, 24, 44, 14, 24, 20]),
                         ("檢查表一作答", [6, 5, 22, 70, 5, 5, 7, 40, 10]), ("AML作答", [5, 5, 8, 60, 6, 6, 40]), ("缺失", [10, 9, 9, 22, 5, 50, 40, 40, 28, 8, 20]), ("事務所", [14, 30])):
        for i, w in enumerate(widths, 1): out[name].column_dimensions[get_column_letter(i)].width = w
        for row in out[name].iter_rows(min_row=2):
            for cell in row: cell.alignment = Alignment(wrap_text=True, vertical="top")
    fn = a.output or os.path.join(ROOT, "input", f"個案檢查_輸入_{year}{firm[:2]}.xlsx"); out.save(fn); print("已建立輸入活頁簿：", fn, f"（{len(cases)} 案）")

def read_input(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    def rows(name):
        ws = wb[name]; data = list(ws.iter_rows(values_only=True)); h = [str(x).strip() if x else "" for x in data[0]]
        return [dict(zip(h, r)) for r in data[1:] if r and any(v not in (None, "") for v in r)]
    cases = {str(c["公司代號"]).strip(): c for c in rows("個案主檔") if c.get("公司代號")}
    a2 = {}
    for r in rows("檢查表二作答"):
        a2.setdefault(str(r["公司代號"]).strip(), {})[r["題號"]] = {"意見": r.get("本局審查意見(V/✗/NA)"), "索引": r.get("會計師工作底稿索引")}
    a1 = {}
    for r in rows("檢查表一作答"):
        if any(r.get(k) for k in ("是", "否", "不適用", "說明及發現", "附件")):
            a1[(int(r["列序"]), int(r["子項"]))] = {k: r.get(k) for k in ("是", "否", "不適用", "說明及發現", "附件")}
    aml = {}
    for r in rows("AML作答"):
        if any(r.get(k) for k in ("正常", "異常", "說明")):
            aml[(int(r["表序"]), int(r["列序"]))] = {"kind": r.get("類型"), "正常": r.get("正常"), "異常": r.get("異常"), "說明": r.get("說明")}
    findings = [dict(f, 類型=str(f.get("類型(個案/品質)") or "").strip()) for f in rows("缺失")]
    firm = {str(r["項目"]).strip(): r["值"] for r in rows("事務所") if r.get("項目")}
    return cases, a2, a1, aml, findings, firm

def cmd_render(a):
    cases, a2, a1, aml, findings, firm = read_input(a.input)
    year = int(firm.get("檢查年度") or 115); fy = int(firm.get("財報年度") or year - 1)
    out = a.out or os.path.join(ROOT, "output", "檢查表產出"); os.makedirs(out, exist_ok=True)
    fname = str(firm.get("事務所名稱") or "事務所")
    for code, c in cases.items():
        tag = f"{code}{c.get('公司簡稱') or ''}"
        render_chk2(c, a2.get(code, {}), fy, os.path.join(out, f"檢查表二_{tag}.docx"))
        fs = [f for f in findings if f["類型"] == "個案" and str(f.get("公司代號") or "").strip() == code]
        if fs: render_3col(c, fs, fy, os.path.join(out, f"缺失三欄式對照表_{tag}.docx"))
        print("  個案", tag, f"檢查表二 {sum(1 for v in a2.get(code, {}).values() if v.get('意見'))}/31 題已填；缺失 {len(fs)} 筆")
    render_chk1(os.path.join(TPL, "檢查表一_115.docx"), firm, a1, os.path.join(out, f"檢查表一_{fname}.docx"))
    render_aml(os.path.join(TPL, "AML評量表_115_範本.docx"), firm, aml, os.path.join(out, f"AML評量表_{fname}.docx"))
    for f in findings:
        if f["類型"] == "個案" and not f.get("公司簡稱"): f["公司簡稱"] = cases.get(str(f.get("公司代號") or "").strip(), {}).get("公司簡稱")
    render_summary(os.path.join(TPL, "彙總表_空白範本.xlsx"), firm, findings, len(cases), year, os.path.join(out, f"缺失分類彙總表_{year}{fname}.xlsx"))
    with open(os.path.join(out, "缺失申報站匯入.csv"), "w", encoding="utf-8-sig", newline="") as fh:
        w = csv.writer(fh); w.writerow(["year", "type", "firm", "companyCode", "company", "content", "localCategory", "themeHint"])
        for f in findings:
            w.writerow([year + 1911, "審計個案缺失" if f["類型"] == "個案" else "品質管理缺失", fname, f.get("公司代號") or "", f.get("公司簡稱") or "",
                        f.get("會計師所涉查核疏失") or "", f.get("缺失分類") or "", f.get("IFIAR主題(選填)") or ""])
    print("完成，輸出於", out, "：", sorted(os.listdir(out)))

if __name__ == "__main__":
    if hasattr(sys.stdout, "reconfigure"): sys.stdout.reconfigure(encoding="utf-8")
    ap = argparse.ArgumentParser(); sp = ap.add_subparsers(dest="cmd", required=True)
    p1 = sp.add_parser("init"); p1.add_argument("--selection", default=os.path.join(ROOT, "output", "115年度_安永_選案工作表.xlsx")); p1.add_argument("--year", default=115)
    p1.add_argument("--firm", default="安永聯合會計師事務所"); p1.add_argument("--output", default=None); p1.set_defaults(fn=cmd_init)
    p2 = sp.add_parser("render"); p2.add_argument("--input", required=True); p2.add_argument("--out", default=None); p2.set_defaults(fn=cmd_render)
    a = ap.parse_args(); a.fn(a)
