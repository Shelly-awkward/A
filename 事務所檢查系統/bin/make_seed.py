# -*- coding: utf-8 -*-
"""產生 apps_script/Seed.gs（題庫＋分類），供 Apps Script setup() 建表用。每年題庫改版時重跑一次。"""
import json, os, sys
HERE = os.path.dirname(os.path.abspath(__file__)); ROOT = os.path.dirname(HERE); sys.path.insert(0, HERE)
from items_chk2_115 import ITEMS
from render_case_docs import chk1_items, aml_items, CASE_CATS, QM_CATS, TPL
chk2, sec = [], ""
for k, t in ITEMS:
    if k == "S": sec = t; continue
    chk2.append({"q": k, "sec": sec, "text": t})
chk1 = [{"row": r, "sub": k, "sec": s, "text": t} for r, k, s, t in chk1_items(os.path.join(TPL, "檢查表一_115.docx"))]
aml = [{"t": ti, "r": ri, "kind": kind, "text": txt} for ti, ri, kind, txt in aml_items(os.path.join(TPL, "AML評量表_115_範本.docx"))]
seed = dict(version="115", chk2=chk2, chk1=chk1, aml=aml, caseCats=CASE_CATS, qmCats=QM_CATS)
out = os.path.join(ROOT, "apps_script", "Seed.gs")
open(out, "w", encoding="utf-8").write("// 自動產生：python bin/make_seed.py（題庫版本 %s）\nvar SEED = %s;\n" % ("115", json.dumps(seed, ensure_ascii=False)))
print("寫入", out, len(chk2), len(chk1), len(aml))
