# -*- coding: utf-8 -*-
"""
85 份「2025 年度溫室氣體排放確信報告」PDF 的出具者判讀結果。
（2026-08-06 判讀 77 份；2026-08-31 母體重抓後新增 8 份）

判讀方式：
  · 50 份可抽文字 → 抽取後比對抬頭／簽章段落
  · 35 份為純掃描影像 → 逐頁渲染後以視覺判讀抬頭與簽章
  ·  2 份下載失敗（平台回 success:false）

type:  cpa=會計師事務所　consult=會計師事務所體系顧問公司　verify=驗證／查驗機構　unknown=未具名
"""

VERDICTS = {
    # ── 會計師事務所出具 ──────────────────────────────────────────
    "1236": ("安永聯合會計師事務所", "cpa", "EY 抬頭「會計師有限確信報告」；惟標的為永續報告書所選定之永續績效資訊、依確信準則3000號，非溫室氣體聲明確信"),
    "1445": ("安永聯合會計師事務所", "cpa", "EY 抬頭「溫室氣體聲明之會計師有限確信報告」，依 ISO 14064-1:2018"),
    "1736": ("KPMG 安侯建業聯合會計師事務所", "cpa", "「溫室氣體聲明確信報告」，附件頁眉為 KPMG"),
    "2480": ("勤業眾信聯合會計師事務所", "cpa", "Deloitte 抬頭確信報告"),
    "2530": ("安永聯合會計師事務所", "cpa", "EY 抬頭確信報告（台北基隆路一段333號9樓）"),
    "2637": ("安永聯合會計師事務所", "cpa", "EY 抬頭確信報告"),
    "3017": ("安永聯合會計師事務所", "cpa", "EY 抬頭確信報告；內文載明範疇一部分排放量係依據其他查驗機構之查證聲明書"),
    "3416": ("思享永續聯合會計師事務所", "cpa", "確信準則3410號有限確信，簽章頁為思享永續聯合會計師事務所（115/06/13）"),
    "3685": ("嘉威聯合會計師事務所", "cpa", "Jia Wei & Co., CPAs 抬頭，確信準則3410號"),
    "4506": ("嘉威聯合會計師事務所", "cpa", "Jia Wei & Co., CPAs 抬頭，GHG Protocol 基準"),
    "5225": ("安永聯合會計師事務所", "cpa", "EY 抬頭「東科控股股份有限公司溫室氣體聲明會計師有限確信報告」"),
    "5519": ("安永聯合會計師事務所", "cpa", "EY 抬頭確信報告（高雄中正三路2號17樓）"),
    "6472": ("國富浩華聯合會計師事務所", "cpa", "Crowe (TW) CPAs 抬頭"),
    "6533": ("思享永續聯合會計師事務所", "cpa", "Live Susthinkability CPAs Firm 抬頭，確信準則3410號"),
    "6771": ("青山永續聯合會計師事務所", "cpa", "青山永續抬頭，確信範疇一、二，GHG Protocol"),
    "8936": ("秉承聯合會計師事務所", "cpa", "Legendary & Steadfast Accountancy 抬頭，確信準則3410號"),

    # ── 2026-08-31 重跑新增（會計師事務所） ──
    "9955": ("正大聯合會計師事務所（Grant Thornton）", "cpa", "Grant Thornton／正大聯合抬頭「溫室氣體聲明確信報告」，確信準則3410號有限確信，標的為類別1、2，期間2025/01/01~12/31"),

    # ── 會計師事務所體系之顧問公司 ────────────────────────────────
    "1721": ("資誠普華綠色科技有限公司（PwC）", "consult", "pwc 抬頭查證意見（編虏 260G012TC）；出具實體為顧問公司非會計師事務所"),

    # ── 驗證／查驗機構 ───────────────────────────────────────────
    "1342": ("台灣德國北德技術監護顧問股份有限公司（TÜV NORD）", "verify", "組織型溫室氣體查驗意見書"),
    "1522": ("台灣衛理國際品保驗證股份有限公司（Bureau Veritas Certification Taiwan）", "verify", "查證聲明"),
    "1525": ("法標國際驗證股份有限公司（AFNOR ASIA）", "verify", "溫室氣體查驗報告意見書，ISO 14064-3:2019"),
    "1536": ("英國標準協會（BSI Group Singapore Pte. Ltd. Taiwan Branch）", "verify", "Verification Opinion"),
    "1541": ("法標國際驗證股份有限公司（AFNOR ASIA）", "verify", "溫室氣體查驗報告意見書"),
    "1558": ("台灣德國萊因技術監護顧問股份有限公司（TÜV Rheinland）", "verify", "查證聲明書；持證者為張家港伸興機電"),
    "1612": ("財團法人台灣大電力研究試驗中心", "verify", "溫室氣體查驗意見書，ISO 14064-3:2019"),
    "2022": ("立恩威國際驗證股份有限公司（DNV）", "verify", "查驗意見書"),
    "2072": ("台灣德國萊因技術監護顧問股份有限公司（TÜV Rheinland）", "verify", "Verification Report 封面；持證者為世紀離岸風電設備"),
    "2303": ("立恩威國際驗證股份有限公司（DNV）", "verify", "DNV Business Assurance 查驗意見書"),
    "2413": ("台灣檢驗科技股份有限公司（SGS）", "verify", "SGS 查驗意見書（TW25/00661GG 等）"),
    "2454": ("台灣衛理國際品保驗證股份有限公司（Bureau Veritas Certification Taiwan）", "verify", "Assurance Opinion for MEDIATEK INC."),
    "2481": ("台灣檢驗科技股份有限公司（SGS）", "verify", "溫室氣體查驗意見書 TW26/00167GG"),
    "2504": ("立恩威國際驗證股份有限公司（DNV）", "verify", "溫室氣體排放量查驗意見書；查驗準則含金管會上市櫃公司永續發展路徑圖"),
    "2646": ("英國標準協會（BSI Taiwan）", "verify", "Verification Opinion"),
    "2850": ("亞瑞仕國際驗證股份有限公司（ARES）", "verify", "溫室氣體查證意見聲明書"),
    "2887": ("英國標準協會（BSI Taiwan）", "verify", "Verification Opinion — Verified with Comments"),
    "3014": ("英商勞盛股份有限公司台灣分公司（LRQA）", "verify", "LRQA Independent Assurance Statement"),
    "3023": ("台灣德國萊因技術監護顧問股份有限公司（TÜV Rheinland）", "verify", "聲明；核查機構為萊茵檢測認證服務（中國）"),
    "3034": ("英商勞盛股份有限公司台灣分公司（LRQA）", "verify", "LRQA Independent Assurance Statement"),
    "3035": ("台灣檢驗科技股份有限公司（SGS）", "verify", "溫室氣體查驗意見書 TW26/00325GG"),
    "3189": ("立恩威國際驗證股份有限公司（DNV）", "verify", "DNV Business Assurance 查驗意見書"),
    "3227": ("台灣德國北德技術監護顧問股份有限公司（TÜV NORD）", "verify", "GHG Verification Opinion for PIXART IMAGING"),
    "3303": ("台灣德國北德技術監護顧問股份有限公司（TÜV NORD）", "verify", "組織溫室氣體查驗意見書"),
    "3504": ("法標國際驗證股份有限公司（AFNOR ASIA）", "verify", "溫室氣體查驗報告意見書"),
    "3516": ("格瑞國際驗證有限公司", "verify", "查證聲明書 Great-GHGER-26-0401"),
    "3652": ("法標國際驗證股份有限公司（AFNOR ASIA）", "verify", "溫室氣體查驗報告意見書 THGHG24472-01"),
    "3715": ("DQS Taiwan Inc.", "verify", "Verification Opinion"),
    "4129": ("英商勞盛股份有限公司台灣分公司（LRQA）", "verify", "LRQA Independent Assurance Statement"),
    "4566": ("英國標準協會（BSI Taiwan）", "verify", "Verification Opinion"),
    "4939": ("安德森國際標準驗證有限公司（Anderson International Certification）", "verify", "溫室氣體查驗意見書"),
    "4979": ("立恩威國際驗證股份有限公司（DNV）", "verify", "DNV Business Assurance 查驗意見書"),
    "5007": ("台灣檢驗科技股份有限公司（SGS）", "verify", "溫室氣體查驗意見書 TW26/00228GG"),
    "5016": ("財團法人金屬工業研究發展中心（MIRDC）", "verify", "溫室氣體查證意見 GHG-2026-08"),
    "5285": ("財團法人金屬工業研究發展中心（MIRDC）", "verify", "溫室氣體查證意見 MIRDC-2026-02G"),
    "5288": ("英國標準協會（BSI Assurance UK Ltd）", "verify", "Verification Opinion；持證者為 VIETNAM PRECISION INDUSTRIAL"),
    "5353": ("台灣德國北德技術監護顧問股份有限公司（TÜV NORD）", "verify", "GHG Verification Opinion for TAILYN TECHNOLOGIES"),
    "5371": ("台灣檢驗科技股份有限公司（SGS）", "verify", "溫室氣體查驗意見書 TW26/00077GG"),
    "6152": ("財團法人台灣商品檢測驗證中心（ETC 商檢中心）", "verify", "溫室氣體查證意見書 113-GHG-013"),
    "6156": ("法標國際驗證股份有限公司（AFNOR ASIA）", "verify", "溫室氣體查驗報告意見書"),
    "6182": ("英國標準協會（BSI Taiwan）", "verify", "Verification Opinion"),
    "6183": ("亞瑞仕國際驗證股份有限公司（ARES）", "verify", "溫室氣體查證意見聲明書 ARES/TW/F2605712G"),
    "6201": ("亞瑞仕國際驗證股份有限公司（ARES）", "verify", "Opinion Statement ARES/TW/F2607709G"),
    "6426": ("亞瑞仕國際驗證股份有限公司（ARES）", "verify", "上傳者為盤查報告書，內載「由外部查證機構（亞瑞仕國際驗證股份有限公司）執行外部查證並取得查證聲明書」"),
    "6446": ("法標國際驗證股份有限公司（AFNOR ASIA）", "verify", "溫室氣體查驗報告意見書 THGHG23002-01"),
    "6491": ("立恩威國際驗證股份有限公司（DNV）", "verify", "DNV 查驗意見書"),
    "6526": ("Bureau Veritas Certification Hong Kong Limited", "verify", "Assurance Opinion for Airoha Technology"),
    "6641": ("中國質量認證中心（CQC）", "verify", "溫室氣體核查陳述；責任方為川源（中國）機械"),
    "6703": ("英國標準協會（BSI Taiwan）", "verify", "AUP 議定程序報告，BSI 明示 do not express any assurance"),
    "6805": ("華測認證有限公司（CTI）", "verify", "溫室氣體排放核查聲明；責任方為深圳市富世達通訊"),
    "6907": ("台灣檢驗科技股份有限公司（SGS）", "verify", "溫室氣體查驗意見書 TW26/00326GG"),
    "7718": ("立恩威國際驗證股份有限公司（DNV）", "verify", "DNV 查驗意見書"),
    "7728": ("英國標準協會（BSI Taiwan）", "verify", "Verification Opinion"),
    "7750": ("台灣德國北德技術監護顧問股份有限公司（TÜV NORD）", "verify", "組織型溫室氣體查驗意見書"),
    "8040": ("Bureau Veritas Certification Hong Kong Limited", "verify", "Assurance Opinion"),
    "8059": ("立恩威國際驗證股份有限公司（DNV）", "verify", "DNV 查驗意見書；委託者為金寶／泰金寶"),
    "8424": ("立恩威國際驗證股份有限公司（DNV）", "verify", "查驗意見書；惟內文所述組織為國產建材實業集團"),
    "9958": ("台灣德國萊因技術監護顧問股份有限公司（TÜV Rheinland）", "verify", "查證聲明書；持證者為世紀鋼鐵結構"),

    # ── 2026-08-31 重跑新增（驗證／查驗機構） ──
    "1319": ("財團法人工業技術研究院（ITRI）", "verify", "ITRI 抬頭「溫室氣體排放量查驗意見」編虏 OC-0908-2026006-01-01；盤查期間 114/01/01~114/12/31，類別1、2 合理保證、類別3、4 有限保證，115/07/27 發行"),
    "1737": ("財團法人台灣大電力研究試驗中心", "verify", "「溫室氣體查驗意見書」聲明書編號 07G0001-20251，ISO 14064-3:2019，涵蓋期間 2025/01/01~12/31，直接及間接排放合理保證、其他間接有限保證"),
    "1802": ("台灣檢驗科技股份有限公司（SGS）", "verify", "SGS 抬頭「溫室氣體查驗意見書－2025 年溫室氣體排放資訊」，ISO 14064-3:2019；惟上傳版本通篇標示 Draft"),
    "2312": ("立恩威國際驗證股份有限公司（DNV）", "verify", "72 頁多機構查驗聲明彙編；台灣母公司 Kinpo Electronics, Inc. 及台灣關係企業由 DNV 查驗（Taipei, 2026-05-08），海外子公司另由 BVQI 巴西、Sustainable SI 墨西哥、IAPMO、SGS、DQS、BSI、法標 AFNOR、華測 CTI 等分別出具"),
    "3515": ("立恩威國際驗證股份有限公司（DNV）", "verify", "DNV「Impartial Engagement Opinion」編虏 C870794-2025-AG-TWN-DNVRev.1（Taipei, 2026-08-13）；受查對象載明為 PEGATRON CORPORATION AND SUBSIDIARIES"),
    "6505": ("英國標準協會（BSI Group Singapore Pte. Ltd. Taiwan Branch）", "verify", "BSI Verification Opinion，FPCC 2025 GHG Report 範疇一、二合理確信（Reasonable Assurance），依 GHG Protocol，2026-07-28 出具"),
    "7711": ("立恩威國際驗證股份有限公司（DNV）", "verify", "DNV Independent Verification Opinion No. C888717-2025-AG-TWN-DNV（Taipei, 2026-08-13），標的為 ASROCK RACK INCORPORATION 2025 年溫室氣體盤查報告"),

    # ── 未具名 ────────────────────────────────────────────────
    "2449": ("（未具名）", "unknown", "上傳者為盤查報告書，僅載「後續由公正第三者查驗機構進行外部查證作業」，未具名"),
    "9906": ("（未具名）", "unknown", "上傳者為盤查報告書，第六章「外部查證單位名稱」欄位空白"),
}

# 需人工確認事由（超出「出具者辨識」本身的疑義）
FLAGS = {
    "1802": "上傳版本通篇為 Draft：意見書編號印為「UNIQUE CODE」、簽署日期為「2026年XX月XX日」，屬未定稿之查驗意見書",
    "2312": "上傳者為 72 頁多機構查驗聲明彙編而非單一確信報告；出具者係以台灣母公司之 DNV 意見為準",
    "3515": "上傳者為母集團和碩（PEGATRON CORPORATION AND SUBSIDIARIES）之查驗意見，附錄A 所列 22 個場址均不含華擎本身，查驗範圍未涵蓋申報公司",
    "1236": "確信標的為永續報告書永續績效資訊（準則3000號），非溫室氣體聲明；填報於溫室氣體確信欄位，範圍認定需確認",
    "2449": "上傳者為盤查報告書而非確信／查證報告，查證機構未具名",
    "6152": "上傳者為 2024 年度（113-GHG-013）查證意見書，非 2025 年度",
    "6426": "上傳者為盤查報告書而非查證聲明書本身",
    "6446": "意見書數據期間為 2024/01/01~2024/12/31，屬 113 年度，非 2025 年度",
    "6703": "為 AUP 議定程序報告，BSI 明示不表示任何確信結論，不應計為確信或查證",
    "6907": "SGS 意見書內文所載受查公司為「智原科技」而非雅特力，疑上傳錯檔或範本誤植",
    "8424": "DNV 意見書內文所載組織為「國產建材實業集團」而非惠普，疑上傳錯檔",
    "9906": "上傳者為盤查報告書，外部查證單位欄位空白",
    "3017": "確信報告載明範疇一部分排放量係依據其他查驗機構之查證聲明書，非全數由安永確信",
}

# 下載失敗（平台回 success:false）
FAILED = {
    "3511": "矽瑪",
    "5244": "弘凱",
}
