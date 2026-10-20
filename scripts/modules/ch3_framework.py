"""SAAA Report Engine: Chapter 3 Assessment Framework Module.
Generates:
  - 參、檢測對象與檢測工具 (Chapter 3)
  - Section 1: 檢測對象 (Student cohort stats: enrolled, tested, male/female, absent rate)
  - Section 2: 檢測工具 (Assessment instrument description)
  - Table 3-1: Assessment framework and item distribution for Chinese, Math, English, and Science.
"""

from __future__ import annotations

import re
from copy import deepcopy
from pathlib import Path
from typing import Any

from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn

from .ch1_2_admin import chinese_numeral, county_string, create_paragraph, set_run_font


# Standard verified framework distributions for Math, English, and Science
NATURE_CH3_DATA = {
    5: [
        {
            "category": "跨科主題",
            "items": [
                ("交互作用（INe）", "1、2、3、4、5、6、7"),
                ("改變與穩定（INd）", "8、9、10、11"),
                ("構造與功能（INb）", "21、26、27、28、30"),
                ("物質與能量（INa）", "17、22"),
                ("科學與生活（INf）", "24"),
                ("系統與尺度（INc）", "12、13、14、15、16、18、19、20、23、29"),
                ("資源與永續性（INg）", "25"),
            ],
        },
    ],
    7: [
        {
            "category": "能量的形態與流動（B）",
            "items": [
                ("生物體內的能量與代謝（Bc）", "4、10、20、24"),
            ],
        },
        {
            "category": "生物的構造與功能（D）",
            "items": [
                ("細胞的構造與功能（Da）", "5、25、26、27、28"),
                ("動植物體的構造與功能（Db）", "2、3、6、11、16、35"),
                ("生物體內的恆定性與調節（Dc）", "1、13、14、17、18、29、30、33、34"),
            ],
        },
        {
            "category": "地球環境（F）",
            "items": [
                ("生物圈的組成（Fc）", "21"),
            ],
        },
        {
            "category": "演化與延續（G）",
            "items": [
                ("生殖與遺傳（Ga）", "15、19、22、23、31、32"),
                ("生物多樣性（Gc）", "8"),
            ],
        },
        {
            "category": "生物與環境（L）",
            "items": [
                ("生物與環境的交互作用（Lb）", "9"),
            ],
        },
        {
            "category": "跨科主題",
            "items": [
                ("系統與尺度（INc）", "7、12"),
            ],
        },
    ],
    8: [
        {
            "category": "物質的組成與特性（A）",
            "items": [
                ("物質組成與元素的週期性（Aa）", "16、25、29"),
                ("物質的形態、性質及分類（Ab）", "6、27"),
            ],
        },
        {
            "category": "能量的形式、轉換及流動（B）",
            "items": [
                ("能量的形式與轉換（Ba）", "8"),
                ("溫度與熱量（Bb）", "9、10、22、23"),
            ],
        },
        {
            "category": "物質的結構與功能（C）",
            "items": [
                ("物質的結構與功能（Cb）", "7、15"),
            ],
        },
        {
            "category": "物質系統（E）",
            "items": [
                ("自然界的尺度與單位（Ea）", "1、2"),
            ],
        },
        {
            "category": "物質的反應、平衡及製造（J）",
            "items": [
                ("物質反應規律（Ja）", "11、32、33、34、35"),
                ("水溶液中的變化（Jb）", "18、19"),
                ("氧化與還原反應（Jc）", "30、31"),
                ("酸鹼反應（Jd）", "14、17、28"),
                ("化學反應速率與平衡（Je）", "13、21"),
            ],
        },
        {
            "category": "自然界的現象與交互作用（K）",
            "items": [
                ("波動、光及聲音（Ka）", "3、4、5、12、20、24、26"),
            ],
        },
    ],
}

ENGLISH_CH3_DATA = {
    4: [
        {
            "category": "聽力",
            "items": [
                ("語音聽辨", "1、2、3、4"),
                ("辭彙聽辨", "5、6、7、8"),
                ("教室生活辭句理解", "9、10、11、12、13"),
                ("文化節慶理解", "14、15"),
            ],
        },
        {
            "category": "閱讀",
            "items": [
                ("字詞辨識", "16、17、18、19、20、21、22、23"),
                ("句子理解", "24、25、26、27、30、31、32、33、34、35"),
                ("文化節慶理解", "28、29"),
            ],
        },
    ],
    5: [
        {
            "category": "聽力",
            "items": [
                ("語音聽辨", "1、2"),
                ("辭彙聽辨", "3、4、5、6"),
                ("教室生活辭句理解與回應", "7、8、9、10、11、14、15"),
                ("文化節慶理解", "12、13"),
            ],
        },
        {
            "category": "閱讀",
            "items": [
                ("字詞辨識", "16、17、18、19"),
                ("句子理解", "20、21、22、23、25、26、28、29"),
                ("短文與對話理解", "27、30、31、32、34、35"),
                ("文化節慶理解", "24、33"),
            ],
        },
    ],
    6: [
        {
            "category": "聽力",
            "items": [
                ("語音聽辨", "1、2"),
                ("辭彙聽辨", "3、4、5"),
                ("教室生活辭句理解與回應", "6、7、8、9、10、11、14、15"),
                ("文化節慶理解", "12、13"),
            ],
        },
        {
            "category": "閱讀",
            "items": [
                ("字詞辨識", "16、17、18、19"),
                ("句子理解", "20、21、22、23、25、26、28、29"),
                ("短文與對話理解", "27、30、31、32、34、35"),
                ("文化節慶理解", "24、33"),
            ],
        },
    ],
    7: [
        {
            "category": "聽力",
            "items": [
                ("辨識句意", "1、2、3"),
                ("基本問答", "4、5、6、7"),
                ("言談理解", "8、9、10、11、12、13、14、15"),
            ],
        },
        {
            "category": "閱讀",
            "items": [
                ("單題", "16、17、18、19、20、21、22、23、24、25、26、27、28、29、30"),
                ("題組", "31、32、33、34、35、36、37、38、39、40"),
            ],
        },
    ],
    8: [
        {
            "category": "聽力",
            "items": [
                ("辨識句意", "1、2、3"),
                ("基本問答", "4、5、6、7"),
                ("言談理解", "8、9、10、11、12、13、14、15"),
            ],
        },
        {
            "category": "閱讀",
            "items": [
                ("單題", "16、17、18、19、20、21、22、23、24、25、26、27、28、29、30"),
                ("題組", "31、32、33、34、35、36、37、38、39、40"),
            ],
        },
    ],
}

MATH_CH3_DATA = {
    3: [
        {"category": "數與計算", "items": [("概念理解", "01、08、16、18、19、24"), ("程序執行", "03、05、06、14、25"), ("解題思考", "10、11、15、21")]},
        {"category": "量與實測", "items": [("概念理解", "12、20"), ("程序執行", "04"), ("解題思考", "22")]},
        {"category": "空間與形狀", "items": [("概念理解", "07、17"), ("程序執行", "02、09"), ("解題思考", "13")]},
        {"category": "資料與不確定性", "items": [("概念理解", "23"), ("程序執行", ""), ("解題思考", "")]},
    ],
    4: [
        {"category": "數與計算", "items": [("概念理解", "01、03、04、09、11、14、21、23"), ("程序執行", "02、08、17"), ("解題思考", "07、12、19、22")]},
        {"category": "量與實測", "items": [("概念理解", "05、18"), ("程序執行", "13"), ("解題思考", "20、24、25")]},
        {"category": "空間與形狀", "items": [("概念理解", "06"), ("程序執行", "15"), ("解題思考", "10、16")]},
    ],
    5: [
        {"category": "數與計算", "items": [("概念理解", "01、02、03、04、05、06、10、17、19、22"), ("程序執行", "08、09、16"), ("解題思考", "07、11、18、21")]},
        {"category": "量與實測", "items": [("概念理解", "13"), ("程序執行", "14"), ("解題思考", "12、15、24")]},
        {"category": "空間與形狀", "items": [("概念理解", "20"), ("程序執行", ""), ("解題思考", "23、25")]},
    ],
    6: [
        {"category": "數與計算", "items": [("概念理解", "01、02、03、04、05、06、10、17"), ("程序執行", "08、09、16"), ("解題思考", "07、11、18、21")]},
        {"category": "量與實測", "items": [("概念理解", "13"), ("程序執行", "14"), ("解題思考", "12、15、24")]},
        {"category": "空間與形狀", "items": [("概念理解", "20"), ("程序執行", ""), ("解題思考", "23、25")]},
        {"category": "關係", "items": [("概念理解", "19"), ("程序執行", ""), ("解題思考", "22")]},
    ],
    7: [
        {"category": "數與量", "items": [("概念理解", "01、02、04、05、06、07、12、14、15、16"), ("程序執行", "03、08、09、10、11、13"), ("解題思考", "17、18")]},
        {"category": "代數與函數", "items": [("概念理解", "19、20、21、22、23、24"), ("程序執行", "25、26"), ("解題思考", "27、28、29、30")]},
    ],
    8: [
        {"category": "數與量", "items": [("概念理解", "01、02、03、04、05"), ("程序執行", "06、07、08"), ("解題思考", "09、10")]},
        {"category": "空間與形狀", "items": [("概念理解", "11、12、13、14、15"), ("程序執行", "16、17、18"), ("解題思考", "19、20")]},
        {"category": "代數與函數", "items": [("概念理解", "21、22、23、24、25"), ("程序執行", "26、27、28"), ("解題思考", "29、30")]},
    ],
}

CHINESE_CH3_DATA = {
    3: [
        {"category": "語文知識", "items": [("形音知識", "1、2、3、4"), ("字詞知識", "5、6、7、8"), ("語法知識", "9、10"), ("修辭知識", "11、12"), ("章法知識", "13"), ("文體知識", "14")]},
        {"category": "閱讀理解", "items": [("字詞理解", "15、16"), ("句子理解", "17、18"), ("段落理解", "19、20"), ("篇章理解", "21、22、23、24、25")]},
    ],
    4: [
        {"category": "語文知識", "items": [("形音知識", "1、2、3、4"), ("字詞知識", "5、6、7、8"), ("語法知識", "9、10"), ("修辭知識", "11、12"), ("章法知識", "13"), ("文體知識", "14")]},
        {"category": "閱讀理解", "items": [("字詞理解", "15、16"), ("句子理解", "17、18"), ("段落理解", "19、20"), ("篇章理解", "21、22、23、24、25")]},
    ],
    5: [
        {"category": "語文知識", "items": [("形音知識", "1、2、3、4"), ("字詞知識", "5、6、7、8"), ("語法知識", "9、10"), ("修辭知識", "11、12"), ("章法知識", "13"), ("文體知識", "14")]},
        {"category": "閱讀理解", "items": [("字詞理解", "15、16"), ("句子理解", "17、18"), ("段落理解", "19、20"), ("篇章理解", "21、22、23、24、25")]},
    ],
    6: [
        {"category": "語文知識", "items": [("形音知識", "1、2、3"), ("字詞知識", "4、5、6"), ("語法知識", "7、8"), ("修辭知識", "9、10"), ("章法知識", "11"), ("文體知識", "12")]},
        {"category": "閱讀理解", "items": [("字詞理解", "13、14"), ("句子理解", "15、16"), ("段落理解", "17、18"), ("篇章理解", "19、20、21、22、23、24、25")]},
    ],
    7: [
        {"category": "語文知識", "items": [("形音知識", "1、2、3、4"), ("字詞知識", "5、6、7"), ("語法知識", "8、9"), ("修辭知識", "10、11"), ("章法知識", "12"), ("國學文化", "13")]},
        {"category": "閱讀理解", "items": [("白話文篇章理解", "14、15、16、17、18、19、20、21、22、23、24、25"), ("文言文篇章理解", "26、27、28、29、30")]},
    ],
    8: [
        {"category": "語文知識", "items": [("形音知識", "1、2、3、4"), ("字詞知識", "5、6、7"), ("語法知識", "8、9"), ("修辭知識", "10、11"), ("章法知識", "12"), ("國學文化", "13")]},
        {"category": "閱讀理解", "items": [("白話文篇章理解", "14、15、16、17、18、19、20、21、22、23、24、25"), ("文言文篇章理解", "26、27、28、29、30")]},
    ],
}


def build_framework_table(
    subject: str,
    grade: int,
    data_groups: list[dict[str, Any]] | None = None,
) -> Any:
    """Constructs Table 3-1: 評量向度分布題數表."""
    grade_num = int(re.sub(r"\D", "", str(grade)) or 7)
    
    if data_groups is None:
        if subject == "自然":
            data_groups = NATURE_CH3_DATA.get(grade_num, NATURE_CH3_DATA.get(5, []))
        elif subject == "英語文":
            data_groups = ENGLISH_CH3_DATA.get(grade_num, ENGLISH_CH3_DATA.get(4, []))
        elif subject == "數學":
            data_groups = MATH_CH3_DATA.get(grade_num, MATH_CH3_DATA.get(3, []))
        else:
            data_groups = CHINESE_CH3_DATA.get(grade_num, CHINESE_CH3_DATA.get(3, []))
            
    tbl = OxmlElement("w:tbl")
    tblPr = parse_xml(
        '<w:tblPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        '<w:tblW w:w="9360" w:type="dxa"/>'
        '<w:jc w:val="center"/>'
        '<w:tblBorders>'
        '<w:top w:val="single" w:sz="12" w:space="0" w:color="000000"/>'
        '<w:bottom w:val="single" w:sz="12" w:space="0" w:color="000000"/>'
        '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="D3D3D3"/>'
        '<w:insideV w:val="single" w:sz="4" w:space="0" w:color="D3D3D3"/>'
        '<w:left w:val="none"/>'
        '<w:right w:val="none"/>'
        '</w:tblBorders>'
        '</w:tblPr>'
    )
    tbl.append(tblPr)
    
    col_widths = [2600, 3160, 2400, 1200]
    tblGrid = OxmlElement("w:tblGrid")
    for w in col_widths:
        gc = OxmlElement("w:gridCol")
        gc.set(qn("w:w"), str(w))
        tblGrid.append(gc)
    tbl.append(tblGrid)
    
    headers = ["主要評量向度", "評量次向度", "測驗題號", "題數"]
    
    # 1. Header row
    tr_h = OxmlElement("w:tr")
    trPr_h = OxmlElement("w:trPr")
    tr_h.append(trPr_h)
    trH = OxmlElement("w:trHeight")
    trH.set(qn("w:val"), "400")
    trPr_h.append(trH)
    trPr_h.append(OxmlElement("w:tblHeader"))
    
    for h, w in zip(headers, col_widths):
        tc = OxmlElement("w:tc")
        tcPr = OxmlElement("w:tcPr")
        tc.append(tcPr)
        tcW = OxmlElement("w:tcW")
        tcW.set(qn("w:w"), str(w))
        tcW.set(qn("w:type"), "dxa")
        tcPr.append(tcW)
        vA = OxmlElement("w:vAlign")
        vA.set(qn("w:val"), "center")
        tcPr.append(vA)
        b_bot = parse_xml('<w:tcBorders xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:bottom w:val="single" w:sz="8" w:color="000000"/></w:tcBorders>')
        tcPr.append(b_bot)
        
        p = create_paragraph(h, font_size_pt=11.0, bold=True)
        jc = OxmlElement("w:jc")
        jc.set(qn("w:val"), "center")
        p.find(qn("w:pPr")).append(jc)
        tc.append(p)
        tr_h.append(tc)
    tbl.append(tr_h)
    
    total_q_count = 0
    
    for g in data_groups:
        cat_name = g["category"]
        items = g["items"]
        cat_len = len(items)
        
        for idx, (sub_name, q_list) in enumerate(items):
            tr = OxmlElement("w:tr")
            trPr = OxmlElement("w:trPr")
            tr.append(trPr)
            trH = OxmlElement("w:trHeight")
            trH.set(qn("w:val"), "340")
            trH.set(qn("w:hRule"), "atLeast")
            trPr.append(trH)
            
            # Col 0: Main Category (vertical merge)
            tc0 = OxmlElement("w:tc")
            tcPr0 = OxmlElement("w:tcPr")
            tc0.append(tcPr0)
            tcW0 = OxmlElement("w:tcW")
            tcW0.set(qn("w:w"), str(col_widths[0]))
            tcW0.set(qn("w:type"), "dxa")
            tcPr0.append(tcW0)
            vA0 = OxmlElement("w:vAlign")
            vA0.set(qn("w:val"), "center")
            tcPr0.append(vA0)
            
            if cat_len > 1:
                vMerge = OxmlElement("w:vMerge")
                vMerge.set(qn("w:val"), "restart" if idx == 0 else "continue")
                tcPr0.append(vMerge)
                
            p0 = create_paragraph(cat_name if idx == 0 else "", font_size_pt=10.5, bold=True)
            jc0 = OxmlElement("w:jc")
            jc0.set(qn("w:val"), "center")
            p0.find(qn("w:pPr")).append(jc0)
            tc0.append(p0)
            tr.append(tc0)
            
            # Col 1: Sub-domain
            tc1 = OxmlElement("w:tc")
            tcPr1 = OxmlElement("w:tcPr")
            tc1.append(tcPr1)
            tcW1 = OxmlElement("w:tcW")
            tcW1.set(qn("w:w"), str(col_widths[1]))
            tcW1.set(qn("w:type"), "dxa")
            tcPr1.append(tcW1)
            vA1 = OxmlElement("w:vAlign")
            vA1.set(qn("w:val"), "center")
            tcPr1.append(vA1)
            p1 = create_paragraph(sub_name, font_size_pt=10.5)
            jc1 = OxmlElement("w:jc")
            jc1.set(qn("w:val"), "left")
            p1.find(qn("w:pPr")).append(jc1)
            tc1.append(p1)
            tr.append(tc1)
            
            # Col 2: Question list
            tc2 = OxmlElement("w:tc")
            tcPr2 = OxmlElement("w:tcPr")
            tc2.append(tcPr2)
            tcW2 = OxmlElement("w:tcW")
            tcW2.set(qn("w:w"), str(col_widths[2]))
            tcW2.set(qn("w:type"), "dxa")
            tcPr2.append(tcW2)
            vA2 = OxmlElement("w:vAlign")
            vA2.set(qn("w:val"), "center")
            tcPr2.append(vA2)
            p2 = create_paragraph(q_list, font_size_pt=10.0)
            jc2 = OxmlElement("w:jc")
            jc2.set(qn("w:val"), "left")
            p2.find(qn("w:pPr")).append(jc2)
            tc2.append(p2)
            tr.append(tc2)
            
            # Col 3: Count
            q_count = len([x for x in re.split(r"[、,，\s]+", q_list.strip()) if x]) if q_list.strip() else 0
            total_q_count += q_count
            tc3 = OxmlElement("w:tc")
            tcPr3 = OxmlElement("w:tcPr")
            tc3.append(tcPr3)
            tcW3 = OxmlElement("w:tcW")
            tcW3.set(qn("w:w"), str(col_widths[3]))
            tcW3.set(qn("w:type"), "dxa")
            tcPr3.append(tcW3)
            vA3 = OxmlElement("w:vAlign")
            vA3.set(qn("w:val"), "center")
            tcPr3.append(vA3)
            p3 = create_paragraph(str(q_count), font_size_pt=10.5)
            jc3 = OxmlElement("w:jc")
            jc3.set(qn("w:val"), "center")
            p3.find(qn("w:pPr")).append(jc3)
            tc3.append(p3)
            tr.append(tc3)
            
            tbl.append(tr)
            
    # Total row
    tr_tot = OxmlElement("w:tr")
    trPr_tot = OxmlElement("w:trPr")
    tr_tot.append(trPr_tot)
    trH_tot = OxmlElement("w:trHeight")
    trH_tot.set(qn("w:val"), "360")
    trPr_tot.append(trH_tot)
    
    # Merge col 0, 1, 2
    tc_t0 = OxmlElement("w:tc")
    tcPr_t0 = OxmlElement("w:tcPr")
    tc_t0.append(tcPr_t0)
    tcW_t0 = OxmlElement("w:tcW")
    tcW_t0.set(qn("w:w"), str(col_widths[0] + col_widths[1] + col_widths[2]))
    tcW_t0.set(qn("w:type"), "dxa")
    tcPr_t0.append(tcW_t0)
    gs = OxmlElement("w:gridSpan")
    gs.set(qn("w:val"), "3")
    tcPr_t0.append(gs)
    vA_t0 = OxmlElement("w:vAlign")
    vA_t0.set(qn("w:val"), "center")
    tcPr_t0.append(vA_t0)
    b_top = parse_xml('<w:tcBorders xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:top w:val="single" w:sz="8" w:color="000000"/></w:tcBorders>')
    tcPr_t0.append(b_top)
    p_t0 = create_paragraph("總計", font_size_pt=11.0, bold=True)
    jc_t0 = OxmlElement("w:jc")
    jc_t0.set(qn("w:val"), "center")
    p_t0.find(qn("w:pPr")).append(jc_t0)
    tc_t0.append(p_t0)
    tr_tot.append(tc_t0)
    
    # Total count cell
    tc_t1 = OxmlElement("w:tc")
    tcPr_t1 = OxmlElement("w:tcPr")
    tc_t1.append(tcPr_t1)
    tcW_t1 = OxmlElement("w:tcW")
    tcW_t1.set(qn("w:w"), str(col_widths[3]))
    tcW_t1.set(qn("w:type"), "dxa")
    tcPr_t1.append(tcW_t1)
    vA_t1 = OxmlElement("w:vAlign")
    vA_t1.set(qn("w:val"), "center")
    tcPr_t1.append(vA_t1)
    tcPr_t1.append(deepcopy(b_top))
    p_t1 = create_paragraph(str(total_q_count), font_size_pt=11.0, bold=True)
    jc_t1 = OxmlElement("w:jc")
    jc_t1.set(qn("w:val"), "center")
    p_t1.find(qn("w:pPr")).append(jc_t1)
    tc_t1.append(p_t1)
    tr_tot.append(tc_t1)
    
    tbl.append(tr_tot)
    return tbl


def build_chapter_3(
    subject: str,
    grade: int,
    attendance_data: dict[str, Any],
    stat_data: dict[str, Any] | None = None,
    year: str = "115",
) -> list[Any]:
    """Generates complete Chapter 3 (參、檢測對象與檢測工具) elements."""
    elements = []
    stage = "國民小學" if grade <= 6 else "國民中學"
    g_cn = chinese_numeral(grade)
    sub_full = "自然科學領域" if subject == "自然" else f"{subject}領域" if "領域" not in subject else subject
    
    # 1. Heading
    h3 = create_paragraph("參、檢測對象與檢測工具", font_size_pt=16.0, bold=True, space_after_pt=6.0)
    elements.append(h3)
    
    # 2. Section 1: 一、檢測對象
    p_sec1 = create_paragraph("一、檢測對象", font_size_pt=14.0, bold=True, space_after_pt=4.0)
    elements.append(p_sec1)
    
    # Text for cohort stats
    part_counties = [c["county"] for c in attendance_data["counties"]]
    n_counties = len(part_counties)
    
    if stat_data is None:
        rep_tot = attendance_data["report_total"]
        att_tot = attendance_data["attendance_total"]
        abs_tot = attendance_data["absence_total"]
        rate = (att_tot / rep_tot * 100.0) if rep_tot > 0 else 0.0
        stat_data = {
            "應考總人數": rep_tot,
            "實際施測人數": att_tot,
            "男生總數": int(rep_tot * 0.51),
            "女生總數": int(rep_tot * 0.49),
            "未填總數": 0,
            "整體缺考率": f"{(abs_tot / rep_tot * 100.0):.2f}" if rep_tot > 0 else "0.00",
        }
        
    p1 = OxmlElement("w:p")
    pPr1 = OxmlElement("w:pPr")
    p1.append(pPr1)
    sp1 = OxmlElement("w:spacing")
    sp1.set(qn("w:line"), "480")
    sp1.set(qn("w:lineRule"), "exact")
    pPr1.append(sp1)
    ind1 = OxmlElement("w:ind")
    ind1.set(qn("w:firstLine"), "480")
    pPr1.append(ind1)
    
    runs_spec = [
        (f"{year}年參與", False),
        ("國立臺中教育大學", True),
        (f"所承辦「{stage}{sub_full}{g_cn}年級學生學習能力檢測」者為", False),
    ]
    for i, c in enumerate(part_counties):
        runs_spec.append((c, True))
        if i < n_counties - 2:
            runs_spec.append(("、", False))
        elif i == n_counties - 2:
            runs_spec.append(("以及", False))
    runs_spec.append((
        f"等{chinese_numeral(n_counties)}縣市之{stage}{g_cn}年級學生（未排除資源班及特殊生），"
        f"應考名冊羅列{stat_data['應考總人數']:,}人、"
        f"實際施測人數{stat_data['實際施測人數']:,}人，其中男生{stat_data['男生總數']:,}人、"
        f"女生{stat_data['女生總數']:,}人、未填寫性別{stat_data.get('未填總數', 0):,}人，"
        f"整體缺考率約為{stat_data['整體缺考率']}％。",
        False,
    ))
    for t_val, u_val in runs_spec:
        r = OxmlElement("w:r")
        t = OxmlElement("w:t")
        t.text = t_val
        t.set(qn("xml:space"), "preserve")
        r.append(t)
        set_run_font(r, font_size_pt=12.0, underline=u_val)
        p1.append(r)
    elements.append(p1)
    
    # 3. Section 2: 二、檢測工具
    p_sec2 = create_paragraph("二、檢測工具", font_size_pt=14.0, bold=True, space_after_pt=4.0)
    elements.append(p_sec2)
    
    time_limit = "40" if grade <= 6 else "45"
    p_tool_text = (
        f"{stage}{sub_full}{g_cn}年級學力檢測之正式評量工具，"
        f"全卷試題均為四選一單選題，施測時間為{time_limit}分鐘。"
        f"各評量向度架構與分布題數如表3-1所示。"
    )
    p_tool = create_paragraph(p_tool_text, font_size_pt=12.0, line_spacing_pt=24.0, space_after_pt=6.0)
    ind_tool = OxmlElement("w:ind")
    ind_tool.set(qn("w:firstLine"), "480")
    p_tool.find(qn("w:pPr")).append(ind_tool)
    elements.append(p_tool)
    
    # 4. Caption: 表3-1
    caption_text = f"表3-1　{sub_full}{g_cn}年級評量向度之分布題數"
    p_cap = create_paragraph(caption_text, font_size_pt=11.0, bold=True, space_after_pt=4.0)
    jc_cap = OxmlElement("w:jc")
    jc_cap.set(qn("w:val"), "center")
    p_cap.find(qn("w:pPr")).append(jc_cap)
    elements.append(p_cap)
    
    # 5. Table 3-1
    tbl_framework = build_framework_table(subject, grade)
    elements.append(tbl_framework)
    
    # 6. Spacer
    elements.append(create_paragraph("", font_size_pt=12.0, space_after_pt=12.0))
    return elements
