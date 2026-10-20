"""SAAA Report Engine: Full Document Assembler Module.
Orchestrates:
  - Base skeleton / template loading (e.g. 114 report template or clean skeleton)
  - Modular injection of Chapters 1, 2, 3, 4, 5, 6, 7
  - Front matter update (Cover, Year, Title, Date: 中華民國 115 年12月)
  - TOC (目錄) generation and page break control
  - Font unification to 微軟正黑體
  - Proper names annotation (專名號單底線) via proper_names.py
"""

from __future__ import annotations

import re
from copy import deepcopy
from pathlib import Path
from typing import Any

from docx import Document
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn

try:
    from ..proper_names import underline_counties
except (ImportError, ValueError):
    from proper_names import underline_counties
from .ch1_2_admin import chinese_numeral, create_paragraph, set_run_font


def unify_all_fonts(document: Document, font_name: str = "微軟正黑體") -> None:
    """Sets all text runs and styles across the document to the specified font."""
    body = document.element.body
    for rPr in body.iter(qn("w:rPr")):
        rFonts = rPr.find(qn("w:rFonts"))
        if rFonts is None:
            rFonts = OxmlElement("w:rFonts")
            rPr.append(rFonts)
        for attr in ("ascii", "eastAsia", "hAnsi", "cs"):
            rFonts.set(qn(f"w:{attr}"), font_name)


def clean_orphaned_bookmarks(document: Document) -> None:
    """Removes broken or orphaned bookmarks to prevent Word opening warnings."""
    body = document.element.body
    for b in list(body.iter(qn("w:bookmarkStart"))) + list(body.iter(qn("w:bookmarkEnd"))):
        p = b.getparent()
        if p is not None:
            p.remove(b)


def build_blank_skeleton(subject: str, grade: int, year: str = "115") -> Document:
    """Builds a fresh standard SAAA report skeleton document from scratch."""
    doc = Document()
    body = doc.element.body
    g_cn = chinese_numeral(grade)
    stage = "國民小學" if grade <= 6 else "國民中學"
    sub_full = "自然科學領域" if subject == "自然" else f"{subject}領域" if "領域" not in subject else subject

    # --- Cover Page ---
    p_cov1 = create_paragraph(f"{year}年度縣市學生學習能力檢測", font_size_pt=26.0, bold=True, space_after_pt=18.0)
    p_cov1_jc = OxmlElement("w:jc")
    p_cov1_jc.set(qn("w:val"), "center")
    p_cov1.find(qn("w:pPr")).append(p_cov1_jc)
    body.append(p_cov1)

    p_cov2 = create_paragraph(f"{subject}{g_cn}年級結案報告", font_size_pt=28.0, bold=True, space_after_pt=180.0)
    p_cov2_jc = OxmlElement("w:jc")
    p_cov2_jc.set(qn("w:val"), "center")
    p_cov2.find(qn("w:pPr")).append(p_cov2_jc)
    body.append(p_cov2)

    p_unit1 = create_paragraph("國立臺中教育大學", font_size_pt=16.0, bold=True, space_after_pt=6.0)
    p_unit1_jc = OxmlElement("w:jc")
    p_unit1_jc.set(qn("w:val"), "center")
    p_unit1.find(qn("w:pPr")).append(p_unit1_jc)
    body.append(p_unit1)

    p_unit2 = create_paragraph("測驗統計與適性學習研究中心", font_size_pt=16.0, bold=True, space_after_pt=12.0)
    p_unit2_jc = OxmlElement("w:jc")
    p_unit2_jc.set(qn("w:val"), "center")
    p_unit2.find(qn("w:pPr")).append(p_unit2_jc)
    body.append(p_unit2)

    p_date = create_paragraph(f"中華民國 {year} 年12月", font_size_pt=14.0, bold=True, space_after_pt=24.0)
    p_date_jc = OxmlElement("w:jc")
    p_date_jc.set(qn("w:val"), "center")
    p_date.find(qn("w:pPr")).append(p_date_jc)
    body.append(p_date)

    # Section break after cover
    p_break = parse_xml(
        '<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        '<w:pPr><w:sectPr><w:type w:val="nextPage"/></w:sectPr></w:pPr>'
        '</w:p>'
    )
    body.append(p_break)

    # --- Table of Contents (TOC) ---
    p_toc_title = create_paragraph("目　錄", font_size_pt=18.0, bold=True, space_after_pt=18.0)
    p_toc_title_jc = OxmlElement("w:jc")
    p_toc_title_jc.set(qn("w:val"), "center")
    p_toc_title.find(qn("w:pPr")).append(p_toc_title_jc)
    body.append(p_toc_title)

    toc_items = [
        "壹、背景與目的",
        "貳、參與縣市與檢測規模",
        "參、檢測對象與檢測工具",
        "肆、試題分析與教學建議",
        "伍、整體學力表現分析",
        "陸、整體教學建議",
        "柒、附錄",
    ]
    for item in toc_items:
        p_t = create_paragraph(item, font_size_pt=12.0, line_spacing_pt=24.0, space_after_pt=4.0)
        ind = OxmlElement("w:ind")
        ind.set(qn("w:left"), "720")
        p_t.find(qn("w:pPr")).append(ind)
        body.append(p_t)

    # Section break after TOC
    p_break2 = parse_xml(
        '<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        '<w:pPr><w:sectPr><w:type w:val="nextPage"/></w:sectPr></w:pPr>'
        '</w:p>'
    )
    body.append(p_break2)

    return doc


class FullReportAssembler:
    """Assembles all report chapters into a complete final Word document."""

    def __init__(
        self,
        subject: str,
        grade: int,
        year: str = "115",
        base_template_path: str | Path | None = None,
    ) -> None:
        self.subject = subject
        self.grade = grade
        self.year = year
        self.base_template_path = Path(base_template_path) if base_template_path else None
        
        if self.base_template_path and self.base_template_path.exists():
            self.document = Document(str(self.base_template_path))
        else:
            self.document = build_blank_skeleton(subject, grade, year=year)

    def append_chapter_elements(self, elements: list[Any], add_page_break: bool = True) -> None:
        """Appends chapter elements to document body before the final sectPr."""
        body = self.document.element.body
        sectPr = body.sectPr
        target_idx = body.index(sectPr) if sectPr is not None else len(body)
        
        if add_page_break and len(elements) > 0:
            # Set pageBreakBefore on first paragraph if possible
            first = elements[0]
            if first.tag == qn("w:p"):
                pPr = first.find(qn("w:pPr"))
                if pPr is None:
                    pPr = OxmlElement("w:pPr")
                    first.insert(0, pPr)
                pbb = pPr.find(qn("w:pageBreakBefore"))
                if pbb is None:
                    pbb = OxmlElement("w:pageBreakBefore")
                    pPr.append(pbb)
                    
        for elem in elements:
            body.insert(target_idx, deepcopy(elem))
            target_idx += 1

    def import_external_chapter_docx(self, docx_path: str | Path, add_page_break: bool = True) -> None:
        """Imports body elements from an externally generated chapter DOCX (e.g. Chapter 4 or 6)."""
        p = Path(docx_path)
        if not p.exists():
            return
        ext_doc = Document(str(p))
        elements = [child for child in ext_doc.element.body.iterchildren() if child.tag in (qn("w:p"), qn("w:tbl"))]
        self.append_chapter_elements(elements, add_page_break=add_page_break)

    def finalize(self, output_path: str | Path) -> Path:
        """Applies final font unification, bookmark cleaning, proper name marks, and saves."""
        out = Path(output_path)
        out.parent.mkdir(parents=True, exist_ok=True)
        
        # 1. Clean bookmarks
        clean_orphaned_bookmarks(self.document)
        
        # 2. Unify fonts
        unify_all_fonts(self.document, "微軟正黑體")
        
        # 3. Apply proper names underline (桃園市, 國立臺中教育大學, etc.)
        underline_counties(self.document)
        
        self.document.save(str(out))
        return out
