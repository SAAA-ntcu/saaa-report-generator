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


def append_to_body_before_sectpr(doc: Document, elem: Any) -> None:
    """Appends an element to body while ensuring body.sectPr remains the final child."""
    body = doc.element.body
    sectPr = body.sectPr
    if sectPr is not None:
        body.insert(body.index(sectPr), elem)
    else:
        body.append(elem)


def build_blank_skeleton(subject: str, grade: int | str, year: str = "115", is_7n: bool = False) -> Document:
    """Builds a fresh standard SAAA report skeleton document from scratch."""
    special_templates_dir = Path(__file__).resolve().parents[2] / "resources" / "templates" / "special_cohorts"
    cov_tpl = special_templates_dir / "7N_cover_template.docx"

    if is_7n and cov_tpl.exists():
        import tempfile
        from docxtpl import DocxTemplate

        tpl = DocxTemplate(str(cov_tpl))
        tpl.render({"YEAR": str(year), "year": str(year), "SUBJECT": str(subject), "subject": str(subject)})
        with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as tmp:
            tmp_path = Path(tmp.name)
        tpl.save(str(tmp_path))
        cov_doc = Document(str(tmp_path))
        doc = Document()
        doc.sections[0].top_margin = cov_doc.sections[0].top_margin
        doc.sections[0].bottom_margin = cov_doc.sections[0].bottom_margin
        doc.sections[0].left_margin = cov_doc.sections[0].left_margin
        doc.sections[0].right_margin = cov_doc.sections[0].right_margin
        for elem in cov_doc.element.body.iterchildren():
            if elem.tag in (qn("w:p"), qn("w:tbl")):
                append_to_body_before_sectpr(doc, deepcopy(elem))
        try:
            tmp_path.unlink()
        except Exception:
            pass
    else:
        doc = Document()
        grade_clean = int(re.sub(r"\D", "", str(grade)) or 7)
        g_cn = chinese_numeral(grade_clean)

        # --- Standard Cover Page ---
        p_cov1 = create_paragraph(f"{year}年度縣市學生學習能力檢測", font_size_pt=26.0, bold=True, space_after_pt=18.0)
        p_cov1_jc = OxmlElement("w:jc")
        p_cov1_jc.set(qn("w:val"), "center")
        p_cov1.find(qn("w:pPr")).append(p_cov1_jc)
        append_to_body_before_sectpr(doc, p_cov1)

        p_cov2 = create_paragraph(f"{subject}{g_cn}年級結案報告", font_size_pt=28.0, bold=True, space_after_pt=180.0)
        p_cov2_jc = OxmlElement("w:jc")
        p_cov2_jc.set(qn("w:val"), "center")
        p_cov2.find(qn("w:pPr")).append(p_cov2_jc)
        append_to_body_before_sectpr(doc, p_cov2)

        p_unit1 = create_paragraph("國立臺中教育大學", font_size_pt=16.0, bold=True, space_after_pt=6.0)
        p_unit1_jc = OxmlElement("w:jc")
        p_unit1_jc.set(qn("w:val"), "center")
        p_unit1.find(qn("w:pPr")).append(p_unit1_jc)
        append_to_body_before_sectpr(doc, p_unit1)

        p_unit2 = create_paragraph("測驗統計與適性學習研究中心", font_size_pt=16.0, bold=True, space_after_pt=12.0)
        p_unit2_jc = OxmlElement("w:jc")
        p_unit2_jc.set(qn("w:val"), "center")
        p_unit2.find(qn("w:pPr")).append(p_unit2_jc)
        append_to_body_before_sectpr(doc, p_unit2)

        p_date = create_paragraph(f"中華民國 {year} 年12月", font_size_pt=14.0, bold=True, space_after_pt=24.0)
        p_date_jc = OxmlElement("w:jc")
        p_date_jc.set(qn("w:val"), "center")
        p_date.find(qn("w:pPr")).append(p_date_jc)
        append_to_body_before_sectpr(doc, p_date)

    # Section break after cover
    p_break = parse_xml(
        '<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        '<w:pPr><w:sectPr><w:type w:val="nextPage"/></w:sectPr></w:pPr>'
        '</w:p>'
    )
    append_to_body_before_sectpr(doc, p_break)

    # --- Table of Contents (TOC) ---
    p_toc_title = create_paragraph("目　錄", font_size_pt=18.0, bold=True, space_after_pt=18.0)
    p_toc_title_jc = OxmlElement("w:jc")
    p_toc_title_jc.set(qn("w:val"), "center")
    p_toc_title.find(qn("w:pPr")).append(p_toc_title_jc)
    append_to_body_before_sectpr(doc, p_toc_title)

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
        append_to_body_before_sectpr(doc, p_t)

    # Section break after TOC
    p_break2 = parse_xml(
        '<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        '<w:pPr><w:sectPr><w:type w:val="nextPage"/></w:sectPr></w:pPr>'
        '</w:p>'
    )
    append_to_body_before_sectpr(doc, p_break2)

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
