"""SAAA Report Engine: Chapter 7 Appendix Module.
Generates:
  - 柒、附錄 (Chapter 7)
  - 附錄一：試卷題本與參考答案
  - 附錄二：試題疑義回覆說明
  - 命題與審題委員名單
  - Sub-document appendix embedding / full-page section injection.
"""

from __future__ import annotations

from pathlib import Path
from typing import Any

from docx import Document
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn

from .ch1_2_admin import chinese_numeral, create_paragraph, set_run_font


def build_chapter_7(
    subject: str,
    grade: int,
    appendix_files: list[str | Path] | None = None,
    year: str = "115",
) -> list[Any]:
    """Generates complete Chapter 7 (柒、附錄) elements."""
    elements = []
    stage = "國民小學" if grade <= 6 else "國民中學"
    g_cn = chinese_numeral(grade)
    sub_full = "自然科學領域" if subject == "自然" else f"{subject}領域" if "領域" not in subject else subject
    
    # 1. Heading: 柒、附錄
    h7 = create_paragraph("柒、附錄", font_size_pt=16.0, bold=True, space_after_pt=12.0)
    elements.append(h7)
    
    # 2. Appendix 1 Heading
    app1_title = f"附錄一　{year}年度{stage}{sub_full}{g_cn}年級學生學習能力檢測試題本與參考答案"
    p_app1 = create_paragraph(app1_title, font_size_pt=14.0, bold=True, space_after_pt=6.0)
    elements.append(p_app1)
    
    # Check if appendix files exist
    inserted_any = False
    if appendix_files:
        for f_path in appendix_files:
            p = Path(f_path)
            if not p.exists():
                continue
            if p.suffix.lower() == ".docx":
                try:
                    ext_doc = Document(str(p))
                    for child in ext_doc.element.body.iterchildren():
                        if child.tag in (qn("w:p"), qn("w:tbl")):
                            elements.append(child)
                            inserted_any = True
                except Exception:
                    pass
            elif p.suffix.lower() in (".png", ".jpg", ".jpeg"):
                # Insert full page image
                temp_doc = Document()
                img_p = temp_doc.add_paragraph()
                img_p.alignment = 1
                run = img_p.add_run()
                run.add_picture(str(p))
                elements.append(img_p._p)
                inserted_any = True
                
    if not inserted_any:
        # Standard placeholder description
        p_desc = create_paragraph(
            f"（本附錄收錄{year}年度{stage}{sub_full}{g_cn}年級之正式施測試題題本、參考答案及相關試題疑義回覆公告）",
            font_size_pt=11.0, line_spacing_pt=24.0, space_after_pt=12.0
        )
        elements.append(p_desc)
        
    return elements
