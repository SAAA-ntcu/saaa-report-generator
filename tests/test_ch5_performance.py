"""Unit tests for Chapter 5 Performance Analysis Module."""
import unittest
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "scripts"))
from modules.ch5_performance import (
    Item,
    INTERVALS,
    build_chapter_5,
    build_area_cross_table,
    clean_description,
    create_standalone_chapter_5_docx,
    sort_areas,
)
from docx import Document
import tempfile


class TestCh5Performance(unittest.TestCase):
    def test_clean_description(self):
        dirty = "加減法計算  加減法計算。"
        cleaned = clean_description(dirty)
        self.assertNotIn("  ", cleaned)

    def test_chapter_5_intervals_grouping(self):
        items = [
            Item(number=1, rate=0.95, codes=("N-4-1",), area="數與計算"),
            Item(number=2, rate=0.82, codes=("N-4-2",), area="數與計算"),
            Item(number=3, rate=0.15, codes=("N-4-3",), area="數與計算"),
        ]
        desc = {"N-4-1": "十進位結構", "N-4-2": "四則運算", "N-4-3": "未知數解題"}
        elems = build_chapter_5("數學", 4, items, desc)
        text = "".join("".join(e.itertext()) for e in elems)
        self.assertIn("伍、整體學力表現分析", text)
        self.assertIn("圖5-1", text)
        self.assertIn("表5-1-1", text)
        self.assertIn("91~100", text)
        self.assertIn("0~10", text)
        self.assertIn("二、國民小學四年級學生於不同學習重點之答對率表現", text)
        self.assertIn("國民小學數學四年級學生學習能力檢測", text)

    def test_7n_cohort_conditional_text(self):
        items = [
            Item(number=1, rate=0.88, codes=("N-6-1",), area="數與計算"),
            Item(number=2, rate=0.45, codes=("S-6-1",), area="空間與形狀"),
        ]
        desc = {"N-6-1": "質數與合數", "S-6-1": "圓周率與圓面積"}
        elems = build_chapter_5("數學", "7N", items, desc, is_7n=True)
        text = "".join("".join(e.itertext()) for e in elems)
        self.assertIn("二、國民中學七年級新生於不同學習重點之答對率表現", text)
        self.assertIn("桃園市國民中學數學七年級新生學習能力檢測", text)
        self.assertIn("圖5-1　數學七年級新生整體學生學力表現分布圖", text)

    def test_chinese_3col_vs_math_4col(self):
        m_items = [Item(number=1, rate=0.92, codes=("N-3-1",), area="數與計算")]
        m_desc = {"N-3-1": "一萬以內的數"}
        m_tbl = build_area_cross_table(m_items, m_desc, is_chinese=False)
        # Verify Math table header has gridSpan=2
        m_xml = m_tbl.xml
        self.assertIn('w:gridSpan w:val="2"', m_xml)

        c_items = [Item(number=1, rate=0.95, codes=("4-Ⅱ-1",), area="形音知識")]
        c_desc = {"4-Ⅱ-1": "認識常用國字"}
        c_tbl = build_area_cross_table(c_items, c_desc, is_chinese=True)
        c_xml = c_tbl.xml
        self.assertNotIn('w:gridSpan w:val="2"', c_xml)

    def test_canonical_area_sorting(self):
        shuffled = ["關係", "空間與形狀", "數與計算", "量與實測"]
        sorted_areas = sort_areas(shuffled, "數學", 3)
        self.assertEqual(sorted_areas, ["數與計算", "量與實測", "空間與形狀", "關係"])

    def test_standalone_export(self):
        items = [Item(number=1, rate=0.75, codes=("N-3-1",), area="數與計算")]
        desc = {"N-3-1": "一萬以內的數"}
        with tempfile.TemporaryDirectory() as tmpdir:
            out_file = Path(tmpdir) / "test_ch5.docx"
            saved = create_standalone_chapter_5_docx(out_file, "數學", 3, items, desc)
            self.assertTrue(saved.exists())
            doc = Document(str(saved))
            full_text = "\n".join(p.text for p in doc.paragraphs)
            self.assertIn("伍、整體學力表現分析", full_text)


if __name__ == "__main__":
    unittest.main()
