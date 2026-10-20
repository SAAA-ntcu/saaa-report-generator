"""Unit tests for Chapter 1 & Chapter 2 Administration Module."""
import unittest
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "scripts"))
from modules.ch1_2_admin import (
    DEFAULT_18_BACKGROUND_COUNTIES,
    build_chapter_1,
    build_chapter_2,
    load_attendance_data,
    chinese_numeral,
)


class TestCh12Admin(unittest.TestCase):
    def test_chinese_numeral(self):
        self.assertEqual(chinese_numeral(3), "三")
        self.assertEqual(chinese_numeral(10), "十")
        self.assertEqual(chinese_numeral(18), "十八")
        self.assertEqual(chinese_numeral(25), "二十五")

    def test_chapter_1_natural_science_clause(self):
        # Annual type with natural science
        elems_annual = build_chapter_1(year="115", has_natural_science=True)
        text_annual = "".join("".join(e.itertext()) for e in elems_annual)
        expected_sent = "檢測對象為國民中小學三年級至八年級學生，檢測科目包含國語文、數學、英語文和自然。"
        self.assertIn(expected_sent, text_annual)
        self.assertIn("等十八縣市共同參與", text_annual)
        self.assertIn("國立臺中教育大學測驗統計與適性學習研究中心", text_annual)

    def test_chapter_1_without_natural_science_fallback(self):
        elems_old = build_chapter_1(year="114", has_natural_science=False)
        text_old = "".join("".join(e.itertext()) for e in elems_old)
        expected_sent = "檢測科目包含國語文、數學和英語文。"
        self.assertIn(expected_sent, text_old)

    def test_chapter_2_table_2_1(self):
        att_data = load_attendance_data("國語文", 3)
        elems = build_chapter_2("國語文", 3, att_data, year="115")
        text = "".join("".join(e.itertext()) for e in elems)
        self.assertIn("表2-1", text)
        self.assertIn("總計", text)
        self.assertGreater(att_data["report_total"], 0)


if __name__ == "__main__":
    unittest.main()
