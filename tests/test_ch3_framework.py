"""Unit tests for Chapter 3 Framework Module."""
import unittest
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "scripts"))
from modules.ch3_framework import (
    build_chapter_3,
    build_framework_table,
    NATURE_CH3_DATA,
    ENGLISH_CH3_DATA,
    MATH_CH3_DATA,
    CHINESE_CH3_DATA,
)
from modules.ch1_2_admin import load_attendance_data


class TestCh3Framework(unittest.TestCase):
    def test_framework_tables_all_subjects(self):
        for sub, gr in [("國語文", 3), ("數學", 4), ("英語文", 7), ("自然", 5)]:
            tbl = build_framework_table(sub, gr)
            self.assertIsNotNone(tbl)
            text = "".join(tbl.itertext())
            self.assertIn("主要評量向度", text)
            self.assertIn("題數", text)
            self.assertIn("總計", text)

    def test_chapter_3_generation(self):
        att = load_attendance_data("自然", 5)
        elems = build_chapter_3("自然", 5, att, year="115")
        text = "".join("".join(e.itertext()) for e in elems)
        self.assertIn("參、檢測對象與檢測工具", text)
        self.assertIn("一、檢測對象", text)
        self.assertIn("二、檢測工具", text)
        self.assertIn("表3-1", text)


if __name__ == "__main__":
    unittest.main()
