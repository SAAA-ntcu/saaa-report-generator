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
)


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


if __name__ == "__main__":
    unittest.main()
