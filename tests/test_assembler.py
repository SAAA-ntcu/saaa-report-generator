"""Unit tests for Full Document Assembler Module."""
import unittest
import sys
import tempfile
from pathlib import Path
from docx import Document

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "scripts"))
from modules.docx_assembler import FullReportAssembler, build_blank_skeleton
from modules.ch1_2_admin import build_chapter_1, build_chapter_2, load_attendance_data
from modules.ch3_framework import build_chapter_3
from modules.ch7_appendix import build_chapter_7


class TestAssembler(unittest.TestCase):
    def test_blank_skeleton_structure(self):
        doc = build_blank_skeleton("國語文", 3, year="115")
        text = "".join(doc.element.body.itertext())
        self.assertIn("115年度縣市學生學習能力檢測", text)
        self.assertIn("國語文三年級結案報告", text)
        self.assertIn("國立臺中教育大學", text)
        self.assertIn("目　錄", text)

    def test_full_assembly_workflow(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            out_path = Path(tmpdir) / "test_report.docx"
            assembler = FullReportAssembler(subject="自然", grade=5, year="115")
            
            # Append Ch 1
            ch1 = build_chapter_1(year="115", has_natural_science=True)
            assembler.append_chapter_elements(ch1)
            
            # Append Ch 2
            att = load_attendance_data("自然", 5)
            ch2 = build_chapter_2("自然", 5, att, year="115")
            assembler.append_chapter_elements(ch2)
            
            # Append Ch 3
            ch3 = build_chapter_3("自然", 5, att, year="115")
            assembler.append_chapter_elements(ch3)
            
            # Append Ch 7
            ch7 = build_chapter_7("自然", 5, year="115")
            assembler.append_chapter_elements(ch7)
            
            saved = assembler.finalize(out_path)
            self.assertTrue(saved.exists())
            
            doc = Document(str(saved))
            full_text = "".join(doc.element.body.itertext())
            self.assertIn("壹、背景與目的", full_text)
            self.assertIn("檢測科目包含國語文、數學、英語文和自然。", full_text)
            self.assertIn("貳、參與縣市與檢測規模", full_text)
            self.assertIn("參、檢測對象與檢測工具", full_text)
            self.assertIn("柒、附錄", full_text)


if __name__ == "__main__":
    unittest.main()
