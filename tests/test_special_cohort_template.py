"""Unit tests for Special Cohort 7N (E7N, C7N & M7N) Chapter 1~3 Templates."""
import unittest
import sys
from pathlib import Path
from docxtpl import DocxTemplate
import docx
from docx.shared import Pt

REPO_ROOT = Path(__file__).resolve().parents[1]
E7N_TPL_PATH = REPO_ROOT / "resources" / "templates" / "special_cohorts" / "E7N_ch1_2_3_template.docx"
C7N_TPL_PATH = REPO_ROOT / "resources" / "templates" / "special_cohorts" / "C7N_ch1_2_3_template.docx"
M7N_TPL_PATH = REPO_ROOT / "resources" / "templates" / "special_cohorts" / "M7N_ch1_2_3_template.docx"
COVER_7N_TPL_PATH = REPO_ROOT / "resources" / "templates" / "special_cohorts" / "7N_cover_template.docx"


class TestSpecialCohortTemplate(unittest.TestCase):
    def test_templates_exist(self):
        self.assertTrue(E7N_TPL_PATH.exists(), f"E7N Template not found at {E7N_TPL_PATH}")
        self.assertTrue(C7N_TPL_PATH.exists(), f"C7N Template not found at {C7N_TPL_PATH}")
        self.assertTrue(M7N_TPL_PATH.exists(), f"M7N Template not found at {M7N_TPL_PATH}")
        self.assertTrue(COVER_7N_TPL_PATH.exists(), f"7N Cover Template not found at {COVER_7N_TPL_PATH}")

    def test_margins_conformance(self):
        for path in [E7N_TPL_PATH, C7N_TPL_PATH, M7N_TPL_PATH, COVER_7N_TPL_PATH]:
            doc = docx.Document(str(path))
            sec = doc.sections[0]
            self.assertAlmostEqual(sec.top_margin.pt, 64.8, delta=0.5, msg=f"Top margin fail: {path.name}")
            self.assertAlmostEqual(sec.bottom_margin.pt, 50.4, delta=0.5, msg=f"Bottom margin fail: {path.name}")
            self.assertAlmostEqual(sec.left_margin.pt, 56.7, delta=0.5, msg=f"Left margin fail: {path.name}")
            self.assertAlmostEqual(sec.right_margin.pt, 49.6, delta=0.5, msg=f"Right margin fail: {path.name}")

    def test_7n_cover_render(self):
        tpl = DocxTemplate(str(COVER_7N_TPL_PATH))
        tpl.render({"YEAR": "115", "year": "115", "SUBJECT": "英語文", "subject": "英語文"})
        out_test = REPO_ROOT / "output" / "test_cover_rendered.docx"
        out_test.parent.mkdir(parents=True, exist_ok=True)
        tpl.save(str(out_test))

        rendered = docx.Document(str(out_test))
        full_text = "\n".join(p.text for p in rendered.paragraphs)
        self.assertIn("115年度桃園市學生學習能力檢測", full_text)
        self.assertIn("英語文七年級新生施測結果報告", full_text)
        self.assertIn("國立臺中教育大學", full_text)
        self.assertIn("National Taichung University of Education", full_text)
        self.assertIn("測驗統計與適性學習研究中心", full_text)
        self.assertIn("中華民國 115 年 12 月", full_text)

    def test_e7n_docxtpl_render_dynamic_loop(self):
        tpl = DocxTemplate(str(E7N_TPL_PATH))
        context = {
            "year": "115",
            "month": "9",
            "day": "25",
            "當年度參與縣市": "桃園市",
            "國字大寫": "一",
            "enrolled_count": "23,273",
            "tested_count": "22,191",
            "boy_count": "11,500",
            "girl_count": "10,691",
            "absence_rate": "4.6",
            "listening_dimensions": "辨識句意、基本問答",
            "listening_dimensions_final": "言談理解",
            "reading_dimensions": "單題",
            "reading_dimensions_final": "題組",
            "framework_rows": [
                {"domain": "聽力", "sub_domain": "辨識句意", "item_numbers": "1、2、3"},
                {"domain": "聽力", "sub_domain": "基本問答", "item_numbers": "4、5、6、7、8"},
                {"domain": "聽力", "sub_domain": "言談理解", "item_numbers": "9、10、11、12、13、14、15"},
                {"domain": "閱讀", "sub_domain": "單題", "item_numbers": "16、17、...、27"},
                {"domain": "閱讀", "sub_domain": "題組", "item_numbers": "28、29、...、35"},
            ],
            "total_items": "35",
        }

        tpl.render(context)
        out_test = REPO_ROOT / "output" / "test_e7n_rendered.docx"
        out_test.parent.mkdir(parents=True, exist_ok=True)
        tpl.save(str(out_test))

        rendered = docx.Document(str(out_test))
        full_text = "\n".join(p.text for p in rendered.paragraphs)
        self.assertIn("學生人數為23,273人，施測人數為22,191人", full_text)
        self.assertIn("男生11,500人、女生10,691人，整體缺考率約為4.6%", full_text)
        self.assertIn("聽力評量向度包含辨識句意、基本問答及言談理解", full_text)
        self.assertIn("閱讀能力則包含單題及題組", full_text)

        table = rendered.tables[0]
        self.assertEqual(len(table.rows), 7)
        self.assertEqual(table.rows[1].cells[0].text.strip(), "聽力")
        self.assertEqual(table.rows[-1].cells[0].text.strip(), "合計")
        self.assertEqual(table.rows[-1].cells[1].text.strip(), "35題")

    def test_c7n_docxtpl_render_taxonomy_table(self):
        tpl = DocxTemplate(str(C7N_TPL_PATH))
        context = {
            "year": "115",
            "month": "9",
            "day": "25",
            "參與縣市_": "桃園市",
            "參與縣市_final": "",
            "國字大寫": "一",
            "enrolled_count": "23,273",
            "tested_count": "22,223",
            "boy_count": "11,500",
            "girl_count": "10,723",
            "absence_rate": "4.5",
            "knowledge_count": "12",
            "reading_count": "18",
        }

        tpl.render(context)
        out_test = REPO_ROOT / "output" / "test_c7n_rendered.docx"
        out_test.parent.mkdir(parents=True, exist_ok=True)
        tpl.save(str(out_test))

        rendered = docx.Document(str(out_test))
        full_text = "\n".join(p.text for p in rendered.paragraphs)
        self.assertIn("學生人數為23,273人，施測人數為22,223人", full_text)
        self.assertIn("男生11,500人、女生10,723人，整體缺考率約為4.5%", full_text)
        self.assertIn("學力檢測施測之正式題本總題數為30題", full_text)

        table = rendered.tables[0]
        self.assertEqual(len(table.rows), 40)
        total_cell_text = table.rows[39].cells[1].text
        self.assertIn("語文知識12題、閱讀理解18題", total_cell_text)

    def test_m7n_docxtpl_render_matrix_tables(self):
        tpl = DocxTemplate(str(M7N_TPL_PATH))
        context = {
            "year": "115",
            "month": "9",
            "day": "25",
            "當年度參與縣市": "桃園市",
            "國字大寫": "一",
            "enrolled_count": "23,273",
            "tested_count": "22,205",
            "boy_count": "11,500",
            "girl_count": "10,705",
            "absence_rate": "4.6",
        }

        tpl.render(context)
        out_test = REPO_ROOT / "output" / "test_m7n_rendered.docx"
        out_test.parent.mkdir(parents=True, exist_ok=True)
        tpl.save(str(out_test))

        rendered = docx.Document(str(out_test))
        full_text = "\n".join(p.text for p in rendered.paragraphs)
        self.assertIn("學生人數為23,273人，施測人數為22,205人", full_text)
        self.assertIn("男生11,500人、女生10,705人，整體缺考率約為4.6%", full_text)
        self.assertIn("學力檢測施測之正式題本總題數為25題", full_text)
        self.assertIn("檢測科目包含國語文、數學、英語文和自然", full_text)

        # Table 0 (Table 3-1: 5 rows x 4 cols)
        t0 = rendered.tables[0]
        self.assertEqual(len(t0.rows), 5)
        self.assertEqual(len(t0.columns), 4)
        self.assertEqual(t0.rows[1].cells[1].text.strip(), "7")

        # Table 1 (Table 3-2: 13 rows x 5 cols)
        t1 = rendered.tables[1]
        self.assertEqual(len(t1.rows), 13)
        self.assertEqual(len(t1.columns), 5)
        self.assertEqual(t1.rows[1].cells[3].text.strip(), "7")


if __name__ == "__main__":
    unittest.main()
