---
name: saaa-report-generator
description: >-
  Automates the complete generation of diagnostic assessment reports (Chapters 1 through 7)
  for Taiwanese academic diagnostic assessments (學力檢測 / 新生學力測驗).
  Supports 4 subjects: Chinese (國語文 C), Math (數學 M), English (英語文 E), and Science (自然科學 S)
  across standard elementary/junior high grades (3~8) and special entry cohorts (C7N, M7N, E7N, S7N).
  Includes Chapter 1 background with the official 18-county union and the new Natural Science curriculum clause,
  Chapter 2 scale/attendance tables, Chapter 3 assessment frameworks, Chapter 4 booklet screenshot integration
  with 24pt line height, Chapter 5 performance distributions and live 108 curriculum PDF descriptor extraction,
  Chapter 6 multi-dimensional pass rate matrices, Chapter 7 appendix attachments, and full document assembly.
---

# SAAA Report Engine Skill (Chapters 1~7 & Full Document Assembly)

This skill provides an end-to-end automated workflow for producing official Taiwanese diagnostic assessment reports:
- **Chapter 1 (壹、背景與目的)**: Standardized background text with official 18-county union, proper name single underlines, and the official updated clause for annual cohorts: `檢測對象為國民中小學三年級至八年級學生，檢測科目包含國語文、數學、英語文和自然。`
- **Chapter 2 (貳、參與縣市與檢測規模)**: Complete Table 2-1 attendance/scale counts from participation and attendance registers.
- **Chapter 3 (參、檢測對象與檢測工具)**: Cohort statistics (enrolled, tested, gender ratio, absence rate) and Table 3-1 assessment framework & item distribution.
- **Chapter 4 (肆、試題分析與教學建議)**: Complete item analysis tables with discrimination indices, distractor distributions, high/low scoring group pass rates, sub-domain codes, booklet PDF screenshots with question-number removal, uniform font-scale planning, and exact 24pt line spacing.
- **Chapter 5 (伍、整體學力表現分析)**: Complete CTT/IRT 3PL & PISA/TIMSS standard text, Figure 5-1 score distribution curves, dynamic stage and cohort condition logic (7N vs annual cohorts), and Tables 5-1-1~N 10-interval pass rate cross-tabulations with canonical domain ordering, subject-adaptive 3-column (Chinese) vs 4-column merged header (Math/English/Science) formatting, and live 108 Curriculum PDF descriptor extraction via PyMuPDF. Supports standalone Chapter 5 docx generation.
- **Chapter 6 (陸、整體教學建議 / 表6-1 各評量向度通過率)**: Structured multi-dimensional pass rate matrices, high/overall/low group breakdowns, cognitive level intersections, diagonal empty cell borders, and customized column widths.
- **Chapter 7 (柒、附錄)**: Automated booklet PDF page rendering (181.4068 DPI / 1500px width), official `DDEBF7` shaded heading with 24pt exact line height, seamless full-page picture injection (489.0pt × 691.8pt) with exact page breaks (`w:pageBreakBefore`), answer keys, inquiry announcements, and standalone/bundled appendix document generation.
- **Master Document Assembler**: Seamlessly stitches any combination of chapters into a finalized Word document with cover page, table of contents (目 錄), unified `微軟正黑體` typography, and proper name marks.

---

## Supported Subjects and Cohorts

| Subject | Code | Standard Grades | Special Cohorts | Templates Used |
| :--- | :--- | :--- | :--- | :--- |
| **國語文** (Chinese) | `C` | 3, 4, 5, 6, 7, 8 | `C7N` | `pydocxTemplet_C_new.docx`, `C_demo.docx` |
| **數學** (Math) | `M` | 3, 4, 5, 6, 7, 8 | `M7N` | `pydocxTemplet_M.docx`, `S_M_demo.docx` |
| **英語文** (English) | `E` | 4, 5, 6, 7, 8 | `E7N` | `pydocxTemplet_E.docx`, `six_templet_E.docx` |
| **自然科學** (Science) | `S` | 5, 7, 8 | `S7N` | `pydocxTemplet_S.docx`, `S_M_demo.docx` |

> [!NOTE]
> **Special Rules for 7th Grade Entrants (7N / 桃園市七年級新生)**:
> - **Chapter 1~3 Standardized Template**: For 7N cohorts (e.g. `E7N`, `C7N`, `M7N`), Chapters 1~3 are automatically generated from dedicated templates at `resources/templates/special_cohorts/{cohort}_ch1_2_3_template.docx`. They provide standardized single-county Taoyuan text, official test date, cohort demographics (enrolled, tested, boy/girl, absence rate), and official assessment framework matrices for Table 3-1 / 3-2 (dynamic Jinja2 row looping for English, the 40-row curriculum taxonomy matrix for Chinese, and the 25-item 4x3 knowledge/cognitive matrix and 12-focus distribution for Math).
> - **Internal Item Codes**: Use grade digit `6` (e.g. `115-C6-A1-01`, `115-M6-D1-01`, `115-E6-A1-01`, `115-S6-A1-01`), representing Grade 6 curriculum scope.
> - **Exam Title & Preamble**: Dynamically formatted as `[年]年桃園市[科目]七年級新生學習能力檢測`.
> - **Math Table 6-1 Header**: Top-left merged cell displays `六`.
> - **Table 6-1 Title**: Standardized to `表6-1　整體學生及高分組、低分組學生於各評量向度之通過率`.
> - **Expert Fields**: Preserve learning-focus and qualitative-analysis slots for domain experts. With a supplied booklet, automatically fill the Chapter 4 question-image row; leave it blank when no booklet is supplied.

---

## Directory Structure

```text
SAAA-4-6-chapter/
├── SKILL.md                          # Main skill instruction file (this document)
├── README.md                         # Complete user manual
├── scripts/
│   ├── build_full_report.py          # Unified Master CLI entry point (Chapters 1~7 & assembly)
│   ├── build_from_source.py          # Booklet + Answer + Analysis end-to-end pipeline
│   ├── generate_reports.py           # Multi-subject generator (Ch 4, Ch 6)
│   ├── modules/                      # Modular chapter generators
│   │   ├── ch1_2_admin.py            # Ch 1 & Ch 2 generator (with Natural Science clause)
│   │   ├── ch3_framework.py          # Ch 3 assessment framework & Table 3-1
│   │   ├── ch5_performance.py        # Ch 5 distributions & Table 5-1-x 108 curriculum extractor
│   │   ├── ch7_appendix.py           # Ch 7 appendix injector
│   │   └── docx_assembler.py         # Full document assembler & OOXML injector
│   ├── qa_verify.py                  # Full verification covering Ch 1, 2, 3, 4, 5, 6, 7
│   ├── question_screenshots.py       # Booklet screenshot crop & text removal
│   ├── report_images.py              # Image layout & passage stitching
│   ├── image_layout.py               # Uniform font-scale image planner
│   ├── proper_names.py               # Proper name underline injector (縣市專名號)
│   └── extract_dimensions.py         # Answer sheet normalizer and Excel inspector
├── resources/
│   └── templates/                    # Standard Word templates for all subjects
└── references/
    ├── full_report_architecture.md   # Pipeline architecture & chapter modularity guide
    ├── screenshot_integration.md     # Booklet screenshot integration reference
    ├── subject_specifications.md     # Detailed data models and field schemas
    ├── word_reference_review.md      # Layout & reference inspection
    └── troubleshooting.md            # Word XML page break rules and error recovery
```

---

## Standard Workflow

### Mode 1: Full Report Assembly (Chapters 1~7)
Run the master CLI to generate all chapters or assemble them into the final report:
```bash
# Assemble complete report for Math Grade 4:
python scripts/build_full_report.py \
  --subject M \
  --grade 4 \
  --year 115 \
  --chapters all \
  --out-dir output/
```

### Mode 2: Modular Chapter Generation
Generate specific chapters according to project milestones:
```bash
# Generate Chapters 1 & 2 (Administration & Attendance):
python scripts/build_full_report.py --subject C --grade 3 --chapters 1,2 --out-dir output/

# Generate Chapters 4 & 6 (Expert review draft with booklet screenshots):
python scripts/build_full_report.py --subject E --grade 7N --chapters 4,6 --booklet-pdf booklet.pdf --out-dir output/

# Generate Chapter 5 (Ability distribution & Table 5-1-x with 108 Curriculum PDF extraction):
python scripts/build_full_report.py --subject S --grade 5 --chapters 5 --curriculum-pdf "自然科學領域.pdf" --out-dir output/
```

### Mode 3: Automated QA Verification
Verify generated Word documents against data standards:
```bash
python scripts/qa_verify.py --dir output/
```
The QA suite verifies:
- Chapter 1: 18-county union, proper name single underlines, and the exact clause: `檢測對象為國民中小學三年級至八年級學生，檢測科目包含國語文、數學、英語文和自然。`
- Chapter 2: Table 2-1 attendance matrix, column sums, and attendance rates.
- Chapter 3: Student statistics and Table 3-1 question distributions.
- Chapter 4: Question tables, item codes, images SHA256, 24pt line height, no black squares.
- Chapter 5: Figure 5-1 caption, Table 5-1-x 10-interval pass rate tables, and curriculum descriptions.
- Chapter 6: Table 6-1 multi-dimensional pass rates and cognitive intersections.
- Chapter 7: Appendix structure.
- Full Doc: Standard font `微軟正黑體` applied uniformly, no orphaned bookmarks.

---

## Detailed Guidelines & Subject Rules

Refer to the reference documentation for specialized topics:
- [full_report_architecture.md](./references/full_report_architecture.md): Overall pipeline architecture and chapter modularity.
- [screenshot_integration.md](./references/screenshot_integration.md): Booklet PDF cropping, question-number removal, passage stitching.
- [subject_specifications.md](./references/subject_specifications.md): Data schemas and field specifications for all subjects.
- [word_reference_review.md](./references/word_reference_review.md): Reference document inspection and review standards.
- [troubleshooting.md](./references/troubleshooting.md): Recovery steps for OOXML page breaks, tables, and encoding.
