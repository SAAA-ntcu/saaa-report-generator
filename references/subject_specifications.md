# SAAA-4-6-chapter: Subject Specifications Reference

This document details the data structures, template variables, column mappings, and layout specifications for generating Chapter 4 and Chapter 6 reports across all four subjects.

---

## 1. Subject Overview & Cohort Naming

| Subject | Code | Chinese Name | Standard Grades | Special Cohorts | Primary Question Types |
| :--- | :--- | :--- | :--- | :--- | :--- |
| **國語文** | `C` | 國語 | 3, 4, 5, 6, 7, 8 | `C7N` | 4-choice single answer (1~4 or A~D) |
| **數學** | `M` | 數學 | 3, 4, 5, 6, 7, 8 | `M7N` | 4-choice single answer (1~4 or A~D) |
| **英語文** | `E` | 英文 / 英語文 | 4, 5, 6, 7, 8 | `E7N` | 聽力測驗 + 閱讀測驗 (1~4 or A~D) |
| **自然科學** | `S` | 自然 / 自然科學 | 5, 7, 8 | `S7N` | 4-choice single answer, supports invalid items |

---

## 2. Chapter 4 (肆、試題分析與教學建議)

### 2.1 Context Variables (docxtpl)
All subject templates utilize the following variables in jinja2 context:

| Variable | Type | Description | Example |
| :--- | :--- | :--- | :--- |
| `year` | `str` | Academic school year | `"115"` |
| `subject` | `str` | Chinese subject name | `"國語"`, `"數學"`, `"英文"`, `"自然"` |
| `grade_cn` | `str` | Chinese numeral for grade | `"三"`, `"四"`, `"五"`, `"六"`, `"七"`, `"八"` |
| `grade` | `str` | Grade code / cohort | `"7"`, `"7N"`, `"5"` |
| `cronbach_s_alpha` | `str` | Reliability index (2 decimals) | `"0.83"` |
| `valid_sample` | `str` | Number of test takers | `"21,400"` |
| `domain_legend` | `str` | Semicolon-separated legend | `"A：數與量；B：空間與形狀..."` |
| `questions` | `list[dict]` | List of question dictionaries | See question structure below |

### 2.2 Question Dictionary Fields
Each entry in `questions` includes:
- `number`: Question sequence number (1, 2, ... N).
- `code`: Item identification code, e.g. `115-C7-A1-01`, `115-M7-A1-01`, `115-E7-A1-01`, `115-S5-A1-01`.
- `ans`: Correct answer text (`A`, `B`, `C`, `D` or `①`~`④`, or `不予計分`).
- `domain`: Content domain name (e.g. `形音知識`, `數與量`, `聽力`, `物質與能量`).
- `cognitive`: Cognitive level or sub-domain name (e.g. `概念理解`, `程序執行`, `解題思考`, `辨識句意`).
- `pass_rate`: Percentage pass rate formatted as integer without `%` sign (e.g. `"68"`).
- `discrimination`: Item discrimination index formatted to 2 decimal places (e.g. `"0.45"`).
- `hi_pass`, `lo_pass`: High-group and low-group pass rates (e.g. `"88"`, `"42"`).
- `h1`~`h4`: Option labels with asterisk attached to correct answer (e.g. `h1="A*"`, `h2="B"`, `h3="C"`, `h4="D"`).
- `all_1`~`all_4`, `all_9`: Option choice proportions for all students (2 decimals). `9` represents missing/invalid responses.
- `hi_1`~`hi_4`, `hi_9`: Choice proportions for the high-scoring group (upper 27%).
- `lo_1`~`lo_4`, `lo_9`: Choice proportions for the low-scoring group (lower 27%).
- `w1_ans`~`w3_ans`, `w1_rate`~`w3_rate`: The three incorrect distractor options sorted in order of option sequence with percentage choice rates.

### 2.3 Subject-Specific Layout Rules (Chapter 4)
- **Chinese (`C`)**:
  - Questions are grouped and ordered by domain prefix: Category A (`一、語文知識：` [A~F]) then Category B (`二、閱讀理解：` [G~J]).
  - Subheadings `(一)形音知識` must be kept on the same page as the first question table of that section.
  - Subsequent question tables under the same section receive an XML page break (`w:pageBreakBefore`).
  - Major headings `二、閱讀理解：` and subheadings `([二三四五六七八九十])` receive `w:pageBreakBefore`.
  - Specific paragraph indents: `評量的` gets left indent 840 twips / 3.5 char and first line indent 480 twips / 2.0 char.
- **Math (`M`), English (`E`), Science (`S`)**:
  - Continuous question sequence from Q1 to QN.
  - Every question table starting from Q2 receives `w:pageBreakBefore`.
  - Empty paragraphs between tables are automatically purged to prevent stray blank pages.
- **Science Invalid Items**:
  - If `is_invalid == True` (unscored questions), rates are replaced with `"-"`, and answer is displayed as `"不予計分"`.

---

## 3. Chapter 6 (陸、整體教學建議 / 表6-1 各評量向度通過率)

### 3.1 Chinese (`C`)
- **Template**: `C_demo.docx`.
- **Columns (5)**:
  1. `向度名稱` (Domain Name)
  2. `題號` (Question Numbers, e.g. `1、2、5`)
  3. `高分組` (High Group Pass Rate, 2 decimals)
  4. `全體` (All Students Pass Rate, 2 decimals)
  5. `低分組` (Low Group Pass Rate, 2 decimals)
- **Rows**: 2 header rows + 10 domain rows + 1 overall row (`整體`) = 13 rows total.
  (If original template has 11 domain rows, row 12 `國學知識` is removed).
- **Paragraphs**:
  - P1: Lists all domain names: `國語文X年級學力檢測，以「選擇題型」的方式，檢測「...」等10項能力。`
  - P2: Overall pass rate: `國語文X年級學力檢測全體學生通過率是XX%，各向度施測結果統計如下表6-1所示：`

### 3.2 Math (`M`)
- **Template**: `S_M_demo.docx`.
- **Columns (13)**:
  - Col 0: Grade (`三`~`八` / `七`)
  - Cols 1~3: `概念理解` (題號, 分組, 通過率)
  - Cols 4~6: `程序執行` (題號, 分組, 通過率)
  - Cols 7~9: `解題思考` (題號, 分組, 通過率)
  - Cols 10~12: `整體` (題數, 分組, 通過率)
- **Rows**: 2 header rows + 3 rows per content domain (高, 全, 低) + 3 overall summary rows.
  - Grade 3~6: 5 content domains = 2 + 15 + 3 = 20 rows.
  - Grade 7~8 / 7N: 4 content domains = 2 + 12 + 3 = 17 rows.
- **Empty Intersections**: Merged 3x3 block with diagonal border (`w:tl2br`).

### 3.3 English (`E`)
- **Template**: `six_templet_E.docx`.
- **Columns (6)**:
  1. `大向度` (Main Domain: 聽力 / 閱讀)
  2. `次向度` (Sub-domain: 辨識句意, 基本常用字詞...)
  3. `題號` (Question Numbers)
  4. `高分組` (High Group Pass Rate)
  5. `全體` (All Students Pass Rate)
  6. `低分組` (Low Group Pass Rate)
- **Merging**: Column 0 is vertically merged across all `聽力` rows, and merged across all `閱讀` rows.
- **Overall Row**: Bottom row merges Cols 0 and 1 with label `整體`.

### 3.4 Science (`S`)
- **Template**: `S_M_demo.docx` (reconstructed to 16 columns).
- **Columns (16)**:
  - Col 0: Grade
  - Cols 1~3: `記憶`
  - Cols 4~6: `理解`
  - Cols 7~9: `應用`
  - Cols 10~12: `分析`
  - Cols 13~15: `整體` (題數, 分組, 通過率)
- **Rows**: 2 header rows + 3 rows per content domain + 3 overall rows.
- **Layout**: 12pt Chinese font with 2-by-2 character wrapping for domain names.

---

## 4. Special Rules for Taoyuan 7th Grade Entrant Cohorts (桃園市七年級新生 7N)

For assessments administered to new 7th-grade entrants in Taoyuan (cohorts `C7N`, `M7N`, `E7N`, `S7N`), the following specialized rules apply:

### 4.1 Assessment Name & Report Preamble
The assessment name in the Chapter 4 preamble dynamically formats as:
- **國語文**: `{year}年桃園市國語文七年級新生學習能力檢測`
- **數學**: `{year}年桃園市數學七年級新生學習能力檢測`
- **英語文**: `{year}年桃園市英語文七年級新生學習能力檢測`
- **自然科學**: `{year}年桃園市自然科學七年級新生學習能力檢測`
*(Do NOT label as standard elementary or junior high school assessments).*

### 4.2 Item Identification Codes (關鍵點 A)
- **Cohort Code**: Files, arguments, and batch processes use `C7N`, `M7N`, `E7N`, `S7N`.
- **Internal Item Code**: Because the assessment tests 6th-grade elementary curriculum, the internal question code within Chapter 4 tables uses grade code **`6`**:
  - 國語文: `{year}-C6-{domain_code}1-{num:02d}` (e.g. `115-C6-A1-01`)
  - 數學: `{year}-M6-{content_code}{cognitive_code}-{num:02d}` (e.g. `115-M6-D1-01`)
  - 英語文: `{year}-E6-{sub_code}-{num:02d}` (e.g. `115-E6-A1-01`)
  - 自然科學: `{year}-S6-{domain_code}{cognitive_code}-{num:02d}` (e.g. `115-S6-A1-01`)

### 4.3 Chapter 6 Table 6-1 Formats
- **Table Title Paragraph**: Standardized across all subjects to:
  `表6-1　整體學生及高分組、低分組學生於各評量向度之通過率` (using full-width space `\u3000`).
- **Math Table 6-1 Header**: Top-left merged cell `cell(0, 0)` displays **`六`** (Grade 6), NOT `七`.
- **Chinese Chapter 6 Preamble**: Formatted as `國語文七年級新生學力檢測...` and `七年級新生國語文學力檢測全體學生通過率是...`.

### 4.4 Boundary with Domain Experts
The automated engine strictly preserves placeholders for domain experts:
- **Chapter 4 Row 2/3 (學習重點與次向度)**: Kept for domain experts to fill.
- **Chapter 4 Row 4 (試題題目呈現)**: Automatically insert the corresponding booklet image when a PDF or image manifest is supplied; otherwise preserve the blank slot. Shared reading material is inserted once before the first related item, in source-page segments. See [screenshot_integration.md](screenshot_integration.md).
- **Chapter 4 Row 9 (單題質性分析與教學建議)**: Left for experts.
- **Chapter 6 Post-Table 6-1 In-Depth Pedagogical Suggestions**: Written by experts.
- **English Table 6-1 Placement**: Adjusted manually after qualitative content is written.

---

## 5. Specifications for Chapters 1, 2, 3, 5, 7 (Full Report Suite)

### 5.1 Chapter 1 (壹、背景與目的)
- **Background Counties**: Unified 18-county union across all participation rows (基隆市至金門縣，排除臺北市、彰化縣、宜蘭縣、連江縣).
- **Proper Names Underline**: Center name (`國立臺中教育大學測驗統計與適性學習研究中心`) and each of the 18 counties carry a single underline (`w:u val="single"`).
- **Annual Cohort Natural Science Addition**:
  The concluding sentence of Paragraph 2 for annual cohorts (Grades 3~8) strictly follows:
  `檢測對象為國民中小學三年級至八年級學生，檢測科目包含國語文、數學、英語文和自然。`

### 5.2 Chapter 2 (貳、參與縣市與檢測規模)
- **Table 2-1 Title**: `表2-1　{year}年度{stage}學生學習能力檢測{subject}{grade_cn}年級縣市參與人數摘要表`
- **Columns (5)**: `縣市別`, `報考學生人數`, `到考人數`, `缺考人數`, `到考率(%)`.
- **Data Source**: Counties marked with `V` in `115年縣市學力檢測參與狀況_1150108.xlsx` and numbers from `到考人數.xlsx`.
- **Summary Row**: `總計` with column totals and weighted overall attendance rate (2 decimal places).

### 5.3 Chapter 3 (參、檢測對象與檢測工具)
- **Cohort Statistics**: Displays enrolled count (`應考名冊羅列{N}人`), tested count (`實際施測人數{N}人`), male/female counts, and absence rate.
- **Table 3-1 Structure**: `表3-1　{subject}{grade_cn}年級評量向度之分布題數`
  - Columns (4): `主要評量向度`, `評量次向度`, `測驗題號`, `題數`.
  - Content domains group sub-dimensions and list test item numbers, ending with a merged `總計` row.

### 5.4 Chapter 5 (伍、整體學力表現分析)
- **Section 1**: Background variable ability distribution analysis with `圖5-1` histogram/curve.
- **Section 2**: Student performance across learning focuses:
  - Tables 5-1-1 through 5-1-N: 10 pass-rate intervals (`91~100`, `81~90`, `71~80`, `61~70`, `51~60`, `41~50`, `31~40`, `21~30`, `11~20`, `0~10`).
  - Columns: `答對率%`, `題號`, `學習重點` (Curriculum code + official descriptor extracted from 108 Curriculum PDF).
  - Alternating row background shading `#DDEBF7`.

### 5.5 Chapter 7 (柒、附錄)
- Section Title: `柒、附錄`
- Contains:
  - `附錄一` Test booklets & answer keys
  - `附錄二` Item inquiry notices and responses
  - Committee lists & administrative appendices.
