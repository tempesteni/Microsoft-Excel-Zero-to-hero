# Excel Zero to Hero

**A complete, self-paced Microsoft Excel curriculum from absolute beginner to advanced user.**

Learn Excel through structured modules, 8 real sample workbooks you can open and practice in, quizzes, cheat sheets, and flashcards. Works for self-study, AI-assisted learning (NotebookLM, ChatGPT, Claude), and AI agent training pipelines. Comes with 5 study plan lengths (12 / 8 / 6 / 4 / 2 weeks): pick your pace, keep all the content.

---

## Table of Contents

1. [Start Here: How to Use This Package](#start-here-how-to-use-this-package)
2. [What's Inside: Full Content Map](#whats-inside-full-content-map)
3. [Workbook Basics: What a Spreadsheet Is and Why It Matters](#workbook-basics-what-a-spreadsheet-is-and-why-it-matters)
4. [Setup Guide: NotebookLM](#setup-guide-notebooklm)
5. [Setup Guide: ChatGPT / Claude / AI Chatbots](#setup-guide-chatgpt--claude--ai-chatbots)
6. [Setup Guide: AI Agents & RAG Systems](#setup-guide-ai-agents--rag-systems)
7. [Setup Guide: Anki Flashcards](#setup-guide-anki-flashcards)
8. [The 8 Modules: What You'll Learn](#the-8-modules-what-youll-learn)
9. [The Sample Workbooks](#the-sample-workbooks)
10. [Study Plans: 12 / 8 / 6 / 4 / 2 Weeks](#study-plans-12--8--6--4--2-weeks)
11. [Supporting Materials](#supporting-materials)
12. [Who This Is For](#who-this-is-for)
13. [Certification Paths](#certification-paths)
14. [Sources & References](#sources--references)
15. [Contributing](#contributing)
16. [License](#license)
17. [Author & Contributors](#author--contributors)

---

## Start Here: How to Use This Package

**Pick the path that matches how you want to learn:**

### Path A: I want to learn on my own (no AI)
1. Download or clone this repo.
2. Open [10-Study-Plan/00-zero-to-hero-study-plan.md](10-Study-Plan/00-zero-to-hero-study-plan.md) and pick one of 5 tracks (12 / 8 / 6 / 4 / 2 weeks).
3. Follow the schedule: read the module, do the exercises, open the matching sample workbook, take the quiz.
4. Check off skills in the [Progress Tracker](10-Study-Plan/progress-tracker.md).
5. Print the [Cheat Sheets](cheat-sheets/) and keep them next to your keyboard.

### Path B: I want AI to teach me (NotebookLM) <sup>Recommended</sup>
1. Go to [notebooklm.google.com](https://notebooklm.google.com) and create a notebook named `Excel Zero to Hero`.
2. Upload the module files (upload order table in [the guide below](#setup-guide-notebooklm)).
3. Start asking questions like:
   - *"Explain VLOOKUP to me like I'm 10 years old"*
   - *"Quiz me on PivotTables from Module 4 and grade my answers"*
   - *"I have 2 hours today on the 6-week track. Make me a minute-by-minute plan"*
   - *"I work in HR. Build me a headcount dashboard using Modules 3 and 4"*
4. Use **Audio Overview** for podcast-style summaries while commuting.

### Path C: I want ChatGPT or Claude to tutor me
1. Open a new chat and paste the setup prompt from [the guide below](#setup-guide-chatgpt--claude--ai-chatbots).
2. Paste one module at a time and let the AI drill you.
3. Use the follow-up prompts for quizzes, formula reviews, and dashboard feedback.

### Path D: I'm building an AI agent or training system
1. All content is structured Markdown with self-contained `##` sections, ready for RAG ingestion.
2. See [AI Agent Setup](#setup-guide-ai-agents--rag-systems) for chunking strategy, metadata schema, and sample code.
3. Quiz files = ready-made assessment content. Sample workbooks = realistic practice data.

---

## What's Inside: Full Content Map

### The Folder Structure

```
EXCEL Learn/
│
├── 01-Fundamentals/                      # Module 1 - Beginner
│   └── 01-excel-fundamentals.md          # Interface, navigation, data entry, references
│
├── 02-Formulas-Functions/                # Module 2 - Beginner+
│   └── 02-formulas-functions.md          # SUM, IF, SUMIF, text/date functions, IFERROR
│
├── 03-Formatting-Visualization/          # Module 3 - Beginner+
│   └── 03-formatting-visualization.md    # Number formats, conditional formatting, charts
│
├── 04-Data-Analysis/                     # Module 4 - Intermediate
│   └── 04-data-analysis.md               # Sorting, filtering, PivotTables, what-if analysis
│
├── 05-Advanced-Functions/                # Module 5 - Intermediate+
│   └── 05-advanced-functions.md          # INDEX/MATCH, XLOOKUP, dynamic arrays, LET/LAMBDA
│
├── 06-Power-Query-Pivot/                 # Module 6 - Advanced
│   └── 06-power-query-power-pivot.md     # Power Query ETL, data model, DAX basics
│
├── 07-VBA-Macros/                        # Module 7 - Advanced+
│   └── 07-vba-macros.md                  # VBA editor, Subs, loops, UserForms
│
├── 08-Real-World-Projects/               # Module 8 - All levels
│   └── 08-real-world-projects.md         # 8 hands-on projects with rubrics
│
├── 09-Sources-References/
│   └── sources-references-index.md       # Full bibliography, every link verified
│
├── 10-Study-Plan/
│   ├── 00-zero-to-hero-study-plan.md     # 5 tracks: 12 / 8 / 6 / 4 / 2 weeks
│   └── progress-tracker.md               # Checkbox mastery checklist + scoreboard
│
├── samples/                              # 8 PRACTICE WORKBOOKS + walkthroughs
│   ├── sample-01-data-entry-practice.xlsx # Module 1: messy entry + text-number cleanup
│   ├── sample-02-formula-challenges.xlsx  # Module 2: 12 graded formula tasks + key
│   ├── sample-03-dashboard-before-after.xlsx # Module 3: raw table vs styled dashboard
│   ├── sample-04-pivottable-source.xlsx  # Module 4: 300 rows + 6 pivot tasks
│   ├── sample-05-lookup-challenges.xlsx  # Module 5: XLOOKUP / INDEX-MATCH / LET tasks
│   ├── sample-06-powerquery-messy.xlsx   # Module 6: 3 messy monthly exports to clean
│   ├── 07-vba-toolkit/                   # Module 7: ready .bas macros + practice data
│   │   ├── mod_report_tools.bas
│   │   ├── mod_bulk_tools.bas
│   │   ├── practice-data.xlsx
│   │   └── README.md
│   ├── exercise-fix-this-workbook.xlsx   # DELIBERATELY BROKEN: find the 20 issues
│   ├── fix-this-workbook-challenge.md    # The 20-issue challenge + answer key
│   ├── walkthrough-01-data-entry.md      # What to notice in each sample
│   ├── walkthrough-02-formulas.md
│   ├── walkthrough-03-dashboard.md
│   ├── walkthrough-04-pivottables.md
│   ├── walkthrough-05-lookups.md
│   ├── walkthrough-06-powerquery.md
│   ├── build_all_samples.py              # Generators (openpyxl): study these too
│   └── build_fixme_workbook.py
│
├── exercises/
│   ├── sales-data-practice.xlsx          # 200 sales rows + 50 employee rows
│   └── create_dataset.py                 # Script that generated the dataset
│
├── cheat-sheets/                         # Printable quick references
│   ├── functions-cheat-sheet.md          # 50 functions with syntax and examples
│   ├── keyboard-shortcuts-cheat-sheet.md # 65+ shortcuts by category
│   ├── pivottable-cheat-sheet.md         # PivotTable reference card
│   └── power-query-cheat-sheet.md        # Top 20 M functions
│
├── quizzes/                              # Self-assessment with answer keys
│   ├── quiz-01-fundamentals.md           # 15 MC + 5 formula challenges = 70 pts each
│   ├── quiz-02-formulas.md
│   ├── quiz-03-visualization.md
│   ├── quiz-04-analysis.md
│   ├── quiz-05-advanced-lookup.md
│   ├── quiz-06-power-tools.md
│   ├── quiz-07-vba.md
│   └── quiz-08-projects.md
│
├── anki/
│   └── excel-flashcards.csv              # 120 cards: Front,Back,Category
│
├── curriculum-full.md                    # ONE-LINK AI edition: the whole course in one file
├── make_curriculum_full.py               # Regenerates the one-link edition
├── HOW-TO-USE-WITH-AI.md                 # Detailed AI usage guides + prompts
├── path.md                               # How this package was built (process doc)
├── CONTRIBUTORS.md                       # Who did what
└── README.md                             # This file
```

### Content Summary

| Category | Files | What You Get |
|----------|-------|-------------|
| **Learning Modules** | 9 files | Full curriculum: basics to VBA, plus 8 real-world projects |
| **Sample Workbooks** | 8 workbooks + 6 walkthroughs | Real practice files incl. the 20-issue fix-me workbook |
| **Quizzes** | 8 files | 120 MC + 40 formula challenges, full answer keys |
| **Cheat Sheets** | 4 files | 50 functions, 65+ shortcuts, PivotTable + Power Query cards |
| **Exercises** | 2 files | Practice dataset (250 rows) + generator script |
| **Flashcards** | 1 .csv | 120 Anki-compatible cards across categories |
| **Study Plans** | 2 files | 5 tracks (12/8/6/4/2 weeks) + progress tracker |
| **Sources** | 1 file | Full bibliography, every URL curl-verified |
| **AI Guides** | 1 file | Platform setup + copy-paste prompts |

---

## Workbook Basics: What a Spreadsheet Is and Why It Matters

New to Excel? This 2-minute primer covers the vocabulary the rest of this package uses.

**What is a workbook?**
Excel stores work in a **workbook** (one `.xlsx` file), which contains **worksheets** (tabs like Sheet1, Sheet2). Each worksheet is a grid of **cells**: the boxes where data lives. Every cell has an address (A1 = column A, row 1) and can hold a value or a **formula** (a live calculation like `=B2*C2` that updates when the data changes).

**Why spreadsheets exist**
A spreadsheet turns raw numbers into answers. A good workbook does three jobs:
- Stores data once, in a clean table (one row per record, one column per field)
- Calculates automatically with formulas, so numbers update when inputs change
- Shows the story with formatting and charts so decisions happen faster

**Workbook vs document vs deck**

| | Workbook | Document | Deck |
|---|---|---|---|
| How it's read | At your own pace, cell by cell | At your own pace | On a screen, paced by a speaker |
| Content | Numbers, tables, calculations | Paragraphs, details | One idea per slide |
| Best for | Analysis, tracking, reporting | Records, plans | Decisions, pitches, teaching |
| The quality test | Can I trust the formula? | Can I find the detail later? | Can the back row read it? |

**Why spreadsheets matter at work**
Nearly every role touches numbers, and Excel is where they live:
- **Finance and ops** run budgets, forecasts, and inventory on spreadsheets
- **Marketing and sales** track campaigns, pipelines, and commissions
- **HR and admin** manage headcount, schedules, and leave
- **Analysts** turn messy exports into decisions (and it is a career by itself)

**The anatomy of a workbook**
Most working files follow the same skeleton: an input sheet (raw data), calculation sheets (formulas), and an output sheet (dashboard or report) with charts and key numbers on top. Clean structure is what separates a trusted workbook from a haunted one. (Modules 1-3)

**How this package uses workbooks**
Everything in `samples/` is a real `.xlsx` file you can open, break, and rebuild. Work through each with its walkthrough, complete the graded challenges, and use `exercise-fix-this-workbook.xlsx` to train your eye on what to avoid. Reading real workbooks is how spreadsheet instincts are built.

---

## Setup Guide: NotebookLM

**What is NotebookLM?** Google's free AI research tool. Upload documents and it answers questions using ONLY your uploaded content.

**Step 1:** Go to [notebooklm.google.com](https://notebooklm.google.com) and sign in.

**Step 2:** Click **"Create new notebook"** and name it `Excel Zero to Hero`.

**Want one link instead of many uploads?** Add this single **Website** source and NotebookLM gets the entire course (every module, quiz, cheat sheet, walkthrough, and answer key):

```
https://raw.githubusercontent.com/tempesteni/Microsoft-Excel-Zero-to-hero/main/curriculum-full.md
```

(Built by `make_curriculum_full.py`. If a tool truncates very large pages, fall back to the per-file list below.)

**Step 3:** Upload your sources in this order:

| Upload Order | File | Why This Order |
|-------------|------|----------------|
| 1st | `10-Study-Plan/00-zero-to-hero-study-plan.md` | Gives the AI the curriculum map |
| 2nd-10th | `01-Fundamentals/` through `08-Real-World-Projects/` | The lessons, in sequence |
| 11th | `09-Sources-References/sources-references-index.md` | Reference context |
| 12th-19th | `quizzes/*.md` (8 files) | Assessment content |
| 20th-23rd | `cheat-sheets/*.md` (4 files) | Quick-reference context |

**Step 4:** Ask away. Copy-paste starters:

```
I'm a complete beginner. Explain VLOOKUP like I'm 10 years old. Give me 3 real exercises I can try in Excel right now.
```
```
Quiz me on Module 4 (Data Analysis). Ask 5 questions one at a time, wait for my answer, then grade me and tell me what to review.
```
```
I have 2 hours today and I'm on the 6-week track. What should I study today? Create a minute-by-minute schedule.
```
```
I work in [your field]. Show me how to build a [monthly report] workbook using Modules 2-4, step by step, with the exact formulas.
```

---

## Setup Guide: ChatGPT / Claude / AI Chatbots

### Quick Start Prompt (copy-paste this first)

```
I'm learning Microsoft Excel from a structured curriculum. I'll share modules one at a time. For each module:
1. Summarize the 5 most important concepts
2. Explain the hardest concept like I'm a beginner
3. Create 3 quiz questions and wait for my answers
4. Tell me what to focus on next

When I say "quiz me": give me 10 questions (mixed multiple-choice and "write the formula") and grade me.
When I say "practice": create a hands-on Excel exercise.
When I say "explain [topic]": break it down step by step.

Ready? Here's Module 1: Fundamentals.

[PASTE THE CONTENT OF 01-Fundamentals/01-excel-fundamentals.md HERE]
```

### Useful Follow-Up Prompts

```
I finished Module [X]. Quiz me with 10 questions, score me out of 100, and tell me which sections to reread.
```
```
I don't understand XLOOKUP. Explain 3 ways: like I'm 10, with a business example, and by comparison with VLOOKUP which I already know.
```
```
Here's my data description: [describe]. Give me the table layout, the formulas (with exact syntax), and 3 things that will break it.
```
```
Review this workbook design like an auditor: [paste structure/formulas]. Find the 5 highest-impact fixes for correctness and readability.
```
```
I have a job interview that requires Excel. Based on my materials: top 10 Excel interview questions, model answers, and a 2-day crash plan.
```

---

## Setup Guide: AI Agents & RAG Systems

Every file is structured Markdown with clear `##` hierarchy: compatible with RAG systems, vector databases, and AI tutoring agents.

### Chunking Strategy
```
Source format: Markdown (.md)
Split by: ## headings (each section is self-contained)
Chunk size: ~500-1500 tokens per chunk
Overlap: include the parent heading as context in each chunk
```

### Metadata Schema
```json
{
  "module_number": "01-08",
  "difficulty": "beginner | intermediate | advanced",
  "topic": "formulas | formatting | analysis | lookups | power_query | vba | projects",
  "content_type": "lesson | exercise | quiz | cheat_sheet | study_plan | walkthrough",
  "file_path": "relative/path/to/file.md"
}
```

### Embedding & Retrieval
```
Recommended model: text-embedding-3-small or all-MiniLM-L6-v2
Retrieval: Top-k=5 with metadata filter on module_number or difficulty
```

### Agent Modes

| Mode | What It Does | Input |
|------|-------------|-------|
| **Quiz** | Runs quiz files, grades answers | Quiz file + user answer |
| **Explain** | Breaks down concepts | Module file + topic |
| **Practice** | Generates exercises from sample data | Sample workbook structure + difficulty |
| **Progress** | Tracks completion, suggests next | progress-tracker.md + scores |
| **Reference** | Quick lookups | Cheat sheet + query |

### Sample Agent Code
```python
def answer_excel_question(question, user_level="beginner"):
    chunks = vector_db.search(query=question, top_k=5,
                              filter={"difficulty": user_level})
    context = "\n\n".join(c.text for c in chunks)
    prompt = f"""You are an Excel tutor. Use ONLY this context.
    Reference the module and section when possible.

    Context: {context}
    Student question: {question}
    Answer clearly, with an example, and suggest what to learn next."""
    return llm.generate(prompt)
```

---

## Setup Guide: Anki Flashcards

### How to Import

1. Open Anki (free: [apps.ankiweb.net](https://apps.ankiweb.net/)).
2. Click **File > Import** and select `anki/excel-flashcards.csv`.
3. Field mapping: Field 1 > Front, Field 2 > Back, Field 3 > Tags.
4. Click **Import**.

### What's in the Deck (120 cards)

| Category | Cards | Examples |
|----------|-------|---------|
| Formulas | 30 | SUMIF syntax, VLOOKUP parameters, IF nesting |
| Text & Date Functions | 22 | LEFT, TRIM, EOMONTH, NETWORKDAYS |
| Error Handling | 10 | #REF!, #DIV/0!, #N/A, IFERROR |
| Lookups | 8 | INDEX/MATCH vs VLOOKUP, XLOOKUP |
| Tables & Formatting | 8 | Structured references, conditional formatting |
| PivotTables | 8 | Field areas, grouping, value settings |
| What-If Analysis | 4 | Goal Seek, Scenario Manager, data tables |
| Power Tools | 20 | Merge vs Append, unpivot, M, DAX, CALCULATE |
| VBA & Shortcuts | 10+ | Recording macros, Sub basics, top shortcuts |

**Study tip:** 5 new cards per day. Anki's spaced repetition shows weak cards more often.

---

## The 8 Modules: What You'll Learn

### Module 1: Fundamentals (Beginner)
Interface tour, cell navigation, data entry, cell references (relative, absolute, mixed), file formats (.xlsx, .csv, .xlsb), saving and autosave.

### Module 2: Formulas & Functions (Beginner+)
SUM, AVERAGE, COUNT, IF, Nested IF, SUMIF/SUMIFS, COUNTIF/COUNTIFS, VLOOKUP, text functions (LEFT, RIGHT, MID, TRIM, CONCAT), date functions (TODAY, DATEDIF, EOMONTH), error handling (IFERROR).

### Module 3: Formatting & Visualization (Beginner+)
Number formats, conditional formatting (highlight rules, color scales, data bars, icon sets), Excel Tables, structured references, charts (Column, Bar, Line, Pie), chart formatting, Sparklines, Slicers.

### Module 4: Data Analysis (Intermediate)
Sorting (single/multi-level), AutoFilter, Advanced Filter, data validation (drop-down lists, custom formulas), PivotTables (creating, field areas, grouping, calculated fields, PivotCharts), Goal Seek, Scenario Manager, Data Tables, Subtotals.

### Module 5: Advanced Functions (Intermediate+)
INDEX/MATCH (single and multi-criteria), XLOOKUP (all parameters), OFFSET, INDIRECT, dynamic arrays (FILTER, SORT, UNIQUE, SEQUENCE), LET, LAMBDA, TEXTBEFORE, TEXTAFTER.

### Module 6: Power Query & Power Pivot (Advanced)
Importing data (CSV, web, databases), Query Editor, transformations (filter, split, merge, pivot/unpivot), Merge Queries (JOINs), Append Queries, M language, Data Model, relationships, DAX (calculated columns, measures, CALCULATE), KPIs, hierarchies.

### Module 7: VBA & Macros (Advanced+)
Recording macros, VBA Editor, Sub/Function, variables, MsgBox/InputBox, Range objects, loops (For Next, For Each), If/Select Case, error handling, working with sheets/workbooks, UserForms.

### Module 8: Real-World Projects (All levels)
8 complete projects applying all skills: Personal Budget Tracker, Sales Dashboard, Inventory Management System, Employee Schedule & Leave Tracker, Invoice Generator, Data Cleaning Pipeline, Financial Model (NPV, IRR, what-if), Automated Report Generator.

---

## The Sample Workbooks

Real `.xlsx` files you open and practice in, each with a walkthrough explaining what to notice.

| Workbook | Sheets | Trains | Highlight |
|----------|--------|--------|-----------|
| `sample-01-data-entry-practice.xlsx` | 3 | Module 1 | Messy real-world entry + text-number cleanup |
| `sample-02-formula-challenges.xlsx` | 3 | Module 2 | 12 graded tasks with expected results + answer key |
| `sample-03-dashboard-before-after.xlsx` | 2 | Module 3 | Identical numbers, raw table vs styled dashboard |
| `sample-04-pivottable-source.xlsx` | 2 | Module 4 | 300 analysis-ready rows + 6 pivot tasks |
| `sample-05-lookup-challenges.xlsx` | 4 | Module 5 | XLOOKUP vs INDEX/MATCH vs LET, with answer key |
| `sample-06-powerquery-messy.xlsx` | 4 | Module 6 | 3 messy monthly exports + 8 cleaning steps |
| `07-vba-toolkit/` | + `.bas` | Module 7 | Importable macro toolkit + practice data |
| `exercise-fix-this-workbook.xlsx` | 4 | All | Deliberately broken: find the 20 planted issues |

Every workbook uses fictional brands and data (YourCompany Inc). The generators (`build_all_samples.py`, `build_fixme_workbook.py`) are included: reading them is automation practice in itself.

---

## Study Plans: 12 / 8 / 6 / 4 / 2 Weeks

The content never changes, only the pace. Every track covers all modules, projects, samples, and quizzes.

| Track | Weeks | Hours/week | Best for |
|------|-------|-----------|----------|
| 12-Week Mastery | 12 | 6-8 | Deep, lasting skill |
| 8-Week Standard | 8 | 8-10 | Most learners with a day job |
| 6-Week Accelerated | 6 | 10-12 | Some Office experience |
| 4-Week Sprint | 4 | 12-15 | Fast upskilling for a job or project |
| 2-Week Crash | 2 | 18-25 | Deadlines: interviews, certification, a big review |

Full weekly schedules, daily session rhythm, and score gates: [10-Study-Plan/00-zero-to-hero-study-plan.md](10-Study-Plan/00-zero-to-hero-study-plan.md).

---

## Supporting Materials

### Quizzes (test yourself)
Each quiz: **15 multiple-choice** (3 pts each) + **5 formula challenges** (5 pts each) = **70 points**, with full answer keys and model solutions.

| Score | Meaning |
|-------|---------|
| 90%+ (63+) | Advance to the next module |
| 70-89% (49-62) | Review weak areas, then advance |
| Below 70% (<49) | Restudy the module before moving on |

### Cheat Sheets (print these)
50 functions with syntax and examples, 65+ keyboard shortcuts, a PivotTable reference card, and the top 20 Power Query M functions.

### Practice Data
`exercises/sales-data-practice.xlsx` contains 200 sales rows and 50 employee records with built-in cleaning issues (duplicates, extra spaces, blanks) for PivotTable, sorting, filtering, and cleaning practice.

---

## Who This Is For

- **Complete beginners** who have never opened Excel
- **Career switchers** needing Excel for data, finance, or operations roles
- **Students** preparing for MOS certification
- **Self-learners** who want structure instead of scattered videos
- **AI engineers** building Excel training agents or RAG systems
- **Teachers and trainers** adapting a ready-made curriculum
- **Professionals** who use Excel daily but want to level up to Power Query, DAX, and VBA

---

## Certification Paths

There are real certifications worth pairing with this curriculum. Every link below was verified live.

| Certification | Provider | What It Tests | Difficulty | Links |
|--------------|----------|---------------|------------|-------|
| MOS: Excel Associate | Microsoft / Certiport | Core functionality, formulas, tables, charts | Beginner-Intermediate | [Program page](https://www.microsoft.com/en-us/learning/mos-certification.aspx) · [Register](https://www.certiport.com/portal/pages/mos/Overview.aspx) |
| MOS: Excel Expert | Microsoft / Certiport | Advanced formulas, PivotTables, macros | Advanced | [Program page](https://www.microsoft.com/en-us/learning/mos-certification.aspx) · [Register](https://www.certiport.com/portal/pages/mos/Overview.aspx) |
| Microsoft 365 Fundamentals (MS-900) and other M365 certifications | Microsoft Learn | Broad Microsoft 365 literacy | Beginner | [Browse certifications](https://learn.microsoft.com/en-us/credentials/certifications/) |
| Excel Skills for Business (Macquarie University) | Coursera | Comprehensive Excel for business | All levels | [Find it on Coursera](https://www.coursera.org/search?query=excel%20skills%20for%20business) |
| Financial Modeling & Valuation Analyst (FMVA) | Corporate Finance Institute | Financial modeling in Excel | Advanced | [Enroll](https://corporatefinanceinstitute.com/certifications/fmva/) |
| Course certificates: data & spreadsheets | edX | Analysis and spreadsheets | All levels | [Search courses](https://www.edx.org/search?q=excel) |
| Certificates of completion | LinkedIn Learning | Practical Excel tracks | All levels | [LinkedIn Learning](https://www.linkedin.com/learning/) |

**Booking MOS exams:**
- Exams are delivered through Certiport authorized test centers: register at [certiport.com](https://www.certiport.com/).
- Exam codes change between generations (Excel exams have carried codes like MO-200 and MO-201). Always confirm the current code and objectives on the official pages above before booking.
- Note: Microsoft's online credential catalog currently returns no results for Office specialist exams (MOS is administered via Certiport), so use the MOS program page and the Certiport link.

**How this package prepares you:**
- Modules 1-3 cover the Associate core (formulas, formatting, charts), Modules 4-7 cover the Expert territory (analysis, advanced functions, Power Query, VBA).
- Book a certification exam only after passing all 8 quizzes at 63+/70 (90%+) and completing at least 5 of the 8 projects in Module 8.
- Use the checklist in `10-Study-Plan/progress-tracker.md` as your readiness gate.

---

## Sources & References

This curriculum draws from a wide, **fully verified** bibliography (every URL curl-tested before release):

- **Official:** Microsoft Support (Excel hub), Microsoft Learn, Microsoft Create
- **Tutorial craft:** ExcelJet, Chandoo, Contextures, Spreadsheeto, Excel Off The Grid, Peltier Tech, Curbal
- **YouTube:** Leila Gharani, ExcelIsFun, and more (with search links where videos move)
- **Books:** Excel Bible (Walkenbach), Power Programming with VBA (Alexander), and more with ISBNs
- **Communities:** r/excel, Microsoft Tech Community, Super User, Stack Overflow

Full bibliography with ISBNs and attribution notes: [Sources & References Index](09-Sources-References/sources-references-index.md).

---

## Contributing

Found an error? Want to add content?

1. Fork this repo.
2. Make your changes (everything is Markdown; sample workbooks regenerate via the included scripts).
3. Submit a pull request.

Contributions welcome: new modules, more sample workbooks, additional challenges, translations, corrections.

---

## License

This project is licensed under the **Creative Commons Attribution 4.0 International License (CC BY 4.0)**.

- **Share** - copy and redistribute in any format
- **Adapt** - remix, transform, and build upon the material
- **Attribution** - give credit and link to this repo

---

## Author & Contributors

Created by **[tempesteni](https://github.com/tempesteni)** with **[Hermes Agent](https://hermes-agent.nousresearch.com)** by Nous Research.

| Contributor | Roles |
|-------------|-------|
| tempesteni | Concept and curriculum direction, scope decisions, review, corrections, quality control, publishing decisions |
| Hermes Agent | Research and compilation, module authoring, quiz/flashcard/cheat sheet generation, sample workbook generation, link verification and repair, content and security verification, documentation |

**Stats:** 53 files | ~75,000 words | 8 modules + 8 real-world projects | 8 sample workbooks | 8 quizzes | 120 flashcards | 5 study tracks | estimated study time 40-120 hours depending on track | prerequisites: basic computer literacy, access to Excel (2019, 2021, or Microsoft 365).
