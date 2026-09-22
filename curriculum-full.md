# Excel Zero to Hero: Full Curriculum (Single-File AI Edition)

> One link, the complete course. Built so AI tools can see everything at once.
>
> **How to use:** in NotebookLM add this page as a single **Website** source. In ChatGPT/Claude, paste the link and ask your questions. In a RAG system, chunk by `## FILE:` headings.
>
> **Companion binaries:** the practice workbooks/decks live as separate files in the GitHub repo (this file keeps their walkthroughs and answer keys).
>
> **Regenerated from source files** by `make_curriculum_full.py`. Prefer per-file sources? Use the raw links listed in the README.

## Table of Contents

- `README.md`
- `01-Fundamentals/01-excel-fundamentals.md`
- `02-Formulas-Functions/02-formulas-functions.md`
- `03-Formatting-Visualization/03-formatting-visualization.md`
- `04-Data-Analysis/04-data-analysis.md`
- `05-Advanced-Functions/05-advanced-functions.md`
- `06-Power-Query-Pivot/06-power-query-power-pivot.md`
- `07-VBA-Macros/07-vba-macros.md`
- `08-Real-World-Projects/08-real-world-projects.md`
- `09-Sources-References/sources-references-index.md`
- `10-Study-Plan/00-zero-to-hero-study-plan.md`
- `10-Study-Plan/progress-tracker.md`
- `HOW-TO-USE-WITH-AI.md`
- `cheat-sheets/functions-cheat-sheet.md`
- `cheat-sheets/keyboard-shortcuts-cheat-sheet.md`
- `cheat-sheets/pivottable-cheat-sheet.md`
- `cheat-sheets/power-query-cheat-sheet.md`
- `quizzes/quiz-01-fundamentals.md`
- `quizzes/quiz-02-formulas.md`
- `quizzes/quiz-03-visualization.md`
- `quizzes/quiz-04-analysis.md`
- `quizzes/quiz-05-advanced-lookup.md`
- `quizzes/quiz-06-power-tools.md`
- `quizzes/quiz-07-vba.md`
- `quizzes/quiz-08-projects.md`
- `samples/07-vba-toolkit/README.md`
- `samples/fix-this-workbook-challenge.md`
- `samples/walkthrough-01-data-entry.md`
- `samples/walkthrough-02-formulas.md`
- `samples/walkthrough-03-dashboard.md`
- `samples/walkthrough-04-pivottables.md`
- `samples/walkthrough-05-lookups.md`
- `samples/walkthrough-06-powerquery.md`
- `anki/excel-flashcards.csv`
- `CONTRIBUTORS.md`
- `path.md`

---


## FILE: README.md

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


## FILE: 01-Fundamentals/01-excel-fundamentals.md

# Module 1: Microsoft Excel Fundamentals

> **Level:** Beginner | **Estimated Time:** 2–3 hours | **Prerequisites:** A computer with Microsoft Excel installed (or access to Excel Online)

---

## Table of Contents

1. [What is Excel?](#1-what-is-excel)
2. [History and Versions](#2-history-and-versions)
3. [The Excel Interface](#3-the-excel-interface)
4. [Workbook vs Worksheet](#4-workbook-vs-worksheet)
5. [Cells, Rows, and Columns](#5-cells-rows-and-columns)
6. [Basic Navigation](#6-basic-navigation)
7. [Entering Data](#7-entering-data)
8. [Cell References](#8-cell-references)
9. [Basic Data Types and Formats](#9-basic-data-types-and-formats)
10. [Saving and File Formats](#10-saving-and-file-formats)
11. [Excel Online vs Desktop](#11-excel-online-vs-desktop)
12. [Practice Exercises](#12-practice-exercises)
13. [Video References](#13-video-references)
14. [Sources](#14-sources)

---

## 1. What is Excel?

**Microsoft Excel** is a spreadsheet application developed by Microsoft. It is used to organize, format, and calculate data with formulas. Excel is one of the most widely used software applications in the world, with hundreds of millions of users across business, education, science, and personal use.

### Key Capabilities

- **Data Organization:** Store and structure data in rows and columns
- **Calculations:** Perform mathematical, statistical, and financial calculations using formulas and functions
- **Data Analysis:** Sort, filter, pivot, and summarize large datasets
- **Visualization:** Create charts, graphs, and conditional formatting to visualize data
- **Automation:** Use macros (VBA) and Power Query to automate repetitive tasks
- **Collaboration:** Share workbooks and co-author in real time (Microsoft 365)

---

## 2. History and Versions

### Timeline of Major Excel Versions

| Year | Version | Key Feature |
|------|---------|-------------|
| 1985 | Excel 1.0 (Mac) | First release, for Apple Macintosh |
| 1987 | Excel 2.0 (Windows) | First Windows version |
| 1993 | Excel 5.0 | Introduced VBA (Visual Basic for Applications) |
| 1997 | Excel 97 | Improved VBA, toolbars, 65,536 rows |
| 2000 | Excel 2000 | HTML saving, improved PivotTables |
| 2003 | Excel 2003 | XML support, improved lists |
| 2007 | Excel 2007 | **Ribbon interface**, .xlsx format, 1,048,576 rows |
| 2010 | Excel 2010 | Sparklines, Slicers, improved PivotTables |
| 2013 | Excel 2013 | Flash Fill, Timeline filters, new chart types |
| 2016 | Excel 2016 | Power Query built-in, new chart types (Treemap, Sunburst) |
| 2019 | Excel 2019 | TEXTJOIN, IFS, SWITCH, funnel charts, 3D maps |
| 2021 | Excel 2021 | XLOOKUP, XMATCH, dynamic arrays, LET function |
| 365 | Microsoft 365 | Continuous updates, Copilot AI, LAMBDA, co-authoring |

### Current Versions

- **Microsoft 365 (Subscription):** Always up-to-date with the latest features, cloud storage (OneDrive), and AI-powered Copilot
- **Excel 2024 (One-time purchase):** Perpetual license, feature set frozen at release
- **Excel 2021 (One-time purchase):** Previous perpetual license version
- **Excel Online (Free):** Browser-based version at [office.com](https://www.office.com) - limited features but free with a Microsoft account

---

## 3. The Excel Interface

When you open Excel, you see the following interface components:

### 3.1 The Ribbon

The **Ribbon** is the toolbar at the top of the Excel window, organized into tabs. It replaced the old menu system starting with Excel 2007.

**Main Tabs:**

| Tab | Purpose |
|-----|---------|
| **Home** | Clipboard, Font, Alignment, Number, Styles, Cells, Editing |
| **Insert** | Tables, Illustrations, Charts, Sparklines, Filters, Links, Text |
| **Page Layout** | Themes, Page Setup, Scale to Fit, Sheet Options |
| **Formulas** | Function Library, Defined Names, Formula Auditing, Calculation |
| **Data** | Get External Data, Connections, Sort & Filter, Data Tools |
| **Review** | Proofing, Language, Comments, Changes (Protection) |
| **View** | Workbook Views, Show, Zoom, Window, Macros |

**Contextual Tabs** appear when you select specific objects (e.g., Chart Design tab when a chart is selected).

### 3.2 Quick Access Toolbar

Located at the very top-left of the window (above or below the Ribbon). Customize it by clicking the dropdown arrow to add frequently used commands like:
- Save
- Undo (Ctrl+Z)
- Redo (Ctrl+Y)
- Quick Print

### 3.3 Name Box

The **Name Box** is the small field to the left of the Formula Bar. It shows:
- The **cell address** of the currently selected cell (e.g., `A1`)
- Any **named range** if the selected cell belongs to one
- You can also type a cell address (e.g., `G50`) and press **Enter** to jump directly to that cell

### 3.4 Formula Bar

The **Formula Bar** displays the contents (value or formula) of the currently selected cell. When you click into the Formula Bar (or press F2), you can edit the cell's contents. It shows:
- Plain text values as-is
- Formulas starting with `=` (e.g., `=SUM(A1:A10)`)

### 3.5 Sheet Tabs

Located at the bottom-left of the window. Each **sheet tab** represents a separate worksheet within the workbook.

- Click a tab to switch to that worksheet
- Right-click a tab to rename, delete, move, copy, or change tab color
- Click the `+` button to add a new worksheet
- Use the navigation arrows `< >` to scroll through tabs when there are many

### 3.6 Status Bar

The bar at the very bottom of the Excel window. It shows:
- **Mode:** Ready, Edit, Enter (current state)
- **Caps Lock / Num Lock** indicators
- **Quick calculations:** Select a range of numbers and the Status Bar shows SUM, AVERAGE, and COUNT automatically (right-click the Status Bar to customize)
- **Zoom slider** on the right side to magnify the view

---

## 4. Workbook vs Worksheet

### Workbook

A **workbook** is the entire Excel file (`.xlsx`). Think of it as a **book**:
- It can contain **multiple worksheets** (like pages in a book)
- A new workbook opens with 1 worksheet by default (you can change this in Options → General → "Include this many sheets")
- The workbook name appears in the **title bar** at the top

### Worksheet

A **worksheet** (also called a **sheet** or **spreadsheet**) is a single page within a workbook:
- It is a grid of **1,048,576 rows** × **16,384 columns** (columns A through XFD)
- Each worksheet has its own tab at the bottom
- You can have multiple worksheets for related data (e.g., "January Sales", "February Sales", "Summary")

### Practical Example

```
Workbook: "2026 Budget.xlsx"
├── Sheet1: "Income"
├── Sheet2: "Expenses"
├── Sheet3: "Summary"
└── Sheet4: "Charts"
```

---

## 5. Cells, Rows, and Columns

### 5.1 The Grid Structure

| Element | Description | Count |
|---------|-------------|-------|
| **Cell** | The intersection of a row and a column - the basic unit of a spreadsheet | Over 17 billion per sheet |
| **Row** | A horizontal line of cells, numbered 1 to 1,048,576 | 1,048,576 |
| **Column** | A vertical line of cells, lettered A to XFD | 16,384 |

### 5.2 Cell Addressing (A1 Notation)

Every cell has a unique **address** (also called a **cell reference**) formed by its column letter + row number:

- **A1** = Column A, Row 1 (top-left cell)
- **B3** = Column B, Row 3
- **Z100** = Column Z, Row 100
- **AA1** = Column AA (the 27th column)

### 5.3 Selecting Cells

| Action | How To |
|--------|--------|
| Select a single cell | Click on it |
| Select a range | Click and drag from one cell to another (e.g., A1:C5) |
| Select an entire row | Click the row number (e.g., `5`) |
| Select an entire column | Click the column letter (e.g., `D`) |
| Select multiple non-adjacent cells | Hold **Ctrl** and click each cell |
| Select all cells | Press **Ctrl+A** or click the corner button (top-left intersection of row/column headers) |

---

## 6. Basic Navigation

### 6.1 Keyboard Shortcuts for Navigation

| Action | Windows Shortcut | Mac Shortcut |
|--------|-----------------|--------------|
| Move one cell up/down/left/right | Arrow Keys | Arrow Keys |
| Move to the beginning of the row | Home | Fn + ← |
| Move to cell A1 | Ctrl + Home | Fn + ⌘ + ← |
| Move to the last used cell | Ctrl + End | Fn + ⌘ + → |
| Move down one screen | Page Down | Fn + ↓ |
| Move up one screen | Page Up | Fn + ↑ |
| Go to a specific cell | Ctrl + G (or F5) | Fn + ⌘ + G |
| Jump to edge of data region | Ctrl + Arrow Key | ⌘ + Arrow Key |
| Go to next sheet | Ctrl + Page Down | ⌘ + Fn + ↓ |
| Go to previous sheet | Ctrl + Page Up | ⌘ + Fn + ↑ |

### 6.2 Using the Name Box to Navigate

1. Click on the **Name Box** (left of the Formula Bar)
2. Type a cell address (e.g., `G500`)
3. Press **Enter**
4. Excel jumps directly to that cell

### 6.3 Using Go To (Ctrl+G)

1. Press **Ctrl+G** (or **F5**)
2. The "Go To" dialog box appears
3. Type a cell reference or range (e.g., `A1:Z100`)
4. Click **OK**

You can also click **Special...** to go to specific types of cells (blanks, formulas, comments, etc.).

---

## 7. Entering Data

### 7.1 Entering Text

1. Click on a cell
2. Start typing
3. Press **Enter** to move down, **Tab** to move right, or **arrow keys** to move in any direction
4. Press **Escape** to cancel entry

**Tips:**
- Text aligns to the **left** by default
- If your text is longer than the column width, it will overflow into the next cell(s) - unless those cells contain data
- To start a new line **within** a cell, press **Alt+Enter**

### 7.2 Entering Numbers

1. Click on a cell
2. Type the number
3. Press **Enter**

**Tips:**
- Numbers align to the **right** by default
- Do not type currency symbols, commas, or percentage signs if you want to apply number formatting later
- To enter a fraction like ½, type `0 1/2` (zero, space, 1/2)
- To enter a number as text (e.g., a zip code starting with 0), type an apostrophe first: `'01234`

### 7.3 Entering Dates

1. Click on a cell
2. Type a date in a recognized format (e.g., `1/15/2026`, `Jan 15, 2026`, `2026-01-15`)
3. Press **Enter**

**Tips:**
- Excel stores dates as **serial numbers** (January 1, 1900 = 1)
- You can format dates in many display formats (right-click → Format Cells → Date)
- Use shortcuts: **Ctrl+;** enters today's date, **Ctrl+Shift+;** enters the current time

### 7.4 Entering Formulas

1. Click on a cell
2. Type `=` to begin a formula
3. Enter the formula (e.g., `=A1+B1`)
4. Press **Enter**

**Example:** To add the values in cells A1 and B1:
```
=A1+B1
```

See [Module 2: Formulas & Functions](../02-Formulas-Functions/02-formulas-functions.md) for detailed formula coverage.

### 7.5 AutoFill

Excel can automatically fill a series of data:

1. Enter a starting value (e.g., `1` in A1, `2` in A2)
2. Select both cells (A1:A2)
3. Drag the **fill handle** (the small square at the bottom-right corner of the selection) down
4. Excel continues the pattern: `3`, `4`, `5`, ...

AutoFill works with:
- Numbers (1, 2, 3, ... or 2, 4, 6, ...)
- Days of the week (Mon, Tue, Wed, ...)
- Months (Jan, Feb, Mar, ...)
- Custom lists

---

## 8. Cell References

Cell references are the foundation of Excel formulas. Understanding how they behave when copied is critical.

### 8.1 Relative References (Default)

A **relative reference** changes when you copy the formula to another cell.

**Example:**
```
Cell C1: =A1+B1
```
If you copy C1 down to C2, the formula automatically becomes:
```
Cell C2: =A2+B2
```

The reference "shifts" relative to the new position.

### 8.2 Absolute References ($A$1)

An **absolute reference** does NOT change when you copy the formula. Use `$` signs to lock the column and/or row.

**Example:**
```
Cell B1: =A1*$D$1
```
If you copy B1 down to B2:
```
Cell B2: =A2*$D$1
```
The `$D$1` stays fixed, while `A1` changes to `A2` (relative).

### 8.3 Mixed References ($A1 or A$1)

A **mixed reference** locks either the column OR the row, but not both.

| Reference | Column | Row | When Copied |
|-----------|--------|-----|-------------|
| `$A1` | **Locked** (always column A) | Relative | Row changes, column stays A |
| `A$1` | Relative | **Locked** (always row 1) | Column changes, row stays 1 |

**Practical Example - Multiplication Table:**

| | A | B | C |
|---|---|---|---|
| **1** | | 2 | 3 |
| **2** | 2 | | |
| **3** | 3 | | |

In cell B2, enter:
```
=$A2*B$1
```

- `$A2` - column is locked to A, row adjusts when copied down
- `B$1` - row is locked to 1, column adjusts when copied right

When copied across B2:C3, this produces a perfect multiplication table.

### 8.4 Quick Toggle: F4 Key

When editing a formula, place your cursor on a cell reference and press **F4** to cycle through reference types:
```
A1  →  $A$1  →  A$1  →  $A1  →  A1
(relative → absolute → mixed → mixed → relative)
```

---

## 9. Basic Data Types and Formats

### 9.1 Data Types in Excel

| Type | Description | Example | Alignment |
|------|-------------|---------|-----------|
| **Text** (String) | Letters, words, alphanumeric | "John Smith", "INV-001" | Left |
| **Number** | Numeric values | 42, 3.14, -100 | Right |
| **Date** | Date values (stored as numbers) | 1/15/2026 | Right |
| **Time** | Time values (stored as decimal numbers) | 2:30 PM | Right |
| **Boolean** | TRUE or FALSE | TRUE | Center |
| **Error** | Formula errors | #VALUE!, #REF! | Center |
| **Currency** | Numbers formatted as currency | $1,234.56 | Right |
| **Percentage** | Numbers formatted as percentages | 25% | Right |

### 9.2 Number Formatting

To change the format of a cell:
1. Select the cell(s)
2. Right-click → **Format Cells** (or press **Ctrl+1**)
3. Choose from categories:
   - **General** - default, no specific format
   - **Number** - decimal places, thousands separator
   - **Currency** - adds currency symbol
   - **Accounting** - aligns currency symbols in a column
   - **Date** - various date display formats
   - **Time** - various time display formats
   - **Percentage** - multiplies by 100 and adds %
   - **Text** - treats content as text
   - **Special** - Zip Code, Phone Number, SSN

**Quick Number Format Shortcuts:**

| Shortcut | Format |
|----------|--------|
| Ctrl+Shift+~ | General |
| Ctrl+Shift+! | Number (with commas, 2 decimals) |
| Ctrl+Shift+$ | Currency |
| Ctrl+Shift+% | Percentage |
| Ctrl+Shift+# | Date (dd-mmm-yy) |
| Ctrl+Shift+@ | Time (h:mm AM/PM) |

---

## 10. Saving and File Formats

### 10.1 Saving Files

| Action | Shortcut |
|--------|----------|
| Save | **Ctrl+S** |
| Save As | **F12** |
| Save a copy | File → Save a Copy |

### 10.2 File Formats

| Extension | Format | Description |
|-----------|--------|-------------|
| `.xlsx` | Excel Workbook | **Default format.** XML-based, no macros |
| `.xlsm` | Excel Macro-Enabled Workbook | Contains VBA macros |
| `.xlsb` | Excel Binary Workbook | Faster for large files, smaller file size |
| `.xls` | Excel 97-2003 Workbook | Legacy format (limited to 65,536 rows) |
| `.csv` | Comma-Separated Values | Plain text, no formatting, no formulas |
| `.pdf` | Portable Document Format | Read-only, for sharing/printing |
| `.ods` | OpenDocument Spreadsheet | LibreOffice/Google Sheets compatible |
| `.xltm` | Macro-Enabled Template | Template with macro support |
| `.xltx` | Excel Template | Template without macros |

### 10.3 AutoSave and AutoRecover

- **AutoSave** (Microsoft 365): Automatically saves to OneDrive/SharePoint as you work
- **AutoRecover**: Saves temporary copies at intervals (default: every 10 minutes). Configure in **File → Options → Save**
- If Excel crashes, AutoRecover files appear in the **Document Recovery** pane when you restart

---

## 11. Excel Online vs Desktop

| Feature | Excel Desktop | Excel Online (Free) |
|---------|--------------|---------------------|
| Cost | Requires Microsoft 365 subscription or one-time purchase | Free with Microsoft account |
| Full feature set | ✅ All features | ❌ Limited subset |
| VBA Macros | ✅ Full support | ❌ Not supported |
| Power Query | ✅ Full support | ❌ Not supported |
| PivotTables | ✅ Full support | ✅ Basic support |
| Co-authoring | ✅ (via OneDrive) | ✅ Real-time by default |
| Add-ins | ✅ Full support | ❌ Limited |
| Offline use | ✅ Yes | ❌ Requires internet |
| Advanced charts | ✅ All chart types | ❌ Limited chart types |
| Data Analysis ToolPak | ✅ Yes | ❌ No |
| Maximum rows | 1,048,576 | 1,048,576 |

**When to use Excel Online:**
- Quick edits when you don't have your computer
- Collaborative editing in real-time
- Light data entry and simple formulas
- Accessing files stored on OneDrive from any device

**When to use Excel Desktop:**
- Complex formulas and large datasets
- VBA/macro automation
- Power Query and Power Pivot
- Advanced data analysis and modeling

---

## 12. Practice Exercises

### Exercise 1: Basic Data Entry

Create a new workbook and enter the following data:

| | A | B | C | D |
|---|---|---|---|---|
| **1** | **Product** | **Quantity** | **Price** | **Total** |
| **2** | Apples | 10 | 1.50 | |
| **3** | Bananas | 15 | 0.75 | |
| **4** | Oranges | 8 | 2.00 | |
| **5** | Grapes | 5 | 3.25 | |

**Tasks:**
1. Enter the headers in Row 1 (A1:D1)
2. Enter the product data in Rows 2-5
3. In cell D2, enter the formula `=B2*C2` and press Enter
4. Copy the formula from D2 down to D5
5. Format column D as Currency (Ctrl+1 → Currency)

### Exercise 2: Cell References Practice

Using the data from Exercise 1:

1. In cell A7, type: **Grand Total**
2. In cell D7, use a SUM formula: `=SUM(D2:D5)`
3. In cell A8, type: **Tax Rate**
4. In cell B8, enter: `0.08` (8%)
5. In cell A9, type: **Tax Amount**
6. In cell D9, enter a formula that multiplies D7 by B8 (using an absolute reference for B8)
7. In cell A10, type: **Grand Total with Tax**
8. In cell D10, add D7 + D9

### Exercise 3: Navigation Practice

1. Press **Ctrl+G** (Go To), type `Z50`, press Enter - Excel jumps to Z50
2. Press **Ctrl+Home** - returns to A1
3. Click the **Name Box**, type `A1:D5`, press Enter - selects the range
4. Press **Ctrl+End** - jumps to the last used cell
5. Press **Ctrl+Home** again - returns to A1

### Exercise 4: Formatting Practice

1. Select cells A1:D1 (headers)
2. Press **Ctrl+B** to make them bold
3. Press **Ctrl+1** → Alignment tab → check "Center" horizontal alignment
4. Select column D (click the column letter D)
5. Right-click → Format Cells → Number → 2 decimal places
6. Double-click the line between column letters D and E in the header to auto-fit column D width

### Exercise 5: File Management

1. Save the workbook as `Practice_Exercises.xlsx` (Ctrl+S)
2. Save a copy as `Practice_Exercises.csv` (File → Save As → CSV)
3. Close the workbook (Ctrl+W)
4. Reopen the .csv file - notice that all formatting and formulas are gone (only values remain)
5. Open the .xlsx file - notice everything is preserved

---

## 13. Video References

> **Note:** These are recommended YouTube videos for visual learners. Open the links in your browser to watch.

- **Excel for Beginners | Excel Tutorial** - by ExcelIsFun
  - https://www.youtube.com/watch?v=rwbho0CgEAE

- **Microsoft Excel Tutorial for Beginners - Full Course** - by freeCodeCamp.org
  - https://www.youtube.com/watch?v=Vl0H-qTcl3s

- **Excel Basics - Getting Started with Excel** - by ExcelJet
  - https://www.youtube.com/watch?v=kOO31qFmi9A

- **Excel Interface Explained (Ribbon, Tabs, Quick Access Toolbar)** - by Leila Gharani
  - https://www.youtube.com/watch?v=ywYD6vMI1Bk

- **Absolute vs Relative Cell References in Excel** - by ExcelIsFun
  - https://www.youtube.com/watch?v=dB9s1hGEKkY

- **Excel Keyboard Shortcuts You MUST Know** - by Leila Gharani
  - https://www.youtube.com/watch?v=m5qjSW_M3gQ

---

## 14. Sources

1. Microsoft Support - "Excel Quick Start Guide"
   - https://support.microsoft.com/en-us/excel

2. Microsoft Support - "What's new in Excel"
   - https://support.microsoft.com/en-us/excel

3. Excel Easy - "Introduction to Excel"
   - https://www.excel-easy.com/introduction.html

4. Excel Easy - "Basics of Excel"
   - https://www.excel-easy.com/basics.html

5. ExcelJet - "Excel Shortcuts"
   - https://exceljet.net/shortcuts

6. GCFGlobal - "Excel 2016: Getting Started with Excel"
   - https://edu.gcfglobal.org/en/excel/

7. Corporate Finance Institute - "Excel Shortcuts Overview"
   - https://corporatefinanceinstitute.com/resources/excel/study/excel-shortcuts/

8. TrumpExcel - "Excel Basics"
   - https://trumpexcel.com/learn-excel/

9. Chandoo - "Excel Basics"
   - https://chandoo.org/?s=excelexcel-basics/

10. Excel Easy - "Format Cells"
    - https://www.excel-easy.com/basics/format-cells.html

---

*Module 1 Complete. Proceed to [Module 2: Formulas & Functions →](../02-Formulas-Functions/02-formulas-functions.md)*


## FILE: 02-Formulas-Functions/02-formulas-functions.md

# Module 2: Microsoft Excel Formulas & Functions

> **Level:** Beginner to Intermediate | **Estimated Time:** 3–4 hours | **Prerequisites:** [Module 1: Excel Fundamentals](../01-Fundamentals/01-excel-fundamentals.md)

---

## Table of Contents

1. [Formula Basics](#1-formula-basics)
2. [Statistical Functions](#2-statistical-functions-sum-average-count-min-max)
3. [Conditional Functions](#3-conditional-functions-sumif-countif-averageif)
4. [Logical Functions](#4-logical-functions-if-and-or-not)
5. [Advanced Logical Functions](#5-advanced-logical-functions-ifs-switch-nested-if)
6. [Text Functions](#6-text-functions)
7. [Date Functions](#7-date-functions)
8. [Lookup Functions](#8-lookup-functions-vlookup-hlookup)
9. [Error Handling](#9-error-handling)
10. [Named Ranges](#10-named-ranges)
11. [Formula Auditing Tools](#11-formula-auditing-tools)
12. [Practice Exercises](#12-practice-exercises)
13. [Quick Reference Cheat Sheet](#13-quick-reference-cheat-sheet)
14. [Video References](#14-video-references)
15. [Sources](#15-sources)

---

## 1. Formula Basics

### 1.1 What is a Formula?

A **formula** is an expression that calculates the value of a cell. Every formula in Excel:
- **Starts with an equals sign** `=`
- Can contain **numbers**, **cell references**, **operators**, and **functions**
- Displays its **result** in the cell and the **formula** in the Formula Bar

### 1.2 Formula Syntax

```
= value1 OPERATOR value2
```

**Examples:**
```
=10+5          → 15
=A1+B1         → Sum of values in A1 and B1
=SUM(A1:A10)   → Sum of values in A1 through A10
```

### 1.3 Arithmetic Operators

| Operator | Meaning | Example | Result |
|----------|---------|---------|--------|
| `+` | Addition | `=5+3` | 8 |
| `-` | Subtraction | `=10-4` | 6 |
| `*` | Multiplication | `=6*7` | 42 |
| `/` | Division | `=20/4` | 5 |
| `^` | Exponentiation | `=2^3` | 8 |
| `%` | Percentage | `=50%` | 0.5 |

### 1.4 Order of Operations (PEMDAS)

Excel follows the standard mathematical order of operations:

1. **P**arentheses `()`
2. **E**xponents `^`
3. **M**ultiplication `*` and **D**ivision `/` (left to right)
4. **A**ddition `+` and **S**ubtraction `-` (left to right)

**Example:**
```
=2+3*4
```
- Multiplication first: `3*4 = 12`
- Then addition: `2+12 = 14`
- Result: **14**

With parentheses:
```
=(2+3)*4
```
- Parentheses first: `2+3 = 5`
- Then multiplication: `5*4 = 20`
- Result: **20**

**Complex Example:**
```
=2^3+4*5-6/3
```
Step by step:
1. Exponents: `2^3 = 8` → `8+4*5-6/3`
2. Multiplication/Division (left to right): `4*5 = 20`, `6/3 = 2` → `8+20-2`
3. Addition/Subtraction (left to right): `8+20 = 28`, `28-2 = 26`
4. **Result: 26**

---

## 2. Statistical Functions: SUM, AVERAGE, COUNT, MIN, MAX

### Sample Data for This Section

Create this table in your Excel sheet:

| | A | B | C | D |
|---|---|---|---|---|
| **1** | **Employee** | **Department** | **Salary** | **Sales** |
| **2** | Alice | Marketing | 55000 | 12000 |
| **3** | Bob | Sales | 48000 | 18500 |
| **4** | Carol | Marketing | 62000 | 15000 |
| **5** | David | IT | 75000 | 9000 |
| **6** | Eve | Sales | 51000 | 22000 |
| **7** | Frank | IT | 70000 | 11000 |

### 2.1 SUM

Adds all numbers in a range.

**Syntax:** `=SUM(number1, [number2], ...)`

**Examples:**
```
=SUM(C2:C7)          → Sum of all salaries: 361000
=SUM(C2,C4,C6)       → Sum of specific cells: 168000
=SUM(C2:C7,D2:D7)    → Sum of two ranges: 448500
=SUM(10,20,30)       → Sum of literal values: 60
```

### 2.2 AVERAGE

Returns the arithmetic mean of a range.

**Syntax:** `=AVERAGE(number1, [number2], ...)`

**Examples:**
```
=AVERAGE(C2:C7)      → Average salary: 60166.67
=AVERAGE(D2:D3)      → Average of first two sales: 15250
```

### 2.3 COUNT

Counts the number of cells that contain **numbers**.

**Syntax:** `=COUNT(value1, [value2], ...)`

**Examples:**
```
=COUNT(C2:C7)        → 6 (all cells contain numbers)
=COUNT(A2:A7)        → 0 (all cells contain text, not numbers)
=COUNT(C2:C7,D2:D7)  → 12
```

### 2.4 COUNTA

Counts the number of **non-empty** cells (text, numbers, errors, etc.).

**Syntax:** `=COUNTA(value1, [value2], ...)`

**Examples:**
```
=COUNTA(A2:A7)       → 6 (all cells have text)
=COUNTA(C2:C7)       → 6
```

### 2.5 MIN and MAX

**MIN** returns the smallest value; **MAX** returns the largest value.

**Syntax:** `=MIN(number1, [number2], ...)` / `=MAX(number1, [number2], ...)`

**Examples:**
```
=MIN(C2:C7)          → 48000 (Bob's salary)
=MAX(C2:C7)          → 75000 (David's salary)
=MIN(D2:D7)          → 9000 (David's sales)
=MAX(D2:D7)          → 22000 (Eve's sales)
```

---

## 3. Conditional Functions: SUMIF, COUNTIF, AVERAGEIF

These functions perform calculations **only on cells that meet a specified condition**.

### 3.1 SUMIF

Sums values in a range that meet a single criterion.

**Syntax:** `=SUMIF(range, criteria, [sum_range])`

- `range` - the range to evaluate (check against the criteria)
- `criteria` - the condition (e.g., `"Marketing"`, `">50000"`, `"A*"`)
- `sum_range` - the range to sum (optional; if omitted, sums the `range`)

**Examples using the sample data above:**
```
=SUMIF(B2:B7,"Marketing",C2:C7)   → 117000 (Alice 55000 + Carol 62000)
=SUMIF(B2:B7,"Sales",C2:C7)       → 99000 (Bob 48000 + Eve 51000)
=SUMIF(B2:B7,"IT",D2:D7)          → 20000 (David 9000 + Frank 11000)
=SUMIF(C2:C7,">60000")            → 205000 (sums salaries over 60000: 62000+75000+70000)
```

### 3.2 SUMIFS

Sums values that meet **multiple** criteria.

**Syntax:** `=SUMIFS(sum_range, criteria_range1, criteria1, [criteria_range2, criteria2], ...)`

**Examples:**
```
=SUMIFS(C2:C7,B2:B7,"Marketing",C2:C7,">55000")  → 62000 (only Carol meets both conditions)
=SUMIFS(D2:D7,B2:B7,"Sales",D2:D7,">15000")      → 405000 (Bob 18500 + Eve 22000)
```

### 3.3 COUNTIF

Counts cells that meet a single criterion.

**Syntax:** `=COUNTIF(range, criteria)`

**Examples:**
```
=COUNTIF(B2:B7,"Marketing")       → 2
=COUNTIF(B2:B7,"Sales")           → 2
=COUNTIF(C2:C7,">60000")          → 3
=COUNTIF(A2:A7,"A*")              → 1 (starts with "A" - Alice)
```

### 3.4 COUNTIFS

Counts cells that meet **multiple** criteria.

**Syntax:** `=COUNTIFS(criteria_range1, criteria1, [criteria_range2, criteria2], ...)`

**Examples:**
```
=COUNTIFS(B2:B7,"Marketing",C2:C7,">55000")  → 1 (only Carol)
=COUNTIFS(B2:B7,"Sales",D2:D7,">15000")      → 2 (Bob and Eve)
```

### 3.5 AVERAGEIF

Averages values that meet a single criterion.

**Syntax:** `=AVERAGEIF(range, criteria, [average_range])`

**Examples:**
```
=AVERAGEIF(B2:B7,"Marketing",C2:C7)    → 58500 ((55000+62000)/2)
=AVERAGEIF(B2:B7,"Sales",C2:C7)        → 49500 ((48000+51000)/2)
=AVERAGEIF(C2:C7,">60000")             → 68333.33 ((62000+75000+70000)/3)
```

### 3.6 AVERAGEIFS

Averages values that meet **multiple** criteria.

**Syntax:** `=AVERAGEIFS(average_range, criteria_range1, criteria1, ...)`

**Examples:**
```
=AVERAGEIFS(D2:D7,B2:B7,"Sales",D2:D7,">15000")  → 20250 ((18500+22000)/2)
```

### 3.7 Criteria Operators Reference

| Operator | Meaning | Example |
|----------|---------|---------|
| `">100"` | Greater than 100 | `COUNTIF(C:C,">100")` |
| `"<50"` | Less than 50 | `COUNTIF(C:C,"<50")` |
| `">=100"` | Greater than or equal to 100 | `COUNTIF(C:C,">=100")` |
| `"<>0"` | Not equal to 0 | `COUNTIF(C:C,"<>0")` |
| `"A*"` | Starts with "A" (wildcard) | `COUNTIF(A:A,"A*")` |
| `"*son"` | Ends with "son" | `COUNTIF(A:A,"*son")` |
| `"*mi*"` | Contains "mi" | `COUNTIF(A:A,"*mi*")` |
| `"???"` | Exactly 3 characters | `COUNTIF(A:A,"???")` |

---

## 4. Logical Functions: IF, AND, OR, NOT

### 4.1 IF Function

Returns one value if a condition is TRUE, and another if it is FALSE.

**Syntax:** `=IF(logical_test, value_if_true, value_if_false)`

**Examples using sample data:**
```
=IF(C2>60000,"High","Low")          → "Low" (Alice's 55000 is not > 60000)
=IF(C5>60000,"High","Low")          → "High" (David's 75000 > 60000)
=IF(D2>15000,"Bonus","No Bonus")    → "No Bonus" (Alice's 12000 is not > 15000)
```

**Practical Examples:**
```
=IF(A2>=90,"A",IF(A2>=80,"B",IF(A2>=70,"C","F")))   → Nested IF for letter grades
=IF(B2="Sales",C2*0.1,C2*0.05)                       → Different commission rates by department
```

### 4.2 AND Function

Returns TRUE if **ALL** conditions are true.

**Syntax:** `=AND(condition1, [condition2], ...)`

**Examples:**
```
=AND(C2>50000,D2>10000)     → TRUE (Alice: 55000>50000 AND 12000>10000)
=AND(C2>70000,D2>15000)     → FALSE (55000 is not > 70000)
```

### 4.3 OR Function

Returns TRUE if **ANY** condition is true.

**Syntax:** `=OR(condition1, [condition2], ...)`

**Examples:**
```
=OR(C2>70000,D2>20000)      → FALSE (neither condition is true for Alice)
=OR(C5>70000,D5>20000)      → TRUE (David: 75000>70000)
```

### 4.4 NOT Function

Reverses the logical value - TRUE becomes FALSE and vice versa.

**Syntax:** `=NOT(logical)`

**Examples:**
```
=NOT(C2>70000)              → TRUE (55000 is NOT > 70000)
=NOT(B2="Sales")            → TRUE (Alice is NOT in Sales)
```

### 4.5 Combining Logical Functions with IF

**Real-world Example - Performance Bonus:**
```
=IF(AND(C2>50000,D2>10000),"Bonus: $2000","No Bonus")
```
- Alice (Salary 55000, Sales 12000): **Bonus: $2000** ✅
- Bob (Salary 48000, Sales 18500): **No Bonus** ❌ (salary ≤ 50000)

**Another Example - Eligibility Check:**
```
=IF(OR(B2="Sales",B2="Marketing"),"Revenue Team","Support Team")
```
- Alice (Marketing): **Revenue Team**
- David (IT): **Support Team**

---

## 5. Advanced Logical Functions: IFS, SWITCH, Nested IF

### 5.1 Nested IF

Nesting multiple IF functions for multi-condition logic.

**Example - Letter Grades:**
```
=IF(C2>=90000,"A",IF(C2>=80000,"B",IF(C2>=70000,"C",IF(C2>=60000,"D","F"))))
```

**Readability Tip:** For more than 3 conditions, use IFS instead.

### 5.2 IFS Function (Excel 2019+ / Microsoft 365)

Tests multiple conditions and returns the value for the first TRUE condition.

**Syntax:** `=IFS(condition1, value1, condition2, value2, ..., TRUE, default_value)`

**Example - Salary Band:**
```
=IFS(C2>=70000,"Senior",C2>=55000,"Mid-Level",C2>=40000,"Junior",TRUE,"Entry-Level")
```
- Alice (55000): **Mid-Level**
- David (75000): **Senior**
- Bob (48000): **Junior**

**Note:** Always include a `TRUE, "default"` catch-all at the end to avoid #N/A errors.

### 5.3 SWITCH Function (Excel 2019+ / Microsoft 365)

Compares a value against a list of values and returns the first match.

**Syntax:** `=SWITCH(expression, value1, result1, value2, result2, ..., default)`

**Example - Department Codes:**
```
=SWITCH(B2,"Marketing","MKT","Sales","SLS","IT","IT","FIN")
```
- Alice (Marketing): **MKT**
- Bob (Sales): **SLS**
- David (IT): **IT**

**Comparison:**
| Function | Best For | Max Conditions |
|----------|----------|---------------|
| IF | Simple yes/no | 1 |
| Nested IF | 2-3 conditions | 64 (practical limit ~7) |
| IFS | Many conditions | 127 |
| SWITCH | Matching exact values | 126 |

---

## 6. Text Functions

### Sample Data for Text Functions

| | A | B |
|---|---|---|
| **1** | **Full Name** | **Email** |
| **2** | John Smith | john.smith@company.com |
| **3** | jane doe | JANE.DOE@CORP.NET |
| **4** | Bob Johnson | bob.j@startup.io |
| **5** | Alice Marie Brown | alice.b@bigco.com |

### 6.1 LEFT, RIGHT, MID

Extract characters from a text string.

**Syntax:**
- `=LEFT(text, num_chars)` - Extract from the **left**
- `=RIGHT(text, num_chars)` - Extract from the **right**
- `=MID(text, start_num, num_chars)` - Extract from the **middle**

**Examples:**
```
=LEFT(A2,4)           → "John" (first 4 characters)
=RIGHT(B2,4)          → ".com" (last 4 characters)
=MID(A2,6,5)          → "Smith" (5 characters starting at position 6)
=LEFT(A2,FIND(" ",A2)-1)  → "John" (everything before the first space)
```

### 6.2 LEN

Returns the number of characters in a text string.

**Syntax:** `=LEN(text)`

**Examples:**
```
=LEN(A2)              → 10 ("John Smith" = 10 characters)
=LEN(A4)              → 12 ("Bob Johnson" = 12 characters)
```

### 6.3 TRIM

Removes extra spaces from text (leading, trailing, and multiple spaces between words).

**Syntax:** `=TRIM(text)`

**Example:**
```
=TRIM("  Hello   World  ")  → "Hello World"
```

### 6.4 UPPER, LOWER, PROPER

Change text case.

**Syntax:**
- `=UPPER(text)` - ALL CAPS
- `=LOWER(text)` - all lowercase
- `=PROPER(text)` - Title Case

**Examples:**
```
=UPPER(A2)            → "JOHN SMITH"
=LOWER(A3)            → "jane doe"
=PROPER(A3)           → "Jane Doe"
```

### 6.5 CONCATENATE / CONCAT

Joins two or more text strings together.

**Syntax:**
- `=CONCATENATE(text1, text2, ...)` - Legacy function
- `=CONCAT(text1, text2, ...)` - Modern replacement (Excel 2016+)
- `=text1 & text2 & ...` - Ampersand operator (most common)

**Examples:**
```
=CONCATENATE(A2," - ",B2)     → "John Smith - john.smith@company.com"
=CONCAT(A2," - ",B2)          → "John Smith - john.smith@company.com"
=A2 & " - " & B2               → "John Smith - john.smith@company.com"
="First: " & LEFT(A2,4)       → "First: John"
```

### 6.6 TEXTJOIN (Excel 2019+ / Microsoft 365)

Joins text with a specified delimiter.

**Syntax:** `=TEXTJOIN(delimiter, ignore_empty, text1, [text2], ...)`

**Examples:**
```
=TEXTJOIN(", ",TRUE,A2,A3,A4,A5)  → "John Smith, jane doe, Bob Johnson, Alice Marie Brown"
=TEXTJOIN(" | ",FALSE,A2,B2)      → "John Smith | john.smith@company.com"
```

### 6.7 FIND and SEARCH

Find the position of text within a string.

**Syntax:**
- `=FIND(find_text, within_text, [start_num])` - Case-sensitive
- `=SEARCH(find_text, within_text, [start_num])` - Case-insensitive, supports wildcards

**Examples:**
```
=FIND(" ",A2)                    → 5 (position of space in "John Smith")
=SEARCH("@",B2)                  → 11 (position of @ in email)
=LEFT(B2,SEARCH("@",B2)-1)      → "john.smith" (extract username from email)
=MID(B2,SEARCH("@",B2)+1,100)   → "company.com" (extract domain from email)
```

### 6.8 SUBSTITUTE and REPLACE

Replace text within a string.

**Syntax:**
- `=SUBSTITUTE(text, old_text, new_text, [instance_num])` - Replace by content
- `=REPLACE(old_text, start_num, num_chars, new_text)` - Replace by position

**Examples:**
```
=SUBSTITUTE(A2," ",".")         → "John.Smith" (replace space with dot)
=SUBSTITUTE(B2,".com",".org")   → "john.smith@company.org"
```

### 6.9 VALUE and TEXT

Convert between text and numbers.

**Syntax:**
- `=VALUE(text)` - Converts text to number
- `=TEXT(value, format_text)` - Converts number to text with a format

**Examples:**
```
=VALUE("123")                   → 123 (number)
=TEXT(TODAY(),"MM/DD/YYYY")     → "09/10/2026" (formatted date as text)
=TEXT(1234.5,"$#,##0.00")       → "$1,234.50" (formatted currency as text)
```

---

## 7. Date Functions

### 7.1 TODAY and NOW

**Syntax:**
- `=TODAY()` - Returns today's date (updates when the sheet recalculates)
- `=NOW()` - Returns today's date AND current time

**Examples:**
```
=TODAY()              → 9/10/2026 (or current date)
=NOW()                → 9/10/2026 2:30 PM (or current date and time)
```

**Note:** These are volatile functions - they recalculate every time the worksheet changes.

### 7.2 DATE Function

Creates a date from individual year, month, and day values.

**Syntax:** `=DATE(year, month, day)`

**Examples:**
```
=DATE(2026,9,10)      → September 10, 2026
=DATE(2026,12,25)     → December 25, 2026
```

### 7.3 YEAR, MONTH, DAY

Extract components from a date.

**Syntax:**
- `=YEAR(date)` - Returns the year (e.g., 2026)
- `=MONTH(date)` - Returns the month (1-12)
- `=DAY(date)` - Returns the day (1-31)

**Examples:**
```
=YEAR(DATE(2026,9,10))   → 2026
=MONTH(DATE(2026,9,10))  → 9
=DAY(DATE(2026,9,10))    → 10
```

### 7.4 DATEDIF

Calculates the difference between two dates in years, months, or days.

**Syntax:** `=DATEDIF(start_date, end_date, unit)`

| Unit | Returns | Description |
|------|---------|-------------|
| `"Y"` | Years | Complete years between dates |
| `"M"` | Months | Complete months between dates |
| `"D"` | Days | Total days between dates |
| `"YM"` | Months | Months ignoring years |
| `"YD"` | Days | Days ignoring years |
| `"MD"` | Days | Days ignoring months and years |

**Examples:**
```
=DATEDIF(DATE(2020,3,15),TODAY(),"Y")    → 6 (years since March 15, 2020)
=DATEDIF(DATE(2020,3,15),TODAY(),"M")    → 78 (months)
=DATEDIF(DATE(2020,3,15),TODAY(),"D")    → 2371 (days, approximate)
```

### 7.5 EOMONTH

Returns the last day of the month, a specified number of months before or after a date.

**Syntax:** `=EOMONTH(start_date, months)`

**Examples:**
```
=EOMONTH(TODAY(),0)       → September 30, 2026 (end of current month)
=EOMONTH(TODAY(),1)       → October 31, 2026 (end of next month)
=EOMONTH(TODAY(),-1)      → August 31, 2026 (end of previous month)
=EOMONTH(DATE(2026,1,15),0)  → January 31, 2026
```

### 7.6 WEEKDAY

Returns the day of the week as a number (1=Sunday by default, or customizable).

**Syntax:** `=WEEKDAY(date, [return_type])`

| return_type | Week starts on |
|-------------|----------------|
| 1 (default) | Sunday (1) through Saturday (7) |
| 2 | Monday (1) through Sunday (7) |
| 3 | Monday (0) through Sunday (6) |

**Examples:**
```
=WEEKDAY(DATE(2026,9,10))     → 5 (Thursday, when 1=Sunday)
=WEEKDAY(DATE(2026,9,10),2)   → 4 (Thursday, when 1=Monday)
```

### 7.7 NETWORKDAYS

Counts the number of working days between two dates (excludes weekends, optionally holidays).

**Syntax:** `=NETWORKDAYS(start_date, end_date, [holidays])`

**Example:**
```
=NETWORKDAYS(DATE(2026,1,1),DATE(2026,12,31))    → 261 (approximate working days in 2026)
```

---

## 8. Lookup Functions: VLOOKUP, HLOOKUP

### Sample Data for Lookup Section

**Employee Table (A1:D7):**

| | A | B | C | D |
|---|---|---|---|---|
| **1** | **Emp ID** | **Name** | **Department** | **Salary** |
| **2** | 101 | Alice | Marketing | 55000 |
| **3** | 102 | Bob | Sales | 48000 |
| **4** | 103 | Carol | Marketing | 62000 |
| **5** | 104 | David | IT | 75000 |
| **6** | 105 | Eve | Sales | 51000 |
| **7** | 106 | Frank | IT | 70000 |

### 8.1 VLOOKUP (Vertical Lookup)

Searches for a value in the **first column** of a range and returns a value from a specified column in the same row.

**Syntax:** `=VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup])`

- `lookup_value` - The value to search for
- `table_array` - The range containing the data (first column must contain the lookup value)
- `col_index_num` - The column number to return from (1 = first column)
- `range_lookup` - `FALSE` for exact match (most common), `TRUE` for approximate match

**Examples:**
```
=VLOOKUP(103,A2:D7,2,FALSE)      → "Carol" (look up Emp ID 103, return column 2)
=VLOOKUP(103,A2:D7,3,FALSE)      → "Marketing" (return column 3)
=VLOOKUP(103,A2:D7,4,FALSE)      → 62000 (return column 4)
=VLOOKUP(105,A2:D7,2,FALSE)      → "Eve"
```

**Practical Example:** If cell F2 contains an employee ID:
```
=VLOOKUP(F2,A2:D7,2,FALSE)       → Returns the employee name for the ID in F2
```

### 8.2 HLOOKUP (Horizontal Lookup)

Same concept as VLOOKUP, but searches across **rows** instead of columns.

**Syntax:** `=HLOOKUP(lookup_value, table_array, row_index_num, [range_lookup])`

**Use case:** When your data is organized horizontally with headers in the first row.

### 8.3 VLOOKUP Limitations

| Limitation | Description |
|-----------|-------------|
| **Left lookup only** | Can only search the leftmost column of the range |
| **No left return** | Can only return values to the RIGHT of the lookup column |
| **Column index is hard-coded** | If you insert/delete columns, the index number may be wrong |
| **Case-insensitive** | Cannot distinguish between "Apple" and "APPLE" |
| **Approximate match default** | If you forget `FALSE`, it may return wrong results |

### 8.4 Modern Alternatives (Microsoft 365 / Excel 2021+)

**XLOOKUP** overcomes all VLOOKUP limitations:
```
=XLOOKUP(lookup_value, lookup_array, return_array, [if_not_found], [match_mode], [search_mode])
```

**Example:**
```
=XLOOKUP(103,A2:A7,B2:B7)        → "Carol"
=XLOOKUP(103,A2:A7,D2:D7,"Not Found")  → 62000 (with error handling)
```

**Advantages over VLOOKUP:**
- Can look left OR right
- Returns a range, not a column number
- Built-in error handling
- Default is exact match
- Supports reverse search

---

## 9. Error Handling

### 9.1 Common Excel Errors

| Error | Cause | Example |
|-------|-------|---------|
| `#VALUE!` | Wrong data type in formula | `="text"+5` |
| `#REF!` | Invalid cell reference (deleted cell) | `=A1` when column A is deleted |
| `#NAME?` | Unrecognized function name or text | `=SUN(A1:A10)` (typo) |
| `#DIV/0!` | Division by zero | `=10/0` |
| `#N/A` | Value not found (lookup) | VLOOKUP can't find the value |
| `#NUM!` | Invalid numeric value | `=SQRT(-1)` |
| `#NULL!` | Incorrect range intersection | `=SUM(A1 B1)` (missing colon/comma) |
| `#SPILL!` | Dynamic array can't spill | Array output blocked by data |

### 9.2 IFERROR

Catches **any** error and returns a specified value instead.

**Syntax:** `=IFERROR(value, value_if_error)`

**Examples:**
```
=IFERROR(VLOOKUP(999,A2:D7,2,FALSE),"Not Found")     → "Not Found" (999 doesn't exist)
=IFERROR(A1/B1,"Cannot divide")                       → "Cannot divide" if B1 is 0
=IFERROR(1/0,"Error")                                 → "Error"
```

### 9.3 IFNA

Catches only `#N/A` errors (useful when you want other errors to still show).

**Syntax:** `=IFNA(value, value_if_na)`

**Examples:**
```
=IFNA(VLOOKUP(999,A2:D7,2,FALSE),"Not Found")         → "Not Found"
=IFNA(VLOOKUP(103,A2:D7,2,FALSE),"Not Found")         → "Carol" (no error)
```

### 9.4 ISERROR and ISNA

Return TRUE/FALSE - useful inside IF statements.

**Syntax:**
- `=ISERROR(value)` - Returns TRUE if ANY error
- `=ISNA(value)` - Returns TRUE only if #N/A

**Example:**
```
=IF(ISERROR(VLOOKUP(999,A2:D7,2,FALSE)),"Check ID","OK")  → "Check ID"
```

### 9.5 Error Handling Best Practice

```
=IFERROR(your_formula, "fallback value")
```

Always wrap lookup formulas and complex calculations in IFERROR to prevent #N/A, #REF!, and other errors from breaking your spreadsheet.

---

## 10. Named Ranges

### 10.1 What is a Named Range?

A **named range** is a descriptive name assigned to a cell or range of cells. Instead of writing `A2:D7`, you write `EmployeeData`.

### 10.2 Creating a Named Range

**Method 1: Name Box**
1. Select the range (e.g., A2:D7)
2. Click the **Name Box** (left of the Formula Bar)
3. Type a name (e.g., `EmployeeData`)
4. Press **Enter**

**Method 2: Define Name (Formulas Tab)**
1. Select the range
2. Go to **Formulas** tab → **Define Name**
3. Enter a name, scope (Workbook or specific sheet), and optional comment
4. Click **OK**

**Method 3: Create from Selection**
1. Select a range that includes headers
2. Go to **Formulas** tab → **Create from Selection**
3. Check where the names are (e.g., "Top row", "Left column")
4. Click **OK** - Excel creates names from the headers

### 10.3 Naming Rules

- Must begin with a letter or underscore
- Cannot contain spaces (use underscores: `Employee_Data`)
- Cannot be a cell reference (e.g., avoid names like `A1` or `R1C1`)
- Maximum 255 characters
- Not case-sensitive (`data` and `DATA` are the same)

### 10.4 Using Named Ranges in Formulas

**Before (cell references):**
```
=SUM(C2:C7)
=VLOOKUP(103,A2:D7,2,FALSE)
```

**After (named ranges):**
```
=SUM(Salary)
=VLOOKUP(103,EmployeeData,2,FALSE)
```

**Benefits:**
- Formulas are easier to read and understand
- Easier to maintain - update the range once, all formulas update
- Named ranges appear in the Name Box for quick navigation

### 10.5 Managing Named Ranges

- **View all:** Formulas tab → **Name Manager** (Ctrl+F3)
- **Edit:** Open Name Manager → select name → Edit
- **Delete:** Open Name Manager → select name → Delete
- **Scope:** Names can be scoped to the entire workbook or a specific sheet

---

## 11. Formula Auditing Tools

### 11.1 Trace Precedents and Dependents

Found on the **Formulas** tab → **Formula Auditing** group.

- **Trace Precedents:** Shows arrows pointing TO the selected cell from cells it references
  - Select a cell with a formula → click "Trace Precedents"
  - Blue arrows show which cells feed into the formula

- **Trace Dependents:** Shows arrows pointing FROM the selected cell to cells that reference it
  - Select a cell → click "Trace Dependents"
  - Shows which formulas use this cell's value

- **Remove Arrows:** Click "Remove Arrows" to clear the display

### 11.2 Show Formulas

Toggle between showing formulas and their results:
- **Formulas** tab → **Show Formulas** button
- Or press **Ctrl+`** (backtick, the key above Tab)

This is extremely useful for auditing - all cells show their formulas instead of results.

### 11.3 Evaluate Formula

**Formulas** tab → **Evaluate Formula**

This steps through a formula one calculation at a time, showing you exactly how Excel evaluates it. Essential for debugging complex nested formulas.

### 11.4 Error Checking

**Formulas** tab → **Error Checking**

Excel identifies cells with errors and offers suggestions to fix them. It will walk through each error in the worksheet.

### 11.5 Watch Window

**Formulas** tab → **Watch Window**

Add cells to a floating watch window that stays visible even when you scroll to other parts of the sheet. Useful for monitoring key values while working on distant cells.

---

## 12. Practice Exercises

### Exercise Set A: Statistical Functions

**Data: Monthly Sales Report**

| | A | B | C | D |
|---|---|---|---|---|
| **1** | **Month** | **Region** | **Product** | **Revenue** |
| **2** | Jan | North | Widget | 15000 |
| **3** | Jan | South | Widget | 12000 |
| **4** | Feb | North | Gadget | 18000 |
| **5** | Feb | South | Gadget | 9500 |
| **6** | Mar | North | Widget | 22000 |
| **7** | Mar | South | Widget | 14000 |
| **8** | Apr | North | Gadget | 16500 |
| **9** | Apr | South | Gadget | 11000 |

**Tasks:**
1. Calculate total revenue: `=SUM(D2:D9)` → **118000**
2. Calculate average revenue: `=AVERAGE(D2:D9)`
3. Find the highest revenue: `=MAX(D2:D9)`
4. Find the lowest revenue: `=MIN(D2:D9)`
5. Count how many months had revenue > 15000: `=COUNTIF(D2:D9,">15000")`
6. Sum revenue for North region only: `=SUMIF(B2:B9,"North",D2:D9)`
7. Sum revenue for Widget products in North: `=SUMIFS(D2:D9,B2:B9,"North",C2:C9,"Widget")`
8. Average revenue for Gadget products: `=AVERAGEIF(C2:C9,"Gadget",D2:D9)`

### Exercise Set B: Logical Functions

Using the same data:

1. Create a "Performance" column in E1: In E2, write: `=IF(D2>15000,"High","Low")` - copy down to E9
2. In F1, header "Bonus": In F2, write: `=IF(AND(D2>15000,B2="North"),D2*0.1,0)` - copy down
3. Sum all bonuses: `=SUM(F2:F9)`

### Exercise Set C: Text Functions

**Data:**

| | A |
|---|---|
| **1** | **Email** |
| **2** | john.smith@company.com |
| **3** | jane.doe@corp.net |
| **4** | bob.j@startup.io |

**Tasks:**
1. Extract username (before @): `=LEFT(A2,FIND("@",A2)-1)`
2. Extract domain (after @): `=MID(A2,FIND("@",A2)+1,LEN(A2))`
3. Capitalize the username: `=PROPER(LEFT(A2,FIND("@",A2)-1))`

### Exercise Set D: Date Functions

1. Enter today's date: `=TODAY()`
2. Calculate your age in years: `=DATEDIF(DATE(1990,5,15),TODAY(),"Y")` (use your own birthdate)
3. Find the last day of the current month: `=EOMONTH(TODAY(),0)`
4. Count working days in the current month: `=NETWORKDAYS(EOMONTH(TODAY(),-1)+1,EOMONTH(TODAY(),0))`
5. What day of the week is your birthday this year? `=TEXT(DATE(2026,5,15),"dddd")`

### Exercise Set E: VLOOKUP

**Lookup Table (A1:C5):**

| | A | B | C |
|---|---|---|---|
| **1** | **Code** | **Product** | **Price** |
| **2** | W01 | Widget A | 25.00 |
| **3** | W02 | Widget B | 30.00 |
| **4** | G01 | Gadget A | 45.00 |
| **5** | G02 | Gadget B | 55.00 |

**Tasks:**
1. In cell E2, type `G01`. In F2: `=VLOOKUP(E2,A2:C5,2,FALSE)` → "Gadget A"
2. In G2: `=VLOOKUP(E2,A2:C5,3,FALSE)` → 45.00
3. In H2 (with error handling): `=IFERROR(VLOOKUP("X99",A2:C5,2,FALSE),"Product not found")`
4. Try changing E2 to different codes and watch F2 and G2 update

---

## 13. Quick Reference Cheat Sheet

### Most-Used Functions

| Category | Function | Syntax | Purpose |
|----------|----------|--------|---------|
| Math | `SUM` | `=SUM(range)` | Add numbers |
| Math | `AVERAGE` | `=AVERAGE(range)` | Calculate mean |
| Math | `COUNT` | `=COUNT(range)` | Count numbers |
| Math | `COUNTA` | `=COUNTA(range)` | Count non-empty cells |
| Math | `MIN` / `MAX` | `=MIN(range)` / `=MAX(range)` | Smallest / Largest |
| Conditional | `SUMIF` | `=SUMIF(range,crit,sum_range)` | Conditional sum |
| Conditional | `COUNTIF` | `=COUNTIF(range,crit)` | Conditional count |
| Conditional | `AVERAGEIF` | `=AVERAGEIF(range,crit,avg_range)` | Conditional average |
| Logical | `IF` | `=IF(test,true,false)` | Conditional value |
| Logical | `AND` / `OR` | `=AND(c1,c2)` / `=OR(c1,c2)` | Multiple conditions |
| Text | `LEFT/RIGHT/MID` | `=LEFT(text,n)` | Extract text |
| Text | `LEN` / `TRIM` | `=LEN(text)` / `=TRIM(text)` | Length / Clean spaces |
| Text | `UPPER/LOWER/PROPER` | `=UPPER(text)` | Change case |
| Text | `CONCAT` / `&` | `=A1&"-"&B1` | Join text |
| Date | `TODAY` / `NOW` | `=TODAY()` / `=NOW()` | Current date/time |
| Date | `DATEDIF` | `=DATEDIF(d1,d2,"Y")` | Date difference |
| Lookup | `VLOOKUP` | `=VLOOKUP(val,range,col,FALSE)` | Vertical lookup |
| Lookup | `XLOOKUP` | `=XLOOKUP(val,range,return)` | Modern lookup (365) |
| Error | `IFERROR` | `=IFERROR(formula,fallback)` | Catch errors |

---

## 14. Video References

> **Note:** These are recommended YouTube videos for visual learners. Open the links in your browser to watch.

- **Excel Formulas and Functions Tutorial** - by ExcelIsFun
  - https://www.youtube.com/watch?v=V2V0GcMYbmo

- **VLOOKUP Tutorial for Beginners** - by ExcelIsFun
  - https://www.youtube.com/watch?v=1Ej-5V7wMOk

- **IF Function Excel Tutorial** - by Leila Gharani
  - https://www.youtube.com/watch?v=g7h8DfOBPRQ

- **SUMIF and COUNTIF Functions Explained** - by ExcelJet
  - https://www.youtube.com/watch?v=QkGHzJhJ3kA

- **Excel Text Functions Tutorial** - by ExcelIsFun
  - https://www.youtube.com/watch?v=CyBV7FkKBOk

- **Excel Date Functions Explained** - by Leila Gharani
  - https://www.youtube.com/watch?v=ZO5jDaVfDhE

- **XLOOKUP vs VLOOKUP - Why XLOOKUP is Better** - by Leila Gharani
  - https://www.youtube.com/watch?v=Hn1JlSjvKkI

- **Named Ranges in Excel - Complete Guide** - by ExcelJet
  - https://www.youtube.com/watch?v=kOO31qFmi9A

- **Excel Error Handling: IFERROR, IFNA** - by ExcelIsFun
  - https://www.youtube.com/watch?v=6h0GN2eH2kY

---

## 15. Sources

1. Microsoft Support - "Excel Functions (by category)"
   - https://support.microsoft.com/en-us/excel

2. Microsoft Support - "VLOOKUP function"
   - https://support.microsoft.com/en-us/office/vlookup-function-0bbc8083-26fe-4963-8ab8-93a18ad188a1

3. Microsoft Support - "IF function"
   - https://support.microsoft.com/en-us/office/if-function-69aed7c9-4e8a-4755-a9bc-aa8bbff73be2

4. Excel Easy - "Formulas and Functions"
   - https://www.excel-easy.com/introduction/formulas-functions.html

5. Excel Easy - "Functions"
   - https://www.excel-easy.com/functions.html

6. ExcelJet - "Excel Functions"
   - https://exceljet.net/functions

7. ExcelJet - "Excel Formulas"
   - https://exceljet.net/formulas

8. Corporate Finance Institute - "Excel Functions"
   - https://corporatefinanceinstitute.com/resources/excel/study/excel-functions/

9. GCFGlobal - "Excel 2016: Working with Multiple Worksheets"
   - https://edu.gcfglobal.org/en/excel/

10. Chandoo - "Excel Formulas"
    - https://chandoo.org/?s=excelexcel-formulas/

11. TrumpExcel - "Excel Formulas"
    - https://trumpexcel.com/excel-formulas/

12. Excel Easy - "VLOOKUP"
    - https://www.excel-easy.com/functions/lookup-reference-functions.html

13. Microsoft Support - "XLOOKUP function"
    - https://support.microsoft.com/en-us/excel

---

*Module 2 Complete. Review [Module 1: Excel Fundamentals ←](../01-Fundamentals/01-excel-fundamentals.md) or practice with the exercises above!*


## FILE: 03-Formatting-Visualization/03-formatting-visualization.md

# Module 3: Formatting & Visualization in Microsoft Excel

## Table of Contents

1. [Cell Formatting: Font, Size, Color, Borders, Fill Colors](#1-cell-formatting)
2. [Number Formats](#2-number-formats)
3. [Format Painter and Cell Styles](#3-format-painter-and-cell-styles)
4. [Conditional Formatting](#4-conditional-formatting)
5. [Custom Conditional Formatting with Formulas](#5-custom-conditional-formatting-with-formulas)
6. [Tables](#6-tables)
7. [Charts](#7-charts)
8. [Chart Elements](#8-chart-elements)
9. [Chart Formatting and Design Best Practices](#9-chart-formatting-and-design-best-practices)
10. [Sparklines](#10-sparklines)
11. [Slicers](#11-slicers)
12. [Dashboard Design Principles](#12-dashboard-design-principles)
13. [Practice Exercises](#13-practice-exercises)
14. [Video References](#14-video-references)
15. [Sources](#15-sources)

---

## 1. Cell Formatting

### 1.1 Font Formatting

Font formatting controls the appearance of text in cells. Access via **Home → Font** group or **Ctrl+1** (Format Cells dialog).

| Property | Description | Shortcut |
|----------|-------------|----------|
| **Font Family** | Typeface (Calibri, Arial, Times New Roman, etc.) | - |
| **Font Size** | Point size (default: 11pt) | - |
| **Bold** | Makes text thicker | **Ctrl+B** |
| **Italic** | Slants text | **Ctrl+I** |
| **Underline** | Adds underline | **Ctrl+U** |
| **Strikethrough** | Draws line through text | Ctrl+1 → Font → Strikethrough |
| **Font Color** | Text color | Home → Font Color dropdown |
| **Superscript/Subscript** | Raises/lowers text | Ctrl+1 → Font → Superscript/Subscript |

**Example:**
```
To make cell A1 bold with a 14pt Arial font in dark blue:
1. Select cell A1
2. Home → Font → Arial
3. Home → Font Size → 14
4. Ctrl+B (bold)
5. Home → Font Color → Dark Blue
```

### 1.2 Borders

Borders add lines around cells to separate and organize data.

**Access:** Home → Borders dropdown or Ctrl+1 → Border tab

| Border Type | Description |
|-------------|-------------|
| **Bottom Border** | Line below selected cells |
| **Top Border** | Line above selected cells |
| **Left Border** | Line on left side |
| **Right Border** | Line on right side |
| **All Borders** | Lines around every cell in selection |
| **Outside Borders** | Line around the outside of the selection |
| **Thick Box Border** | Thick line around the outside |
| **Bottom Double Border** | Double line below (accounting style) |

**Custom Borders (Ctrl+1 → Border tab):**
- Choose line style (thin, thick, double, dashed, dotted)
- Choose line color
- Click specific border positions in the preview area
- Use presets: None, Outline, Inside

**Example:**
```
To create a formatted header row:
1. Select row 1 (A1:E1)
2. Ctrl+1 → Border tab
3. Choose thick line style
4. Click "Outline" preset
5. Click OK
6. Apply Bold + Fill Color
```

### 1.3 Fill Colors (Background)

Fill colors change the background color of cells.

**Access:** Home → Fill Color dropdown

**Methods:**
1. **Quick Fill:** Home → Fill Color → select color
2. **More Colors:** Fill Color → More Colors → Custom tab (enter RGB values)
3. **Format Cells dialog:** Ctrl+1 → Fill tab → choose color
4. **Pattern Fill:** Ctrl+1 → Fill tab → Pattern Color and Pattern Style

**Pro Tip:** Use `Ctrl+1` → Fill tab to access pattern fills (diagonal stripes, dots, etc.) for printed reports.

---

## 2. Number Formats

Number formats change how numbers appear without changing the underlying value.

### 2.1 Built-in Number Formats

| Format | Input | Display | Description |
|--------|-------|---------|-------------|
| **General** | 1234.5 | 1234.5 | Default, no specific format |
| **Number** | 1234.5 | 1,234.50 | Thousands separator, 2 decimals |
| **Currency** | 1234.5 | $1,234.50 | Currency symbol, thousands separator |
| **Accounting** | 1234.5 | $ 1,234.50 | Currency aligned, parentheses for negatives |
| **Date** | 44927 | 1/1/2023 | Serial number to date |
| **Time** | 0.5 | 12:00:00 PM | Decimal to time |
| **Percentage** | 0.75 | 75% | Multiplied by 100, % symbol |
| **Fraction** | 0.75 | 3/4 | Decimal to fraction |
| **Scientific** | 1234.5 | 1.23E+03 | Scientific notation |
| **Text** | 1234.5 | 1234.5 | Treated as text |

**Access:** Home → Number group dropdown, or Ctrl+1 → Number tab

### 2.2 Custom Number Formats

Custom formats use special codes to control display. Access via **Ctrl+1 → Number → Custom**.

**Format Code Structure:** `positive;negative;zero;text`

#### Common Format Codes

| Code | Description | Example Input | Display |
|------|-------------|---------------|---------|
| `0` | Digit placeholder (shows zeros) | 45.2 | 45 |
| `#` | Digit placeholder (hides zeros) | 45.2 | 45 |
| `?` | Digit placeholder (aligns decimals) | 45.2 | 45.2 |
| `.` | Decimal point | `0.00` | 45.20 |
| `,` | Thousands separator | `#,##0` | 1,234 |
| `%` | Percentage (multiplies by 100) | `0%` | 75% |
| `$` | Dollar sign | `$#,##0` | $1,234 |
| `*` | Repeat character to fill cell | `@ *-` | text---- |
| `_` | Skip width of next character | `_( )` | space for alignment |
| `"text"` | Literal text | `0.0 "ft"` | 45.2 ft |
| `@` | Text placeholder | `@ is text` | Hello is text |
| `[Color]` | Apply color | `[Red]#,##0` | 1,234 in red |

#### Practical Custom Format Examples

| Format Code | Input | Display | Use Case |
|-------------|-------|---------|----------|
| `00000` | 41 | 00041 | Zip codes, product codes |
| `#,##0.00` | 1234567 | 1,234,567.00 | Financial reports |
| `$#,##0.00_);($#,##0.00)` | -500 | ($500.00) | Accounting |
| `0.0,, "M"` | 2500000 | 2.5 M | Millions display |
| `#,##0;[Red]-#,##0` | -500 | -500 in red | Red negatives |
| `0.0%` | 0.856 | 85.6% | Percentage with 1 decimal |
| `mmmm d, yyyy` | 44927 | January 1, 2023 | Full date |
| `hh:mm AM/PM` | 0.75 | 6:00 PM | Time display |
| `[Green]$#,##0;[Red]$(#,##0);"Zero"` | 0 | Zero | Multi-condition |

**Creating a Custom Format:**
```
1. Select the cell(s)
2. Press Ctrl+1
3. Click Number tab → Custom
4. In the Type box, enter your format code
5. Preview in Sample box
6. Click OK
```

---

## 3. Format Painter and Cell Styles

### 3.1 Format Painter

The Format Painter copies formatting from one cell and applies it to another.

**How to Use:**
```
1. Select the cell with the formatting you want to copy
2. On the Home tab, in the Clipboard group, click Format Painter
   (A moving dashed border appears around the source cell)
3. Click the destination cell (or drag across a range)
4. Formatting is applied to the destination
```

**Double-Click Trick:** Double-click the Format Painter button to keep it active. You can then apply the same formatting to multiple cells/ranges. Press **Esc** or click Format Painter again to deactivate.

**What Format Painter Copies:**
- Font formatting (type, size, bold, italic, color)
- Number format
- Alignment
- Borders
- Fill color
- Cell protection settings

### 3.2 Cell Styles

Cell Styles are predefined formatting combinations that you can apply with one click.

**Access:** Home → Styles group → Cell Styles

**Built-in Style Categories:**
- **Good, Bad, Neutral** - green, red, gray backgrounds
- **Data and Model** - input, calculation, check cell, linked cell
- **Titles and Headings** - Heading 1-4, Title, Total
- **Themed Cell Styles** - Accent 1-6 (20% and 40% lighter/darker variants)
- **Number Format** - Comma, Currency, Percent

**Creating a Custom Cell Style:**
```
1. Home → Styles → Cell Styles → New Cell Style
2. Enter a Style name (e.g., "My Header")
3. Click Format button
4. Set Number, Alignment, Font, Border, Fill, Protection
5. Uncheck any categories you don't want to control
6. Click OK twice
7. Your style now appears in the Cell Styles gallery
```

**Modifying a Style:** Right-click any cell style → Modify. Changes apply to ALL cells using that style across the workbook.

**Merging Styles:** You can import styles from another open workbook:
```
1. Open both workbooks
2. Home → Styles → Cell Styles → Merge Styles
3. Select the source workbook
4. Click OK
```

---

## 4. Conditional Formatting

Conditional formatting automatically applies formatting based on cell values or formulas.

**Access:** Home → Styles → Conditional Formatting

### 4.1 Highlight Cells Rules

Highlight cells that meet specific criteria:

| Rule | Description | Example |
|------|-------------|---------|
| **Greater Than** | Values above a threshold | Highlight sales > $1000 |
| **Less Than** | Values below a threshold | Highlight inventory < 10 |
| **Between** | Values within a range | Highlight scores between 80-100 |
| **Equal To** | Exact match | Highlight status = "Complete" |
| **Text That Contains** | Cells containing specific text | Highlight cells with "Urgent" |
| **A Date Occurring** | Date-based rules | Highlight dates in "Last Week" |
| **Duplicate Values** | Find duplicates or unique values | Highlight duplicate entries |

**Example: Highlight cells greater than 80**
```
1. Select the range A1:A10
2. Home → Conditional Formatting → Highlight Cells Rules → Greater Than
3. Enter: 80
4. Choose formatting style (e.g., Light Red Fill with Dark Red Text)
5. Click OK
```

### 4.2 Top/Bottom Rules

| Rule | Description | Example |
|------|-------------|---------|
| **Top 10 Items** | Highlight highest N values | Top 10 sales figures |
| **Top 10 %** | Highlight top N percentage | Top 10% of performers |
| **Bottom 10 Items** | Highlight lowest N values | Bottom 10 scores |
| **Bottom 10 %** | Highlight bottom N percentage | Bottom 10% of results |
| **Above Average** | Values above the mean | Above-average grades |
| **Below Average** | Values below the mean | Below-average grades |

**Example: Highlight above average**
```
1. Select the range A1:A10
2. Home → Conditional Formatting → Top/Bottom Rules → Above Average
3. Choose formatting style
4. Click OK
5. Excel calculates the average (e.g., 42.5) and highlights cells above it
```

### 4.3 Data Bars

Data bars add a visual bar inside each cell, proportional to the value.

```
1. Select the range
2. Home → Conditional Formatting → Data Bars
3. Choose a color style (gradient or solid fill)
4. Bars automatically size relative to min/max values
```

**Data Bar Options (More Rules):**
- Minimum/Maximum type (Lowest/Highest value, Number, Percent, Percentile, Formula)
- Bar color and fill (solid, gradient)
- Show bar only (hide the number)
- Bar direction (left-to-right, right-to-left)
- Negative value handling (separate color, axis)

### 4.4 Color Scales

Color scales apply a two-color or three-color gradient based on cell values.

```
1. Select the range
2. Home → Conditional Formatting → Color Scales
3. Choose a preset:
   - Green-Yellow-Red (3-color)
   - Green-White-Red (3-color)
   - Red-Yellow-Green (3-color)
   - White-Red (2-color)
   - Green-White (2-color)
   - Green-Yellow (2-color)
```

**Use Cases:**
- Heatmaps: Red = high values, Green = low values
- Performance tracking: Color intensity shows relative performance
- Geographic data: Color variation by region/metric

### 4.5 Icon Sets

Icon sets add small icons (arrows, shapes, indicators) to cells based on values.

```
1. Select the range
2. Home → Conditional Formatting → Icon Sets
3. Choose a style:
   - Directional: Arrows (3, 4, or 5 arrows)
   - Shapes: Traffic lights, circles, flags
   - Indicators: Check marks, crosses, exclamation
   - Ratings: Stars, ratings
```

**Customizing Icon Sets (More Rules):**
```
1. Home → Conditional Formatting → Icon Sets → More Rules
2. Set Type to "Percent" or "Number"
3. Set values for each icon threshold:
   - Green arrow: >= 67%
   - Yellow arrow: >= 33%
   - Red arrow: < 33%
4. Reverse icon order if needed
5. Show icon only (hide the number)
```

### 4.6 Managing Conditional Formatting Rules

```
1. Home → Conditional Formatting → Manage Rules
2. The Rules Manager shows all rules for the current selection
3. You can:
   - Add new rules
   - Edit existing rules
   - Delete rules
   - Reorder rules (higher rules take priority)
   - Change the "Applies to" range
```

**Rule Priority:** Rules are evaluated top-to-bottom. Check "Stop If True" to prevent lower-priority rules from applying.

---

## 5. Custom Conditional Formatting with Formulas

Use formulas to create powerful, flexible conditional formatting rules.

**Access:** Home → Conditional Formatting → New Rule → "Use a formula to determine which cells to format"

### Key Rules for Formula-Based Conditional Formatting:
1. The formula must evaluate to **TRUE** or **FALSE**
2. Write the formula for the **upper-left cell** in the selected range
3. Excel automatically copies the formula to other cells (relative references adjust)
4. Use absolute ($) references to lock specific cells

### Common Formula Examples

#### Highlight Entire Row Based on a Cell Value
```
Formula: =$C2="Completed"
Applied to: $A$2:$F$100

Explanation: The $ before C locks the column. As Excel moves
across columns, it still checks column C. The row reference
(2) is relative, so it adjusts for each row.
```

#### Highlight Alternating Rows (Zebra Striping)
```
Formula: =MOD(ROW(),2)=0
Applied to: $A$1:$F$100

Explanation: MOD(ROW(),2)=0 returns TRUE for even rows.
Apply a light gray fill for readability.
```

#### Highlight Cells with Errors
```
Formula: =ISERROR(A1)
Applied to: $A$1:$F$100

Explanation: ISERROR returns TRUE for #N/A, #VALUE!, #REF!, etc.
```

#### Highlight Dates Older Than 30 Days
```
Formula: =TODAY()-A1>30
Applied to: $A$1:$A$100

Explanation: Calculates the difference between today and the date.
```

#### Highlight Top N Values
```
Formula: =A1>=LARGE($A$1:$A$100,5)
Applied to: $A$1:$A$100

Explanation: LARGE returns the 5th largest value.
Cells >= that value get formatted.
```

#### Highlight Cells Based on Another Cell
```
Formula: =$B1>$C$1
Applied to: $A$1:$A$100

Explanation: Compares each cell in column A to a
threshold value in cell C1.
```

#### Highlight Weekends in a Date Column
```
Formula: =OR(WEEKDAY(A1,2)=6, WEEKDAY(A1,2)=7)
Applied to: $A$1:$A$31

Explanation: WEEKDAY with return_type 2 returns
6 for Saturday and 7 for Sunday.
```

#### Dynamic Conditional Formatting with Data Validation
```
Formula: =A1>=$E$1
Applied to: $A$1:$A$100

Where E1 contains a dropdown (Data Validation list).
Change the dropdown to dynamically update formatting.
```

---

## 6. Tables

Excel Tables are structured data ranges with built-in features for sorting, filtering, formatting, and formulas.

### 6.1 Creating a Table

```
1. Click any cell inside your data
2. On the Insert tab, in the Tables group, click Table
   (or press Ctrl+T)
3. Excel auto-detects the data range
4. Check "My table has headers" if your first row contains headers
5. Click OK
```

**Keyboard Shortcut:** `Ctrl+T` creates a table instantly.

### 6.2 Table Features

| Feature | Description |
|---------|-------------|
| **Auto-filter** | Drop-down filters in every header |
| **Banding** | Alternating row colors for readability |
| **Total Row** | Built-in SUBTOTAL calculations |
| **Auto-expand** | Table grows automatically when you add data |
| **Structured References** | Use column names in formulas |
| **Calculated Columns** | Formula fills down automatically |
| **Table Name** | Reference the entire table by name |

### 6.3 Structured References

Tables use structured references instead of cell addresses.

| Reference Style | Example | Description |
|-----------------|---------|-------------|
| **Column Name** | `[@Sales]` | Current row's Sales value |
| **Full Column** | `[Sales]` | Entire Sales column |
| **Table + Column** | `Table1[Sales]` | Sales column from Table1 |
| **Table + Column + @** | `Table1[@Sales]` | Current row's Sales in Table1 |
| **#Headers** | `Table1[#Headers]` | Header row of Table1 |
| **#All** | `Table1[#All]` | Entire table including headers |
| **#Data** | `Table1[#Data]` | Data rows only (no header) |
| **#Totals** | `Table1[#Totals]` | Total row only |

**Example Formulas:**
```
=SUM(Table1[Sales])              Sum of all Sales
=AVERAGE(Table1[Amount])         Average of Amount column
=COUNTA(Table1[Product])         Count of Product entries
=Table1[@Price]*Table1[@Qty]     Price × Quantity in current row
```

### 6.4 Table Styles

```
1. Click inside the table
2. On the Table Design tab, in the Table Styles group:
   - Choose from Light, Medium, or Dark styles
   - Hover to preview
   - Click to apply
3. In Table Style Options:
   ☑ Header Row     - Show/hide header formatting
   ☑ Total Row      - Show/hide total row
   ☑ Banded Rows    - Alternating row colors
   ☑ First Column   - Bold first column
   ☑ Last Column    - Bold last column
   ☑ Banded Columns - Alternating column colors
   ☑ Filter Button  - Show/hide filter dropdowns
```

### 6.5 Total Row

```
1. Click inside the table
2. Table Design → Table Style Options → check Total Row
   (or press Ctrl+Shift+T)
3. Click any cell in the total row
4. Choose a function from the dropdown:
   - Average, Count, Count Numbers, Max, Min, Sum
   - StdDev, Var, More Functions...
```

**Important:** The Total Row uses the SUBTOTAL function (not SUM), which correctly handles filtered data.

### 6.6 Converting Table Back to Range

```
1. Click inside the table
2. Table Design → Tools → Convert to Range
3. Click Yes
```

---

## 7. Charts

### 7.1 Chart Types Overview

| Chart Type | Best For | Example Use |
|------------|----------|-------------|
| **Column** | Comparing values across categories | Sales by product |
| **Bar** | Comparing values (horizontal) | Rankings, long category names |
| **Line** | Trends over time | Monthly revenue trend |
| **Pie** | Part-to-whole relationships | Market share (≤7 slices) |
| **Scatter (XY)** | Correlation between two variables | Height vs. weight |
| **Combo** | Mixed data types on different axes | Revenue (columns) + growth % (line) |
| **Area** | Volume over time | Cumulative sales |
| **Stock** | High-Low-Close financial data | Stock prices |
| **Radar** | Comparison across categories | Feature comparison |
| **Treemap** | Hierarchical data | Sales by region/category |
| **Sunburst** | Multi-level hierarchy | Budget allocation |

### 7.2 Creating a Chart

```
1. Select your data range (including headers)
2. On the Insert tab, in the Charts group:
   - Click Recommended Charts to see Excel's suggestions
   - Or click a specific chart type dropdown
3. Choose a chart subtype
4. The chart appears on your worksheet
5. Use Chart Design and Format tabs to customize
```

### 7.3 Column Charts

**Clustered Column:** Compare values across categories
```
Data Layout:
| Product   | Q1    | Q2    | Q3    | Q4    |
|-----------|-------|-------|-------|-------|
| Widgets   | 1200  | 1500  | 1300  | 1800  |
| Gadgets   | 800   | 900   | 1100  | 1200  |
| Gizmos    | 500   | 600   | 700   | 900   |
```

**Stacked Column:** Show part-to-whole across categories
**100% Stacked Column:** Compare proportions (each column = 100%)

### 7.4 Bar Charts

Bar charts are horizontal columns - ideal when:
- Category labels are long
- You're comparing rankings
- You want a different visual style

### 7.5 Line Charts

**Best for:** Showing trends over time

```
Data Layout:
| Month    | Sales  | Expenses |
|----------|--------|----------|
| Jan      | 10000  | 8000     |
| Feb      | 12000  | 8500     |
| Mar      | 11000  | 9000     |
| Apr      | 15000  | 8200     |
```

**Subtypes:**
- **Line:** Basic trend
- **Line with Markers:** Points at each data value
- **Stacked Line:** Cumulative trends
- **100% Stacked Line:** Proportional trends

### 7.6 Pie Charts

**Best for:** Showing composition (part-to-whole)

**Rules:**
- Use for ≤7 slices (more becomes hard to read)
- All values should be positive
- Slices should sum to a meaningful whole
- Consider a Bar/Column chart if slices are similar in size

**Subtypes:**
- **Pie:** Standard 2D
- **3D Pie:** Adds depth (avoid - distorts perception)
- **Pie of Pie:** Separates small slices into a second pie
- **Bar of Pie:** Separates small slices into a bar chart
- **Doughnut:** Multiple rings for multiple data series

### 7.7 Scatter (XY) Charts

**Best for:** Showing correlation between two variables

```
Data Layout:
| Height | Weight |
|--------|--------|
| 60     | 120    |
| 65     | 150    |
| 70     | 180    |
| 72     | 200    |
```

**Use Cases:**
- Correlation analysis
- Regression visualization
- Scientific data
- Outlier detection

### 7.8 Combo Charts

Combo charts combine two or more chart types on the same axes.

```
Example: Sales (columns) + Profit Margin (line on secondary axis)

1. Select data
2. Insert → Recommended Charts → All Charts → Combo
3. For each data series, choose:
   - Chart type (Column, Line, Area, etc.)
   - Which axis (Primary or Secondary)
4. Click OK
```

**Common Combo:** Clustered Column + Line with Secondary Axis

---

## 8. Chart Elements

### 8.1 Adding Chart Elements

```
1. Click the chart to select it
2. Click the + (plus) icon next to the chart
3. Check/uncheck elements:
   ☑ Chart Title
   ☑ Axis Titles
   ☑ Legend
   ☑ Data Labels
   ☑ Data Table
   ☑ Gridlines
   ☑ Trendline
   ☑ Error Bars
   ☑ Up/Down Bars
   ☑ Lines (Drop Lines, High-Low Lines)
```

### 8.2 Chart Title

```
1. Click the + icon → Chart Title
2. Click the title text to edit
3. Options:
   - Above Chart (default)
   - Centered Overlay (overlaps chart)
   - Use formula: Click title → type = → click cell
```

### 8.3 Axis Titles

```
1. Click + → Axis Titles
2. Click each axis title to edit
3. For 3D charts: also a Depth Axis title
```

### 8.4 Legend

```
1. Click + → Legend
2. Positions: Top, Bottom, Left, Right, Top Right
3. Click and drag to reposition
4. Right-click legend → Select Data to control which series appear
```

### 8.5 Data Labels

```
1. Click + → Data Labels
2. Positions: Center, Inside End, Inside Base, Outside End
3. Right-click labels → Format Data Labels:
   ☑ Value
   ☑ Category Name
   ☑ Series Name
   ☑ Percentage (for pie charts)
   ☑ Separator (comma, semicolon, newline)
```

### 8.6 Trendlines

```
1. Click the data series in the chart
2. Click + → Trendline
3. Choose type:
   - Linear: Straight line (y = mx + b)
   - Exponential: Growth curve (y = ae^(bx))
   - Logarithmic: Log curve (y = a*ln(x) + b)
   - Polynomial: Curve (y = a + bx + cx² + ...)
   - Power: Power curve (y = ax^b)
   - Moving Average: Smoothed trend
4. Right-click trendline → Format Trendline:
   ☑ Display Equation on chart
   ☑ Display R-squared value
   ☑ Set Forecast periods (forward/backward)
```

### 8.7 Error Bars

```
1. Click the data series
2. Click + → Error Bars
3. Options:
   - Standard Error
   - Percentage (e.g., 5%)
   - Standard Deviation (e.g., 1 SD)
   - Fixed value
   - Custom (specify + and - values from cells)
```

---

## 9. Chart Formatting and Design Best Practices

### 9.1 Design Principles

| Principle | Guidance |
|-----------|----------|
| **Simplify** | Remove unnecessary elements (gridlines, borders, 3D effects) |
| **Contrast** | Use contrasting colors for different data series |
| **Consistency** | Same colors/fonts across all charts in a report |
| **Label clearly** | Always include titles and axis labels |
| **Avoid distortion** | Start Y-axis at zero (unless showing small changes) |
| **Right chart type** | Match chart type to the story you're telling |
| **Color-blind friendly** | Don't rely on color alone; use patterns or labels |
| **White space** | Don't overcrowd; let data breathe |

### 9.2 Formatting Tips

```
1. Remove gridlines: Click gridlines → Delete
2. Change colors: Chart Design → Change Colors
3. Format data series: Right-click → Format Data Series
4. Format axis: Right-click axis → Format Axis
   - Set min/max bounds
   - Change number format
   - Adjust tick marks
5. Format plot area: Right-click → Format Plot Area
   - Add/remove border
   - Fill color
6. Add shapes/text: Insert → Shapes/Text Box
```

### 9.3 Common Chart Mistakes to Avoid

| Mistake | Fix |
|---------|-----|
| Using 3D charts | Use 2D for accurate comparison |
| Too many data series | Limit to 4-5 series per chart |
| Pie chart with too many slices | Group small slices into "Other" |
| Y-axis doesn't start at zero | Start at zero unless showing small changes |
| Missing axis labels | Always label both axes |
| Inconsistent colors | Use the same color for the same category |
| No title | Always include a descriptive title |
| Misleading scale | Use consistent intervals |

---

## 10. Sparklines

Sparklines are small, cell-sized charts that show data trends in a compact format.

**Access:** Insert → Sparklines group

### 10.1 Line Sparklines

```
1. Select where you want the sparkline (e.g., E2:E10)
2. Insert → Sparklines → Line
3. Data Range: The values (e.g., A2:D10)
4. Location Range: Where to place sparklines (e.g., E2:E10)
5. Click OK
```

### 10.2 Column Sparklines

```
1. Insert → Sparklines → Column
2. Same process as line sparklines
3. Each column in the sparkline represents one data point
```

### 10.3 Win/Loss Sparklines

```
1. Insert → Sparklines → Win/Loss
2. Shows positive (wins) and negative (losses) values
3. Ideal for: Game scores, profit/loss, +/- changes
```

### 10.4 Customizing Sparklines

```
1. Click the sparkline(s) to select
2. Sparkline Design tab appears:
   - Edit Data: Change data source or location
   - Sparkline Color: Change line/bar color
   - Marker Color: Highlight High Point, Low Point,
     First Point, Last Point, Negative Points, Markers
   - Style: Choose from preset styles
   - Axis: Change min/max bounds, date axis
   - Group: Group/ungroup sparklines for consistent scaling
   - Clear: Remove sparklines
```

**Key Options:**
```
☑ High Point    - Highlights the maximum value
☑ Low Point     - Highlights the minimum value
☑ First Point   - Highlights the first data point
☑ Last Point    - Highlights the last data point
☑ Negative Points - Highlights negative values
☑ Markers       - Shows all data points
```

---

## 11. Slicers

Slicers are visual filter buttons that make it easy to filter data in tables and PivotTables.

### 11.1 Slicers for Tables

```
1. Click inside the table
2. On the Table Design tab, click Insert Slicer
3. Check the columns you want slicers for
4. Click OK
5. Slicers appear as floating panels
6. Click buttons to filter
7. Ctrl+click to select multiple items
8. Click the clear button (×) to remove filter
```

### 11.2 Slicers for PivotTables

```
1. Click inside the PivotTable
2. On the PivotTable Analyze tab, click Insert Slicer
3. Check the fields you want slicers for
4. Click OK
5. Use slicers to filter the PivotTable interactively
```

### 11.3 Formatting Slicers

```
1. Click the slicer to select it
2. Slicer tab appears:
   - Slicer Settings: Name, caption, item sorting
   - Slicer Styles: Choose from preset styles
   - Buttons: Set number of columns, button height/width
   - Size: Set exact dimensions
   - Arrange: Position, alignment, grouping
```

### 11.4 Connecting Slicers to Multiple PivotTables

```
1. Click the slicer
2. On the Slicer tab, click Report Connections
   (or right-click → Report Connections)
3. Check all PivotTables you want connected
4. Click OK
5. Now filtering the slicer affects all connected PivotTables
```

---

## 12. Dashboard Design Principles

### 12.1 What is a Dashboard?

A dashboard is a single-page visual display of the most important information, designed for at-a-glance monitoring.

### 12.2 Design Principles

| Principle | Description |
|-----------|-------------|
| **Know your audience** | Executive vs. operational - adjust detail level |
| **Single page** | Fits on one screen; no scrolling |
| **Key metrics first** | Most important KPIs in the top-left (reading order) |
| **Consistent formatting** | Same fonts, colors, sizes throughout |
| **Minimal text** | Use numbers and visuals, not paragraphs |
| **Color coding** | Red = bad, Yellow = caution, Green = good |
| **Interactive elements** | Slicers, dropdowns, buttons for filtering |
| **White space** | Don't overcrowd - give elements breathing room |
| **Actionable** | Shows what needs attention, not just what happened |

### 12.3 Dashboard Building Blocks in Excel

| Element | Use For |
|---------|---------|
| **PivotTables** | Summarizing large datasets |
| **PivotCharts** | Visual representation of PivotTable data |
| **Sparklines** | Inline trend indicators |
| **Conditional Formatting** | Color-coded KPIs |
| **Slicers** | Interactive filters |
| **Data Validation Dropdowns** | User-controlled parameters |
| **Named Ranges** | Dynamic data sources |
| **CHOOSE/INDEX functions** | Dynamic chart data selection |

### 12.4 Dashboard Layout Tips

```
┌─────────────────────────────────────────────┐
│  KPI 1        KPI 2        KPI 3        KPI 4│
│  $1.2M        847          92%          ↑12% │
├─────────────────────────────────────────────┤
│  [Slicer 1]  [Slicer 2]  [Slicer 3]        │
├────────────────────────┬────────────────────┤
│                        │                    │
│   Chart 1              │   Chart 2          │
│   (Main metric)        │   (Secondary)      │
│                        │                    │
├────────────────────────┼────────────────────┤
│                        │                    │
│   Chart 3              │   Table            │
│   (Trend)              │   (Details)        │
│                        │                    │
└────────────────────────┴────────────────────┘
```

### 12.5 KPI Cards in Excel

Create KPI summary cards using cell formatting:

```
1. Merge cells for the card area
2. Add a subtle border
3. Large font for the KPI value
4. Small font for the KPI label
5. Use conditional formatting for indicators:
   - Green: On target
   - Yellow: Close to target
   - Red: Below target
6. Add a small sparkline for trend
```

---

## 13. Practice Exercises

### Exercise 1: Cell Formatting & Number Formats

**Objective:** Format a sales report with proper number formatting.

**Sample Data:**
| A | B | C | D |
|---|---|---|---|
| Date | Product | Units Sold | Revenue |
| 1/15/2024 | Widget A | 150 | 4500.5 |
| 2/20/2024 | Widget B | 200 | 8000.75 |
| 3/10/2024 | Widget C | 75 | 2250.25 |
| 4/5/2024 | Widget A | 300 | 9000 |
| 5/18/2024 | Widget B | 125 | 5000.5 |

**Tasks:**
1. Format column A as Date (MM/DD/YYYY)
2. Format column D as Currency with 2 decimal places
3. Bold the header row and add a bottom border
4. Apply a fill color to the header row (e.g., dark blue with white text)
5. Auto-fit all columns
6. Apply a custom number format to column D that shows positive values in green and negative in red

### Exercise 2: Conditional Formatting

**Objective:** Apply various conditional formatting rules to highlight important data.

**Sample Data:**
| A | B | C | D |
|---|---|---|---|
| Student | Math | Science | English |
| Alice | 92 | 88 | 95 |
| Bob | 78 | 82 | 71 |
| Carol | 95 | 97 | 89 |
| Dave | 65 | 58 | 72 |
| Eve | 88 | 91 | 84 |
| Frank | 45 | 52 | 61 |

**Tasks:**
1. Highlight cells >90 with green fill, <60 with red fill
2. Apply a color scale (green-yellow-red) to all scores
3. Add data bars to column B (Math scores)
4. Highlight the top 2 scores in each column
5. Create a formula-based rule to highlight entire row if any score < 60
6. Add icon sets (traffic lights) based on: ≥90 green, ≥70 yellow, <70 red

### Exercise 3: Tables & Structured References

**Objective:** Convert data to a table and use structured references.

**Sample Data:**
| A | B | C | D | E |
|---|---|---|---|---|
| OrderID | Product | Category | Price | Quantity |
| 1001 | Laptop | Electronics | 999.99 | 5 |
| 1002 | Desk Chair | Furniture | 249.50 | 10 |
| 1003 | Monitor | Electronics | 399.99 | 8 |
| 1004 | Keyboard | Electronics | 79.99 | 20 |
| 1005 | Bookshelf | Furniture | 149.00 | 3 |

**Tasks:**
1. Convert to a Table (Ctrl+T) named "Orders"
2. Add a Total Row showing the SUM of Quantity
3. Add a calculated column "Total" = Price × Quantity (use structured references)
4. Apply a Medium style table format with banded rows
5. Use the table name in a formula: `=SUM(Orders[Total])`
6. Insert slicers for Category

### Exercise 4: Charts & Visualization

**Objective:** Create and format various chart types.

**Sample Data:**
| A | B | C | D | E |
|---|---|---|---|---|
| Month | Sales | Expenses | Profit | Margin% |
| Jan | 50000 | 35000 | 15000 | 30% |
| Feb | 55000 | 36000 | 19000 | 35% |
| Mar | 48000 | 34000 | 14000 | 29% |
| Apr | 62000 | 38000 | 24000 | 39% |
| May | 70000 | 42000 | 28000 | 40% |
| Jun | 65000 | 40000 | 25000 | 38% |

**Tasks:**
1. Create a clustered column chart for Sales and Expenses
2. Add a line for Profit on a secondary axis (Combo chart)
3. Add a trendline to the Sales series (linear)
4. Format the chart: title, axis labels, legend, data labels
5. Create a pie chart showing expense breakdown by month
6. Add sparklines in column F showing the trend for each metric
7. Apply conditional formatting with data bars to the Profit column

### Exercise 5: Dashboard Building

**Objective:** Create a mini dashboard combining multiple visual elements.

**Steps:**
1. Start with the sample data from Exercise 4
2. Create KPI cards at the top showing:
   - Total Sales (large number)
   - Average Margin %
   - Best Month
   - Worst Month
3. Add a Combo chart (columns + line)
4. Add sparklines for each metric
5. Add a slicer or dropdown for filtering
6. Apply consistent formatting (colors, fonts, spacing)
7. Ensure it fits on one screen

---

## 14. Video References

### Conditional Formatting
- **Excel Conditional Formatting in 10 Minutes** - https://www.youtube.com/watch?v=yVixz6pMk8Q
- **Conditional Formatting with Formulas** - https://www.youtube.com/watch?v=9NUjHBNWe2E
- **Excel Conditional Formatting - Complete Guide** - https://www.youtube.com/watch?v=7Fj4Qb_b7Tw

### Charts & Visualization
- **Excel Charts - The Complete Guide** - https://www.youtube.com/watch?v=ewYorxOGwTk
- **Excel Combo Charts Tutorial** - https://www.youtube.com/watch?v=3k7ZsT3Gk5Y
- **Excel Chart Formatting Tips** - https://www.youtube.com/watch?v=K79UMRWKDF8

### Tables
- **Excel Tables - Complete Guide** - https://www.youtube.com/watch?v=2G4r2r_1T9Y
- **Structured References in Excel Tables** - https://www.youtube.com/watch?v=3b4nF1BFwBg

### Sparklines & Slicers
- **Excel Sparklines Tutorial** - https://www.youtube.com/watch?v=9vJRopuVJ9Q
- **Excel Slicers - Tables & PivotTables** - https://www.youtube.com/watch?v=Zb2h2nUOZuE

### Dashboard Design
- **Excel Dashboard Tutorial** - https://www.youtube.com/watch?v=M2NH_PQuWnE
- **Excel Dashboard Design Tips** - https://www.youtube.com/watch?v=LjWPUqEBZ2E

### Number Formatting
- **Custom Number Formats in Excel** - https://www.youtube.com/watch?v=2f1mSq2sGMk
- **Excel Format Cells Tutorial** - https://www.youtube.com/watch?v=VlIHvXNJKUQ

---

## 15. Sources

### Primary Sources

1. **Microsoft Support - Format Cells**
   https://support.microsoft.com/en-us/excel

2. **Microsoft Support - Conditional Formatting**
   https://support.microsoft.com/en-us/office/use-conditional-formatting-to-highlight-information-fed60dfa-1d3f-4e13-9ecb-f1951ff89d7f

3. **Microsoft Support - Excel Tables**
   https://support.microsoft.com/en-us/excel

4. **Microsoft Support - Charts**
   https://support.microsoft.com/en-us/excel

5. **Microsoft Support - Sparklines**
   https://support.microsoft.com/en-us/excel

6. **Microsoft Support - Slicers**
   https://support.microsoft.com/en-us/office/use-slicers-to-filter-data-249f966b-a9d5-4b0f-b31a-12651785d29d

7. **Excel Easy - Format Cells**
   https://www.excel-easy.com/basics/format-cells.html

8. **Excel Easy - Custom Number Format**
   https://www.excel-easy.com/examples/custom-number-format.html

9. **Excel Easy - Format Painter**
   https://www.excel-easy.com/examples/format-painter.html

10. **Excel Easy - Cell Styles**
    https://www.excel-easy.com/examples/cell-styles.html

11. **Excel Easy - Conditional Formatting**
    https://www.excel-easy.com/data-analysis/conditional-formatting.html

12. **Excel Easy - Tables**
    https://www.excel-easy.com/data-analysis/tables.html

13. **Excel Easy - Charts**
    https://www.excel-easy.com/

14. **Exceljet - Conditional Formatting**
    https://exceljet.net/conditional-formatting

15. **Exceljet - Excel Charts**
    https://exceljet.net/charts

16. **Exceljet - Pivot Table Examples**
    https://exceljet.net/pivot-tables

17. **CFI - Conditional Formatting**
    https://corporatefinanceinstitute.com/resources/excel/conditional-formatting/

18. **Chandoo - Excel Charts**
    https://chandoo.org/?s=excel+charts

19. **Contextures - Conditional Formatting**
    https://www.contextures.com/

20. **Peltier Tech Blog - Chart Types**
    https://peltiertech.com/

---

*Last Updated: September 2026*


## FILE: 04-Data-Analysis/04-data-analysis.md

# Module 4: Data Analysis

> **Level:** Beginner to Intermediate | **Estimated Time:** 4–5 hours | **Prerequisites:** Modules 1–3 (Fundamentals, Formulas & Functions, Formatting & Visualization)

---

## Table of Contents

1. [Sorting Data](#1-sorting-data)
2. [Filtering Data](#2-filtering-data)
3. [Advanced Filtering](#3-advanced-filtering)
4. [Data Validation](#4-data-validation)
5. [What-If Analysis](#5-what-if-analysis)
6. [PivotTables (Detailed)](#6-pivottables-detailed)
7. [Subtotals](#7-subtotals)
8. [Data Consolidation](#8-data-consolidation)
9. [Practice Exercises](#9-practice-exercises)
10. [Video References](#10-video-references)
11. [Sources](#11-sources)

---

## Sample Dataset

Throughout this module, we'll use the following **Sales Data** table. Create this in a new workbook on a sheet named `SalesData` to follow along:

| OrderID | Date       | Region  | SalesRep  | Product     | Category    | Units | UnitPrice | Total   |
|---------|------------|---------|-----------|-------------|-------------|-------|-----------|---------|
| 1001    | 2025-01-15 | North   | Alice     | Widget A    | Electronics | 25    | 10.00     | 250.00  |
| 1002    | 2025-01-20 | South   | Bob       | Widget B    | Hardware    | 40    | 8.50      | 340.00  |
| 1003    | 2025-02-03 | East    | Charlie   | Gadget X    | Electronics | 15    | 25.00     | 375.00  |
| 1004    | 2025-02-14 | West    | Diana     | Gadget Y    | Software    | 60    | 15.00     | 900.00  |
| 1005    | 2025-03-01 | North   | Alice     | Widget A    | Electronics | 30    | 10.00     | 300.00  |
| 1006    | 2025-03-10 | South   | Eve       | Widget C    | Hardware    | 20    | 12.00     | 240.00  |
| 1007    | 2025-03-22 | East    | Charlie   | Widget B    | Hardware    | 55    | 8.50      | 467.50  |
| 1008    | 2025-04-05 | West    | Frank     | Gadget X    | Electronics | 35    | 25.00     | 875.00  |
| 1009    | 2025-04-18 | North   | Bob       | Gadget Y    | Software    | 45    | 15.00     | 675.00  |
| 1010    | 2025-05-02 | South   | Alice     | Widget A    | Electronics | 50    | 10.00     | 500.00  |
| 1011    | 2025-05-15 | East    | Diana     | Widget C    | Hardware    | 25    | 12.00     | 300.00  |
| 1012    | 2025-06-01 | West    | Eve       | Gadget Y    | Software    | 70    | 15.00     | 1050.00 |
| 1013    | 2025-06-20 | North   | Charlie   | Gadget X    | Electronics | 20    | 25.00     | 500.00  |
| 1014    | 2025-07-04 | South   | Frank     | Widget A    | Electronics | 65    | 10.00     | 650.00  |
| 1015    | 2025-07-18 | East    | Bob       | Widget B    | Hardware    | 30    | 8.50      | 255.00  |
| 1016    | 2025-08-01 | West    | Alice     | Gadget X    | Electronics | 40    | 25.00     | 1000.00 |
| 1017    | 2025-08-15 | North   | Diana     | Widget C    | Hardware    | 35    | 12.00     | 420.00  |
| 1018    | 2025-09-03 | South   | Eve       | Gadget Y    | Software    | 80    | 15.00     | 1200.00 |
| 1019    | 2025-09-20 | East    | Frank     | Widget A    | Electronics | 45    | 10.00     | 450.00  |
| 1020    | 2025-10-05 | West    | Charlie   | Widget B    | Hardware    | 25    | 8.50      | 212.50  |

**Tip:** Type this data into Excel starting at cell `A1`. Use `=G2*H2` in the Total column (cell I2) and copy down to auto-calculate totals.

---

## 1. Sorting Data

Sorting rearranges rows based on the values in one or more columns. It does **not** change your data - it reorders it.

### 1.1 Single-Column Sort

**Quick Method:**
1. Click any cell in the column you want to sort by (e.g., click on a cell in the `Total` column)
2. Go to **Data** tab → **Sort & Filter** group
3. Click **Sort A to Z** (ascending ↑) or **Sort Z to A** (descending ↓)

**Example:** Sort by Total (highest to lowest):
- Click any cell in column I (Total)
- Data → Sort Z to A
- Result: Order 1018 ($1,200.00) appears first, Order 1020 ($212.50) appears last

> **Keyboard Shortcuts:**
> - `Alt + D, S, S` - Quick ascending sort
> - `Alt + D, S, O` - Quick descending sort
> - `Ctrl + Shift + L` - Toggle AutoFilter on/off

### 1.2 Multi-Level Sort

When two rows have the same value in the primary sort column, a secondary sort breaks the tie.

**Steps:**
1. Click any cell in your data range
2. Go to **Data** tab → **Sort** (the larger Sort button, not the A→Z icon)
3. The Sort dialog box opens:
   - **Sort by:** Select `Region` - Sort On: `Cell Values` - Order: `A to Z`
   - Click **Add Level**
   - **Then by:** Select `Total` - Sort On: `Cell Values` - Order: `Largest to Smallest`
4. Click **OK**

**Result:** All East records are grouped together (sorted by Total descending within East), then North, South, West.

### 1.3 Custom Sort Order

For data that doesn't sort alphabetically (e.g., days of the week, months, priority levels):

**Steps:**
1. Data → Sort
2. For the Order dropdown, select **Custom List...**
3. Choose an existing list (e.g., `Jan, Feb, Mar...`) or type your own (e.g., `High, Medium, Low`)
4. Click **Add** → **OK** → **OK**

**Example Custom List for Regions:**
```
North, South, East, West
```

### 1.4 Sort by Color

If cells have font colors or fill colors (e.g., conditional formatting, manual highlights):

1. Data → Sort
2. Under **Sort On**, choose `Cell Color`, `Font Color`, or `Cell Icon`
3. Under **Order**, pick the color and choose `On Top` or `On Bottom`
4. Add levels for additional colors

**Use Case:** Move all red-highlighted (flagged) rows to the top of the list.

### 1.5 Sort Left to Right (By Column)

Normally Excel sorts by rows, but you can sort columns:

1. Data → Sort → **Options...**
2. Select **Sort left to right**
3. Now sort by Row number instead of Column name

### Sort Pitfalls

| Problem | Cause | Solution |
|---------|-------|----------|
| Headers sort into the data | "My data has headers" checkbox unchecked | Re-sort and check the box |
| Formulas return wrong values after sort | Relative references shifted | Use absolute references `$A$1` or structured table references |
| Mixed numbers and text in one column | Numbers stored as text | Use `Data → Text to Columns → Finish` to convert |

---

## 2. Filtering Data

Filtering **hides** rows that don't match your criteria - the data stays intact, but you only see what's relevant.

### 2.1 Enable AutoFilter

1. Click any cell inside your data range
2. Go to **Data** tab → click **Filter** (or press `Ctrl + Shift + L`)
3. Dropdown arrows (▼) appear in each header cell

### 2.2 Filter by Selection (Quick Filter)

**Steps:**
1. Right-click a cell with the value you want to filter by (e.g., a cell containing "North")
2. Choose **Filter** → **Filter by Selected Cell's Value**
3. Only rows where Region = "North" are shown; row numbers turn blue to indicate hidden rows

### 2.3 Text Filters

Click the dropdown arrow on the `Region` column:

1. Uncheck **(Select All)** to clear all
2. Check only **North** and **South**
3. Click **OK**

**Advanced Text Filters** (click the dropdown → Text Filters):
- **Equals...** - exact match (e.g., `East`)
- **Does Not Equal...** - exclude a value
- **Begins With...** - e.g., `W` → matches "West"
- **Ends With...** - e.g., `th` → matches "North", "South"
- **Contains...** - e.g., `ou` → matches "South"
- **Does Not Contain...** - exclude rows with a substring

### 2.4 Number Filters

Click the dropdown arrow on the `Total` column:

- **Equals / Does Not Equal**
- **Greater Than...** - e.g., `500` → shows only totals above $500
- **Less Than...**
- **Between...** - e.g., `300` to `700`
- **Top 10...** - shows the top/bottom N items or top/bottom N%
- **Above Average / Below Average**
- **Custom Filter...** - combine two conditions with AND/OR

**Example - Top 5 Sales:**
1. Click dropdown on `Total`
2. Number Filters → Top 10...
3. Set: `Top` `5` `Items`
4. Click **OK**

### 2.5 Date Filters

Click the dropdown on the `Date` column:

- Date filters offer a **calendar picker** for selecting dates
- **Dynamic filters:** Today, Yesterday, Tomorrow, This Week, Last Month, Next Quarter, This Year, Year to Date...
- **Between...** - specify a date range (e.g., 2025-03-01 to 2025-06-30)
- **Custom Filter...** - combine conditions (e.g., After 2025-03-01 AND Before 2025-09-01)

### 2.6 Search Filter

When you click any column dropdown, there's a **search box** at the top:

1. Click the dropdown on `SalesRep`
2. Type `Alice` in the search box
3. Only "Alice" is shown in the checkbox list - check it → **OK**

### 2.7 Clear Filters

- **Clear one column:** Click the column dropdown → **Clear Filter from [ColumnName]**
- **Clear all filters:** Data tab → **Clear** (the funnel icon with an X)
- **Remove filter arrows entirely:** Data tab → **Filter** (toggle off)

### 2.8 Filter with Wildcards

In the search box or custom filter, use wildcards:
- `*` = any number of characters: `W*` matches "Widget A", "Widget B", "Widget C"
- `?` = exactly one character: `Gadget ?` matches "Gadget X", "Gadget Y"
- `~` = escape the wildcard: `~*` searches for a literal asterisk

---

## 3. Advanced Filtering

Advanced Filtering lets you use a **criteria range** - a separate area of the sheet where you define complex conditions with AND/OR logic.

### 3.1 Setting Up a Criteria Range

**Rules for criteria ranges:**
- The criteria range **must** have the same column headers as your data
- Place it on the same sheet (or a different sheet) away from your data
- **Same row** = AND logic (both conditions must be true)
- **Different row** = OR logic (either condition can be true)

**Example - Set up in cells K1:M4:**

| Region | Category    | Total     |
|--------|-------------|-----------|
| North  | Electronics | >400      |
| South  |             | >600      |

This means:
- Row 2 (AND): Region="North" AND Category="Electronics" AND Total>400
- Row 3 (OR): Region="South" AND Total>600

### 3.2 Apply Advanced Filter

1. Go to **Data** tab → **Advanced** (in the Sort & Filter group)
2. **Action:** Choose one:
   - **Filter the list, in-place** - hides non-matching rows (like AutoFilter)
   - **Copy to another location** - outputs matching rows to a different area
3. **List range:** Select your data including headers (e.g., `A1:I21`)
4. **Criteria range:** Select your criteria including headers (e.g., `K1:M4`)
5. If copying to another location: set **Copy to** (e.g., `K10`)
6. Check **Unique records only** to remove duplicates
7. Click **OK**

### 3.3 Criteria Range Examples

**AND conditions (same row):**

| Product | Units |
|---------|-------|
| Widget A | >30 |

→ Shows only Widget A orders with more than 30 units

**OR conditions (different rows):**

| Region |
|--------|
| North  |
| South  |

→ Shows all North OR South records

**Wildcard in criteria:**

| Product |
|---------|
| Widget* |

→ Shows all products starting with "Widget"

---

## 4. Data Validation

Data Validation controls what users can enter in a cell. It prevents bad data before it enters your spreadsheet.

### 4.1 Access Data Validation

1. Select the cell(s) where you want to restrict input
2. Go to **Data** tab → **Data Validation** (the checkmark icon with a dropdown)
3. The Data Validation dialog has three tabs: **Settings**, **Input Message**, **Error Alert**

### 4.2 Validation Types

#### Whole Number
Restricts to whole numbers only.

**Example - Units column must be 1–1000:**
1. Select the Units column (e.g., G2:G100)
2. Data → Data Validation → Settings tab
3. Allow: `Whole number`
4. Data: `between`
5. Minimum: `1`
6. Maximum: `1000`

#### Decimal
Allows decimal numbers.

**Example - Unit Price must be 0.01 to 9999.99:**
1. Select the UnitPrice column (H2:H100)
2. Allow: `Decimal`
3. Data: `between`
4. Minimum: `0.01`
5. Maximum: `9999.99`

#### List (Drop-down)
Creates a dropdown menu in the cell.

**Example - Region dropdown:**
1. Select the Region column (C2:C100)
2. Allow: `List`
3. Source: `North,South,East,West`
4. **Or** Source: `=$K$1:$K$4` (a range containing the options)
5. Check **In-cell dropdown** (should be on by default)
6. Click **OK**

**Drop-down from a Named Range:**
1. First, create a named range: select cells K1:K4 → type `RegionList` in the Name Box → Enter
2. In Data Validation, set Source: `=RegionList`

> **Tip:** Use an Excel Table (`Ctrl+T`) for dynamic dropdown lists. When you add items to the table, the dropdown automatically expands.

#### Date
Restricts to valid dates within a range.

**Example - Date must be in 2025:**
1. Allow: `Date`
2. Data: `between`
3. Start date: `2025-01-01`
4. End date: `2025-12-31`

#### Time
Restricts to times.

**Example - Business hours only:**
1. Allow: `Time`
2. Data: `between`
3. Start: `9:00:00 AM`
4. End: `5:00:00 PM`

#### Text Length
Restricts by character count.

**Example - Product code must be exactly 5 characters:**
1. Allow: `Text length`
2. Data: `equal to`
3. Length: `5`

### 4.3 Custom Formula Validation

Use a formula for the most flexible validation.

**Example - Total column must equal Units × Unit Price:**
1. Select the Total column (I2:I100)
2. Allow: `Custom`
3. Formula: `=I2=G2*H2`

**Example - No duplicate Order IDs:**
1. Select the OrderID column (A2:A100)
2. Allow: `Custom`
3. Formula: `=COUNTIF($A:$A,A2)<=1`

**Example - Date must be a weekday (Mon–Fri):**
1. Allow: `Custom`
2. Formula: `=WEEKDAY(B2,2)<=5`

### 4.4 Input Message

This is a tooltip that appears when the user selects the cell.

1. Go to the **Input Message** tab
2. Check **Show input message when cell is selected**
3. Title: `Region`
4. Message: `Select a region from the dropdown: North, South, East, or West`

### 4.5 Error Alert

Controls what happens when invalid data is entered.

1. Go to the **Error Alert** tab
2. Check **Show error alert after invalid data is entered**
3. Style: Choose one:
   - **Stop** 🛑 - Prevents the entry entirely (default)
   - **Warning** ⚠️ - Warns but allows the user to proceed
   - **Information** ℹ️ - Notifies but accepts the entry
4. Title: `Invalid Entry`
5. Message: `Please enter a value between 1 and 1000.`

### 4.6 Circle Invalid Data

After setting up validation, you can highlight cells that already contain invalid data:

1. Go to **Data** tab → **Data Validation** (dropdown) → **Circle Invalid Data**
2. Red circles appear around cells that violate the validation rules
3. To clear: **Data Validation** dropdown → **Clear Validation Circles**

---

## 5. What-If Analysis

What-If Analysis tools let you experiment with different input values to see how they affect your results.

### 5.1 Goal Seek

Goal Seek works **backwards** from a desired result. It asks: "I know what answer I want - what input value gives me that answer?"

**Scenario:** You have a formula `Total = Units × 10`. You want Total to equal $750. How many units do you need?

**Setup:**
- Cell G2 = Units (e.g., currently 25)
- Cell H2 = Unit Price = `10`
- Cell I2 = formula `=G2*H2` (currently 250)

**Steps:**
1. Click on the cell with the formula you want to change (I2)
2. Go to **Data** tab → **What-If Analysis** → **Goal Seek**
3. **Set cell:** `I2` (the formula cell)
4. **To value:** `750` (your desired result)
5. **By changing cell:** `G2` (the input cell)
6. Click **OK**

Excel iterates and finds: G2 = 75 (i.e., you need 75 units). Click **OK** to keep the result or **Cancel** to revert.

### 5.2 Scenario Manager

Scenarios let you save and compare multiple sets of input values for the same model.

**Example - Best Case, Worst Case, Expected Case for a profit model:**

Suppose you have:
- B2: Units Sold (input)
- B3: Price per Unit (input)
- B4: Cost per Unit (input)
- B5: `=B2*(B3-B4)` - Profit formula

**Create Scenarios:**
1. Go to **Data** tab → **What-If Analysis** → **Scenario Manager**
2. Click **Add...**
3. Scenario name: `Worst Case`
4. Changing cells: `B2,B3,B4` (hold Ctrl to select multiple)
5. Click **OK** → enter values: `100`, `15`, `12` → click **Add**
6. Scenario name: `Best Case` → values: `500`, `25`, `8` → click **Add**
7. Scenario name: `Expected` → values: `300`, `20`, `10` → click **OK**

**View a Scenario:** Select it from the list → click **Show**. The sheet updates to display those input values and recalculates formulas.

**Create a Summary Report:**
1. In Scenario Manager, click **Summary...**
2. Choose **Scenario summary** (table format) or **Scenario PivotTable report**
3. **Result cells:** Select B5 (your profit formula)
4. Click **OK**

Excel creates a new sheet with a comparison table showing all scenarios side by side.

### 5.3 Data Tables

Data Tables show how changing **one or two input variables** affects a formula result, creating a matrix of outcomes.

#### One-Variable Data Table

**Scenario:** How does profit change at different unit quantities?

**Setup:**
- C1: `Profit Model` (label)
- C2: `=B2*(B3-B4)` (link to your profit formula, or duplicate the formula here)
- Column D (input column): List different unit values: 100, 200, 300, 400, 500
- Row E1: leave blank or label

| D        | E          |
|----------|------------|
|          | `=C2`      |
| 100      |            |
| 200      |            |
| 300      |            |
| 400      |            |
| 500      |            |

**Steps:**
1. Select the entire table range (D1:E6)
2. **Data** tab → **What-If Analysis** → **Data Table...**
3. **Column input cell:** `B2` (the cell that the column values replace)
4. Click **OK**

Excel fills in the profit for each unit quantity.

#### Two-Variable Data Table

**Scenario:** How does profit change with BOTH different unit quantities AND different prices?

**Setup (E1:H6):**

|          | 15       | 20       | 25       |
|----------|----------|----------|----------|
|          | `=C2`    |          |          |
| 100      |          |          |          |
| 200      |          |          |          |
| 300      |          |          |          |
| 400      |          |          |          |
- Top-left corner (E2) contains the formula: `=C2`
- Row 2 (F2:H2) contains different price values
- Column E (E3:E6) contains different unit quantities

**Steps:**
1. Select the entire table (E2:H6)
2. Data → What-If Analysis → Data Table...
3. **Row input cell:** `B3` (price)
4. **Column input cell:** `B2` (units)
5. Click **OK**

Excel fills in the profit matrix.

### 5.4 Solver Add-in

Solver finds the **optimal** value for a formula cell subject to constraints. It's like Goal Seek on steroids.

**Enable Solver:**
1. File → Options → Add-ins
2. At the bottom, Manage: `Excel Add-ins` → **Go...**
3. Check **Solver Add-in** → **OK**
4. Now find it under **Data** tab → **Solver**

**Example - Maximize Profit with Constraints:**

Setup:
- B2: Units of Product A (variable)
- B3: Units of Product B (variable)
- B4: Profit formula = `=B2*10 + B3*15`
- Constraint: Total units ≤ 500 → `B2 + B3 <= 500`
- Constraint: Product A units ≥ 50 → `B2 >= 50`

**Steps:**
1. Data → Solver
2. **Set Objective:** `B4`
3. **To:** `Max` (maximize profit)
4. **By Changing Variable Cells:** `B2,B3`
5. **Subject to Constraints:** Click **Add**
   - `B2+B3 <= 500`
   - `B2 >= 50`
   - `B3 >= 0`
6. Solving Method: `GRG Nonlinear` (or `LP Simplex` for linear problems)
7. Click **Solve**

Solver finds the optimal combination (e.g., B2=50, B3=450) that maximizes profit.

---

## 6. PivotTables (Detailed)

PivotTables are the **most powerful** data analysis tool in Excel. They summarize, group, and analyze large datasets without writing formulas.

### 6.1 Creating a PivotTable

**Steps:**
1. Click any cell in your data range (e.g., the SalesData table)
2. Go to **Insert** tab → **PivotTable**
3. In the dialog:
   - **Table/Range:** Excel auto-detects your data range (e.g., `SalesData!$A$1:$I$21`)
   - **Choose where to place the PivotTable:**
     - `New Worksheet` (recommended) - creates a new sheet
     - `Existing Worksheet` - place it on the current sheet at a specific cell
4. Click **OK**

A blank PivotTable appears on a new sheet with the **PivotTable Fields** pane on the right.

### 6.2 Understanding the Field Areas

The PivotTable Fields pane has four areas:

| Area | Purpose | Example |
|------|---------|---------|
| **Filters** | Creates a report-level filter (dropdown at top of PivotTable) | `Region` → filter to show only North |
| **Columns** | Creates column headers across the top | `Category` → Electronics, Hardware, Software as columns |
| **Rows** | Creates row labels down the left side | `SalesRep` → each rep as a row |
| **Values** | The data to calculate (aggregated) | `Sum of Total` → sum of sales for each intersection |

**Example - Build your first PivotTable:**
1. Drag `SalesRep` to the **Rows** area
2. Drag `Total` to the **Values** area
3. You now see: each sales rep with their total sales

**Add a second dimension:**
4. Drag `Category` to the **Columns** area
5. Now you see: Sales reps × Product Categories with subtotals

**Add a filter:**
6. Drag `Region` to the **Filters** area
7. Use the dropdown at the top to filter by region

### 6.3 Value Field Settings

Click on the value in the Values area (e.g., "Sum of Total") → **Value Field Settings...**

#### Summarize By (Calculation Type)

| Function | What It Does | Example |
|----------|-------------|---------|
| **Sum** | Adds all values (default for numbers) | Total sales |
| **Count** | Counts cells with data | Number of orders |
| **Average** | Calculates mean | Average order value |
| **Max** | Largest value | Highest single sale |
| **Min** | Smallest value | Lowest single sale |
| **Product** | Multiplies all values | Compound growth |
| **StdDev / StdDevp** | Standard deviation | Sales volatility |
| **Var / Varp** | Variance | Data spread |

**Steps to change:**
1. Right-click any value in the PivotTable
2. **Summarize Values By** → choose `Average` (or any function)
3. Or: Value Field Settings → Summarize By tab → select function

### 6.4 Show Values As (Display Type)

This powerful feature shows values as **percentages, rankings, running totals**, and more - without extra formulas.

**Access:** Right-click a value → **Show Values As** → choose:

| Display Type | What It Shows | Example |
|-------------|---------------|---------|
| **% of Grand Total** | Each cell as % of the grand total | Alice's Widget A = 3.1% of all sales |
| **% of Column Total** | % within each column | Alice's % of Electronics column |
| **% of Row Total** | % within each row | Alice's % of her own total |
| **% of Parent Row Total** | % of parent row item | In a grouped row, % of the group |
| **% of Parent Column Total** | % of parent column item | % of parent column group |
| **Running Total In** | Cumulative sum across rows/columns | Year-to-date running total |
| **Rank Smallest to Largest** | Rank from lowest to highest | Rank reps by sales |
| **Rank Largest to Smallest** | Rank from highest to lowest | Top salesperson |
| **Difference From** | Difference vs. a base item | vs. previous month, vs. a specific region |
| **% Difference From** | % change vs. a base item | Month-over-month % change |
| **Index** | Relative importance of each value | Weighted comparison |

**Example - Show each rep's sales as % of Grand Total:**
1. Add `SalesRep` to Rows, `Total` to Values
2. Right-click any value cell → **Show Values As** → `% of Grand Total`
3. Alice now shows `13.2%` (her share of all sales)

**Example - Running Total by Date:**
1. Add `Date` (grouped by month) to Rows, `Total` to Values
2. Right-click → **Show Values As** → **Running Total In** → select the Date field
3. Each month now shows cumulative sales from January onward

### 6.5 Grouping

Grouping collapses detailed data into meaningful categories.

#### Group Dates

1. Place a date field in the Rows area
2. Right-click any date in the PivotTable
3. Select **Group...**
4. In the Grouping dialog, select the levels:
   - `Months`, `Quarters`, `Years` - can select multiple
5. Click **OK**

**Result:** Instead of 20 individual dates, you see:
```
2025
  Q1
    Jan          590.00
    Feb          1,275.00
    Mar          1,007.50
  Q2
    Apr          1,550.00
    ...
```

#### Group Numbers

If you have numeric data (e.g., order totals), you can create ranges:

1. Place `Total` in Rows
2. Right-click → **Group...**
3. Set: Starting at `0`, Ending at `1500`, By `250`
4. Click **OK**

**Result:** Groups like `0-250`, `250-500`, `500-750`, etc.

#### Ungroup

Right-click a grouped item → **Ungroup** (or select and press `Alt + Shift + Left Arrow`)

### 6.6 Calculated Fields

A **Calculated Field** creates a new field using a formula based on existing fields.

**Example - Create a "Commission" field (5% of Total):**
1. Click anywhere in the PivotTable
2. Go to **PivotTable Analyze** tab (or **Analyze** tab) → **Fields, Items & Sets** → **Calculated Field...**
3. Name: `Commission`
4. Formula: `=Total * 0.05`
5. Click **Add** → **OK**

A new "Commission" column appears in your PivotTable, calculated for each row/column intersection.

### 6.7 Calculated Items

A **Calculated Item** creates a new item within an existing field.

**Example - Create a "Coastal" item combining East and West:**
1. Click a cell in the `Region` field in the PivotTable
2. PivotTable Analyze → Fields, Items & Sets → **Calculated Item...**
3. Name: `Coastal`
4. Formula: `=East + West`
5. Click **Add** → **OK**

> **Note:** Calculated Items can cause unexpected results with subtotals. Use with caution.

### 6.8 PivotTable Design Tab

When you click inside a PivotTable, the **PivotTable Design** tab (or **Design** tab) appears:

#### Report Layout
- **Compact Form** (default): All row fields stacked in one column
- **Outline Form**: Each row field in its own column, subtotals at top
- **Tabular Form**: Each row field in its own column, subtotals at bottom

#### Subtotals
- **Do Not Show Subtotals**
- **Show All Subtotals at Top of Group**
- **Show All Subtotals at Bottom of Group**

#### Grand Totals
- **Off for Rows and Columns**
- **On for Rows and Columns** (default)
- **On for Rows Only**
- **On for Columns Only**

#### Report Filters
- Show in **Rows** (stacked vertically) or **Columns** (side by side)

#### Banded Rows / Columns
- Alternate shading for readability

#### PivotTable Styles
- Choose from 85+ pre-built styles in the Styles gallery
- Options: Banded rows, banded columns, first/last column headers

### 6.9 Refreshing PivotTables

PivotTables **do not automatically update** when source data changes.

**Manual Refresh:**
- Right-click the PivotTable → **Refresh**
- Or: **PivotTable Analyze** tab → **Refresh** (or press `Alt + F5`)
- Refresh all PivotTables in the workbook: `Ctrl + Alt + F5`

**Auto-Refresh on File Open:**
1. Right-click PivotTable → **PivotTable Options...**
2. Go to the **Data** tab
3. Check **Refresh data when opening the file**
4. Click **OK**

**Dynamic Source Data (Best Practice):**
Convert your source data to an **Excel Table** (`Ctrl+T`) before creating the PivotTable. When you add rows to the table, the PivotTable automatically includes them after refresh.

### 6.10 GETPIVOTDATA Function

When you reference a PivotTable cell in a formula, Excel uses `GETPIVOTDATA`:

```
=GETPIVOTDATA("Total",$A$3,"SalesRep","Alice","Category","Electronics")
```

This extracts Alice's Electronics total from the PivotTable.

**Arguments:**
- `"Total"` - the data field name
- `$A$3` - reference to any cell in the PivotTable
- `"SalesRep","Alice"` - field/item pair (filter)
- `"Category","Electronics"` - another field/item pair

**Toggle GETPIVOTDATA:**
If you prefer regular cell references instead:
- **PivotTable Analyze** → **Options** dropdown → uncheck **Generate GetPivotData**

### 6.11 PivotCharts

A PivotChart is a chart linked to a PivotTable - it updates automatically when the PivotTable changes.

**Create a PivotChart:**
1. Click inside your PivotTable
2. **PivotTable Analyze** tab → **PivotChart**
3. Choose a chart type (Column, Bar, Line, Pie, etc.)
4. Click **OK**

**PivotChart Features:**
- **Field buttons** on the chart let you filter directly on the chart
- Right-click field buttons → **Hide All Field Buttons** for a cleaner look
- The chart filters when you filter the PivotTable (and vice versa)

**Best Chart Types for PivotChart:**
| Data Type | Recommended Chart |
|-----------|-------------------|
| Sales by region | Clustered Column |
| Trend over time | Line |
| Category breakdown | Pie or Donut |
| Comparison + trend | Combo (Column + Line) |
| Ranking | Bar (horizontal) |

---

## 7. Subtotals

Subtotals automatically insert summary rows (SUM, AVERAGE, COUNT, etc.) at group boundaries in a sorted list.

### 7.1 Prerequisites

**⚠️ Important:** You MUST sort your data by the grouping column BEFORE applying subtotals. Excel inserts subtotal rows wherever the grouping column value changes.

### 7.2 Automatic Subtotals

**Example - Total sales by Region:**

1. Sort your data by `Region` (Data → Sort A to Z on Region column)
2. Click any cell in the data
3. Go to **Data** tab → **Subtotal** (in the Outline group)
4. In the Subtotal dialog:
   - **At each change in:** `Region`
   - **Use function:** `Sum`
   - **Add subtotal to:** check `Total` (and `Units` if desired)
   - Check **Summary below data** (default)
   - Check **Replace current subtotals** (if re-applying)
   - Check **Page break between groups** (optional - for printing)
5. Click **OK**

**Result:**
```
Alice    North    Widget A    25    10.00    250.00
Alice    North    Widget A    30    10.00    300.00
Bob      North    Gadget Y    45    15.00    675.00
Charlie  North    Gadget X    20    25.00    500.00
Diana    North    Widget C    35    12.00    420.00
                                        North Total   2145.00
...
```

### 7.3 Other Subtotal Functions

Besides `Sum`, you can use:
- **Count** - number of items per group
- **Average** - average per group
- **Max** - maximum per group
- **Min** - minimum per group
- **Product** - product of all values
- **StdDev / StdDevp** - standard deviation
- **Var / Varp** - variance

### 7.4 Nesting Subtotals (Multiple Levels)

You can add subtotals at multiple levels. For example, subtotal by Region, then by Category within each Region:

1. Sort by `Region` first, then by `Category` (multi-level sort)
2. Apply subtotals: At each change in `Region`, Sum of `Total` → OK
3. Apply subtotals again: At each change in `Category`, Sum of `Total`
4. **⚠️ UNCHECK** "Replace current subtotals" - this adds the second level on top of the first

### 7.5 Group and Ungroup (Outline)

After applying subtotals, Excel creates an **outline** with level buttons in the top-left corner:

- **Level 1:** Grand total only (most collapsed)
- **Level 2:** Subtotals + grand total
- **Level 3:** All data + subtotals + grand total (most expanded)

**Click the level numbers** (1, 2, 3) at the top-left to collapse/expand.

**Manual Grouping (without Subtotals):**
1. Select the rows you want to group
2. **Data** tab → **Group** (or `Alt + Shift + Right Arrow`)
3. To ungroup: select grouped rows → **Ungroup** (`Alt + Shift + Left Arrow`)

### 7.6 Remove Subtotals

1. Data → **Subtotal**
2. Click **Remove All** at the bottom-left of the dialog
3. All subtotal rows and the outline are removed

---

## 8. Data Consolidation

Consolidation combines data from multiple ranges, worksheets, or workbooks into a single summary.

### 8.1 Consolidate by Category

Use when data is organized differently across sources but has **matching labels** (e.g., same product names in different regional reports).

**Example - Consolidate quarterly sales:**

Imagine three sheets: `Q1_Sales`, `Q2_Sales`, `Q3_Sales`, each with:

| Product     | Sales |
|-------------|-------|
| Widget A    | 500   |
| Widget B    | 350   |
| Gadget X    | 750   |

**Steps:**
1. Go to a new summary sheet
2. Click where you want the consolidated data to start (e.g., A1)
3. **Data** tab → **Consolidate** (in the Data Tools group)
4. In the Consolidate dialog:
   - **Function:** `Sum` (or Average, Count, etc.)
   - **Reference:** Click the range picker, go to `Q1_Sales!$A$1:$B$4`, click **Add**
   - Repeat for Q2 and Q3
   - **Use labels in:** check `Top row` and `Left column`
5. Click **OK**

**Result:** Excel creates a merged summary:
| Product     | Sales |
|-------------|-------|
| Widget A    | 1650  |
| Widget B    | 1100  |
| Gadget X    | 2400  |

### 8.2 Consolidate by Position

Use when all source ranges have **identical layout** (same rows, same columns, same order).

**Steps:**
1. Same as above, but **do NOT check** "Top row" or "Left column"
2. Excel simply sums (or averages, etc.) the values in the same cell positions across all ranges

### 8.3 Creating Links to Source Data

In the Consolidate dialog, check **Create links to source data**.

- The consolidated result will be grouped/linked to the source data
- When source data changes, the consolidation updates automatically
- Creates an outline similar to Subtotals that you can expand/collapse

> **Note:** When linking is enabled, you cannot edit the consolidated cells directly - they are formulas.

### 8.4 Consolidation Tips

- **Different workbooks:** You can consolidate data from files that are not open - browse to the file path
- **Named ranges:** Use named ranges for cleaner references
- **Multiple functions:** You can run consolidation multiple times with different functions (Sum, Average, Count) and place them side by side
- **3D References:** An alternative to Consolidate - use formulas like `=SUM(Q1_Sales:Q3_Sales!B2)` to sum the same cell across multiple sheets

---

## 9. Practice Exercises

### Exercise 1: Sorting & Filtering (Beginner)

Using the sample SalesData from the beginning of this module:

1. Sort the data by `Region` (A→Z), then by `Total` (largest to smallest)
2. Filter to show only `Electronics` category with more than 50 units
3. Clear filters, then use AutoFilter to find the **Top 3** sales by Total
4. Sort by color: Highlight all orders over $800 in yellow, then sort yellow cells to the top

### Exercise 2: Advanced Filtering (Intermediate)

Using the SalesData:

1. Set up a criteria range that shows:
   - East region AND Electronics category
   - OR West region AND total > $800
2. Apply Advanced Filter and copy results to cell K15
3. Use a wildcard filter to show all "Gadget" products (begins with "Gadget")

### Exercise 3: Data Validation (Beginner)

1. Add a dropdown for the `Region` column using a list: North, South, East, West
2. Add validation to `Units` column: must be a whole number between 1 and 1000
3. Add an input message: "Enter number of units (1-1000)"
4. Add an error alert (Stop style): "Units must be between 1 and 1000!"
5. Create a formula-based validation: Order Date must not be in the future (`<=TODAY()`)

### Exercise 4: What-If Analysis (Intermediate)

Create a simple pricing model:
- A1: "Units", A2: 100 (input)
- B1: "Price", B2: 25.00 (input)
- C1: "Cost/Unit", C2: 15.00 (input)
- D1: "Profit", D2: `=A2*(B2-C2)` (formula)

1. Use **Goal Seek**: What price do you need to achieve $2,000 profit with 100 units?
2. Create **Scenarios**: Best Case (200 units, $30 price, $10 cost), Worst Case (50 units, $20 price, $18 cost), Base Case (100 units, $25 price, $15 cost)
3. Create a **Scenario Summary** report
4. Create a **1-Variable Data Table** showing profit at 50, 100, 150, 200, 250, 300 units
5. Create a **2-Variable Data Table** with units (rows: 50–300) and price (columns: $20–$35)

### Exercise 5: PivotTables (Intermediate–Advanced)

Using the SalesData:

1. Create a PivotTable showing **Total Sales by SalesRep**
2. Add `Category` to Columns - now see sales per rep per category
3. Change the value display to **% of Column Total**
4. Group the dates by **Quarters**
5. Create a **Calculated Field** called "AvgPrice" = Total / Units
6. Add a **Region** filter and show only North and South
7. Change the layout to **Tabular Form** with subtotals at bottom
8. Create a **PivotChart** (Clustered Column) showing sales by rep by category
9. Refresh the PivotTable after adding two new rows to the source data

### Exercise 6: Subtotals & Consolidation (Intermediate)

**Part A - Subtotals:**
1. Sort SalesData by `Region`, then by `Category`
2. Apply **nested subtotals**: Sum of Total at each Region change, then Sum of Total at each Category change
3. Collapse to Level 2 to see only subtotals
4. Remove all subtotals

**Part B - Consolidation:**
1. Create three sheets: `Store_A`, `Store_B`, `Store_C` each with the same products but different sales numbers
2. Use Consolidate (Sum, by category) to create a combined total on a `Summary` sheet
3. Try consolidating with "Create links to source data" checked
4. Edit a value in Store_A and verify the Summary updates

---

## 10. Video References

### Sorting & Filtering
- Excel Sorting and Filtering - Leila Gharani
  https://www.youtube.com/watch?v=ylI8uHb_gmg
- Advanced Filter in Excel - Leila Gharani
  https://www.youtube.com/watch?v=HK1UOB2bPBI

### Data Validation
- Excel Data Validation Drop-down List - Leila Gharani
  https://www.youtube.com/watch?v=7moKWkCKbzE
- Dynamic Drop-down List with Data Validation - ExcelJet
  https://www.youtube.com/watch?v=5sD2CBLSMkU

### What-If Analysis
- Goal Seek, Scenario Manager, Data Tables - Leila Gharani
  https://www.youtube.com/watch?v=SfGkEGvq4JE
- Excel Solver Explained - Leila Gharani
  https://www.youtube.com/watch?v=kOO_PQ7VSzQ

### PivotTables
- PivotTable Tutorial - Leila Gharani
  https://www.youtube.com/watch?v=qu-Acm06sVY
- PivotTable Master Class - Chandoo
  https://www.youtube.com/watch?v=S2PFMBons0A
- PivotTable Grouping (Dates, Numbers) - Leila Gharani
  https://www.youtube.com/watch?v=bWpGPYctav4
- PivotTable Show Values As - Leila Gharani
  https://www.youtube.com/watch?v=Cz0VBjApM0M
- Calculated Fields in PivotTables - MyOnlineTrainingHub
  https://www.youtube.com/watch?v=nOkTIL8K3GM

### Subtotals & Consolidation
- Excel Subtotals - MyOnlineTrainingHub
  https://www.youtube.com/watch?v=3CJzmYUMgE0
- Consolidate Data from Multiple Sheets - Leila Gharani
  https://www.youtube.com/watch?v=HK1UOB2bPBI

---

## 11. Sources

### Microsoft Support
- Sort data in a range or table
  https://support.microsoft.com/en-us/excel
- Filter data in a range or table
  https://support.microsoft.com/en-us/excel
- Advanced Filter
  https://support.microsoft.com/en-us/excel
- Apply data validation to cells
  https://support.microsoft.com/en-us/office/apply-data-validation-to-cells-29fecbcc-d1b9-42c1-9d76-eff3ce5f7249
- Introduction to What-If Analysis
  https://support.microsoft.com/en-us/excel
- Goal Seek
  https://support.microsoft.com/en-us/excel
- Create a PivotTable
  https://support.microsoft.com/en-us/excel
- Group or Ungroup data in a PivotTable
  https://support.microsoft.com/en-us/excel
- Insert subtotals in a list of data
  https://support.microsoft.com/en-us/excel
- Consolidate data from multiple ranges
  https://support.microsoft.com/en-us/excel

### ExcelJet.net
- Excel PivotTable - Detailed Guide
  https://exceljet.net/
- Excel Sorting Guide
  https://exceljet.net/
- Excel Filtering Guide
  https://exceljet.net/
- Data Validation Guide
  https://exceljet.net/
- Goal Seek
  https://exceljet.net/

### Chandoo.org
- Pivot Tables - Comprehensive Guide
  https://chandoo.org/?s=excelpivot-tables/
- Sorting and Filtering Tips
  https://chandoo.org/?s=sorting+filtering
- Data Validation Tricks
  https://chandoo.org/?s=data+validation

### Contextures.com
- Excel Pivot Table Tutorial
  https://www.contextures.com/
- Excel Data Validation Examples
  https://www.contextures.com/xlDataVal01.html
- Excel Subtotals
  https://www.contextures.com/
- Advanced Filter Examples
  https://www.contextures.com/xladvfilter01.html

### Corporate Finance Institute (CFI)
- PivotTable Guide
  https://corporatefinanceinstitute.com/resources/excel/
- What-If Analysis
  https://corporatefinanceinstitute.com/resources/excel/what-if-analysis/
- Data Tables in Excel
  https://corporatefinanceinstitute.com/resources/excel/data-tables/

---

> **Next Module:** [Module 5 - Advanced Functions](../05-Advanced-Functions/05-advanced-functions.md) covers XLOOKUP, INDEX/MATCH, dynamic arrays, LAMBDA, and more.


## FILE: 05-Advanced-Functions/05-advanced-functions.md

# Module 5: Advanced Functions in Microsoft Excel

> **Target Audience:** Intermediate to Advanced Excel users  
> **Prerequisites:** Modules 1-4 (Basic formulas, cell references, named ranges, IF/SUMIFS/COUNTIFS)  
> **Excel Version:** Microsoft 365 / Excel 2021 (some functions require Excel 365)  
> **Last Updated:** September 2026

---

## Table of Contents

1. [XLOOKUP Function](#1-xlookup-function)
2. [XMATCH Function](#2-xmatch-function)
3. [INDEX/MATCH Combinations](#3-indexmatch-combinations)
4. [OFFSET Function and Dynamic Ranges](#4-offset-function-and-dynamic-ranges)
5. [INDIRECT Function and Dynamic References](#5-indirect-function-and-dynamic-references)
6. [CHOOSE Function](#6-choose-function)
7. [Dynamic Arrays (Excel 365/2021)](#7-dynamic-arrays)
8. [LET Function for Formula Readability](#8-let-function)
9. [LAMBDA Function (Custom Functions)](#9-lambda-function)
10. [Lambda Helper Functions](#10-lambda-helper-functions)
11. [New Text Functions](#11-new-text-functions)
12. [VSTACK and HSTACK](#12-vstack-and-hstack)
13. [TOCOL and TOROW](#13-tocol-and-torow)
14. [WRAPCOLS and WRAPROWS](#14-wrapcols-and-wraprows)
15. [Practice Exercises](#15-practice-exercises)
16. [Sources and Further Reading](#16-sources-and-further-reading)

---

## 1. XLOOKUP Function

### What is XLOOKUP?

XLOOKUP is Excel's modern, all-in-one lookup function introduced in Microsoft 365. It searches a range or array for a value and returns the corresponding value from another range. It replaces VLOOKUP, HLOOKUP, and the legacy LOOKUP function with a cleaner, more powerful syntax.

### Syntax

```
=XLOOKUP(lookup_value, lookup_array, return_array, [if_not_found], [match_mode], [search_mode])
```

| Argument | Required? | Description |
|----------|-----------|-------------|
| `lookup_value` | ✅ Yes | The value to search for |
| `lookup_array` | ✅ Yes | The range or array to search within |
| `return_array` | ✅ Yes | The range or array to return values from |
| `if_not_found` | ❌ Optional | Value to return if no match is found (default: #N/A error) |
| `match_mode` | ❌ Optional | 0 = exact (default), -1 = exact or next smaller, 1 = exact or next larger, 2 = wildcard, 3 = regex |
| `search_mode` | ❌ Optional | 1 = first-to-last (default), -1 = last-to-first, 2 = binary ascending, -2 = binary descending |

### Example 1: Basic Exact Match

**Sample Data - Employee Table:**

| A (Emp ID) | B (Name) | C (Department) | D (Salary) |
|------------|----------|----------------|------------|
| E001 | Alice Johnson | Engineering | $95,000 |
| E002 | Bob Smith | Marketing | $72,000 |
| E003 | Carol Davis | Engineering | $88,000 |
| E004 | David Lee | Sales | $67,000 |
| E005 | Eve Wilson | Marketing | $78,000 |

**Formula:** Look up the salary for employee "E003"

```excel
=XLOOKUP("E003", A2:A6, D2:D6)
```

**Result:** `$88,000`

### Example 2: Custom "Not Found" Message

```excel
=XLOOKUP("E099", A2:A6, D2:D6, "Employee Not Found")
```

**Result:** `Employee Not Found` (instead of #N/A)

### Example 3: Approximate Match - Discount Tiers

**Sample Data - Discount Table:**

| A (Min Qty) | B (Discount %) |
|-------------|----------------|
| 1 | 0% |
| 10 | 5% |
| 50 | 10% |
| 100 | 15% |
| 500 | 20% |

```excel
=XLOOKUP(75, A2:A6, B2:B6, , -1)
```

- `match_mode = -1` means "exact match or next smaller value"
- 75 falls between 50 and 100, so it returns the discount for 50 → **10%**

### Example 4: Return Multiple Values (Entire Row)

XLOOKUP can return an entire row of data at once:

```excel
=XLOOKUP("E003", A2:A6, B2:D6)
```

**Result:** Spills across 3 cells: `Carol Davis | Engineering | $88,000`

### Example 5: Nested XLOOKUP (Two-Way Lookup)

Look up a value by matching both a row header AND a column header:

**Sample Data - Sales by Region and Quarter:**

| | E (Q1) | F (Q2) | G (Q3) | H (Q4) |
|---|--------|--------|--------|--------|
| **North** | 12000 | 15000 | 13000 | 18000 |
| **South** | 9000 | 11000 | 10000 | 14000 |
| **East** | 14000 | 16000 | 15000 | 19000 |
| **West** | 8000 | 10000 | 9500 | 12000 |

```excel
=XLOOKUP("South", A2:A5, XLOOKUP("Q3", E1:H1, E2:H5))
```

**How it works:**
1. Inner XLOOKUP: Finds "Q3" in the header row → returns column G (the Q3 data for all regions)
2. Outer XLOOKUP: Finds "South" in that column → returns **10,000**

### Example 6: Reverse Search (Find Last Match)

```excel
=XLOOKUP("Engineering", C2:C6, D2:D6, , 0, -1)
```

- `search_mode = -1` searches from **last to first**
- Returns the salary of the **last** employee in Engineering

### Example 7: Wildcard Match

```excel
=XLOOKUP("*Wilson*", B2:B6, D2:D6, , 2)
```

- `match_mode = 2` enables wildcard matching with `*` and `?`
- Returns the salary of any name containing "Wilson"

### Advantages Over VLOOKUP

| Feature | VLOOKUP | XLOOKUP |
|---------|---------|---------|
| Look left | ❌ No | ✅ Yes |
| Return multiple columns | ❌ Single column | ✅ Entire rows |
| Default match | Approximate | Exact |
| Custom #N/A message | ❌ Requires IFERROR | ✅ Built-in |
| Reverse search | ❌ No | ✅ Yes |
| Column insert safe | ❌ Breaks on insert | ✅ References are ranges |

---

## 2. XMATCH Function

### What is XMATCH?

XMATCH is the modern replacement for the MATCH function. It returns the **relative position** of an item in an array. Unlike MATCH, it supports reverse search and binary search.

### Syntax

```
=XMATCH(lookup_value, lookup_array, [match_mode], [search_mode])
```

| Argument | Required? | Description |
|----------|-----------|-------------|
| `lookup_value` | ✅ Yes | The value to find |
| `lookup_array` | ✅ Yes | The range or array to search |
| `match_mode` | ❌ Optional | 0 = exact (default), -1 = next smaller, 1 = next larger, 2 = wildcard |
| `search_mode` | ❌ Optional | 1 = first-to-last (default), -1 = last-to-first, 2 = binary ascending, -2 = binary descending |

### Example 1: Find Position of a Value

```excel
=XMATCH("Carol Davis", B2:B6)
```

**Result:** `3` (Carol Davis is the 3rd name in the list)

### Example 2: Find Last Occurrence

```excel
=XMATCH("Engineering", C2:C6, 0, -1)
```

Returns the position of the **last** Engineering employee.

### Example 3: Approximate Match with XMATCH

```excel
=XMATCH(75, {1, 10, 50, 100, 500}, -1)
```

**Result:** `3` (75 falls closest to 50 from below, which is the 3rd item)

### Example 4: XMATCH + INDEX (Alternative to XLOOKUP)

```excel
=INDEX(D2:D6, XMATCH("E003", A2:A6))
```

This combination works like XLOOKUP but uses INDEX/MATCH syntax.

---

## 3. INDEX/MATCH Combinations

### Why INDEX/MATCH Still Matters

Even with XLOOKUP, INDEX/MATCH is valuable because:
- It works in **all Excel versions** (including 2010, 2013, 2016)
- It's the foundation many professionals learned
- It's more flexible for certain multi-criteria scenarios

### Basic Syntax

```
=INDEX(return_range, MATCH(lookup_value, lookup_range, match_type))
```

### Example 1: Single Criteria Lookup

Using the employee table from Section 1:

```excel
=INDEX(D2:D6, MATCH("E003", A2:A6, 0))
```

**Result:** `$88,000`

- `MATCH("E003", A2:A6, 0)` → finds "E003" is in position 3
- `INDEX(D2:D6, 3)` → returns the 3rd value in column D

### Example 2: Left Lookup (Something VLOOKUP Can't Do)

```excel
=INDEX(A2:A6, MATCH("Carol Davis", B2:B6, 0))
```

**Result:** `E003` (looks LEFT from the Name column to return the Emp ID)

### Example 3: Two-Way Lookup (Row + Column)

```excel
=INDEX(E2:H5, MATCH("South", A2:A5, 0), MATCH("Q3", E1:H1, 0))
```

- Row match: "South" is in row 3 of the data range
- Column match: "Q3" is in column 3 of the data range
- Returns the value at intersection → **10,000**

### Example 4: Multiple Criteria with INDEX/MATCH

To look up a value based on **two criteria** (e.g., Name = "Alice Johnson" AND Department = "Engineering"):

```excel
=INDEX(D2:D6, MATCH(1, (B2:B6="Alice Johnson") * (C2:C6="Engineering"), 0))
```

> ⚠️ **Important:** This is an **array formula**. In Excel 2019 and earlier, press **Ctrl+Shift+Enter** instead of just Enter. In Excel 365/2021, just press Enter.

**How it works:**
- `(B2:B6="Alice Johnson")` returns `{TRUE;FALSE;FALSE;FALSE;FALSE}` → `{1;0;0;0;0}`
- `(C2:C6="Engineering")` returns `{TRUE;FALSE;TRUE;FALSE;FALSE}` → `{1;0;1;0;0}`
- Multiplying: `{1;0;0;0;0}` × `{1;0;1;0;0}` = `{1;0;0;0;0}`
- MATCH finds the first `1` → position 1
- INDEX returns the 1st salary → **$95,000**

### Example 5: Return All Matches (with SMALL/IF - Legacy Approach)

```excel
=INDEX(D2:D6, SMALL(IF(C2:C6="Engineering", ROW(C2:C6)-ROW(C2)+1), ROW(1:1)))
```

> ⚠️ Array formula - Ctrl+Shift+Enter in older Excel versions.

This returns the 1st Engineering salary. Copy down to get the 2nd, 3rd, etc.

> 💡 **Modern Alternative:** Use `=FILTER(D2:D6, C2:C6="Engineering")` in Excel 365!

---

## 4. OFFSET Function and Dynamic Ranges

### What is OFFSET?

OFFSET returns a reference to a range that is a specified number of rows and columns from a starting cell or range. It's commonly used to create **dynamic named ranges** that automatically expand as data is added.

### Syntax

```
=OFFSET(reference, rows, cols, [height], [width])
```

| Argument | Description |
|----------|-------------|
| `reference` | The starting point (cell or range) |
| `rows` | Number of rows to move from the starting point (can be negative) |
| `cols` | Number of columns to move from the starting point (can be negative) |
| `height` | [Optional] Height of the returned range in rows |
| `width` | [Optional] Width of the returned range in columns |

### Example 1: Basic OFFSET

```excel
=OFFSET(A1, 2, 3)
```

Starting from A1, move **2 rows down** and **3 columns right** → returns the value in **D3**.

### Example 2: Return a Range

```excel
=SUM(OFFSET(A1, 0, 0, 5, 1))
```

Returns the sum of A1:A5 - a range that is 5 rows tall and 1 column wide starting from A1.

### Example 3: Dynamic Named Range

Create a named range that automatically grows as you add data:

1. Go to **Formulas → Name Manager → New**
2. Name: `SalesData`
3. Refers to: `=OFFSET(Sheet1!$A$1, 0, 0, COUNTA(Sheet1!$A:$A), 1)`

**How it works:**
- `COUNTA($A:$A)` counts non-empty cells in column A
- OFFSET starts at A1 and creates a range that many rows tall
- As you add data, the range automatically expands

### Example 4: Rolling Average (Last N Values)

```excel
=AVERAGE(OFFSET(A1, COUNTA(A:A)-7, 0, 7, 1))
```

Returns the average of the **last 7 values** in column A.

### ⚠️ OFFSET Limitations

- **Volatile function**: Recalculates every time the worksheet recalculates, even if its inputs haven't changed. This can slow down large workbooks.
- **Not compatible with structured table references.**
- **Modern alternative:** Use `INDEX` to create non-volatile dynamic ranges:
  ```
  =A1:INDEX(A:A, COUNTA(A:A))
  ```

---

## 5. INDIRECT Function and Dynamic References

### What is INDIRECT?

INDIRECT converts a **text string** into a valid cell reference. This allows you to build cell references dynamically - for example, pulling data from different sheets based on a dropdown selection.

### Syntax

```
=INDIRECT(ref_text, [a1])
```

| Argument | Description |
|----------|-------------|
| `ref_text` | A text string that represents a cell reference (e.g., "A1", "Sheet2!B5") |
| `a1` | [Optional] TRUE = A1-style reference (default), FALSE = R1C1-style |

### Example 1: Basic Dynamic Reference

If cell A1 contains the text `"B5"`:

```excel
=INDIRECT(A1)
```

**Result:** Returns the value in cell **B5**.

### Example 2: Dynamic Sheet Reference

Create a dropdown in cell A1 with sheet names (e.g., "January", "February", "March"). Then:

```excel
=INDIRECT("'" & A1 & "'!B10")
```

This pulls the value from cell B10 on whichever sheet is selected in A1.

**Explanation:**
- If A1 = "January", the formula builds: `='January'!B10`
- If A1 = "February", it becomes: `='February'!B10`

### Example 3: Dynamic Range with INDIRECT

```excel
=SUM(INDIRECT("D2:D" & A1))
```

If A1 = 100, this sums D2:D100. Change A1 to 200, and it sums D2:D200.

### Example 4: Dependent Dropdown Lists

Use INDIRECT for cascading/dependent dropdown validation:

1. Name your lists: `Fruits` = {"Apple","Banana","Cherry"}, `Vegetables` = {"Carrot","Broccoli","Spinach"}
2. Cell A1: Data Validation list → "Fruits", "Vegetables"
3. Cell B1: Data Validation list → `=INDIRECT(A1)`

Now B1's dropdown changes based on A1's selection!

### ⚠️ INDIRECT Limitations

- **Volatile function** (recalculates on every change)
- **Cannot reference closed workbooks**
- **Text-based references are fragile** - typos in sheet names cause #REF! errors

---

## 6. CHOOSE Function

### What is CHOOSE?

CHOOSE returns a value from a list of values based on a **position number**. It's useful for creating simple lookup tables or selecting from predefined options.

### Syntax

```
=CHOOSE(index_num, value1, [value2], [value3], ...)
```

### Example 1: Day of the Week Name

```excel
=CHOOSE(WEEKDAY(A1), "Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat")
```

If A1 is a date, this returns the day name.

### Example 2: Month Name from Number

```excel
=CHOOSE(MONTH(A1), "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")
```

### Example 3: Conditional Category Labels

```excel
=CHOOSE((A1>=90)*1 + (A1>=80)*2 + (A1>=70)*3 + (A1>=60)*4 + (A1<60)*5, "A", "B", "C", "D", "F")
```

> 💡 Note: For more complex scenarios, `IFS()` or nested `IF()` may be clearer.

### Example 4: CHOOSE with INDEX for Dynamic Column Selection

```excel
=SUM(CHOOSE({1,2,3}, B2:B10, D2:D10, F2:F10))
```

This sums columns B, D, and F together - useful when column selection needs to be dynamic.

---

## 7. Dynamic Arrays

### What Are Dynamic Arrays?

Dynamic Arrays are a **revolutionary feature** in Excel 365 and Excel 2021. Instead of returning a single value, a formula can now return an **array of values** that automatically "spills" into neighboring cells.

### The Spill Operator `#`

When a formula spills results into multiple cells, you can reference the entire spill range using the `#` operator:

```excel
=A2#    → References the entire spill range starting at A2
```

**Example:** If `=FILTER(A2:C10, B2:B10="Engineering")` spills into A2:C4, then `A2#` refers to that entire range.

### Common Spill Errors

| Error | Meaning |
|-------|---------|
| `#SPILL!` | Spill range is blocked by existing data in neighboring cells |
| `#CALC!` | No results match the criteria (e.g., FILTER returns nothing) |

---

### 7.1 FILTER Function

**Purpose:** Extract rows from a range that meet one or more conditions.

```
=FILTER(array, include, [if_empty])
```

#### Example 1: Single Condition

```excel
=FILTER(A2:D6, C2:C6="Engineering", "No results")
```

Returns all rows where Department = "Engineering".

**Result (spills into multiple cells):**

| A | B | C | D |
|---|---|---|---|
| E001 | Alice Johnson | Engineering | $95,000 |
| E003 | Carol Davis | Engineering | $88,000 |

#### Example 2: Multiple Conditions (AND)

Use `*` for AND logic:

```excel
=FILTER(A2:D6, (C2:C6="Engineering") * (D2:D6>90000))
```

Returns Engineering employees with salary > $90,000.

#### Example 3: Multiple Conditions (OR)

Use `+` for OR logic:

```excel
=FILTER(A2:D6, (C2:C6="Engineering") + (C2:C6="Sales"))
```

Returns employees in Engineering OR Sales.

#### Example 4: Filter and Sort

Combine FILTER with SORT:

```excel
=SORT(FILTER(A2:D6, C2:C6="Engineering"), 4, -1)
```

Filters for Engineering employees, then sorts by salary (column 4) in **descending** order (-1).

---

### 7.2 SORT and SORTBY Functions

#### SORT

```
=SORT(array, [sort_index], [sort_order], [by_col])
```

| Argument | Default | Description |
|----------|---------|-------------|
| `array` | Required | Range to sort |
| `sort_index` | 1 | Column (or row) number to sort by |
| `sort_order` | 1 (ascending) | 1 = ascending, -1 = descending |
| `by_col` | FALSE | FALSE = sort by rows, TRUE = sort by columns |

**Example - Sort employees by salary descending:**

```excel
=SORT(A2:D6, 4, -1)
```

**Result:**

| E001 | Alice Johnson | Engineering | $95,000 |
|------|--------------|-------------|---------|
| E003 | Carol Davis | Engineering | $88,000 |
| E005 | Eve Wilson | Marketing | $78,000 |
| E002 | Bob Smith | Marketing | $72,000 |
| E004 | David Lee | Sales | $67,000 |

#### SORTBY

Sort by a column that is **not in the output range**:

```
=SORTBY(array, by_array1, [sort_order1], [by_array2], [sort_order2], ...)
```

**Example - Sort names alphabetically but return salary column:**

```excel
=SORTBY(D2:D6, B2:B6, 1)
```

Returns salaries sorted by employee name (A→Z).

**Multi-level sort - Department ascending, then Salary descending:**

```excel
=SORTBY(A2:D6, C2:C6, 1, D2:D6, -1)
```

---

### 7.3 UNIQUE Function

```
=UNIQUE(array, [by_col], [exactly_once])
```

| Argument | Description |
|----------|-------------|
| `array` | Range from which to extract unique values |
| `by_col` | FALSE (default) = unique rows, TRUE = unique columns |
| `exactly_once` | FALSE (default) = all unique values, TRUE = values that appear exactly once |

**Example 1 - Unique departments:**

```excel
=UNIQUE(C2:C6)
```

**Result:** `Engineering | Marketing | Sales`

**Example 2 - Unique rows (full records):**

```excel
=UNIQUE(A2:D6)
```

**Example 3 - Combine with SORT:**

```excel
=SORT(UNIQUE(C2:C6))
```

Returns unique departments sorted alphabetically.

---

### 7.4 SEQUENCE Function

Generates an array of sequential numbers.

```
=SEQUENCE(rows, [columns], [start], [step])
```

**Example 1 - Numbers 1 to 10:**

```excel
=SEQUENCE(10)
```

**Result:** `{1; 2; 3; 4; 5; 6; 7; 8; 9; 10}` (spills down)

**Example 2 - 4×5 multiplication table:**

```excel
=SEQUENCE(4, 5, 1, 1)
```

Creates a 4-row, 5-column grid of numbers.

**Example 3 - Generate dates for 30 days:**

```excel
=TODAY() + SEQUENCE(30) - 1
```

**Example 4 - Odd numbers from 1 to 19:**

```excel
=SEQUENCE(10, 1, 1, 2)
```

**Result:** `{1; 3; 5; 7; 9; 11; 13; 15; 17; 19}`

---

### 7.5 RANDARRAY Function

Generates an array of random numbers.

```
=RANDARRAY([rows], [columns], [min], [max], [whole_number])
```

**Example 1 - 5 random decimals between 0 and 1:**

```excel
=RANDARRAY(5)
```

**Example 2 - 10 random integers between 1 and 100:**

```excel
=RANDARRAY(10, 1, 1, 100, TRUE)
```

**Example 3 - Random sample (shuffle a list):**

```excel
=SORTBY(A2:A100, RANDARRAY(COUNTA(A2:A100)))
```

Randomly shuffles the list in A2:A100.

---

## 8. LET Function

### What is LET?

The LET function lets you **define named variables inside a formula**. This makes complex formulas:
- **Easier to read** - meaningful variable names instead of repeated expressions
- **Faster to calculate** - expensive sub-calculations run only once
- **Easier to maintain** - change a value in one place instead of multiple

### Syntax

```
=LET(name1, value1, [name2, value2, ...], calculation)
```

- You can define up to **126 name/value pairs**
- The **last argument** must be the final calculation
- Variables are scoped to the formula only (can't reference them from other cells)

### Example 1: Simple LET

Without LET:
```excel
=IF(SUM(A1:A10) > 1000, SUM(A1:A10) * 0.1, SUM(A1:A10) * 0.05)
```
(SUM is calculated 3 times!)

With LET:
```excel
=LET(total, SUM(A1:A10), IF(total > 1000, total * 0.1, total * 0.05))
```
(SUM is calculated once!)

### Example 2: Complex Formula Made Readable

**Calculate the tax on income using progressive brackets:**

```excel
=LET(
  income, A1,
  bracket1, MIN(income, 10000) * 0.10,
  bracket2, MAX(MIN(income, 40000) - 10000, 0) * 0.20,
  bracket3, MAX(income - 40000, 0) * 0.30,
  bracket1 + bracket2 + bracket3
)
```

With line breaks (Alt+Enter in the formula bar), this reads almost like code!

### Example 3: Filtering and Sorting with LET

```excel
=LET(
  data, A2:D100,
  dept, "Engineering",
  filtered, FILTER(data, CHOOSECOLS(data, 3) = dept),
  SORT(filtered, 4, -1)
)
```

---

## 9. LAMBDA Function

### What is LAMBDA?

LAMBDA lets you create **custom reusable functions** without VBA or macros. You define a function with parameters and a formula, then use it like any built-in function.

### Syntax

```
=LAMBDA(parameter1, [parameter2, ...], calculation)
```

### Example 1: Simple LAMBDA

```excel
=LAMBDA(x, y, x + y)(3, 5)
```

**Result:** `8`

The `(3, 5)` at the end **invokes** the LAMBDA with those values.

### Example 2: Named LAMBDA (Reusable)

1. Go to **Formulas → Name Manager → New**
2. Name: `CelsiusToFahrenheit`
3. Refers to: `=LAMBDA(celsius, celsius * 9/5 + 32)`

Now use it anywhere in your workbook:

```excel
=CelsiusToFahrenheit(100)
```

**Result:** `212`

### Example 3: LAMBDA with Text Processing

Create a named LAMBDA called `Initials`:

```
=LAMBDA(name, TEXTJOIN(".", TRUE, LEFT(TEXTSPLIT(name, " "), 1) & "."))
```

Usage: `=Initials("John Robert Smith")` → `J.R.S.`

### Example 4: Percentage of Total (Named LAMBDA)

```
=LAMBDA(value, total, IF(total = 0, 0, value / total))
```

Usage: `=PctOfTotal(A1, SUM(A:A))`

---

## 10. Lambda Helper Functions

These functions are designed to work with LAMBDA for array processing.

### 10.1 MAP

Applies a LAMBDA to each element of an array and returns an array of results.

```
=MAP(array1, [array2, ...], LAMBDA(parameter, calculation))
```

**Example - Convert temperatures from Celsius to Fahrenheit:**

```excel
=MAP(A2:A10, LAMBDA(c, c * 9/5 + 32))
```

**Example - Concatenate first and last name:**

```excel
=MAP(A2:A10, B2:B10, LAMBDA(first, last, first & " " & last))
```

### 10.2 REDUCE

Reduces an array to a **single value** by applying a LAMBDA cumulatively.

```
=REDUCE(initial_value, array, LAMBDA(accumulator, value, calculation))
```

**Example - Sum of all values:**

```excel
=REDUCE(0, A2:A10, LAMBDA(acc, val, acc + val))
```

(This is equivalent to `=SUM(A2:A10)` but demonstrates the concept.)

**Example - Count values greater than 100:**

```excel
=REDUCE(0, A2:A10, LAMBDA(acc, val, acc + IF(val > 100, 1, 0)))
```

### 10.3 SCAN

Like REDUCE, but returns **every intermediate result** (running total).

```
=SCAN(initial_value, array, LAMBDA(accumulator, value, calculation))
```

**Example - Running total:**

```excel
=SCAN(0, A2:A10, LAMBDA(acc, val, acc + val))
```

**Result:** Spills a cumulative sum into each row.

### 10.4 BYROW

Applies a LAMBDA to each **row** of an array and returns one result per row.

```
=BYROW(array, LAMBDA(row, calculation))
```

**Example - Sum of each row:**

```excel
=BYROW(A2:D10, LAMBDA(row, SUM(row)))
```

**Example - Max value per row:**

```excel
=BYROW(A2:D10, LAMBDA(row, MAX(row)))
```

### 10.5 BYCOL

Applies a LAMBDA to each **column** of an array and returns one result per column.

```
=BYCOL(array, LAMBDA(col, calculation))
```

**Example - Average of each column:**

```excel
=BYCOL(A2:D10, LAMBDA(col, AVERAGE(col)))
```

---

## 11. New Text Functions

### 11.1 TEXTBEFORE

Extracts text **before** a specified delimiter.

```
=TEXTBEFORE(text, delimiter, [instance_num], [match_mode], [match_end], [if_not_found])
```

**Examples:**

```excel
=TEXTBEFORE("John-Smith-Jones", "-")       → "John"
=TEXTBEFORE("John-Smith-Jones", "-", 2)     → "John-Smith"
=TEXTBEFORE("John-Smith-Jones", "-", -1)    → "John-Smith" (from end)
```

### 11.2 TEXTAFTER

Extracts text **after** a specified delimiter.

```
=TEXTAFTER(text, delimiter, [instance_num], [match_mode], [match_end], [if_not_found])
```

**Examples:**

```excel
=TEXTAFTER("John-Smith-Jones", "-")       → "Smith-Jones"
=TEXTAFTER("John-Smith-Jones", "-", -1)   → "Jones" (after last dash)
=TEXTAFTER("report.pdf", ".")             → "pdf"
```

### 11.3 TEXTSPLIT

Splits text into rows and/or columns based on delimiters.

```
=TEXTSPLIT(text, col_delimiter, [row_delimiter], [ignore_empty], [match_mode], [pad_with])
```

**Example 1 - Split into columns:**

```excel
=TEXTSPLIT("Apple,Banana,Cherry", ",")
```

**Result:** `Apple | Banana | Cherry` (spills across 3 cells)

**Example 2 - Split into rows:**

```excel
=TEXTSPLIT("Apple;Banana;Cherry", , ";")
```

**Result:** Spills down - Apple, Banana, Cherry in separate rows.

**Example 3 - Split a CSV-like block:**

```excel
=TEXTSPLIT("A1,B1,C1;A2,B2,C2", ",", ";")
```

**Result:** A 2×3 table:

| A1 | B1 | C1 |
|----|----|----|
| A2 | B2 | C2 |

---

## 12. VSTACK and HSTACK

### VSTACK - Stack Arrays Vertically

```
=VSTACK(array1, [array2], ...)
```

**Example:** Combine two tables stacked on top of each other:

```excel
=VSTACK(A1:C3, E1:G4)
```

**Result:** A combined range with rows from both tables.

**Use case:** Consolidate data from multiple sheets/regions into one view.

### HSTACK - Stack Arrays Horizontally

```
=HSTACK(array1, [array2], ...)
```

**Example:** Combine names and scores side by side:

```excel
=HSTACK(A2:A10, D2:D10)
```

**Result:** A range with names in column 1 and scores in column 2.

---

## 13. TOCOL and TOROW

### TOCOL - Convert Array to Single Column

```
=TOCOL(array, [ignore], [scan_by_col])
```

| `ignore` | Behavior |
|----------|----------|
| 0 | Keep all values |
| 1 | Ignore blanks |
| 2 | Ignore errors |
| 3 | Ignore blanks and errors |

**Example:**

```excel
=TOCOL(A1:C3)
```

Converts a 3×3 range into a single column of 9 values.

### TOROW - Convert Array to Single Row

```
=TOROW(array, [ignore], [scan_by_col])
```

**Example:**

```excel
=TOROW(A1:C3)
```

Converts a 3×3 range into a single row of 9 values.

---

## 14. WRAPCOLS and WRAPROWS

### WRAPCOLS - Wrap Values into Columns

```
=WRAPCOLS(vector, wrap_count, [pad_with])
```

**Example:**

```excel
=WRAPCOLS(SEQUENCE(10), 3, "N/A")
```

Wraps numbers 1-10 into a table with 3 rows per column:

| 1 | 4 | 7 | 10 |
|---|---|---|----|
| 2 | 5 | 8 | N/A |
| 3 | 6 | 9 | N/A |

### WRAPROWS - Wrap Values into Rows

```
=WRAPROWS(vector, wrap_count, [pad_with])
```

**Example:**

```excel
=WRAPROWS(SEQUENCE(10), 4, "")
```

Wraps numbers 1-10 into a table with 4 columns per row:

| 1 | 2 | 3 | 4 |
|---|---|---|---|
| 5 | 6 | 7 | 8 |
| 9 | 10 | | |

---

## 15. Practice Exercises

### Exercise 1: XLOOKUP Mastery

**Setup:** Create a product catalog with columns: ProductID, ProductName, Category, Price, Stock

**Tasks:**
1. Use XLOOKUP to find the price of product "P1042"
2. Use XLOOKUP with a custom "Out of Stock" message when ProductID isn't found
3. Use nested XLOOKUP to find the stock level by matching both ProductName and Category
4. Use XLOOKUP with `search_mode = -1` to find the last product in the "Electronics" category

### Exercise 2: Dynamic Array Dashboard

**Setup:** Create a sales data table with: Date, Salesperson, Region, Product, Amount

**Tasks:**
1. Use FILTER to extract all sales for "North" region
2. Use SORT + FILTER to show top 5 sales by amount for the "South" region
3. Use UNIQUE to list all unique salespeople, then SORT alphabetically
4. Use SEQUENCE to generate a month number list and combine with TEXT to create month names
5. Create a formula that filters, sorts, and extracts unique values all at once using LET

### Exercise 3: INDEX/MATCH Multi-Criteria

**Setup:** Employee table with: EmpID, FirstName, LastName, Department, HireDate, Salary

**Tasks:**
1. Find salary using INDEX/MATCH with EmpID as the lookup value
2. Find EmpID using INDEX/MATCH (demonstrating left-lookup)
3. Use INDEX/MATCH with two criteria (FirstName AND Department)
4. Create a two-way lookup using INDEX with two MATCH functions

### Exercise 4: LET and LAMBDA

**Tasks:**
1. Create a LAMBDA that calculates Body Mass Index (BMI = weight / height²)
2. Name it `BMI` and use it: `=BMI(70, 1.75)`
3. Write a LET formula that calculates the average, max, and count of a range, then returns a formatted string
4. Create a LAMBDA called `Grades` that converts a numeric score to letter grades (A/B/C/D/F)

### Exercise 5: Text Processing

**Tasks:**
1. Use TEXTSPLIT to parse "Last, First | Department | Salary" into separate columns
2. Use TEXTBEFORE and TEXTAFTER to extract the domain from email addresses
3. Use VSTACK to combine two separate regional sales tables into one
4. Use TOCOL and WRAPROWS to reshape a vertical list into a 4-column table

### Exercise 6: Lambda Helpers

**Setup:** Sales table with amounts by product and month

**Tasks:**
1. Use BYROW to calculate the total for each product (row sum)
2. Use BYCOL to find the maximum sale in each month (column max)
3. Use SCAN to create a running total of monthly sales
4. Use MAP to apply a 10% markup to all prices
5. Use REDUCE to count how many sales exceed $10,000

---

## 16. Sources and Further Reading

### Microsoft Official Documentation
- [XLOOKUP Function - Microsoft Support](https://support.microsoft.com/office/xlookup-function-b7fd680e-6d10-43e6-84f9-88eae8bf5929)
- [XMATCH Function - Microsoft Support](https://support.microsoft.com/office/xmatch-function-d966da31-7a6b-4a13-a1c6-5a33ed6a0312)
- [FILTER Function - Microsoft Support](https://support.microsoft.com/en-us/excel)
- [SORT Function - Microsoft Support](https://support.microsoft.com/office/sort-function-22f63bd0-ccc8-492f-953d-c20e8e44b86c)
- [UNIQUE Function - Microsoft Support](https://support.microsoft.com/office/unique-function-c5ab87fd-30a3-4ce9-9d1a-40204fb85e1e)
- [SEQUENCE Function - Microsoft Support](https://support.microsoft.com/office/sequence-function-57467a98-57e0-4817-9f14-2eb78519ca90)
- [LET Function - Microsoft Support](https://support.microsoft.com/en-us/excel)
- [LAMBDA Function - Microsoft Support](https://support.microsoft.com/office/lambda-function-bd212d27-1cd1-4321-a34a-ccbf254b8b67)
- [TEXTBEFORE, TEXTAFTER, TEXTSPLIT - Microsoft Support](https://support.microsoft.com/en-us/excel)

### Exceljet (Recommended for Examples and Videos)
- [XLOOKUP Function - Exceljet](https://exceljet.net/functions/xlookup-function)
- [FILTER Function - Exceljet](https://exceljet.net/functions/filter-function)
- [LET Function - Exceljet](https://exceljet.net/functions/let-function)
- [Dynamic Array Functions - Exceljet](https://exceljet.net/)

### Community Resources
- [XLOOKUP vs INDEX/MATCH - Chandoo.org](https://chandoo.org/?s=xlookup)
- [Dynamic Arrays Guide - MyOnlineTrainingHub](https://www.myonlinetraininghub.com/excel-dynamic-array-functions)
- [LAMBDA Function Tutorial - Excel Off The Grid](https://exceloffthegrid.com/)

### YouTube Video References
- [XLOOKUP - ExcelIsFun (Mike Girvin)](https://www.youtube.com/watch?v=GPSY2lE7PSA)
- [Dynamic Arrays - Leila Gharani](https://www.youtube.com/watch?v=0EiOaCMq_FA)
- [FILTER Function Deep Dive - ExcelIsFun](https://www.youtube.com/watch?v=MQ3WqK5eJzI)
- [LET and LAMBDA - Leila Gharani](https://www.youtube.com/watch?v=JTXixVvfSZU)
- [INDEX/MATCH vs XLOOKUP - MyOnlineTrainingHub](https://www.youtube.com/watch?v=PN8nUASSEm4)
- [TEXTBEFORE, TEXTAFTER, TEXTSPLIT - Excel Off The Grid](https://www.youtube.com/watch?v=xl_8DLK9oXw)
- [Dynamic Arrays Crash Course - Chandoo](https://www.youtube.com/watch?v=ZP1HjV4M4Y4)
- [LAMBDA Functions Explained - Curbal](https://www.youtube.com/watch?v=miMjEFMNE0Y)

---

> **Next Module:** [Module 6 - Power Query & Power Pivot](../06-Power-Query-Pivot/06-power-query-power-pivot.md)


## FILE: 06-Power-Query-Pivot/06-power-query-power-pivot.md

# Module 6: Power Query & Power Pivot

> **Target Audience:** Intermediate to Advanced Excel users  
> **Prerequisites:** Modules 1-5 (Formulas, PivotTables, Advanced Functions)  
> **Excel Version:** Microsoft 365 / Excel 2019+ (Power Pivot requires Pro Plus or standalone license)  
> **Last Updated:** September 2026

---

## Table of Contents

### Part A: Power Query (Get & Transform)
1. [What is Power Query?](#1-what-is-power-query)
2. [Getting Data from Various Sources](#2-getting-data-from-various-sources)
3. [The Query Editor Interface](#3-the-query-editor-interface)
4. [Basic Transformations](#4-basic-transformations)
5. [Split, Merge, and Extract Columns](#5-split-merge-and-extract-columns)
6. [Pivot and Unpivot Columns](#6-pivot-and-unpivot-columns)
7. [Merge Queries (SQL-Style JOINs)](#7-merge-queries)
8. [Append Queries (Stacking Tables)](#8-append-queries)
9. [Custom Columns with M Language](#9-custom-columns-with-m-language)
10. [Parameters](#10-parameters)
11. [Refreshing and Updating Queries](#11-refreshing-and-updating-queries)

### Part B: Power Pivot
12. [What is Power Pivot?](#12-what-is-power-pivot)
13. [Data Model Concepts](#13-data-model-concepts)
14. [Relationships Between Tables](#14-relationships-between-tables)
15. [DAX Basics](#15-dax-basics)
16. [Calculated Columns vs Measures](#16-calculated-columns-vs-measures)
17. [Key Performance Indicators (KPIs)](#17-key-performance-indicators-kpis)
18. [Hierarchies](#18-hierarchies)

### Part C: Integration & Practice
19. [Connecting Power Query to Power Pivot](#19-connecting-power-query-to-power-pivot)
20. [Practice Exercises](#20-practice-exercises)
21. [Sources and Further Reading](#21-sources-and-further-reading)

---

# Part A: Power Query (Get & Transform)

---

## 1. What is Power Query?

### The ETL Process Explained

Power Query is Excel's built-in **ETL** (Extract, Transform, Load) tool. ETL is the standard process for preparing data for analysis:

| Stage | What It Does | Power Query Equivalent |
|-------|-------------|----------------------|
| **Extract** | Pull data from various sources | **Get Data** - connect to files, databases, web, APIs |
| **Transform** | Clean, reshape, and prepare data | **Query Editor** - remove columns, filter, split, merge, pivot |
| **Load** | Load the cleaned data into your workbook | **Close & Load** - output to worksheet table or Data Model |

### Why Power Query?

- **Repeatable:** Set up once, refresh with one click when source data changes
- **No formulas needed:** Transformations are recorded as steps, not cell formulas
- **Handles large data:** Can process millions of rows efficiently
- **Multiple sources:** Combine data from CSV, databases, web pages, folders, and more
- **Non-destructive:** Original data is never modified; transformations create a new view

### Where to Find Power Query

| Excel Version | Location |
|---------------|----------|
| Excel 365 / 2021 | **Data** tab → **Get & Transform Data** group |
| Excel 2016+ | **Data** tab → **Get & Transform Data** group |
| Excel 2010-2013 | Download the free **Power Query add-in** from Microsoft |

---

## 2. Getting Data from Various Sources

### 2.1 From CSV / Text Files

1. **Data** → **Get Data** → **From File** → **From Text/CSV**
2. Browse to your `.csv` or `.txt` file
3. Excel previews the data and auto-detects:
   - Delimiter (comma, tab, semicolon, etc.)
   - Data types
4. Click **Transform Data** (opens Query Editor) or **Load** (loads directly)

**Common scenarios:**
- Monthly sales exports from an ERP system
- Data dumps from reporting tools
- Log files

### 2.2 From Excel Workbooks

1. **Data** → **Get Data** → **From File** → **From Workbook**
2. Select the `.xlsx` file
3. Navigator shows all sheets and named ranges
4. Select which table(s) to import

### 2.3 From the Web

1. **Data** → **Get Data** → **From Other Sources** → **From Web**
2. Enter a URL
3. Power Query detects **HTML tables** on the page
4. Select the table you want from the Navigator

**Example:** Import a Wikipedia table of countries by GDP.

> ⚠️ **Note:** Web scraping depends on the page structure. Tables in standard HTML `<table>` tags work best.

### 2.4 From Databases

1. **Data** → **Get Data** → **From Database**
2. Choose your database type:
   - **SQL Server** - enter server name, optional database
   - **Access** - select `.accdb` file
   - **SQL Server Analysis Services** - for OLAP cubes
   - **Oracle, MySQL, PostgreSQL** - via connectors
3. Navigator shows available tables and views
4. Select and load

### 2.5 From a Folder (Combining Files)

This is one of Power Query's most powerful features:

1. **Data** → **Get Data** → **From File** → **From Folder**
2. Select a folder containing multiple files (e.g., 12 monthly CSV files)
3. Power Query shows all files in the folder
4. Click **Combine & Transform** to:
   - Automatically stack all files into one table
   - Apply consistent column mapping
   - Handle files with the same structure but different data

**Use case:** Combine 12 monthly sales CSV files into a single annual dataset.

---

## 3. The Query Editor Interface

When you click **Transform Data**, the **Power Query Editor** opens:

### Key Areas of the Interface

```
┌─────────────────────────────────────────────────────────┐
│  Ribbon (Home, Transform, Add Column, View)             │
├──────────┬──────────────────────┬───────────────────────┤
│          │                      │                       │
│ Queries  │   Preview Pane       │  Query Settings       │
│ (left    │   (data grid)        │  (Applied Steps)      │
│  panel)  │                      │                       │
│          │                      │                       │
├──────────┴──────────────────────┴───────────────────────┤
│  Formula Bar (M language expression for current step)   │
└─────────────────────────────────────────────────────────┘
```

| Area | Purpose |
|------|---------|
| **Queries Pane** (left) | Lists all queries in the workbook |
| **Preview Pane** (center) | Shows a preview of the data at the current step |
| **Query Settings** (right) | **Applied Steps** - lists every transformation step in order |
| **Formula Bar** | Shows the M language formula for the selected step |
| **Ribbon Tabs** | Home (common tasks), Transform (reshape), Add Column (new columns), View (display options) |

### The Applied Steps Panel

Every action you take in Power Query is recorded as a **step**. You can:
- **Click any step** to see the data state at that point
- **Delete a step** to undo that transformation
- **Reorder steps** by dragging (be careful - order matters!)
- **Rename steps** for clarity (right-click → Rename)

**Example Applied Steps:**
```
1. Source              → Connected to SalesData.csv
2. Promoted Headers    → First row became column headers
3. Changed Type        → Columns assigned data types
4. Filtered Rows       → Kept only rows where Region = "North"
5. Removed Columns     → Deleted unnecessary columns
6. Renamed Columns     → Renamed "amt" to "Amount"
```

---

## 4. Basic Transformations

### 4.1 Remove Columns

**Method 1:** Select columns → **Home** → **Remove Columns**
**Method 2:** Right-click column header → **Remove**
**Method 3:** Select columns to KEEP → **Home** → **Remove Other Columns**

> 💡 **Best Practice:** Use "Remove Other Columns" when you want to keep only specific columns - it's more resilient to source data changes.

### 4.2 Change Data Types

Click the **data type icon** (left of column header) to change:
- **ABC** → Text
- **123** → Whole Number
- **1.2** → Decimal Number
- **📅** → Date
- **🕐** → Time
- **TRUE/FALSE** → Boolean

Or: **Transform** → **Data Type** dropdown

**Auto-detection:** Power Query tries to detect types when you first load data, but always verify!

### 4.3 Filter Rows

Click the **filter dropdown** on any column header:

**For text columns:**
- Text Filters → Equals, Contains, Begins With, Ends With, etc.
- Select/deselect specific values from the list

**For number columns:**
- Number Filters → Greater Than, Less Than, Between, Top 10, etc.

**For date columns:**
- Date Filters → Before, After, Between, This Month, Last Year, etc.

**M Language example:**
```m
= Table.SelectRows(#"Previous Step", each [Region] = "North" and [Amount] > 1000)
```

### 4.4 Sort Data

- Click column header dropdown → **Sort Ascending** or **Sort Descending**
- **Transform** → **Sort** for multi-level sorting

### 4.5 Rename Columns

- **Double-click** the column header → type new name
- Right-click → **Rename**
- Select column → **Transform** → **Rename Column**

### 4.6 Replace Values

Right-click a column → **Replace Values**

**Example:** Replace "N/A" with null:
- Value to Find: `N/A`
- Replace With: *(leave empty)*

**Bulk replacements:** Use **Transform** → **Replace Values** for find-and-replace operations across the entire column.

---

## 5. Split, Merge, and Extract Columns

### 5.1 Split Columns

**Transform** → **Split Column**

| Option | Use Case |
|--------|----------|
| **By Delimiter** | Split "John-Smith" by "-" → "John" and "Smith" |
| **By Number of Characters** | Split at a fixed position (e.g., first 3 chars are a code) |
| **By Positions** | Split at specific character positions |
| **By Lowercase to Uppercase** | Split "excelPower" → "excel" and "Power" |
| **By Digit to Non-Digit** | Split "ABC123DEF" → "ABC", "123", "DEF" |

**Example - Split Full Name:**

| Before | After (Split by space) |
|--------|----------------------|
| John Smith | John | Smith |
| Mary Jane Watson | Mary | Jane | Watson |

**Delimiter options:** Comma, Semicolon, Space, Tab, or Custom

**Split into:** Rows (vertical) or Columns (horizontal)

### 5.2 Merge Columns

Select two or more columns → **Transform** → **Merge Columns**

- Choose a **separator** (space, comma, custom, or none)
- Enter a **new column name**

**Example:**

| FirstName | LastName | → Merged (space separator): |
|-----------|----------|---------------------------|
| John | Smith | John Smith |

### 5.3 Extract

**Transform** → **Extract** (or **Add Column** → **Extract**)

| Option | Description | Example |
|--------|-------------|---------|
| **Text Length** | Count characters | "Hello" → 5 |
| **First Characters** | Extract N chars from start | "ABCDEF" (3) → "ABC" |
| **Last Characters** | Extract N chars from end | "ABCDEF" (2) → "EF" |
| **Text Range** | Extract from position, N chars | "ABCDEF" (pos 2, len 3) → "BCD" |
| **Text Before Delimiter** | Like TEXTBEFORE function | "A-B-C" (before "-") → "A" |
| **Text After Delimiter** | Like TEXTAFTER function | "A-B-C" (after last "-") → "C" |
| **Text Between Delimiters** | Extract between two delimiters | "[A][B][C]" (between "[", "]") → "A", "B", "C" |

---

## 6. Pivot and Unpivot Columns

### 6.1 Unpivot Columns (Wide → Long)

**Use case:** Convert a cross-tab/table layout into a flat list.

**Before (pivoted):**

| Product | Q1 | Q2 | Q3 | Q4 |
|---------|-----|-----|-----|-----|
| Widget A | 100 | 120 | 110 | 130 |
| Widget B | 200 | 210 | 190 | 220 |

**Select Q1-Q4 columns** → **Transform** → **Unpivot Columns**

**After (unpivoted):**

| Product | Attribute | Value |
|---------|-----------|-------|
| Widget A | Q1 | 100 |
| Widget A | Q2 | 120 |
| Widget A | Q3 | 110 |
| Widget A | Q4 | 130 |
| Widget B | Q1 | 200 |
| Widget B | Q2 | 210 |
| Widget B | Q3 | 190 |
| Widget B | Q4 | 220 |

> 💡 Rename "Attribute" to "Quarter" and "Value" to "Sales" for clarity.

**Variants:**
- **Unpivot Columns** - unpivot only selected columns
- **Unpivot Other Columns** - unpivot everything except selected columns

### 6.2 Pivot Columns (Long → Wide)

**Use case:** Reverse of unpivot - convert a flat list into a cross-tab.

Select the column whose values should become headers → **Transform** → **Pivot Column**

**Settings:**
- **Values Column:** Which column provides the values
- **Aggregate Value Function:** Don't Aggregate, Count, Sum, Min, Max, Average, Median

**Example:** Pivot the unpivoted data back into quarterly columns.

---

## 7. Merge Queries

### What is a Merge?

A **Merge** combines two queries side-by-side (horizontally), similar to a **SQL JOIN**. You match rows from one table to another based on a common column (key).

### How to Merge

1. **Home** → **Merge Queries** (or **Merge Queries as New** for a new query)
2. Select the **first table** (top) and the **second table** (bottom)
3. Click the **matching column(s)** in each table
4. Choose the **Join Kind**
5. Click **OK**

### Join Types Explained

| Join Kind | SQL Equivalent | What It Returns |
|-----------|---------------|-----------------|
| **Left Outer** | LEFT JOIN | All rows from Table 1 + matching rows from Table 2 |
| **Right Outer** | RIGHT JOIN | All rows from Table 2 + matching rows from Table 1 |
| **Inner** | INNER JOIN | Only rows that match in BOTH tables |
| **Full Outer** | FULL OUTER JOIN | All rows from BOTH tables, matched where possible |
| **Left Anti** | NOT IN | Rows in Table 1 that have NO match in Table 2 |
| **Right Anti** | NOT IN | Rows in Table 2 that have NO match in Table 1 |

### Example: Merge Orders with Customer Details

**Table 1 - Orders:**

| OrderID | CustomerID | Amount |
|---------|------------|--------|
| 1001 | C01 | $500 |
| 1002 | C02 | $300 |
| 1003 | C05 | $450 |
| 1004 | C01 | $200 |

**Table 2 - Customers:**

| CustomerID | Name | City |
|------------|------|------|
| C01 | Alice | New York |
| C02 | Bob | Chicago |
| C03 | Carol | Boston |
| C04 | David | Denver |

**Left Outer Merge (Orders + Customer info):**

| OrderID | CustomerID | Amount | Name | City |
|---------|------------|--------|------|------|
| 1001 | C01 | $500 | Alice | New York |
| 1002 | C02 | $300 | Bob | Chicago |
| 1003 | C05 | $450 | null | null |
| 1004 | C01 | $200 | Alice | New York |

> Notice: Customer C05 has no match → null values. C03 and C04 (no orders) are excluded.

**Inner Merge:** Would exclude Order 1003 (no matching customer).

**Left Anti Merge:** Would show only Customer C05 from Orders (those without a customer record).

### Expanding Merged Columns

After merging, you get a **Table** column. Click the **expand icon** (↔) in the column header to:
- Select which columns to bring in
- Choose to keep the original column name as prefix

---

## 8. Append Queries

### What is an Append?

An **Append** stacks queries on top of each other (vertically), similar to a **SQL UNION**. It combines rows from multiple tables into one.

### How to Append

1. **Home** → **Append Queries**
2. Choose:
   - **Two tables** - append one table to another
   - **Three or more tables** - append multiple tables
3. Select the tables to append
4. Click **OK**

### Rules for Appending

| Scenario | Behavior |
|----------|----------|
| Same columns | Rows stack perfectly |
| Extra columns in one table | Columns are included; missing values become `null` |
| Different column names | Treated as separate columns (won't merge) |

> 💡 **Tip:** Rename columns to match before appending if the structure differs.

### Example: Combine Monthly Reports

**January:**

| Date | Product | Sales |
|------|---------|-------|
| 2024-01-05 | Widget A | $1,000 |
| 2024-01-12 | Widget B | $800 |

**February:**

| Date | Product | Sales |
|------|---------|-------|
| 2024-02-03 | Widget A | $1,200 |
| 2024-02-18 | Widget C | $600 |

**Appended Result:**

| Date | Product | Sales |
|------|---------|-------|
| 2024-01-05 | Widget A | $1,000 |
| 2024-01-12 | Widget B | $800 |
| 2024-02-03 | Widget A | $1,200 |
| 2024-02-18 | Widget C | $600 |

---

## 9. Custom Columns with M Language

### What is M Language?

Every Power Query transformation is written in **M language** (also called Power Query Formula Language). You can see it in the **Formula Bar** and the **Advanced Editor**.

### Adding a Custom Column

**Add Column** → **Custom Column**

**Example 1 - Simple Calculation:**

```m
= [Price] * [Quantity]
```

Column name: `Revenue`

**Example 2 - Conditional Logic:**

```m
if [Amount] > 1000 then "High"
else if [Amount] > 500 then "Medium"
else "Low"
```

Column name: `Priority`

**Example 3 - Text Manipulation:**

```m
Text.Upper(Text.Start([FirstName], 1)) & Text.Lower(Text.End([FirstName], Text.Length([FirstName]) - 1))
```

Capitalizes the first letter of a name.

### Common M Functions

| Category | Function | Example |
|----------|----------|---------|
| **Text** | `Text.Length` | `Text.Length([Name])` → character count |
| **Text** | `Text.StartsWith` | `Text.StartsWith([Code], "US")` → true/false |
| **Text** | `Text.Contains` | `Text.Contains([Email], "@gmail")` → true/false |
| **Text** | `Text.Replace` | `Text.Replace([Name], "Inc", "LLC")` |
| **Text** | `Text.Combine` | `Text.Combine({[First], [Last]}, " ")` |
| **Number** | `Number.Round` | `Number.Round([Amount], 2)` |
| **Number** | `Number.Abs` | `Number.Abs([Variance])` |
| **Date** | `Date.Year` | `Date.Year([OrderDate])` → 2024 |
| **Date** | `Date.Month` | `Date.Month([OrderDate])` → 6 |
| **Date** | `Date.DayOfWeekName` | `Date.DayOfWeekName([OrderDate])` → "Monday" |
| **Date** | `Date.AddDays` | `Date.AddDays([ShipDate], 7)` |
| **Logic** | `if ... then ... else` | Conditional expressions |
| **Logic** | `try ... otherwise` | Error handling |

### The Advanced Editor

**Home** → **Advanced Editor** shows the complete M code for a query:

```m
let
    Source = Csv.Document(File.Contents("C:\Sales.csv")),
    PromotedHeaders = Table.PromoteHeaders(Source),
    ChangedTypes = Table.TransformColumnTypes(PromotedHeaders, {
        {"Date", type date}, {"Amount", type number}
    }),
    FilteredRows = Table.SelectRows(ChangedTypes, each [Amount] > 0)
in
    FilteredRows
```

**Structure:**
- `let` block: Define each step as a variable
- `in` block: The final result (what gets loaded)

---

## 10. Parameters

### What are Parameters?

Parameters make your queries **flexible and reusable**. Instead of hardcoding values (like a file path, date range, or filter value), you use a parameter that can be changed from a single place.

### Creating a Parameter

1. **Home** → **Manage Parameters** → **New Parameter**
2. Set:
   - **Name:** e.g., `FilterYear`
   - **Type:** Text, Number, Date, Boolean, etc.
   - **Current Value:** e.g., `2024`
   - **Suggested Values:** List of values, query, or any value

### Using Parameters in Queries

**Example - Filter by Year:**

```m
= Table.SelectRows(Source, each Date.Year([OrderDate]) = FilterYear)
```

**Example - Dynamic File Path:**

```m
= Csv.Document(File.Contents("C:\Data\" & DataFile & ".csv"))
```

Where `DataFile` is a text parameter (e.g., "Q1_2024_Sales").

### Using Parameters with Dropdowns

1. Create a parameter with **Suggested Values** → **List of values**
2. The parameter shows as a dropdown in the Queries pane
3. Users can change the value without editing the query

---

## 11. Refreshing and Updating Queries

### Manual Refresh

- **Data** tab → **Refresh All** (or press **Ctrl+Alt+F5**)
- Right-click a specific query → **Refresh**

### Auto-Refresh on File Open

1. **Data** → **Queries & Connections**
2. Right-click a query → **Properties**
3. Check **"Refresh data when opening the file"**

### Scheduled Refresh (Power BI Service only)

For cloud-hosted workbooks via Power BI:
- Configure **Scheduled Refresh** in the Power BI service
- Set frequency (daily, hourly, etc.)
- Requires a **data gateway** for on-premises sources

### Refresh Considerations

| Issue | Solution |
|-------|----------|
| Source file moved | Update the file path in the Source step |
| Column names changed | Update the relevant transformation steps |
| New columns added | "Remove Other Columns" step may exclude them - update |
| Slow refresh | Consider **query folding** (pushes transformations to the source) |
| Credentials expired | **Data** → **Queries & Connections** → enter credentials |

### Data Load Settings

Right-click a query → **Properties**:
- **Connection Name:** Display name
- **Load to:** Table (worksheet), Connection Only (no output), or **Data Model** (for Power Pivot)
- **Include in Refresh:** Yes/No

---

# Part B: Power Pivot

---

## 12. What is Power Pivot?

### Power Pivot vs Regular PivotTable

| Feature | Regular PivotTable | Power Pivot |
|---------|-------------------|-------------|
| Data source | Single table/range | Multiple tables via Data Model |
| Relationships | Not supported | Yes - define table relationships |
| Row limit | ~1,048,576 rows (worksheet limit) | Millions of rows (in-memory engine) |
| Calculations | Standard formulas in cells | DAX (Data Analysis Expressions) |
| Calculated fields | Limited | Rich DAX measures and calculated columns |
| KPIs | Not available | Built-in KPI support |
| Hierarchies | Manual grouping | Named hierarchies (drill-down) |
| Memory | Worksheet memory | VertiPaq column-store engine (compressed) |

### Enabling Power Pivot

1. **File** → **Options** → **Add-ins**
2. Manage: **COM Add-ins** → **Go...**
3. Check **Microsoft Power Pivot for Excel**
4. Click **OK**

A new **Power Pivot** tab appears on the ribbon.

> ⚠️ **Note:** Power Pivot is available in Excel Professional Plus and standalone Excel licenses. It is NOT included in Excel Standard or Excel for Mac.

---

## 13. Data Model Concepts

### What is the Data Model?

The **Data Model** is an in-memory database inside your Excel workbook. It:
- Stores data in a **compressed columnar format** (VertiPaq engine)
- Supports **relationships** between tables
- Enables **DAX** calculations
- Can hold **millions of rows** without worksheet limits

### Adding Data to the Data Model

**Method 1: From Power Query**
- When loading a query, choose **"Only Create Connection"** and check **"Add this data to the Data Model"**

**Method 2: From a Table**
- Click any cell in an Excel table → **Power Pivot** tab → **Add to Data Model**

**Method 3: From External Data**
- **Data** → **Get Data** → connect to a source → load to **Data Model**

### Power Pivot Window

**Power Pivot** tab → **Manage** opens the Power Pivot window:

```
┌─────────────────────────────────────────────────┐
│  Ribbon (Home, Design, Advanced)                │
├──────────┬──────────────────────────────────────┤
│          │                                      │
│  Table   │   Data View (rows & columns)         │
│  Tabs    │   or                                 │
│ (bottom) │   Diagram View (relationship map)    │
│          │                                      │
├──────────┴──────────────────────────────────────┤
│  Calculation Area (measures for active table)   │
└─────────────────────────────────────────────────┘
```

---

## 14. Relationships Between Tables

### Why Relationships?

Instead of one massive flat table, you use a **star schema**:

```
         ┌──────────┐
         │ DimDate  │
         │ (DateKey)│
         └────┬─────┘
              │
┌──────────┐  │  ┌──────────────┐
│DimProduct├──┼──┤ FactSales    │
│(ProductID)│  │ │(DateKey,     │
└──────────┘  │  │ ProductID,   │
              │  │ CustomerID)  │
┌──────────┐  │  └──────────────┘
│DimCust.  ├──┘
│(CustID) │
└──────────┘
```

- **Fact table** (center): Contains measurable data - sales amounts, quantities
- **Dimension tables** (surrounding): Contain descriptive data - dates, products, customers

### Creating Relationships

**Method 1: Diagram View**
1. Open Power Pivot → **Home** → **Diagram View**
2. **Drag** a field from one table to the matching field in another table
3. The relationship line appears

**Method 2: Ribbon**
1. **Home** → **Create Relationship**
2. Select Table 1, Column, Table 2, Related Column

### Relationship Settings

| Setting | Description |
|---------|-------------|
| **Cardinality** | Many-to-One (most common), One-to-One, Many-to-Many |
| **Cross Filter Direction** | Single (filter flows one way) or Both (bidirectional) |
| **Active** | Only one active relationship between two tables (inactive ones used via DAX USERELATIONSHIP) |

### Example: Sales and Products

**FactSales table:**
| ProductID | DateKey | Amount |
|-----------|---------|--------|
| P01 | 20240101 | 500 |
| P02 | 20240101 | 300 |
| P01 | 20240102 | 450 |

**DimProduct table:**
| ProductID | ProductName | Category |
|-----------|-------------|----------|
| P01 | Widget A | Electronics |
| P02 | Widget B | Hardware |
| P03 | Widget C | Electronics |

**Relationship:** FactSales[ProductID] → DimProduct[ProductID] (Many-to-One)

Now you can create PivotTables that show Sales by Category - even though "Category" lives in DimProduct, not FactSales.

---

## 15. DAX Basics

### What is DAX?

**DAX (Data Analysis Expressions)** is the formula language for Power Pivot. It looks similar to Excel formulas but has key differences:

| Feature | Excel Formula | DAX |
|---------|--------------|-----|
| Operates on | Individual cells | Columns and tables |
| Context | Cell context | Row context + Filter context |
| Aggregation | SUM(A1:A10) | SUM(Table[Column]) |
| Time intelligence | Limited | Rich built-in functions |

### Core DAX Concepts

#### Filter Context
When you place fields on a PivotTable's rows/columns/filters, they create a **filter context** that determines which rows a measure sees.

#### Row Context
Exists when a formula is being evaluated **row by row** (calculated columns, iterators like SUMX).

### Essential DAX Functions

#### SUM - Simple Aggregation

```dax
Total Sales = SUM(FactSales[Amount])
```

#### SUMX - Iterator (Row-by-Row Calculation)

```dax
Revenue = SUMX(FactSales, FactSales[Price] * FactSales[Quantity])
```

SUMX iterates through each row, calculates `Price * Quantity`, then sums the results.

#### COUNTX - Count with Expression

```dax
Order Count = COUNTX(FactSales, FactSales[OrderID])
```

#### RELATED - Pull from Related Table

```dax
Category = RELATED(DimProduct[Category])
```

Used in **calculated columns** to bring in a value from a related table.

#### CALCULATE - The Most Important DAX Function

```dax
Electronics Sales = CALCULATE(
    SUM(FactSales[Amount]),
    DimProduct[Category] = "Electronics"
)
```

CALCULATE modifies the **filter context**:
1. Starts with the existing filter context
2. Applies additional filters (or overrides existing ones)
3. Evaluates the expression in the modified context

**Key uses:**
- **Override filters:** `CALCULATE(SUM(...), ALL(DimProduct))` - ignores product filters
- **Add filters:** `CALCULATE(SUM(...), DimDate[Year] = 2024)` - adds year filter
- **Time intelligence:** `CALCULATE(SUM(...), DATEADD(DimDate[Date], -1, YEAR))` - prior year

#### FILTER - Return Filtered Table

```dax
High Value Orders = CALCULATE(
    COUNTROWS(FactSales),
    FILTER(FactSales, FactSales[Amount] > 1000)
)
```

FILTER returns a table that is a subset of the original, used inside CALCULATE.

#### RELATEDTABLE - Count/Sum Related Rows

```dax
Product Sales Count = COUNTROWS(RELATEDTABLE(FactSales))
```

Counts the sales rows related to each product.

#### DIVIDE - Safe Division

```dax
Avg Order Value = DIVIDE(SUM(FactSales[Amount]), COUNTROWS(FactSales), 0)
```

DIVIDE handles division by zero gracefully (returns 0 instead of error).

### DAX Examples Table

| DAX Measure | Description |
|-------------|-------------|
| `Total = SUM(Sales[Amount])` | Total sales amount |
| `Avg = AVERAGE(Sales[Amount])` | Average order value |
| `Orders = COUNTROWS(Sales)` | Number of orders |
| `Distinct Products = DISTINCTCOUNT(Sales[ProductID])` | Unique products sold |
| `YoY Growth = DIVIDE([Total] - [Prior Year], [Prior Year])` | Year-over-year growth % |
| `Prior Year = CALCULATE([Total], DATEADD(DimDate[Date], -1, YEAR))` | Sales from prior year |
| `YTD Sales = TOTALYTD([Total], DimDate[Date])` | Year-to-date sales |
| `MTD Sales = TOTALMTD([Total], DimDate[Date])` | Month-to-date sales |

---

## 16. Calculated Columns vs Measures

### Calculated Columns

**Created in:** Power Pivot data view or as a column in the Data Model
**Calculated:** When data is loaded/refreshed (stored in the model)
**Context:** Row context (evaluates per row)

**Example - Profit calculated column:**

```dax
= [Revenue] - [Cost]
```

**When to use:**
- You need the value for **row-by-row** filtering, slicing, or sorting
- The result is used in **relationships** or **groupings**
- The calculation doesn't change based on the PivotTable layout

### Measures (Calculated Fields)

**Created in:** Power Pivot → Calculation Area, or PivotTable → New Measure
**Calculated:** At query time (when the PivotTable is rendered)
**Context:** Filter context (evaluates based on slicers, rows, columns)

**Example - Profit margin measure:**

```dax
Profit Margin = DIVIDE(
    SUM(FactSales[Revenue]) - SUM(FactSales[Cost]),
    SUM(FactSales[Revenue]),
    0
)
```

**When to use:**
- The result depends on **user selections** (slicers, filters)
- You need **aggregations** (sum, average, count, etc.)
- The calculation should **respond dynamically** to the PivotTable layout

### Comparison

| Aspect | Calculated Column | Measure |
|--------|------------------|---------|
| Stored in model | ✅ Yes (uses memory) | ❌ No (calculated on demand) |
| Row context | ✅ Yes | ❌ No |
| Filter context | ❌ No | ✅ Yes |
| Use in PivotTable rows/values | Rows, filters | Values area |
| Performance impact | Increases file size | No file size impact |
| Best for | Static per-row values | Dynamic aggregations |

> 💡 **Rule of Thumb:** Prefer **measures** unless you specifically need a calculated column.

---

## 17. Key Performance Indicators (KPIs)

### What are KPIs?

KPIs in Power Pivot compare an **actual value** (base measure) against a **target value** to show performance status (e.g., on track, needs attention, off track).

### Creating a KPI

1. Create a **base measure** (e.g., `Total Sales`)
2. Create a **target measure** (e.g., `Sales Target`)
3. In Power Pivot → **Home** → **KPIs** → **New KPI**
4. Configure:
   - **KPI base measure:** Total Sales
   - **Target:** Sales Target (or Absolute Value)
   - **Thresholds:** Define ranges for Red/Yellow/Green

### Example KPI Setup

**Base Measure:**
```dax
Total Sales = SUM(FactSales[Amount])
```

**Target Measure:**
```dax
Sales Target = SUM(FactSales[TargetAmount])
```

**KPI Thresholds:**
| Status | Threshold |
|--------|-----------|
| 🔴 Red (Bad) | Actual < 80% of Target |
| 🟡 Yellow (Caution) | 80% ≤ Actual < 100% of Target |
| 🟢 Green (Good) | Actual ≥ 100% of Target |

### Using KPIs in PivotTables

When you add a KPI to a PivotTable, you can display:
- **Value:** The actual measure result
- **Status:** Icon (traffic light, gauge, etc.)
- **Goal:** The target value

---

## 18. Hierarchies

### What are Hierarchies?

Hierarchies define **drill-down paths** in your data. For example:
- **Date Hierarchy:** Year → Quarter → Month → Day
- **Geography Hierarchy:** Country → State → City
- **Product Hierarchy:** Category → Subcategory → Product

### Creating a Hierarchy

**In Diagram View:**
1. Right-click a field → **Create Hierarchy**
2. Name it (e.g., "Date Hierarchy")
3. Drag related fields into the hierarchy in order:
   - Year (Level 1)
   - Quarter (Level 2)
   - Month (Level 3)
   - Day (Level 4)

### Using Hierarchies in PivotTables

Once created, the hierarchy appears as a single expandable field in the PivotTable field list. Users can:
- Click **drill-down arrows** (⊞) to go from Year → Quarter → Month
- Click **drill-up arrows** (⊟) to go back up
- Use **Expand to Entire Hierarchy** for multi-level analysis

### Auto Date Hierarchies

Excel can automatically create date hierarchies:
1. **Power Pivot** → **Diagram View**
2. Select a date column
3. **Design** → **Mark as Date Table** → specify the date column

Excel auto-generates: Year, Quarter, Month, Day hierarchies.

---

# Part C: Integration & Practice

---

## 19. Connecting Power Query to Power Pivot

### The Ideal Workflow

```
┌──────────────┐     ┌──────────────┐     ┌──────────────┐
│  Data        │     │  Power       │     │  Power       │
│  Sources     │────▶│  Query       │────▶│  Pivot       │
│  (CSV, DB,   │     │  (Clean &    │     │  (Data Model │
│   Web, etc.) │     │   Transform) │     │   + DAX)     │
└──────────────┘     └──────────────┘     └──────┬───────┘
                                                  │
                                                  ▼
                                          ┌──────────────┐
                                          │  PivotTables │
                                          │  & Charts    │
                                          │  (Dashboard) │
                                          └──────────────┘
```

### Step-by-Step Integration

**Step 1: Load data with Power Query**
1. **Data** → **Get Data** → connect to your sources
2. Transform and clean in the Query Editor
3. In **Close & Load** dropdown → **Close & Load To...**
4. Select **"Only Create Connection"**
5. Check **"Add this data to the Data Model"**
6. Click **Load**

**Step 2: Define relationships in Power Pivot**
1. **Power Pivot** tab → **Manage**
2. Switch to **Diagram View**
3. Drag fields to create relationships

**Step 3: Create DAX measures**
1. In the **Calculation Area** below the data
2. Write DAX measures for your analysis

**Step 4: Build PivotTables from the Data Model**
1. **Insert** → **PivotTable**
2. Select **"Use this workbook's Data Model"**
3. Build your PivotTable using fields from multiple tables

### Benefits of This Workflow

- **Single source of truth:** Power Query handles all data prep
- **Refreshable:** One click updates all data, transformations, and PivotTables
- **Scalable:** Handle millions of rows without worksheet limits
- **Maintainable:** Transformation steps are visible and editable
- **Shareable:** Data Model connections persist when sharing the workbook

---

## 20. Practice Exercises

### Exercise 1: Power Query Basics

**Scenario:** You receive a messy CSV file with employee data.

**Tasks:**
1. Import the CSV file using Power Query
2. Promote the first row to headers
3. Remove unnecessary columns (keep only: ID, Name, Department, Salary, HireDate)
4. Change data types (Salary = decimal, HireDate = date)
5. Filter to keep only active employees
6. Split the "Name" column into "FirstName" and "LastName"
7. Add a custom column for "Years of Service" (current year minus hire year)
8. Load the cleaned data to a worksheet

### Exercise 2: Merge and Append

**Scenario:** Combine data from multiple sources.

**Tasks:**
1. Import two tables: **Orders** (OrderID, CustomerID, Amount) and **Customers** (CustomerID, Name, City)
2. **Merge** Orders with Customers using a Left Outer join on CustomerID
3. Expand the merged columns to bring in Name and City
4. Import three monthly CSV files (January, February, March)
5. **Append** them into a single "Q1 Sales" query
6. Group by Product and calculate total sales per product

### Exercise 3: Unpivot and Reshape

**Scenario:** A report exports data in cross-tab format.

**Tasks:**
1. Import a table with months as columns (Jan, Feb, Mar...) and products as rows
2. **Unpivot** the month columns to create a flat list
3. Rename the "Attribute" column to "Month" and "Value" column to "Sales"
4. Split a combined "ProductCode-ProductName" column into two columns
5. Pivot the data back to cross-tab format for verification

### Exercise 4: M Language and Parameters

**Tasks:**
1. Create a parameter called `MinAmount` (type: Decimal Number, value: 100)
2. Write a custom column formula that categorizes amounts: "Low" (<MinAmount), "Medium" (MinAmount-500), "High" (>500)
3. Open the Advanced Editor and read the M code for your query
4. Add a custom step using M: `Table.SelectRows(#"Previous Step", each [Amount] > MinAmount)`
5. Create a parameter for a file path and use it in the Source step

### Exercise 5: Data Model and Relationships

**Scenario:** Build a star schema for sales analysis.

**Tasks:**
1. Create/load these tables into the Data Model:
   - **FactSales** (OrderID, DateKey, ProductID, CustomerID, Amount, Quantity)
   - **DimDate** (DateKey, Date, Year, Quarter, Month, DayOfWeek)
   - **DimProduct** (ProductID, ProductName, Category, SubCategory)
   - **DimCustomer** (CustomerID, CustomerName, City, State)
2. Create relationships between FactSales and each dimension table
3. Build a PivotTable showing **Total Sales by Category and Year**
4. Add a slicer for City
5. Create a hierarchy: Year → Quarter → Month

### Exercise 6: DAX Measures

**Tasks:**
1. Create a measure: `Total Sales = SUM(FactSales[Amount])`
2. Create a measure: `Order Count = COUNTROWS(FactSales)`
3. Create a measure: `Avg Order Value = DIVIDE([Total Sales], [Order Count], 0)`
4. Create a measure for **Prior Year Sales** using CALCULATE and DATEADD
5. Create a measure for **YoY Growth %** = (Current - Prior) / Prior
6. Create a **KPI** comparing Total Sales to a Sales Target
7. Create a measure using SUMX: `Revenue = SUMX(FactSales, FactSales[Amount] * FactSales[Quantity])`

### Exercise 7: End-to-End Project

**Scenario:** Build a complete Sales Analytics Dashboard.

**Tasks:**
1. **Power Query:** Import sales CSV files from a folder, clean and transform
2. **Power Query:** Merge with a product lookup table and a customer table
3. **Data Model:** Load all tables, create relationships
4. **DAX:** Create measures for Total Revenue, Profit, Profit Margin, YoY Growth
5. **PivotTable:** Build a summary with Category × Year, with drill-down hierarchies
6. **KPI:** Add a KPI showing performance against targets
7. **Dashboard:** Create PivotCharts connected to your PivotTable
8. **Refresh:** Add new data to the source folder, refresh all queries, verify updates

---

## 21. Sources and Further Reading

### Microsoft Official Documentation
- [Power Query Overview - Microsoft Support](https://support.microsoft.com/en-us/excel)
- [Connect to Data Sources - Microsoft Support](https://support.microsoft.com/en-us/excel)
- [Merge Queries - Microsoft Support](https://support.microsoft.com/en-us/excel)
- [Power Pivot Overview - Microsoft Support](https://support.microsoft.com/en-us/excel)
- [DAX Reference - Microsoft Support](https://learn.microsoft.com/en-us/dax/)
- [CALCULATE Function - Microsoft DAX Reference](https://learn.microsoft.com/en-us/dax/calculate-function-dax)
- [Data Model in Excel - Microsoft Support](https://support.microsoft.com/en-us/excel)

### Exceljet
- [Power Query Overview - Exceljet](https://exceljet.net/)
- [DAX Functions Reference - Exceljet](https://exceljet.net/)

### Community Resources
- [Power Query Tutorial - MyOnlineTrainingHub](https://www.myonlinetraininghub.com/power-query-overview)
- [DAX Guide - Curbal](https://www.curbal.com/)
- [Power Query Tips - Chandoo.org](https://chandoo.org/?s=power+query)
- [Power Pivot and DAX - Excel Off The Grid](https://exceloffthegrid.com/)
- [M Language Reference - Microsoft](https://learn.microsoft.com/en-us/powerquery-m/)

### YouTube Video References
- [Power Query Full Tutorial - Leila Gharani](https://www.youtube.com/watch?v=ohG6zAslhXM)
- [Power Query for Beginners - ExcelIsFun](https://www.youtube.com/watch?v=fHF6ZMPiFmI)
- [Merge Queries (JOIN) Explained - Curbal](https://www.youtube.com/watch?v=YBzFgt5qWOc)
- [Append Queries - Curbal](https://www.youtube.com/watch?v=kwtmE5sMf6k)
- [M Language Basics - Curbal](https://www.youtube.com/watch?v=0yCyxjPFGAM)
- [Power Pivot Full Tutorial - Leila Gharani](https://www.youtube.com/watch?v=Q3EiNBMPxHg)
- [DAX for Beginners - Curbal](https://www.youtube.com/watch?v=TMR2MBKDOuI)
- [CALCULATE Function Deep Dive - Curbal](https://www.youtube.com/watch?v=Jb2DxhfDCb8)
- [Data Model and Relationships - MyOnlineTrainingHub](https://www.youtube.com/watch?v=Yb-Sahv0K3g)
- [Power Query Parameters - Excel Off The Grid](https://www.youtube.com/watch?v=5D3K0CbOFoM)
- [DAX SUMX vs SUM - Curbal](https://www.youtube.com/watch?v=ciHIenRJaKI)
- [Power Query Folder Connection - ExcelIsFun](https://www.youtube.com/watch?v=dHxNdVJSPCE)

### Books
- *Power Query for Power BI and Excel* - Chris Webb
- *The Definitive Guide to DAX* - Alberto Ferrari & Marco Russo
- *Power Pivot and Power BI* - Rob Collie & Avichal Singh
- *M Is for (Data) Monkey* - Ken Puls & Miguel Escobar

---

> **Previous Module:** [Module 5 - Advanced Functions](../05-Advanced-Functions/05-advanced-functions.md)


## FILE: 07-VBA-Macros/07-vba-macros.md

# Module 7: VBA & Macros - Excel Automation Mastery

> **Module Goal:** Master Visual Basic for Applications (VBA) to automate repetitive Excel tasks, build custom functions, create interactive forms, and develop professional-grade automation solutions.

---

## Table of Contents

1. [What is VBA? Why Automate Excel?](#1-what-is-vba-why-automate-excel)
2. [Developer Tab: Enabling and Using](#2-developer-tab-enabling-and-using)
3. [Recording Macros: Step by Step](#3-recording-macros-step-by-step)
4. [Running Macros](#4-running-macros)
5. [VBA Editor (Alt+F11)](#5-vba-editor-altf11)
6. [VBA Basics](#6-vba-basics)
7. [Working with Worksheets and Workbooks via VBA](#7-working-with-worksheets-and-workbooks-via-vba)
8. [Error Handling](#8-error-handling)
9. [UserForms Basics](#9-userforms-basics)
10. [Common Automation Tasks](#10-common-automation-tasks)
11. [Best Practices for VBA Code](#11-best-practices-for-vba-code)
12. [Practice Exercises](#12-practice-exercises)
13. [Video References](#13-video-references)
14. [Sources & Further Reading](#14-sources--further-reading)

---

## 1. What is VBA? Why Automate Excel?

### What is VBA?

**VBA** stands for **Visual Basic for Applications**. It is a programming language developed by Microsoft and embedded into several Office applications including Excel, Word, Outlook, and Access.

VBA allows you to:

- **Automate repetitive tasks** - Format reports, clean data, generate summaries with one click
- **Create custom functions** - Build functions that don't exist in Excel's built-in library
- **Build interactive applications** - Create UserForms for data entry, dashboards, and mini-apps
- **Control other Office apps** - Send emails via Outlook, create Word documents, interact with Access databases
- **Extend Excel's capabilities** - Perform tasks impossible with formulas alone

### Why Automate Excel?

| Manual Approach | VBA Automation |
|----------------|----------------|
| Format 50 reports manually | One macro formats all 50 in seconds |
| Copy-paste data daily | Auto-import with a single click |
| Build charts one by one | Generate all charts programmatically |
| Send weekly email summaries | VBA sends them automatically via Outlook |
| Clean messy data row by row | Loop through thousands of rows instantly |

### Real-World Impact

- **Finance:** Automated financial models, scenario analysis, report generation
- **HR:** Employee data management, leave tracking, payroll calculations
- **Marketing:** Campaign analysis, data consolidation, automated dashboards
- **Operations:** Inventory tracking, order processing, supply chain monitoring
- **Education:** Grade calculations, attendance tracking, report cards

### VBA vs. Other Approaches

| Feature | Formulas | Power Query | VBA |
|---------|----------|-------------|-----|
| Repetitive formatting | ❌ | ❌ | ✅ |
| Custom dialogs | ❌ | ❌ | ✅ |
| Data transformation | Limited | ✅ | ✅ |
| Conditional logic | Limited | ✅ | ✅ |
| External app control | ❌ | ❌ | ✅ |
| User interaction | ❌ | ❌ | ✅ |
| Learning curve | Low | Medium | Medium-High |

> **Key Insight:** VBA is not a replacement for formulas or Power Query - it complements them. Use formulas for calculations, Power Query for data transformation, and VBA for automation and interaction.

---

## 2. Developer Tab: Enabling and Using

### Why Enable the Developer Tab?

The Developer Tab provides access to:
- **Macro recording and running**
- **VBA Editor** (Visual Basic button)
- **Form Controls** (buttons, checkboxes, dropdowns)
- **ActiveX Controls** (advanced interactive controls)
- **XML and Add-in management**

### How to Enable the Developer Tab

#### Excel 365 / 2021 / 2019 / 2016:

1. Go to **File** → **Options**
2. Click **Customize Ribbon** in the left panel
3. In the right panel, check the box next to **Developer**
4. Click **OK**

#### Excel for Mac:

1. Go to **Excel** → **Preferences** → **Ribbon & Toolbar**
2. In the right panel, check **Developer**
3. Click **Save**

### Developer Tab Components

```
Developer Tab
├── Code Group
│   ├── Visual Basic (opens VBA Editor)
│   ├── Macros (view/run macros)
│   ├── Record Macro
│   ├── Use Relative References
│   └── Macro Security
├── Add-ins Group
│   ├── Excel Add-ins
│   └── COM Add-ins
├── Controls Group
│   ├── Insert (Form & ActiveX controls)
│   ├── Design Mode
│   ├── Properties
│   └── View Code
├── XML Group
│   ├── Source
│   ├── Map
│   └── Expansion Packs
└── Modify Group
    └── Document Panel
```

### Macro Security Settings

1. Go to **Developer** → **Macro Security**
2. Choose your security level:

| Setting | Description | Recommended For |
|---------|-------------|-----------------|
| Disable all macros without notification | Blocks all macros | High-security environments |
| Disable all macros with notification | Shows warning bar | **Most users (recommended)** |
| Disable all macros except digitally signed | Only trusted macros | Corporate environments |
| Enable all macros | Runs everything (not recommended) | Development only |

> **Best Practice:** Use "Disable all macros with notification" - you'll see a yellow security bar when opening files with macros, and can choose to enable them.

---

## 3. Recording Macros: Step by Step

### What is Macro Recording?

Macro recording captures your actions in Excel and converts them into VBA code. It's the easiest way to start learning VBA - record actions, then examine the generated code.

### Step-by-Step: Recording Your First Macro

#### Example: Auto-Format a Sales Report

**Scenario:** You receive a weekly sales report and always format it the same way.

1. **Start Recording:**
   - Go to **Developer** → **Record Macro**
   - Or click the small circle in the bottom-left status bar

2. **Set Macro Properties:**
   - **Macro name:** `FormatSalesReport`
   - **Shortcut key:** `Ctrl+Shift+F`
   - **Store in:** This Workbook
   - **Description:** Formats sales report with standard formatting

3. **Perform Your Actions:**
   ```
   1. Select A1 (title cell)
   2. Bold the text, set font size to 14
   3. Select A1:F1, apply Merge & Center
   4. Select the header row (A2:F2)
   5. Bold, apply background color (blue), white font
   6. Select data range A2:F100
   7. Apply borders (all borders)
   8. Auto-fit columns
   9. Select column D (revenue), apply Currency format
   10. Freeze panes at row 3
   ```

4. **Stop Recording:**
   - Go to **Developer** → **Stop Recording**
   - Or click the square in the status bar

### Examining the Recorded Code

Press **Alt+F11** to open the VBA Editor. The recorded macro looks like this:

```vba
Sub FormatSalesReport()
    ' Formats sales report with standard formatting
    ' Recorded on 2026-09-10
    
    ' Format title
    Range("A1").Select
    Selection.Font.Bold = True
    Selection.Font.Size = 14
    Range("A1:F1").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .MergeCells = True
    End With
    
    ' Format headers
    Range("A2:F2").Select
    With Selection
        .Font.Bold = True
        .Interior.Color = RGB(0, 51, 102)
        .Font.Color = RGB(255, 255, 255)
    End With
    
    ' Apply borders
    Range("A2:F100").Select
    Selection.Borders.LineStyle = xlContinuous
    
    ' Auto-fit columns
    Cells.Select
    Cells.EntireColumn.AutoFit
    
    ' Format revenue column as currency
    Range("D2:D100").Select
    Selection.NumberFormat = "$#,##0.00"
    
    ' Freeze panes
    Range("A3").Select
    ActiveWindow.FreezePanes = True
End Sub
```

### Improving Recorded Macros

Recorded macros work but are often inefficient. Here's the improved version:

```vba
Sub FormatSalesReport_Improved()
    ' Improved version - no Select statements
    
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    ' Format title
    With ws.Range("A1")
        .Font.Bold = True
        .Font.Size = 14
    End With
    With ws.Range("A1:F1")
        .HorizontalAlignment = xlCenter
        .MergeCells = True
    End With
    
    ' Format headers
    With ws.Range("A2:F2")
        .Font.Bold = True
        .Interior.Color = RGB(0, 51, 102)
        .Font.Color = RGB(255, 255, 255)
    End With
    
    ' Apply borders
    With ws.Range("A2:F100").Borders
        .LineStyle = xlContinuous
    End With
    
    ' Auto-fit columns
    ws.Cells.EntireColumn.AutoFit
    
    ' Format revenue column as currency
    ws.Range("D2:D100").NumberFormat = "$#,##0.00"
    
    ' Freeze panes
    ws.Range("A3").Select
    ActiveWindow.FreezePanes = True
    
    MsgBox "Report formatted successfully!", vbInformation
End Sub
```

### Macro Recording Best Practices

1. **Plan your actions** before recording
2. **Use relative references** when needed (Developer → Use Relative References)
3. **Keep recordings short** - break complex tasks into multiple macros
4. **Always examine and clean** the generated code
5. **Remove unnecessary `.Select` statements** - they slow down execution
6. **Test immediately** after recording
7. **Add comments** to explain what the macro does

---

## 4. Running Macros

### Method 1: Macro Dialog Box

1. Press **Alt+F8**
2. Select the macro from the list
3. Click **Run**

### Method 2: Keyboard Shortcut

- Assign during recording or via **Alt+F8** → **Options**
- Default: `Ctrl+letter` (e.g., `Ctrl+Shift+F`)

### Method 3: Button on Worksheet

1. Go to **Developer** → **Insert** → **Button (Form Control)**
2. Draw the button on your worksheet
3. In the Assign Macro dialog, select your macro
4. Right-click the button to edit text

```vba
' You can also create buttons programmatically:
Sub CreateMacroButton()
    Dim btn As Button
    Set btn = ActiveSheet.Buttons.Add(10, 10, 120, 30)
    With btn
        .Caption = "Format Report"
        .OnAction = "FormatSalesReport"
        .Font.Size = 11
    End With
End Sub
```

### Method 4: Quick Access Toolbar

1. Click the dropdown arrow on the Quick Access Toolbar
2. Select **More Commands**
3. Choose **Macros** from the "Choose commands from" dropdown
4. Select your macro and click **Add**
5. Click **Modify** to change the icon

### Method 5: Ribbon Button

1. Go to **File** → **Options** → **Customize Ribbon**
2. Create a new tab or group
3. Select **Macros** from the dropdown
4. Add your macro to the custom group

### Method 6: From VBA Editor

- Press **Alt+F11** to open VBA Editor
- Place cursor inside the macro
- Press **F5** or click **Run**

### Method 7: Auto-Run Macros

```vba
' Runs automatically when the workbook opens
Private Sub Workbook_Open()
    MsgBox "Welcome! Workbook is ready."
End Sub

' Runs automatically when a cell changes
Private Sub Worksheet_Change(ByVal Target As Range)
    If Target.Address = "$A$1" Then
        MsgBox "Cell A1 changed to: " & Target.Value
    End If
End Sub
```

---

## 5. VBA Editor (Alt+F11)

### Opening the VBA Editor

Press **Alt+F11** from Excel to open the Visual Basic Editor.

### VBA Editor Layout

```
┌─────────────────────────────────────────────────────────────┐
│  VBA Editor - Microsoft Visual Basic for Applications       │
├─────────────────────────────────────────────────────────────┤
│ File  Edit  View  Insert  Format  Debug  Run  Tools  Add-Ins│
├──────────┬──────────────────────────────────────────────────┤
│ Project  │  Code Window                                     │
│ Explorer │  ┌──────────────────────────────────────────────┐│
│          │  │  (Your VBA code goes here)                   ││
│ VBAProj  │  │                                              ││
│  │       │  │                                              ││
│  ├─Sheet1 │  │                                              ││
│  ├─Sheet2 │  │                                              ││
│  ├─ThisWb │  │                                              ││
│  └─Module1│  │                                              ││
│          │  └──────────────────────────────────────────────┘│
│──────────┤──────────────────────────────────────────────────┤
│Properties│  Immediate Window (Ctrl+G)                       │
│          │  ┌──────────────────────────────────────────────┐│
│ Name:    │  │  ? Range("A1").Value                         ││
│ Module1  │  │  Hello World                                 ││
│          │  └──────────────────────────────────────────────┘│
└──────────┴──────────────────────────────────────────────────┘
```

### Project Explorer (Ctrl+R)

Shows the hierarchy of all open workbooks and their components:

```
VBAProject (Book1.xlsm)
├── Microsoft Excel Objects
│   ├── Sheet1 (Sheet1)
│   ├── Sheet2 (Sheet2)
│   └── ThisWorkbook
├── Forms
│   └── UserForm1
└── Modules
    ├── Module1
    └── Module2
```

**Where to place code:**

| Location | Use For |
|----------|---------|
| **Module** | General-purpose macros and functions |
| **Sheet object** | Sheet-specific event code (e.g., Worksheet_Change) |
| **ThisWorkbook** | Workbook-level events (e.g., Workbook_Open) |
| **UserForm** | Form control event code |
| **Class Module** | Custom objects and events |

### Properties Window (F4)

Displays properties of the selected object. You can change:
- **Name** - The code name of the object
- **Visible** - Show/hide sheets (xlSheetVisible, xlSheetHidden, xlSheetVeryHidden)
- **Tab Color** - Sheet tab color

### Code Window

Where you write and edit VBA code. Features:
- **Syntax highlighting** - Keywords, comments, strings in different colors
- **Auto-complete** - Type `Range("` and see suggestions
- **Procedure dropdown** - Jump to any Sub/Function in the module
- **Split bar** - Divide the window to view two code sections

### Immediate Window (Ctrl+G)

A powerful debugging tool:

```vba
' Test expressions immediately (prefix with ?)
? Range("A1").Value          ' Returns: Hello World
? ActiveSheet.Name           ' Returns: Sheet1
? Cells(1, 1).Address        ' Returns: $A$1

' Execute statements
Range("A1").Value = "Updated"
Sheets("Sheet1").Activate

' Debug output
Debug.Print "Variable x = " & x
```

### Essential VBA Editor Shortcuts

| Shortcut | Action |
|----------|--------|
| Alt+F11 | Toggle between Excel and VBA Editor |
| F5 | Run the current procedure |
| F8 | Step through code (one line at a time) |
| F9 | Toggle breakpoint |
| Ctrl+G | Show Immediate Window |
| Ctrl+R | Show Project Explorer |
| F4 | Show Properties Window |
| Ctrl+Space | Auto-complete |
| Ctrl+J | List members |
| Ctrl+Shift+I | Quick Info |
| Ctrl+Z | Undo |
| Ctrl+F | Find |
| Ctrl+H | Find and Replace |

---

## 6. VBA Basics

### 6.1 Sub and Function Procedures

#### Sub Procedures

A **Sub** performs an action but doesn't return a value:

```vba
Sub SayHello()
    MsgBox "Hello, World!"
End Sub

' Sub with parameters
Sub GreetUser(userName As String)
    MsgBox "Hello, " & userName & "!"
End Sub

' Call it:
' GreetUser "Alice"
```

#### Function Procedures

A **Function** performs a calculation and returns a value:

```vba
' Custom function usable in worksheets
Function CalculateTax(price As Double, taxRate As Double) As Double
    CalculateTax = price * taxRate
End Function

' Usage in a cell: =CalculateTax(100, 0.08) returns 8

' Function with optional parameter
Function FullName(firstName As String, lastName As String, _
                  Optional middleName As String = "") As String
    If middleName = "" Then
        FullName = firstName & " " & lastName
    Else
        FullName = firstName & " " & middleName & " " & lastName
    End If
End Function
```

#### Difference Between Sub and Function

| Feature | Sub | Function |
|---------|-----|----------|
| Returns value | No | Yes |
| Called from worksheet | No | Yes (=FunctionName()) |
| Called from VBA | `Call SubName` or `SubName` | `result = FunctionName()` |
| Can modify cells | Yes | Yes (but not recommended) |
| Entry point | Can be assigned to buttons | Used in formulas |

### 6.2 Variables and Data Types

#### Declaring Variables

```vba
' Explicit declaration (recommended)
Dim employeeName As String
Dim age As Integer
Dim salary As Double
Dim isActive As Boolean
Dim startDate As Date

' Multiple declarations on one line
Dim x As Integer, y As Integer, z As Integer

' Module-level variable (available to all procedures in the module)
Dim totalCount As Long

' Public variable (available to all modules)
Public companyName As String
```

#### VBA Data Types

| Data Type | Storage | Range | Example |
|-----------|---------|-------|---------|
| **Byte** | 1 byte | 0 to 255 | `Dim b As Byte` |
| **Boolean** | 2 bytes | True/False | `Dim flag As Boolean` |
| **Integer** | 2 bytes | -32,768 to 32,767 | `Dim i As Integer` |
| **Long** | 4 bytes | -2.1B to 2.1B | `Dim l As Long` |
| **LongLong** | 8 bytes | Very large numbers | `Dim ll As LongLong` |
| **Single** | 4 bytes | Decimal (7 digits) | `Dim s As Single` |
| **Double** | 8 bytes | Decimal (15 digits) | `Dim d As Double` |
| **Currency** | 8 bytes | Financial precision | `Dim c As Currency` |
| **Date** | 8 bytes | Dates and times | `Dim dt As Date` |
| **String** | Varies | Text (up to 2GB) | `Dim str As String` |
| **Variant** | Varies | Any type (default) | `Dim v As Variant` |
| **Object** | 4 bytes | Object reference | `Dim ws As Worksheet` |

#### Choosing the Right Data Type

```vba
' Use Long instead of Integer (faster on modern systems)
Dim rowNumber As Long  ' NOT Integer

' Use Double for calculations
Dim revenue As Double  ' NOT Single (precision matters)

' Use String for text
Dim customerName As String

' Use Date for dates
Dim orderDate As Date

' Avoid Variant when possible (uses more memory)
' BAD:  Dim x As Variant  (or just Dim x)
' GOOD: Dim x As Long
```

### 6.3 Option Explicit

**Always use Option Explicit!** It forces you to declare all variables, preventing typos from creating new variables.

```vba
Option Explicit  ' Must be at the TOP of every module

Sub Example()
    Dim counter As Integer
    conuter = 5  ' TYPO! Without Option Explicit, this creates a new variable
    ' With Option Explicit: ERROR! "Variable not defined"
    Debug.Print counter  ' Returns 0 (not 5!) without Option Explicit
End Sub
```

**To set Option Explicit automatically:**
1. In VBA Editor, go to **Tools** → **Options**
2. Check **Require Variable Declaration**
3. Click **OK**

### 6.4 MsgBox and InputBox

#### MsgBox

```vba
' Simple message
MsgBox "Operation completed!"

' With title
MsgBox "Data saved successfully.", vbInformation, "Success"

' With buttons and capturing response
Dim response As VbMsgBoxResult
response = MsgBox("Do you want to continue?", vbYesNoCancel + vbQuestion, "Confirm")

Select Case response
    Case vbYes
        ' User clicked Yes
        MsgBox "You chose Yes"
    Case vbNo
        ' User clicked No
        MsgBox "You chose No"
    Case vbCancel
        ' User clicked Cancel
        MsgBox "Operation cancelled"
End Select

' MsgBox buttons and icons:
' vbOKOnly (0)           vbOKCancel (1)
' vbAbortRetryIgnore (2) vbYesNoCancel (3)
' vbYesNo (4)            vbRetryCancel (5)
'
' vbCritical (16)        vbQuestion (32)
' vbExclamation (48)     vbInformation (64)
```

#### InputBox

```vba
' Simple input
Dim userName As String
userName = InputBox("Enter your name:", "User Input")

' With default value
Dim city As String
city = InputBox("Enter your city:", "Location", "New York")

' Validate input
Dim age As String
age = InputBox("Enter your age:", "Age")
If age = "" Then
    MsgBox "No input provided."
ElseIf Not IsNumeric(age) Then
    MsgBox "Please enter a valid number."
Else
    MsgBox "You are " & age & " years old."
End If

' Application.InputBox (more powerful, allows range selection)
Dim rng As Range
On Error Resume Next
Set rng = Application.InputBox("Select a range:", Type:=8)  ' Type 8 = Range
On Error GoTo 0
If rng Is Nothing Then
    MsgBox "No range selected."
Else
    MsgBox "You selected: " & rng.Address
End If
```

### 6.5 Range and Cells Object Model

#### The Range Object

```vba
' Referencing ranges
Range("A1")                    ' Single cell
Range("A1:D10")                ' Range of cells
Range("A1, C1, E1")           ' Non-contiguous range
Range("A1:" & "D" & lastRow)  ' Dynamic range

' Using Cells (row, column) - great for loops
Cells(1, 1)        ' A1
Cells(5, 3)        ' C5
Cells(Rows.Count, 1)  ' Last cell in column A

' Combining Range and Cells
Range(Cells(1, 1), Cells(10, 5))  ' A1:E10

' Named ranges
Range("SalesData")

' Entire rows and columns
Range("A:A")           ' Entire column A
Range("1:1")           ' Entire row 1
Columns("A:D")         ' Columns A through D
Rows("1:5")            ' Rows 1 through 5
```

#### Working with Range Properties

```vba
' Value
Range("A1").Value = "Hello"
Dim val As String
val = Range("A1").Value

' Address
Debug.Print Range("A1").Address          ' $A$1
Debug.Print Range("A1").Address(False, False)  ' A1 (relative)

' Offset
Range("A1").Offset(1, 0)   ' A2 (down 1, right 0)
Range("A1").Offset(0, 2)   ' C1 (down 0, right 2)

' Resize
Range("A1").Resize(3, 4)   ' A1:D3

' End (like Ctrl+Arrow)
Range("A1").End(xlDown)     ' Last filled cell down from A1
Range("A1").End(xlToRight)  ' Last filled cell right from A1

' Find last row and column
Dim lastRow As Long
Dim lastCol As Long
lastRow = Cells(Rows.Count, 1).End(xlUp).Row
lastCol = Cells(1, Columns.Count).End(xlToLeft).Column

' CurrentRegion (like Ctrl+Shift+*)
Range("A1").CurrentRegion.Select

' SpecialCells
' Find all cells with formulas
ActiveSheet.Cells.SpecialCells(xlCellTypeFormulas).Select
' Find all blank cells
ActiveSheet.Cells.SpecialCells(xlCellTypeBlanks).Select
```

#### Range Actions

```vba
' Clear contents and formatting
Range("A1:D10").Clear           ' Clear everything
Range("A1:D10").ClearContents   ' Clear values only
Range("A1:D10").ClearFormats    ' Clear formatting only

' Copy and Paste
Range("A1:A10").Copy Destination:=Range("C1")
Range("A1:A10").Copy
Range("C1").PasteSpecial xlPasteValues

' Sort
Range("A1:D100").Sort Key1:=Range("A1"), Order1:=xlAscending, Header:=xlYes

' Filter
Range("A1:D100").AutoFilter Field:=1, Criteria1:="Sales"

' Find
Dim foundCell As Range
Set foundCell = Range("A:A").Find(What:="Target", LookIn:=xlValues)
If Not foundCell Is Nothing Then
    MsgBox "Found at: " & foundCell.Address
End If
```

### 6.6 With...End With

```vba
' Without With (repetitive)
Range("A1").Font.Bold = True
Range("A1").Font.Size = 14
Range("A1").Font.Color = RGB(255, 0, 0)
Range("A1").Font.Name = "Arial"

' With With (cleaner)
With Range("A1").Font
    .Bold = True
    .Size = 14
    .Color = RGB(255, 0, 0)
    .Name = "Arial"
End With

' Nested With
With Range("A1:D1")
    .Font.Bold = True
    .Interior.Color = RGB(0, 51, 102)
    .Font.Color = RGB(255, 255, 255)
    With .Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
End With
```

### 6.7 Conditional Statements

#### If...Then...Else

```vba
' Simple If
If Range("A1").Value > 100 Then
    MsgBox "Greater than 100"
End If

' If...Else
If Range("A1").Value > 100 Then
    Range("B1").Value = "High"
Else
    Range("B1").Value = "Low"
End If

' If...ElseIf...Else
Dim score As Double
score = Range("A1").Value

If score >= 90 Then
    Range("B1").Value = "A"
ElseIf score >= 80 Then
    Range("B1").Value = "B"
ElseIf score >= 70 Then
    Range("B1").Value = "C"
ElseIf score >= 60 Then
    Range("B1").Value = "D"
Else
    Range("B1").Value = "F"
End If

' Single-line If (for simple cases)
If Range("A1").Value = "" Then Range("A1").Value = "N/A"

' If with logical operators
If Range("A1").Value > 0 And Range("B1").Value > 0 Then
    MsgBox "Both positive"
End If

If Range("A1").Value = "Yes" Or Range("A1").Value = "Y" Then
    MsgBox "Confirmed"
End If
```

#### Select Case

```vba
' Select Case is cleaner for multiple conditions
Dim dept As String
dept = Range("A1").Value

Select Case dept
    Case "Sales"
        Range("B1").Value = "Revenue Generation"
    Case "Marketing"
        Range("B1").Value = "Brand Awareness"
    Case "IT"
        Range("B1").Value = "Technology Support"
    Case "HR"
        Range("B1").Value = "People Management"
    Case "Finance"
        Range("B1").Value = "Financial Operations"
    Case Else
        Range("B1").Value = "Unknown Department"
End Select

' Select Case with ranges
Select Case score
    Case Is >= 90
        grade = "A"
    Case 80 To 89
        grade = "B"
    Case 70 To 79
        grade = "C"
    Case 60 To 69
        grade = "D"
    Case Else
        grade = "F"
End Select

' Select Case with multiple values
Select Case weekday
    Case 1, 7
        dayType = "Weekend"
    Case 2 To 6
        dayType = "Weekday"
End Select
```

### 6.8 Loops

#### For...Next

```vba
' Basic For loop
Dim i As Long
For i = 1 To 10
    Cells(i, 1).Value = i * 10
Next i

' With Step
For i = 1 To 100 Step 5
    Cells(i, 1).Value = "Every 5th row"
Next i

' Reverse loop
For i = 100 To 1 Step -1
    If Cells(i, 1).Value = "" Then
        Cells(i, 1).EntireRow.Delete
    End If
Next i

' Loop through rows with data
Dim lastRow As Long
lastRow = Cells(Rows.Count, 1).End(xlUp).Row

For i = 2 To lastRow  ' Start at 2 to skip header
    If Cells(i, 3).Value > 1000 Then
        Cells(i, 3).Interior.Color = RGB(0, 255, 0)  ' Highlight
    End If
Next i
```

#### For Each...Next

```vba
' Loop through cells in a range
Dim cell As Range
For Each cell In Range("A1:A100")
    If cell.Value < 0 Then
        cell.Font.Color = RGB(255, 0, 0)  ' Red for negatives
    End If
Next cell

' Loop through all worksheets
Dim ws As Worksheet
For Each ws In ThisWorkbook.Worksheets
    Debug.Print ws.Name
    ws.Range("A1").Value = "Updated"
Next ws

' Loop through all open workbooks
Dim wb As Workbook
For Each wb In Application.Workbooks
    Debug.Print wb.Name
Next wb

' Loop through selected cells
For Each cell In Selection
    cell.Value = UCase(cell.Value)  ' Convert to uppercase
Next cell
```

#### Do While / Do Until

```vba
' Do While - runs while condition is True
Dim row As Long
row = 1
Do While Cells(row, 1).Value <> ""
    Cells(row, 2).Value = Cells(row, 1).Value * 2
    row = row + 1
Loop

' Do Until - runs until condition becomes True
row = 1
Do Until IsEmpty(Cells(row, 1))
    Cells(row, 2).Value = Cells(row, 1).Value * 2
    row = row + 1
Loop

' Do...Loop While (runs at least once)
Do
    Cells(row, 2).Value = Cells(row, 1).Value * 2
    row = row + 1
Loop While Cells(row, 1).Value <> ""

' Exit Do - break out of loop
Do While True
    If Cells(row, 1).Value = "STOP" Then Exit Do
    row = row + 1
Loop
```

### 6.9 Arrays

#### Static Arrays

```vba
' Declare a static array
Dim names(1 To 5) As String
names(1) = "Alice"
names(2) = "Bob"
names(3) = "Charlie"
names(4) = "Diana"
names(5) = "Eve"

' Multi-dimensional array
Dim matrix(1 To 3, 1 To 3) As Double
matrix(1, 1) = 1: matrix(1, 2) = 2: matrix(1, 3) = 3
matrix(2, 1) = 4: matrix(2, 2) = 5: matrix(2, 3) = 6
matrix(3, 1) = 7: matrix(3, 2) = 8: matrix(3, 3) = 9

' Initialize with Array function
Dim colors As Variant
colors = Array("Red", "Green", "Blue", "Yellow")
```

#### Dynamic Arrays

```vba
' Dynamic array - size determined at runtime
Dim data() As Variant
Dim lastRow As Long
lastRow = Cells(Rows.Count, 1).End(xlUp).Row

' Read range into array (FAST!)
data = Range("A1:D" & lastRow).Value

' Process array (much faster than cell-by-cell)
Dim i As Long
For i = 1 To UBound(data, 1)
    data(i, 4) = data(i, 2) * data(i, 3)  ' Calculate total
Next i

' Write array back to range (FAST!)
Range("A1:D" & lastRow).Value = data
```

#### Array Performance Tip

```vba
' SLOW - reading/writing cells one by one
Sub SlowMethod()
    Dim i As Long
    For i = 1 To 10000
        Cells(i, 2).Value = Cells(i, 1).Value * 2
    Next i
End Sub

' FAST - using arrays
Sub FastMethod()
    Dim data() As Variant
    Dim result() As Variant
    Dim i As Long
    
    ' Read all data at once
    data = Range("A1:A10000").Value
    
    ' Prepare result array
    ReDim result(1 To UBound(data, 1), 1 To 1)
    
    ' Process in memory
    For i = 1 To UBound(data, 1)
        result(i, 1) = data(i, 1) * 2
    Next i
    
    ' Write all results at once
    Range("B1:B10000").Value = result
End Sub
```

---

## 7. Working with Worksheets and Workbooks via VBA

### Worksheet Operations

```vba
' Reference worksheets
Dim ws As Worksheet
Set ws = ThisWorkbook.Sheets("Sales")          ' By name
Set ws = ThisWorkbook.Sheets(1)                ' By index
Set ws = ActiveSheet                            ' Currently active

' Create new worksheet
Dim newWs As Worksheet
Set newWs = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
newWs.Name = "NewSheet"

' Delete worksheet (with confirmation)
Application.DisplayAlerts = False
ThisWorkbook.Sheets("TempSheet").Delete
Application.DisplayAlerts = True

' Copy worksheet
ThisWorkbook.Sheets("Template").Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)

' Move worksheet
ThisWorkbook.Sheets("Sheet1").Move Before:=ThisWorkbook.Sheets(1)

' Hide/Show worksheets
ws.Visible = xlSheetHidden       ' Hidden (user can unhide)
ws.Visible = xlSheetVeryHidden   ' Very Hidden (only via VBA)
ws.Visible = xlSheetVisible      ' Visible

' Rename worksheet
ws.Name = "Q1 Sales"

' Protect/Unprotect
ws.Protect Password:="myPassword", AllowFiltering:=True
ws.Unprotect Password:="myPassword"

' Loop through all sheets
Dim sheet As Worksheet
For Each sheet In ThisWorkbook.Worksheets
    Debug.Print sheet.Name & " - " & sheet.Cells(Rows.Count, 1).End(xlUp).Row & " rows"
Next sheet
```

### Workbook Operations

```vba
' Create new workbook
Dim newWb As Workbook
Set newWb = Workbooks.Add

' Open workbook
Dim existingWb As Workbook
Set existingWb = Workbooks.Open("C:\Data\Sales.xlsx")

' Save workbook
ThisWorkbook.Save
existingWb.SaveAs "C:\Data\Sales_Backup.xlsx", xlOpenXMLWorkbook

' Close workbook
existingWb.Close SaveChanges:=True

' ThisWorkbook vs ActiveWorkbook
' ThisWorkbook = workbook containing the VBA code
' ActiveWorkbook = currently active workbook (may be different!)

' Get workbook path
Debug.Print ThisWorkbook.Path        ' C:\Users\Name\Documents
Debug.Print ThisWorkbook.FullName    ' C:\Users\Name\Documents\Book1.xlsm
Debug.Print ThisWorkbook.Name        ' Book1.xlsm

' Copy data between workbooks
SourceWorkbook.Sheets("Data").Range("A1:D100").Copy _
    Destination:=ThisWorkbook.Sheets("Import").Range("A1")
```

---

## 8. Error Handling

### Why Error Handling Matters

Without error handling, VBA crashes with a confusing dialog. With proper error handling, you can:
- Display meaningful error messages
- Log errors for debugging
- Recover gracefully from unexpected situations
- Prevent data corruption

### On Error GoTo

```vba
Sub ProcessData()
    On Error GoTo ErrorHandler
    
    ' Code that might fail
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("NonExistentSheet")  ' This will error!
    
    ' More processing...
    ws.Range("A1").Value = "Done"
    
    Exit Sub  ' Important: exit before error handler
    
ErrorHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description & vbCrLf & _
           "In procedure: ProcessData", vbCritical, "Error"
    
    ' Optional: log error to a file
    Open ThisWorkbook.Path & "\error.log" For Append As #1
    Print #1, Now & " - Error " & Err.Number & ": " & Err.Description
    Close #1
End Sub
```

### On Error Resume Next

```vba
Sub SafeDeleteSheet()
    On Error Resume Next  ' Ignore errors temporarily
    
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets("TempData").Delete
    Application.DisplayAlerts = True
    
    If Err.Number <> 0 Then
        Debug.Print "Sheet didn't exist or couldn't be deleted"
        Err.Clear  ' Reset error
    End If
    
    On Error GoTo 0  ' Resume normal error handling
End Sub

' Check specific error numbers
Sub ReadFile()
    On Error Resume Next
    
    Dim filePath As String
    filePath = "C:\NonExistent\file.xlsx"
    
    Workbooks.Open filePath
    
    Select Case Err.Number
        Case 0
            MsgBox "File opened successfully"
        Case 1004
            MsgBox "File not found: " & filePath
        Case Else
            MsgBox "Unexpected error: " & Err.Description
    End Select
    
    On Error GoTo 0
End Sub
```

### Defensive Programming

```vba
Sub ProcessWorksheet(sheetName As String)
    ' Check if sheet exists before processing
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    
    If ws Is Nothing Then
        MsgBox "Sheet '" & sheetName & "' not found!", vbExclamation
        Exit Sub
    End If
    
    ' Check if range has data
    If Application.WorksheetFunction.CountA(ws.UsedRange) = 0 Then
        MsgBox "Sheet '" & sheetName & "' is empty!", vbExclamation
        Exit Sub
    End If
    
    ' Safe to proceed
    ' ... process data ...
End Sub
```

### Error Handling Best Practices

1. **Always use error handling** in production code
2. **Exit Sub before ErrorHandler** to prevent fall-through
3. **Use specific error handling** (Resume Next) sparingly
4. **Clear errors** after handling with `Err.Clear`
5. **Log errors** for debugging
6. **Don't ignore errors** - at minimum, log them
7. **Test error paths** - don't just test the happy path

---

## 9. UserForms Basics

### What is a UserForm?

A UserForm is a custom dialog box that provides a user-friendly interface for data input, selection, and interaction.

### Creating a UserForm

1. In VBA Editor, go to **Insert** → **UserForm**
2. Add controls from the **Toolbox**
3. Set properties in the **Properties Window**
4. Write code for control events

### Common Controls

| Control | Purpose | Key Properties |
|---------|---------|---------------|
| **Label** | Display text | Caption, Font |
| **TextBox** | Text input | Value, MaxLength |
| **ComboBox** | Drop-down list | List, Value, Style |
| **ListBox** | List selection | List, MultiSelect |
| **CommandButton** | Click action | Caption, Default |
| **CheckBox** | Toggle option | Value (True/False) |
| **OptionButton** | Radio selection | Value, GroupName |
| **Frame** | Group controls | Caption |
| **SpinButton** | Numeric up/down | Value, Min, Max |
| **Image** | Display images | Picture |

### Example: Employee Data Entry Form

```vba
' UserForm Code (frmEmployee)
' Controls: txtName, txtEmail, cboDepartment, txtSalary, btnSubmit, btnCancel

' Initialize form - populate department dropdown
Private Sub UserForm_Initialize()
    With cboDepartment
        .AddItem "Sales"
        .AddItem "Marketing"
        .AddItem "IT"
        .AddItem "Finance"
        .AddItem "HR"
        .AddItem "Operations"
    End With
End Sub

' Submit button click
Private Sub btnSubmit_Click()
    ' Validate inputs
    If Trim(txtName.Value) = "" Then
        MsgBox "Please enter employee name.", vbExclamation
        txtName.SetFocus
        Exit Sub
    End If
    
    If Trim(txtEmail.Value) = "" Then
        MsgBox "Please enter email address.", vbExclamation
        txtEmail.SetFocus
        Exit Sub
    End If
    
    If cboDepartment.ListIndex = -1 Then
        MsgBox "Please select a department.", vbExclamation
        cboDepartment.SetFocus
        Exit Sub
    End If
    
    If Not IsNumeric(txtSalary.Value) Or Val(txtSalary.Value) <= 0 Then
        MsgBox "Please enter a valid salary.", vbExclamation
        txtSalary.SetFocus
        Exit Sub
    End If
    
    ' Find next empty row
    Dim nextRow As Long
    nextRow = Cells(Rows.Count, 1).End(xlUp).Row + 1
    
    ' Write data
    With Cells
        Cells(nextRow, 1).Value = txtName.Value
        Cells(nextRow, 2).Value = txtEmail.Value
        Cells(nextRow, 3).Value = cboDepartment.Value
        Cells(nextRow, 4).Value = CDbl(txtSalary.Value)
        Cells(nextRow, 5).Value = Date
    End With
    
    MsgBox "Employee added successfully!", vbInformation
    
    ' Clear form for next entry
    Call ClearForm
End Sub

' Cancel button
Private Sub btnCancel_Click()
    Unload Me
End Sub

' Helper to clear form
Private Sub ClearForm()
    txtName.Value = ""
    txtEmail.Value = ""
    cboDepartment.ListIndex = -1
    txtSalary.Value = ""
    txtName.SetFocus
End Sub
```

### Showing the Form

```vba
' Show form modally (blocks Excel interaction)
Sub ShowEmployeeForm()
    frmEmployee.Show vbModal
End Sub

' Show form modeless (allows Excel interaction)
Sub ShowEmployeeFormModeless()
    frmEmployee.Show vbModeless
End Sub
```

---

## 10. Common Automation Tasks

### 10.1 Auto-Format Reports

```vba
Sub AutoFormatReport()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim lastCol As Long
    
    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    With ws
        ' Title formatting
        With .Range(.Cells(1, 1), .Cells(1, lastCol))
            .Font.Bold = True
            .Font.Size = 14
            .Font.Color = RGB(0, 51, 102)
            .HorizontalAlignment = xlCenter
            .MergeCells = True
        End With
        
        ' Header formatting
        With .Range(.Cells(2, 1), .Cells(2, lastCol))
            .Font.Bold = True
            .Interior.Color = RGB(0, 51, 102)
            .Font.Color = RGB(255, 255, 255)
            .HorizontalAlignment = xlCenter
        End With
        
        ' Data formatting
        With .Range(.Cells(2, 1), .Cells(lastRow, lastCol))
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
            .RowHeight = 20
        End With
        
        ' Alternating row colors
        Dim i As Long
        For i = 3 To lastRow
            If i Mod 2 = 1 Then
                .Range(.Cells(i, 1), .Cells(i, lastCol)).Interior.Color = RGB(240, 240, 240)
            End If
        Next i
        
        ' Auto-fit columns
        .Cells.EntireColumn.AutoFit
    End With
    
    MsgBox "Report formatted!", vbInformation
End Sub
```

### 10.2 Loop Through Files in a Folder

```vba
Sub ProcessAllFilesInFolder()
    Dim folderPath As String
    Dim fileName As String
    Dim wb As Workbook
    Dim summaryWs As Worksheet
    Dim summaryRow As Long
    
    ' Select folder
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Select Folder"
        If .Show = -1 Then
            folderPath = .SelectedItems(1) & "\"
        Else
            Exit Sub
        End If
    End With
    
    ' Setup summary sheet
    Set summaryWs = ThisWorkbook.Sheets("Summary")
    summaryRow = 2
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    ' Loop through all Excel files
    fileName = Dir(folderPath & "*.xlsx")
    
    Do While fileName <> ""
        Set wb = Workbooks.Open(folderPath & fileName)
        
        ' Process each workbook
        With wb.Sheets(1)
            summaryWs.Cells(summaryRow, 1).Value = fileName
            summaryWs.Cells(summaryRow, 2).Value = .Range("B5").Value  ' e.g., total sales
            summaryWs.Cells(summaryRow, 3).Value = .Range("B10").Value ' e.g., profit
        End With
        
        wb.Close SaveChanges:=False
        summaryRow = summaryRow + 1
        fileName = Dir()  ' Get next file
    Loop
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    MsgBox "Processed " & (summaryRow - 2) & " files.", vbInformation
End Sub
```

### 10.3 Send Emails from Excel (Outlook)

```vba
Sub SendBulkEmails()
    Dim outlookApp As Object
    Dim outlookMail As Object
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    
    Set ws = ThisWorkbook.Sheets("EmailList")
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    ' Create Outlook application
    Set outlookApp = CreateObject("Outlook.Application")
    
    For i = 2 To lastRow  ' Skip header
        If ws.Cells(i, 3).Value = "Pending" Then  ' Column C = Status
            Set outlookMail = outlookApp.CreateItem(0)  ' 0 = olMailItem
            
            With outlookMail
                .To = ws.Cells(i, 1).Value           ' Column A = Email
                .Subject = ws.Cells(i, 2).Value      ' Column B = Subject
                .HTMLBody = "<h3>Hello " & ws.Cells(i, 4).Value & "</h3>" & _
                           "<p>Please find attached the Q3 report.</p>" & _
                           "<p>Best regards,<br>Your Team</p>"
                .Attachments.Add ThisWorkbook.Path & "\Reports\Q3_Report.pdf"
                .Send  ' Or .Display to review before sending
            End With
            
            ' Update status
            ws.Cells(i, 3).Value = "Sent"
            ws.Cells(i, 4).Value = Now
            
            Set outlookMail = Nothing
        End If
    Next i
    
    Set outlookApp = Nothing
    
    MsgBox "Emails sent successfully!", vbInformation
End Sub
```

### 10.4 Create Charts Programmatically

```vba
Sub CreateSalesChart()
    Dim ws As Worksheet
    Dim chartObj As ChartObject
    Dim chart As chart
    Dim lastRow As Long
    
    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    ' Delete existing charts
    Dim c As ChartObject
    For Each c In ws.ChartObjects
        c.Delete
    Next c
    
    ' Create new chart
    Set chartObj = ws.ChartObjects.Add(Left:=300, Top:=10, Width:=500, Height:=300)
    Set chart = chartObj.chart
    
    With chart
        .ChartType = xlColumnClustered
        
        ' Set data source
        .SetSourceData Source:=ws.Range("A1:B" & lastRow)
        
        ' Format chart
        .HasTitle = True
        .ChartTitle.Text = "Monthly Sales Report"
        .ChartTitle.Font.Size = 14
        
        ' Format axes
        .Axes(xlValue).HasTitle = True
        .Axes(xlValue).AxisTitle.Text = "Revenue ($)"
        
        ' Add data labels
        .SeriesCollection(1).HasDataLabels = True
        .SeriesCollection(1).DataLabels.NumberFormat = "$#,##0"
        
        ' Apply style
        .ChartStyle = 42
        
        ' Legend
        .HasLegend = False
    End With
    
    MsgBox "Chart created!", vbInformation
End Sub
```

### 10.5 Import/Export Data

```vba
' Import CSV file
Sub ImportCSV()
    Dim filePath As String
    filePath = Application.GetOpenFilename("CSV Files (*.csv), *.csv", , "Select CSV File")
    
    If filePath = "False" Then Exit Sub
    
    ' Method 1: Open as workbook
    Dim csvWb As Workbook
    Set csvWb = Workbooks.Open(filePath)
    csvWb.Sheets(1).UsedRange.Copy ThisWorkbook.Sheets("Import").Range("A1")
    csvWb.Close False
    
    MsgBox "CSV imported!", vbInformation
End Sub

' Export range to CSV
Sub ExportToCSV()
    Dim filePath As String
    filePath = ThisWorkbook.Path & "\Export_" & Format(Now, "yyyymmdd") & ".csv"
    
    ' Copy to new workbook and save as CSV
    Dim exportWb As Workbook
    Set exportWb = Workbooks.Add
    ThisWorkbook.Sheets("Export").UsedRange.Copy exportWb.Sheets(1).Range("A1")
    
    Application.DisplayAlerts = False
    exportWb.SaveAs filePath, xlCSV
    exportWb.Close False
    Application.DisplayAlerts = True
    
    MsgBox "Exported to: " & filePath, vbInformation
End Sub

' Export to PDF
Sub ExportToPDF()
    Dim filePath As String
    filePath = ThisWorkbook.Path & "\Report_" & Format(Now, "yyyymmdd") & ".pdf"
    
    ThisWorkbook.Sheets("Report").ExportAsFixedFormat _
        Type:=xlTypePDF, _
        fileName:=filePath, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False
    
    MsgBox "PDF exported to: " & filePath, vbInformation
End Sub
```

---

## 11. Best Practices for VBA Code

### Code Structure

```vba
Option Explicit  ' ALWAYS use this

' Module-level declarations
Private Const MAX_ROWS As Long = 100000
Private Const REPORT_DATE As String = "2026-09-10"

' Main procedure - entry point
Sub MainProcess()
    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    ' Call helper procedures
    Call ValidateData
    Call ProcessData
    Call GenerateReport
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    
    MsgBox "Process complete!", vbInformation
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    MsgBox "Error: " & Err.Description, vbCritical
End Sub
```

### Performance Optimization

```vba
Sub OptimizedCode()
    ' 1. Turn off screen updating
    Application.ScreenUpdating = False
    
    ' 2. Turn off automatic calculation
    Application.Calculation = xlCalculationManual
    
    ' 3. Disable events
    Application.EnableEvents = False
    
    ' 4. Use arrays instead of cell-by-cell
    Dim data() As Variant
    data = Range("A1:D10000").Value
    ' Process data in array...
    Range("A1:D10000").Value = data
    
    ' 5. Avoid Select/Activate
    ' BAD:  Range("A1").Select: Selection.Value = 10
    ' GOOD: Range("A1").Value = 10
    
    ' 6. Use With blocks
    With Range("A1").Font
        .Bold = True
        .Size = 12
    End With
    
    ' 7. Use specific objects, not Variant
    Dim ws As Worksheet  ' NOT: Dim ws As Object
    
    ' 8. Disable alerts for batch operations
    Application.DisplayAlerts = False
    ' ... delete sheets, etc.
    Application.DisplayAlerts = True
    
    ' Restore settings
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
End Sub
```

### Naming Conventions

```vba
' Variables: camelCase
Dim lastRow As Long
Dim customerName As String
Dim isValid As Boolean

' Constants: UPPER_CASE
Const MAX_RETRIES As Long = 3
Const TAX_RATE As Double = 0.08

' Procedures: PascalCase
Sub ProcessSalesData()
Function CalculateCommission()

' Controls: prefix + name
' txtName, cboDepartment, btnSubmit, lblTitle

' Meaningful names (avoid: x, tmp, data1, val)
' GOOD: employeeName, totalRevenue, isFirstRun
' BAD:  x, temp, d, v
```

### Documentation

```vba
' ============================================================
' Procedure: ProcessSalesData
' Purpose:   Processes raw sales data and calculates totals
' Author:    John Smith
' Date:      2026-09-10
' Parameters:
'   ws - Source worksheet containing sales data
'   startDate - Filter start date
'   endDate - Filter end date
' Returns:   Number of records processed
' ============================================================
Function ProcessSalesData(ws As Worksheet, startDate As Date, endDate As Date) As Long
    ' ... implementation
End Function
```

### Complete Checklist

- [ ] `Option Explicit` at top of every module
- [ ] All variables declared with specific types
- [ ] Error handling in every procedure
- [ ] ScreenUpdating/Calculation/Events managed
- [ ] No unnecessary Select/Activate
- [ ] Meaningful variable and procedure names
- [ ] Comments explaining "why", not "what"
- [ ] Arrays for large data operations
- [ ] Constants for magic numbers
- [ ] Modular code (small, focused procedures)

---

## 12. Practice Exercises

### Exercise 1: Hello World

Create a macro that displays a message box with your name.

```vba
Sub HelloWorld()
    Dim name As String
    name = InputBox("What is your name?", "Hello")
    If name <> "" Then
        MsgBox "Hello, " & name & "! Welcome to VBA!", vbInformation, "Greeting"
    End If
End Sub
```

### Exercise 2: Range Operations

Create a macro that fills cells A1:A10 with numbers 1-10, then highlights even numbers in green.

```vba
Sub HighlightEvenNumbers()
    Dim i As Long
    
    For i = 1 To 10
        Cells(i, 1).Value = i
        If i Mod 2 = 0 Then
            Cells(i, 1).Interior.Color = RGB(0, 200, 0)
        Else
            Cells(i, 1).Interior.ColorIndex = xlNone
        End If
    Next i
    
    MsgBox "Done! Even numbers highlighted.", vbInformation
End Sub
```

### Exercise 3: Data Validation Loop

Loop through column A and mark rows where sales exceed $1000.

```vba
Sub MarkHighSales()
    Dim lastRow As Long
    Dim i As Long
    
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To lastRow  ' Skip header
        If IsNumeric(Cells(i, 2).Value) Then
            If Cells(i, 2).Value > 1000 Then
                Cells(i, 3).Value = "High"
                Cells(i, 3).Font.Color = RGB(0, 128, 0)
            Else
                Cells(i, 3).Value = "Normal"
            End If
        End If
    Next i
End Sub
```

### Exercise 4: Sheet Consolidator

Combine data from multiple sheets into one summary sheet.

```vba
Sub ConsolidateSheets()
    Dim summaryWs As Worksheet
    Dim ws As Worksheet
    Dim summaryRow As Long
    Dim lastRow As Long
    
    ' Create or clear summary sheet
    On Error Resume Next
    Set summaryWs = ThisWorkbook.Sheets("Summary")
    On Error GoTo 0
    
    If summaryWs Is Nothing Then
        Set summaryWs = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Sheets(1))
        summaryWs.Name = "Summary"
    End If
    
    summaryWs.Cells.Clear
    summaryWs.Range("A1:D1").Value = Array("Sheet", "Row", "Column A", "Column B")
    summaryRow = 2
    
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> "Summary" Then
            lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
            If lastRow > 1 Then
                ws.Range("A2:B" & lastRow).Copy summaryWs.Cells(summaryRow, 3)
                summaryWs.Range(summaryWs.Cells(summaryRow, 1), _
                    summaryWs.Cells(summaryRow + lastRow - 2, 1)).Value = ws.Name
                summaryRow = summaryRow + lastRow - 1
            End If
        End If
    Next ws
    
    summaryWs.Columns.AutoFit
    MsgBox "Consolidated " & (summaryRow - 2) & " rows!", vbInformation
End Sub
```

### Exercise 5: Random Password Generator (Function)

```vba
Function GeneratePassword(length As Long) As String
    Dim chars As String
    Dim password As String
    Dim i As Long
    
    chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789!@#$%"
    password = ""
    
    Randomize  ' Seed random number generator
    
    For i = 1 To length
        password = password & Mid(chars, Int(Rnd() * Len(chars)) + 1, 1)
    Next i
    
    GeneratePassword = password
End Function

' Usage: =GeneratePassword(12) in a cell, or call from another Sub
Sub TestPassword()
    MsgBox "Generated password: " & GeneratePassword(16)
End Sub
```

### Exercise 6: Interactive Form

Build a simple calculator using a UserForm with two text boxes, an operator dropdown, and a result display.

```vba
' UserForm: frmCalculator
' Controls: txtNum1, txtNum2, cboOperator, btnCalculate, lblResult

Private Sub UserForm_Initialize()
    cboOperator.AddItem "+"
    cboOperator.AddItem "-"
    cboOperator.AddItem "*"
    cboOperator.AddItem "/"
    cboOperator.ListIndex = 0
End Sub

Private Sub btnCalculate_Click()
    If Not IsNumeric(txtNum1.Value) Or Not IsNumeric(txtNum2.Value) Then
        MsgBox "Please enter valid numbers.", vbExclamation
        Exit Sub
    End If
    
    Dim num1 As Double, num2 As Double, result As Double
    num1 = CDbl(txtNum1.Value)
    num2 = CDbl(txtNum2.Value)
    
    Select Case cboOperator.Value
        Case "+": result = num1 + num2
        Case "-": result = num1 - num2
        Case "*": result = num1 * num2
        Case "/"
            If num2 = 0 Then
                MsgBox "Cannot divide by zero!", vbExclamation
                Exit Sub
            End If
            result = num1 / num2
    End Select
    
    lblResult.Caption = "Result: " & result
End Sub
```

---

## 13. Video References

### Wise Owl Tutorials (YouTube)
- Excel VBA Introduction: https://www.youtube.com/watch?v=6t8M1bB6r7I
- VBA for Beginners Full Course: https://www.youtube.com/watch?v=GhGg4JkxKEM
- Excel VBA UserForms: https://www.youtube.com/watch?v=KHOJ5_pb5yc

### ExcelMacroMastery (YouTube)
- VBA Tutorial for Beginners: https://www.youtube.com/watch?v=y5Y1Jd_STCE
- VBA Arrays Guide: https://www.youtube.com/watch?v=hyGc9I6xJXg

### Chandoo.org (YouTube)
- Excel VBA Crash Course: https://www.youtube.com/watch?v=kS3fwbJ1cOA
- VBA Loops Tutorial: https://www.youtube.com/watch?v=Aw2bJjj8c5E

### TrumpExcel (YouTube)
- Excel VBA Tutorial for Beginners: https://www.youtube.com/watch?v=Wwg_KbxWfyk
- Recording Macros: https://www.youtube.com/watch?v=JkDKRg_jGgM
- VBA UserForm Tutorial: https://www.youtube.com/watch?v=9WN-2NknhzA

### Leila Gharani (YouTube)
- Excel VBA Tutorial: https://www.youtube.com/watch?v=BaKdFe1JY4I

### Kevin Stratvert (YouTube)
- Excel VBA for Beginners: https://www.youtube.com/watch?v=BaKdFe1JY4I

---

## 14. Sources & Further Reading

### Websites

- **TrumpExcel** - Excel VBA Programming: https://trumpexcel.com/excel-vba/
- **ExcelMacroMastery** - VBA Articles: https://excelmacromastery.com/vba-articles/
- **Chandoo.org** - VBA Tutorials: https://chandoo.org/?s=vba
- **Microsoft Docs** - VBA Language Reference: https://docs.microsoft.com/en-us/office/vba/api/overview/excel
- **Excel Easy** - VBA Tutorial: https://www.excel-easy.com/vba.html
- **Automate Excel** - VBA Tutorials: https://www.automateexcel.com/vba/

### Books

- *Excel VBA Programming for Dummies* by Michael Alexander & John Walkenbach
- *Power Programming with VBA* by John Walkenbach
- *Excel 2019 Power Programming with VBA* by Michael Alexander & Dick Kusleika
- *Professional Excel Development* by Rob Bovey, Dennis Wallentin, Stephen Bullen

### Free Resources

- ExcelMacroMastery VBA Cheat Sheet: https://excelmacromastery.com/
- Chandoo.org VBA Classes: https://chandoo.org/?s=vba
- TrumpExcel Free VBA Course: https://trumpexcel.com/excel-vba/

---

> **Next Module:** [Module 8: Real-World Projects](../08-Real-World-Projects/08-real-world-projects.md) - Apply everything you've learned in practical, industry-relevant projects.


## FILE: 08-Real-World-Projects/08-real-world-projects.md

# Module 8: Real-World Excel Projects

> **Level:** Beginner → Advanced (applied)
> **Duration:** 40–60 hours (8 projects × 5–7 hours each)
> **Prerequisites:** Modules 1–7 (all core Excel skills)

---

## Table of Contents

1. [Project 1: Personal Budget Tracker](#project-1-personal-budget-tracker)
2. [Project 2: Sales Dashboard](#project-2-sales-dashboard)
3. [Project 3: Inventory Management System](#project-3-inventory-management-system)
4. [Project 4: Employee Schedule & Leave Tracker](#project-4-employee-schedule--leave-tracker)
5. [Project 5: Invoice Generator Template](#project-5-invoice-generator-template)
6. [Project 6: Data Cleaning Pipeline](#project-6-data-cleaning-pipeline)
7. [Project 7: Financial Model](#project-7-financial-model)
8. [Project 8: Automated Report Generator](#project-8-automated-report-generator)
9. [Industry Use Cases](#industry-use-cases)
10. [Excel Certification Paths](#excel-certification-paths)
11. [Excel vs Google Sheets](#excel-vs-google-sheets)
12. [Recommended Templates & Add-ins](#recommended-templates--add-ins)
13. [YouTube Video References](#youtube-video-references)
14. [Sources](#sources)

---

# Project 1: Personal Budget Tracker

## Overview

Build a comprehensive personal finance tracker that records income and expenses by category, calculates running balances, and presents spending insights through a visual dashboard. This is the ideal first project because it uses foundational skills in a context everyone understands.

**Difficulty:** ⭐⭐ Beginner
**Time:** 5–6 hours
**Skills Used:** SUM, SUMIF, SUMIFS, charts (pie, bar, line), conditional formatting, data validation, named ranges, basic formatting

## Step-by-Step Instructions

### Step 1 - Set Up the Workbook Structure

Create four sheets:

| Sheet Name | Purpose |
|---|---|
| `Transactions` | Raw data entry for every income/expense |
| `Categories` | Reference list of income & expense categories |
| `Monthly Summary` | Aggregated totals by month and category |
| `Dashboard` | Charts and KPIs |

### Step 2 - Build the Categories Sheet

In the `Categories` sheet, create two tables:

**Income Categories (Column A):**

| Category |
|---|
| Salary |
| Freelance |
| Investments |
| Gifts |
| Other Income |

**Expense Categories (Column C):**

| Category |
|---|
| Housing |
| Utilities |
| Groceries |
| Transportation |
| Dining Out |
| Entertainment |
| Healthcare |
| Insurance |
| Clothing |
| Education |
| Savings |
| Debt Payments |
| Miscellaneous |

**Named Ranges:**
- Select Income Categories → Formulas → Define Name → `IncomeCategories`
- Select Expense Categories → Define Name → `ExpenseCategories`

### Step 3 - Build the Transactions Sheet

Set up headers in Row 1:

| A | B | C | D | E | F | G |
|---|---|---|---|---|---|---|
| **Date** | **Type** | **Category** | **Description** | **Amount** | **Account** | **Running Balance** |

**Data Validation (dropdowns):**

1. **Type column (B):** Select B2:B5000 → Data → Data Validation → List → Source: `Income,Expense`
2. **Category column (C):** Use a dependent dropdown:
   ```excel
   =IF(B2="Income", IncomeCategories, ExpenseCategories)
   ```
   (Data Validation → List → Source: the formula above)
3. **Account column (F):** Create a named range `Accounts` with: `Checking,Savings,Credit Card,Cash`

**Running Balance Formula (G2):**
```excel
=IF(B2="Income", G1+E2, G1-E2)
```
Drag down. Set G1 to your starting balance.

### Step 4 - Build the Monthly Summary Sheet

Create a summary table:

| | A | B | C | D | E |
|---|---|---|---|---|---|
| 1 | **Month** | **Total Income** | **Total Expenses** | **Net Savings** | **Savings Rate** |
| 2 | Jan 2025 | | | | |

**Formulas (assuming Transactions data in rows 2–5000):**

```excel
B2 (Total Income):
=SUMIFS(Transactions!E:E, Transactions!B:B, "Income",
        Transactions!A:A, ">="&DATE(2025,1,1),
        Transactions!A:A, "<"&DATE(2025,2,1))

C2 (Total Expenses):
=SUMIFS(Transactions!E:E, Transactions!B:B, "Expense",
        Transactions!A:A, ">="&DATE(2025,1,1),
        Transactions!A:A, "<"&DATE(2025,2,1))

D2 (Net Savings): =B2-C2

E2 (Savings Rate): =IF(B2>0, D2/B2, 0)  → Format as percentage
```

For category breakdowns, add a second table:

| Category | Budget | Actual | Difference | % of Budget |
|---|---|---|---|---|

```excel
Actual (for "Groceries" in cell A10):
=SUMIFS(Transactions!E:E, Transactions!C:C, A10,
        Transactions!B:B, "Expense",
        Transactions!A:A, ">="&DATE(2025,1,1),
        Transactions!A:A, "<"&DATE(2025,2,1))

Difference: =B10-C10
% of Budget: =IF(B10>0, C10/B10, 0)
```

### Step 5 - Apply Conditional Formatting

**Budget vs Actual (Difference column):**

1. Select the Difference column
2. Home → Conditional Formatting → New Rule → Format cells that contain
3. Cell Value < 0 → Red fill (over budget)
4. Cell Value > 0 → Green fill (under budget)
5. Cell Value = 0 → Yellow fill (on target)

**Spending Alerts:**
Select the % of Budget column → Conditional Formatting → Color Scale:
- Green (0%) → Yellow (80%) → Red (100%+)

### Step 6 - Build the Dashboard

**Chart 1 - Monthly Income vs Expenses (Clustered Bar)**
1. Select Month, Total Income, Total Expenses columns from Monthly Summary
2. Insert → Chart → Clustered Bar
3. Title: "Monthly Income vs Expenses"

**Chart 2 - Expense Breakdown (Pie/Donut)**
1. Select Category and Actual from the category breakdown table
2. Insert → Chart → Doughnut
3. Title: "Spending by Category"
4. Add data labels with percentages

**Chart 3 - Savings Trend (Line)**
1. Select Month and Net Savings
2. Insert → Chart → Line with Markers
3. Title: "Monthly Savings Trend"
4. Add a trendline (Linear)

**Chart 4 - Budget Gauge (for single month)**
Use a stacked bar showing spent vs remaining:
```
=BudgetUsed  |  =BudgetRemaining
```

**KPI Cards (top of dashboard):**
Use merged cells with large font:
```excel
Total Income:    =SUMIFS(Transactions!E:E, Transactions!B:B, "Income")
Total Expenses:  =SUMIFS(Transactions!E:E, Transactions!B:B, "Expense")
Net Savings:     =TotalIncome-TotalExpenses
Savings Rate:    =NetSavings/TotalIncome
```

## Sample Data

| Date | Type | Category | Description | Amount | Account |
|---|---|---|---|---|---|
| 2025-01-01 | Income | Salary | January paycheck | 4500.00 | Checking |
| 2025-01-03 | Expense | Housing | Rent payment | 1200.00 | Checking |
| 2025-01-05 | Expense | Groceries | Weekly groceries | 85.50 | Credit Card |
| 2025-01-06 | Expense | Transportation | Gas | 45.00 | Credit Card |
| 2025-01-08 | Expense | Dining Out | Restaurant dinner | 62.30 | Credit Card |
| 2025-01-10 | Income | Freelance | Web design project | 750.00 | Savings |
| 2025-01-12 | Expense | Utilities | Electric bill | 95.00 | Checking |
| 2025-01-15 | Expense | Insurance | Car insurance | 150.00 | Checking |
| 2025-01-18 | Expense | Entertainment | Streaming services | 25.99 | Credit Card |
| 2025-01-20 | Expense | Healthcare | Doctor copay | 30.00 | Cash |
| 2025-01-22 | Expense | Groceries | Weekly groceries | 92.15 | Credit Card |
| 2025-01-25 | Income | Investments | Dividend payment | 125.00 | Savings |
| 2025-01-28 | Expense | Clothing | Winter jacket | 89.99 | Credit Card |
| 2025-01-30 | Expense | Savings | Transfer to savings | 500.00 | Checking |

## Expected Output

- A `Transactions` sheet with 50+ entries over 3+ months
- A `Monthly Summary` showing income, expenses, net savings per month with % savings rate
- Category breakdown with budget vs actual comparison
- A `Dashboard` sheet with 4 charts and KPI cards
- Conditional formatting highlighting over-budget categories in red
- Dropdown menus for Type, Category, and Account

## Tips & Tricks

- **Use Excel Tables** (Ctrl+T) for the Transactions range - formulas auto-expand
- **Freeze Row 1** on the Transactions sheet for scrolling
- Use `TEXT(A2, "MMM YYYY")` to group transactions by month in pivot-style summaries
- Add a **"Quick Add"** section at the top of Transactions for rapid data entry
- Use `SUBTOTAL(9, range)` instead of `SUM` if you filter transactions

## Common Pitfalls

- ❌ **Forgetting to lock the date range** in SUMIFS (use absolute references for the summary sheet)
- ❌ **Circular references** in running balance if row 1 formula references itself
- ❌ **Not using consistent category names** - always use the dropdown, never type manually
- ❌ **Including the header row** in SUMIFS ranges
- ❌ **Date format mismatches** - ensure all dates are actual Excel dates, not text

---

# Project 2: Sales Dashboard

## Overview

Build an interactive sales analytics dashboard using PivotTables, PivotCharts, slicers, and timelines. Analyze sales data by region, product, salesperson, and time period. Create KPI cards and dynamic charts that respond to user selections.

**Difficulty:** ⭐⭐⭐ Intermediate
**Time:** 6–7 hours
**Skills Used:** PivotTables, PivotCharts, slicers, timeline, GETPIVOTDATA, INDEX/MATCH, conditional formatting, dashboard layout, sparklines

## Step-by-Step Instructions

### Step 1 - Prepare the Sales Data

Create a `SalesData` sheet with the following columns:

| Column | Header | Format |
|---|---|---|
| A | OrderID | Text (e.g., ORD-0001) |
| B | OrderDate | Date |
| C | Region | Text |
| D | Salesperson | Text |
| E | Product | Text |
| F | Category | Text |
| G | Quantity | Number |
| H | UnitPrice | Currency |
| I | Revenue | Currency |
| J | Cost | Currency |
| K | Profit | Currency |
| L | ProfitMargin | Percentage |

**Key Formulas:**
```excel
I2: =G2*H2          (Revenue)
K2: =I2-J2          (Profit)
L2: =IF(I2>0, K2/I2, 0)  (Profit Margin)
```

Convert to Excel Table (Ctrl+T) and name it `SalesTable`.

### Step 2 - Create the PivotTable

1. Insert → PivotTable → From Table/Range
2. Place on a new sheet called `PivotAnalysis`
3. **Configure the PivotTable:**
   - **Rows:** Region, then Salesperson (nested)
   - **Columns:** Product Category
   - **Values:** Sum of Revenue, Sum of Profit, Count of OrderID

4. **Right-click the PivotTable** → PivotTable Options:
   - Check "Show items with no data"
   - Set "For empty cells show: 0"

### Step 3 - Add Calculated Fields

In the PivotTable:
1. PivotTable Analyze → Fields, Items & Sets → Calculated Field
2. Add `AvgOrderValue` = `Revenue / Quantity`
3. Add `ProfitMargin` = `Profit / Revenue`

### Step 4 - Create PivotCharts

**Chart 1 - Revenue by Region (Clustered Bar)**
1. Click inside PivotTable → Insert → PivotChart → Clustered Bar
2. Move chart to `Dashboard` sheet

**Chart 2 - Monthly Revenue Trend (Line)**
1. Create a second PivotTable with OrderDate grouped by Month
2. Values: Sum of Revenue, Sum of Profit
3. Insert Line Chart with markers

**Chart 3 - Product Category Mix (Pie)**
1. Create a PivotChart → Pie Chart
2. Show percentage data labels

**Chart 4 - Salesperson Performance (Horizontal Bar)**
1. Rows: Salesperson, Values: Sum of Revenue
2. Sort descending by revenue

### Step 5 - Add Slicers and Timeline

**Slicers:**
1. Click PivotTable → PivotTable Analyze → Insert Slicer
2. Add slicers for: Region, Category, Salesperson
3. **Connect to all PivotTables:** Right-click each slicer → Report Connections → check all PivotTables
4. Arrange slicers at the top of the Dashboard

**Timeline:**
1. PivotTable Analyze → Insert Timeline → OrderDate
2. Set to show by Quarter (dropdown in timeline)
3. Connect to all PivotTables

### Step 6 - Build KPI Cards

Above the charts, create KPI cards using GETPIVOTDATA:

```excel
Total Revenue:
=GETPIVOTDATA("Revenue", PivotAnalysis!$A$3)

Total Orders:
=GETPIVOTDATA("OrderID", PivotAnalysis!$A$3, "OrderID", "")

Avg Order Value:
=TotalRevenue/TotalOrders

Top Region:
=INDEX(RegionList, MATCH(MAX(RegionRevenue), RegionRevenue, 0))
```

Format KPIs with:
- Large bold numbers
- Icon indicators (▲/▼) for comparison to previous period
- Color coding (green for up, red for down)

### Step 7 - Dashboard Layout and Formatting

1. Set the `Dashboard` sheet to a fixed view: View → Page Layout
2. Arrange elements in a grid:
   - **Row 1-2:** Title + KPI cards
   - **Row 3:** Slicers and Timeline
   - **Row 4-5:** Main charts
   - **Row 6:** Detail tables
3. Use consistent colors: pick a 4-color palette
4. Remove chart clutter: no gridlines, minimal legends
5. Add a company logo placeholder

## Sample Data (first 15 rows)

| OrderID | OrderDate | Region | Salesperson | Product | Category | Qty | UnitPrice | Cost |
|---|---|---|---|---|---|---|---|---|
| ORD-0001 | 2024-01-05 | North | Alice Chen | Widget Pro | Hardware | 25 | 49.99 | 625.00 |
| ORD-0002 | 2024-01-08 | South | Bob Martinez | Service Plan | Services | 10 | 199.00 | 800.00 |
| ORD-0003 | 2024-01-12 | East | Carol Davis | Widget Lite | Hardware | 50 | 29.99 | 900.00 |
| ORD-0004 | 2024-01-15 | West | David Kim | Premium Suite | Software | 5 | 499.00 | 1250.00 |
| ORD-0005 | 2024-01-20 | North | Alice Chen | Service Plan | Services | 15 | 199.00 | 1200.00 |
| ORD-0006 | 2024-02-03 | South | Eve Johnson | Widget Pro | Hardware | 30 | 49.99 | 900.00 |
| ORD-0007 | 2024-02-10 | East | Frank Wilson | Basic Suite | Software | 20 | 99.00 | 1000.00 |
| ORD-0008 | 2024-02-14 | West | Grace Lee | Widget Lite | Hardware | 100 | 29.99 | 1800.00 |
| ORD-0009 | 2024-03-01 | North | Henry Patel | Premium Suite | Software | 8 | 499.00 | 2000.00 |
| ORD-0010 | 2024-03-05 | South | Alice Chen | Service Plan | Services | 12 | 199.00 | 960.00 |

## Expected Output

- Raw data table with 200+ rows across 4 quarters
- 3–4 PivotTables analyzing different dimensions
- 4 PivotCharts on the Dashboard
- 3 slicers + 1 timeline, all interconnected
- KPI cards showing Total Revenue, Total Orders, Avg Order Value, Top Region
- Charts update dynamically when slicers change
- Professional color scheme and layout

## Tips & Tricks

- **Alt + F1** creates an instant chart from selected data
- Use **PivotTable → Repeat All Item Labels** to avoid blank cells
- **Slicer shortcuts:** Ctrl+click to multi-select, Alt+S to clear filter
- Use **sparklines** (Insert → Sparklines) for inline mini-charts in tables
- Name your PivotTable ranges for cleaner GETPIVOTDATA formulas

## Common Pitfalls

- ❌ **Slicers not connected** to all PivotTables (must do manually)
- ❌ **Dates not recognized** - ensure column is formatted as Date, not Text
- ❌ **PivotTable doesn't refresh** when source data changes (right-click → Refresh)
- ❌ **Overcrowding the dashboard** - leave white space between elements
- ❌ **GETPIVOTDATA errors** when PivotTable field names change

---

# Project 3: Inventory Management System

## Overview

Build a product inventory database with automated stock tracking, reorder alerts, and in/out transaction logging. Use lookup functions to pull product details, conditional formatting to flag low stock, and basic VBA to automate repetitive tasks.

**Difficulty:** ⭐⭐⭐ Intermediate
**Time:** 6–7 hours
**Skills Used:** XLOOKUP, INDEX/MATCH, data validation, conditional formatting, VBA macros, COUNTIF, SUMIFS, running totals

## Step-by-Step Instructions

### Step 1 - Create the Product Database Sheet

Headers for `Products` sheet:

| A | B | C | D | E | F | G | H | I |
|---|---|---|---|---|---|---|---|---|
| SKU | Product Name | Category | Supplier | Unit Cost | Sell Price | Current Stock | Reorder Level | Status |

**Status Formula (I2):**
```excel
=IF(G2<=0, "OUT OF STOCK",
  IF(G2<=H2, "REORDER NOW",
    IF(G2<=H2*1.5, "LOW STOCK", "IN STOCK")))
```

**Named Ranges:**
- `SKUList` - Column A of Products
- `ProductTable` - entire Products table

### Step 2 - Add Reorder Conditional Formatting

Select the Status column (I2:I500):

1. **OUT OF STOCK:** Red fill, white bold text
   - Formula: `=$I2="OUT OF STOCK"`
2. **REORDER NOW:** Orange fill, dark text
   - Formula: `=$I2="REORDER NOW"`
3. **LOW STOCK:** Yellow fill, dark text
   - Formula: `=$I2="LOW STOCK"`
4. **IN STOCK:** Green fill, dark text
   - Formula: `=$I2="IN STOCK"`

**Additional: Highlight entire row when stock < reorder level:**
Select A2:I500 → New Rule → Use formula:
```excel
=$G2<$H2
```
Format: Light red fill

### Step 3 - Create the Transaction Log Sheet

Headers for `Transactions` sheet:

| A | B | C | D | E | F | G | H |
|---|---|---|---|---|---|---|---|
| TransID | Date | SKU | Type | Quantity | Unit Cost | Total | Notes |

**Type column (D):** Data Validation list: `Stock In,Stock Out,Adjustment,Return`

**Auto-fill Product Info using XLOOKUP:**
```excel
E2 (Product Name lookup - in a helper column or use XLOOKUP):
=XLOOKUP(C2, Products!A:A, Products!B:B, "Unknown SKU")

F2 (Unit Cost):
=XLOOKUP(C2, Products!A:A, Products!E:E, 0)

G2 (Total):
=IF(D2="Stock In", E2*F2, -E2*F2)
```

### Step 4 - Build the Stock Summary with Running Totals

Create a `StockSummary` sheet using SUMIFS:

```excel
For SKU "WGT-001":

Total Stock In:
=SUMIFS(Transactions!E:E, Transactions!C:C, "WGT-001",
        Transactions!D:D, "Stock In") +
 SUMIFS(Transactions!E:E, Transactions!C:C, "WGT-001",
        Transactions!D:D, "Return")

Total Stock Out:
=SUMIFS(Transactions!E:E, Transactions!C:C, "WGT-001",
        Transactions!D:D, "Stock Out") +
 SUMIFS(Transactions!E:E, Transactions!C:C, "WGT-001",
        Transactions!D:D, "Adjustment")*−1

Current Stock: =TotalIn - TotalOut
```

### Step 5 - Build the Dashboard

**Chart 1 - Stock by Category (Bar)**
Group products by category, show total current stock.

**Chart 2 - Top 10 Products by Value (Horizontal Bar)**
```excel
Stock Value: =CurrentStock * UnitCost
```

**Chart 3 - Transactions Over Time (Line)**
Sum of stock in and stock out by month.

**KPI Cards:**
- Total SKUs: `=COUNTA(Products!A:A)-1`
- Items to Reorder: `=COUNTIF(Products!I:I, "REORDER NOW")+COUNTIF(Products!I:I, "OUT OF STOCK")`
- Total Inventory Value: `=SUMPRODUCT(Products!G:G, Products!E:E)`

### Step 6 - Basic VBA: One-Click Stock Update

Press **Alt+F11** → Insert → Module:

```vba
Sub UpdateStockLevels()
    Dim wsProd As Worksheet, wsTrans As Worksheet
    Dim lastRow As Long, i As Long
    Dim sku As String, qty As Long, transType As String

    Set wsProd = Sheets("Products")
    Set wsTrans = Sheets("Transactions")

    ' Reset current stock to 0
    lastRow = wsProd.Cells(wsProd.Rows.Count, "A").End(xlUp).Row
    For i = 2 To lastRow
        wsProd.Cells(i, "G").Value = 0
    Next i

    ' Recalculate from transactions
    lastRow = wsTrans.Cells(wsTrans.Rows.Count, "A").End(xlUp).Row
    For i = 2 To lastRow
        sku = wsTrans.Cells(i, "C").Value
        qty = wsTrans.Cells(i, "E").Value
        transType = wsTrans.Cells(i, "D").Value

        ' Find the product row
        Dim prodRow As Range
        Set prodRow = wsProd.Columns("A").Find(sku, LookAt:=xlWhole)

        If Not prodRow Is Nothing Then
            Select Case transType
                Case "Stock In", "Return"
                    wsProd.Cells(prodRow.Row, "G").Value = _
                        wsProd.Cells(prodRow.Row, "G").Value + qty
                Case "Stock Out"
                    wsProd.Cells(prodRow.Row, "G").Value = _
                        wsProd.Cells(prodRow.Row, "G").Value - qty
                Case "Adjustment"
                    wsProd.Cells(prodRow.Row, "G").Value = _
                        wsProd.Cells(prodRow.Row, "G").Value + qty
            End Select
        End If
    Next i

    MsgBox "Stock levels updated!", vbInformation
End Sub
```

## Sample Data

**Products (first 10):**

| SKU | Product Name | Category | Supplier | Unit Cost | Sell Price | Stock | Reorder Lvl |
|---|---|---|---|---|---|---|---|
| WGT-001 | Widget Pro | Hardware | Acme Corp | 15.00 | 49.99 | 150 | 50 |
| WGT-002 | Widget Lite | Hardware | Acme Corp | 8.00 | 29.99 | 30 | 100 |
| SFT-001 | Basic Suite | Software | TechSoft | 40.00 | 99.00 | 200 | 50 |
| SFT-002 | Premium Suite | Software | TechSoft | 150.00 | 499.00 | 15 | 20 |
| SRV-001 | Service Plan | Services | InHouse | 50.00 | 199.00 | 500 | 100 |
| CAB-001 | USB Cable 6ft | Cables | WireCo | 2.50 | 12.99 | 8 | 200 |
| CAB-002 | HDMI Cable | Cables | WireCo | 5.00 | 24.99 | 180 | 100 |
| ACC-001 | Mouse Pad | Accessories | PadCo | 3.00 | 9.99 | 0 | 50 |
| ACC-002 | Webcam HD | Accessories | CamTech | 25.00 | 79.99 | 45 | 30 |
| PKG-001 | Starter Pack | Bundles | InHouse | 30.00 | 89.99 | 75 | 25 |

## Expected Output

- Product database with 50+ SKUs and automatic status indicators
- Transaction log with 100+ entries showing stock movements
- Conditional formatting: red rows for out-of-stock, orange for reorder
- Stock summary with running totals matching physical counts
- Dashboard with inventory value, reorder alerts, category breakdown
- VBA macro that recalculates all stock levels in one click

## Tips & Tricks

- Use **Data Validation → Error Alert** to prevent negative stock entries
- **Ctrl+Shift+L** to toggle filters on the product list
- Create a **"Quick Stock In"** form using VBA UserForm for faster data entry
- Use `SUMPRODUCT` for weighted average cost calculations
- Protect the Products sheet to prevent accidental formula edits

## Common Pitfalls

- ❌ **XLOOKUP returns #N/A** for new SKUs - always use the `if_not_found` argument
- ❌ **Stock going negative** - add a check: `=MAX(0, calculated_stock)`
- ❌ **Transaction types inconsistent** - always use the dropdown, never free-text
- ❌ **Not backing up before running VBA** - macros can't be undone (Ctrl+Z)
- ❌ **Forgetting to refresh** stock levels after manual transaction edits

---

# Project 4: Employee Schedule & Leave Tracker

## Overview

Create a workforce management tool that handles weekly schedules, tracks leave balances, and integrates company holidays. Use date functions to calculate working days, conditional formatting to visualize availability, and formulas to manage accruals.

**Difficulty:** ⭐⭐⭐ Intermediate
**Time:** 5–6 hours
**Skills Used:** NETWORKDAYS, WORKDAY, EOMONTH, conditional formatting, data validation, date functions, COUNTIF, DATEDIF

## Step-by-Step Instructions

### Step 1 - Set Up the Employee Database

Create an `Employees` sheet:

| A | B | C | D | E | F | G |
|---|---|---|---|---|---|---|
| EmpID | Name | Department | HireDate | Annual Leave | Sick Leave | Personal Days |
| EMP001 | John Smith | Engineering | 2020-03-15 | 15 | 10 | 3 |
| EMP002 | Maria Garcia | Marketing | 2021-07-01 | 15 | 10 | 3 |

### Step 2 - Create the Holiday Calendar

Create a `Holidays` sheet with a named range `HolidayList`:

| Date | Holiday Name |
|---|---|
| 2025-01-01 | New Year's Day |
| 2025-01-20 | MLK Day |
| 2025-02-17 | Presidents Day |
| 2025-05-26 | Memorial Day |
| 2025-07-04 | Independence Day |
| 2025-09-01 | Labor Day |
| 2025-11-27 | Thanksgiving |
| 2025-11-28 | Day After Thanksgiving |
| 2025-12-25 | Christmas Day |

### Step 3 - Build the Weekly Schedule Grid

Create a `Schedule` sheet:

**Header Row:** Employee Name | Mon 1/6 | Tue 1/7 | Wed 1/8 | ... | Fri 1/10

**Data Validation for shift cells:** List → `Day Shift,Night Shift,Off,Leave - Annual,Leave - Sick,Leave - Personal,Training,Remote`

**Conditional Formatting for the grid:**
- `Day Shift` → Light blue fill
- `Night Shift` → Dark blue fill, white text
- `Off` → Gray fill
- `Leave - Annual` → Green fill
- `Leave - Sick` → Orange fill
- `Leave - Personal` → Purple fill
- Holiday dates → Red fill with holiday name

### Step 4 - Build the Leave Tracker

Create a `LeaveTracker` sheet:

| A | B | C | D | E | F | G | H | I |
|---|---|---|---|---|---|---|---|---|
| EmpID | Name | Leave Type | Start Date | End Date | Working Days | Status | Approved By | Notes |

**Working Days Formula (F2):**
```excel
=NETWORKDAYS(D2, E2, HolidayList)
```
This automatically excludes weekends and company holidays.

**Leave Balance Summary (separate table):**

| EmpID | Name | Annual Entitled | Annual Used | Annual Remaining | Sick Entitled | Sick Used | Sick Remaining |
|---|---|---|---|---|---|---|---|

```excel
Annual Used:
=SUMIFS(LeaveTracker!F:F, LeaveTracker!A:A, A2,
        LeaveTracker!C:C, "Leave - Annual",
        LeaveTracker!G:G, "Approved")

Annual Remaining: =C2-E2
```

### Step 5 - Create the Monthly Calendar View

Build a visual calendar for each month:

1. Create a grid: Columns = days 1–31, Rows = employees
2. Use WORKDAY and EOMONTH to auto-generate dates:
   ```excel
   =WORKDAY(EOMONTH(DATE(2025,1,1),-1), 1)  → First workday of Jan
   ```
3. Use INDEX/MATCH to pull schedule data for each employee/day
4. Apply conditional formatting based on shift type

### Step 6 - Dashboard: Attendance Overview

**Charts:**
1. **Leave Usage by Type (Pie):** Annual vs Sick vs Personal usage
2. **Monthly Absences (Bar):** Count of leave days per month
3. **Department Attendance (Stacked Bar):** Present vs Absent by dept
4. **Leave Balance Summary (Horizontal Bar):** Remaining days per employee

**KPIs:**
- Total employees on leave today: `=COUNTIF(today_column, "*Leave*")`
- Average attendance rate: `=(total_workdays - total_absences) / total_workdays`
- Upcoming holidays: Next holiday from HolidayList

## Sample Data

**Leave Requests:**

| EmpID | Name | Leave Type | Start Date | End Date | Working Days | Status |
|---|---|---|---|---|---|---|
| EMP001 | John Smith | Leave - Annual | 2025-03-10 | 2025-03-14 | 5 | Approved |
| EMP002 | Maria Garcia | Leave - Sick | 2025-02-03 | 2025-02-04 | 2 | Approved |
| EMP003 | James Wilson | Leave - Annual | 2025-04-21 | 2025-04-25 | 5 | Pending |
| EMP001 | John Smith | Leave - Personal | 2025-05-02 | 2025-05-02 | 1 | Approved |
| EMP004 | Sarah Lee | Leave - Annual | 2025-12-22 | 2025-12-31 | 6 | Approved |

## Expected Output

- Employee database with leave entitlements
- Holiday calendar with named range
- Weekly schedule grid with color-coded shifts
- Leave tracker calculating working days automatically
- Leave balance summary per employee showing used and remaining
- Monthly calendar view with conditional formatting
- Dashboard with attendance KPIs and charts

## Tips & Tricks

- **NETWORKDAYS(start, end, holidays)** is your best friend for leave calculations
- Use `TODAY()` in formulas to highlight the current day on the schedule
- Create a **printable weekly schedule** with Page Layout view
- Use **Data Validation Input Message** to explain what each shift code means
- Protect cells with formulas so only schedule cells are editable

## Common Pitfalls

- ❌ **Holiday list not updated** annually - set a reminder for January
- ❌ **NETWORKDAYS including the holiday list range but missing new holidays**
- ❌ **Leave spanning month boundaries** - ensure SUMIFS covers the full year
- ❌ **Time zones not considered** for remote workers in different zones
- ❌ **Not distinguishing between "Pending" and "Approved"** leave in balance calculations

---

# Project 5: Invoice Generator Template

## Overview

Build a professional, print-ready invoice template that auto-calculates subtotals, taxes, and totals. Include a customer database, product catalog with lookup, and professional formatting suitable for client-facing documents.

**Difficulty:** ⭐⭐ Beginner-Intermediate
**Time:** 4–5 hours
**Skills Used:** Named ranges, VLOOKUP/XLOOKUP, templates, print setup, cell formatting, IF, ROUND, page layout

## Step-by-Step Instructions

### Step 1 - Create Supporting Sheets

**Sheet 1: `CompanyInfo`**

| Field | Value |
|---|---|
| Company Name | Your Company LLC |
| Address | 123 Business Ave, Suite 100 |
| City, State ZIP | New York, NY 10001 |
| Phone | (555) 123-4567 |
| Email | billing@yourcompany.com |
| Website | www.yourcompany.com |
| Tax ID | XX-XXXXXXX |
| Bank Name | First National Bank |
| Account Number | XXXX-XXXX-1234 |
| Routing Number | XXXXXXXXX |

**Sheet 2: `CustomerDB`**

| CustID | Company Name | Contact | Address | City | State | ZIP | Email | Phone |
|---|---|---|---|---|---|---|---|---|
| C001 | Acme Corp | John Doe | 456 Client St | Boston | MA | 02101 | john@acme.com | (555) 234-5678 |

**Sheet 3: `ProductCatalog`**

| ProdID | Description | Unit Price | Taxable |
|---|---|---|---|
| P001 | Widget Pro | 49.99 | Yes |
| P002 | Consulting (hourly) | 150.00 | Yes |
| P003 | Shipping | 15.00 | No |

**Named Ranges:**
- `CustomerTable` - CustomerDB data
- `ProductTable` - ProductCatalog data
- `CompanyName`, `CompanyAddress`, etc. - from CompanyInfo

### Step 2 - Design the Invoice Layout

Create the `Invoice` sheet with the following layout:

**Rows 1-5: Header**
```
[Company Logo Placeholder]     YOUR COMPANY LLC
                               123 Business Ave, Suite 100
                               New York, NY 10001
                               (555) 123-4567

                               INVOICE
```

**Rows 7-12: Invoice Details**
```
Invoice #:     [INV-2025-001]          Bill To:
Date:          [Auto: =TODAY()]         [Customer Name from lookup]
Due Date:      [Auto: +30 days]        [Address]
Payment Terms: Net 30                  [City, State ZIP]
                                       [Email]
```

**Customer Lookup:**
```excel
Customer Name: =XLOOKUP(InvoiceCustomerID, CustomerDB!A:A, CustomerDB!B:B)
Address:       =XLOOKUP(InvoiceCustomerID, CustomerDB!A:A, CustomerDB!D:D)
```

**Rows 14-25: Line Items Table**

| # | Description | Qty | Unit Price | Line Total |
|---|---|---|---|---|
| 1 | [Product lookup] | 10 | [Auto from catalog] | [Qty × Price] |
| 2 | | | | |
| 3 | | | | |
| ... | | | | |

**Formulas:**
```excel
Description: =XLOOKUP(B14, ProductCatalog!A:A, ProductCatalog!B:B, "")
Unit Price:  =XLOOKUP(B14, ProductCatalog!A:A, ProductCatalog!C:C, 0)
Line Total:  =C14*D14
```

**Rows 27-31: Totals**
```
                            Subtotal:     =SUM(E14:E25)
                            Tax (8.875%): =Subtotal * TaxRate
                            Shipping:     [manual entry]
                            TOTAL DUE:    =Subtotal + Tax + Shipping
```

**Rows 33-37: Footer**
```
Notes: [Optional text]
Payment Instructions: Please make check payable to Your Company LLC
Bank: First National Bank | Acct: XXXX-XXXX-1234 | Routing: XXXXXXXXX
Thank you for your business!
```

### Step 3 - Professional Formatting

- **Company name:** 18pt bold, dark blue
- **"INVOICE" title:** 24pt bold, centered
- **Line items header:** Bold, dark blue fill, white text
- **Alternating row colors:** Light gray on even rows
- **TOTAL DUE:** Bold, 14pt, top border, light blue fill
- **Currency format:** `$#,##0.00` for all money fields
- **Date format:** `MMMM D, YYYY` (e.g., "January 15, 2025")

### Step 4 - Print Setup

1. Page Layout → Size: Letter
2. Margins: Narrow (0.5" all sides)
3. Orientation: Portrait
4. Print Area: Select the invoice area → Set Print Area
5. Header/Footer: Add page number ("Page &P of &N")
6. Scale: Fit to 1 page wide by 1 page tall
7. Gridlines: Uncheck (for clean print)
8. **Test print** to PDF before using

### Step 5 - Make It Reusable

1. **Save as template:** File → Save As → Excel Template (.xltx)
2. **Invoice number auto-increment:** Use a helper cell:
   ```excel
   ="INV-" & TEXT(YEAR(TODAY()), "0000") & "-" & TEXT(NextInvoiceNum, "000")
   ```
3. **Due date auto-calc:** `=InvoiceDate + PaymentTermsDays`

## Sample Data

**Product Catalog:**

| ProdID | Description | Unit Price | Taxable |
|---|---|---|---|
| P001 | Widget Pro | 49.99 | Yes |
| P002 | Widget Lite | 29.99 | Yes |
| P003 | Consulting - Standard | 150.00 | Yes |
| P004 | Consulting - Premium | 250.00 | Yes |
| P005 | Service Plan - Annual | 1,999.00 | Yes |
| P006 | Setup Fee | 500.00 | Yes |
| P007 | Shipping - Standard | 15.00 | No |
| P008 | Shipping - Express | 35.00 | No |
| P009 | Training (per session) | 300.00 | Yes |
| P010 | Custom Development (hourly) | 175.00 | Yes |

## Expected Output

- Professional-looking invoice template
- Auto-populating company and customer info
- Product catalog lookup for line items
- Automatic subtotal, tax, and total calculations
- Print-ready layout that fits on one page
- Saved as .xltx for reuse
- Invoice numbering system

## Tips & Tricks

- Use **conditional formatting** to highlight the TOTAL row
- Add a **"Paid" stamp** as a WordArt or text box overlay
- Create a **duplicate sheet** for credit notes (negative quantities)
- Use `ROUND(tax_calc, 2)` to avoid fractional cent errors
- Save multiple templates for different tax rates or currencies

## Common Pitfalls

- ❌ **Tax calculated on non-taxable items** - use SUMPRODUCT with a taxable flag
- ❌ **Print area not set** - prints blank pages or cuts off content
- ❌ **Currency formatting inconsistent** - apply to entire columns at once
- ❌ **Invoice number not unique** - use a separate counter sheet
- ❌ **Forgetting to update the year** in invoice numbering

---

# Project 6: Data Cleaning Pipeline

## Overview

Build a repeatable data cleaning workflow using Power Query and text functions. Import messy CSV data, fix encoding issues, remove duplicates, standardize formats, and output a clean dataset ready for analysis. This project teaches one of the most valuable Excel skills: transforming dirty data.

**Difficulty:** ⭐⭐⭐ Intermediate
**Time:** 5–6 hours
**Skills Used:** Power Query, TRIM, CLEAN, SUBSTITUTE, PROPER, UPPER, LOWER, TEXT, LEFT, RIGHT, MID, Find & Replace, Remove Duplicates, Text to Columns

## Step-by-Step Instructions

### Step 1 - Import Messy Data

**Source:** Download or create a messy CSV with common data issues:

1. Data → Get Data → From File → From Text/CSV
2. Select the file → Preview dialog appears
3. Check encoding (UTF-8 usually correct) → Load

### Step 2 - Identify Data Quality Issues

Open Power Query Editor (Data → Get Data → Launch Power Query Editor):

**Common issues to look for:**
- Leading/trailing spaces
- Mixed case ("new york", "NEW YORK", "New York")
- Duplicates (exact and fuzzy)
- Missing values (blank, N/A, "null", "-")
- Inconsistent date formats (MM/DD/YYYY vs DD/MM/YYYY)
- Phone numbers in different formats
- Addresses with abbreviations
- Special characters and encoding artifacts

### Step 3 - Power Query Transformations

**In Power Query Editor, apply these steps:**

**3a. Remove extra spaces:**
```
Transform → Format → Trim
```

**3b. Standardize text case:**
```
For names:   Transform → Format → Capitalize Each Word
For emails:  Transform → Format → Lowercase
For codes:   Transform → Format → Uppercase
```

**3c. Remove special characters:**
```
Transform → Replace Values
Value to Find: specific characters
Replace With: (empty)
```

**3d. Handle missing values:**
```
Transform → Replace Values
Value: null → Replace with: (blank or "N/A")
Value: "-" → Replace with: (blank)
Value: "N/A" → Replace with: (blank)
```

**3e. Remove duplicates:**
```
Home → Remove Rows → Remove Duplicates
Select key columns (e.g., Email, FullName)
```

**3f. Standardize dates:**
```
Transform → Data Type → Date
If mixed formats, use: Date.From([DateString], "MM/dd/yyyy")
```

**3g. Parse and clean phone numbers:**
```
Add Column → Custom Column:
= Text.Select([Phone], {"0".."9"})
Then format: = "(" & Text.Range(CleanPhone,0,3) & ") " & ...
```

**3h. Split/merge columns:**
```
FullName → Split Column by Delimiter (space) → First, Last
Address → Split Column by Delimiter (comma) → Street, City, State ZIP
```

### Step 4 - Advanced: Fuzzy Matching for Deduplication

Power Query supports fuzzy matching:

```
Home → Remove Rows → Remove Duplicates
→ Right-click column → Remove Duplicates
→ Options: Similarity threshold = 0.8
```

Or use a Merge query with fuzzy matching:
```
Home → Merge Queries
Select join columns → Check "Use fuzzy matching"
Set similarity threshold
```

### Step 5 - Text Function Alternatives (Without Power Query)

If not using Power Query, use these formulas on a helper column:

```excel
Remove spaces:        =TRIM(A2)
Remove non-printable: =CLEAN(A2)
Capitalize names:     =PROPER(A2)
Remove commas:        =SUBSTITUTE(A2, ",", "")
Extract first name:   =LEFT(A2, FIND(" ", A2)-1)
Extract last name:    =RIGHT(A2, LEN(A2)-FIND(" ", A2))
Standardize phone:    ="("&MID(A2,1,3)&") "&MID(A2,4,3)&"-"&MID(A2,7,4)
Clean email:          =LOWER(TRIM(SUBSTITUTE(A2, " ", "")))
```

### Step 6 - Output Clean Dataset

1. In Power Query: Home → Close & Load → To new worksheet
2. Rename the output sheet: `CleanData`
3. **Verify:** Spot-check 10-20 rows against the original
4. **Document changes:** Create a `CleaningLog` sheet listing all transformations applied

### Step 7 - Build a Reusable Cleaning Template

1. Save the Power Query steps as a reusable query
2. Create parameters for file path so you can change the source
3. Add a "Refresh" button that re-runs the entire pipeline
4. Document the expected input format

## Sample Data (Messy)

| FullName | Email | Phone | City | State | ZipCode | DateJoined |
|---|---|---|---|---|---|---|
| john smith | John.Smith@acme.example | (555)123-4567 | new york | NY | 10001 | 01/15/2023 |
| JANE DOE | jane.doe@acme.example | 555.234.5678 | Los Angeles | ca | 90001 | 2023-03-20 |
| bob  Johnson | BOB@company.org | 5553456789 | chicago | Illinois | 60601 | 03/15/2023 |
| John Smith | john.smith@acme.example | (555) 123-4567 | New York | NY | 10001 | 01/15/2023 |
| Alice Brown | alice.brown@ | 555.456.7890 |  | TX | 75001 | - |
| null | - | N/A | Miami | FL | 33101 | 2023-06-01 |
| María García | maria@email.com | +1-555-567-8901 | San Antonio | TX | 78201 | 07/04/2023 |
| Charlie Wilson | CHARLIE.W@WORK.COM | (555)678-9012 | seattle | WA | 98101 | 08/30/2023 |

**Issues present:**
- Mixed case names, emails
- Extra spaces (bob  Johnson)
- Duplicate person (john smith / John Smith)
- Different phone formats
- Missing values (null, -, N/A, blank)
- Different date formats
- Invalid email (alice.brown@)
- Inconsistent state (ca vs CA vs California)

## Expected Output

- Clean dataset with all issues resolved
- Consistent capitalization (Proper Case for names, lowercase for emails)
- Standardized phone format: (XXX) XXX-XXXX
- All dates in MM/DD/YYYY format
- Duplicates removed (2 → 1)
- Missing values handled consistently
- A cleaning log documenting all transformations
- Reusable Power Query pipeline

## Tips & Tricks

- **Always keep the original data** - clean into a new sheet/table
- Use `TRIM(CLEAN(SUBSTITUTE(A2,CHAR(160)," ")))` for stubborn spaces (non-breaking)
- **Flash Fill** (Ctrl+E) can auto-detect patterns for name splitting
- Power Query's **Applied Steps** pane is your undo history - review each step
- Use `=IFERROR(VALUE(A2), A2)` to convert text-numbers to numbers

## Common Pitfalls

- ❌ **Deleting rows** instead of cleaning them - you lose data
- ❌ **Not checking for fuzzy duplicates** - "Jon Smith" vs "John Smith"
- ❌ **Changing data types before cleaning** - clean first, then set types
- ❌ **Power Query case sensitivity** - "New York" ≠ "new york" in some operations
- ❌ **Not refreshing the query** after source data changes

---

# Project 7: Financial Model

## Overview

Build a 5-year financial projection model with revenue forecasting, expense modeling, break-even analysis, loan amortization, and sensitivity analysis. This is the most formula-intensive project and is directly applicable to business planning and investment analysis.

**Difficulty:** ⭐⭐⭐⭐ Advanced
**Time:** 7–8 hours
**Skills Used:** NPV, IRR, PMT, PPMT, IPMT, FV, PV, Data Tables (1-var, 2-var), Scenario Manager, Goal Seek, INDEX/MATCH, conditional formatting

## Step-by-Step Instructions

### Step 1 - Model Structure

Create five sheets:

| Sheet | Purpose |
|---|---|
| `Assumptions` | All input variables in one place |
| `Revenue` | Revenue projections |
| `Expenses` | Cost and expense model |
| `Financials` | Income statement, cash flow |
| `Analysis` | NPV, IRR, break-even, sensitivity |

### Step 2 - Build the Assumptions Sheet

**Revenue Assumptions:**

| Parameter | Value | Notes |
|---|---|---|
| Starting Customers | 100 | Year 1 beginning |
| Monthly Growth Rate | 5% | New customer acquisition |
| Churn Rate | 2% | Monthly customer loss |
| Avg Revenue Per User (ARPU) | $50/month | |
| ARPU Growth | 3%/year | Annual price increase |

**Cost Assumptions:**

| Parameter | Value |
|---|---|
| Fixed Costs (monthly) | $10,000 |
| Variable Cost per Customer | $15 |
| Cost Inflation | 2%/year |
| Initial Investment | $250,000 |
| Loan Amount | $200,000 |
| Loan Interest Rate | 6% annually |
| Loan Term (years) | 5 |
| Tax Rate | 25% |

**Name each cell** for easy reference (e.g., `GrowthRate`, `ChurnRate`, `ARPU`)

### Step 3 - Revenue Projections

**Monthly Customer Model:**
```excel
Month 1 Customers: =StartingCustomers
Month 2 Customers: =Month1 * (1 + GrowthRate - ChurnRate)
Month N Customers: =Month(N-1) * (1 + GrowthRate - ChurnRate)
```

**Monthly Revenue:**
```excel
=Customers * ARPU
```

**Annual Revenue:** Sum of monthly revenues for the year.

### Step 4 - Expense Model

**Monthly Expenses:**
```excel
Fixed:   =FixedCosts * (1 + CostInflation) ^ (YearNum - 1)
Variable: =Customers * VariableCostPerCustomer
Total:   =Fixed + Variable
```

### Step 5 - Income Statement

| | Year 1 | Year 2 | Year 3 | Year 4 | Year 5 |
|---|---|---|---|---|---|
| Revenue | | | | | |
| COGS (Variable) | | | | | |
| Gross Profit | | | | | |
| Operating Expenses (Fixed) | | | | | |
| EBITDA | | | | | |
| Depreciation | | | | | |
| EBIT | | | | | |
| Interest Expense | | | | | |
| EBT | | | | | |
| Taxes | | | | | |
| Net Income | | | | | |

**Key Formulas:**
```excel
Gross Profit:    =Revenue - COGS
EBITDA:          =Gross Profit - Operating Expenses
EBIT:            =EBITDA - Depreciation
Interest:        (from amortization schedule)
EBT:             =EBIT - Interest
Taxes:           =MAX(0, EBT * TaxRate)
Net Income:      =EBT - Taxes
```

### Step 6 - Loan Amortization Schedule

| Month | Beginning Balance | Payment | Principal | Interest | Ending Balance |
|---|---|---|---|---|---|

```excel
Monthly Payment (cell B1): =PMT(Rate/12, Term*12, -LoanAmount)
Beginning Balance (Month 1): =LoanAmount
Interest: =BeginningBalance * (Rate/12)
Principal: =Payment - Interest
Ending Balance: =BeginningBalance - Principal
```

**Built-in functions for specific periods:**
```excel
Principal in month 13: =PPMT(Rate/12, 13, Term*12, -LoanAmount)
Interest in month 13:  =IPMT(Rate/12, 13, Term*12, -LoanAmount)
```

### Step 7 - NPV and IRR Analysis

**Net Present Value:**
```excel
Cash Flows: Year 0 = -InitialInvestment, Years 1-5 = Net Income + Depreciation
=NPV(DiscountRate, Year1:Year5) + Year0
```

**Internal Rate of Return:**
```excel
=IRR(Year0:Year5)
```

**Payback Period:** Use cumulative cash flow to find when it turns positive.

### Step 8 - Break-Even Analysis

**Break-Even Point (units):**
```excel
=FixedCosts / (ARPU - VariableCostPerCustomer)
```

**Break-Even with Goal Seek:**
1. Data → What-If Analysis → Goal Seek
2. Set cell: Net Income (Year 1)
3. To value: 0
4. By changing: Growth Rate
5. Click OK → Excel finds the minimum growth rate for profitability

### Step 9 - Sensitivity Analysis with Data Tables

**1-Variable Data Table (Revenue sensitivity to growth rate):**

Set up:
| Growth Rate | Revenue |
|---|---|
| 1% | =RevenueFormula |
| 2% | |
| 3% | |
| ... | |

1. Select the table range (including the formula cell)
2. Data → What-If Analysis → Data Table
3. Column input cell: GrowthRate cell
4. Excel fills in revenue for each growth rate

**2-Variable Data Table (Revenue sensitivity to growth rate AND ARPU):**

| | $30 ARPU | $40 ARPU | $50 ARPU | $60 ARPU |
|---|---|---|---|---|
| 2% growth | =RevenueFormula | | | |
| 4% growth | | | | |
| 6% growth | | | | |
| 8% growth | | | | |

1. Select the table range
2. Data → What-If Analysis → Data Table
3. Row input cell: ARPU cell
4. Column input cell: GrowthRate cell

**Conditional Formatting on the 2-var table:**
- Color scale: Red (low) → Green (high)
- Highlight profitable combinations in green

### Step 10 - Scenario Manager

1. Data → What-If Analysis → Scenario Manager
2. Create three scenarios:

| Scenario | Growth Rate | ARPU | Variable Cost |
|---|---|---|---|
| Best Case | 8% | $60 | $12 |
| Base Case | 5% | $50 | $15 |
| Worst Case | 2% | $40 | $20 |

3. Show each scenario to see the impact on Net Income and NPV
4. Create a Scenario Summary report

## Sample Data

**Assumptions for Base Case:**

| Input | Value |
|---|---|
| Starting Customers | 100 |
| Monthly Growth Rate | 5% |
| Monthly Churn Rate | 2% |
| ARPU | $50/month |
| Fixed Costs | $10,000/month |
| Variable Cost/Customer | $15/month |
| Initial Investment | $250,000 |
| Loan | $200,000 @ 6% for 5 years |
| Discount Rate | 10% |
| Tax Rate | 25% |

## Expected Output

- 5-year revenue projection with monthly granularity
- Complete income statement (Revenue → Net Income)
- Loan amortization schedule (60 months)
- NPV and IRR calculations
- Break-even point calculation
- 1-variable and 2-variable sensitivity data tables
- 3 scenarios (Best/Base/Worst) with summary
- Dashboard showing key financial metrics over 5 years

## Tips & Tricks

- **Keep all assumptions on one sheet** - never hardcode numbers in formulas
- Use **named ranges** for every assumption cell
- Color code: Blue font = input, Black font = formula (financial modeling convention)
- **Ctrl+`** (backtick) toggles formula view - useful for auditing
- Use `IFERROR` to handle division by zero in margin calculations

## Common Pitfalls

- ❌ **Circular references** in interest calculations (interest depends on balance, which depends on payments)
- ❌ **Not using absolute references** for assumptions in data tables
- ❌ **Forgetting to add back depreciation** in cash flow (it's non-cash)
- ❌ **Data tables recalculating slowly** - use Ctrl+Alt+F9 to force recalc
- ❌ **Inconsistent time periods** - don't mix monthly and annual figures

---

# Project 8: Automated Report Generator

## Overview

Build a VBA-powered macro that automatically refreshes data, generates PivotTables and charts, exports the report as a PDF, and emails it to a distribution list. This is the capstone project that ties together everything learned.

**Difficulty:** ⭐⭐⭐⭐⭐ Advanced
**Time:** 7–8 hours
**Skills Used:** VBA macros, PivotTables, charts, Outlook automation, PDF export, error handling, user forms, progress indicators

## Step-by-Step Instructions

### Step 1 - Design the Report Template

Create a `Report` sheet with:
- Header section: Title, date, prepared by
- KPI summary section
- PivotTable section
- Chart section
- Footer: Page numbers, confidentiality notice

### Step 2 - Create the Data Connection

1. Data → Get Data → From your data source (SQL, CSV, SharePoint, etc.)
2. Load to PivotTable cache
3. Name the query: `ReportData`

### Step 3 - Build the VBA Macro

Press **Alt+F11** → Insert → Module:

```vba
Option Explicit

Sub GenerateReport()
    ' ============================================
    ' Automated Report Generator
    ' Refreshes data, updates charts, exports PDF,
    ' and emails the report
    ' ============================================

    On Error GoTo ErrorHandler

    Dim wsReport As Worksheet
    Dim pvt As PivotTable
    Dim reportDate As String
    Dim pdfPath As String
    Dim startTime As Double

    startTime = Timer
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' --- Step 1: Update status ---
    UpdateStatus "Starting report generation..."

    ' --- Step 2: Refresh all data connections ---
    UpdateStatus "Refreshing data..."
    ThisWorkbook.RefreshAll
    Application.CalculateUntilAsyncQueriesDone

    ' --- Step 3: Update report date ---
    Set wsReport = Sheets("Report")
    wsReport.Range("ReportDate").Value = Format(Date, "MMMM D, YYYY")

    ' --- Step 4: Refresh PivotTables ---
    UpdateStatus "Updating PivotTables..."
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        For Each pvt In ws.PivotTables
            pvt.RefreshTable
            pvt.Update
        Next pvt
    Next ws

    ' --- Step 5: Export to PDF ---
    UpdateStatus "Exporting PDF..."
    reportDate = Format(Date, "YYYY-MM-DD")
    pdfPath = ThisWorkbook.Path & "\Reports\Report_" & reportDate & ".pdf"

    ' Create Reports folder if it doesn't exist
    If Dir(ThisWorkbook.Path & "\Reports", vbDirectory) = "" Then
        MkDir ThisWorkbook.Path & "\Reports"
    End If

    wsReport.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=pdfPath, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=False

    ' --- Step 6: Send email ---
    UpdateStatus "Sending email..."
    Call SendReportEmail(pdfPath)

    ' --- Step 7: Complete ---
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    UpdateStatus "Report generated in " & _
        Format(Timer - startTime, "0.0") & " seconds"
    MsgBox "Report generated and emailed successfully!" & vbNewLine & _
           "Saved to: " & pdfPath, vbInformation, "Complete"
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "Error " & Err.Number & ": " & Err.Description, _
           vbCritical, "Report Generation Failed"
End Sub

Sub SendReportEmail(pdfPath As String)
    ' Requires Microsoft Outlook
    Dim outlookApp As Object
    Dim mailItem As Object
    Dim recipientList As String

    ' Get recipient list from Recipients sheet
    recipientList = ""
    Dim wsRecip As Worksheet
    Set wsRecip = Sheets("Recipients")
    Dim i As Long
    For i = 2 To wsRecip.Cells(wsRecip.Rows.Count, "A").End(xlUp).Row
        If wsRecip.Cells(i, "B").Value = "Active" Then
            recipientList = recipientList & wsRecip.Cells(i, "A").Value & ";"
        End If
    Next i

    ' Create email
    Set outlookApp = CreateObject("Outlook.Application")
    Set mailItem = outlookApp.CreateItem(0)

    With mailItem
        .To = Left(recipientList, Len(recipientList) - 1)
        .Subject = "Weekly Report - " & Format(Date, "MMMM D, YYYY")
        .Body = "Hello," & vbNewLine & vbNewLine & _
                "Please find attached this week's report." & vbNewLine & _
                vbNewLine & "Best regards," & vbNewLine & _
                "Automated Report System"
        .Attachments.Add pdfPath
        .Display  ' Use .Send to auto-send
    End With

    Set mailItem = Nothing
    Set outlookApp = Nothing
End Sub

Sub UpdateStatus(statusMsg As Application.StatusBar)
    Application.StatusBar = statusMsg
    DoEvents
End Sub
```

### Step 4 - Create the Recipients Sheet

| Email | Status | Department |
|---|---|---|
| manager@company.com | Active | Management |
| team-lead@company.com | Active | Operations |
| analyst@company.com | Active | Finance |
| former-employee@company.com | Inactive | Former |

### Step 5 - Add a Run Button

1. Insert → Shapes → Rounded Rectangle
2. Right-click → Assign Macro → `GenerateReport`
3. Format the button with color and text "📊 Generate Report"
4. Place it at the top of the Report sheet

### Step 6 - Error Handling and Logging

Add a `Log` sheet and logging subroutine:

```vba
Sub LogAction(action As String)
    Dim wsLog As Worksheet
    Set wsLog = Sheets("Log")
    Dim nextRow As Long
    nextRow = wsLog.Cells(wsLog.Rows.Count, "A").End(xlUp).Row + 1

    wsLog.Cells(nextRow, 1).Value = Now()
    wsLog.Cells(nextRow, 2).Value = Application.UserName
    wsLog.Cells(nextRow, 3).Value = action
End Sub
```

### Step 7 - Schedule Automatic Execution (Optional)

To run the report on a schedule, use Windows Task Scheduler:

1. Create a .bat file:
   ```bat
   "C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE" /r "C:\Reports\ReportGenerator.xlsm"
   ```
2. Open Task Scheduler → Create Basic Task
3. Set trigger (e.g., every Monday at 8 AM)
4. Set action to run the .bat file

## Sample Data

**Recipients (already shown above)**

**Report Content (auto-generated from source data):**
- Sales summary by region (PivotTable)
- Monthly trend chart
- Top 10 products bar chart
- KPI cards: Revenue, Orders, Avg Order Value
- Comparison to previous period

## Expected Output

- One-click report generation button
- Auto-refreshed data and PivotTables
- Professional PDF report with charts and KPIs
- Auto-email to distribution list
- Generation log with timestamps
- Error handling for common issues (no data, Outlook not running, etc.)
- Execution time displayed in status bar

## Tips & Tricks

- **Always test with `.Display` first**, then switch to `.Send` when confident
- Use `Application.StatusBar` for progress updates during long operations
- **Save as .xlsm** (macro-enabled) - regular .xlsx won't save VBA code
- Add `Application.CutCopyMode = False` after paste operations to clear clipboard
- Use `DoEvents` to keep Excel responsive during long operations

## Common Pitfalls

- ❌ **Forgetting to save as .xlsm** - macros disappear when saved as .xlsx
- ❌ **Outlook security prompts** - user may need to allow programmatic access
- ❌ **PDF export failing** if the report sheet has print area issues
- ❌ **Not handling "no data" scenarios** - add checks before PivotTable refresh
- ❌ **Macro security blocking execution** - ensure trusted location or signed macro

---

# Industry Use Cases

## Finance & Banking

| Use Case | Excel Features | Example |
|---|---|---|
| Financial reporting | PivotTables, charts, Power Query | Monthly P&L, balance sheets |
| Budgeting & forecasting | What-If Analysis, Data Tables, scenarios | Annual budget with variance analysis |
| Loan analysis | PMT, NPV, IRR, amortization | Mortgage calculator, loan comparison |
| Risk assessment | Monte Carlo (add-in), probability distributions | Portfolio risk modeling |
| Audit trails | Track Changes, comments, cell protection | Financial audit workpapers |

## Human Resources

| Use Case | Excel Features | Example |
|---|---|---|
| Payroll processing | VLOOKUP, tax tables, SUMPRODUCT | Monthly payroll calculation |
| Employee scheduling | NETWORKDAYS, conditional formatting | Shift planning, leave management |
| Recruitment tracking | Data validation, filters, COUNTIF | Candidate pipeline management |
| Performance reviews | Rating scales, radar charts | Employee evaluation scorecards |
| Headcount planning | PivotTables, trend analysis | Workforce planning dashboard |

## Marketing

| Use Case | Excel Features | Example |
|---|---|---|
| Campaign analysis | PivotTables, charts, A/B test calculations | Email campaign performance |
| Customer segmentation | PivotTables, conditional formatting | RFM analysis (Recency, Frequency, Monetary) |
| Content calendar | Date functions, conditional formatting | Monthly content planning |
| ROI tracking | SUMIFS, percentage calculations | Marketing spend vs revenue attribution |
| Social media reporting | Power Query (API import), charts | Weekly social media metrics |

## Operations & Supply Chain

| Use Case | Excel Features | Example |
|---|---|---|
| Inventory management | XLOOKUP, conditional formatting, VBA | Stock tracking, reorder alerts |
| Production planning | Solver add-in, linear programming | Optimal production mix |
| Quality control | Statistical functions, control charts | Defect tracking, SPC charts |
| Vendor management | Data validation, scoring models | Supplier evaluation matrix |
| Logistics optimization | Distance matrices, route planning | Delivery route cost analysis |

## Education

| Use Case | Excel Features | Example |
|---|---|---|
| Grade book | AVERAGE, weighted grades, conditional formatting | Student grade tracking |
| Attendance tracking | COUNTIF, NETWORKDAYS | Daily attendance with % calculations |
| Test analysis | Statistical functions, histograms | Item analysis, score distributions |
| Budget management | SUMIFS, charts | Department budget tracking |
| Scheduling | Date functions, Gantt charts | Course schedule, room allocation |

## Healthcare

| Use Case | Excel Features | Example |
|---|---|---|
| Patient tracking | Data validation, conditional formatting | Patient intake, appointment scheduling |
| Clinical trial data | Statistical analysis, data cleaning | Trial results analysis |
| Inventory management | XLOOKUP, alerts | Medical supply tracking |
| Staff scheduling | NETWORKDAYS, shift patterns | Nurse scheduling, on-call rotation |
| Financial reporting | PivotTables, charts | Department cost analysis |

---

# Excel Certification Paths

## Microsoft Office Specialist (MOS) - Excel Associate

**Exam:** MO-200 (Excel 2019) or MO-210 (Microsoft 365 Apps)

**Skills Measured:**
- Manage worksheets and workbooks (15–20%)
- Manage data cells and ranges (20–25%)
- Manage tables and table data (15–20%)
- Perform operations by using formulas and functions (20–25%)
- Manage charts (15–20%)

**Preparation:**
- Study time: 40–60 hours
- Practice with Excel 2019/365 (not older versions)
- Use GMetrix or Certiport practice tests
- Focus areas: PivotTables, formulas, charts, cell formatting

**Career Value:** Entry-level validation, good for administrative and analyst roles

## Microsoft Office Specialist (MOS) - Excel Expert

**Exam:** MO-201 (Excel 2019) or MO-211 (Microsoft 365 Apps)

**Skills Measured:**
- Manage workbook options and settings (10–15%)
- Apply custom data formats and layouts (15–20%)
- Create advanced formulas (20–25%)
- Create advanced charts and tables (25–30%)
- Manage PivotTables and PivotCharts (15–20%)

**Key Differences from Associate:**
- Advanced formulas: INDEX/MATCH, array formulas, XLOOKUP, LAMBDA
- Power Query basics
- Advanced PivotTable features (calculated fields, groups)
- Complex chart types (combo, waterfall, funnel)

**Preparation:**
- Study time: 60–80 hours
- Requires Associate certification first (recommended, not required)
- Focus on scenario-based questions

**Career Value:** Strong differentiator for data analysts and finance professionals

## Microsoft Office Specialist (MOS) - Master

**Requirements:** Pass 4 exams:
1. Excel Expert (MO-201 or MO-211)
2. Word Expert (MO-101 or MO-111)
3. PowerPoint Associate (MO-300 or MO-310)
4. Outlook Associate (MO-400 or MO-410) OR Access Expert (MO-500)

**Career Value:** Highest Office certification, signals comprehensive Microsoft 365 proficiency. Valued for executive assistants, office managers, and training roles.

## Other Relevant Certifications

| Certification | Focus | Level |
|---|---|---|
| **CFA (Chartered Financial Analyst)** | Financial modeling in Excel | Advanced |
| **FMVA (Financial Modeling & Valuation Analyst)** | Excel financial modeling | Intermediate-Advanced |
| **Google Sheets Certification** | Google Workspace equivalent | Intermediate |
| **Tableau Desktop Specialist** | Visualization (Excel skills transfer) | Intermediate |
| **Power BI Data Analyst (PL-300)** | Power BI (builds on Excel skills) | Intermediate |

---

# Excel vs Google Sheets

## Feature Comparison

| Feature | Excel | Google Sheets | Winner |
|---|---|---|---|
| **Max rows** | 1,048,576 | 10,000,000 cells total | Google |
| **Max columns** | 16,384 | 18,278 | Google |
| **Offline access** | Full (desktop app) | Limited (requires setup) | Excel |
| **Real-time collaboration** | Good (Microsoft 365) | Excellent (native) | Google |
| **PivotTables** | Full-featured | Basic but improving | Excel |
| **Power Query** | Full ETL tool | Not available | Excel |
| **VBA/Macros** | Full VBA | Apps Script (JavaScript) | Excel |
| **PivotCharts** | Yes | No (manual charts only) | Excel |
| **Conditional Formatting** | Advanced | Good | Excel |
| **Array Formulas** | Dynamic arrays (FILTER, SORT, UNIQUE) | Similar (with ARRAYFORMULA) | Tie |
| **XLOOKUP** | Yes (native) | Yes (XLOOKUP added 2022) | Tie |
| **Data Tables (What-If)** | Yes | No | Excel |
| **Solver** | Yes (add-in) | Yes (add-on) | Tie |
| **Python integration** | Yes (Excel 365) | No (external only) | Excel |
| **Collaboration** | Good | Excellent | Google |
| **Cost** | $6.99–$22/user/month (M365) | Free (personal) / $6/user/month (Workspace) | Google |
| **File size** | Can handle large files | Performance degrades >5MB | Excel |
| **Add-in ecosystem** | Massive | Growing but smaller | Excel |
| **Learning curve** | Steeper (more features) | Gentler (simpler UI) | Google |
| **Version history** | Good (OneDrive/SharePoint) | Excellent (native) | Google |
| **Scripting** | VBA (proprietary) | Apps Script (open JS) | Depends |
| **Import from web** | Power Query (very powerful) | IMPORTHTML, IMPORTXML (basic) | Excel |

## When to Use Excel

- **Financial modeling** - What-If Analysis, Data Tables, Scenario Manager
- **Large datasets** - Power Query, Power Pivot, millions of rows
- **Complex automation** - VBA is more mature and capable than Apps Script
- **Enterprise environments** - Microsoft 365 integration, compliance features
- **Advanced analytics** - Power Pivot, DAX, Python integration
- **Offline-heavy work** - Full desktop application with no internet dependency
- **Professional reports** - Better charting, formatting, and print controls

## When to Use Google Sheets

- **Team collaboration** - Real-time editing is seamless and intuitive
- **Simple data tracking** - Inventory, to-do lists, basic budgets
- **Web-based workflows** - No installation, accessible from any device
- **Cost-sensitive** - Free for personal use
- **Integration with Google ecosystem** - Forms, BigQuery, Apps Script
- **Quick sharing** - Link sharing without file attachments
- **Education** - Easy distribution to students, no license needed

## Migration Tips

**Excel → Google Sheets:**
1. Upload .xlsx to Google Drive → Open with Sheets
2. Most formulas transfer, but check: VBA (won't work), complex conditional formatting
3. Data validation, PivotTables transfer but may behave differently
4. Named ranges transfer

**Google Sheets → Excel:**
1. File → Download → .xlsx
2. Apps Script won't transfer - rewrite in VBA
3. QUERY/IMPORTDATA functions won't work - use Power Query
4. Check conditional formatting rules

---

# Recommended Templates & Add-ins

## Built-in Templates (File → New)

| Template | Use Case |
|---|---|
| Monthly Budget | Personal finance tracking |
| Invoice | Professional billing |
| Project Planner | Gantt chart-style project management |
| Calendar (any year) | Printable monthly/annual calendar |
| To-Do List | Task management |
| Loan Amortization | Loan payment schedule |
| Sales Report | Pre-built dashboard with sample data |
| Gantt Chart | Project timeline visualization |
| Blood Pressure Tracker | Health monitoring |
| Meal Planner | Weekly meal planning with grocery list |

## Recommended Add-ins

| Add-in | Purpose | Cost |
|---|---|---|
| **Power Query** (built-in 2016+) | Data import and transformation | Free |
| **Power Pivot** (built-in) | Data modeling, DAX, large datasets | Free |
| **Solver** (built-in) | Optimization, linear programming | Free |
| **Analysis ToolPak** (built-in) | Statistical analysis, histograms | Free |
| **XLMiner Analysis ToolPak** | Advanced statistics (Google Sheets) | Free |
| **Fuzzy Lookup** (Microsoft) | Match similar but not identical data | Free |
| **Geography/Stocks** (built-in 365) | Live data types for geography and stocks | Free (365) |
| **Power BI Publisher** | Export ranges to Power BI | Free |
| **Kutools for Excel** | 300+ productivity tools | $49 |
| **Ablebits** | Data cleaning, merge, dedup | $59.95 |
| **Think-Cell** | Professional charting (consulting firms) | $300+/yr |
| **@RISK** | Monte Carlo simulation | $$$ |
| **Crystal Ball** | Forecasting and simulation | $$$ |
| **Python in Excel** (365) | Run Python directly in cells | Included in 365 |

## Free Online Resources

| Resource | URL | Description |
|---|---|---|
| Excel Easy | excel-easy.com | Beginner tutorials with examples |
| Chandoo | chandoo.org | Advanced formulas, dashboards, VBA |
| ExcelJet | exceljet.net | Formula reference, shortcuts, tips |
| Contextures | contextures.com | PivotTables, data validation, VBA |
| Mr. Excel | mrexcel.com | Forum for troubleshooting |
| Leila Gharani (YouTube) | youtube.com/@LeilaGharani | Professional Excel tutorials |
| ExcelIsFun (YouTube) | youtube.com/@excelisfun | 3,000+ videos, all levels |

---

# YouTube Video References

## Beginner Projects
- "How to Create a Personal Budget in Excel" - Leila Gharani
- "Excel Budget Template Tutorial" - ExcelIsFun
- "Invoice Template in Excel" - Leila Gharani
- "Excel for Beginners Full Course" - freeCodeCamp

## Intermediate Projects
- "Excel Sales Dashboard Tutorial" - MyOnlineTrainingHub
- "PivotTable Tutorial for Beginners" - Leila Gharani
- "Excel Inventory Management System" - ExcelIsFun
- "Employee Schedule in Excel" - MyOnlineTrainingHub
- "Data Cleaning in Excel - Full Tutorial" - Leila Gharani

## Advanced Projects
- "Financial Modeling in Excel - Full Tutorial" - Aswath Damodaran
- "Excel VBA Tutorial for Beginners" - Leila Gharani
- "Power Query Tutorial" - ExcelIsFun
- "Excel Dashboard Course" - MyOnlineTrainingHub
- "Automate Excel with VBA" - WiseOwlTutorials

## Certification Prep
- "MOS Excel Associate Exam Prep" - Certiport
- "MOS Excel Expert Exam Tips" - Leila Gharani
- "Excel Certification Study Guide" - Simon Sez IT

---

# Sources

1. Microsoft. "Excel documentation and training." support.microsoft.com
2. Microsoft. "Microsoft Office Specialist (MOS) certification." learn.microsoft.com
3. ExcelJet. "Excel formulas and functions reference." exceljet.net
4. Chandoo. "Excel tips, tutorials, and templates." chandoo.org
5. Gharani, Leila. "Excel tutorials and courses." leilagharani.com
6. ExcelIsFun. "Excel video tutorials." youtube.com/excelisfun
7. MyOnlineTrainingHub. "Excel dashboard and analysis tutorials." myonlinetraininghub.com
8. Contextures. "Excel tutorials and sample files." contextures.com
9. Mr. Excel. "Excel forum and resources." mrexcel.com
10. Google. "Google Sheets documentation." support.google.com/docs
11. Investopedia. "Financial modeling in Excel." investopedia.com
12. Corporate Finance Institute. "Excel for financial analysis." corporatefinanceinstitute.com

---

> **Next Steps:** After completing all 8 projects, you should be comfortable using Excel for any professional task. Consider pursuing MOS certification, exploring Power BI for advanced visualization, or learning Python integration for data science workflows.

> **Practice Dataset Sources:**
> - Kaggle (kaggle.com/datasets) - free datasets for practice
> - data.gov - U.S. government open data
> - Google Dataset Search (datasetsearch.research.google.com)
> - UCI Machine Learning Repository - classic datasets


## FILE: 09-Sources-References/sources-references-index.md

# Sources & References Index

## Excel 0 to Hero - Complete Source Bibliography

All sources used in the creation of this learning package. Organized by category for NotebookLM reference.

---

## 1. Microsoft Official Documentation

| Source | URL | Used In |
|--------|-----|---------|
| Microsoft Excel Help & Learning | https://support.microsoft.com/en-us/excel | All Modules |
| Excel functions (alphabetical) | https://support.microsoft.com/en-us/office/excel-functions-alphabetical-b3944572-255d-4efb-bb96-c6d90033e188 | Module 2, 5 |
| Dynamic array formulas | https://support.microsoft.com/en-us/excel | Module 5 |
| PivotTable reports | https://support.microsoft.com/en-us/excel | Module 4 |
| Power Query overview | https://learn.microsoft.com/en-us/power-query/power-query-what-is-power-query | Module 6 |
| VBA Language Reference | https://learn.microsoft.com/en-us/office/vba/api/overview/excel | Module 7 |
| What-If Analysis | https://support.microsoft.com/en-us/excel | Module 4 |
| Conditional Formatting | https://support.microsoft.com/en-us/excel | Module 3 |
| Excel Data Types & Stocks | https://support.microsoft.com/en-us/excel | Module 1 |

---

## 2. Educational Websites

| Source | URL | Used In |
|--------|-----|---------|
| ExcelJet | https://exceljet.net/ | All Modules |
| ExcelJet - 500 Excel Formula Examples | https://exceljet.net/formulas | Module 2, 5 |
| ExcelJet - Excel Shortcuts | https://exceljet.net/keyboard-shortcuts | Module 1 |
| Excel Easy | https://www.excel-easy.com/ | Modules 1-4 |
| Excel Easy - VBA | https://www.excel-easy.com/vba.html | Module 7 |
| Corporate Finance Institute (CFI) | https://corporatefinanceinstitute.com/resources/excel/ | Modules 2-5 |
| GCFGlobal - Excel Tutorials | https://edu.gcfglobal.org/en/excel/ | Module 1 |
| Chandoo.org | https://chandoo.org/?s=excel | All Modules |
| TrumpExcel | https://trumpexcel.com/ | Modules 1-4, 8 |
| Contextures | https://www.contextures.com/ | Modules 3-4 |
| MyOnlineTrainingHub | https://www.myonlinetraininghub.com/ | Modules 3, 5-6 |
| Excel Macro Mastery | https://www.excelmacromastery.com/ | Module 7 |
| Excel Off The Grid | https://exceloffthegrid.com/ | Modules 5-6 |
| Spreadsheeto | https://spreadsheeto.com/ | Modules 1-2 |
| Ablebits | https://www.ablebits.com/ | Modules 2-3 |
| Wall Street Prep | https://wallstreetprep.com/ | Module 5 |

---

## 3. Video Channels (Reference Only - Not Downloaded)

| Channel | URL | Topics |
|---------|-----|--------|
| Leila Gharani | https://www.youtube.com/@LeilaGharani | Advanced Excel, Power Query, Dashboards |
| ExcelIsFun (Mike Girvin) | https://www.youtube.com/@ExcelIsFun | All Levels, Dynamic Arrays, Array Formulas |
| MyOnlineTrainingHub | https://www.youtube.com/@MyOnlineTrainingHub | PivotTables, Power Query, Charts |
| Chandoo | https://www.youtube.com/@chandoo | Dashboards, Formulas, VBA |
| Wise Owl Tutorials | https://www.youtube.com/@WiseOwlTutorials | VBA, Macros |
| Curbal | https://www.youtube.com/@Curbal | Power Query, Power Pivot, DAX |
| Contextures (Debra Dalgleish) | https://www.youtube.com/@contextures | PivotTables, Data Validation |
| Kenji Explains | https://www.youtube.com/@KenjiExplains | Excel for Finance |
| Kevin Stratvert | https://www.youtube.com/@KevinStratvert | Excel Basics, New Features |
| Technology for Teachers | https://www.youtube.com/@TechnologyforTeachers | Beginner Excel |

### Recommended Starting Videos (by Level)

**Beginner:**
- Excel Tutorial for Beginners - Kevin Stratvert: https://www.youtube.com/watch?v=rwbk0c6MkG4
- Excel Full Course - freeCodeCamp: https://www.youtube.com/watch?v=NBI_ush7HhQ
- Excel Basics - Technology for Teachers: https://www.youtube.com/playlist?list=PL41LFaxdoQqOCiYE4JYJIhXrjXlPS-J1c

**Intermediate:**
- Excel PivotTables - Leila Gharani: https://www.youtube.com/watch?v=qu-Acm0tSag
- VLOOKUP vs INDEX MATCH vs XLOOKUP - Leila Gharani: https://www.youtube.com/watch?v=H1NQhMsMb_0
- Conditional Formatting Deep Dive - ExcelIsFun (search on YouTube): https://www.youtube.com/results?search_query=excelisfun+conditional+formatting

**Advanced:**
- Dynamic Arrays Masterclass - ExcelIsFun: https://www.youtube.com/watch?v=0P6wMb8GXoI
- Power Query Full Tutorial - MyOnlineTrainingHub: https://www.youtube.com/watch?v=Ohj5sXgRqSg
- VBA for Beginners - Wise Owl: https://www.youtube.com/playlist?list=PLNIs-AWhQ_nntIasbjq2ZnoXJeY53vp__

---

## 4. Academic & Professional Sources

| Source | URL/Reference | Used In |
|--------|--------------|---------|
| Journal of Accountancy - Excel Best Practices | https://www.journalofaccountancy.com/ | Modules 2, 4 |
| MIT OpenCourseWare - Excel | https://ocw.mit.edu/ | Module 1 |
| Coursera - Excel Skills for Business (Macquarie University) | https://www.coursera.org/specializations/excel | All Modules |
| edX - Excel for Beginners (free) | https://www.edx.org/learn/excel | Modules 4-5 |
| Khan Academy - Spreadsheets | https://www.khanacademy.org/computing/computer-programming | Module 1 |
| Google Scholar - Excel Analysis Papers | https://scholar.google.com/ | Module 4 |

---

## 5. Books & Publications

| Book | Author | ISBN | Used In |
|------|--------|------|---------|
| Excel Bible | John Walkenbach | 978-1119514787 | All Modules |
| Excel 2019 Power Programming with VBA | Michael Alexander & Dick Kusleika | 978-1119514923 | Module 7 |
| Power Query for Power BI and Excel | Chris Webb | 978-1484209271 | Module 6 |
| Dashboards and Reports with Excel | Michael Alexander | 978-1118490426 | Module 3-4 |
| Excel Data Analysis for Dummies | Paul McFedries | 978-1119518167 | Module 4 |
| Financial Modeling in Excel for Dummies | Danielle Stein Fairhurst | 978-1119357544 | Module 8 |
| MrExcel LIVe - The Excel Master Class | Bill Jelen | 978-1615470594 | All Modules |

---

## 6. Certification Bodies

| Certification | Provider | URL |
|--------------|----------|-----|
| MOS: Microsoft Excel Associate | Microsoft / Certiport | https://learn.microsoft.com/en-us/credentials/certifications/mos-excel-2019/ |
| MOS: Microsoft Excel Expert | Microsoft / Certiport | https://learn.microsoft.com/en-us/credentials/certifications/ |
| Excel Skills for Business Specialization | Coursera / Macquarie University | https://www.coursera.org/specializations/excel |
| CFI Financial Modeling Certification | Corporate Finance Institute | https://corporatefinanceinstitute.com/certifications/ |

---

## 7. Community & Forums

| Source | URL |
|--------|-----|
| r/excel (Reddit) | https://www.reddit.com/r/excel/ |
| MrExcel Forum | https://www.mrexcel.com/board/ |
| Stack Overflow - Excel Tag | https://stackoverflow.com/questions/tagged/excel |
| Excel Forum (Microsoft) | https://answers.microsoft.com/en-us/msoffice/forum/excel |
| Chandoo Forum | https://chandoo.org/forum/ |

---

## 8. Tools & Add-ins Referenced

| Tool | Purpose | URL |
|------|---------|-----|
| Power Query (built-in) | ETL / Data transformation | Included in Excel 2016+ |
| Power Pivot (built-in) | Data modeling with DAX | Included in Excel 2016+ |
| Solver (built-in add-in) | Optimization | Enable via File > Options > Add-ins |
| Analysis ToolPak | Statistical analysis | Enable via File > Options > Add-ins |
| Fuzzy Lookup Add-in | Fuzzy matching | https://support.microsoft.com/en-us/excel |
| XLSTAT | Statistical analysis | https://www.xlstat.com/ |
| Power BI Desktop | Advanced visualization | https://powerbi.microsoft.com/ |

---

## 9. Published Articles & Blog Posts Referenced

These articles and blog posts were referenced during the creation of this learning package. They provide additional depth, alternative explanations, and real-world context.

### Excel Fundamentals & Best Practices
| Article | Author/Source | URL | Used In |
|---------|--------------|-----|---------|
| "10 Excel Functions Every Beginner Should Know" | ExcelJet | https://exceljet.net/articles | Module 2 |
| "Excel Formulas: The Ultimate Guide" | TrumpExcel | https://trumpexcel.com/excel-formulas/ | Module 2 |
| "How to Use VLOOKUP: A Step-by-Step Guide" | Spreadsheeto | https://spreadsheeto.com/vlookup | Module 2 |
| "Conditional Formatting in Excel: Complete Guide" | Ablebits | https://www.ablebits.com/office-addins-blog/excel-conditional-formatting/ | Module 3 |
| "Excel PivotTable Tutorial: A Complete Guide" | Contextures | https://www.contextures.com/xlPivot01.html | Module 4 |

### Advanced Techniques
| Article | Author/Source | URL | Used In |
|---------|--------------|-----|---------|
| "XLOOKUP vs VLOOKUP vs INDEX/MATCH" | Leila Gharani | https://www.youtube.com/watch?v=H1NQhMsMb_0 | Module 5 |
| "Dynamic Arrays: FILTER, SORT, UNIQUE" | ExcelIsFun | https://www.youtube.com/watch?v=0P6wMb8GXoI | Module 5 |
| "Power Query Tutorial: A Complete Guide" | MyOnlineTrainingHub | https://www.myonlinetraininghub.com/power-query-overview | Module 6 |
| "DAX Basics: Introduction to DAX" | Curbal | https://www.youtube.com/@Curbal | Module 6 |
| "VBA for Beginners: A Complete Course" | Wise Owl Tutorials | https://www.youtube.com/playlist?list=PLNIs-AWhQ_nntIasbjq2ZnoXJeY53vp__ | Module 7 |

### Industry Applications
| Article | Author/Source | URL | Used In |
|---------|--------------|-----|---------|
| "Financial Modeling in Excel" | Wall Street Prep | https://wallstreetprep.com/ | Module 8 |
| "Excel for Data Analysis" | Corporate Finance Institute | https://corporatefinanceinstitute.com/resources/excel/ | Modules 4-5 |
| "Building Dashboards in Excel" | Chandoo | https://chandoo.org/?s=excelexcel-dashboards/ | Module 3 |
| "Excel vs Google Sheets: Key Differences" | Spreadsheeto | https://spreadsheeto.com/ | Module 1 |

### Certification & Career
| Article | Author/Source | URL | Used In |
|---------|--------------|-----|---------|
| "MOS Certification Guide" | Microsoft Learn | https://learn.microsoft.com/en-us/credentials/certifications/ | Study Plan |
| "Excel Skills for Business" | Coursera / Macquarie | https://www.coursera.org/specializations/excel | Study Plan |
| "How to Learn Excel: A Complete Roadmap" | Excel Off The Grid | https://exceloffthegrid.com/ | Study Plan |

---

## NotebookLM Upload Notes

**How to use these materials with NotebookLM:**
1. Upload all `.md` files from each module folder as individual sources
2. NotebookLM processes Markdown natively with proper formatting
3. The structured headings allow NotebookLM to create accurate citations
4. Use the Study Plan as a guide for which modules to study first
5. Ask NotebookLM questions like: "Explain VLOOKUP with examples from Module 2" or "Summarize the key PivotTable concepts from Module 4"

**Recommended upload order:**
1. Upload the Study Plan first for overview context
2. Upload modules in sequence (Module 1 through 8)
3. Upload this Sources file for reference context


## FILE: 10-Study-Plan/00-zero-to-hero-study-plan.md

# Zero to Hero Study Plans: 12 / 8 / 6 / 4 / 2 Weeks

**Choose the track that fits your schedule.** The content never changes, only the pace. Every track covers all 8 modules, the real-world projects, the sample workbooks, and all quizzes.

| Track | Weeks | Hours per week | Best for |
|------|-------|----------------|----------|
| [12-Week Mastery Track](#track-a-12-week-mastery) | 12 | 6-8 h | Beginners who want deep, lasting skill |
| [8-Week Standard Track](#track-b-8-week-standard) | 8 | 8-10 h | Most learners with a day job |
| [6-Week Accelerated Track](#track-c-6-week-accelerated) | 6 | 10-12 h | Learners with some Office experience |
| [4-Week Sprint Track](#track-d-4-week-sprint) | 4 | 12-15 h | Fast upskilling for a job or project |
| [2-Week Crash Track](#track-e-2-week-crash) | 2 | 18-25 h | Emergency prep (interview, certification, deadline) |

## How to Use Any Track

1. Pick a track below and copy its schedule into your own calendar.
2. Every study session follows the same rhythm:

| Step | Time | Activity |
|------|------|----------|
| 1 | 10 min | Review flashcards (Anki) from previous sessions |
| 2 | 40-60 min | Read the module section, follow along in Excel |
| 3 | 30-45 min | Do the exercises and the matching sample workbook |
| 4 | 20 min | Take the module quiz, score it, note weak areas |
| 5 | 10 min | Log progress in the [Progress Tracker](progress-tracker.md) |

3. **Score rule (same in every track):** quiz 90%+ advance, 70-89% review weak areas then advance, below 70% restudy the module before moving on.
4. Work in the sample workbooks in `../samples/` as you go: each one is paired with a walkthrough.
5. Want variety? Swap in the AI-assisted prompts from `../HOW-TO-USE-WITH-AI.md`.

---

## Track A: 12-Week Mastery

**Total: 72-96 hours. 5 study days per week (4 study + 1 review/build).**

| Week | Focus | Study Days (Mon-Thu) | Friday: Review & Build |
|------|-------|----------------------|------------------------|
| 1 | Module 1: Fundamentals | Interface, navigation, data entry, references, formats | Quiz 1 + `sample-01-data-entry-practice.xlsx` |
| 2 | Module 2: Formulas & Functions | SUM family, IF, COUNTIF/SUMIF, text/date functions | Quiz 2 + `sample-02-formula-challenges.xlsx` (tasks 1-6) |
| 3 | Module 3: Formatting & Visualization | Number formats, conditional formatting, tables, charts | Quiz 3 + `sample-03-dashboard-before-after.xlsx` rebuild |
| 4 | Module 4: Data Analysis | Sorting, filtering, validation, PivotTables, what-if | Quiz 4 + `sample-04-pivottable-source.xlsx` tasks |
| 5 | Module 5: Advanced Functions | INDEX/MATCH, XLOOKUP, dynamic arrays, LET/LAMBDA | Quiz 5 + `sample-05-lookup-challenges.xlsx` |
| 6 | Module 6: Power Query & Power Pivot | Query Editor, transforms, merge/append, DAX basics | Quiz 6 + `sample-06-powerquery-messy.xlsx` |
| 7 | Module 7: VBA & Macros | Recorder, editor, Subs, loops, error handling | Quiz 7 + `samples/07-vba-toolkit/` import and run |
| 8 | Module 8: Real-World Projects (start) | Budget tracker + Sales dashboard projects | Mid-course: `exercise-fix-this-workbook.xlsx` challenge |
| 9 | Projects (cont.) | Inventory system + Invoice generator | Peer review with the pitfalls tables |
| 10 | Projects (cont.) | Data cleaning pipeline + Financial model | Rebuild the fix-me workbook cleanly |
| 11 | Projects (finish) | Automated report generator + portfolio polish | Consistent formatting across all your workbooks |
| 12 | Capstone & Assessment | Review weak areas from quiz scores | Final: build a workbook for a real need, show it to a real human |

**12-week extras:** read one 'Further Learning' source per week beyond the modules. By week 12 you should have 8+ portfolio workbooks.

---

## Track B: 8-Week Standard

**Total: 64-80 hours. 5 study days per week. Modules pair up.**

| Week | Focus | Study Days (Mon-Thu) | Friday: Review & Build |
|------|-------|----------------------|------------------------|
| 1 | Modules 1-2 | Mon-Tue: Module 1 core. Wed-Thu: Module 2 core | Quizzes 1-2 + samples 01-02 |
| 2 | Module 3 | Formatting, conditional formatting, tables, charts | Quiz 3 + `sample-03` rebuild |
| 3 | Module 4 | Sorting, filtering, validation, PivotTables | Quiz 4 + `sample-04` tasks |
| 4 | Modules 5-6 | Mon-Tue: lookups + arrays. Wed-Thu: Power Query | Quizzes 5-6 + samples 05-06 |
| 5 | Module 7 | VBA basics + the toolkit macros | Quiz 7 + run and break a macro |
| 6 | Module 8 (half) | Budget tracker, sales dashboard, invoice generator | Fix-this-workbook challenge |
| 7 | Module 8 (half) | Cleaning pipeline, financial model, report generator | Rebuild the fix-me workbook |
| 8 | Capstone | Final real-need workbook + presentation to a human | Portfolio polish |

**8-week rule:** skip nothing, but do the 3 most relevant practice projects per module instead of all 5.

---

## Track C: 6-Week Accelerated

**Total: 60-72 hours. 6 study days per week. For people who already know Word/PowerPoint well.**

| Week | Focus | Daily Split |
|------|-------|-------------|
| 1 | Modules 1-2 | 3 days per module. Quizzes Fri-Sat |
| 2 | Modules 3-4 | 3 days per module. Samples 03-04 as practice |
| 3 | Modules 5-6 | 3 days per module. Samples 05-06 as practice |
| 4 | Module 7 + Projects 1-2 | Toolkit macros, then 2 build projects |
| 5 | Projects 3-5 | One project per 2 days |
| 6 | Projects 6-8 + Capstone | 1 day each project, final real-need workbook |

**6-week rule:** read every module fully, but do the 3 most relevant practice projects per module. Sunday off.

---

## Track D: 4-Week Sprint

**Total: 48-60 hours. 6-7 study days per week. Intense but complete.**

| Week | Focus | Daily Split |
|------|-------|-------------|
| 1 | Modules 1-2-3 (2 days each) | Fundamentals, formulas, formatting. Quizzes day 6 |
| 2 | Modules 4-5-6 (2 days each) | Analysis, lookups, Power Query. Samples 04-06 |
| 3 | Module 7 + Projects 1-3 | VBA day 1-2, then one project per day |
| 4 | Projects 4-8 (1 day each) | 5 mini-projects at speed, day 7: capstone + show it |

**4-week rule:** do 2 practice projects per module, review flashcards daily (20 cards/day), treat every quiz as a gate.

---

## Track E: 2-Week Crash

**Total: 36-50 hours. 7 days a week, 2-3 sessions per day. For deadlines: interviews, certification exams, a big review.**

| Days | Focus | Sessions per day |
|------|-------|------------------|
| Day 1 | Module 1 + Module 2 (core sections only) | 3 |
| Day 2 | Module 3 (formats, conditional formatting, charts) | 3 |
| Day 3 | Module 4 (sorting, filtering, PivotTables) | 3 |
| Day 4 | Module 5 (XLOOKUP, INDEX/MATCH) + Module 6 (Power Query core) | 3 |
| Day 5 | Module 7 (record + edit simple macros) + Quiz 7 | 3 |
| Day 6 | Quizzes 1-6 retakes, fix weak areas | 3 |
| Day 7 | Projects 1-2 (budget tracker + sales dashboard) | 3 |
| Day 8 | Projects 3-4 (inventory + invoice generator) | 3 |
| Day 9 | Projects 5-6 (cleaning pipeline + financial model) | 2 |
| Day 10 | Project 7 (report generator) + `exercise-fix-this-workbook.xlsx` | 2 |
| Day 11 | Rebuild the fix-me workbook cleanly end to end | 2 |
| Day 12 | Capstone: build your real-need workbook | 2 |
| Day 13 | Rehearse presenting it (Module 8 delivery tips) | 2 |
| Day 14 | Buffer: retake lowest quiz, polish portfolio, present | 2 |

**2-week rule:** read only the sections listed in each day's goal, skip 'Further Learning' reading, do 1 practice project per module, and use the AI tutor prompts to compress review time.

---

## After Any Track: Keep Growing

- Rebuild the sample workbooks in a different industry context (swap YourCompany Inc data for your own)
- Turn real work needs into workbooks every week (practice is the real teacher)
- Read one 'Further Learning' source per week from any module
- Retake quizzes after 30 days to test retention
- Explore LAMBDA and DAX further once the fundamentals are automatic

## Files You Need

| File | Purpose |
|------|---------|
| [progress-tracker.md](progress-tracker.md) | Checkbox checklist for every skill |
| `../quizzes/quiz-01-fundamentals.md` to `quiz-08-projects.md` | Module quizzes with answer keys |
| `../anki/excel-flashcards.csv` | Daily spaced repetition |
| `../samples/` | 8 practice workbooks + walkthroughs |
| `../HOW-TO-USE-WITH-AI.md` | AI tutor prompts to speed up any track |


## FILE: 10-Study-Plan/progress-tracker.md

# Excel Zero to Hero - Progress Tracker

Use this checklist to track your learning journey. Check off each item as you complete it.

**Pace note:** the daily checklist below follows the 12-week Mastery track. On the 8/6/4/2-week tracks in [00-zero-to-hero-study-plan.md](00-zero-to-hero-study-plan.md), compress the same items into your schedule: the skills are identical, only the calendar changes.

**Rating Scale:** 1 = Don't understand | 3 = Can do with help | 5 = Confident independently

---

## PHASE 1: FOUNDATION (Weeks 1-3)

### Week 1: Getting Started
- [ ] Day 1: Explored the Ribbon interface and identified all tabs
- [ ] Day 2: Practiced cell navigation (Arrow keys, Ctrl+Arrow, Ctrl+Home)
- [ ] Day 3: Entered text, numbers, dates and applied basic formatting
- [ ] Day 4: Created formulas with relative and absolute references ($)
- [ ] Day 5: Saved files in .xlsx, .csv, and .xlsb formats
- [ ] Day 6: Completed all Module 1 exercises
- [ ] Day 7: Watched Kevin Stratvert beginner video
- **Milestone:** Can navigate Excel and enter data confidently
- **Self-rating (1-5):** ___

### Week 2: Formulas & Functions
- [ ] Day 1: Wrote basic formulas, understood PEMDAS order
- [ ] Day 2: Used SUM, AVERAGE, COUNT, COUNTA, MIN, MAX
- [ ] Day 3: Applied SUMIF, SUMIFS, COUNTIF, COUNTIFS
- [ ] Day 4: Built IF, Nested IF, IFS, SWITCH formulas
- [ ] Day 5: Used LEFT, RIGHT, MID, TRIM, CONCATENATE
- [ ] Day 6: Used TODAY, NOW, DATEDIF, EOMONTH
- [ ] Day 7: Built VLOOKUP formulas with IFERROR
- **Milestone:** Can write formulas confidently and use 20+ functions
- **Self-rating (1-5):** ___

### Week 3: Formatting & Visualization
- [ ] Day 1: Applied number formats (currency, %, date, custom)
- [ ] Day 2: Created conditional formatting rules
- [ ] Day 3: Created Excel Tables with structured references
- [ ] Day 4: Built Column, Bar, Line, and Pie charts
- [ ] Day 5: Formatted charts with titles, legends, trendlines
- [ ] Day 6: Added Sparklines and Slicers
- [ ] Day 7: Built a mini-dashboard with charts
- **Milestone:** Can present data visually and format professionally
- **Self-rating (1-5):** ___

---

## PHASE 2: INTERMEDIATE (Weeks 4-6)

### Week 4: Data Analysis
- [ ] Day 1: Sorted by multiple columns, used AutoFilter
- [ ] Day 2: Created advanced filter criteria ranges
- [ ] Day 3: Built data validation drop-down lists
- [ ] Day 4: Created first PivotTable from scratch
- [ ] Day 5: Grouped PivotTable data, changed value settings
- [ ] Day 6: Built PivotCharts with interactive slicers
- [ ] Day 7: Analyzed a full dataset with PivotTables
- **Milestone:** Can analyze large datasets and create interactive reports
- **Self-rating (1-5):** ___

### Week 5: What-If Analysis & Dashboards
- [ ] Day 1: Used Goal Seek for break-even analysis
- [ ] Day 2: Created scenarios and 1/2-variable data tables
- [ ] Day 3: Applied automatic subtotals and grouping
- [ ] Day 4: Built formula-based conditional formatting rules
- [ ] Day 5: Learned dashboard design principles
- [ ] Day 6: Built a complete sales dashboard
- [ ] Day 7: Refined and polished dashboard
- **Milestone:** Can build interactive dashboards and perform what-if analysis
- **Self-rating (1-5):** ___

### Week 6: Advanced Lookups
- [ ] Day 1: Built INDEX/MATCH single-criteria lookups
- [ ] Day 2: Built INDEX/MATCH multi-criteria lookups
- [ ] Day 3: Used XLOOKUP with all parameters
- [ ] Day 4: Used OFFSET and INDIRECT for dynamic ranges
- [ ] Day 5: Used FILTER, SORT, UNIQUE dynamic arrays
- [ ] Day 6: Used LET, LAMBDA, TEXTBEFORE/AFTER
- [ ] Day 7: Completed multi-sheet lookup challenge
- **Milestone:** Can write advanced lookup formulas and use dynamic arrays
- **Self-rating (1-5):** ___

---

## PHASE 3: ADVANCED (Weeks 7-9)

### Week 7: Power Query
- [ ] Day 1: Imported data from CSV, text, and web sources
- [ ] Day 2: Applied basic transformations (filter, change types, rename)
- [ ] Day 3: Split/merged columns, pivoted/unpivoted data
- [ ] Day 4: Merged queries with different join types
- [ ] Day 5: Appended queries and created custom columns
- [ ] Day 6: Set up parameters and auto-refresh
- [ ] Day 7: Cleaned a messy dataset end-to-end
- **Milestone:** Can transform and clean data using Power Query
- **Self-rating (1-5):** ___

### Week 8: Power Pivot & DAX
- [ ] Day 1: Created data model with relationships
- [ ] Day 2: Built PivotTables from the Data Model
- [ ] Day 3: Wrote calculated columns with SUMX, RELATED
- [ ] Day 4: Created measures with CALCULATE
- [ ] Day 5: Built KPIs and hierarchies
- [ ] Day 6: Built full Power Query to Power Pivot pipeline
- [ ] Day 7: Completed multi-table sales analysis
- **Milestone:** Can build data models with relationships and write DAX measures
- **Self-rating (1-5):** ___

### Week 9: VBA & Macros
- [ ] Day 1: Recorded and edited macros
- [ ] Day 2: Used VBA Editor, wrote Sub/Function, MsgBox
- [ ] Day 3: Worked with Range objects and loops
- [ ] Day 4: Used If/Select Case and error handling
- [ ] Day 5: Automated sheet and workbook operations
- [ ] Day 6: Built practical automation tasks
- [ ] Day 7: Created a basic UserForm
- **Milestone:** Can write VBA macros to automate repetitive tasks
- **Self-rating (1-5):** ___

---

## PHASE 4: MASTERY (Weeks 10-12)

### Week 10: Projects Part 1
- [ ] Day 1-2: Built Personal Budget Tracker
- [ ] Day 3-4: Built Sales Dashboard
- [ ] Day 5-6: Built Inventory Management System
- [ ] Day 7: Reviewed and polished all 3 projects

### Week 11: Projects Part 2
- [ ] Day 1: Built Employee Schedule Tracker
- [ ] Day 2: Built Invoice Generator
- [ ] Day 3-4: Built Data Cleaning Pipeline
- [ ] Day 5-6: Built Financial Model
- [ ] Day 7: Built Automated Report Generator

### Week 12: Certification Prep
- [ ] Day 1: Completed self-assessment of all modules
- [ ] Day 2: Practiced MOS Associate test questions
- [ ] Day 3: Practiced MOS Expert test questions
- [ ] Day 4: Reviewed weak areas
- [ ] Day 5: Built portfolio project combining ALL skills
- [ ] Day 6: Explored Power BI as next step
- [ ] Day 7: Final review (and certification exam if ready)

---

## Skills Acquired Checklist

### Formulas & Functions
- [ ] Can write SUM, AVERAGE, COUNT, MIN, MAX
- [ ] Can write SUMIF/SUMIFS, COUNTIF/COUNTIFS
- [ ] Can write IF, Nested IF, IFS, SWITCH
- [ ] Can use text functions (LEFT, RIGHT, MID, TRIM, CONCAT)
- [ ] Can use date functions (TODAY, DATEDIF, EOMONTH, NETWORKDAYS)
- [ ] Can write VLOOKUP with error handling
- [ ] Can write INDEX/MATCH (single and multi-criteria)
- [ ] Can write XLOOKUP
- [ ] Can use dynamic arrays (FILTER, SORT, UNIQUE)
- [ ] Can use OFFSET and INDIRECT

### Data Analysis
- [ ] Can sort and filter data effectively
- [ ] Can create and configure PivotTables
- [ ] Can group PivotTable data and change value settings
- [ ] Can use Goal Seek and Scenario Manager
- [ ] Can create 1-variable and 2-variable data tables
- [ ] Can set up data validation with drop-down lists

### Visualization
- [ ] Can create and format charts (Column, Bar, Line, Pie)
- [ ] Can apply conditional formatting rules
- [ ] Can build interactive dashboards with slicers
- [ ] Can create Sparklines

### Power Tools
- [ ] Can import and transform data with Power Query
- [ ] Can merge and append queries
- [ ] Can create data models with relationships
- [ ] Can write basic DAX measures

### Automation
- [ ] Can record and edit macros
- [ ] Can write VBA Sub procedures
- [ ] Can use loops and conditional logic in VBA
- [ ] Can create basic UserForms

---

## Sample Workbooks

- [ ] `sample-01-data-entry-practice.xlsx` cleaned with zero green triangles
- [ ] `sample-02-formula-challenges.xlsx` all 12 tasks matched expected results
- [ ] `sample-03-dashboard-before-after.xlsx` After sheet rebuilt from Before
- [ ] `sample-04-pivottable-source.xlsx` all 6 pivot tasks done
- [ ] `sample-05-lookup-challenges.xlsx` lookups filled + answers matched
- [ ] `sample-06-powerquery-messy.xlsx` clean refreshable table built
- [ ] `samples/07-vba-toolkit/` macros imported, run, and one broken-then-fixed
- [ ] `exercise-fix-this-workbook.xlsx` all 20 planted issues found
- [ ] Fix-me workbook rebuilt cleanly

---

## Final Assessment

**Date completed:** _______________

**Total weeks taken:** ___

**Modules completed:** ___/8

**Quizzes passed (70%+):** ___/8

**Projects built:** ___/8

**Ready for MOS certification:** Yes / No / Almost

**Next steps planned:** _______________


## FILE: HOW-TO-USE-WITH-AI.md

# How to Use This Package with AI Agents & NotebookLM

This guide explains how to feed the Excel Zero to Hero learning package to different AI platforms for maximum effectiveness.

---

## 1. NotebookLM (Google)

**URL:** https://notebooklm.google.com

### Step-by-Step Setup

1. Go to notebooklm.google.com and sign in with your Google account
2. Click **"Create new notebook"**
3. Name it: `Excel Zero to Hero`

### Simplest path: one link, no downloads (recommended)

Skip the download-and-upload dance entirely. Click **Add source > Website** and paste this single link. NotebookLM then holds the whole course at once: every module, quiz, cheat sheet, walkthrough, and answer key.

```
https://raw.githubusercontent.com/tempesteni/Microsoft-Excel-Zero-to-hero/main/curriculum-full.md
```

That is it, you are ready to learn. (If a tool ever truncates such a large page, use the per-file path below.)

4. Upload sources in this order (manual path, one file at a time):

| Order | File | Why This Order |
|-------|------|---------------|
| 1 | `10-Study-Plan/00-zero-to-hero-study-plan.md` | Gives NotebookLM the full curriculum overview |
| 2 | `01-Fundamentals/01-excel-fundamentals.md` | Foundation context |
| 3 | `02-Formulas-Functions/02-formulas-functions.md` | Builds on fundamentals |
| 4 | `03-Formatting-Visualization/03-formatting-visualization.md` | Sequential |
| 5 | `04-Data-Analysis/04-data-analysis.md` | Sequential |
| 6 | `05-Advanced-Functions/05-advanced-functions.md` | Sequential |
| 7 | `06-Power-Query-Pivot/06-power-query-power-pivot.md` | Sequential |
| 8 | `07-VBA-Macros/07-vba-macros.md` | Sequential |
| 9 | `08-Real-World-Projects/08-real-world-projects.md` | Projects apply all skills |
| 10 | `09-Sources-References/sources-references-index.md` | Reference context |
| 11 | `10-Study-Plan/progress-tracker.md` | Tracking context |
| 12-19 | `quizzes/quiz-01.md` through `quiz-08.md` | Assessment content |
| 20-23 | `cheat-sheets/*.md` | Quick reference context |

### Prompts to Use in NotebookLM

**Learning & Explanation:**
```
Explain VLOOKUP with a real-world example. Compare it to INDEX/MATCH and tell me when to use each.
```
```
I'm a complete beginner. Walk me through creating my first PivotTable step by step using the sales data example from Module 4.
```
```
What's the difference between Power Query and Power Pivot? Give me a decision framework for when to use each.
```
```
Summarize the key concepts from Module 6 (Power Query) in bullet points I can review before my quiz.
```

**Practice & Quizzing:**
```
Quiz me on Module 2 (Formulas & Functions). Ask me 5 questions, wait for my answers, then tell me my score and what to review.
```
```
Give me a practice challenge: I have a spreadsheet with employee names, departments, and salaries. Write 5 formulas I should be able to create, then give me the answers.
```
```
Create a fill-in-the-blank exercise for XLOOKUP syntax. Leave blanks for each parameter and give me the answers at the end.
```

**Study Planning:**
```
I have 2 hours today. Based on the study plan, what should I focus on if I just finished Week 3?
```
```
I'm struggling with PivotTables. Create a focused 3-day mini-study plan just for PivotTables using the Module 4 content.
```

**Real-World Application:**
```
I work in HR. How can I use the skills from Modules 3-4 to build a headcount dashboard? Give me specific steps.
```
```
I need to clean a messy CSV file with inconsistent date formats and duplicate rows. Walk me through the Power Query approach from Module 6.
```

**Audio Overview (NotebookLM Feature):**
- After uploading, click **"Audio Overview"** to generate a podcast-style summary
- Listen while commuting or exercising
- Use it to review before quizzes

---

## 2. ChatGPT / Claude / General AI Chatbots

### How to Feed as Context

**Option A: Paste Module Content Directly**
```
I'm going to paste the content of an Excel tutorial module. After I paste it, I'll ask you questions about it. Please use ONLY the content I provide to answer, and reference specific sections when possible.

[PASTE MODULE CONTENT HERE]

Now explain conditional formatting with 3 practical examples from this module.
```

**Option B: Upload as File (if supported)**
```
I'm uploading an Excel learning module. Read it thoroughly, then:
1. Summarize the key concepts in 10 bullet points
2. Create 5 practice exercises based on the content
3. Identify the 3 most important formulas to memorize
4. Suggest a 1-hour study plan for this module
```

**Option C: Multi-Module Conversation**
```
I'm studying Excel from a structured curriculum. I'll share modules one at a time. After each module:
- Confirm what you understood
- Create 3 quiz questions
- Tell me what I should focus on next

Let's start with Module 1: Fundamentals.

[PASTE MODULE 1]
```

### Prompts for Specific Learning Goals

**For Beginners:**
```
I have zero Excel experience. I'm sharing Module 1 of a learning curriculum. Read it and then:
1. Explain the 5 most important concepts like I'm 10 years old
2. Give me a "cheat code" - the one thing that will make everything else click
3. Create a 30-minute practice exercise I can do right now
```

**For Interview Prep:**
```
I have a job interview next week that requires Excel skills. I'm sharing my learning materials. Based on this content:
1. What are the top 10 Excel questions interviewers ask?
2. Give me model answers for each
3. Create a 2-day crash study plan focusing on what matters most for interviews
```

**For Certification (MOS):**
```
I'm preparing for the MOS Excel certification. I'm sharing my study materials. Based on this:
1. Which topics from these materials map to MOS exam objectives?
2. Create a practice test with 20 questions in MOS exam style
3. Score my answers and tell me which areas need more study
```

**For Teaching Others:**
```
I need to teach Excel to my team. I'm sharing a learning curriculum. Based on this:
1. Create a 1-hour workshop outline for absolute beginners
2. Include 3 hands-on exercises everyone should try
3. Suggest how to break this into 4 weekly 1-hour sessions
```

---

## 3. GitHub Copilot / Cursor / Code Assistants

### How to Use for VBA Learning (Module 7)

```
I'm learning VBA from this Excel tutorial. Read Module 7 and then:
1. Explain each VBA concept with a simpler example
2. If I paste my VBA code, debug it using the patterns from this module
3. Suggest improvements to make my code more efficient

Here's the module content:
[PASTE MODULE 7]
```

### How to Use for Formula Building

```
I need to build a complex Excel formula. I'm sharing my reference material on Excel functions (Module 2). Help me:
1. Build the formula I describe step by step
2. Explain each part of the formula
3. Suggest a simpler alternative if one exists

The formula I need: [DESCRIBE WHAT YOU WANT TO CALCULATE]
```

---

## 4. RAG Systems / Vector Databases

### For AI Engineers Building Excel Training Agents

**How to index this package for RAG:**

```
Source format: Markdown (.md)
Chunking strategy: Split by ## headings (each section is self-contained)
Metadata to attach per chunk:
  - module_number (01-08)
  - difficulty (beginner/intermediate/advanced)
  - topic (formulas, pivot, vba, etc.)
  - content_type (lesson, exercise, quiz, cheat_sheet)

Embedding recommendation: Use text-embedding-3-small or similar
Retrieval: Top-k=5 with metadata filter on module_number
```

**Sample RAG query pipeline:**
```python
# Pseudocode for an Excel tutor agent
def answer_excel_question(question, user_level="beginner"):
    # Retrieve relevant chunks
    chunks = vector_db.search(
        query=question,
        top_k=5,
        filter={"difficulty": user_level}
    )
    
    # Build context
    context = "\n\n".join([c.text for c in chunks])
    
    # Generate answer
    prompt = f"""You are an Excel tutor. Use ONLY the following context to answer.
    Reference the specific module and section when possible.
    
    Context:
    {context}
    
    Student question: {question}
    
    Answer clearly, with an example, and suggest what to learn next."""
    
    return llm.generate(prompt)
```

**Recommended agent capabilities:**
- Quiz mode: Generate questions from quiz files, grade answers
- Explain mode: Break down complex concepts from module content
- Practice mode: Generate exercises using practice dataset structure
- Progress mode: Track which modules are completed, suggest next steps

---

## 5. Microsoft Copilot (Excel)

### How to Use Inside Excel

While working in Excel with Copilot enabled, reference the learning package:

```
Based on Excel best practices, help me create a PivotTable from this data.
Show me how to group dates by quarter and add a calculated field for profit margin.
```

```
I'm learning conditional formatting. Show me how to:
1. Highlight cells above average in green
2. Add data bars to this column
3. Create a formula-based rule to highlight entire rows where status = "Overdue"
```

```
Help me write a VBA macro that:
1. Loops through all sheets
2. Applies the same header formatting
3. Adds auto-filters
Reference the VBA patterns from Excel best practices.
```

---

## 6. Perplexity / AI Search Engines

### How to Use for Research

```
Based on this Excel learning curriculum, what are the best free resources to practice PivotTables? Include specific websites, YouTube channels, and practice datasets.
```

```
I'm following a structured Excel course that covers Power Query in Module 6. What are the most common Power Query errors beginners face, and how do I fix them?
```

```
Compare the MOS Excel Associate certification vs. the Excel Skills for Business Coursera specialization. Which is better for someone following a self-study curriculum?
```

---

## Quick Reference: Which Platform for What

| Goal | Best Platform | Why |
|------|--------------|-----|
| Deep learning & explanation | NotebookLM | Source-grounded, cites modules |
| Practice & quizzing | ChatGPT / Claude | Interactive, can grade answers |
| Formula building | Copilot in Excel | Lives inside the tool |
| VBA debugging | Cursor / Copilot | Code-aware |
| Building training bots | RAG system | Scalable, customizable |
| Quick research | Perplexity | Web-augmented answers |
| Study planning | NotebookLM | Can reference all modules at once |

---

## Tips for Best Results

1. **Always provide module context** - Don't assume the AI knows the curriculum
2. **Specify your level** - "I'm a beginner" vs "I'm intermediate" changes the answer depth
3. **Ask for examples** - Abstract explanations are less useful than "show me with this data"
4. **Request practice** - After learning a concept, ask for an exercise
5. **Use the quiz files** - Feed them to the AI and ask it to quiz you interactively
6. **Reference the cheat sheets** - Paste the relevant cheat sheet when working on that topic
7. **Track progress** - Use the progress tracker and ask the AI to update it with you


## FILE: cheat-sheets/functions-cheat-sheet.md

# Excel Functions Cheat Sheet - Top 50

> **Quick Reference** | Print on A4/Letter | 2 pages

---

## MATH & TRIG FUNCTIONS

| Function | Syntax | Description | Example |
|----------|--------|-------------|---------|
| `SUM` | `=SUM(number1, [number2], …)` | Adds all numbers in a range | `=SUM(A1:A10)` → 55 |
| `SUMIF` | `=SUMIF(range, criteria, [sum_range])` | Sums cells matching criteria | `=SUMIF(A:A,"Sales",B:B)` |
| `SUMIFS` | `=SUMIFS(sum_range, crit_range1, crit1, …)` | Sums with multiple criteria | `=SUMIFS(C:C,A:A,"East",B:B,">100")` |
| `SUMPRODUCT` | `=SUMPRODUCT(array1, [array2], …)` | Multiplies then sums arrays | `=SUMPRODUCT(A1:A5,B1:B5)` |
| `ROUND` | `=ROUND(number, num_digits)` | Rounds to specified digits | `=ROUND(3.456,2)` → 3.46 |
| `ROUNDUP` | `=ROUNDUP(number, num_digits)` | Always rounds up | `=ROUNDUP(3.1,0)` → 4 |
| `ROUNDDOWN` | `=ROUNDDOWN(number, num_digits)` | Always rounds down | `=ROUNDDOWN(3.9,0)` → 3 |
| `INT` | `=INT(number)` | Rounds down to nearest integer | `=INT(8.9)` → 8 |
| `ABS` | `=ABS(number)` | Returns absolute value | `=ABS(-5)` → 5 |
| `MOD` | `=MOD(number, divisor)` | Returns remainder after division | `=MOD(10,3)` → 1 |
| `RAND` | `=RAND()` | Random number between 0 and 1 | `=RAND()` → 0.73 |
| `RANDBETWEEN` | `=RANDBETWEEN(bottom, top)` | Random integer in range | `=RANDBETWEEN(1,100)` |
| `SQRT` | `=SQRT(number)` | Square root | `=SQRT(144)` → 12 |
| `POWER` | `=POWER(number, power)` | Raises to a power | `=POWER(2,10)` → 1024 |
| `SUBTOTAL` | `=SUBTOTAL(function_num, ref)` | Sum/avg/count ignoring filtered rows | `=SUBTOTAL(109,A1:A100)` |

---

## TEXT FUNCTIONS

| Function | Syntax | Description | Example |
|----------|--------|-------------|---------|
| `LEN` | `=LEN(text)` | Returns number of characters | `=LEN("Hello")` → 5 |
| `LEFT` | `=LEFT(text, [num_chars])` | Extracts characters from left | `=LEFT("Excel",2)` → "Ex" |
| `RIGHT` | `=RIGHT(text, [num_chars])` | Extracts characters from right | `=RIGHT("Excel",3)` → "cel" |
| `MID` | `=MID(text, start_num, num_chars)` | Extracts from middle of text | `=MID("Excel",2,3)` → "xce" |
| `TRIM` | `=TRIM(text)` | Removes extra spaces | `=TRIM(" Hi ")` → "Hi" |
| `UPPER` | `=UPPER(text)` | Converts to uppercase | `=UPPER("hello")` → "HELLO" |
| `LOWER` | `=LOWER(text)` | Converts to lowercase | `=LOWER("HELLO")` → "hello" |
| `CONCATENATE` / `&` | `=text1 & text2` | Joins text strings | `="Hello" & " " & "World"` |
| `CONCAT` | `=CONCAT(text1, [text2], …)` | Joins text (range support) | `=CONCAT(A1:A5)` |
| `TEXTJOIN` | `=TEXTJOIN(delimiter, ignore_empty, …)` | Joins with delimiter | `=TEXTJOIN(",",TRUE,A1:A5)` |
| `SUBSTITUTE` | `=SUBSTITUTE(text, old, new, [n])` | Replaces text | `=SUBSTITUTE("A-B-C","-","/")` |
| `TEXT` | `=TEXT(value, format_text)` | Formats number as text | `=TEXT(0.5,"0%")` → "50%" |
| `FIND` | `=FIND(find_text, within, [start])` | Finds position (case-sensitive) | `=FIND("e","Excel")` → 2 |

---

## DATE & TIME FUNCTIONS

| Function | Syntax | Description | Example |
|----------|--------|-------------|---------|
| `TODAY` | `=TODAY()` | Current date | `=TODAY()` → 9/17/2026 |
| `NOW` | `=NOW()` | Current date and time | `=NOW()` → 9/17/2026 14:30 |
| `DATE` | `=DATE(year, month, day)` | Creates a date from components | `=DATE(2026,9,17)` |
| `YEAR` | `=YEAR(serial_number)` | Extracts year from date | `=YEAR(TODAY())` → 2026 |
| `MONTH` | `=MONTH(serial_number)` | Extracts month from date | `=MONTH(TODAY())` → 9 |
| `DAY` | `=DAY(serial_number)` | Extracts day from date | `=DAY(TODAY())` → 17 |
| `DATEDIF` | `=DATEDIF(start, end, unit)` | Difference between dates | `=DATEDIF(A1,B1,"D")` → days |
| `EOMONTH` | `=EOMONTH(start_date, months)` | End of month | `=EOMONTH(TODAY(),0)` |
| `NETWORKDAYS` | `=NETWORKDAYS(start, end, [holidays])` | Working days between dates | `=NETWORKDAYS(A1,B1)` |
| `WEEKDAY` | `=WEEKDAY(serial, [return_type])` | Day of week (1=Sun) | `=WEEKDAY(TODAY())` |

---

## LOOKUP & REFERENCE FUNCTIONS

| Function | Syntax | Description | Example |
|----------|--------|-------------|---------|
| `VLOOKUP` | `=VLOOKUP(lookup, table, col, [match])` | Vertical lookup (legacy) | `=VLOOKUP(A1,B:D,3,FALSE)` |
| `XLOOKUP` | `=XLOOKUP(lookup_value, lookup_array, return_array, [if_not_found], [match_mode], [search_mode])` | Modern lookup (any direction) | `=XLOOKUP(A1,A:A,C:C)` |
| `INDEX` | `=INDEX(array, row_num, [col_num])` | Returns value at position | `=INDEX(B2:D10,3,2)` |
| `MATCH` | `=MATCH(lookup, array, [match_type])` | Returns position of match | `=MATCH("Sales",A:A,0)` → 5 |
| `HLOOKUP` | `=HLOOKUP(lookup, table, row, [match])` | Horizontal lookup | `=HLOOKUP("Q1",A1:D3,2,FALSE)` |
| `INDIRECT` | `=INDIRECT(ref_text, [a1])` | Creates reference from text | `=INDIRECT("A"&B1)` |

---

## LOGICAL FUNCTIONS

| Function | Syntax | Description | Example |
|----------|--------|-------------|---------|
| `IF` | `=IF(logical_test, if_true, if_false)` | Conditional value | `=IF(A1>100,"High","Low")` |
| `IFS` | `=IFS(test1,val1, test2,val2, …)` | Multiple conditions (no nesting) | `=IFS(A1>90,"A",A1>80,"B")` |
| `AND` | `=AND(logical1, [logical2], …)` | TRUE if all conditions met | `=AND(A1>0,B1>0)` |
| `OR` | `=OR(logical1, [logical2], …)` | TRUE if any condition met | `=OR(A1="Yes",B1="Yes")` |
| `IFERROR` | `=IFERROR(value, value_if_error)` | Returns alt value on error | `=IFERROR(A1/B1,0)` |
| `IFNA` | `=IFNA(value, value_if_na)` | Returns alt value on #N/A | `=IFNA(VLOOKUP(…),"Not Found")` |
| `SWITCH` | `=SWITCH(expr, val1, result1, …)` | Match expression to values | `=SWITCH(A1,1,"Mon",2,"Tue")` |

---

## STATISTICAL FUNCTIONS

| Function | Syntax | Description | Example |
|----------|--------|-------------|---------|
| `AVERAGE` | `=AVERAGE(number1, [number2], …)` | Arithmetic mean | `=AVERAGE(A1:A10)` |
| `AVERAGEIF` | `=AVERAGEIF(range, criteria, [avg_range])` | Average with condition | `=AVERAGEIF(A:A,"East",B:B)` |
| `COUNT` | `=COUNT(value1, [value2], …)` | Counts cells with numbers | `=COUNT(A1:A100)` |
| `COUNTA` | `=COUNTA(value1, [value2], …)` | Counts non-empty cells | `=COUNTA(A1:A100)` |
| `COUNTIF` | `=COUNTIF(range, criteria)` | Counts cells matching criteria | `=COUNTIF(A:A,"Sales")` |
| `COUNTIFS` | `=COUNTIFS(range1, crit1, …)` | Counts with multiple criteria | `=COUNTIFS(A:A,"East",B:B,">100")` |
| `MAX` | `=MAX(number1, [number2], …)` | Largest value | `=MAX(A1:A10)` |
| `MIN` | `=MIN(number1, [number2], …)` | Smallest value | `=MIN(A1:A10)` |
| `MEDIAN` | `=MEDIAN(number1, [number2], …)` | Middle value | `=MEDIAN(1,2,3,4,5)` → 3 |
| `MODE` | `=MODE(number1, [number2], …)` | Most frequent value | `=MODE(1,2,2,3)` → 2 |

---

## 📋 Pro Tips

- **`Ctrl+`` `** toggles showing formulas vs results
- Use **named ranges** (`Formulas > Define Name`) for readable formulas
- Press **`F4`** while editing to cycle absolute/relative references (`$A$1` → `A1`)
- **`XLOOKUP`** replaces `VLOOKUP`/`HLOOKUP`/`INDEX-MATCH` in Microsoft 365
- Use **`SUBTOTAL`** instead of `SUM` for filtered tables (function 109 = SUM)
- Wrap any function in **`IFERROR`** to handle errors gracefully


## FILE: cheat-sheets/keyboard-shortcuts-cheat-sheet.md

# Excel Keyboard Shortcuts Cheat Sheet - 60+ Essential Shortcuts

> **Quick Reference** | Print on A4/Letter | 1-2 pages

---

## NAVIGATION SHORTCUTS

| Shortcut | Action |
|----------|--------|
| `Ctrl + Home` | Go to cell A1 |
| `Ctrl + End` | Go to last used cell |
| `Ctrl + G` or `F5` | Go To dialog |
| `Ctrl + ↑/↓/←/→` | Jump to edge of data region |
| `Ctrl + Page Up` | Previous worksheet tab |
| `Ctrl + Page Down` | Next worksheet tab |
| `Ctrl + Tab` | Switch between open workbooks |
| `Alt + Page Down` | Move one screen right |
| `Alt + Page Up` | Move one screen left |
| `Ctrl + Backspace` | Scroll to show active cell |
| `Ctrl + F` | Find |
| `Ctrl + H` | Find and Replace |

---

## SELECTION SHORTCUTS

| Shortcut | Action |
|----------|--------|
| `Ctrl + Space` | Select entire column |
| `Shift + Space` | Select entire row |
| `Ctrl + A` | Select entire worksheet (or current region) |
| `Ctrl + Shift + End` | Select from active cell to last used cell |
| `Ctrl + Shift + Home` | Select from active cell to A1 |
| `Shift + Arrow Keys` | Extend selection by one cell |
| `Ctrl + Shift + Arrow` | Extend selection to edge of data |
| `Shift + Click` | Select range from anchor to clicked cell |
| `Ctrl + Click` | Add non-adjacent cell to selection |
| `Ctrl + Shift + L` | Toggle AutoFilter (select headers first) |

---

## EDITING SHORTCUTS

| Shortcut | Action |
|----------|--------|
| `F2` | Edit active cell |
| `Enter` | Confirm and move down |
| `Tab` | Confirm and move right |
| `Shift + Enter` | Confirm and move up |
| `Shift + Tab` | Confirm and move left |
| `Escape` | Cancel edit |
| `Delete` | Clear cell contents |
| `Ctrl + -` (minus) | Delete selected cells/rows/columns |
| `Ctrl + Shift + +` (plus) | Insert cells/rows/columns |
| `Ctrl + D` | Fill down (copy cell above) |
| `Ctrl + R` | Fill right (copy cell to left) |
| `Ctrl + '` (apostrophe) | Copy formula from cell above |
| `Ctrl + ;` (semicolon) | Insert current date |
| `Ctrl + Shift + ;` | Insert current time |
| `Alt + Enter` | New line within a cell |
| `Ctrl + Z` | Undo |
| `Ctrl + Y` | Redo / Repeat last action |

---

## FORMATTING SHORTCUTS

| Shortcut | Action |
|----------|--------|
| `Ctrl + B` | **Bold** |
| `Ctrl + I` | *Italic* |
| `Ctrl + U` | Underline |
| `Ctrl + 5` | ~~Strikethrough~~ |
| `Ctrl + Shift + ~` | General number format |
| `Ctrl + Shift + $` | Currency format |
| `Ctrl + Shift + %` | Percentage format |
| `Ctrl + Shift + #` | Date format |
| `Ctrl + Shift + @` | Time format |
| `Ctrl + Shift + !` | Number with 2 decimals & comma |
| `Ctrl + 1` | Format Cells dialog |
| `Ctrl + Shift + &` | Add outline border |
| `Ctrl + Shift + _` | Remove all borders |
| `Alt + H + H` | Fill color picker |
| `Alt + H + FC` | Font color picker |
| `Alt + H + O + I` | Auto-fit column width |
| `Ctrl + Shift + F` | Format Cells > Font tab |

---

## FORMULA SHORTCUTS

| Shortcut | Action |
|----------|--------|
| `=` | Start a formula |
| `F4` | Toggle absolute/relative reference (`$A$1` ↔ `A1`) |
| `Ctrl + `` ` | Show/hide formulas (toggle) |
| `F9` | Calculate all workbooks |
| `Shift + F9` | Calculate active worksheet only |
| `Ctrl + Shift + Enter` | Enter as array formula (legacy) |
| `Ctrl + [` | Select all precedent cells |
| `Ctrl + ]` | Select all dependent cells |
| `Alt + =` | AutoSum |
| `Ctrl + Shift + U` | Expand/collapse formula bar |
| `Tab` | Accept AutoComplete suggestion in formulas |
| `Ctrl + A` | Insert function arguments (inside function tooltip) |

---

## DATA & TABLE SHORTCUTS

| Shortcut | Action |
|----------|--------|
| `Ctrl + T` | Create Table from selection |
| `Ctrl + Shift + L` | Toggle AutoFilter |
| `Alt + ↓` | Open filter dropdown in filtered column |
| `Ctrl + Shift + F3` | Create names from selection |
| `F11` | Create chart from selection (new sheet) |
| `Alt + F1` | Create chart from selection (same sheet) |
| `Ctrl + K` | Insert hyperlink |
| `Alt + D + S` | Sort dialog (legacy) |
| `Alt + A + S + S` | Sort dialog (ribbon) |
| `Alt + D + F + F` | Toggle AutoFilter (legacy) |
| `Ctrl + Shift + M` | Insert Power Query data (M365) |

---

## GENERAL / WORKBOOK SHORTCUTS

| Shortcut | Action |
|----------|--------|
| `Ctrl + N` | New workbook |
| `Ctrl + O` | Open workbook |
| `Ctrl + S` | Save |
| `F12` | Save As |
| `Ctrl + W` | Close workbook |
| `Ctrl + P` | Print |
| `Ctrl + F2` | Print Preview |
| `F1` | Help |
| `F7` | Spell Check |
| `F10` | Activate menu bar (ribbon shortcuts) |
| `Ctrl + F1` | Toggle ribbon visibility |
| `Alt` | Show ribbon shortcut keys |
| `Ctrl + Shift + J` | Fill selection with value from cell above (Ctrl+D alternative) |

---

## 📋 Pro Tips

- **`Alt` key sequences**: Press `Alt`, then follow the letter badges that appear on the ribbon to navigate menus (e.g., `Alt → H → B → A` = Add All Borders)
- **Double-click fill handle**: Auto-fills formulas/data down to match adjacent column
- **`Ctrl+Shift+V`**: Paste Special (in newer Excel / M365)
- **Quick Access Toolbar**: Add frequently used commands and access them with `Alt+1`, `Alt+2`, etc.
- **`Ctrl+D`** with multiple cells selected: copies top row down through the entire selection


## FILE: cheat-sheets/pivottable-cheat-sheet.md

# Excel PivotTable Cheat Sheet - Quick Reference

> **Printable Reference** | Print on A4/Letter | 1-2 pages

---

## CREATING A PIVOTTABLE

| Step | Action |
|------|--------|
| 1 | Click anywhere in your data range |
| 2 | **Insert → PivotTable** (or `Alt + N + V`) |
| 3 | Confirm data range and choose location (New/Existing Worksheet) |
| 4 | Click **OK** - empty PivotTable + Field List appears |

> **Prerequisite**: Data must have headers in row 1, no blank rows/columns, no merged cells.

---

## PIVOTTABLE FIELD LIST LAYOUT

The 4 field areas determine how data is displayed:

| Area | Position | What Goes Here | Example |
|------|----------|---------------|---------|
| **Filters** | Top of report | Field to filter entire table | Year, Region |
| **Columns** | Across top | Field as column headers | Quarter, Product |
| **Rows** | Down left side | Field as row labels | Department, Salesperson |
| **Values** | Center (data) | Field to calculate (sum, count, etc.) | Revenue, Units |

### Drag & Drop Methods
- **Drag** fields from field list to areas
- **Check** a field → Excel auto-assigns (text → Rows, numbers → Values)
- **Right-click** field → Move to → choose area

---

## VALUE FIELD SETTINGS

Right-click any value → **Value Field Settings**:

| Calculation | Use When | Example |
|------------|----------|---------|
| **Sum** | Totalling numeric data | Total revenue |
| **Count** | Counting entries | Number of orders |
| **Average** | Finding mean values | Average order value |
| **Max** | Finding highest value | Highest sale |
| **Min** | Finding lowest value | Lowest price |
| **Count Distinct** *(M365)* | Counting unique values | Unique customers |
| **% of Grand Total** | Showing proportion | % of total sales |
| **% of Column Total** | Column-wise percentage | % within each quarter |
| **Running Total** | Cumulative sum | Year-to-date total |
| **Difference From** | Comparing to a base | Change from prior year |
| **Index** | Relative importance | Cross-tabulation weight |

### Changing Calculation
1. Right-click value cell → **Summarize Values By** → select
2. Or: **Value Field Settings → Summarize Values By** tab

### Number Formatting
- Right-click value → **Number Format** (not Format Cells!)
- This ensures formatting persists after pivot refresh

---

## GROUPING DATA

### Group Dates
1. Right-click any date in Rows/Columns
2. Select **Group**
3. Choose: **Years, Quarters, Months, Days, Hours, Minutes**
4. Can select multiple levels (e.g., Years + Months)

### Group Numbers
1. Right-click any number in Rows
2. Select **Group**
3. Set **Starting at**, **Ending at**, and **By** (interval)

### Group Text (Manual)
1. Select items to group (hold `Ctrl`)
2. Right-click → **Group**
3. Rename the group as needed

### Ungroup
Right-click grouped items → **Ungroup**

---

## CALCULATED FIELDS & ITEMS

### Calculated Field (new field using existing fields)
1. Click PivotTable → **PivotTable Analyze → Fields, Items & Sets → Calculated Field**
2. Name: e.g., `Profit`
3. Formula: `=Revenue - Cost`
4. Click **Add → OK**

| Example Calculated Fields | Formula |
|--------------------------|---------|
| Profit Margin | `=Profit / Revenue` |
| Tax | `=Revenue * 0.08` |
| Avg per Unit | `=Revenue / Units` |

### Calculated Item (new item within an existing field)
1. Click a field label → **Fields, Items & Sets → Calculated Item**
2. Name: e.g., `Total Q1+Q2`
3. Formula: `= Q1 + Q2`

> ⚠️ Calculated items can cause issues with grouping. Use with caution.

---

## SLICERS (Visual Filters)

### Add a Slicer
1. Click PivotTable → **PivotTable Analyze → Insert Slicer**
2. Check fields you want as slicers → **OK**
3. Drag slicers to position them on the worksheet

### Using Slicers
| Action | How |
|--------|-----|
| Select single item | Click the item |
| Select multiple items | `Ctrl + Click` items |
| Clear filter | Click the **clear filter** icon (⊘) in slicer header |
| Multi-select mode | Click the **multi-select** icon (☐) in header |

### Connect Slicer to Multiple PivotTables
1. Right-click slicer → **Report Connections** (or **PivotTable Connections**)
2. Check all PivotTables that should share this slicer

### Slicer Formatting
- **Slicer tab** → choose styles, columns, button size
- Right-click → **Slicer Settings** for sorting and hiding items

---

## REFRESHING PIVOTTABLES

| Action | How |
|--------|-----|
| Refresh one PivotTable | Right-click → **Refresh** |
| Refresh all PivotTables | **PivotTable Analyze → Refresh → Refresh All** |
| Auto-refresh on open | **PivotTable Analyze → Options → Data → Refresh data when opening the file** |
| Change data source | **PivotTable Analyze → Change Data Source** |
| Use Tables as source | Convert data to **Table** (`Ctrl+T`) first - auto-expands on refresh |

---

## PIVOTTABLE OPTIONS & TIPS

| Setting | Where | What It Does |
|---------|-------|-------------|
| Classic layout | Options → Display → **Classic PivotTable layout** | Drag fields directly into grid |
| Repeat item labels | Design → Report Layout → **Repeat All Item Labels** | Fills blank row labels |
| Remove subtotals | Design → Subtotals → **Do Not Show Subtotals** | Cleaner look |
| Show in Tabular Form | Design → Report Layout → **Show in Tabular Form** | One column per field |
| Conditional formatting | Home → Conditional Formatting (on pivot cells) | Visual cues on values |
| Drill down | **Double-click** a value cell | Creates new sheet with underlying data |

---

## 📋 Quick Workflow Tips

1. **Start with clean data** → convert to Table (`Ctrl+T`) first
2. **Rows** = categories going down, **Columns** = categories going across
3. **Right-click** is your best friend - nearly every setting is accessible from context menu
4. **Number Format** via right-click pivot value (not Format Cells) - survives refresh
5. **Group dates** immediately after adding - saves dozens of individual date entries
6. **Slicers > Filters** - more visual and easier for end users
7. **Double-click any number** in a pivot to see the raw data behind it
8. **PivotTable Analyze → Options → Totals & Filters** → uncheck "Show grand totals" to remove


## FILE: cheat-sheets/power-query-cheat-sheet.md

# Power Query Cheat Sheet - M Language & Transformations

> **Quick Reference** | Print on A4/Letter | 1-2 pages

---

## GETTING STARTED

| Action | How |
|--------|-----|
| Open Power Query Editor | **Data → Get Data → Launch Power Query Editor** |
| Import from file | **Data → Get Data → From File** (CSV, Excel, JSON, XML, etc.) |
| Import from database | **Data → Get Data → From Database** (SQL Server, Access, etc.) |
| Import from web | **Data → Get Data → From Web** |
| Append queries | **Home → Append Queries** (stack tables vertically) |
| Merge queries | **Home → Merge Queries** (join tables like VLOOKUP) |
| Close & Load | **Home → Close & Load** (returns data to Excel) |

---

## TOP 20 M LANGUAGE FUNCTIONS

### Table Functions

| Function | Syntax | Description | Example |
|----------|--------|-------------|---------|
| `Table.SelectRows` | `Table.SelectRows(table, condition)` | Filter rows by condition | `= Table.SelectRows(Source, each [Status] = "Active")` |
| `Table.SelectColumns` | `Table.SelectColumns(table, columns)` | Keep only listed columns | `= Table.SelectRows(Source, {"Name","Sales"})` |
| `Table.RemoveColumns` | `Table.RemoveColumns(table, columns)` | Remove listed columns | `= Table.RemoveColumns(Source, {"Temp","Debug"})` |
| `Table.AddColumn` | `Table.AddColumn(table, name, function)` | Add computed column | `= Table.AddColumn(Source, "Tax", each [Price] * 0.08)` |
| `Table.RenameColumns` | `Table.RenameColumns(table, list)` | Rename columns | `= Table.RenameColumns(Source, {{"old","new"}})` |
| `Table.Sort` | `Table.Sort(table, ordering)` | Sort by columns | `= Table.Sort(Source, {{"Sales", Order.Descending}})` |
| `Table.Group` | `Table.Group(table, key, agg)` | Group and aggregate | `= Table.Group(Source, {"Region"}, {{"Total", each List.Sum([Sales]), type number}})` |
| `Table.TransformColumnTypes` | `Table.TransformColumnTypes(table, types)` | Set column data types | `= Table.TransformColumnTypes(Source, {{"Date", type date}})` |
| `Table.NestedJoin` | `Table.NestedJoin(t1, key1, t2, key2, name)` | Left join two tables | `= Table.NestedJoin(Orders, {"ID"}, Customers, {"ID"}, "Customer")` |
| `Table.Combine` | `Table.Combine({table1, table2, …})` | Append/stack tables | `= Table.Combine({Q1, Q2, Q3, Q4})` |

### List Functions

| Function | Syntax | Description | Example |
|----------|--------|-------------|---------|
| `List.Sum` | `List.Sum(list)` | Sum of list values | `= List.Sum({1, 2, 3})` → 6 |
| `List.Average` | `List.Average(list)` | Average of list | `= List.Average({10, 20, 30})` → 20 |
| `List.Count` | `List.Count(list)` | Count items | `= List.Count({1, 2, 3})` → 3 |
| `List.Min` / `List.Max` | `List.Min(list)` | Minimum / Maximum value | `= List.Max({5, 2, 8})` → 8 |
| `List.Distinct` | `List.Distinct(list)` | Remove duplicates | `= List.Distinct({"A","B","A"})` → {"A","B"} |

### Text Functions

| Function | Syntax | Description | Example |
|----------|--------|-------------|---------|
| `Text.Length` | `Text.Length(text)` | Character count | `= Text.Length("Hello")` → 5 |
| `Text.Split` | `Text.Split(text, delimiter)` | Split text into list | `= Text.Split("a-b-c", "-")` → {"a","b","c"} |
| `Text.Combine` | `Text.Combine(list, delimiter)` | Join list into text | `= Text.Combine({"a","b"}, ",")` → "a,b" |
| `Text.Lower` / `Upper` | `Text.Lower(text)` | Change case | `= Text.Upper("hello")` → "HELLO" |
| `Text.Trim` | `Text.Trim(text)` | Remove leading/trailing spaces | `= Text.Trim(" hi ")` → "hi" |

---

## COMMON TRANSFORMATIONS (GUI → M Code)

### Filter Rows
```
// Via GUI: Click column header filter dropdown → select values
// M Code generated:
= Table.SelectRows(Source, each [Region] = "East")

// Multiple conditions:
= Table.SelectRows(Source, each [Region] = "East" and [Sales] > 1000)

// Contains text:
= Table.SelectRows(Source, each Text.Contains([Name], "Corp"))

// Not null:
= Table.SelectRows(Source, each [Email] <> null)
```

### Add Conditional Column
```
// GUI: Add Column → Conditional Column
// M Code:
= Table.AddColumn(Source, "Tier", each
    if [Sales] >= 10000 then "Gold"
    else if [Sales] >= 5000 then "Silver"
    else "Bronze"
)
```

### Merge (Join) Tables
```
// GUI: Home → Merge Queries → select key columns → Join kind
// M Code (Left Outer Join):
= Table.NestedJoin(Orders, {"CustomerID"}, Customers, {"CustomerID"}, "Customer", JoinKind.LeftOuter)

// Join types: LeftOuter, RightOuter, FullOuter, Inner, LeftAnti, RightAnti
```

### Pivot / Unpivot
```
// Unpivot (wide → long): Select columns → Transform → Unpivot Columns
= Table.UnpivotOtherColumns(Source, {"ID","Name"}, "Attribute", "Value")

// Pivot (long → wide): Select column → Transform → Pivot Column
= Table.Pivot(Source, List.Distinct(Source[Month]), "Month", "Sales", List.Sum)
```

### Replace Values
```
// GUI: Right-click column → Replace Values
= Table.ReplaceValue(Source, "old", "new", Replacer.ReplaceText, {"ColumnName"})

// Replace nulls:
= Table.ReplaceValue(Source, null, 0, Replacer.ReplaceValue, {"Sales"})
```

### Custom Column with Text Extraction
```
// Extract domain from email:
= Table.AddColumn(Source, "Domain", each Text.AfterDelimiter([Email], "@"))

// Extract first 3 characters:
= Table.AddColumn(Source, "Code", each Text.Start([ID], 3))

// Parse JSON column:
= Table.AddColumn(Source, "Parsed", each Json.Document([JsonCol]))
```

---

## POWER QUERY PARAMETERS & VARIABLES

### Define a Parameter
1. **Home → Manage Parameters → New Parameter**
2. Set Name, Type, Default Value
3. Reference in code: `= Table.SelectRows(Source, each [Date] >= Parameter1)`

### Reference Previous Steps
Each step is a variable. Reference by name:
```
let
    Source = Excel.Workbook(File.Contents("data.xlsx")),
    Sheet1 = Source{[Name="Sheet1"]}[Data],
    Promoted = Table.PromoteHeaders(Sheet1),
    Filtered = Table.SelectRows(Promoted, each [Status] = "Active")
in
    Filtered
```

---

## ERROR HANDLING

| Pattern | M Code |
|---------|--------|
| Replace errors with value | `= Table.ReplaceErrorValues(Source, {"Col", 0})` |
| Try/catch | `= try [risky_operation] otherwise "default"` |
| Remove error rows | `= Table.RemoveRowsWithErrors(Source, {"Col"})` |
| Check for null | `= if [Value] = null then 0 else [Value]` |

---

## PERFORMANCE TIPS

| Tip | Why |
|-----|-----|
| **Filter early** | Remove rows/columns as first steps - reduces data flowing through pipeline |
| **Disable auto data type detection** | File → Options → uncheck "Auto-detect column types" (prevents auto-changes) |
| **Use `Table.Buffer`** | `= Table.Buffer(Source)` - caches result in memory to avoid re-evaluation |
| **Fold queries** | Ensure filters push to source (check via View → Query Dependencies → right-click → View native query) |
| **Avoid row-by-row** | Prefer column-level transforms over `Table.AddColumn` with expensive per-row logic |
| **Reference, don't duplicate** | Base new queries on existing steps rather than re-loading source data |

---

## 📋 Quick Tips

- **`fx` button** in formula bar: Insert step → write M code
- **Applied Steps pane** (right): Click any step to preview at that point
- **Advanced Editor**: Home → Advanced Editor - see/edit full M code
- **Duplicate query**: Right-click query → Duplicate (for variations without reloading)
- **`Ctrl+Click`** column headers to select multiple columns for batch transforms
- All M functions are **case-sensitive**: `Table.Group` ≠ `table.group`
- Use `each` shorthand: `each [Column] > 10` is shorthand for `(_ as record) => _[Column] > 10`


## FILE: quizzes/quiz-01-fundamentals.md

# Quiz 01 - Excel Fundamentals

**Module:** Interface, Navigation, Cell References, Data Entry, File Formats
**Total Questions:** 20 (15 Multiple Choice + 5 Formula Challenges)
**Passing Score:** 90%+ = Ready to advance | 70–89% = Review weak areas | <70% = Re-study module

---

## Part A: Multiple Choice (3 points each)

**1.** You need to select cells A1 through D10 quickly. What is the most efficient method?
- A) Click A1, hold Ctrl, click D10
- B) Click A1, hold Shift, click D10
- C) Type `A1:D10` in the Name Box and press Enter
- D) Both B and C work

**2.** You type `3/15` into a cell and press Enter. By default, Excel will:
- A) Display the text "3/15"
- B) Display the date March 15 of the current year
- C) Display the fraction 3/15 reduced to 1/5
- D) Show an error

**3.** Which cell reference, when copied from B2 to D5, still points to the original cell?
- A) `A1`
- B) `$A$1`
- C) `A$1`
- D) `$A1`

**4.** You press `Ctrl + ;` in a cell. What gets entered?
- A) Current time
- B) Current date
- C) The value from the cell above
- D) A formula

**5.** A column shows `#####` in every cell. What is the most likely fix?
- A) The formula has an error
- B) The column is too narrow to display the value
- C) The cells contain circular references
- D) The sheet is protected

**6.** Which file format preserves ALL Excel features including macros and Power Query connections?
- A) `.xlsx`
- B) `.csv`
- C) `.xlsm`
- D) `.xls`

**7.** You want to enter the same value into cells A1, B3, and C5 simultaneously. How?
- A) Type in A1, copy, paste to B3 and C5
- B) Select A1, hold Ctrl, click B3 and C5, type the value, press Ctrl+Enter
- C) It's not possible to enter into non-contiguous cells at once
- D) Use Fill > Series

**8.** What does pressing `Tab` after entering data in a cell do?
- A) Inserts a tab character into the cell
- B) Moves the selection one cell to the right
- C) Moves the selection one cell down
- D) Confirms the entry and stays on the same cell

**9.** Which shortcut jumps directly to the last used cell in a worksheet?
- A) `Ctrl + End`
- B) `Ctrl + Home`
- C) `Ctrl + ↓`
- D) Both A and C, but they behave differently

**10.** You enter `00125` into a cell formatted as General. What displays?
- A) `00125`
- B) `125`
- C) `0.00125`
- D) An error message

**11.** Which of the following is a mixed cell reference?
- A) `A1`
- B) `$A$1`
- C) `$A1`
- D) `Sheet1!A1`

**12.** You want to prevent accidental edits to a specific sheet while allowing edits on others. You should:
- A) Save the file as read-only
- B) Protect that specific sheet with a password
- C) Hide the sheet
- D) Move the sheet to a new workbook

**13.** In Excel's Name Box (left of the formula bar), typing `F15` and pressing Enter will:
- A) Insert the text "F15" into the current cell
- B) Move the cursor to cell F15
- C) Create a named range called F15
- D) Open the Find dialog

**14.** You paste data from a website and it appears in a single cell with line breaks. To split each line into its own row, you should use:
- A) Text to Columns with a space delimiter
- B) Flash Fill
- C) Find & Replace, replacing `Ctrl+J` with a unique delimiter, then Text to Columns
- D) It's impossible to split in-cell line breaks

**15.** Which of these is NOT a valid Excel data type when entering into a cell?
- A) `TRUE`
- B) `#N/A`
- C) `=SUM(A1:A5)`
- D) `12:00 AM`

---

## Part B: Fill-in-the-Formula Challenge (5 points each)

Write the formula or describe the action that produces the required result.

**16.** In cell C1, write a formula that adds the values in A1 and B1 using the `+` operator, but only works correctly when copied across columns (both references should shift). What is the formula?

> **Your answer:** `_________________________________`

**17.** You have data in A1:A100. Write a formula in B1 that, when copied down, shows the running row number starting at 1.

> **Your answer:** `_________________________________`

**18.** Write the formula to display the text "High" if cell A1 is greater than 100, otherwise display "Low", using the IF function.

> **Your answer:** `_________________________________`

**19.** You need to reference cell B5 on a sheet named "Q1 Data" from another sheet. Write the complete cell reference.

> **Your answer:** `_________________________________`

**20.** What Excel feature (not a formula) lets you auto-fill a pattern like "Jan, Feb, Mar" after typing just the first entry?

> **Your answer:** `_________________________________`

---

## Answer Key

| # | Answer | Explanation |
|---|--------|-------------|
| 1 | **D** | Both Shift+click and Name Box entry select a range. |
| 2 | **B** | Excel auto-converts short date entries to the current year. |
| 3 | **B** | `$A$1` is absolute - it doesn't shift when copied. |
| 4 | **B** | `Ctrl+;` inserts today's date; `Ctrl+Shift+;` inserts time. |
| 5 | **B** | `#####` means the column is too narrow - widen it. |
| 6 | **C** | `.xlsm` is the macro-enabled format that preserves VBA. |
| 7 | **B** | Ctrl+click for non-contiguous selection, then Ctrl+Enter fills all. |
| 8 | **B** | Tab moves right after data entry; Enter moves down. |
| 9 | **D** | `Ctrl+End` goes to last used cell; `Ctrl+↓` goes to last filled cell in column. |
| 10 | **B** | General format drops leading zeros. Use Text or custom format `00000`. |
| 11 | **C** | `$A1` locks the column but not the row - that's mixed. |
| 12 | **B** | Sheet protection allows granular control per sheet. |
| 13 | **B** | The Name Box navigates to any cell or named range. |
| 14 | **C** | Ctrl+J represents a line break in Find & Replace; use it to create a delimiter, then split. |
| 15 | **B** | `#N/A` is an error result, not something you type as data. |
| 16 | `=A1+B1` | With relative references, both shift when copied across columns. |
| 17 | `=ROW()-ROW($A$1)+1` or `=ROW(A1)` | ROW(A1) returns 1 and increments as copied down. |
| 18 | `=IF(A1>100,"High","Low")` | Standard IF with text outputs. |
| 19 | `='Q1 Data'!B5` | Sheet names with spaces need single quotes. |
| 20 | **AutoFill (Fill Handle)** | Drag the fill handle or use Data > Series for pattern recognition. |

---

**Scoring:**
- Part A: 15 questions × 3 points = 45 points
- Part B: 5 questions × 5 points = 25 points
- **Total: 70 points**
- 63+ (90%): ✅ Ready for Module 2
- 49–62 (70–89%): 📖 Review weak areas
- Below 49 (<70%): 🔄 Re-study Module 1


## FILE: quizzes/quiz-02-formulas.md

# Quiz 02 - Formulas & Functions

**Module:** Formulas, SUM/IF/VLOOKUP, Text/Date Functions, Error Handling
**Total Questions:** 20 (15 Multiple Choice + 5 Formula Challenges)
**Passing Score:** 90%+ = Ready to advance | 70–89% = Review weak areas | <70% = Re-study module

---

## Part A: Multiple Choice (3 points each)

**1.** Cell A1 contains `500`, B1 contains `200`. What does `=A1/B1` return?
- A) `2.5`
- B) `2`
- C) `2.50` (always two decimals)
- D) Depends on cell formatting

**2.** What does `=IFERROR(VLOOKUP("X",A:B,2,FALSE),"Not Found")` return when "X" doesn't exist in column A?
- A) `#N/A`
- B) `0`
- C) `"Not Found"`
- D) Blank cell

**3.** You want to sum all values in column B where column A equals "East". Which formula is correct?
- A) `=SUMIF(A:A,"East",B:B)`
- B) `=SUMIFS(B:B,A:A,"East")`
- C) `=SUM(IF(A:A="East",B:B))`
- D) All of the above

**4.** Cell A1 contains `"Hello World"`. What does `=MID(A1,7,5)` return?
- A) `"Hello"`
- B) `"World"`
- C) `" Worl"`
- D) `"World "`

**5.** What is the result of `=LEN(TRIM(" Excel "))`?
- A) `7`
- B) `5`
- C) `9`
- D) `8`

**6.** You have `=VLOOKUP(A1,Sheet2!A:D,3,FALSE)`. What does the `3` represent?
- A) The third row to return
- B) The third column of the lookup table to return
- C) Three criteria to match
- D) A three-column range

**7.** Which formula correctly extracts the last 3 characters from cell A1?
- A) `=RIGHT(A1,3)`
- B) `=MID(A1,LEN(A1)-2,3)`
- C) `=LEFT(A1,LEN(A1)-3)`
- D) Both A and B

**8.** `=TODAY()` returns `2025-01-15`. What does `=YEAR(TODAY())+1` return?
- A) `2025`
- B) `2026`
- C) `2026-01-15`
- D) `#VALUE!`

**9.** You see `#REF!` in a cell. What most likely happened?
- A) A formula divides by zero
- B) A referenced cell or range was deleted
- C) The lookup value wasn't found
- D) A circular reference exists

**10.** What does `=COUNTA(A1:A10)` return if A1:A10 contains 7 numbers, 2 text entries, and 1 blank?
- A) `10`
- B) `9`
- C) `7`
- D) `8`

**11.** You want to concatenate A1 and B1 with a space between them. Which works?
- A) `=A1&" "&B1`
- B) `=CONCATENATE(A1," ",B1)`
- C) `=TEXTJOIN(" ",TRUE,A1,B1)`
- D) All of the above

**12.** What does `=ROUND(3.456, 2)` return?
- A) `3.45`
- B) `3.46`
- C) `3.5`
- D) `3.4`

**13.** You want to find the position of the `@` in an email address in A1. Which formula?
- A) `=SEARCH("@",A1)`
- B) `=FIND("@",A1)`
- C) `=MATCH("@",A1,0)`
- D) Both A and B

**14.** `=DATE(2025,14,1)` returns what?
- A) `#VALUE!`
- B) `February 1, 2026`
- C) `January 14, 2025`
- D) `January 1, 2025`

**15.** You have nested IFs returning `#N/A`. Which function lets you provide a fallback value instead?
- A) `=ISNA()`
- B) `=IFERROR()`
- C) `=IFNA()`
- D) Both B and C

---

## Part B: Fill-in-the-Formula Challenge (5 points each)

Write the formula that produces the required result.

**16.** Column A has product names and column B has prices. Write a formula that returns the price for "Widget" using VLOOKUP. The product name column is A and price is in column B.

> **Your answer:** `_________________________________`

**17.** Write a formula that counts how many cells in A1:A100 contain values greater than 50 AND less than 100.

> **Your answer:** `_________________________________`

**18.** Cell A1 contains a full name like `"John Smith"`. Write a formula to extract just the first name.

> **Your answer:** `_________________________________`

**19.** Write a formula that converts a date in A1 to the format `"DD-MMM-YYYY"` (e.g., `15-Jan-2025`).

> **Your answer:** `_________________________________`

**20.** Write a formula using SUMPRODUCT that multiplies quantities in B1:B10 by unit prices in C1:C10 and returns the total.

> **Your answer:** `_________________________________`

---

## Answer Key

| # | Answer | Explanation |
|---|--------|-------------|
| 1 | **D** | Display depends on cell format; the stored value is 2.5 but may show as `2`, `2.50`, etc. |
| 2 | **C** | IFERROR catches #N/A and returns the specified fallback text. |
| 3 | **D** | SUMIF, SUMIFS, and array SUM(IF(...)) all work. SUMIFS is preferred for new workbooks. |
| 4 | **B** | `MID(A1,7,5)` starts at position 7 and grabs 5 chars: "World". |
| 5 | **A** | TRIM removes leading/trailing spaces, leaving "Excel" (5 chars). |
| 6 | **B** | The third argument of VLOOKUP is the column index within the table array. |
| 7 | **D** | `RIGHT(A1,3)` is simplest; `MID` approach also works. |
| 8 | **B** | YEAR extracts 2025, adds 1 → 2026. |
| 9 | **B** | #REF! means a cell/range reference became invalid (usually deleted). |
| 10 | **B** | COUNTA counts non-blank cells: 7 numbers + 2 text = 9. |
| 11 | **D** | All three concatenation methods work. |
| 12 | **B** | ROUND(3.456, 2) rounds to 2 decimal places → 3.46. |
| 13 | **D** | Both SEARCH and FIND locate "@" - FIND is case-sensitive, SEARCH supports wildcards. |
| 14 | **B** | Excel rolls over month 14 to February of the next year. |
| 15 | **D** | IFERROR catches all errors; IFNA catches only #N/A. Both can provide fallback values. |
| 16 | `=VLOOKUP("Widget",A:B,2,FALSE)` | Exact match lookup on "Widget", returns column 2 (price). |
| 17 | `=COUNTIFS(A1:A100,">"&50,A1:A100,"<"&100)` | COUNTIFS with two criteria. |
| 18 | `=LEFT(A1,FIND(" ",A1)-1)` | Finds the space, takes everything left of it. |
| 19 | `=TEXT(A1,"DD-MMM-YYYY")` | TEXT function with custom format string. |
| 20 | `=SUMPRODUCT(B1:B10,C1:C10)` | Element-wise multiplication then sum. |

---

**Scoring:**
- Part A: 15 questions × 3 points = 45 points
- Part B: 5 questions × 5 points = 25 points
- **Total: 70 points**
- 63+ (90%): ✅ Ready for Module 3
- 49–62 (70–89%): 📖 Review weak areas
- Below 49 (<70%): 🔄 Re-study Module 2


## FILE: quizzes/quiz-03-visualization.md

# Quiz 03 - Data Visualization & Formatting

**Module:** Conditional Formatting, Charts, Tables, Sparklines
**Total Questions:** 20 (15 Multiple Choice + 5 Formula Challenges)
**Passing Score:** 90%+ = Ready to advance | 70–89% = Review weak areas | <70% = Re-study module

---

## Part A: Multiple Choice (3 points each)

**1.** You apply a conditional formatting rule with the formula `=B2>100` to range B2:B50. What happens?
- A) Only B2 is highlighted if it's over 100
- B) Each cell in B2:B50 is highlighted independently if its value exceeds 100
- C) All cells highlight because B2 is checked against 100
- D) An error - formulas can't be used in conditional formatting

**2.** You convert a data range to an Excel Table (Ctrl+T). Which of these is NOT an automatic benefit?
- A) Auto-filling formulas in new columns
- B) Banded rows for readability
- C) Automatic PivotTable creation
- D) Auto-expanding range references

**3.** Which chart type best shows the part-to-whole relationship of 5 budget categories?
- A) Line chart
- B) Bar chart
- C) Pie chart
- D) Scatter chart

**4.** You want to highlight the top 10% of values in a column. Which conditional formatting preset should you use?
- A) Color Scales
- B) Icon Sets
- C) Top/Bottom Rules > Top 10%
- D) Data Bars

**5.** What happens when you add a new row below an Excel Table?
- A) Nothing - you need to manually extend the table
- B) The table auto-expands to include the new row
- C) A dialog asks if you want to extend
- D) The new row stays outside the table

**6.** You create a line chart but the x-axis labels show numbers instead of your month names. What's the fix?
- A) The months are in the wrong row - move them above the data
- B) Right-click the chart > Select Data > Edit Horizontal Axis Labels
- C) Change the chart type to a bar chart
- D) Format cells as Text

**7.** A sparkline in cell E2 shows data from A2:D2. You insert a new column at C. What happens to the sparkline?
- A) It now shows A2:E2
- B) It breaks and shows an error
- C) It still shows A2:D2 (now A2:B2 + D2:E2)
- D) It shows A2:C2 only

**8.** You want conditional formatting that shows a gradient from red (low) to green (high) across a range. Use:
- A) Data Bars
- B) Color Scales
- C) Icon Sets
- D) A custom formula with RGB

**9.** Which is TRUE about structured references in Excel Tables?
- A) `=SUM(Table1[Sales])` sums the Sales column
- B) `=Table1[@Revenue]` refers to the Revenue cell in the current row
- C) `=Table1[[#Totals],[Profit]]` refers to the Profit total row
- D) All of the above

**10.** You have a combo chart with bars on the primary axis and a line on the secondary axis. When is this useful?
- A) When you want two chart types to look fancy
- B) When comparing data with very different scales (e.g., revenue in millions vs. growth rate %)
- C) When you have more than 255 data points
- D) When you want to animate the chart

**11.** What does "Clear Rules > Clear Rules from Entire Sheet" do?
- A) Deletes all data from the sheet
- B) Removes all conditional formatting rules from the sheet
- C) Clears the cell contents of formatted cells
- D) Removes all chart formatting

**12.** You want a conditional formatting rule that highlights an entire row when column C contains "Overdue". What formula should you use?
- A) `=C2="Overdue"`
- B) `=$C2="Overdue"`
- C) `=C$2="Overdue"`
- D) `=$C$2="Overdue"`

**13.** Which sparkline type shows trends across a row of monthly data?
- A) Win/Loss
- B) Line
- C) Column
- D) Both B and C

**14.** You want to show data labels on a pie chart that display both the category name and percentage. How?
- A) Right-click data labels > Format Data Labels > check Category Name and Percentage
- B) Manually type each label
- C) This is not possible in Excel
- D) Use a text box linked to cells

**15.** An Excel Table has a Total Row enabled. What function does it default to for a column containing dates?
- A) SUM
- B) COUNT
- C) AVERAGE
- D) MAX

---

## Part B: Fill-in-the-Formula Challenge (5 points each)

Write the formula or describe the action needed.

**16.** You want conditional formatting that highlights cells in A1:A100 that contain the text "URGENT" (case-insensitive). Write the formula.

> **Your answer:** `_________________________________`

**17.** Write a conditional formatting formula that highlights every other row in A1:F100 (creating a zebra-stripe effect).

> **Your answer:** `_________________________________`

**18.** You have a Table named `SalesData` with columns `[Region]`, `[Product]`, and `[Amount]`. Write a formula that returns the total Amount where Region is "North".

> **Your answer:** `_________________________________`

**19.** Describe the steps to create a PivotChart from an existing Excel Table named `Orders`.

> **Your answer:** `_________________________________`

**20.** Write the structured reference formula to calculate the average of the `Price` column in a Table named `Inventory`.

> **Your answer:** `_________________________________`

---

## Answer Key

| # | Answer | Explanation |
|---|--------|-------------|
| 1 | **B** | Relative reference `B2` adjusts per row - each cell is evaluated independently. |
| 2 | **C** | Tables don't auto-create PivotTables; you must insert one manually. |
| 3 | **C** | Pie charts are designed for part-to-whole with a small number of categories. |
| 4 | **C** | Top/Bottom Rules have a "Top 10%" preset. |
| 5 | **B** | Tables auto-expand when you type in the row directly below. |
| 6 | **B** | Select Data > Edit Horizontal Axis Labels lets you specify the label range. |
| 7 | **A** | Sparklines auto-adjust their data range when columns are inserted. |
| 8 | **B** | Color Scales apply gradient fills based on cell values. |
| 9 | **D** | All three are valid structured reference syntaxes. |
| 10 | **B** | Combo charts with dual axes handle data with vastly different scales. |
| 11 | **B** | "Clear Rules" only removes conditional formatting - not data or charts. |
| 12 | **B** | `$C2` locks the column (C) but lets the row change - highlights the entire row. |
| 13 | **D** | Both Line and Column sparklines show trends; Win/Loss shows positive/negative only. |
| 14 | **A** | Format Data Labels has checkboxes for Category Name, Percentage, Value, etc. |
| 15 | **D** | Date columns default to MAX in the Total Row. |
| 16 | `=SEARCH("URGENT",A1)>0` or `=ISNUMBER(SEARCH("URGENT",A1))` | SEARCH is case-insensitive; ISNUMBER converts the result to TRUE/FALSE. |
| 17 | `=MOD(ROW(),2)=0` | Even-numbered rows get the format; use `=1` for odd rows. |
| 18 | `=SUMIF(SalesData[Region],"North",SalesData[Amount])` | SUMIF with structured references. |
| 19 | Click anywhere in the `Orders` Table → Insert tab → PivotChart → choose location → configure fields. | - |
| 20 | `=AVERAGE(Inventory[Price])` | Structured reference for the Price column average. |

---

**Scoring:**
- Part A: 15 questions × 3 points = 45 points
- Part B: 5 questions × 5 points = 25 points
- **Total: 70 points**
- 63+ (90%): ✅ Ready for Module 4
- 49–62 (70–89%): 📖 Review weak areas
- Below 49 (<70%): 🔄 Re-study Module 3


## FILE: quizzes/quiz-04-analysis.md

# Quiz 04 - Data Analysis & PivotTables

**Module:** Sorting, Filtering, PivotTables, What-If Analysis, Data Validation
**Total Questions:** 20 (15 Multiple Choice + 5 Formula Challenges)
**Passing Score:** 90%+ = Ready to advance | 70–89% = Review weak areas | <70% = Re-study module

---

## Part A: Multiple Choice (3 points each)

**1.** You have a dataset with columns: Date, Region, Product, Sales. You want to see total sales per region per product. The fastest approach is:
- A) Use SUMIF for each region-product combination
- B) Create a PivotTable with Region and Product as rows, Sales as values (SUM)
- C) Sort by Region, then manually add subtotals
- D) Create a separate sheet for each region

**2.** You apply a filter to column A showing only "East" and "North". What happens to the row numbers?
- A) They stay the same with gaps
- B) Hidden rows are deleted
- C) They become sequential (1, 2, 3...)
- D) Hidden rows turn gray but numbers stay

**3.** In a PivotTable, you drag "Region" to Rows and "Sales" to Values. Excel defaults the value calculation to:
- A) SUM of Sales
- B) COUNT of Sales
- C) AVERAGE of Sales
- D) It shows a dialog asking you to choose

**4.** You want to restrict a cell so users can only enter numbers between 1 and 100. You should use:
- A) Conditional Formatting
- B) Data Validation with Whole Number between 1 and 100
- C) A formula in the cell
- D) Sheet protection

**5.** What-If Analysis's Goal Seek finds:
- A) The input value needed to achieve a desired output
- B) The maximum possible output
- C) All scenarios for a formula
- D) Errors in your formulas

**6.** You sort a dataset by Date (oldest to newest), then by Region (A to Z). The result:
- A) Sorts by Region first, then Date
- B) Sorts by Date first, then Region within each date
- C) Sorts alphabetically by Region only
- D) Sorts by Date only

**7.** A PivotTable shows quarterly sales. You want to display each quarter as a percentage of the year's total. How?
- A) Add a calculated field with `=Sales/TotalSales`
- B) Right-click a value > Show Values As > % of Grand Total
- C) Manually calculate percentages in a new column
- D) Change the number format to percentage

**8.** You create a custom error message in Data Validation. When a user enters an invalid value, they:
- A) See a red triangle in the cell corner
- B) Get a pop-up error dialog with your message
- C) The cell turns red
- D) The value is silently deleted

**9.** Which Scenario Manager feature lets you compare multiple what-if scenarios side by side?
- A) Scenario Summary
- B) Data Table
- C) Goal Seek
- D) Solver

**10.** You want to filter a PivotTable to show only the top 5 products by sales. How?
- A) Sort descending and delete the bottom rows
- B) Click the Row Labels dropdown > Value Filters > Top 10... then set to Top 5
- C) Use a slicer
- D) Create a separate PivotTable for top 5

**11.** What happens when you add a Slicer to a PivotTable?
- A) It replaces the PivotTable filters
- B) It creates a visual filter panel that can be clicked to filter the PivotTable
- C) It exports the data to a new sheet
- D) It creates a chart

**12.** You have a 2-variable Data Table where the row input cell is interest rate and the column input cell is loan amount. The table shows:
- A) One result per combination of rate and amount
- B) Only the interest rate results
- C) Only the loan amount results
- D) An error unless you use Solver

**13.** You want a data validation dropdown that shows options from a named range called `DeptList`. Which source formula?
- A) `=DeptList`
- B) `=INDIRECT("DeptList")`
- C) You type the range `=$H$1:$H$10`
- D) Both A and B work

**14.** What does "Refresh" do on a PivotTable?
- A) Deletes and recreates the PivotTable
- B) Updates the PivotTable with any changes made to the source data
- C) Clears all filters
- D) Resets the layout to default

**15.** Solver differs from Goal Seek because Solver:
- A) Can optimize with multiple constraints and variables
- B) Only works with one variable
- C) Is built into Excel without add-in
- D) Only finds minimum values

---

## Part B: Fill-in-the-Formula Challenge (5 points each)

**16.** Write a Data Validation custom formula that only allows entries greater than the value in cell A1.

> **Your answer:** `_________________________________`

**17.** You have a PivotTable named `PivotTable1`. Write the GETPIVOTDATA formula to extract the sales amount for Region "East".

> **Your answer:** `_________________________________`

**18.** Describe the steps to create a 1-variable Data Table that shows how monthly payments change for different interest rates (rates in column starting at A2, PMT formula in A1).

> **Your answer:** `_________________________________`

**19.** Write a formula using SUBTOTAL that calculates the average of visible cells in B2:B100 (accounting for filtered rows).

> **Your answer:** `_________________________________`

**20.** You want to add a calculated field to a PivotTable that divides Sales by Quantity to get Unit Price. Describe the steps.

> **Your answer:** `_________________________________`

---

## Answer Key

| # | Answer | Explanation |
|---|--------|-------------|
| 1 | **B** | PivotTables are designed exactly for this cross-tabulation. |
| 2 | **A** | Filters hide rows but don't change row numbers - you see gaps. |
| 3 | **A** | Numeric fields default to SUM in PivotTable Values. |
| 4 | **B** | Data Validation restricts input; conditional formatting only visualizes. |
| 5 | **A** | Goal Seek adjusts one input to reach a target output. |
| 6 | **B** | Multi-level sorts: primary sort first, secondary sort within ties. |
| 7 | **B** | "Show Values As > % of Grand Total" is built into PivotTables. |
| 8 | **B** | Data Validation shows a customizable error alert dialog. |
| 9 | **A** | Scenario Summary creates a comparison table of all scenarios. |
| 10 | **B** | Value Filters > Top 10 lets you specify top N by any measure. |
| 11 | **B** | Slicers are visual, clickable filter buttons. |
| 12 | **A** | A 2-variable data table shows one result per row×column combination. |
| 13 | **D** | Both direct name and INDIRECT work as Data Validation sources. |
| 14 | **B** | Refresh pulls in updated source data without recreating the PivotTable. |
| 15 | **A** | Solver handles multiple variables, constraints, and optimization types. |
| 16 | `=B1>A1` | Apply to the validation range; B1 is the first cell in the range being validated. |
| 17 | `=GETPIVOTDATA("Sales",PivotTable1,"Region","East")` | Extracts a specific value from a PivotTable. |
| 18 | Select range containing PMT formula + rate column → Data tab → What-If Analysis → Data Table → Column input cell = the interest rate cell used in the PMT formula. Leave Row input cell blank. | - |
| 19 | `=SUBTOTAL(1,B2:B100)` | Function 1 = AVERAGE; SUBTOTAL ignores filtered-out rows. |
| 20 | PivotTable Analyze tab → Fields, Items & Sets → Calculated Field → Name: "Unit Price" → Formula: `=Sales/Quantity` → OK. | - |

---

**Scoring:**
- Part A: 15 questions × 3 points = 45 points
- Part B: 5 questions × 5 points = 25 points
- **Total: 70 points**
- 63+ (90%): ✅ Ready for Module 5
- 49–62 (70–89%): 📖 Review weak areas
- Below 49 (<70%): 🔄 Re-study Module 4


## FILE: quizzes/quiz-05-advanced-lookup.md

# Quiz 05 - Advanced Lookup & Dynamic Arrays

**Module:** INDEX/MATCH, XLOOKUP, Dynamic Arrays, OFFSET/INDIRECT
**Total Questions:** 20 (15 Multiple Choice + 5 Formula Challenges)
**Passing Score:** 90%+ = Ready to advance | 70–89% = Review weak areas | <70% = Re-study module

---

## Part A: Multiple Choice (3 points each)

**1.** Why is `INDEX/MATCH` often preferred over `VLOOKUP`?
- A) It's shorter to type
- B) It can look left (return a column to the left of the lookup column)
- C) It's always faster
- D) It doesn't need a lookup value

**2.** What does `=XLOOKUP("Widget",A:A,B:B,"Not Found",0)` return if "Widget" is not in column A?
- A) `#N/A`
- B) `0`
- C) `"Not Found"`
- D) Blank cell

**3.** You enter `=SEQUENCE(3,2)` in cell A1. What appears?
- A) A single column of 1, 2, 3
- B) A 3-row, 2-column array: `{1,2; 3,4; 5,6}`
- C) An error
- D) The number 6

**4.** What is the "spill range" in dynamic arrays?
- A) The area where a formula result overflows into adjacent cells
- B) A named range you create manually
- C) The error when a formula can't display results
- D) A PivotTable feature

**5.** `=OFFSET(A1,2,3,1,1)` returns the value in which cell?
- A) `A3`
- B) `D3`
- C) `D1`
- D) `A4`

**6.** You have `=XLOOKUP(A1, Products[Name], Products[Price])`. What's the 4th optional argument that controls "not found" behavior?
- A) `match_mode`
- B) `search_mode`
- C) `if_not_found`
- D) `return_array`

**7.** What does `=UNIQUE(A1:A100)` return?
- A) The first unique value only
- B) An array of all distinct values from the range
- C) A count of unique values
- D) TRUE/FALSE for each cell

**8.** A dynamic array formula shows `#SPILL!` error. What's the most likely cause?
- A) The formula is wrong
- B) Cells in the spill range are occupied by other data
- C) The range is too large
- D) The workbook is protected

**9.** Which formula correctly uses INDEX/MATCH to return the price of "Gadget" from a table where names are in A1:A10 and prices in B1:B10?
- A) `=INDEX(B1:B10,MATCH("Gadget",A1:A10,0))`
- B) `=INDEX(A1:A10,MATCH("Gadget",B1:B10,0))`
- C) `=MATCH(B1:B10,"Gadget",0)`
- D) `=INDEX("Gadget",A1:A10,B1:B10)`

**10.** `=SORT(A2:C10,3,-1)` does what?
- A) Sorts by column 3 in ascending order
- B) Sorts by column 3 in descending order
- C) Sorts 3 columns randomly
- D) Returns an error

**11.** What does `=INDIRECT("Sheet"&A1&"!B5")` do when A1 contains `2`?
- A) Returns the value in B5 of Sheet2
- B) Returns the text "Sheet2!B5"
- C) Returns `#REF!`
- D) Returns the value in B5 of the current sheet

**12.** You need a formula that returns the last non-empty value in column A. Which works?
- A) `=LOOKUP(2,1/(A:A<>""),A:A)`
- B) `=INDEX(A:A,COUNTA(A:A))`
- C) `=XLOOKUP("*",A:A,A:A,,2)`
- D) Both A and B (though B assumes no gaps)

**13.** What does `=FILTER(A2:C100, B2:B100>50, "None")` return?
- A) Rows where column B exceeds 50; shows "None" if no matches
- B) Column B values over 50 only
- C) An error
- D) All rows with "None" in column B

**14.** The `match_mode` parameter in XLOOKUP set to `-1` means:
- A) Exact match
- B) Exact match or next smaller item
- C) Exact match or next larger item
- D) Wildcard match

**15.** Which function can create a dynamic named range that automatically expands?
- A) `OFFSET` with `COUNTA`
- B) `INDEX` with `COUNTA`
- C) Both A and B
- D) Neither - dynamic ranges must be Excel Tables

---

## Part B: Fill-in-the-Formula Challenge (5 points each)

**16.** Write an XLOOKUP formula that searches for the value in A1 in the range D1:D100, returns the corresponding value from E1:E100, and shows "N/A" if not found.

> **Your answer:** `_________________________________`

**17.** Write a formula using INDEX/MATCH that finds the sales amount for "Product X" in column B where product names are in column A and sales are in column C.

> **Your answer:** `_________________________________`

**18.** Write a formula that returns the top 5 values from A1:A100 using SORT and SEQUENCE (or SORTBY).

> **Your answer:** `_________________________________`

**19.** Write a formula using FILTER that returns all rows from A1:D100 where column C equals "Active" AND column D is greater than 1000.

> **Your answer:** `_________________________________`

**20.** Create a dynamic named range formula using OFFSET that covers column A from A1 to the last non-empty cell.

> **Your answer:** `_________________________________`

---

## Answer Key

| # | Answer | Explanation |
|---|--------|-------------|
| 1 | **B** | INDEX/MATCH doesn't require the lookup column to be the leftmost. |
| 2 | **C** | The 4th argument ("Not Found") is the fallback when no match exists. |
| 3 | **B** | SEQUENCE(3,2) creates a 3×2 array: 1,2 / 3,4 / 5,6. |
| 4 | **A** | Spill is when a multi-result formula auto-fills adjacent cells. |
| 5 | **B** | OFFSET(A1, row_offset=2, col_offset=3) → A1+2 rows, +3 cols = D3. |
| 6 | **C** | `if_not_found` is the 4th argument for custom fallback values. |
| 7 | **B** | UNIQUE returns an array of all distinct values. |
| 8 | **B** | #SPILL! occurs when the output area has existing data blocking it. |
| 9 | **A** | INDEX returns from the price column; MATCH finds "Gadget" in the name column. |
| 10 | **B** | Third argument `-1` means descending sort. |
| 11 | **A** | INDIRECT evaluates the string as a reference to Sheet2's B5. |
| 12 | **D** | LOOKUP trick and INDEX+COUNTA both work (LOOKUP handles gaps better). |
| 13 | **A** | FILTER returns matching rows; the third argument is the "if_empty" fallback. |
| 14 | **B** | `-1` = exact match or next smaller item (useful for bracket lookups). |
| 15 | **C** | Both OFFSET+COUNTA and INDEX+COUNTA can define expanding ranges. |
| 16 | `=XLOOKUP(A1,D1:D100,E1:E100,"N/A")` | Basic XLOOKUP with custom not-found message. |
| 17 | `=INDEX(C1:C100,MATCH("Product X",A1:A100,0))` | MATCH finds the row, INDEX returns from sales column. |
| 18 | `=SORT(A1:A100,1,-1)` then reference the first 5 rows, or `=TAKE(SORT(A1:A100,1,-1),5)` | TAKE limits the sorted result to top 5. |
| 19 | `=FILTER(A1:D100,(C1:C100="Active")*(D1:D100>1000))` | Multiply conditions for AND logic in FILTER. |
| 20 | `=OFFSET(A1,0,0,COUNTA(A:A),1)` | Starts at A1, spans COUNTA(A:A) rows tall, 1 column wide. |

---

**Scoring:**
- Part A: 15 questions × 3 points = 45 points
- Part B: 5 questions × 5 points = 25 points
- **Total: 70 points**
- 63+ (90%): ✅ Ready for Module 6
- 49–62 (70–89%): 📖 Review weak areas
- Below 49 (<70%): 🔄 Re-study Module 5


## FILE: quizzes/quiz-06-power-tools.md

# Quiz 06 - Power Query, Power Pivot & DAX

**Module:** Power Query, Power Pivot, DAX, Data Modeling
**Total Questions:** 20 (15 Multiple Choice + 5 Formula Challenges)
**Passing Score:** 90%+ = Ready to advance | 70–89% = Review weak areas | <70% = Re-study module

---

## Part A: Multiple Choice (3 points each)

**1.** What is the primary purpose of Power Query?
- A) Creating PivotTables
- B) Connecting to, transforming, and loading data from various sources
- C) Writing VBA macros
- D) Creating charts

**2.** In Power Query, you apply a "Remove Duplicates" step on column A. The result:
- A) Deletes duplicate rows from the original data
- B) Creates a new query result with duplicates removed - original data unchanged
- C) Highlights duplicates with color
- D) Moves duplicates to a new sheet

**3.** What is a data model in Excel?
- A) A 3D chart template
- B) A collection of tables with relationships defined between them
- C) A VBA script that processes data
- D) A Power Query connection

**4.** In DAX, what does `CALCULATE(SUM(Sales[Amount]), Sales[Region]="East")` do?
- A) Sums all sales amounts
- B) Sums sales amounts filtered to only the "East" region
- C) Returns an error
- D) Counts East region sales

**5.** You have two tables: Orders (OrderID, CustomerID, Amount) and Customers (CustomerID, Name, City). To create a relationship, you link:
- A) OrderID to CustomerID
- B) CustomerID in Orders to CustomerID in Customers
- C) Amount to Name
- D) Any column - Excel figures it out

**6.** What is the difference between "Close & Load" and "Close & Load To..." in Power Query?
- A) There's no difference
- B) "Close & Load" puts data in a new sheet; "Close & Load To..." lets you choose the destination
- C) "Close & Load To..." saves the file
- D) "Close & Load" creates a PivotTable

**7.** In DAX, which function creates a calculated column that ranks products by sales?
- A) `RANKX(ALL(Products), Products[Sales])`
- B) `RANK.EQ(Products[Sales])`
- C) `SUMX(Products, Products[Sales])`
- D) `INDEX(Products[Sales])`

**8.** Power Pivot differs from a regular PivotTable because:
- A) It can handle millions of rows without slowing down
- B) It can use DAX measures
- C) It works with the data model
- D) All of the above

**9.** You merge two queries in Power Query using a Left Outer join. What does this return?
- A) Only matching rows from both tables
- B) All rows from the left table, matching rows from the right (null if no match)
- C) All rows from both tables
- D) Only rows that don't match

**10.** A DAX measure using `ALL(Sales[Region])` does what within CALCULATE?
- A) Filters to all regions
- B) Removes the Region filter, showing totals across all regions
- C) Returns an error
- D) Filters to only blank regions

**11.** In Power Query's M language, what does `= Table.AddColumn(Source, "Tax", each [Amount] * 0.1)` do?
- A) Multiplies all amounts by 0.1
- B) Adds a new column "Tax" that calculates 10% of each row's Amount
- C) Replaces the Amount column
- D) Creates a separate table

**12.** Which DAX function is the time-intelligence equivalent of "year-to-date total"?
- A) `TOTALYTD()`
- B) `DATESYTD()`
- C) `YTD()`
- D) `SUMYTD()`

**13.** You append (union) two queries in Power Query. Both must have:
- A) The same number of rows
- B) The same column names and compatible data types
- C) The same number of columns only
- D) Identical data

**14.** What does "Edit Queries" in a PivotTable connected to Power Query let you do?
- A) Modify the data transformation steps
- B) Change the PivotTable layout
- C) Write SQL queries
- D) Nothing - it's read-only

**15.** In DAX, `SUMX` differs from `SUM` because:
- A) SUMX is faster
- B) SUMX evaluates an expression row by row before summing
- C) SUMX only works in calculated columns
- D) There's no difference

---

## Part B: Fill-in-the-Formula Challenge (5 points each)

**16.** Write a DAX measure that calculates the total sales amount: `Total Sales = ...`

> **Your answer:** `_________________________________`

**17.** Write a DAX measure that calculates the percentage of total sales for the current filter context: `Sales % = ...`

> **Your answer:** `_________________________________`

**18.** Describe the Power Query steps to split a column containing "Last, First" into two separate columns.

> **Your answer:** `_________________________________`

**19.** Write a DAX calculated column formula that flags orders as "High" if Amount > 1000, otherwise "Low".

> **Your answer:** `_________________________________`

**20.** Describe how to create a date table in Power Query that spans from Jan 1, 2023 to Dec 31, 2025.

> **Your answer:** `_________________________________`

---

## Answer Key

| # | Answer | Explanation |
|---|--------|-------------|
| 1 | **B** | Power Query is Excel's ETL tool for data import and transformation. |
| 2 | **B** | Power Query is non-destructive - it creates a new output from transformation steps. |
| 3 | **B** | The data model stores tables and their relationships for cross-table analysis. |
| 4 | **B** | CALCULATE modifies the filter context - here it filters to Region="East". |
| 5 | **B** | Relationships require matching key columns between tables. |
| 6 | **B** | "Close & Load To..." gives options: Table, PivotTable, Connection Only, etc. |
| 7 | **A** | RANKX iterates over a table and ranks based on an expression. |
| 8 | **D** | Power Pivot leverages the data model, DAX, and handles large datasets efficiently. |
| 9 | **B** | Left Outer keeps all left rows; right rows are null when no match exists. |
| 10 | **B** | ALL removes filters - used to calculate totals ignoring the current filter. |
| 11 | **B** | Table.AddColumn adds a new computed column to the table. |
| 12 | **A** | TOTALYTD computes year-to-date totals with a date column reference. |
| 13 | **B** | Append requires compatible columns (names and types). |
| 14 | **A** | "Edit Queries" opens Power Query Editor to modify transformation steps. |
| 15 | **B** | SUMX iterates row by row evaluating an expression; SUM just adds a column. |
| 16 | `Total Sales = SUM(Sales[Amount])` | Simple aggregation measure. |
| 17 | `Sales Pct = DIVIDE(SUM(Sales[Amount]), CALCULATE(SUM(Sales[Amount]), ALL(Sales)))` | DIVIDE handles division by zero; ALL removes filters for the denominator. |
| 18 | Select the column → Transform tab → Split Column → By Delimiter → choose Comma → set "At each occurrence" → OK. This creates two columns. Rename as needed. | - |
| 19 | `Order Flag = IF(Sales[Amount] > 1000, "High", "Low")` | Calculated column with IF. |
| 20 | Home > New Source > Blank Query → In Advanced Editor, enter: `= List.Dates(#date(2023,1,1), 1096, #duration(1,0,0,0))` → Convert to Table → Add date-related columns (Year, Month, etc.) → Close & Load. | 1096 days covers 2023-01-01 to 2025-12-31. |

---

**Scoring:**
- Part A: 15 questions × 3 points = 45 points
- Part B: 5 questions × 5 points = 25 points
- **Total: 70 points**
- 63+ (90%): ✅ Ready for Module 7
- 49–62 (70–89%): 📖 Review weak areas
- Below 49 (<70%): 🔄 Re-study Module 6


## FILE: quizzes/quiz-07-vba.md

# Quiz 07 - VBA Macros

**Module:** VBA Macros, Loops, UserForms, Error Handling
**Total Questions:** 20 (15 Multiple Choice + 5 Formula Challenges)
**Passing Score:** 90%+ = Ready to advance | 70–89% = Review weak areas | <70% = Re-study module

---

## Part A: Multiple Choice (3 points each)

**1.** What keyboard shortcut opens the VBA Editor?
- A) `Alt + F11`
- B) `Ctrl + M`
- C) `F5`
- D) `Alt + F8`

**2.** You record a macro that formats cell A1 bold. You then run it. Cell B5:
- A) Also becomes bold
- B) Stays unchanged - the macro only affects A1
- C) Gets deleted
- D) Shows an error

**3.** In VBA, what does `MsgBox "Hello"` display?
- A) A text box on the worksheet
- B) A pop-up dialog box with the text "Hello"
- C) Writes "Hello" to cell A1
- D) Prints "Hello" to the Immediate window

**4.** Which loop structure runs at least once before checking the condition?
- A) `For...Next`
- B) `Do While...Loop`
- C) `Do...Loop While`
- D) `For Each...Next`

**5.** You want to loop through all cells in column A from A1 to A100. Which is correct?
- A) `For i = 1 To 100: Cells(i, 1).Value = ... : Next i`
- B) `For Each cell In Range("A1:A100"): cell.Value = ... : Next`
- C) Both A and B
- D) Neither - you must use `Do While`

**6.** What does `Application.ScreenUpdating = False` do?
- A) Hides Excel completely
- B) Prevents screen redraws during macro execution, speeding it up
- C) Disables all user input
- D) Turns off the formula bar

**7.** A UserForm has a TextBox named `txtName` and a CommandButton. To read the user's input in the button's click event:
- A) `MsgBox txtName.Value`
- B) `MsgBox txtName.Text`
- C) Both A and B work
- D) `MsgBox txtName.Caption`

**8.** What does `On Error Resume Next` do?
- A) Stops the macro on any error
- B) Skips the error-causing line and continues execution
- C) Shows a custom error message
- D) Deletes the error

**9.** Which statement correctly declares a variable in VBA?
- A) `Dim x As Integer`
- B) `var x = Integer`
- C) `Integer x`
- D) `Declare x As Integer`

**10.** You want to call a macro named `ProcessData` from another macro. The correct syntax is:
- A) `Call ProcessData`
- B) `ProcessData`
- C) `Run ProcessData`
- D) Both A and B

**11.** What is the difference between `Sub` and `Function` in VBA?
- A) There is no difference
- B) A Sub performs actions; a Function returns a value
- C) A Function is faster
- D) A Sub can't be called from a worksheet cell

**12.** You want a macro to run automatically when the workbook opens. Name the sub:
- A) `Auto_Open()`
- B) `Workbook_Open()`
- C) `Sub Start()`
- D) Both A and B work

**13.** What does `Cells(Rows.Count, 1).End(xlUp).Row` return?
- A) Always row 1
- B) The last used row in column A
- C) The total number of rows
- D) An error

**14.** To add a dropdown list to a UserForm, you use:
- A) TextBox
- B) ComboBox or ListBox
- C) CheckBox
- D) Label

**15.** What does `Err.Number` give you after an error occurs?
- A) The line number where the error happened
- B) The error code identifying the type of error
- C) The number of errors so far
- D) A random number

---

## Part B: Fill-in-the-Formula Challenge (5 points each)

Write the VBA code that accomplishes the task.

**16.** Write a Sub that loops through cells A1:A10 and makes any cell with a value greater than 100 bold.

> **Your answer:** `_________________________________`

**17.** Write a `For Each` loop that clears all cells in the range B2:B50 that contain the text "TBD".

> **Your answer:** `_________________________________`

**18.** Write a Function called `DoubleIt` that takes a number as input and returns double the value.

> **Your answer:** `_________________________________`

**19.** Write the error handling block that shows a custom message box when an error occurs, then exits the Sub.

> **Your answer:** `_________________________________`

**20.** Write VBA code that creates a new worksheet named "Summary" at the end of the workbook (handling the case where it already exists).

> **Your answer:** `_________________________________`

---

## Answer Key

| # | Answer | Explanation |
|---|--------|-------------|
| 1 | **A** | `Alt+F11` opens VBA Editor; `Alt+F8` opens the macro list dialog. |
| 2 | **B** | Recorded macros use absolute references by default - only the recorded cells are affected. |
| 3 | **B** | MsgBox displays a modal dialog box to the user. |
| 4 | **C** | `Do...Loop While` executes the body first, then checks the condition. |
| 5 | **C** | Both `For i` with `Cells()` and `For Each` with a range work. |
| 6 | **B** | Disabling screen updating prevents flicker and speeds up macros significantly. |
| 7 | **C** | Both `.Value` and `.Text` return the TextBox content. |
| 8 | **B** | `On Error Resume Next` suppresses errors and continues to the next line. |
| 9 | **A** | `Dim x As Integer` is the standard VBA variable declaration. |
| 10 | **D** | Both `Call ProcessData` and just `ProcessData` invoke the sub. |
| 11 | **B** | Subs execute code; Functions return values and can be used in expressions. |
| 12 | **D** | Both `Auto_Open()` and `Workbook_Open()` run on open (Workbook_Open is preferred). |
| 13 | **B** | Starting from the bottom of column A and pressing Ctrl+Up gives the last used row. |
| 14 | **B** | ComboBox (editable dropdown) or ListBox (scrollable list) provides dropdown behavior. |
| 15 | **B** | `Err.Number` is the numeric error code (e.g., 1004 for application-defined errors). |
| 16 | ```vba Sub BoldOver100() Dim cell As Range For Each cell In Range("A1:A10") If cell.Value > 100 Then cell.Font.Bold = True Next cell End Sub``` | - |
| 17 | ```vba Sub ClearTBD() Dim cell As Range For Each cell In Range("B2:B50") If cell.Value = "TBD" Then cell.ClearContents Next cell End Sub``` | - |
| 18 | ```vba Function DoubleIt(num As Double) As Double DoubleIt = num * 2 End Function``` | - |
| 19 | ```vba Sub MyMacro() On Error GoTo ErrHandler ' ... code ... Exit Sub ErrHandler: MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical Exit Sub End Sub``` | - |
| 20 | ```vba Sub CreateSummary() Dim ws As Worksheet On Error Resume Next Set ws = ThisWorkbook.Sheets("Summary") On Error GoTo 0 If ws Is Nothing Then Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)) ws.Name = "Summary" End If End Sub``` | - |

---

**Scoring:**
- Part A: 15 questions × 3 points = 45 points
- Part B: 5 questions × 5 points = 25 points
- **Total: 70 points**
- 63+ (90%): ✅ Ready for Module 8
- 49–62 (70–89%): 📖 Review weak areas
- Below 49 (<70%): 🔄 Re-study Module 7


## FILE: quizzes/quiz-08-projects.md

# Quiz 08 - Real-World Projects & Dashboards

**Module:** Real-World Projects, Dashboards, Financial Models
**Total Questions:** 20 (15 Multiple Choice + 5 Formula Challenges)
**Passing Score:** 90%+ = Ready to advance | 70–89% = Review weak areas | <70% = Re-study module

---

## Part A: Multiple Choice (3 points each)

**1.** You're building a dashboard that needs to show KPIs from a large dataset refreshed weekly. The best approach is:
- A) Copy-paste data each week
- B) Use Power Query to load and transform data, with PivotTables and charts that auto-refresh
- C) Manually re-create charts each week
- D) Use only VBA macros

**2.** In a financial model, which practice best prevents accidental formula changes?
- A) Use bright colors to highlight input cells
- B) Color-code cells: blue font for inputs, black for formulas, and protect formula cells
- C) Print the formula bar
- D) Hide all formulas

**3.** Your dashboard has 6 charts and 3 PivotTables. Users complain it's slow. What's the first optimization to try?
- A) Buy a faster computer
- B) Reduce data volume, use data model instead of cell formulas, and limit volatile functions
- C) Delete half the charts
- D) Convert everything to CSV

**4.** When building a revenue forecast model, which Excel feature lets users toggle between optimistic, pessimistic, and base scenarios?
- A) Data Validation dropdown driving scenario inputs
- B) Manual editing of each cell
- C) Separate workbooks for each scenario
- D) Conditional formatting only

**5.** A dashboard should include all EXCEPT:
- A) Clear section headers and labels
- B) Interactive filters (slicers, dropdowns)
- C) Raw data tables with thousands of rows
- D) KPI summary cards

**6.** You're building a P&L statement in Excel. Revenue is in row 10, COGS in row 11. The gross profit formula should be:
- A) `=B10-B11`
- B) `=SUM(B10:B11)`
- C) `=B10/B11`
- D) `=B10*B11`

**7.** What does a "sensitivity analysis" in a financial model show?
- A) How output changes as one or more input assumptions change
- B) Whether the model has errors
- C) The model's file size
- D) How fast the model recalculates

**8.** You want a dashboard chart to update based on a user's dropdown selection. The technique involves:
- A) VBA only
- B) Using INDEX/MATCH or XLOOKUP to pull selected data into a helper range, then charting that range
- C) Creating a separate chart for each option
- D) Using Word Art

**9.** In project management dashboards, a Gantt chart in Excel is typically built using:
- A) A scatter plot with error bars
- B) A stacked bar chart with invisible base bars
- C) A pie chart
- D) Conditional formatting only

**10.** What is the purpose of an "assumptions sheet" in a financial model?
- A) It lists all input variables and assumptions in one place for easy review
- B) It assumes the model is correct
- C) It contains only hardcoded values
- D) It's an optional decorative element

**11.** You need to track budget vs. actual across 12 months. Which visualization works best?
- A) 12 separate pie charts
- B) A clustered bar chart with Budget and Actual as series, months on the x-axis
- C) A single number showing the total difference
- D) A line chart of actuals only

**12.** Dynamic array formulas can make dashboards more powerful because:
- A) They auto-spill results without copying formulas
- B) They integrate with SORT, FILTER, and UNIQUE for live data views
- C) They reduce formula maintenance
- D) All of the above

**13.** For a professional financial model, which layout is recommended?
- A) All inputs, calculations, and outputs on one sheet
- B) Separate sheets for inputs, calculations, and outputs (clear flow)
- C) Random placement wherever there's space
- D) All formulas in one column

**14.** You're building an inventory dashboard. To show stock levels that need reorder, you use:
- A) Conditional formatting (red when stock < reorder point)
- B) A separate "reorder" sheet with manual checks
- C) A macro that deletes items below threshold
- D) Nothing - just print the full list

**15.** What is "model auditing" in Excel?
- A) Checking the model for errors, tracing precedents/dependents, and validating logic
- B) Auditing the file size
- C) Checking if the model runs on Mac
- D) A VBA feature only

---

## Part B: Fill-in-the-Formula Challenge (5 points each)

**16.** Write a formula for a KPI card that calculates the percentage change from last month's revenue (in A1) to this month's (in B1).

> **Your answer:** `_________________________________`

**17.** You have a dropdown in cell D1 that selects a department name. Write a formula that returns the total expenses for the selected department from a Table named `Expenses` with columns `[Department]` and `[Amount]`.

> **Your answer:** `_________________________________`

**18.** Write a formula that calculates the Compound Annual Growth Rate (CAGR) given Beginning Value in A1, Ending Value in B1, and Number of Years in C1.

> **Your answer:** `_________________________________`

**19.** Write a formula that flags rows in a project tracker where the End Date (column D) is before today AND Status (column E) is not "Complete".

> **Your answer:** `_________________________________`

**20.** Describe the structure of a 3-statement financial model (what are the 3 statements and how they link).

> **Your answer:** `_________________________________`

---

## Answer Key

| # | Answer | Explanation |
|---|--------|-------------|
| 1 | **B** | Power Query + PivotTables + auto-refresh is the standard dashboard data pipeline. |
| 2 | **B** | Color-coding (blue=input, black=formula) + sheet protection is financial modeling best practice. |
| 3 | **B** | Reduce data, use data model, avoid volatile functions (NOW, INDIRECT, OFFSET in excess). |
| 4 | **A** | A Data Validation dropdown lets users select scenarios; formulas reference the selection. |
| 5 | **C** | Dashboards should show summaries - raw data belongs in a separate data sheet. |
| 6 | **A** | Gross Profit = Revenue − COGS. |
| 7 | **A** | Sensitivity analysis shows how outputs respond to changes in input assumptions. |
| 8 | **B** | INDEX/MATCH or XLOOKUP dynamically pull data for the selected item into a chart-ready range. |
| 9 | **B** | Gantt charts use stacked bars: invisible base + visible duration bar per task. |
| 10 | **A** | The assumptions sheet centralizes all inputs for transparency and easy updates. |
| 11 | **B** | Clustered bar charts are ideal for comparing two series across categories (months). |
| 12 | **D** | Dynamic arrays auto-fill, reduce maintenance, and power interactive views. |
| 13 | **B** | Separation of inputs, calculations, and outputs is the standard modeling convention. |
| 14 | **A** | Conditional formatting provides instant visual alerts for low stock. |
| 15 | **A** | Model auditing traces formula logic, finds errors, and validates assumptions. |
| 16 | `=(B1-A1)/A1` or `=(B1-A1)/ABS(A1)` | Percentage change: (New − Old) / Old. Use ABS if values can be negative. |
| 17 | `=SUMIF(Expenses[Department],D1,Expenses[Amount])` | SUMIF with the dropdown value as criteria. |
| 18 | `=(B1/A1)^(1/C1)-1` | CAGR formula: (End/Begin)^(1/years) − 1. |
| 19 | `=AND(D2<TODAY(),E2<>"Complete")` | Use in a helper column or conditional formatting. Returns TRUE for overdue incomplete tasks. |
| 20 | The 3 statements are: **Income Statement** (revenue, expenses, net income), **Balance Sheet** (assets, liabilities, equity), and **Cash Flow Statement** (operating, investing, financing). Net Income flows from the IS to the BS (retained earnings) and to the CFS (starting point for operating cash flow). The BS balances: Assets = Liabilities + Equity. Cash from CFS links back to the BS cash balance. | - |

---

**Scoring:**
- Part A: 15 questions × 3 points = 45 points
- Part B: 5 questions × 5 points = 25 points
- **Total: 70 points**
- 63+ (90%): ✅ Excel Proficiency Certificate Ready!
- 49–62 (70–89%): 📖 Review weak areas
- Below 49 (<70%): 🔄 Re-study Module 8

---

🎓 **Congratulations on completing all 8 module quizzes!**
If you scored 90%+ on all quizzes, you have demonstrated strong Excel proficiency across fundamentals through advanced projects.


## FILE: samples/07-vba-toolkit/README.md

# Sample 07: VBA Toolkit (Module 7)

Ready-to-use VBA modules plus a practice workbook. Real `.xlsm` files with embedded macros are binary artifacts, so this toolkit ships plain-text `.bas` modules instead: importing them takes 10 seconds and teaches you the VBA editor along the way.

## Contents

| File | Macros | What they do |
|------|--------|--------------|
| `mod_report_tools.bas` | `FormatReport`, `AddSummaryRow` | One-click report formatting + totals row |
| `mod_bulk_tools.bas` | `RemoveEmptyRows`, `ExportEachSheetToCSV`, `RenameSheetsFromColumnA` | Bulk cleaning and export jobs |
| `practice-data.xlsx` | (data only) | Messy multi-sheet workbook to practice on |

## How to Import and Run

1. Open `practice-data.xlsx` and immediately **Save As `practice-data.xlsm`** (macro-enabled). Excel will not keep macros in `.xlsx`.
2. Press **Alt+F11** to open the VBA editor.
3. **File > Import File** and pick `mod_report_tools.bas`, then repeat for `mod_bulk_tools.bas`.
4. Close the editor, press **Alt+F8**, pick a macro, click **Run**.
5. If Excel shows a security banner, click **Enable Content** (this is your own code: you can read every line first).

## Practice Path

1. Run `FormatReport` on the `Raw` sheet and read what each block changes.
2. Run `AddSummaryRow` and check the `=SUM` range it wrote.
3. Run `RemoveEmptyRows` on `Messy`, then `RenameSheetsFromColumnA`.
4. Break one macro on purpose (change `Columns(2)` to `Columns(99)`), run it, and read the error. Fixing broken code is half of Module 7.
5. Write your own 5-line macro that stamps today's date in `Z1` of every sheet.

## Safety Note

Macros are code. Never import `.bas` files you have not read, and never enable macros in workbooks from strangers. These two modules are plain text: open them in any editor first.


## FILE: samples/fix-this-workbook-challenge.md

# The Fix-This-Workbook Challenge (Find 20 Issues)

Companion to `exercise-fix-this-workbook.xlsx` (deliberately broken Q2 report for a fictional product, YourCompany Inc).

## How the Challenge Works

1. Open `exercise-fix-this-workbook.xlsx` and inspect all 4 sheets (Results, Data, Lookups, Charts) for 5 minutes without editing.
2. List every problem: there are **20 planted issues** in 4 categories.
3. Score: 16-20 found = hero eyes, 11-15 = solid, 6-10 = revisit Modules 1-4, 0-5 = restudy then retry.
4. Then rebuild the workbook properly: clean Data, working formulas, honest charts. Rebuilding is where the learning happens.

## Issue Worksheet (write your findings before checking the key)

| # | Sheet | Category | What is wrong | Which module fixes it |
|---|-------|----------|---------------|-----------------------|
| 1 | | | | |
| ... up to 20 | | | | |

---

## ANSWER KEY (spoilers below the line)

### Formula Issues (5)

1. **Hardcoded total**: `Results!B3` holds the literal value `25` where the label promises a formula. Fix: live `=SUM` formulas (Module 2).
2. **#REF! error**: `Results!B4` is `=SUM(#REF!)`, a deleted-range leftover. Fix: rebuild the reference (Module 2 error handling).
3. **#DIV/0! error**: `Results!B5` divides by the zero in `B7`. Fix: `IFERROR` or fix the denominator (Module 2).
4. **Off-by-one SUM**: `Data!H16` is `=SUM(H2:H12)` but data runs to row 14, silently missing 2 rows. Fix: structured table references (Modules 2, 3).
5. **VLOOKUP approximate match**: `Lookups!B4` uses `VLOOKUP(...,TRUE)` against an unsorted SKU column, returning wrong prices silently. Fix: exact match `FALSE`/`0` or XLOOKUP (Module 5).

### Data Issues (5)

6. **Numbers stored as text**: `Data` column D (Qty) is text-formatted, breaking every SUM. Fix: Text to Columns or `--` coercion (Module 1).
7. **Dates stored as text**: `Data` column B holds strings like `03/05/2026`, so no date math works. Fix: real date values (Module 1).
8. **Trailing spaces in keys**: `"Northwind Traders "` and `"Harborline Credit Union "` break lookups and grouping. Fix: TRIM (Module 2 text functions).
9. **Duplicate rows**: rows 10-11 duplicate rows 4-5 exactly. Fix: Remove Duplicates (Module 4).
10. **Blank header**: `Data!F1` is empty, so the column cannot be part of a table. Fix: name every header (Module 3).

### Formatting Issues (5)

11. **Three fonts in one workbook**: Impact (Results title), Comic Sans MS (Data), Calibri everywhere else. Fix: one theme font pair (Module 3).
12. **Merged cells inside data**: `Data!A8:B8` and header `G1:H1` are merged, breaking sorting and tables. Fix: unmerge, use Center Across Selection (Modules 1, 3).
13. **Low contrast**: `Results!A10` is light gray on light yellow. Fix: contrast-safe colors (Module 3).
14. **Rainbow header row**: `Data` row 1 uses 5 bright fills with no system. Fix: one palette (Module 3).
15. **Mixed number formats**: `Data` column E mixes currency and General formats. Fix: consistent number formats (Module 3).

### Charts & Structure Issues (5)

16. **Pie chart for 12 months of trend data**: `Charts` pie of monthly revenue. Fix: line or column over time (Module 4).
17. **Truncated y-axis**: the "Revenue growth" chart starts at 78 on data ranging 42-91, exaggerating the climb. Fix: zero baseline (Module 4).
18. **Chart junk**: same chart shows legend + data labels + gridlines at once. Fix: data-ink cleanup (Module 4).
19. **Hidden column with real data**: `Data` column C ("corrections") is hidden, so nobody sees the notes that change the numbers. Fix: unhide, audit before trusting (Module 4).
20. **Circular reference**: `Data!E2` contains `=E2`. Fix: never let a cell depend on itself (Module 2).

## Grading Yourself

- 1 pt per issue spotted (sheet + what is wrong).
- Rebuild the workbook cleanly and re-run every formula: totals must match the rows.
- Log your score in `../10-Study-Plan/progress-tracker.md`.


## FILE: samples/walkthrough-01-data-entry.md

# Walkthrough: sample-01-data-entry-practice.xlsx

Companion to **Module 1: Fundamentals**.

## What to Notice

1. **The mess is realistic on purpose.** Order IDs mix `YC-0008` and `YC-3`; dates arrive as `2026-01-06`, `03/13/2026`, `2/14/26`, and `March 5 2026`. This is exactly what pasted exports look like in real jobs.
2. **Formulas are already waiting.** Column G carries `=E*F` formulas on every row, including the rows with missing inputs. Watch how they sit quiet until the data is fixed, then come alive. (Module 2)
3. **The Format-Me sheet stores numbers as text.** Totals refuse to add until you fix it. The fast fix is Data > Text to Columns > Finish, which converts in place. (Module 1)
4. **Green triangles are free lessons.** Every warning triangle marks a place where Excel suspects your data lies.

## Success Check

No green triangles anywhere, every Order ID matches `YC-0001` format, and column G totals are correct to the cent.


## FILE: samples/walkthrough-02-formulas.md

# Walkthrough: sample-02-formula-challenges.xlsx

Companion to **Module 2: Formulas & Functions**.

## What to Notice

1. **Three sheets, one loop.** Data (40 clean rows) > Challenges (12 graded tasks with expected results) > Answer Key (model formulas as text so you cannot peek by accident).
2. **Difficulty ramps on purpose.** Tasks 1-4 build single-cell fluency (arithmetic, IF, SUM, AVERAGE). Tasks 5-7 add conditions (COUNTIF, SUMIF, SUMIFS). Tasks 8-12 mix lookups, text joins, and error handling.
3. **The expected-result column is your self-grader.** If your number disagrees, the fault is in your formula, not the data.
4. **Task 2 teaches the IF mindset**: one condition, two outcomes, filled down once. Every fancy dashboard formula is this shape wearing a suit.

## Success Check

All 12 challenge cells match the expected results, and you can explain each formula out loud without reading it.


## FILE: samples/walkthrough-03-dashboard.md

# Walkthrough: sample-03-dashboard-before-after.xlsx

Companion to **Module 3: Formatting & Visualization**.

## What to Notice

1. **Before and After hold identical numbers.** Only formatting, formulas-in-view, and one chart differ. That is the whole lesson: presentation is a layer, not a rewrite.
2. **The After sheet's title states a claim** ("profit up 6 straight months"), not a topic. Assertion titles work in workbooks too. (Module 3)
3. **KPI callouts sit next to the table**: `=SUM` and `=INDEX/MATCH` style summaries, currency-formatted, so a manager reads the page in 5 seconds.
4. **Conditional formatting carries the analysis**: data bars on Profit show the climb at a glance; the cost highlight flags the month that broke the trend. (Module 3)
5. **One chart, no chart junk.** Compare with the chart crimes in `exercise-fix-this-workbook.xlsx` (issues 16-18).

## Success Check

Rebuild the After sheet from the Before sheet in under 10 minutes without peeking.


## FILE: samples/walkthrough-04-pivottables.md

# Walkthrough: sample-04-pivottable-source.xlsx

Companion to **Module 4: Data Analysis**.

## What to Notice

1. **300 clean rows, zero clutter.** PivotTables love boring, rectangular data: one header row, no merges, no blanks, consistent types. This sheet is what "analysis-ready" means.
2. **The Tasks sheet is the lesson.** Six pivot moves in order: totals by field, two-level grouping, date grouping, PivotChart, Slicer, and Show Values As percentage. (Module 4)
3. **Why a source workbook ships without a pivot**: building the pivot IS the exercise. Follow the task steps with Insert > PivotTable and compare results with a colleague or the quiz answer key style: the numbers must match regardless of who builds it.
4. **Refresh is the superpower.** Change one Revenue cell and every pivot updates on Data > Refresh All. That relationship between flat data and pivoted views is the core mental model of Module 4.

## Success Check

All six tasks done, and your Region totals match: sum of pivoted regions equals total revenue of the flat data.


## FILE: samples/walkthrough-05-lookups.md

# Walkthrough: sample-05-lookup-challenges.xlsx

Companion to **Module 5: Advanced Functions**.

## What to Notice

1. **Two tables that want to be one.** Orders knows SKUs but not names or prices; Products knows everything about SKUs. Lookups are just "borrow the missing columns". (Module 5)
2. **XLOOKUP is the hero, INDEX/MATCH is the resume.** Task 1 vs task 4 solve the same problem both ways so you can feel the difference in effort and readability.
3. **`if_not_found` is not optional in real life.** Task 2 bakes in `Discontinued` because missing SKUs happen. Task 3 then forces you to handle the text result gracefully with ISNUMBER. That pairing is how production spreadsheets stay honest.
4. **Tasks 5-7 graduate to array thinking**: SUMPRODUCT with XLOOKUP-returned arrays, and LET for one-pass readability. Task 8's dynamic-array spill is the future-proof version of task 1.

## Success Check

All challenge answers match the Answer Key, and you can explain why VLOOKUP's approximate match would be wrong here (preview: `exercise-fix-this-workbook.xlsx` issue 5).


## FILE: samples/walkthrough-06-powerquery.md

# Walkthrough: sample-06-powerquery-messy.xlsx

Companion to **Module 6: Power Query & Power Pivot**.

## What to Notice

1. **Three sheets, three different messes.** The Jan/Feb/Mar exports disagree on headers (`Date` vs `date ` vs `Transaction Date`), region casing (`North`/`north`), number storage (`" 830"`, `"1,450"`), and date formats. This is the real world of monthly exports. (Module 6)
2. **Each sheet carries the same 3 sins**: trailing spaces in names, exact duplicate rows, and blank separator rows. Once you build one cleaning query, it fixes every future month automatically.
3. **The point of Power Query is repeatability.** Your cleaned output must survive the test in task 8: edit Jan and Refresh All, and the cleaned table updates with zero rework. If you cleaned by hand, it will not.
4. **M language appears after you click.** Every step you build in the ribbon writes M into the Advanced Editor. Read it after each click; that is how Module 6 teaches the language without memorization.

## Success Check

One appended, cleaned, refreshable table where regions are `North/South/East/West`, amounts are real numbers, dates are real dates, and rows are deduplicated.


## FILE: anki/excel-flashcards.csv

```csv
Front,Back,Category
"=SUM(A1:A10)","Adds all values in range A1 to A10",Formulas
"=AVERAGE(B1:B20)","Returns the arithmetic mean of the range",Formulas
"=COUNT(C1:C100)","Counts cells containing numbers",Formulas
"=COUNTA(C1:C100)","Counts non-empty cells",Formulas
"=MAX(D1:D50)","Returns the largest value in range",Formulas
"=MIN(D1:D50)","Returns the smallest value in range",Formulas
"=SUMIF(A:A,""North"",B:B)","Sums B where A equals 'North'",Formulas
"=SUMIFS(C:C,A:A,""North"",B:B,">100")","Sums C where A='North' AND B>100",Formulas
"=COUNTIF(A:A,""Yes"")","Counts cells equal to 'Yes'",Formulas
"=IF(A1>100,""High"",""Low"")","Returns 'High' if A1>100, else 'Low'",Formulas
"=IFS(A1>90,""A"",A1>80,""B"",A1>70,""C"",TRUE,""F"")","Multiple conditions without nesting",Formulas
"=VLOOKUP(lookup_value,table,col_index,FALSE)","Vertical lookup; FALSE for exact match",Formulas
"=HLOOKUP(lookup_value,table,row_index,FALSE)","Horizontal lookup across rows",Formulas
"=INDEX(range,MATCH(lookup,lookup_range,0))","Flexible lookup; more powerful than VLOOKUP",Formulas
"=XLOOKUP(lookup,lookup_array,return_array)","Modern lookup; replaces VLOOKUP/INDEX-MATCH",Formulas
"=XLOOKUP(lookup,range,default,0,-1)","XLOOKUP with default value and search last-to-first",Formulas
"=FILTER(A2:C100,B2:B100>100)","Returns rows where column B > 100 (dynamic array)",Formulas
"=SORT(A2:D100,3,-1)","Sorts by column 3 in descending order",Formulas
"=UNIQUE(A2:A100)","Returns unique values from range",Formulas
"=LEFT(A1,5)","Extracts first 5 characters from text",Text Functions
"=RIGHT(A1,3)","Extracts last 3 characters from text",Text Functions
"=MID(A1,3,4)","Extracts 4 characters starting at position 3",Text Functions
"=LEN(A1)","Returns number of characters in cell",Text Functions
"=TRIM(A1)","Removes leading/trailing/extra spaces",Text Functions
"=UPPER(A1) / LOWER(A1) / PROPER(A1)","Converts text case",Text Functions
"=CONCATENATE(A1,"" "",B1)","Joins text strings (old method)",Text Functions
"=A1&"" ""&B1","Joins text with ampersand operator",Text Functions
"=TEXT(A1,""MM/DD/YYYY"")","Formats number as date string",Text Functions
"=SUBSTITUTE(A1,""old"",""new"")","Replaces all occurrences of text",Text Functions
"=TEXTBEFORE(A1,""@"")","Extracts text before delimiter (365/2021)",Text Functions
"=TEXTAFTER(A1,""@"")","Extracts text after delimiter (365/2021)",Text Functions
"=TODAY()","Returns current date (volatile)",Date Functions
"=NOW()","Returns current date and time (volatile)",Date Functions
"=YEAR(A1) / MONTH(A1) / DAY(A1)","Extracts year/month/day from date",Date Functions
"=DATEDIF(start,end,""Y"")","Difference between dates in years",Date Functions
"=EOMONTH(A1,0)","Last day of the month for given date",Date Functions
"=NETWORKDAYS(start,end)","Working days between two dates (excludes weekends)",Date Functions
"=NETWORKDAYS(start,end,holidays)","Working days excluding weekends and holidays",Date Functions
"=WORKDAY(A1,10)","Date that is 10 working days after A1",Date Functions
"=DATE(2024,12,31)","Creates a date from year/month/day components",Date Functions
"=IFERROR(formula,""fallback"")","Returns fallback value if formula errors",Error Handling
"=IFNA(formula,""not found"")","Returns fallback only for #N/A errors",Error Handling
"=ISERROR(A1)","Returns TRUE if A1 is any error",Error Handling
"What does #REF! mean?","Invalid cell reference (deleted row/column or bad formula)",Error Handling
"What does #VALUE! mean?","Wrong data type in formula (text where number expected)",Error Handling
"What does #N/A mean?","Lookup value not found",Error Handling
"What does #SPILL! mean?","Dynamic array can't expand (blocked by other data)",Error Handling
"What does #NAME? mean?","Excel doesn't recognize function name or named range",Error Handling
"What does #DIV/0! mean?","Division by zero",Error Handling
"What does #NUM! mean?","Invalid numeric value in formula",Error Handling
"What does #NULL! mean?","Incorrect range intersection (space instead of comma)",Error Handling
"INDEX/MATCH vs VLOOKUP: main advantage?","INDEX/MATCH can look LEFT; VLOOKUP only looks right",Lookups
"XLOOKUP vs VLOOKUP: key differences?","XLOOKUP: left lookup, default values, no col index needed",Lookups
"What is a volatile function?","Recalculates on EVERY sheet change (OFFSET INDIRECT NOW RAND)",Lookups
"What is a spilled array?","Dynamic array result that 'spills' into adjacent cells automatically",Lookups
"What does the $ sign do in formulas?","Makes reference absolute (doesn't change when copied)",Cell References
"Difference between A1 and $A$1?","A1 is relative (changes when copied); $A$1 is absolute (stays fixed)",Cell References
"What is $A1 vs A$1?","$A1 locks column only; A$1 locks row only (mixed reference)",Cell References
"What does F4 do while editing a formula?","Cycles through absolute/relative: A1 → $A$1 → A$1 → $A1",Cell References
"How to create a table?","Select data → Ctrl+T or Insert → Table",Tables
"What are structured references?","Table column names in formulas: =SUM(Table1[Sales])",Tables
"How to add a Total Row to a table?","Table Design tab → check Total Row",Tables
"What is a PivotTable?","Interactive summary tool that groups/aggregates data by categories",PivotTables
"How to create a PivotTable?","Select data → Insert → PivotTable → choose location",PivotTables
"What are the4 PivotTable field areas?","Rows Columns Values Filters",PivotTables
"How to group dates in PivotTable?","Right-click date field → Group → select Months/Quarters/Years",PivotTables
"What is a Slicer?","Visual clickable filter button for PivotTables/charts",PivotTables
"What is a Calculated Field?","Custom formula field added inside a PivotTable",PivotTables
"How to refresh a PivotTable?","Right-click → Refresh or Alt+F5 (all PivotTables: Ctrl+Alt+F5)",PivotTables
"What does Goal Seek do?","Finds input value needed to reach a target result",What-If Analysis
"What is a Data Table?","Shows how formula results change with1 or2 input variables",What-If Analysis
"What is Scenario Manager?","Saves and compares different sets of input values",What-If Analysis
"How to enable Solver?","File → Options → Add-ins → Solver Add-in → Go → check",What-If Analysis
"What is conditional formatting?","Automatically formats cells based on their values or formulas",Formatting
"How to highlight top10 values?","Home → Conditional Formatting → Top/Bottom Rules → Top10 Items",Formatting
"How to add data bars?","Home → Conditional Formatting → Data Bars → choose color",Formatting
"What is a formula-based conditional format?","Rule using custom formula (=A1>AVERAGE($A:$A)) to determine formatting",Formatting
"How to create a chart?","Select data → Insert → choose chart type",Charts
"How to add a trendline?","Click chart → Chart Design → Add Chart Element → Trendline",Charts
"What is a Sparkline?","Tiny chart inside a single cell showing data trend",Charts
"What is Power Query?","ETL tool for importing transforming and cleaning data",Power Query
"How to open Power Query?","Data tab → Get Data → choose source OR Data → From Table/Range",Power Query
"What is Merge Queries?","Joins two tables (like SQL JOIN) by matching columns",Power Query
"What is Append Queries?","Stacks tables vertically (like UNION in SQL)",Power Query
"What is Unpivot?","Converts columns to rows (wide to long format)",Power Query
"What is the M language?","Power Query's formula language for custom transformations",Power Query
"What is Power Pivot?","Data modeling tool for creating relationships between tables",Power Pivot
"What is DAX?","Data Analysis Expressions; formula language for Power Pivot/measures",Power Pivot
"What is a Measure in DAX?","Calculated value that changes based on context (not a column)",Power Pivot
"What does CALCULATE do in DAX?","Changes filter context for a calculation",Power Pivot
"What is a Data Model?","Collection of tables with relationships for PivotTable analysis",Power Pivot
"How to record a macro?","Developer tab → Record Macro → perform actions → Stop Recording",VBA
"How to open VBA Editor?","Alt+F11",VBA
"What is a Sub in VBA?","A procedure that performs actions: Sub Name() ... End Sub",VBA
"What is For Each loop?","Loops through each item in a collection: For Each ws In Worksheets",VBA
"How to handle errors in VBA?","On Error GoTo ErrorHandler ... ErrorHandler: MsgBox Err.Description",VBA
"What is a UserForm?","Custom dialog box with input controls (textboxes buttons etc.)",VBA
"What is the Range object?","Represents a cell or range: Range(""A1:B10"") or Cells(1 1)",VBA
"How to reference active sheet?","ActiveSheet or Activesheet",VBA
"What does Option Explicit do?","Forces variable declaration; catches typos",VBA
"What is a Function in VBA?","Returns a value: Function Name() As Type ... End Function",VBA
"Ctrl+C / Ctrl+V / Ctrl+X","Copy / Paste / Cut",Shortcuts
"Ctrl+Z / Ctrl+Y","Undo / Redo",Shortcuts
"Ctrl+S","Save",Shortcuts
"Ctrl+B / Ctrl+I / Ctrl+U","Bold / Italic / Underline",Shortcuts
"Ctrl+1","Format Cells dialog",Shortcuts
"Ctrl+Shift+L","Toggle AutoFilter",Shortcuts
"Ctrl+T","Create Table",Shortcuts
"Ctrl+Space","Select entire column",Shortcuts
"Shift+Space","Select entire row",Shortcuts
"Ctrl+Arrow","Jump to edge of data region",Shortcuts
"Ctrl+Home","Go to cell A1",Shortcuts
"Ctrl+End","Go to last used cell",Shortcuts
"F2","Edit active cell",Shortcuts
"F4","Toggle absolute/relative reference ($)",Shortcuts
"Alt+=","AutoSum",Shortcuts
"Alt+F11","Open VBA Editor",Shortcuts
"Ctrl+`","Show/hide formulas",Shortcuts
"Ctrl+D","Fill Down (copy cell above)",Shortcuts
"Ctrl+R","Fill Right (copy cell to left)",Shortcuts
"Ctrl+;","Insert current date",Shortcuts
"Ctrl+Shift+;","Insert current time",Shortcuts
"Ctrl+-","Delete cells/rows/columns",Shortcuts
"Ctrl+Shift++","Insert cells/rows/columns",Shortcuts
"Alt+Enter","New line within a cell",Shortcuts
"Ctrl+Page Up/Down","Switch between worksheets",Shortcuts
"Ctrl+F","Find",Shortcuts
"Ctrl+H","Find and Replace",Shortcuts
"Ctrl+G or F5","Go To dialog",Shortcuts
"Ctrl+Shift+$","Format as currency",Shortcuts
"Ctrl+Shift+%","Format as percentage",Shortcuts
"Ctrl+Shift+#","Format as date",Shortcuts
"Ctrl+Shift+@","Format as time",Shortcuts
```


## FILE: CONTRIBUTORS.md

# Contributors

- **[tempesteni](https://github.com/tempesteni)** - Original concept, curriculum direction, content review, corrections, quality control, publishing decisions, GitHub setup
- **[Hermes Agent](https://hermes-agent.nousresearch.com)** by Nous Research - Research, content compilation, dataset creation, quiz generation, cheat sheet authoring, URL verification, README and repo preparation


## FILE: path.md

# PATH.md - Project Log: Excel 0 to Hero Learning Package

## Project Overview
**Goal:** Build a comprehensive, self-contained Microsoft Excel learning package from zero to hero, sourced from verified web resources, articles, journals, and expert content. Package must be NotebookLM-compatible.

**Started:** September 10, 2026
**Completed:** September 10, 2026

---

## Process Log

### Step 1: Directory Structure Created
- Created `~/EXCEL Learn/` with 10 subdirectories:
  - `01-Fundamentals/` - Excel basics, interface, navigation
  - `02-Formulas-Functions/` - Core formulas and functions
  - `03-Formatting-Visualization/` - Charts, conditional formatting, design
  - `04-Data-Analysis/` - PivotTables, sorting, filtering, what-if
  - `05-Advanced-Functions/` - XLOOKUP, INDEX/MATCH, dynamic arrays
  - `06-Power-Query-Pivot/` - Power Query, Power Pivot, data model
  - `07-VBA-Macros/` - VBA programming, automation
  - `08-Real-World-Projects/` - Case studies and practical projects
  - `09-Sources-References/` - All sources, links, bibliography
  - `10-Study-Plan/` - Master study plan and curriculum

### Step 2: Web Research & Content Creation
- Dispatched 5 parallel research agents to create content simultaneously
- Sources used:
  - **Microsoft Official Docs:** support.microsoft.com, learn.microsoft.com
  - **Educational Sites:** ExcelJet.net, ExcelEasy.com, CorporateFinanceInstitute.com, GCFGlobal.org, Chandoo.org, TrumpExcel.com, Contextures.com, MyOnlineTrainingHub.com, ExcelMacroMastery.com, ExcelOffTheGrid.com, Spreadsheeto.com, Ablebits.com
  - **Academic:** MIT OCW, Coursera (Macquarie University), edX (Microsoft), Khan Academy, Google Scholar
  - **Books:** Excel Bible (Walkenbach), Power Query for Power BI (Webb), Financial Modeling for Dummies (Fairhurst), and more
  - **Video Channels (reference only, no downloads):** Leila Gharani, ExcelIsFun, MyOnlineTrainingHub, Chandoo, Wise Owl Tutorials, Curbal, Contextures, Kenji Explains, Kevin Stratvert, Technology for Teachers
  - **Forums:** r/excel, MrExcel, Stack Overflow, Excel Forum (Microsoft)
  - **Certifications:** MOS Associate/Expert/Master, CFI FMVA, Coursera Specialization

### Step 3: Files Created

| Module | File | Size | Lines | Description |
|--------|------|------|-------|-------------|
| 1 | `01-Fundamentals/01-excel-fundamentals.md` | 20,980 bytes | 590 | Interface, navigation, cells, references, data types |
| 2 | `02-Formulas-Functions/02-formulas-functions.md` | 32,647 bytes | 1,095 | 50+ functions: SUM, IF, VLOOKUP, TEXT, DATE, error handling |
| 3 | `03-Formatting-Visualization/03-formatting-visualization.md` | 38,085 bytes | 1,202 | Conditional formatting, charts, sparklines, tables, dashboards |
| 4 | `04-Data-Analysis/04-data-analysis.md` | 16,506 bytes | 539 | Sorting, filtering, validation, PivotTables, what-if analysis |
| 5 | `05-Advanced-Functions/05-advanced-functions.md` | 31,643 bytes | 1,211 | XLOOKUP, INDEX/MATCH, dynamic arrays, LAMBDA, LET |
| 6 | `06-Power-Query-Pivot/06-power-query-power-pivot.md` | 42,866 bytes | 1,207 | Power Query ETL, Power Pivot, DAX, data model |
| 7 | `07-VBA-Macros/07-vba-macros.md` | 55,524 bytes | 2,008 | VBA programming, macros, automation, UserForms |
| 8 | `08-Real-World-Projects/08-real-world-projects.md` | 19,251 bytes | 590 | 8 projects, industry use cases, certifications |
| 9 | `09-Sources-References/sources-references-index.md` | 8,554 bytes | 158 | Complete bibliography, video links, tools list |
| 10 | `10-Study-Plan/00-zero-to-hero-study-plan.md` | 13,138 bytes | 278 | 12-week curriculum, daily schedule, milestones |
| - | `path.md` | ~2,500 bytes | ~80 | This process log |

**Total: ~280 KB of learning content across 11 files, 8,960+ lines**

### Step 4: NotebookLM Compatibility
- All files are clean Markdown (.md) format
- Structured with proper headings (H1, H2, H3) for NotebookLM's source parser
- Tables use pipe syntax for clean parsing
- Code blocks use triple backticks
- Sources cited with URLs for verifiability
- Each module is self-contained (can be uploaded independently)
- Upload instructions included in the Study Plan

### Step 5: Packaging
- All files zipped into `EXCEL_Learn_Zero_to_Hero.zip`
- No video files downloaded (YouTube URLs referenced only)

---

## File Structure
```
EXCEL Learn/
├── 01-Fundamentals/
│   └── 01-excel-fundamentals.md
├── 02-Formulas-Functions/
│   └── 02-formulas-functions.md
├── 03-Formatting-Visualization/
│   └── 03-formatting-visualization.md
├── 04-Data-Analysis/
│   └── 04-data-analysis.md
├── 05-Advanced-Functions/
│   └── 05-advanced-functions.md
├── 06-Power-Query-Pivot/
│   └── 06-power-query-power-pivot.md
├── 07-VBA-Macros/
│   └── 07-vba-macros.md
├── 08-Real-World-Projects/
│   └── 08-real-world-projects.md
├── 09-Sources-References/
│   └── sources-references-index.md
├── 10-Study-Plan/
│   └── 00-zero-to-hero-study-plan.md
└── path.md
```

---

## How to Use with NotebookLM
1. Go to https://notebooklm.google.com
2. Create a new notebook
3. Upload all .md files as sources (drag & drop)
4. Start with the Study Plan for an overview
5. Ask questions like: "Explain PivotTables step by step" or "What VBA code do I need to automate reports?"

---

*Document generated by Hermes Agent - September 10, 2026*
