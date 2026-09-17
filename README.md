# Excel Zero to Hero

**A complete, self-paced Microsoft Excel curriculum from absolute beginner to advanced user.**

Learn Excel through structured modules, practice datasets, quizzes, and cheat sheets. Works for self-study, AI-assisted learning (NotebookLM, ChatGPT, Claude), and AI agent training pipelines.

---

## Table of Contents

1. [Start Here: How to Use This Package](#start-here-how-to-use-this-package)
2. [What's Inside: Full Content Map](#whats-inside-full-content-map)
3. [Setup Guide: NotebookLM](#setup-guide-notebooklm)
4. [Setup Guide: ChatGPT / Claude / AI Chatbots](#setup-guide-chatgpt--claude--ai-chatbots)
5. [Setup Guide: AI Agents & RAG Systems](#setup-guide-ai-agents--rag-systems)
6. [Setup Guide: Anki Flashcards](#setup-guide-anki-flashcards)
7. [The8 Modules: What You'll Learn](#the-8-modules-what-youll-learn)
8. [Supporting Materials](#supporting-materials)
9. [Who This Is For](#who-this-is-for)
10. [Certification Paths](#certification-paths)
11. [Sources & References](#sources--references)
12. [Contributing](#contributing)
13. [License](#license)

---

## Start Here: How to Use This Package

**Pick the path that matches how you want to learn:**

### Path A: I want to learn on my own (no AI)
1. Download or clone this repo
2. Open the [12-Week Study Plan](10-Study-Plan/00-zero-to-hero-study-plan.md)
3. Follow the daily schedule: read the module, do the exercises, take the quiz
4. Use the [Progress Tracker](10-Study-Plan/progress-tracker.md) to check off skills as you learn them
5. Print the [Cheat Sheets](cheat-sheets/) and keep them at your desk

### Path B: I want AI to teach me (NotebookLM) <sup>Recommended</sup>
1. Go to [notebooklm.google.com](https://notebooklm.google.com)
2. Create a new notebook called "Excel Zero to Hero"
3. Upload all the `.md` files (see [detailed setup below](#setup-guide-notebooklm))
4. Start asking questions like:
   - *"Explain VLOOKUP to me like I'm10 years old"*
   - *"Quiz me on PivotTables from Module 4 and grade my answers"*
   - *"I have 1 hour today, what should I study?"*
   - *"I work in HR, how do I use these Excel skills for headcount dashboards?"*
5. Use the **Audio Overview** feature to listen to summaries while commuting

### Path C: I want ChatGPT or Claude to tutor me
1. Open a new chat in ChatGPT or Claude
2. Copy-paste the setup prompt from [the guide below](#setup-guide-chatgpt--claude--ai-chatbots)
3. Then paste a module's content and start learning
4. Ask the AI to quiz you, explain concepts, or build exercises

### Path D: I'm building an AI agent or training system
1. All content is structured Markdown, ready for RAG ingestion
2. See [AI Agent Setup](#setup-guide-ai-agents--rag-systems) for chunking strategy, metadata schema, and sample code
3. Quiz files provide ready-made assessment content
4. Practice datasets give your agent realistic data to work with

---

## What's Inside: Full Content Map

### The Folder Structure

```
EXCEL Learn/
│
├── 01-Fundamentals/                    # Module 1 - Beginner
│   └── 01-excel-fundamentals.md
│
├── 02-Formulas-Functions/              # Module 2 - Beginner+
│   └── 02-formulas-functions.md
│
├── 03-Formatting-Visualization/        # Module 3 - Beginner+
│   └── 03-formatting-visualization.md
│
├── 04-Data-Analysis/                   # Module 4 - Intermediate
│   └── 04-data-analysis.md
│
├── 05-Advanced-Functions/              # Module 5 - Intermediate+
│   └── 05-advanced-functions.md
│
├── 06-Power-Query-Pivot/               # Module 6 - Advanced
│   └── 06-power-query-power-pivot.md
│
├── 07-VBA-Macros/                      # Module 7 - Advanced
│   └── 07-vba-macros.md
│
├── 08-Real-World-Projects/             # Module 8 - All Levels
│   └── 08-real-world-projects.md
│
├── 09-Sources-References/              # Bibliography
│   └── sources-references-index.md
│
├── 10-Study-Plan/                      # Curriculum & Tracking
│   ├── 00-zero-to-hero-study-plan.md   # 12-week schedule
│   └── progress-tracker.md             # Checkbox checklist
│
├── cheat-sheets/                       # Printable Quick References
│   ├── functions-cheat-sheet.md        # 50 functions with syntax & examples
│   ├── keyboard-shortcuts-cheat-sheet.md # 65+ shortcuts by category
│   ├── pivottable-cheat-sheet.md       # PivotTable reference card
│   └── power-query-cheat-sheet.md      # Top 20 M functions
│
├── quizzes/                            # Self-Assessment
│   ├── quiz-01-fundamentals.md         # 15 MC + 5 formula challenges
│   ├── quiz-02-formulas.md
│   ├── quiz-03-visualization.md
│   ├── quiz-04-analysis.md
│   ├── quiz-05-advanced-lookup.md
│   ├── quiz-06-power-tools.md
│   ├── quiz-07-vba.md
│   └── quiz-08-projects.md
│
├── exercises/                          # Practice Data
│   ├── sales-data-practice.xlsx        # 200 sales rows + 50 employee rows
│   └── create_dataset.py               # Script that generated the dataset
│
├── anki/                               # Spaced Repetition
│   └── excel-flashcards.csv            # 100+ cards, importable to Anki
│
├── HOW-TO-USE-WITH-AI.md              # Detailed AI platform guides
└── README.md                           # This file
```

### Content Summary

| Category | Files | What You Get |
|----------|-------|-------------|
| **Learning Modules** | 8 files, ~350 KB | Complete curriculum from basics to VBA |
| **Cheat Sheets** | 4 files, ~26 KB | Printable desk references |
| **Quizzes** | 8 files | 120 multiple-choice + 40 formula challenges with answer keys |
| **Practice Data** | 1 .xlsx file | Realistic sales & employee data with intentional cleaning issues |
| **Flashcards** | 1 .csv file | 100+ Anki-compatible cards covering functions, shortcuts, concepts |
| **Study Plan** | 2 files | 12-week daily schedule + progress tracker checklist |
| **Sources** | 1 file | 60+ verified references with URLs |
| **AI Guides** | 1 file | Platform-specific setup and prompts |

---

## Setup Guide: NotebookLM

**What is NotebookLM?** Google's free AI research tool. Upload documents, and it answers questions using ONLY your uploaded content. No hallucination from the web.

### Step-by-Step Setup

**Step 1:** Go to [notebooklm.google.com](https://notebooklm.google.com) and sign in with your Google account

**Step 2:** Click **"Create new notebook"** and name it `Excel Zero to Hero`

**Step 3:** Upload your sources in this order (click "Add Source" for each):

| Upload Order | File to Upload | Why This Order |
|-------------|---------------|---------------|
| 1st | `10-Study-Plan/00-zero-to-hero-study-plan.md` | Gives overview of the whole curriculum |
| 2nd | `01-Fundamentals/01-excel-fundamentals.md` | Foundation context |
| 3rd | `02-Formulas-Functions/02-formulas-functions.md` | Builds on Module 1 |
| 4th | `03-Formatting-Visualization/03-formatting-visualization.md` | Sequential |
| 5th | `04-Data-Analysis/04-data-analysis.md` | Sequential |
| 6th | `05-Advanced-Functions/05-advanced-functions.md` | Sequential |
| 7th | `06-Power-Query-Pivot/06-power-query-power-pivot.md` | Sequential |
| 8th | `07-VBA-Macros/07-vba-macros.md` | Sequential |
| 9th | `08-Real-World-Projects/08-real-world-projects.md` | Projects apply everything |
| 10th | `09-Sources-References/sources-references-index.md` | Reference context |
| 11th | `10-Study-Plan/progress-tracker.md` | Tracking context |
| 12th-19th | `quizzes/quiz-01.md` through `quiz-08.md` | Assessment content |
| 20th-23rd | `cheat-sheets/*.md` | Quick reference context |

**Step 4:** Start asking questions in the chat panel on the right

### Copy-Paste Prompts for NotebookLM

**Learning a new topic:**
```
I'm a complete beginner. Explain conditional formatting from Module 3 like I'm 10 years old. Give me 3 real-world examples I can try right now.
```

**Getting quizzed:**
```
Quiz me on Module 2 (Formulas & Functions). Ask me 5 questions one at a time. Wait for my answer before giving the next question. After all 5, tell me my score and what to review.
```

**Study planning:**
```
I just finished Week 3 of the study plan. I have 2 hours today. What should I focus on? Create a minute-by-minute schedule for me.
```

**Real-world application:**
```
I work in marketing. Show me how to use PivotTables (Module 4) and charts (Module 3) to build a campaign performance dashboard. Give me step-by-step instructions.
```

**Review before quiz:**
```
I'm about to take quiz-04 (Data Analysis). Give me a quick 5-minute refresher of the most important concepts I need to remember.
```

**Audio learning:**
After uploading, click **"Audio Overview"** in the left panel to generate a podcast-style summary. Listen while commuting or exercising.

---

## Setup Guide: ChatGPT / Claude / AI Chatbots

### Quick Start Prompt (Copy-Paste This First)

```
I'm learning Microsoft Excel from a structured curriculum. I'll share modules one at a time. For each module I share:

1. Summarize the 5 most important concepts
2. Explain the hardest concept like I'm a beginner
3. Create 3 quiz questions and wait for my answers
4. Tell me what I should focus on next

When I say "quiz me", give me 10 questions and grade my answers.
When I say "practice", create a hands-on exercise I can do.
When I say "explain [topic]", break it down step by step.

Ready? Here's Module 1: Fundamentals.

[PASTE THE CONTENT OF 01-Fundamentals/01-excel-fundamentals.md HERE]
```

### Useful Follow-Up Prompts

**After finishing a module:**
```
I finished Module [X]. Quiz me with 10 questions. Mix multiple-choice and "write the formula" questions. Give me a score at the end.
```

**When stuck on a concept:**
```
I don't understand XLOOKUP. Explain it 3 different ways:
1. Like I'm 10 years old
2. With a real business example
3. By comparing it to VLOOKUP which I already know
```

**For practice:**
```
I just learned about PivotTables. Create a practice exercise:
- Give me a small dataset (20 rows)
- Tell me 5 things I should be able to do with it
- Then give me the solutions
```

**For interview prep:**
```
I have a job interview that requires Excel. Based on my learning materials:
1. What are the top 10 Excel interview questions?
2. Give me model answers
3. Create a 2-day crash study plan
```

**For certification prep:**
```
I'm preparing for MOS Excel certification. Based on the content I've shared:
1. Create a 20-question practice test in MOS style
2. Score my answers
3. Tell me which modules I need to review
```

---

## Setup Guide: AI Agents & RAG Systems

This package is designed for AI-native use. Every file is structured Markdown with clear headings, making it compatible with RAG systems, vector databases, and AI tutoring agents.

### Chunking Strategy
```
Source format: Markdown (.md)
Split by: ## headings (each section is self-contained)
Chunk size: ~500-1500 tokens per chunk
Overlap: Include parent heading as context in each chunk
```

### Metadata Schema
```json
{
  "module_number": "01-08",
  "difficulty": "beginner | intermediate | advanced",
  "topic": "formulas | pivot | vba | formatting | analysis | power_query",
  "content_type": "lesson | exercise | quiz | cheat_sheet | study_plan",
  "file_path": "relative/path/to/file.md"
}
```

### Embedding & Retrieval
```
Recommended model: text-embedding-3-small or all-MiniLM-L6-v2
Retrieval: Top-k=5 with metadata filter on module_number or difficulty
Reranking: Optional cross-encoder for relevance
```

### Agent Modes

| Mode | What It Does | Input |
|------|-------------|-------|
| **Quiz** | Generates questions from quiz files, grades answers | Quiz file + user answer |
| **Explain** | Breaks down concepts from module content | Module file + topic |
| **Practice** | Generates exercises using dataset structure | Module file + difficulty |
| **Progress** | Tracks completed modules, suggests next steps | Progress tracker + quiz scores |
| **Cheat Sheet** | Quick reference lookups | Cheat sheet file + query |

### Sample Agent Code (Python pseudocode)
```python
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

---

## Setup Guide: Anki Flashcards

### How to Import

1. Open Anki (download free from [apps.ankiweb.net](https://apps.ankiweb.net/))
2. Click **File → Import**
3. Select `anki/excel-flashcards.csv`
4. Set field mapping:
   - Field 1 → Front
   - Field 2 → Back
   - Field 3 → Tags (Category)
5. Click **Import**

### What's in the Deck

| Category | Card Count | Examples |
|----------|-----------|---------|
| Formulas | 30 | SUMIF syntax, VLOOKUP parameters, IF nesting |
| Text Functions | 12 | LEFT, RIGHT, MID, TRIM, CONCATENATE |
| Date Functions | 10 | TODAY, DATEDIF, EOMONTH, NETWORKDAYS |
| Error Handling | 10 | #REF!, #VALUE!, #N/A, IFERROR, IFNA |
| Lookups | 8 | INDEX/MATCH vs VLOOKUP, XLOOKUP, volatile functions |
| Cell References | 4 | Absolute vs relative, $ sign, F4 key |
| Tables | 4 | Structured references, Total Row |
| PivotTables | 8 | Creating, field areas, grouping, slicers |
| What-If Analysis | 4 | Goal Seek, Data Tables, Scenario Manager |
| Formatting | 4 | Conditional formatting, data bars |
| Charts | 3 | Creating, trendlines, Sparklines |
| Power Query | 7 | What it is, Merge vs Append, Unpivot, M language |
| Power Pivot | 5 | DAX, Measures, CALCULATE, Data Model |
| VBA | 10 | Recording macros, Sub/Function, loops, UserForms |
| Shortcuts | 30 | Ctrl+C, F4, Alt+=, Ctrl+T, and more |

**Study tip:** Start with5 new cards per day. Anki's spaced repetition will automatically show you cards you struggle with more often.

---

## The 8 Modules: What You'll Learn

### Module 1: Fundamentals (Beginner)
Interface tour, cell navigation, data entry, cell references (relative, absolute, mixed), file formats (.xlsx, .csv, .xlsb)

### Module 2: Formulas & Functions (Beginner+)
SUM, AVERAGE, COUNT, IF, Nested IF, SUMIF/SUMIFS, COUNTIF/COUNTIFS, VLOOKUP, text functions (LEFT, RIGHT, MID, TRIM), date functions (TODAY, DATEDIF, EOMONTH), error handling (IFERROR)

### Module 3: Formatting & Visualization (Beginner+)
Number formats, conditional formatting (highlight rules, color scales, data bars, icon sets), Excel Tables, structured references, charts (Column, Bar, Line, Pie), chart formatting, Sparklines, Slicers

### Module 4: Data Analysis (Intermediate)
Sorting (single/multi-level), AutoFilter, Advanced Filter, data validation (drop-down lists, custom formulas), PivotTables (creating, field areas, grouping, calculated fields, PivotCharts), Goal Seek, Scenario Manager, Data Tables, Subtotals

### Module 5: Advanced Functions (Intermediate+)
INDEX/MATCH (single and multi-criteria), XLOOKUP (all parameters), OFFSET, INDIRECT, dynamic arrays (FILTER, SORT, UNIQUE, SEQUENCE), LET, LAMBDA, TEXTBEFORE, TEXTAFTER

### Module 6: Power Query & Power Pivot (Advanced)
Importing data (CSV, web, databases), Query Editor, transformations (filter, split, merge, pivot/unpivot), Merge Queries (JOINs), Append Queries, M language, Data Model, relationships, DAX (calculated columns, measures, CALCULATE), KPIs, hierarchies

### Module 7: VBA & Macros (Advanced)
Recording macros, VBA Editor, Sub/Function, variables, MsgBox/InputBox, Range objects, loops (For Next, For Each), If/Select Case, error handling, working with sheets/workbooks, UserForms

### Module 8: Real-World Projects (All Levels)
8 complete projects applying all skills:
1. Personal Budget Tracker (SUM, SUMIF, charts)
2. Sales Dashboard (PivotTables, slicers, KPIs)
3. Inventory Management System (XLOOKUP, data validation, VBA)
4. Employee Schedule & Leave Tracker (date functions, NETWORKDAYS)
5. Invoice Generator (templates, named ranges, print setup)
6. Data Cleaning Pipeline (Power Query, text functions)
7. Financial Model (NPV, IRR, what-if analysis, data tables)
8. Automated Report Generator (VBA, PivotTables, charts)

---

## Supporting Materials

### Cheat Sheets (Print These)
| Sheet | Content | Best For |
|-------|---------|----------|
| [functions-cheat-sheet.md](cheat-sheets/functions-cheat-sheet.md) | 50 functions with syntax and examples | Desk reference while working |
| [keyboard-shortcuts-cheat-sheet.md](cheat-sheets/keyboard-shortcuts-cheat-sheet.md) | 65+ shortcuts organized by category | Speed up your workflow |
| [pivottable-cheat-sheet.md](cheat-sheets/pivottable-cheat-sheet.md) | PivotTable creation, fields, grouping, calculated fields | Data analysis tasks |
| [power-query-cheat-sheet.md](cheat-sheets/power-query-cheat-sheet.md) | Top 20 M functions and transformations | Power Query work |

### Quizzes (Test Yourself)
Each quiz has **15 multiple-choice questions** (3 points each) + **5 fill-in-the-formula challenges** (5 points each) = **70 points total**.

| Score | Meaning |
|-------|---------|
| 90%+ (63+) | Ready to advance to next module |
| 70-89% (49-62) | Review weak areas, then advance |
| Below 70% (<49) | Restudy the module before moving on |

### Practice Dataset
[sales-data-practice.xlsx](exercises/sales-data-practice.xlsx) contains:
- **Sheet 1 (Sales-Data):** 200 rows of sales data with formulas (Total, Commission)
- **Sheet 2 (Employee-Data):** 50 rows of employee records
- **Built-in issues for cleaning practice:** 2 duplicate rows, 3 cells with extra spaces, 1 blank cell

Use this dataset to practice PivotTables, sorting, filtering, charts, and data cleaning.

---

## Who This Is For

- **Complete beginners** with zero Excel experience
- **Career switchers** needing Excel for data, finance, or operations roles
- **Students** preparing for MOS (Microsoft Office Specialist) certification
- **Self-learners** who prefer structured content over scattered YouTube videos
- **AI engineers** building Excel training agents, chatbots, or RAG systems
- **Teachers and trainers** looking for curriculum to adapt for their classes
- **Professionals** who use Excel daily but want to level up to advanced features

---

## Certification Paths

After completing this curriculum, you'll be prepared for:

| Certification | Provider | What It Tests | Difficulty |
|--------------|----------|---------------|------------|
| MOS: Excel Associate | Microsoft / Certiport | Core functionality, formulas, tables, charts | Beginner-Intermediate |
| MOS: Excel Expert | Microsoft / Certiport | Advanced formulas, PivotTables, macros | Advanced |
| Excel Skills for Business | Coursera / Macquarie University | Comprehensive Excel for business | All Levels |
| CFI Financial Modeling | Corporate Finance Institute | Financial modeling in Excel | Advanced |

---

## Sources & References

This curriculum draws from **60+ verified sources** including:

- **Official:** Microsoft documentation, Microsoft Learn
- **Educational sites:** ExcelJet, Excel Easy, Chandoo, TrumpExcel, Contextures, Ablebits, Spreadsheeto
- **YouTube channels:** Leila Gharani, ExcelIsFun (Mike Girvin), MyOnlineTrainingHub, Wise Owl Tutorials, Curbal, Kenji Explains, Kevin Stratvert
- **Academic:** Coursera (Macquarie University), edX, MIT OpenCourseWare, Khan Academy
- **Books:** Excel Bible (Walkenbach), Power Programming with VBA (Alexander), Power Query for Power BI (Webb)

See the full bibliography: [Sources & References Index](09-Sources-References/sources-references-index.md)

---

## Contributing

Found an error? Want to add content?

1. Fork this repo
2. Make your changes (all content is Markdown, edit any file directly)
3. Submit a pull request

Contributions welcome: new modules, better examples, additional practice datasets, translations, corrections.

---

## License

This project is licensed under the **Creative Commons Attribution 4.0 International License (CC BY 4.0)**.

You are free to:
- **Share** - copy and redistribute the material in any format
- **Adapt** - remix, transform, and build upon the material for any purpose

Under the following terms:
- **Attribution** - You must give appropriate credit and link to this repo

---

## Author

Created by [tempesteni](https://github.com/tempesteni)

**Stats:**
- Estimated study time: 80-120 hours over 12 weeks
- Total content: ~612 KB across 30 files
- Prerequisites: Basic computer literacy, access to Microsoft Excel (2019, 2021, or 365)
