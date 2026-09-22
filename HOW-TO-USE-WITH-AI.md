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
