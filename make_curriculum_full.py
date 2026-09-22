#!/usr/bin/env python3
"""Build curriculum-full.md: the ENTIRE course in one file for AI tools.

Usage: add this single URL as a NotebookLM "Website" source (or paste into any LLM):
  https://raw.githubusercontent.com/tempesteni/Microsoft-Excel-Zero-to-hero/main/curriculum-full.md

Regenerate after edits: python3 make_curriculum_full.py
"""
import os

TITLE = "Excel Zero to Hero"
ROOT = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(ROOT, "curriculum-full.md")
SKIP_DIRS = {".git", "render", "__pycache__"}

ORDER = {"HOW-TO-USE-WITH-AI.md": 2, "cheat-sheets": 3, "quizzes": 4,
         "samples": 5, "exercises": 6, "anki": 7}

def sort_key(path):
    rel = os.path.relpath(path, ROOT)
    if rel == "README.md":
        return (0, rel)
    if rel[0].isdigit():
        return (1, rel)
    head = rel.split(os.sep)[0]
    return (ORDER.get(head, 8), rel)

files = []
for dirpath, dirs, names in os.walk(ROOT):
    dirs[:] = [d for d in dirs if d not in SKIP_DIRS]
    for n in names:
        p = os.path.join(dirpath, n)
        if (n.endswith(".md") and os.path.basename(OUT) != n) or n.endswith(".csv"):
            files.append(p)
files.sort(key=sort_key)

toc = "\n".join(f"- `{os.path.relpath(f, ROOT)}`" for f in files)
parts = [f"""# {TITLE}: Full Curriculum (Single-File AI Edition)

> One link, the complete course. Built so AI tools can see everything at once.
>
> **How to use:** in NotebookLM add this page as a single **Website** source. In ChatGPT/Claude, paste the link and ask your questions. In a RAG system, chunk by `## FILE:` headings.
>
> **Companion binaries:** the practice workbooks/decks live as separate files in the GitHub repo (this file keeps their walkthroughs and answer keys).
>
> **Regenerated from source files** by `make_curriculum_full.py`. Prefer per-file sources? Use the raw links listed in the README.

## Table of Contents

{toc}

---
"""]

for f in files:
    rel = os.path.relpath(f, ROOT)
    body = open(f, encoding="utf8").read().strip()
    if f.endswith(".csv"):
        body = "```csv\n" + body + "\n```"
    parts.append(f"\n\n## FILE: {rel}\n\n{body}\n")

open(OUT, "w", encoding="utf8").write("".join(parts))
words = sum(len(p.split()) for p in parts)
print(f"saved curriculum-full.md | {len(files)} files | ~{words:,} words | {os.path.getsize(OUT)/1024:.0f} KB")
