# Walkthrough: sample-01-data-entry-practice.xlsx

Companion to **Module 1: Fundamentals**.

## What to Notice

1. **The mess is realistic on purpose.** Order IDs mix `YC-0008` and `YC-3`; dates arrive as `2026-01-06`, `03/13/2026`, `2/14/26`, and `March 5 2026`. This is exactly what pasted exports look like in real jobs.
2. **Formulas are already waiting.** Column G carries `=E*F` formulas on every row, including the rows with missing inputs. Watch how they sit quiet until the data is fixed, then come alive. (Module 2)
3. **The Format-Me sheet stores numbers as text.** Totals refuse to add until you fix it. The fast fix is Data > Text to Columns > Finish, which converts in place. (Module 1)
4. **Green triangles are free lessons.** Every warning triangle marks a place where Excel suspects your data lies.

## Success Check

No green triangles anywhere, every Order ID matches `YC-0001` format, and column G totals are correct to the cent.
