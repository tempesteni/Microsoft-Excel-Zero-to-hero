# Walkthrough: sample-06-powerquery-messy.xlsx

Companion to **Module 6: Power Query & Power Pivot**.

## What to Notice

1. **Three sheets, three different messes.** The Jan/Feb/Mar exports disagree on headers (`Date` vs `date ` vs `Transaction Date`), region casing (`North`/`north`), number storage (`" 830"`, `"1,450"`), and date formats. This is the real world of monthly exports. (Module 6)
2. **Each sheet carries the same 3 sins**: trailing spaces in names, exact duplicate rows, and blank separator rows. Once you build one cleaning query, it fixes every future month automatically.
3. **The point of Power Query is repeatability.** Your cleaned output must survive the test in task 8: edit Jan and Refresh All, and the cleaned table updates with zero rework. If you cleaned by hand, it will not.
4. **M language appears after you click.** Every step you build in the ribbon writes M into the Advanced Editor. Read it after each click; that is how Module 6 teaches the language without memorization.

## Success Check

One appended, cleaned, refreshable table where regions are `North/South/East/West`, amounts are real numbers, dates are real dates, and rows are deduplicated.
