# Walkthrough: sample-05-lookup-challenges.xlsx

Companion to **Module 5: Advanced Functions**.

## What to Notice

1. **Two tables that want to be one.** Orders knows SKUs but not names or prices; Products knows everything about SKUs. Lookups are just "borrow the missing columns". (Module 5)
2. **XLOOKUP is the hero, INDEX/MATCH is the resume.** Task 1 vs task 4 solve the same problem both ways so you can feel the difference in effort and readability.
3. **`if_not_found` is not optional in real life.** Task 2 bakes in `Discontinued` because missing SKUs happen. Task 3 then forces you to handle the text result gracefully with ISNUMBER. That pairing is how production spreadsheets stay honest.
4. **Tasks 5-7 graduate to array thinking**: SUMPRODUCT with XLOOKUP-returned arrays, and LET for one-pass readability. Task 8's dynamic-array spill is the future-proof version of task 1.

## Success Check

All challenge answers match the Answer Key, and you can explain why VLOOKUP's approximate match would be wrong here (preview: `exercise-fix-this-workbook.xlsx` issue 5).
