# Using Microsoft Copilot with the stress test

Copilot is used in three places. In all three, the **verdict comes from the VBA**.
Copilot explains the results and suggests ideas, but a Copilot answer never turns a
FAIL into a PASS.

| Where | What Copilot does |
|---|---|
| `ST_Copilot` sheet (`=COPILOT()` formulas written by the macro) | Summarises the run, suggests fixes, reads the performance figures, finds patterns in mismatches, and generates new test rows. |
| Copilot chat pane in Excel (**Home ▸ Copilot**) | The same prompts, if the `=COPILOT()` function isn't available to you, plus the prompts below. |
| Copilot in the VBA editor or Copilot chat | Explains or adapts `OvertimeStressTest.bas` (see "Working on the macro" below). |

## Requirements

* **`=COPILOT()` function:** a Microsoft 365 Copilot licence, in Excel for Windows or Mac
  on a channel where the function is enabled. If a cell shows `#NAME?`, use the chat pane.
* Each `=COPILOT()` cell uses your Copilot allowance whenever it recalculates. The macro
  clears `ST_Copilot` before a run starts and rebuilds it only at the end, so the tests
  never call Copilot.
* Don't paste the overtime workbook's personal data (names, IDs) into prompts outside your
  organisation's Copilot tenancy.

## Loop: Copilot writes test cases, VBA checks them

1. After a run, look at block **5. New test scenarios** on `ST_Copilot`. It spills a
   6-column table: Description, Date, From, Until/From, Until/From, Until.
2. Copy the spilled range, then **Paste ▸ Values** into `ST_Scenarios` column **B**, on
   the first free row (row 5 or below).
3. Run **`RunScenariosOnly`**. Each row gets **Form travel / work**, **Reference travel /
   work**, and one of these statuses:
   * **PASS**: the form matches the reference model.
   * **FAIL**: the form doesn't match the reference model. A real bug.
   * **REVIEW**: the form matches the reference model but not the hours you typed in
     columns H–I. Your expectation or the policy needs a second look.
   * **INVALID**: a date or time couldn't be read.
4. Keep the useful rows. They run again on every full stress test.

## Prompts for the Copilot chat pane

Select the range named in brackets before you send each prompt.

**Summarise the run** *(select `ST_Summary`, the ISSUES table)*
> These are the FAIL and WARN results of an automated stress test of an Excel overtime-claim form. Summarise them in at most 6 bullet points for a manager: what is broken, what is risky, what is fine. Put FAIL items first.

**Explain one failure** *(select the row on `ST_Results`)*
> Explain this test result in plain English. What input caused it, what did the form return, what should it have returned, and what is the smallest change to the workbook that would fix it? Do not suggest changing the Power Query.

**Mismatch patterns** *(select `ST_Mismatches`)*
> Each row is an input where the overtime form's result differed from an independent reference calculation. What do the failing inputs have in common? Consider weekday or weekend, public holiday, midnight crossing, and which time columns are blank.

**New edge cases** *(select the holiday list on `ST_Copilot`, columns H:I)*
> Create 15 edge-case rows for an overtime form as a table with columns Description, Date (yyyy-mm-dd), From, Until/From, Until/From, Until (24-hour hh:mm, blank when unused). The times are travel, then work, then travel, and may cross midnight once. On weekdays, only time outside 08:00–17:30 counts. Weekends and the listed public holidays count in full. Include office-hour boundaries, midnight, and rows with only some columns filled.

**Performance** *(select the PERFORMANCE table on `ST_Summary`)*
> These are millisecond timings from repeated Power Query refreshes and recalculations of an overtime form. Is this acceptable for an engineer filling in a monthly claim? Point out outliers.

**Explain the overtime formula** *(select `ST_Copilot!M6`, the Travel formula text)*
> Explain this Excel formula step by step, as a list of rules an engineer can check by hand. Then give three example inputs that exercise different branches.

## Working on the macro

Paste a procedure from `OvertimeStressTest.bas` into Copilot chat with one of these:

* *"Explain what this VBA procedure tests and what would make it report FAIL."*
* *"Add a check to this VBA procedure that ___. Keep the LogResult call style and never write to the Power Query or to the original workbook."*
* *"This VBA line raises run-time error ___ in Excel ___. Why, and how do I fix it without changing the test's meaning?"*

After Copilot changes the macro, run the full stress test once. Q98 and Q99 must still
PASS, which proves the query and the original file were not touched.
