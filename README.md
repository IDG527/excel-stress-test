# Overtime form stress test

A VBA stress test for `OVERTIME_FORM_2026_MM__NAME__rev_1.2_QUERY_LINKED.xlsx`, the
Singapore field-service overtime claim form whose public holidays are loaded by the
Power Query **Holidays** from mom.gov.sg. Microsoft Copilot is used to explain the
results and to suggest new test cases.

**The query is never edited.** Every test runs on a temporary copy of the workbook.
The copy's query is only refreshed. Its M code and connection string are compared
character by character before and after the run (check `Q99`), and the original
file's size and timestamp are checked too (`Q98`).

## What's in the repository

| Path | What it is |
|---|---|
| `vba/OvertimeStressTest.bas` | The stress test. Import it into any macro-enabled workbook. |
| `docs/COPILOT_PROMPTS.md` | How Copilot is used, plus prompts for the Copilot chat pane. |
| `docs/FINDINGS.md` | What a first analysis of rev 1.2 found, verified by recalculating the real formulas. |
| `tools/oracle.py` | Python copy of the reference model the VBA compares the form against. |
| `tools/cross_check_libreoffice.py` | Checks the reference model against the real formulas using LibreOffice (no Excel needed). |

## Setup (once)

1. In Excel desktop (Windows, Microsoft 365 or 2016+), create a new blank workbook.
2. Press **Alt+F11**, then **File ▸ Import File…** and choose `vba/OvertimeStressTest.bas`.
3. Save the workbook as **StressTestRunner.xlsm** (macro-enabled).
4. Open the overtime form once and click **Data ▸ Refresh All**. This lets Excel ask,
   one time only, for permission to read mom.gov.sg. If you skip this, the refresh
   tests stop at that prompt. Close the form afterwards; you don't need to save it.

## Run

* **Alt+F8 ▸ `RunOvertimeStressTest` ▸ Run**, then choose the overtime form.
  Close other large workbooks first so the timings are clean. A full run takes about
  1–3 minutes; most of that is the 10 web refreshes.
* **Alt+F8 ▸ `RunScenariosOnly`** runs just the rows on `ST_Scenarios`. It doesn't
  refresh, fuzz, or clear your last full results.

The runner workbook gets these sheets:

| Sheet | Contents |
|---|---|
| `ST_Summary` | The verdict, PASS/FAIL/WARN counts, run details, performance table, and every FAIL and WARN |
| `ST_Results` | One line per check: ID, category, expected, actual, details |
| `ST_Timings` | Every refresh and recalculation time (ms) |
| `ST_Mismatches` | Inputs where the form disagreed with the reference model (up to 500 rows) |
| `ST_Scenarios` | Your own or Copilot-generated test rows, checked on every run |
| `ST_Copilot` | `=COPILOT()` analysis of this run, and the prompts it uses |

The statuses mean:

* **PASS**: the check passed.
* **FAIL**: the form is wrong or broken.
* **WARN**: the form works as built, but the result is risky or differs from the real-world answer.
* **SKIP**: the check couldn't run (the reason is given).
* **INFO**: for your information only.

## What gets tested

Every row of the form and every part of the file is covered:

| ID | Area | What is checked |
|---|---|---|
| S01–S18 | Structure | The four sheets are present. The engine sheet is hidden and the form is protected. Input cells are unlocked and output cells locked. Office hours `Q1`/`Q2` are valid. The headings are in place. Name, ID, month and approvers are filled in. The LOCAL/OVERSEAS list points at `M9:M10`. Data validation, conditional formatting and number formats are present (Project ID must be Text). The total formulas are correct. There are no stray validations on the hidden sheet. The print area and file size are sensible. |
| F01–F07 | Formulas | All 540 engine formulas are consistent across rows 11–70, and all 240 form cells link to the engine. There is no `#REF!` inside any formula. Office hours and the `Holidays_1[Date]` lookup are wired in. There are no error values in the file as saved (F07) or after a full recalculation (F06). |
| Q01–Q06, Q98, Q99 | Power Query (read only) | The query, its connection and the table binding exist. The M code fingerprint is recorded. The query covers only the current year and depends on the live web page. The query and the original file are unchanged at the end. |
| B01–B10 | As-found entries | The rows already in the file calculate correctly. Also checks: dates all in one month and in the holiday year, chronological order, rows with times but no date, valid LOCAL/OVERSEAS values, and Project IDs stored as text. |
| R01–R04 | Refresh stress | 10 refreshes in a row (configurable): success rate, identical data every time, timing statistics, and the form's public-holiday column following the refreshed table each time. |
| H01–H15 | Holiday table | Columns, row count, real dates, current year, Day matching Date, allowed Types, and clean names. In-lieu Mondays follow Sunday holidays. Rows are sorted with no duplicates. Fixed and moving holidays are present, and no stale rows are left under the table. Every holiday is flagged **Y** on the form and the day after it **N**. |
| K01–K22 | Known answers | 21 hand-checked cases with fixed expected hours: overnight shifts, office-hour boundaries (07:59–08:01, 17:30), midnight crossings, blank middle columns, weekday public holidays and 23:59 days. Also the sample month's totals (15.5 / 7.5 / 23). |
| L01–L07 | Known limitations | 24-hour trips, two midnight crossings, a holiday date with a time part, holidays from the previous or next year, an entry with only **From** filled, and overnight travel into the next day's office hours. |
| P01–P07 | Bad pasted input | Text in a time cell, a date typed as text, 25:00 and negative times, seconds, 255-character names, trailing-zero Project IDs, and dates before 2020. These show what happens when pasting bypasses data validation. |
| Z01–Z04 | Random fills | 200 fills × 60 rows (12,000 rows) of random dates (25% public holidays), time patterns and projects. Each row, the totals and the Project IDs are compared with the reference model. Recalculation time is measured. |
| C01–C04 | Capacity | All 60 rows filled with near-24-hour days, the totals with a full form, and 50 forced full recalculations. |
| O1xx–O4xx | Office hours | Random fills with office hours 07:30–16:30, 09:00–18:00, 07:00–19:00 and 00:00–23:59, on the copy only. Needs the sheet password in `FORM_PASSWORD` if the form has one. |
| U… | Scenarios | Each row on `ST_Scenarios`: the form against the reference model, and against your expected hours if you gave them. |

### The reference model

The VBA contains a second, independent implementation of the form's rules, written in
minutes and not copied from the Excel formulas. The random and scenario tests compare
the form with it. Its Python twin (`tools/oracle.py`) was checked against the workbook's
real formulas, recalculated in LibreOffice: **540 random rows and 21 hand cases matched
with 0 differences**. You can re-run that check without Excel:

```bash
pip install openpyxl            # and install LibreOffice Calc
python tools/cross_check_libreoffice.py path/to/OVERTIME_FORM....xlsx --runs 9
```

## Settings

Change these at the top of `OvertimeStressTest.bas`:

| Constant | Default | Meaning |
|---|---|---|
| `REFRESH_ITERATIONS` | 10 | Power Query refreshes. Each one reads mom.gov.sg. |
| `REFRESH_PAUSE_MS` | 1000 | Pause between refreshes |
| `FUZZ_ITERATIONS` | 200 | Random full-form fills (× 60 rows) |
| `RECALC_ITERATIONS` | 50 | Forced full recalculations |
| `OFFICE_HOURS_FUZZ` | 10 | Random fills for each alternative office-hours setting |
| `RANDOM_SEED` | 20260929 | Same seed = same random inputs, so failures can be reproduced |
| `FORM_PASSWORD` | `""` | Sheet password of `Engineer Name`, if any |
| `KEEP_TEST_COPY` | False | Keep the scratch copy for inspection |
| `SLOW_RECALC_MS` / `SLOW_REFRESH_MS` | 500 / 30000 | WARN thresholds |

## Copilot

`ST_Copilot` is rebuilt after each run with five `=COPILOT()` formulas:

1. An executive summary of the issues.
2. Root causes and fixes. Copilot is given the Travel and Work formulas and told not to change the query.
3. A read-out of the performance figures.
4. The pattern behind any mismatches.
5. A generator for new edge-case test rows. Paste them into `ST_Scenarios` and run `RunScenariosOnly`.

A few things to know:

* The `=COPILOT()` function needs a Microsoft 365 Copilot licence.
* If a cell shows `#NAME?`, the function isn't available to you. Copy the prompt printed
  above the cell into the Copilot chat pane instead.
* The sheet is cleared before each test starts, so Copilot never runs mid-test.
* Copilot only explains. **PASS/FAIL always comes from the VBA checks.**

See `docs/COPILOT_PROMPTS.md` for more prompts.

## Notes

* The workbook itself isn't in this repository: it contains a named employee's claim
  and is about 24 MB. Keep it wherever you normally do and select it when the macro asks.
* Mac Excel: the macro runs, but timings use the lower-resolution `Timer`. Whether the
  query can refresh depends on Mac Power Query's support for `Web.BrowserContents`. If it
  can't, R01 reports the error and the other tests use the last loaded holiday table.
