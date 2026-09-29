# Findings: OVERTIME_FORM_2026 rev 1.2 (QUERY_LINKED)

This is a first analysis of the uploaded workbook, done without Excel. The file was
opened read-only and its formulas were recalculated in LibreOffice 24.2 with test inputs.
The Power Query was decoded and read, never edited or refreshed. `RunOvertimeStressTest`
repeats each of these checks inside real Excel. The IDs in brackets are the stress-test
checks that cover each finding.

## Verified correct

* **Formula integrity.** All 540 engine formulas (`Formula - Do Not Edit`!A11:I70) follow
  the same pattern, and all 240 form cells (B, C, H, I on rows 11–70) link to the matching
  engine cell. No formula contains `#REF!`. *(F01–F03)*
* **Travel and Work rules.** An independent reference model agreed with the real formulas
  on **540 random rows** (25% public holidays, midnight crossings, blank columns) and on
  **21 hand-checked cases**, with 0 differences. *(K01–K22, Z01)*
* **Sample month.** After recalculation, the six entries in the file give Travel 15.5 h,
  Work 7.5 h and Total 23 h. *(K22, B01)*
* **Holiday table (2026).** It has 14 rows: 11 gazetted holiday days plus 3 in-lieu
  Mondays (1 Jun, 10 Aug, 9 Nov), each following a Sunday holiday. The rows are sorted,
  have no duplicates, and have clean names. Hari Raya Puasa falls on Saturday 21 Mar and
  correctly has no in-lieu day. *(H01–H13)*

## Issues found

| # | Finding | Effect | Check |
|---|---|---|---|
| 1 | **The saved file shows `#REF!`** in the Public holiday, Travel and Work columns and in all three totals. These are stale cached values; a recalculation gives the right numbers. | Anything that shows the file without recalculating it (previews, some viewers, exports) shows `#REF!` instead of hours. Open the file in Excel, press Ctrl+Alt+F9, and save. | F07 |
| 2 | **One bad pasted value breaks the whole claim total.** Text such as `abc` in a time cell turns that row, and `H71`/`I71`/`K71`, into `#VALUE!`. | Data validation only stops typed entries, not pasted ones. | P01, P02 |
| 3 | **Holidays from another year are not recognised.** The query loads `Date.Year(DateTime.LocalNow())` only. | A December claim submitted in January gets weekday rates on 25 Dec (0.5 h instead of 2 h travel + 8 h work in the test). The same applies to claims that run into the next year. | Q05, L04, L05, B04 |
| 4 | **Out-of-range pasted times are accepted.** 25:00 gives 16 h and a negative time gives 12.4 h. | The only warning is the red highlight; the hours still reach the total. | P03, P04 |
| 5 | **Trips of 24 h or more collapse.** 00:00 → 00:00 gives 0 h. A shift crossing midnight twice loses its last travel leg (4 h instead of 8 h). | Long overseas trips must be split over two rows. The form doesn't say so. | L01, L02 |
| 6 | **An incomplete entry gives 0 h silently.** A row with only **From** filled claims nothing and shows no warning. | Engineers may under-claim. | L06 |
| 7 | **A date pasted with a time** (for example 25 Dec 12:00) isn't matched as a holiday, because `COUNTIF` needs the exact date. | Weekday rates are applied to a public holiday. | L03 |
| 8 | **Next-morning office hours aren't deducted.** Weekday travel from 22:00 to 09:00 counts as 11 h, including 08:00–09:00 on the following day. | This is a policy question, not a bug. Confirm with HR which answer is intended. | L07 |
| 9 | **The header isn't filled in.** ID (`I5`) and MONTH (`K5`) are blank in the uploaded copy, and nothing forces them to be filled. | Claims can be submitted without an ID or month. | S08 |
| 10 | **The file is 24 MB.** Almost all of it is a single 1800 × 3336 PNG on `HOW TO USE (FSE Instructions)`. | Slow to open, e-mail and sync. Compressing the picture (Picture Format ▸ Compress Pictures) should shrink the file to around 1 MB. | S18 |
| 11 | **Stray drop-downs on the hidden engine sheet** (D33, E33, F33, G33, F43, G43) point at empty ranges. | Harmless clutter. | S16 |
| 12 | **The two totals are formatted differently.** `H71` is General and `I71` is `0.0`. | Cosmetic: the totals can display differently. | S15 |

## Not yet verified (needs Excel)

* **Live refresh.** Whether the `Web.BrowserContents` refresh succeeds reliably, how long
  it takes, and whether it returns the same data every time. *(R01–R04)*
* **Recalculation speed.** The recalculation time in real Excel. *(Z04, C04)*
* **Date typed as text.** In LibreOffice with a US locale, a date typed as text
  (`25/12/2026`) gave `#VALUE!`. Excel with Singapore (day/month) settings may convert
  it instead. P02 reports whichever happens.
