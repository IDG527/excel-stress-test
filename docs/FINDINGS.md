# Findings: OVERTIME_FORM_MONTH_MANUAL_INPUT_CLEAN_1

This is an analysis of the version with the month entry (`K5`) and input pop-ups, done
without Excel. The file was opened read-only. Its formulas and validation rules were
recalculated in LibreOffice 24.2 with test inputs, and it was compared cell by cell with
rev 1.2 (QUERY_LINKED). `RunOvertimeStressTest` repeats each check inside real Excel. The
IDs in brackets are the stress-test checks that cover each finding.

## What changed since rev 1.2

| Change | Where |
|---|---|
| Month entry: `K5` must be a full month name in capitals and a 4-digit year (e.g. `OCTOBER 2026`). | `Engineer Name`!K5 validation |
| Date rule: a date must fall in the claim month, or on the last day of the month before (for overnight jobs). | A11:A70 validation, which reads the month from `Formula - Do Not Edit`!J6 |
| Input pop-ups on Name, Employee ID, Month, Date, the four time columns and "Submitted by". Error alerts on the date, time and month cells. | B5, I5, K5, A11:A70, D11:G70, C77 |
| The public-holiday flag also matches `TEXT(date,"yyyy-mm-dd")`, so a date carrying a time of day is still recognised. | Engine column C |
| "Submitted By" moved from row 76 to row 77, and its name cell C77 is now an input. | A77, C77 |
| The `SG Public Holidays` sheet is hidden, and the workbook forces a full recalculation on every change. | sheet state, `forceFullCalc` |
| The form sheet is **no longer protected** (rev 1.2 was password-protected). | `Engineer Name` |
| Unchanged: the Power Query **Holidays** (its M code is identical byte for byte), the connection, and the Travel / Work formulas. | |

## Verified correct

* **Travel and Work rules.** An independent reference model matched the real formulas
  on 360 random rows in this version (540 in rev 1.2) and on 21 hand-checked cases, with
  0 differences. *(K01–K22, Z01)*
* **Formula integrity.** All 540 engine formulas and all 240 form links are consistent.
  The saved file has no cached `#REF!` any more. *(F01–F07)*
* **Month entry rule.** It accepts `OCTOBER 2026` and `MAY 2026`, and rejects `October 2026`,
  `OCT 2026`, `OCTOBER 26`, `2026 OCTOBER`, double spaces, and leading or trailing spaces. *(V04)*
* **A holiday date pasted with a time** (25 Dec 12:00) is now flagged **Y** and paid at
  holiday rates. This was issue 7 in rev 1.2 and is now fixed. *(L03)*

## Issues found

| # | Finding | Effect | Check |
|---|---|---|---|
| 1 | **Blocker: every date is rejected.** The date rule compares each date with `'Formula - Do Not Edit'!J6`, but J6 is empty. With `K5 = OCTOBER 2026`, 5 Oct and 30 Sep are rejected like any other date. Once J6 is filled from K5, 5 Oct and 30 Sep are accepted and 29 Sep and 1 Nov rejected, as intended. | Engineers can't type any date into column A. | V06, V07 |
| 2 | **The form sheet is not protected.** | Anyone can type over the Day, Public holiday, Travel and Work formulas or the totals. The input cells are already unlocked, so protecting the sheet again is safe. | S03 |
| 3 | **LOCAL / OVERSEAS has no pop-up and its error alert is off.** Project ID and Vessel have no pop-up either. | Anything typed in column J is accepted. | V02 |
| 4 | **The month rule doesn't check that the year makes sense.** `OCTOBER 0000`, `OCTOBER 9999` and `OCTOBER -202` are accepted. | A typo such as `OCTOBER 2062` passes. | V05 |
| 5 | **Only 00:00 can come "before" the previous time.** Typing 21:00 → 02:00 is blocked; the engineer must end the row at 00:00 and continue on a new row. The pop-up says this; the formulas themselves would handle 02:00. | A usability point, not an error. | V09 |
| 6 | **The holiday list covers only the current year.** The query uses `Date.Year(DateTime.LocalNow())`. | In early January, before the query has loaded the new year, new-year holidays are paid at normal rates. In early January the query may also stop with "No public holiday records were found". A December claim finished in January gets normal rates on 25 Dec. | Q05, Y01, Y13, L04, L05 |
| 7 | **One bad pasted value breaks the whole claim total.** Text in a time cell turns the row and `H71`/`I71`/`K71` into `#VALUE!`. | Validation only stops typed entries, not pasted ones. | P01, P02 |
| 8 | **Out-of-range pasted times are accepted.** 25:00 gives 16 h and a negative time gives 12.4 h. | The only warning is the red highlight. | P03, P04 |
| 9 | **Trips of 24 h or more collapse.** 00:00 → 00:00 gives 0 h, and a second midnight crossing loses the last travel leg. | Long trips must be split over rows. | L01, L02 |
| 10 | **A row with only "From" filled claims 0 h with no warning.** | Engineers may under-claim. | L06 |
| 11 | **Next-morning office hours aren't deducted.** Weekday travel from 22:00 to 09:00 counts as 11 h. | A policy question for HR. | L07 |
| 12 | **The file is still 24 MB**, almost all of it one 1800 × 3336 PNG on the instructions sheet. | Slow to open, e-mail and sync. Compress the picture. | S18 |
| 13 | **Leftovers:** stray drop-downs on the engine sheet (D33:G33, F43:G43), and `H71` formatted General while `I71` is `0.0`. | Cosmetic. | S15, S16 |

### Suggested formula for J6 (issue 1)

Put this in `'Formula - Do Not Edit'!J6`. It turns `OCTOBER 2026` into 1-Oct-2026, and
gives a blank while K5 is empty or invalid. It was checked in LibreOffice with the
form's own date rule:

```
=IFERROR(DATE(VALUE(RIGHT('Engineer Name'!K5,4)),MATCH(LEFT('Engineer Name'!K5,FIND(" ",'Engineer Name'!K5)-1),{"JANUARY","FEBRUARY","MARCH","APRIL","MAY","JUNE","JULY","AUGUST","SEPTEMBER","OCTOBER","NOVEMBER","DECEMBER"},0),1),"")
```

## Update: OVERTIME_FORM_MONTH_CAPITALIZED_ONLY

The file was compared part by part with the version above, and the rules were re-run in LibreOffice.

* **The K5 capitals rule is unchanged.** It is identical, character for character, to
  the previous version, which already rejected `October 2026`, `october 2026`, `OCTOBEr 2026`,
  `Oct 2026` and `OCT 2026`. All 12 wrong formats are still rejected. *(V04)*
* **Fixed since the previous version:** the form sheet is password-protected again, and the
  engine sheet is hidden again *(S02, S03)*. Objects stay editable, so engineers can still
  place a signature picture as the C77 pop-up asks. All input cells, C77 included, are unlocked.
* **Still a blocker: J6 is empty**, so every date is rejected, including 1 Oct and 31 Oct with
  `K5 = OCTOBER 2026`. With the suggested J6 formula, 1 Oct, 31 Oct and 30 Sep are accepted and
  29 Sep, 1 Nov and 15 Oct 2025 rejected. *(V06, V07)*
* **More impossible years pass the K5 rule:** besides `0000`, `9999` and `-202`, `OCTOBER 2.26` and
  `OCTOBER 1E03` are accepted, because the rule checks for four characters that `VALUE()` can read
  rather than four digits. *(V05)*
* **K5 is empty in this copy.** Until a month is entered, every date is rejected (once J6 is fixed),
  so the date error message should also say "enter the month in K5 first".
* **Office hours (with the sheet password):** Travel and Work stay correct when the office
  hours in Q1/Q2 change: 07:30–16:30, 09:00–18:00, 07:00–19:00 and 00:00–23:59, 120 random rows
  each, 0 differences. Q1/Q2 are locked and in a hidden column, so only someone with the password
  can change them. In the VBA, put the password in `FORM_PASSWORD` on your own copy of the module
  (it is not stored in this repository). *(O1–O4)*
* **Unchanged:** the Power Query (byte for byte), all Travel/Work formulas (240 more random rows,
  0 differences), and the open issues 3 and 6–13 above.

## Keeping J6 empty: point the date rule at K5

K5 stays a text month (`OCTOBER 2026`), and `'Formula - Do Not Edit'!J6` stays empty. For dates
to be accepted, the A11:A70 rule then has to read K5 itself. Replace the custom rule on A11:A70
(select A11:A70 ▸ Data ▸ Data Validation ▸ Custom) with:

```
=IF(A11="",TRUE,IFERROR(OR(YEAR(A11)*12+MONTH(A11)=(RIGHT($K$5,4)*12+(FIND(LEFT($K$5,3),"JANFEBMARAPRMAYJUNJULAUGSEPOCTNOVDEC")+2)/3),YEAR(A11+1)*12+MONTH(A11+1)=(RIGHT($K$5,4)*12+(FIND(LEFT($K$5,3),"JANFEBMARAPRMAYJUNJULAUGSEPOCTNOVDEC")+2)/3)),FALSE))
```

* It is 253 characters long, under Excel's 255-character limit for validation formulas.
* It works out the month from the first three letters of K5, so it doesn't depend on the
  computer's language settings. A `DATEVALUE("1 "&K5)` version was rejected for that reason.
* Checked in LibreOffice on 19 cases, with 0 wrong:
  * **Accepted:** 1 and 31 Oct plus 30 Sep with `OCTOBER 2026`; 31 Dec 2026 and 1 Jan 2027 with
    `JANUARY 2027`; 29 Feb 2028 with `MARCH 2028`; 28 Feb 2027 with `MARCH 2027`; 30 Apr with `MAY 2026`.
  * **Rejected:** 29 Sep, 1 Nov and 15 Oct 2025 with `OCTOBER 2026`; 30 Dec 2026 with `JANUARY 2027`;
    28 Feb 2028 with `MARCH 2028`.
  * **K5 empty:** every date is rejected. A blank date cell is always allowed.
* Because errors become a clean "rejected", Excel no longer shows "The formula currently
  evaluates to an error" when you save the rule.
* **K5 accepts month text only.** The current K5 rule rejects every date-style entry:
  `1/10/2026`, `01-Oct-2026`, `1 OCTOBER 2026`, `2026-10-01`, `Oct-26`, a pasted real date, and a
  raw date number. *(V04)*

The stress test accepts either design. V06 passes when the date rule reads K5 directly, and
only fails when the rule points at an empty J6.

## Applied to the form: K5 digits-only rule and date rule reading K5

Applied to `OVERTIME_FORM_MONTH_CAPITALIZED_ONLY.xlsx` (the copy with the compressed 3 MB picture).
Only these two validation formulas were changed; every other part of the file is byte for byte
the same.

**K5:** a capitalised full month name, one space, and four plain digits. Any year is allowed.

```
=IFERROR(AND(FIND(" ",K5)=LEN(K5)-4,ISNUMBER(FIND("|"&LEFT(K5,LEN(K5)-5)&"|","|JANUARY|FEBRUARY|MARCH|APRIL|MAY|JUNE|JULY|AUGUST|SEPTEMBER|OCTOBER|NOVEMBER|DECEMBER|")),TEXT(--RIGHT(K5,4),"0000")=RIGHT(K5,4)),FALSE)
```

* Checked on 37 entries in LibreOffice, with 0 wrong.
* **Accepted:** `OCTOBER 2026`, `JUNE 2026`, `OCTOBER 2019`, `OCTOBER 0000`, `OCTOBER 9999`.
* **Rejected:**
  * years that are not plain digits: `-202`, `2.26`, `1E03`, `+202`, `2O26`, `20 6`, ` 202`
  * lower-case and short forms, and extra spaces
  * every date-style entry
* 215 characters.

**A11:A70:** the date rule reads K5 directly, so `'Formula - Do Not Edit'!J6` stays empty. This is
the rule from the section "Keeping J6 empty" above.

Still open in that copy: the Engineer Name sheet is not protected (S03), and the K5 error
message doesn't mention the year rule.

## Not yet verified (needs Excel)

* **Live refresh.** Whether the `Web.BrowserContents` refresh succeeds reliably, and its
  timing. *(R01–R04)*
* **Rule results inside Excel.** The validation results come from `Range.Validation.Value`,
  which needs Excel. The LibreOffice results above were obtained by evaluating the same
  rule formulas in cells. *(V04–V08)*
* **Recalculation speed** with `forceFullCalc` switched on. *(Z04, C04)*
