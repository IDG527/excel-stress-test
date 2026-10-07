Attribute VB_Name = "OvertimeStressTest"
'==============================================================================
'  OVERTIME FORM STRESS TEST
'  Target: OVERTIME_FORM_MONTH_MANUAL_INPUT_CLEAN_1.xlsx (month entry in K5, input pop-ups)
'          Also runs against the earlier rev 1.2 QUERY_LINKED form.
'==============================================================================
'  HOW TO RUN
'    1. Open a new blank workbook, press Alt+F11, File > Import File... and
'       choose this .bas file. Save the workbook as StressTestRunner.xlsm.
'    2. Alt+F8 > RunOvertimeStressTest > Run, then pick the overtime form.
'
'  SAFETY
'    * Every test runs on a temporary COPY of the form. The original file is
'      never opened for writing and never saved.
'    * The Power Query "Holidays" is only read and refreshed - never edited.
'      Its M code and connection are fingerprinted before and after the run
'      and the run is marked FAIL if a single character changed.
'
'  OUTPUT (sheets added to this runner workbook)
'    ST_Summary    - verdict, pass/fail counts, performance, list of issues
'    ST_Results    - one line per check
'    ST_Timings    - every refresh / recalculation timing
'    ST_Mismatches - inputs where the form disagreed with the reference model
'    ST_Scenarios  - your own (or Copilot-generated) test rows, run every time
'    ST_Copilot    - =COPILOT() analysis of the results + ready-made prompts
'==============================================================================
Option Explicit

' ---------- settings you can change ----------
Private Const REFRESH_ITERATIONS As Long = 10     ' Power Query refreshes (each one reads mom.gov.sg)
Private Const REFRESH_PAUSE_MS As Long = 1000     ' pause between refreshes, to be polite to the website
Private Const FUZZ_ITERATIONS As Long = 200       ' random full-form fills, 60 rows each
Private Const RECALC_ITERATIONS As Long = 50      ' forced full recalculations of a full form
Private Const OFFICE_HOURS_FUZZ As Long = 10      ' random fills per alternative office-hours setting
Private Const ROLLOVER_FUZZ As Long = 10          ' random fills per simulated future holiday year
Private Const RANDOM_SEED As Long = 20260929      ' same seed = same random inputs every run
Private Const FORM_PASSWORD As String = ""        ' sheet password of 'Engineer Name', if it has one
Private Const KEEP_TEST_COPY As Boolean = False   ' True keeps the scratch copy for inspection
Private Const MAX_MISMATCH_LOG As Long = 500
Private Const SLOW_RECALC_MS As Double = 500      ' WARN when an average recalculation is slower
Private Const SLOW_REFRESH_MS As Double = 30000   ' WARN when an average refresh is slower
Private Const LARGE_FILE_MB As Double = 5
Private Const TOL As Double = 0.001

' ---------- workbook layout (rev 1.2) ----------
Private Const SH_FORM As String = "Engineer Name"
Private Const SH_ENGINE As String = "Formula - Do Not Edit"
Private Const SH_HELP As String = "HOW TO USE (FSE Instructions)"   ' rev 1.2; later versions use README sheets
Private Const SH_HOL As String = "SG Public Holidays"
Private Const QUERY_NAME As String = "Holidays"
Private Const CONN_NAME As String = "Query - Holidays"
Private Const TABLE_NAME As String = "Holidays_1"
Private Const FIRST_ROW As Long = 11
Private Const LAST_ROW As Long = 70
Private Const TOTAL_ROW As Long = 71
Private Const NROWS As Long = 60
Private Const DEFAULT_Q1 As Long = 480            ' 08:00 in minutes
Private Const DEFAULT_Q2 As Long = 1050           ' 17:30 in minutes
Private Const MONTH_CELL As String = "K5"         ' claim month typed by the engineer, e.g. OCTOBER 2026
Private Const ENGINE_MONTH_CELL As String = "J6"  ' engine cell the date rule reads the claim month from
Private Const DATE_ALLOWS_NEXT_FIRST As Boolean = False ' rev2 allowed the 1st of the next month; rev4 does not

' ---------- output sheets ----------
Private Const OUT_SUMMARY As String = "ST_Summary"
Private Const OUT_RESULTS As String = "ST_Results"
Private Const OUT_TIMINGS As String = "ST_Timings"
Private Const OUT_MISMATCH As String = "ST_Mismatches"
Private Const OUT_SCENARIOS As String = "ST_Scenarios"
Private Const OUT_COPILOT As String = "ST_Copilot"
Private Const SCEN_FIRST_ROW As Long = 5

#If Mac Then
#ElseIf VBA7 Then
Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (ByRef lpCount As Currency) As Long
Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (ByRef lpFrequency As Currency) As Long
#Else
Private Declare Function QueryPerformanceCounter Lib "kernel32" (ByRef lpCount As Currency) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (ByRef lpFrequency As Currency) As Long
#End If

' Result of the reference model for one form row
Private Type OTResult
    DayText As String
    PH As String
    Travel As Variant       ' hours (Double) or "" when the form shows blank
    Work As Variant
End Type

Private mWb As Workbook                 ' the scratch copy under test
Private mForm As Worksheet
Private mEng As Worksheet
Private mRes As Worksheet
Private mTim As Worksheet
Private mMis As Worksheet
Private mResRow As Long
Private mTimRow As Long
Private mMisRow As Long
Private mPass As Long, mFail As Long, mWarn As Long, mSkip As Long, mInfo As Long
Private mHol() As Long                  ' holiday date serials from the query table
Private mHolCount As Long
Private mHolYear As Long
Private mQ1 As Long                     ' office hours on the form, minutes after midnight
Private mQ2 As Long
Private mOrigLen As Double
Private mSourcePath As String
Private mQueryFingerprint As String
Private mEngFormulaH As String          ' engine H11 / I11 formulas, captured for Copilot
Private mEngFormulaI As String
Private mPopups() As String               ' section | cell | pop-up title | pop-up text | error title | error text
Private mPopupCount As Long
Private mColLocal As Long                 ' LOCAL / OVERSEAS column, 0 when the form has none (rev2)
Private mColProject As Long               ' Project ID column
Private mColVessel As Long                ' Vessel name column
Private mColActivity As Long              ' Activity number column, 0 when the form has none (added in rev4)
Private mLastInput As Long                ' last input column of the job rows

'==============================================================================
'  ENTRY POINTS
'==============================================================================
Public Sub RunOvertimeStressTest()
    Dim srcPath As Variant, copyPath As String, origStamp As Date
    Dim mBefore As String, connBefore As String, mAfter As String, connAfter As String
    Dim foundBefore As Boolean, foundAfter As Boolean
    Dim oldCalc As Long, oldScreen As Boolean, oldAlerts As Boolean, oldEvents As Boolean
    Dim t0 As Double, baseline As Variant, finishing As Boolean, seconds As Double

    srcPath = Application.GetOpenFilename("Excel workbooks (*.xlsx;*.xlsm),*.xlsx;*.xlsm", , _
                                          "Choose the OVERTIME FORM workbook to stress test")
    If VarType(srcPath) = vbBoolean Then Exit Sub
    If StrComp(CStr(srcPath), ThisWorkbook.FullName, vbTextCompare) = 0 Then
        MsgBox "Pick the overtime form, not this runner workbook.", vbExclamation
        Exit Sub
    End If
    mSourcePath = CStr(srcPath)

    oldCalc = Application.Calculation
    oldScreen = Application.ScreenUpdating
    oldAlerts = Application.DisplayAlerts
    oldEvents = Application.EnableEvents
    On Error GoTo Fatal

    t0 = NowMs()
    PrepareOutputSheets
    mOrigLen = FileLen(mSourcePath)
    origStamp = FileDateTime(mSourcePath)

    Progress "copying the workbook"
    copyPath = MakeWorkingCopy(mSourcePath)
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual      ' keep the values saved in the file until F07 has looked at them
    Set mWb = Workbooks.Open(Filename:=copyPath, UpdateLinks:=0, ReadOnly:=False, AddToMru:=False)

    If Not BindSheets() Then GoTo Finish
    mEngFormulaH = mEng.Range("H11").Formula
    mEngFormulaI = mEng.Range("I11").Formula
    mQ1 = TimeCellMinutes(mForm.Range("Q1").Value2, DEFAULT_Q1)
    mQ2 = TimeCellMinutes(mForm.Range("Q2").Value2, DEFAULT_Q2)
    baseline = mForm.Range(mForm.Cells(FIRST_ROW, 1), mForm.Cells(LAST_ROW, 12)).Value2
    mBefore = ReadQueryFormula(foundBefore)
    connBefore = ReadConnectionText()
    mQueryFingerprint = Fingerprint(mBefore)
    LoadHolidays

    Progress "structure checks":       TestStructure
    Progress "formula integrity":      TestFormulas
    Progress "query definition":       TestQueryDefinition mBefore, foundBefore
    Progress "as-found entries":       TestBaseline baseline
    Progress "Power Query refresh":    TestRefreshStress baseline
    Progress "holiday table content":  TestHolidayOutput
    Progress "pop-ups + month entry":  TestPopups
    Progress "known answers":          TestKnownAnswers
    Progress "known limitations":      TestLimitations
    Progress "bad pasted input":       TestRobustness
    Progress "random fills":           SeedRandom: TestFuzz FUZZ_ITERATIONS, "Z", "Random fills"
    Progress "capacity + performance": TestCapacity
    Progress "office hours variations": TestOfficeHours
    Progress "year rollover":          TestYearRollover
    Progress "your scenarios":         TestScenarioSheet

Finish:
    finishing = True
    On Error Resume Next
    If Not mWb Is Nothing Then
        ' ---- the query must be byte-for-byte what it was before the run ----
        mAfter = ReadQueryFormula(foundAfter)
        connAfter = ReadConnectionText()
        If Not foundBefore Then
            LogResult "Q99", "Query safety", "Power Query M code unchanged by the stress test", "SKIP", , , _
                      "Query '" & QUERY_NAME & "' could not be read (Workbook.Queries needs Excel 2016 or later)."
        ElseIf StrComp(mBefore, mAfter, vbBinaryCompare) = 0 And StrComp(connBefore, connAfter, vbBinaryCompare) = 0 Then
            LogResult "Q99", "Query safety", "Power Query M code and connection unchanged by the stress test", "PASS", _
                      Fingerprint(mBefore), Fingerprint(mAfter)
        Else
            LogResult "Q99", "Query safety", "Power Query M code and connection unchanged by the stress test", "FAIL", _
                      Fingerprint(mBefore), Fingerprint(mAfter), "The query text or connection differs after the run."
        End If
        mWb.Close SaveChanges:=False
        Set mWb = Nothing
    End If
    If Len(copyPath) > 0 And Not KEEP_TEST_COPY Then Kill copyPath
    If Len(mSourcePath) > 0 Then
        If FileLen(mSourcePath) = mOrigLen And FileDateTime(mSourcePath) = origStamp Then
            LogResult "Q98", "Query safety", "Original workbook left untouched", "PASS", _
                      Format$(mOrigLen, "#,##0") & " bytes", Format$(FileLen(mSourcePath), "#,##0") & " bytes"
        Else
            LogResult "Q98", "Query safety", "Original workbook left untouched", "FAIL", _
                      Format$(origStamp, "yyyy-mm-dd hh:nn:ss"), Format$(FileDateTime(mSourcePath), "yyyy-mm-dd hh:nn:ss"), _
                      "The original file's size or timestamp changed during the run."
        End If
    End If
    Application.Calculation = oldCalc
    seconds = (NowMs() - t0) / 1000#
    WriteSummary seconds
    WriteCopilotSheet
    Application.ScreenUpdating = oldScreen
    Application.DisplayAlerts = oldAlerts
    Application.EnableEvents = oldEvents
    Application.StatusBar = False
    ThisWorkbook.Worksheets(OUT_SUMMARY).Activate
    MsgBox "Stress test finished." & vbCrLf & vbCrLf & _
           mPass & " PASS, " & mFail & " FAIL, " & mWarn & " WARN, " & mSkip & " SKIP, " & mInfo & " INFO" & vbCrLf & _
           "See the " & OUT_SUMMARY & " sheet.", IIf(mFail > 0, vbExclamation, vbInformation)
    Exit Sub

Fatal:
    If finishing Then Resume Next
    If Not mRes Is Nothing Then
        LogResult "X00", "Runner", "Stress test stopped early", "FAIL", , , "Error " & Err.Number & ": " & Err.Description
    Else
        MsgBox "Stress test could not start: " & Err.Description, vbCritical
    End If
    Resume Finish
End Sub

' Runs only the rows on ST_Scenarios against the form (no refresh, no fuzzing).
Public Sub RunScenariosOnly()
    Dim srcPath As Variant, copyPath As String, oldCalc As Long

    srcPath = Application.GetOpenFilename("Excel workbooks (*.xlsx;*.xlsm),*.xlsx;*.xlsm", , _
                                          "Choose the OVERTIME FORM workbook")
    If VarType(srcPath) = vbBoolean Then Exit Sub
    mSourcePath = CStr(srcPath)
    oldCalc = Application.Calculation
    On Error GoTo Done
    Set mRes = Nothing                       ' results of the last full run stay as they are
    Set mMis = Nothing
    mPass = 0: mFail = 0: mWarn = 0: mSkip = 0: mInfo = 0
    EnsureScenarioSheet
    copyPath = MakeWorkingCopy(mSourcePath)
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Set mWb = Workbooks.Open(Filename:=copyPath, UpdateLinks:=0, ReadOnly:=False, AddToMru:=False)
    If BindSheets() Then
        mPass = 0: mFail = 0: mWarn = 0: mSkip = 0: mInfo = 0
        mQ1 = TimeCellMinutes(mForm.Range("Q1").Value2, DEFAULT_Q1)
        mQ2 = TimeCellMinutes(mForm.Range("Q2").Value2, DEFAULT_Q2)
        LoadHolidays
        TestScenarioSheet
    End If
Done:
    On Error Resume Next
    If Not mWb Is Nothing Then mWb.Close SaveChanges:=False
    Set mWb = Nothing
    If Len(copyPath) > 0 And Not KEEP_TEST_COPY Then Kill copyPath
    Application.Calculation = oldCalc
    Application.ScreenUpdating = True
    Application.StatusBar = False
    ThisWorkbook.Worksheets(OUT_SCENARIOS).Activate
    MsgBox "Scenarios finished: " & mPass & " PASS, " & mFail & " FAIL, " & mWarn & " WARN.", vbInformation
End Sub

'==============================================================================
'  S - STRUCTURE
'==============================================================================
Private Sub TestStructure()
    Dim cat As String, v As Variant, i As Long, bad As String, lockedState As Variant
    Dim headers As Variant, f As String, sz As Double, rng As Range
    cat = "Structure"
    On Error GoTo Boom

    If mEng.Visible = xlSheetVisible Then
        LogResult "S02", cat, "Formula engine sheet is hidden from users", "WARN", "hidden", "visible", _
                  "Users can see and edit '" & SH_ENGINE & "'."
    Else
        LogResult "S02", cat, "Formula engine sheet is hidden from users", "PASS", "hidden", "hidden"
    End If

    If mForm.ProtectContents Then
        LogResult "S03", cat, "Form sheet is protected", "PASS", "protected", "protected"
    Else
        LogResult "S03", cat, "Form sheet is protected", "WARN", "protected", "not protected", _
                  "Users can type over the Day / Public holiday / Travel / Work formulas and the totals. " & _
                  "Protect the sheet again (the input cells are already unlocked)."
    End If

    bad = ""
    For Each v In Array("A11:A70", "D11:G70", ColRange(10, mLastInput), "B5", "I5", MONTH_CELL, "C77")
        lockedState = mForm.Range(v).Locked
        If IsNull(lockedState) Then
            bad = bad & " " & v & "(mixed)"
        ElseIf lockedState Then
            bad = bad & " " & v
        End If
    Next v
    If Len(bad) = 0 Then
        LogResult "S04", cat, "All input cells are unlocked (users can type in them)", "PASS"
    Else
        LogResult "S04", cat, "All input cells are unlocked (users can type in them)", "FAIL", "unlocked", "locked:" & bad
    End If
    lockedState = mForm.Range(ColRange(mLastInput + 1, mLastInput + 1)).Locked
    If Not IsNull(lockedState) Then
        If Not lockedState And Len(Trim$(CStr(mForm.Cells(9, mLastInput + 1).Value))) = 0 Then
            LogResult "S04b", cat, "No unlocked cells without a heading next to the job rows", "WARN", "locked", _
                      ColRange(mLastInput + 1, mLastInput + 1) & " unlocked, no heading", _
                      "Left over from the removed column: engineers can type there, outside the print area, and nothing uses it."
        End If
    End If

    bad = ""
    For Each v In Array("B11:C70", "H11:I70", "H71:I71", "K71")
        lockedState = mForm.Range(v).Locked
        If IsNull(lockedState) Then
            bad = bad & " " & v & "(mixed)"
        ElseIf Not lockedState Then
            bad = bad & " " & v
        End If
    Next v
    If Len(bad) = 0 Then
        LogResult "S05", cat, "All calculated cells are locked", "PASS"
    Else
        LogResult "S05", cat, "All calculated cells are locked", "WARN", "locked", "unlocked:" & bad
    End If

    If IsNumeric(mForm.Range("Q1").Value2) And IsNumeric(mForm.Range("Q2").Value2) And mQ1 < mQ2 Then
        LogResult "S06", cat, "Office hours (Q1 to Q2) are valid times, start before end", "PASS", _
                  HHMM(DEFAULT_Q1) & "-" & HHMM(DEFAULT_Q2), HHMM(mQ1) & "-" & HHMM(mQ2)
    Else
        LogResult "S06", cat, "Office hours (Q1 to Q2) are valid times, start before end", "FAIL", _
                  "two times, Q1 < Q2", ToText(mForm.Range("Q1").Value) & " / " & ToText(mForm.Range("Q2").Value)
    End If

    headers = Array("A2", "OVERTIME CLAIM", "A5", "NAME", "H5", "ID", "J5", "MONTH", "A9", "DATE", "B9", "DAY", _
                    "C9", "PUBLIC HOLIDAY (Y / N)", "D9", "TRAVEL - WORK - TRAVEL", "H9", "OVERTIME RATE", _
                    "D10", "From", _
                    "E10", "Until / From", "F10", "Until / From", "G10", "Until", "H10", "Travel", "I10", "Work", _
                    "J71", "TOTAL", "A77", "Submitted By:")
    bad = ""
    For i = LBound(headers) To UBound(headers) Step 2
        If StrComp(Squash(CStr(mForm.Range(headers(i)).Value)), Squash(CStr(headers(i + 1))), vbTextCompare) <> 0 Then
            bad = bad & " " & headers(i) & "='" & CStr(mForm.Range(headers(i)).Value) & "'"
        End If
    Next i
    If mColProject = 0 Then bad = bad & " Project ID heading not found in J9:L9"
    If mColVessel = 0 Then bad = bad & " VESSEL NAME heading not found in J9:M9"
    If Len(bad) = 0 Then
        LogResult "S07", cat, "Form headings are where the formulas expect them", "PASS", , _
                  "Project ID in " & ColLetter(mColProject) & ", Vessel in " & ColLetter(mColVessel) & _
                  IIf(mColActivity > 0, ", Activity number in " & ColLetter(mColActivity), "") & _
                  IIf(mColLocal > 0, ", LOCAL / OVERSEAS in " & ColLetter(mColLocal), ", no LOCAL / OVERSEAS column")
    Else
        LogResult "S07", cat, "Form headings are where the formulas expect them", "WARN", , bad, _
                  "A heading moved or changed; check the layout was not shifted."
    End If

    bad = ""
    If Len(Trim$(CStr(mForm.Range("B5").Value))) = 0 Then bad = bad & " NAME (B5)"
    If Len(Trim$(CStr(mForm.Range("I5").Value))) = 0 Then bad = bad & " ID (I5)"
    If Len(Trim$(CStr(mForm.Range(MONTH_CELL).Value))) = 0 Then
        bad = bad & " MONTH (" & MONTH_CELL & ")"
    ElseIf ParseMonthText(CStr(mForm.Range(MONTH_CELL).Value)) = 0 Then
        bad = bad & " MONTH (" & MONTH_CELL & ") not in the form 'OCTOBER 2026': '" & mForm.Range(MONTH_CELL).Value & "'"
    End If
    If Len(bad) = 0 Then
        LogResult "S08", cat, "Claim header filled in (name, ID, month)", "PASS"
    Else
        LogResult "S08", cat, "Claim header filled in (name, ID, month)", "WARN", "filled", "blank:" & bad, _
                  "The form does not force these fields; a claim can be submitted without them."
    End If

    LogResult "S09", cat, "Verifier / approver names present", _
              IIf(Len(Trim$(CStr(mForm.Range("K77").Value))) > 0 And Len(Trim$(CStr(mForm.Range("K83").Value))) > 0, "PASS", "WARN"), _
              "K77 and K83 filled", "'" & Trim$(CStr(mForm.Range("K77").Value)) & "' / '" & Trim$(CStr(mForm.Range("K83").Value)) & "'"

    If mColLocal = 0 Then
        LogResult "S10", cat, "LOCAL / OVERSEAS column", "INFO", , "removed", _
                  "The hidden list in M9:M10 (" & ToText(mForm.Range("M9").Value) & ", " & ToText(mForm.Range("M10").Value) & ") is no longer used."
    Else
        f = ValidationFormula(mForm.Cells(FIRST_ROW, mColLocal))
        If UCase$(mForm.Range("M9").Value) = "LOCAL" And UCase$(mForm.Range("M10").Value) = "OVERSEAS" And _
           InStr(1, Replace(f, "$", ""), "M9:M10", vbTextCompare) > 0 Then
            LogResult "S10", cat, "LOCAL / OVERSEAS drop-down points at M9:M10", "PASS", "LOCAL, OVERSEAS", _
                      mForm.Range("M9").Value & ", " & mForm.Range("M10").Value
        Else
            LogResult "S10", cat, "LOCAL / OVERSEAS drop-down points at M9:M10", "WARN", "=$M$9:$M$10 -> LOCAL, OVERSEAS", _
                      IIf(Len(f) = 0, "(no drop-down)", f) & " -> " & ToText(mForm.Range("M9").Value) & ", " & ToText(mForm.Range("M10").Value)
        End If
    End If

    bad = ""
    If Not (ValidationType(mForm.Range("A11")) = xlValidateDate Or ValidationType(mForm.Range("A11")) = xlValidateCustom) Or _
       Not (ValidationType(mForm.Range("A70")) = xlValidateDate Or ValidationType(mForm.Range("A70")) = xlValidateCustom) Then bad = bad & " A11:A70(date)"
    For Each v In Array("D", "E", "F", "G")
        If ValidationType(mForm.Range(v & "11")) <> xlValidateCustom Or ValidationType(mForm.Range(v & "70")) <> xlValidateCustom Then bad = bad & " " & v & "11:" & v & "70(time)"
    Next v
    If mColLocal > 0 Then
        If ValidationType(mForm.Cells(FIRST_ROW, mColLocal)) <> xlValidateList Or ValidationType(mForm.Cells(LAST_ROW, mColLocal)) <> xlValidateList Then bad = bad & " " & ColRange(mColLocal, mColLocal) & "(list)"
    End If
    If Len(bad) = 0 Then
        LogResult "S11", cat, "Data validation on date, time and LOCAL/OVERSEAS columns (rows 11-70)", "PASS"
    Else
        LogResult "S11", cat, "Data validation on date, time and LOCAL/OVERSEAS columns (rows 11-70)", "WARN", , "missing:" & bad
    End If
    LogResult "S11b", cat, "Data validation is bypassed by paste / VBA", "INFO", , , _
              "Excel validation only checks typed entries. Pasted values are accepted - see the P tests for what that does."

    v = SheetByName(mWb, SH_HOL).Visible
    LogResult "S19", cat, "Holiday list sheet visibility", "INFO", , IIf(v = xlSheetVisible, "visible", "hidden"), _
              IIf(v = xlSheetVisible, "Engineers can see the holiday list.", _
                  "'" & SH_HOL & "' is hidden; engineers see only the Y/N column. The lookup still works (H15).")
    On Error Resume Next
    v = Empty
    v = mWb.ForceFullCalculation
    On Error GoTo Boom
    If VarType(v) = vbBoolean Then
        LogResult "S20", cat, "Workbook forces a full recalculation on every change", "INFO", , CStr(v), _
                  IIf(v, "Every edit recalculates all 600+ formulas; C04 shows what that costs.", "")
    End If

    If mForm.Range("D11").FormatConditions.Count > 0 And mForm.Range("G70").FormatConditions.Count > 0 Then
        LogResult "S12", cat, "Out-of-range time highlight (conditional format) on D11:G70", "PASS"
    Else
        LogResult "S12", cat, "Out-of-range time highlight (conditional format) on D11:G70", "WARN", "present", "missing"
    End If

    bad = ""
    If NzStr(mForm.Range("D11:G70").NumberFormat) <> "hh:mm" Then bad = bad & " D:G=" & NzStr(mForm.Range("D11:G70").NumberFormat)
    If mColProject > 0 Then
        If NzStr(mForm.Range(ColRange(mColProject, mColProject)).NumberFormat) <> "@" Then bad = bad & " Project ID=" & NzStr(mForm.Range(ColRange(mColProject, mColProject)).NumberFormat)
    End If
    If mColActivity > 0 Then
        If NzStr(mForm.Range(ColRange(mColActivity, mColActivity)).NumberFormat) <> "@" Then bad = bad & " Activity number=" & NzStr(mForm.Range(ColRange(mColActivity, mColActivity)).NumberFormat)
    End If
    If Len(NzStr(mForm.Range("A11:A70").NumberFormat)) = 0 Then bad = bad & " A=(mixed)"
    If Len(bad) = 0 Then
        LogResult "S13", cat, "Number formats: times hh:mm, Project ID as text", "PASS"
    Else
        LogResult "S13", cat, "Number formats: times hh:mm, Project ID as text", "WARN", "hh:mm / @", bad, _
                  "An ID column that is not Text changes what is typed: 100084981.010 becomes 100084981.01, " & _
                  "0010 becomes 10, and a 16+ digit number loses its last digits."
    End If

    If NzStr(mForm.Cells(TOTAL_ROW, 8).Formula) = "=SUM(H11:H70)" And _
       NzStr(mForm.Cells(TOTAL_ROW, 9).Formula) = "=SUM(I11:I70)" And NzStr(mForm.Cells(TOTAL_ROW, 11).Formula) = "=H71+I71" Then
        LogResult "S14", cat, "Totals: H71=SUM(H11:H70), I71=SUM(I11:I70), K71=H71+I71", "PASS"
    Else
        LogResult "S14", cat, "Totals: H71=SUM(H11:H70), I71=SUM(I11:I70), K71=H71+I71", "FAIL", _
                  "=SUM(H11:H70) / =SUM(I11:I70) / =H71+I71", _
                  mForm.Range("H71").Formula & " / " & mForm.Range("I71").Formula & " / " & mForm.Range("K71").Formula
    End If

    If mForm.Range("H71").NumberFormat <> mForm.Range("I71").NumberFormat Then
        LogResult "S15", cat, "Travel and Work totals use the same number format", "WARN", _
                  mForm.Range("I71").NumberFormat, mForm.Range("H71").NumberFormat, _
                  "H71 and I71 display differently (e.g. 15.5 vs 7.5 / 15.50 vs 7.5)."
    Else
        LogResult "S15", cat, "Travel and Work totals use the same number format", "PASS"
    End If

    Set rng = Nothing
    On Error Resume Next
    Set rng = mEng.Cells.SpecialCells(xlCellTypeAllValidation)
    On Error GoTo Boom
    If rng Is Nothing Then
        LogResult "S16", cat, "No stray data validation on the hidden engine sheet", "PASS"
    Else
        LogResult "S16", cat, "No stray data validation on the hidden engine sheet", "WARN", "none", rng.Address(False, False), _
                  "Leftover drop-downs on '" & SH_ENGINE & "' point at empty ranges. Harmless, but clutter."
    End If

    f = ""
    On Error Resume Next
    f = mForm.PageSetup.PrintArea
    On Error GoTo Boom
    LogResult "S17", cat, "Print area covers the whole form", IIf(Replace(f, "$", "") = "A1:" & ColLetter(mLastInput) & "84", "PASS", "WARN"), _
              "A1:" & ColLetter(mLastInput) & "84", f

    sz = mOrigLen / 1048576#
    If sz > LARGE_FILE_MB Then
        LogResult "S18", cat, "Workbook file size", "WARN", "< " & LARGE_FILE_MB & " MB", Format$(sz, "0.0") & " MB", _
                  "The instruction sheets hold " & HelpShapeCount() & " picture(s); a full-resolution screenshot there is " & _
                  "the usual cause. Large files open, e-mail and sync slowly. Compress the picture (Picture Format > Compress)."
    Else
        LogResult "S18", cat, "Workbook file size", "PASS", "< " & LARGE_FILE_MB & " MB", Format$(sz, "0.0") & " MB"
    End If
    Exit Sub
Boom:
    LogCrash "S99", cat
End Sub

'==============================================================================
'  F - FORMULA INTEGRITY
'==============================================================================
Private Sub TestFormulas()
    Dim cat As String, c As Long, r As Long, base As String, bad As String, nBad As Long
    Dim cell As Range, rng As Range, sh As Variant, errs As String, nErr As Long, v As Variant
    cat = "Formulas"
    On Error GoTo Boom

    ' F01 every engine row uses the same formula as row 11
    For c = 1 To 9
        base = mEng.Cells(FIRST_ROW, c).FormulaR1C1
        For r = FIRST_ROW + 1 To LAST_ROW
            If mEng.Cells(r, c).FormulaR1C1 <> base Then
                nBad = nBad + 1
                If nBad <= 10 Then bad = bad & " " & mEng.Cells(r, c).Address(False, False)
            End If
        Next r
    Next c
    If nBad = 0 Then
        LogResult "F01", cat, "Engine formulas identical on all 60 rows (A:I, rows 11-70)", "PASS", "540 consistent", "540 consistent"
    Else
        LogResult "F01", cat, "Engine formulas identical on all 60 rows (A:I, rows 11-70)", "FAIL", "0 different", nBad & " different", "First: " & bad
    End If

    ' F02 every form output cell points at the same cell on the engine sheet
    nBad = 0: bad = ""
    For Each v In Array(2, 3, 8, 9)
        For r = FIRST_ROW To LAST_ROW
            If mForm.Cells(r, v).FormulaR1C1 <> "='" & SH_ENGINE & "'!RC" Then
                nBad = nBad + 1
                If nBad <= 10 Then bad = bad & " " & mForm.Cells(r, v).Address(False, False)
            End If
        Next r
    Next v
    If nBad = 0 Then
        LogResult "F02", cat, "Form columns B, C, H, I link to the same row on the engine sheet", "PASS", "240 links", "240 links"
    Else
        LogResult "F02", cat, "Form columns B, C, H, I link to the same row on the engine sheet", "FAIL", "0 broken", nBad & " broken", "First: " & bad
    End If

    ' F03 no #REF! written into any formula
    nBad = 0: bad = ""
    For Each sh In Array(mForm, mEng)
        Set rng = Nothing
        On Error Resume Next
        Set rng = sh.UsedRange.SpecialCells(xlCellTypeFormulas)
        On Error GoTo Boom
        If Not rng Is Nothing Then
            For Each cell In rng.Cells
                If InStr(1, cell.Formula, "#REF!") > 0 Then
                    nBad = nBad + 1
                    If nBad <= 10 Then bad = bad & " " & sh.Name & "!" & cell.Address(False, False)
                End If
            Next cell
        End If
    Next sh
    LogResult "F03", cat, "No #REF! inside any formula", IIf(nBad = 0, "PASS", "FAIL"), "0", CStr(nBad), bad

    ' F04 office hours and holiday table are wired in
    If Replace(mEng.Range("C6").Formula, "$", "") = "='" & SH_FORM & "'!Q1" And _
       Replace(mEng.Range("E6").Formula, "$", "") = "='" & SH_FORM & "'!Q2" Then
        LogResult "F04", cat, "Engine reads office hours from 'Engineer Name'!Q1:Q2", "PASS"
    Else
        LogResult "F04", cat, "Engine reads office hours from 'Engineer Name'!Q1:Q2", "WARN", , _
                  mEng.Range("C6").Formula & " / " & mEng.Range("E6").Formula
    End If
    If InStr(1, mEng.Range("C11").Formula, TABLE_NAME & "[Date]", vbTextCompare) > 0 Then
        LogResult "F05", cat, "Public-holiday flag looks up " & TABLE_NAME & "[Date] (the query output)", "PASS"
    Else
        LogResult "F05", cat, "Public-holiday flag looks up " & TABLE_NAME & "[Date] (the query output)", "FAIL", _
                  TABLE_NAME & "[Date]", mEng.Range("C11").Formula
    End If

    ' F07 values exactly as saved in the file (nothing recalculated yet)
    For Each v In Array("B11:C70", "H11:I70", "H71:I71", "K71")
        For Each cell In mForm.Range(v).Cells
            If IsError(cell.Value) Then
                nErr = nErr + 1
                If nErr <= 10 Then errs = errs & " " & cell.Address(False, False) & "=" & ErrText(cell.Value)
            End If
        Next cell
    Next v
    If nErr = 0 Then
        LogResult "F07", cat, "Values saved in the file contain no error cells", "PASS", "0 error cells", "0 error cells"
    Else
        LogResult "F07", cat, "Values saved in the file contain no error cells", "WARN", "0 error cells", nErr & " error cells", _
                  "First:" & errs & ". Anything that shows the file without recalculating it (previews, some mobile/web " & _
                  "viewers, PDF exports) shows these errors. Open it in Excel, press Ctrl+Alt+F9 and save again."
    End If
    nErr = 0: errs = ""

    ' F06 full recalculation of the file - no error values
    ForceRecalc
    For Each v In Array("B11:C70", "H11:I70", "H71:I71", "K71")
        For Each cell In mForm.Range(v).Cells
            If IsError(cell.Value) Then
                nErr = nErr + 1
                If nErr <= 10 Then errs = errs & " " & cell.Address(False, False) & "=" & ErrText(cell.Value)
            End If
        Next cell
    Next v
    LogResult "F06", cat, "After a full recalculation the form shows no #REF!/#VALUE!/#N/A", IIf(nErr = 0, "PASS", "FAIL"), _
              "0 error cells", nErr & " error cells", errs
    Exit Sub
Boom:
    LogCrash "F99", cat
End Sub

'==============================================================================
'  Q - POWER QUERY DEFINITION (read only - nothing here writes to the query)
'==============================================================================
Private Sub TestQueryDefinition(ByVal mCode As String, ByVal found As Boolean)
    Dim cat As String, cn As WorkbookConnection, s As String, lo As ListObject
    cat = "Power Query"
    On Error GoTo Boom

    If found Then
        LogResult "Q01", cat, "Query '" & QUERY_NAME & "' exists", "PASS", , Len(mCode) & " characters of M code"
    Else
        LogResult "Q01", cat, "Query '" & QUERY_NAME & "' exists", "FAIL", QUERY_NAME, "not found / not readable"
    End If

    Set cn = Nothing
    On Error Resume Next
    Set cn = mWb.Connections(CONN_NAME)
    On Error GoTo Boom
    If cn Is Nothing Then
        LogResult "Q02", cat, "Connection '" & CONN_NAME & "' exists and targets the query", "FAIL", CONN_NAME, "missing"
    Else
        s = ReadConnectionText()
        If InStr(1, s, "Location=" & QUERY_NAME, vbTextCompare) > 0 And InStr(1, s, "[" & QUERY_NAME & "]", vbTextCompare) > 0 Then
            LogResult "Q02", cat, "Connection '" & CONN_NAME & "' exists and targets the query", "PASS", , s
        Else
            LogResult "Q02", cat, "Connection '" & CONN_NAME & "' exists and targets the query", "FAIL", _
                      "Location=" & QUERY_NAME & " / SELECT * FROM [" & QUERY_NAME & "]", s
        End If
    End If

    Set lo = GetHolidayTable()
    s = ""
    If Not lo Is Nothing Then
        On Error Resume Next
        s = lo.QueryTable.WorkbookConnection.Name
        On Error GoTo Boom
    End If
    If lo Is Nothing Then
        LogResult "Q03", cat, "Table " & TABLE_NAME & " on '" & SH_HOL & "' is loaded by the query", "FAIL", TABLE_NAME, "table missing"
    ElseIf StrComp(s, CONN_NAME, vbTextCompare) = 0 Then
        LogResult "Q03", cat, "Table " & TABLE_NAME & " on '" & SH_HOL & "' is loaded by the query", "PASS", CONN_NAME, s
    Else
        LogResult "Q03", cat, "Table " & TABLE_NAME & " on '" & SH_HOL & "' is loaded by the query", "FAIL", CONN_NAME, _
                  IIf(Len(s) = 0, "(not linked to a connection - static copy?)", s)
    End If

    If found Then
        LogResult "Q04", cat, "M code fingerprint recorded (compared again at the end, Q99)", "INFO", , mQueryFingerprint, _
                  "Source: " & IIf(InStr(1, mCode, "mom.gov.sg", vbTextCompare) > 0, "Ministry of Manpower public-holiday page", "(not mom.gov.sg)")
        If InStr(1, mCode, "DateTime.LocalNow", vbTextCompare) > 0 Then
            LogResult "Q05", cat, "Holiday list covers only the current calendar year", "WARN", , _
                      "Year = Date.Year(DateTime.LocalNow())", _
                      "A claim for last year's dates (e.g. a December claim submitted in January) gets no public-holiday rate. See L04/L05."
        End If
        If InStr(1, mCode, "Web.BrowserContents", vbTextCompare) > 0 Then
            LogResult "Q06", cat, "Query scrapes a live web page", "INFO", , "Web.BrowserContents + Html.Table", _
                      "Needs internet access. The query stops with its own error if MOM changes the page layout or has not published the year yet."
        End If
    End If
    Exit Sub
Boom:
    LogCrash "Q97", cat
End Sub

'==============================================================================
'  B - THE ENTRIES AS FOUND IN THE FILE
'==============================================================================
Private Sub TestBaseline(ByRef base As Variant)
    Dim cat As String, i As Long, c As Long, t(0 To 3) As Long, ds As Double, e As OTResult
    Dim v As Variant, used As Long, bad As Long, badList As String, notes As String
    Dim months As String, yrs As String, unsorted As Boolean, lastD As Double, dupes As String
    Dim claimMonth As Double, outside As String
    Dim noTimes As String, noDate As String, badJ As String, numK As String, anyTime As Boolean, skip As Boolean
    Dim spanS(1 To NROWS) As Double, spanE(1 To NROWS) As Double, hasSpan(1 To NROWS) As Boolean, overlaps As String
    cat = "As-found entries"
    On Error GoTo Boom

    v = mForm.Range(mForm.Cells(FIRST_ROW, 1), mForm.Cells(LAST_ROW, 12)).Value
    claimMonth = ParseMonthText(CStr(mForm.Range(MONTH_CELL).Value))
    For i = 1 To NROWS
        skip = False
        anyTime = False
        ds = 0
        If Not IsEmpty(base(i, 1)) Then
            If VarType(base(i, 1)) = vbDouble Then
                ds = base(i, 1)
            Else
                notes = notes & " A" & (i + 10) & " is text;"
                skip = True
            End If
        End If
        For c = 0 To 3
            t(c) = -1
            If Not IsEmpty(base(i, 4 + c)) Then
                anyTime = True
                If VarType(base(i, 4 + c)) = vbDouble Then
                    If base(i, 4 + c) >= 0 And base(i, 4 + c) < 1 Then
                        t(c) = Int(base(i, 4 + c) * 1440 + 0.5) Mod 1440
                    Else
                        notes = notes & " " & Chr$(68 + c) & (i + 10) & " outside 00:00-23:59;"
                        skip = True
                    End If
                Else
                    notes = notes & " " & Chr$(68 + c) & (i + 10) & " is text;"
                    skip = True
                End If
            End If
        Next c
        If ds <> 0 Or anyTime Then used = used + 1
        If ds <> 0 And Not anyTime Then noTimes = noTimes & " " & (i + 10)
        If ds = 0 And anyTime And Not skip Then noDate = noDate & " " & (i + 10)
        If ds <> 0 Then
            If InStr(months, Format$(CDate(ds), "yyyy-mm")) = 0 Then months = months & " " & Format$(CDate(ds), "yyyy-mm")
            If claimMonth > 0 Then
                If Not InClaimMonth(ds, claimMonth) Then outside = outside & " A" & (i + 10) & "=" & Format$(CDate(ds), "d-mmm-yy")
            End If
            If Year(CDate(ds)) <> mHolYear And InStr(yrs, CStr(Year(CDate(ds)))) = 0 Then yrs = yrs & " " & Year(CDate(ds))
            If ds < lastD Then unsorted = True
            If ds = lastD Then dupes = dupes & " " & Format$(CDate(ds), "d-mmm")
            lastD = ds
        End If
        If mColLocal > 0 Then
            If Not IsEmpty(base(i, mColLocal)) Then
                If UCase$(CStr(base(i, mColLocal))) <> "LOCAL" And UCase$(CStr(base(i, mColLocal))) <> "OVERSEAS" Then badJ = badJ & " " & ColLetter(mColLocal) & (i + 10)
            End If
        End If
        If mColProject > 0 Then
            If Not IsEmpty(base(i, mColProject)) And VarType(base(i, mColProject)) <> vbString Then numK = numK & " " & ColLetter(mColProject) & (i + 10)
        End If

        If Not skip And ds <> 0 Then hasSpan(i) = RowSpan(ds, t, spanS(i), spanE(i))
        If Not skip And (ds <> 0 Or anyTime) Then
            e = Oracle(ds, t, mQ1, mQ2)
            If Not RowMatches(e, v(i, 2), v(i, 3), v(i, 8), v(i, 9), ds) Then
                bad = bad + 1
                badList = badList & " row " & (i + 10) & ";"
                LogMismatch "As found", 0, i + 10, ds, t, e, v(i, 2), v(i, 3), v(i, 8), v(i, 9)
            End If
        End If
    Next i

    If used = 0 Then
        LogResult "B01", cat, "Entries saved in the file calculate correctly", "SKIP", , , "The form has no entries."
    Else
        LogResult "B01", cat, "Entries saved in the file calculate correctly", IIf(bad = 0, "PASS", "FAIL"), _
                  used & " rows correct", (used - bad) & " rows correct", badList & notes
    End If
    LogResult "B02", cat, "Totals of the entries saved in the file", "INFO", , _
              "Travel " & ToText(mForm.Range("H71").Value) & " h, Work " & ToText(mForm.Range("I71").Value) & " h, Total " & ToText(mForm.Range("K71").Value) & " h"
    If Len(months) > 0 Then
        If claimMonth > 0 Then
            LogResult "B03", cat, "All dates are in the claim month " & mForm.Range(MONTH_CELL).Value & " (or the last day before it)", _
                      IIf(Len(outside) = 0, "PASS", "WARN"), "all inside", IIf(Len(outside) = 0, "all inside", "outside:" & outside)
        Else
            LogResult "B03", cat, "All dates are in one claim month", IIf(InStr(2, Trim$(months), " ") = 0, "PASS", "WARN"), "1 month", Trim$(months)
        End If
    End If
    If Len(yrs) > 0 Then
        LogResult "B04", cat, "All dates are in the holiday table's year (" & mHolYear & ")", "WARN", CStr(mHolYear), Trim$(yrs), _
                  "Public holidays in other years are not recognised."
    End If
    If lastD > 0 Then LogResult "B05", cat, "Dates are in chronological order", IIf(unsorted, "WARN", "PASS")
    If Len(dupes) > 0 Then LogResult "B06", cat, "Same date on consecutive rows", "INFO", , Trim$(dupes), "Allowed (two jobs in a day), listed for review."
    If Len(noTimes) > 0 Then LogResult "B07", cat, "Rows with a date but no times (0 hours claimed)", "INFO", , "rows" & noTimes
    If Len(noDate) > 0 Then LogResult "B08", cat, "Rows with times but no date", "WARN", "none", "rows" & noDate, _
                                   "Hours on a row without a date are silently dropped from the total."
    If Len(badJ) > 0 Then LogResult "B09", cat, "LOCAL / OVERSEAS values valid", "WARN", "LOCAL or OVERSEAS", Trim$(badJ)
    If Len(numK) > 0 Then LogResult "B10", cat, "Project IDs stored as text", "WARN", "text", "numbers in" & numK, _
                                  "Numeric IDs lose trailing zeros (100084981.010 -> 100084981.01)."
    overlaps = OverlapList(spanS, spanE, hasSpan)
    If used > 0 Then
        LogResult "B11", cat, "No two rows claim the same hours (overlapping times)", IIf(Len(overlaps) = 0, "PASS", "WARN"), _
                  "no overlaps", IIf(Len(overlaps) = 0, "no overlaps", overlaps), _
                  IIf(Len(overlaps) = 0, "", "The form pays each row separately, so overlapping time is paid twice. Check these rows.")
    End If
    Exit Sub
Boom:
    LogCrash "B99", cat
End Sub

'==============================================================================
'  R - POWER QUERY REFRESH STRESS
'==============================================================================
Private Sub TestRefreshStress(ByRef base As Variant)
    Dim cat As String, i As Long, ok As Boolean, errMsg As String, t As Double, ms As Double
    Dim sig As String, firstSig As String, nOk As Long, diffSig As Long, firstErr As String
    Dim wrongPH As Long, r As Long, v As Variant, expPH As String, ds As Double, st As Variant
    cat = "Query refresh"
    On Error GoTo Boom

    If REFRESH_ITERATIONS <= 0 Then
        LogResult "R01", cat, "Power Query refresh stress", "SKIP", , , "REFRESH_ITERATIONS = 0"
        Exit Sub
    End If
    For i = 1 To REFRESH_ITERATIONS
        Progress "Power Query refresh " & i & " of " & REFRESH_ITERATIONS
        t = NowMs()
        ok = RefreshHolidayQuery(errMsg)
        ms = NowMs() - t
        If ok Then
            nOk = nOk + 1
            Recalc
            LoadHolidays
            sig = TableSignature()
            If Len(firstSig) = 0 Then
                firstSig = sig
            ElseIf sig <> firstSig Then
                diffSig = diffSig + 1
            End If
            ' the form's public-holiday column must follow the refreshed table
            v = mForm.Range(mForm.Cells(FIRST_ROW, 1), mForm.Cells(LAST_ROW, 3)).Value
            For r = 1 To NROWS
                If VarType(base(r, 1)) = vbDouble Then
                    ds = base(r, 1)
                    expPH = IIf(IsHoliday(ds), "Y", "N")
                    If IsError(v(r, 3)) Then
                        wrongPH = wrongPH + 1
                    ElseIf CStr(v(r, 3)) <> expPH Then
                        wrongPH = wrongPH + 1
                    End If
                End If
            Next r
            AddTiming "Query refresh", i, ms, True, mHolCount & " rows"
        Else
            If Len(firstErr) = 0 Then firstErr = errMsg
            AddTiming "Query refresh", i, ms, False, errMsg
        End If
        If i < REFRESH_ITERATIONS Then PauseMs REFRESH_PAUSE_MS
    Next i

    If nOk = REFRESH_ITERATIONS Then
        LogResult "R01", cat, "Query refreshes without error (" & REFRESH_ITERATIONS & " times in a row)", "PASS", _
                  REFRESH_ITERATIONS & " / " & REFRESH_ITERATIONS, nOk & " / " & REFRESH_ITERATIONS
    Else
        LogResult "R01", cat, "Query refreshes without error (" & REFRESH_ITERATIONS & " times in a row)", _
                  IIf(nOk = 0, "FAIL", "WARN"), REFRESH_ITERATIONS & " / " & REFRESH_ITERATIONS, nOk & " / " & REFRESH_ITERATIONS, _
                  "First error: " & firstErr & " | Check internet access, that the query was refreshed once by hand on this PC " & _
                  "(to accept web-content access), and that mom.gov.sg is reachable."
    End If
    If nOk > 1 Then
        LogResult "R02", cat, "Every refresh returns identical holiday data", IIf(diffSig = 0, "PASS", "WARN"), _
                  "0 differences", diffSig & " differences", IIf(diffSig = 0, "", "The website returned different data between refreshes.")
    End If
    If nOk > 0 Then
        st = TimingStats("Query refresh")
        LogResult "R03", cat, "Average refresh time", IIf(st(2) <= SLOW_REFRESH_MS, "PASS", "WARN"), _
                  "<= " & Format$(SLOW_REFRESH_MS, "0") & " ms", Format$(st(2), "0") & " ms", _
                  "min " & Format$(st(1), "0") & " / p95 " & Format$(st(3), "0") & " / max " & Format$(st(4), "0") & " ms"
        LogResult "R04", cat, "Form's public-holiday column follows the refreshed table after every refresh", _
                  IIf(wrongPH = 0, "PASS", "FAIL"), "0 wrong", wrongPH & " wrong"
    End If
    Exit Sub
Boom:
    LogCrash "R99", cat
End Sub

'==============================================================================
'  H - CONTENT OF THE QUERY'S HOLIDAY TABLE
'==============================================================================
Private Sub TestHolidayOutput()
    Dim cat As String, lo As ListObject, hdr As Variant, v As Variant, n As Long, i As Long, j As Long
    Dim bad As String, dayNames As Variant, d As Double, nm As String, tp As String, yrs As String
    Dim allowed As String, fixedList As Variant, named As Variant, cnt As Long, below As Range, r As Long
    Dim aA() As Variant, expected() As String, got As Variant, k As Long, w As String
    cat = "Holiday table"
    On Error GoTo Boom

    Set lo = GetHolidayTable()
    If lo Is Nothing Then
        LogResult "H00", cat, "Holiday table present", "FAIL", TABLE_NAME, "missing"
        Exit Sub
    End If
    hdr = lo.HeaderRowRange.Value
    If lo.ListColumns.Count = 4 And hdr(1, 1) = "Date" And hdr(1, 2) = "Day" And hdr(1, 3) = "Holiday name" And hdr(1, 4) = "Type" Then
        LogResult "H01", cat, "Columns are Date, Day, Holiday name, Type", "PASS"
    Else
        LogResult "H01", cat, "Columns are Date, Day, Holiday name, Type", "FAIL", "Date | Day | Holiday name | Type", JoinRow(hdr)
    End If
    If lo.DataBodyRange Is Nothing Then
        LogResult "H02", cat, "Number of holiday rows", "FAIL", "11-16", "0"
        Exit Sub
    End If
    v = lo.DataBodyRange.Value2
    n = UBound(v, 1)
    LogResult "H02", cat, "Number of holiday rows", IIf(n >= 11 And n <= 16, "PASS", "WARN"), "11-16", CStr(n), _
              IIf(n >= 11 And n <= 16, "", "Singapore has 11 gazetted holiday days plus in-lieu days; an unusual count can mean MOM changed the page.")

    bad = "": yrs = ""
    For i = 1 To n
        If VarType(v(i, 1)) <> vbDouble Then
            bad = bad & " row " & i
        ElseIf InStr(yrs, CStr(Year(v(i, 1)))) = 0 Then
            yrs = yrs & " " & Year(v(i, 1))
        End If
    Next i
    LogResult "H03", cat, "Every Date is a real date", IIf(Len(bad) = 0, "PASS", "FAIL"), "all dates", IIf(Len(bad) = 0, "all dates", "text/blank:" & bad)
    If Len(bad) > 0 Then Exit Sub
    LogResult "H04", cat, "All holidays are in the current year", IIf(Trim$(yrs) = CStr(Year(Date)), "PASS", "WARN"), _
              CStr(Year(Date)), Trim$(yrs)

    dayNames = Array("Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday")
    bad = ""
    For i = 1 To n
        If CStr(v(i, 2)) <> dayNames(Weekday(v(i, 1), vbMonday) - 1) Then bad = bad & " " & Format$(v(i, 1), "d-mmm") & "=" & v(i, 2)
    Next i
    LogResult "H05", cat, "Day column matches the date", IIf(Len(bad) = 0, "PASS", "FAIL"), , bad

    allowed = "|Public Holiday|Public Holiday in lieu|"
    bad = ""
    For i = 1 To n
        If InStr(allowed, "|" & CStr(v(i, 4)) & "|") = 0 Then bad = bad & " '" & v(i, 4) & "'"
    Next i
    LogResult "H06", cat, "Type is 'Public Holiday' or 'Public Holiday in lieu'", IIf(Len(bad) = 0, "PASS", "FAIL"), , bad

    bad = ""
    For i = 1 To n
        nm = CStr(v(i, 3))
        If Len(nm) = 0 Or nm Like "*[0-9]*" Or nm <> Trim$(nm) Or InStr(nm, Chr$(160)) > 0 Then bad = bad & " '" & nm & "'"
    Next i
    LogResult "H07", cat, "Holiday names are clean (no dates, digits or stray spaces)", IIf(Len(bad) = 0, "PASS", "FAIL"), , bad

    bad = ""
    For i = 1 To n
        If v(i, 4) = "Public Holiday in lieu" Then
            nm = CStr(v(i, 3))
            If Not nm Like "* (In lieu)" Then bad = bad & " name '" & nm & "';"
            If Weekday(v(i, 1), vbMonday) <> 1 Then bad = bad & " " & Format$(v(i, 1), "d-mmm") & " not a Monday;"
            If Not HolidayRowExists(v, v(i, 1) - 1, Replace(nm, " (In lieu)", ""), "Public Holiday") Then
                bad = bad & " no Sunday holiday before " & Format$(v(i, 1), "d-mmm") & ";"
            End If
        End If
    Next i
    LogResult "H08", cat, "In-lieu rows are Mondays following a Sunday holiday of the same name", IIf(Len(bad) = 0, "PASS", "WARN"), , bad

    bad = ""
    For i = 1 To n
        If v(i, 4) = "Public Holiday" And Weekday(v(i, 1), vbMonday) = 7 Then
            If Not HolidayRowExists(v, v(i, 1) + 1, CStr(v(i, 3)) & " (In lieu)", "Public Holiday in lieu") Then
                bad = bad & " " & v(i, 3) & " " & Format$(v(i, 1), "d-mmm") & ";"
            End If
        End If
    Next i
    LogResult "H09", cat, "Every Sunday holiday has its Monday in-lieu day", IIf(Len(bad) = 0, "PASS", "WARN"), , bad, _
              IIf(Len(bad) = 0, "", "The query only adds in-lieu days that MOM's page spells out.")

    bad = ""
    For i = 2 To n
        If v(i, 1) < v(i - 1, 1) Then bad = bad & " row " & i
    Next i
    LogResult "H10", cat, "Holidays sorted by date", IIf(Len(bad) = 0, "PASS", "FAIL"), , bad

    bad = ""
    For i = 1 To n - 1
        For j = i + 1 To n
            If v(i, 1) = v(j, 1) And v(i, 3) = v(j, 3) And v(i, 4) = v(j, 4) Then bad = bad & " " & Format$(v(i, 1), "d-mmm")
        Next j
    Next i
    LogResult "H11", cat, "No duplicate holidays", IIf(Len(bad) = 0, "PASS", "FAIL"), , bad

    fixedList = Array(1, 1, "New Year", 5, 1, "Labour Day", 8, 9, "National Day", 12, 25, "Christmas")
    bad = ""
    For i = 0 To UBound(fixedList) Step 3
        d = DateSerial(mHolYear, fixedList(i), fixedList(i + 1))
        If Not HolidayRowExists(v, d, CStr(fixedList(i + 2)), "", True) Then bad = bad & " " & fixedList(i + 2)
    Next i
    LogResult "H12", cat, "Fixed-date holidays present (1 Jan, 1 May, 9 Aug, 25 Dec)", IIf(Len(bad) = 0, "PASS", "FAIL"), , _
              IIf(Len(bad) = 0, "all 4", "missing:" & bad)

    named = Array("Chinese New Year", "Good Friday", "Hari Raya Puasa", "Hari Raya Haji", "Vesak Day", "Deepavali")
    bad = ""
    For i = 0 To UBound(named)
        cnt = 0
        For j = 1 To n
            If InStr(1, v(j, 3), named(i), vbTextCompare) = 1 And v(j, 4) = "Public Holiday" Then cnt = cnt + 1
        Next j
        If cnt = 0 Or (named(i) = "Chinese New Year" And cnt < 2) Then bad = bad & " " & named(i)
    Next i
    LogResult "H13", cat, "Moving holidays present (CNY x2, Good Friday, Hari Raya x2, Vesak, Deepavali)", _
              IIf(Len(bad) = 0, "PASS", "WARN"), , IIf(Len(bad) = 0, "all present", "missing:" & bad)

    Set below = lo.Range.Rows(lo.Range.Rows.Count).Offset(1, 0)
    LogResult "H14", cat, "No stale rows left under the table", IIf(Application.WorksheetFunction.CountA(below) = 0, "PASS", "WARN"), _
              "empty", below.Address(False, False) & IIf(Application.WorksheetFunction.CountA(below) = 0, " empty", " has data")

    ' H15 put every holiday (and the day after it) into the form: PH must be Y / N
    ClearInputs
    ReDim aA(1 To NROWS, 1 To 1)
    ReDim expected(1 To NROWS)
    k = 0
    For i = 1 To n
        If k < NROWS Then
            k = k + 1: aA(k, 1) = CDate(v(i, 1)): expected(k) = "Y"
        End If
        If k < NROWS And Not IsHoliday(v(i, 1) + 1) Then
            k = k + 1: aA(k, 1) = CDate(v(i, 1) + 1): expected(k) = "N"
        End If
    Next i
    mForm.Range(mForm.Cells(FIRST_ROW, 1), mForm.Cells(LAST_ROW, 1)).Value = aA
    Recalc
    got = mForm.Range(mForm.Cells(FIRST_ROW, 1), mForm.Cells(LAST_ROW, 3)).Value
    bad = ""
    For r = 1 To k
        If IsError(got(r, 3)) Then
            bad = bad & " " & Format$(aA(r, 1), "d-mmm") & "=" & ErrText(got(r, 3))
        ElseIf CStr(got(r, 3)) <> expected(r) Then
            bad = bad & " " & Format$(aA(r, 1), "d-mmm") & "=" & got(r, 3)
        End If
        w = Format$(aA(r, 1), "ddd")
        If Not IsError(got(r, 2)) Then
            If CStr(got(r, 2)) <> w And CStr(got(r, 2)) <> Application.WorksheetFunction.Text(aA(r, 1), "ddd") Then bad = bad & " day " & got(r, 2)
        End If
    Next r
    LogResult "H15", cat, "Form flags every holiday Y and the day after N", IIf(Len(bad) = 0, "PASS", "FAIL"), _
              k & " dates correct", IIf(Len(bad) = 0, k & " dates correct", "wrong:" & bad)
    ClearInputs
    Exit Sub
Boom:
    LogCrash "H99", cat
End Sub

'==============================================================================
'  K - KNOWN ANSWERS (hand-checked expected hours, office hours 08:00-17:30)
'==============================================================================
Private Sub TestKnownAnswers()
    Dim cat As String, fri As Date, sat As Date, sat2 As Date, sun As Date, tue As Date, wed As Date, ph As Double
    Dim g As Variant, ok As Boolean
    cat = "Known answers"
    On Error GoTo Boom
    fri = DateSerial(2026, 9, 18): sat = DateSerial(2026, 9, 19): sat2 = DateSerial(2026, 9, 26)
    sun = DateSerial(2026, 9, 27): tue = DateSerial(2026, 9, 22): wed = DateSerial(2026, 9, 30)

    KA "K01", "Weekday overnight travel 21:00-00:00", fri, "21:00", "00:00", "", "", 3, 0, True
    KA "K02", "Saturday travel-work-travel from midnight", sat, "00:00", "01:00", "06:30", "09:30", 4, 5.5, False
    KA "K03", "Weekday: travel before office hours only counts outside 08:00", tue, "06:30", "08:30", "11:30", "12:30", 1.5, 0, True
    KA "K04", "Saturday ending exactly at midnight", sat2, "16:00", "20:30", "22:30", "00:00", 6, 2, False
    KA "K05", "Sunday: From + Until only (travel straight through)", sun, "00:00", "", "", "01:00", 1, 0, False
    KA "K06", "Date with no times gives blank hours", wed, "", "", "", "", "", "", False
    KA "K07", "Weekday: only the 17:30-18:00 part is overtime", tue, "08:00", "09:00", "17:00", "18:00", 0.5, 0, True
    KA "K08", "Weekday after office hours", tue, "17:30", "18:00", "23:00", "23:30", 1, 5, True
    KA "K09", "Weekday work crosses midnight", tue, "20:00", "22:00", "02:00", "03:00", 3, 4, True
    KA "K10", "Weekday work straddles office start", tue, "05:00", "07:00", "10:00", "11:00", 2, 1, True
    KA "K11", "Saturday with From blank", sat, "", "09:00", "10:00", "11:00", 1, 1, False
    KA "K12", "Saturday with 2nd column blank (work = From to 3rd)", sat, "09:00", "", "13:00", "14:00", 1, 4, False
    KA "K13", "Saturday with 3rd column blank (work = 2nd to Until)", sat, "09:00", "10:00", "", "14:00", 1, 4, False
    KA "K15", "All four times equal", sat, "09:00", "09:00", "09:00", "09:00", 0, 0, False
    KA "K16", "Times without a date are ignored", Empty, "09:00", "10:00", "", "", "", "", False
    KA "K17", "Two minutes across midnight", sat, "23:59", "00:01", "", "", 0.03, 0, False
    KA "K18", "Two minutes across 08:00 on a weekday", tue, "07:59", "08:01", "", "", 0.02, 0, True
    KA "K19", "Weekday travel fully inside office hours", tue, "12:00", "13:00", "", "", 0, 0, True
    KA "K20", "Weekday 00:00-23:59 minus office hours", tue, "00:00", "", "", "23:59", 14.48, 0, True
    KA "K21", "Saturday 00:00-23:59", sat, "00:00", "", "", "23:59", 23.98, 0, False

    ph = FirstWeekdayHoliday()
    If ph = 0 Then
        LogResult "K14", cat, "Weekday public holiday pays every hour", "SKIP", , , "No weekday holiday in the table."
    Else
        KA "K14", "Weekday public holiday (" & Format$(ph, "d mmm yyyy") & ") pays every hour", CDate(ph), _
           "08:00", "09:00", "17:00", "18:00", 2, 8, False
    End If

    ' K22 the six sample rows shipped in the file, all at once, and the totals
    If mQ1 <> DEFAULT_Q1 Or mQ2 <> DEFAULT_Q2 Then
        LogResult "K22", cat, "Sample month totals (Travel 15.5, Work 7.5, Total 23)", "SKIP", , , "Office hours are not 08:00-17:30."
    Else
        ClearInputs
        PutRow 11, fri, "21:00", "00:00", "", ""
        PutRow 12, sat, "00:00", "01:00", "06:30", "09:30"
        PutRow 13, tue, "06:30", "08:30", "11:30", "12:30"
        PutRow 14, sat2, "16:00", "20:30", "22:30", "00:00"
        PutRow 15, sun, "00:00", "", "", "01:00"
        PutRow 16, wed, "", "", "", ""
        Recalc
        g = mForm.Range("H71:K71").Value
        ok = SameValue(g(1, 1), 15.5) And SameValue(g(1, 2), 7.5) And SameValue(g(1, 4), 23)
        LogResult "K22", cat, "Sample month totals (Travel 15.5, Work 7.5, Total 23)", IIf(ok, "PASS", "FAIL"), _
                  "15.5 / 7.5 / 23", ToText(g(1, 1)) & " / " & ToText(g(1, 2)) & " / " & ToText(g(1, 4))
    End If
    ClearInputs
    Exit Sub
Boom:
    LogCrash "K99", cat
End Sub

Private Sub KA(ByVal id As String, ByVal title As String, ByVal dVal As Variant, ByVal sD As String, ByVal sE As String, _
               ByVal sF As String, ByVal sG As String, ByVal expT As Variant, ByVal expW As Variant, ByVal usesOfficeHours As Boolean)
    Dim g As Variant, ok As Boolean
    If usesOfficeHours And (mQ1 <> DEFAULT_Q1 Or mQ2 <> DEFAULT_Q2) Then
        LogResult id, "Known answers", title, "SKIP", , , "Expected values assume office hours 08:00-17:30; the form uses " & HHMM(mQ1) & "-" & HHMM(mQ2) & "."
        Exit Sub
    End If
    ClearInputs
    PutRow FIRST_ROW, dVal, sD, sE, sF, sG
    Recalc
    g = mForm.Range(mForm.Cells(FIRST_ROW, 1), mForm.Cells(FIRST_ROW, 9)).Value
    ok = SameValue(g(1, 8), expT) And SameValue(g(1, 9), expW)
    If IsEmpty(dVal) Then ok = ok And Not IsError(g(1, 2)) And Not IsError(g(1, 3))
    If ok And IsEmpty(dVal) Then ok = (Trim$(CStr(g(1, 2))) = "" And CStr(g(1, 3)) = "")
    LogResult id, "Known answers", title, IIf(ok, "PASS", "FAIL"), FmtPair(expT, expW), FmtPair(g(1, 8), g(1, 9)), _
              DescribeInput(dVal, sD, sE, sF, sG)
End Sub

'==============================================================================
'  L - KNOWN LIMITATIONS (WARN = the form gives a different answer from reality)
'==============================================================================
Private Sub TestLimitations()
    Dim cat As String, sat As Date, tue As Date, ph As Double, y As Long, d As Date, g As Variant
    cat = "Limitations"
    On Error GoTo Boom
    sat = DateSerial(2026, 9, 19): tue = DateSerial(2026, 9, 22)

    LimitCase "L01", "A 24-hour trip (00:00 to 00:00 next day)", sat, "00:00", "", "", "00:00", 24, 0, _
              "Durations of 24 h or more collapse to 0 h. Split the trip over two rows."
    LimitCase "L02", "A shift that crosses midnight twice", sat, "22:00", "02:00", "21:00", "01:00", 8, 19, _
              "Only one midnight roll-over is supported; the last travel leg is lost. Split over two rows."

    ph = FirstWeekdayHoliday()
    If ph > 0 Then
        LimitCase "L03", "Public-holiday date pasted with a time (" & Format$(ph + 0.5, "d mmm hh:nn") & ")", CDate(ph + 0.5), _
                  "08:00", "09:00", "17:00", "18:00", 2, 8, _
                  "COUNTIF needs an exact date; a date with a time part is not recognised as a holiday."
    End If

    For y = mHolYear - 1 To mHolYear - 7 Step -1
        d = DateSerial(y, 12, 25)
        If Weekday(d, vbMonday) <= 5 Then Exit For
    Next y
    LimitCase "L04", "Christmas of the previous year (" & Format$(d, "d mmm yyyy") & ")", d, "08:00", "09:00", "17:00", "18:00", 2, 8, _
              "The query loads the current year only, so last year's holidays are paid at weekday rates."
    For y = mHolYear + 1 To mHolYear + 7
        d = DateSerial(y, 1, 1)
        If Weekday(d, vbMonday) <= 5 Then Exit For
    Next y
    LimitCase "L05", "New Year's Day of a later year (" & Format$(d, "d mmm yyyy") & ")", d, "08:00", "09:00", "17:00", "18:00", 2, 8, _
              "The query loads the current year only."

    ClearInputs
    PutRow FIRST_ROW, sat, "10:00", "", "", ""
    Recalc
    g = mForm.Range("H11:I11").Value
    If SameValue(g(1, 1), 0) And SameValue(g(1, 2), 0) Then
        LogResult "L06", cat, "Incomplete entry (only 'From' filled) is flagged", "WARN", "a warning", "0 h / 0 h, no warning", _
                  "An engineer who forgets the end time silently claims 0 hours."
    Else
        LogResult "L06", cat, "Incomplete entry (only 'From' filled) is flagged", "PASS", , FmtPair(g(1, 1), g(1, 2))
    End If

    ' L08 two jobs on the same day whose times overlap
    ClearInputs
    PutRow FIRST_ROW, sat, "07:00", "", "", "10:00"
    PutRow FIRST_ROW + 1, sat, "09:00", "", "", "11:00"
    Recalc
    g = mForm.Range("H71").Value
    LogResult "L08", cat, "Two rows on the same day with overlapping times (Sat 07:00-10:00 and 09:00-11:00)", _
              IIf(SameValue(g, 4), "PASS", "WARN"), "4 h (07:00-11:00 once)", ToText(g) & " h travel", _
              "The form has no overlap check, so 09:00-10:00 is paid twice. The stress test's own overlap check " & _
              "(B11) " & IIf(Len(OverlapOfFormRows(FIRST_ROW, FIRST_ROW + 1)) > 0, "detects", "DID NOT detect") & " this case."
    LogResult "L09", cat, "Form highlights the two overlapping rows", _
              IIf(Highlighted(FIRST_ROW) And Highlighted(FIRST_ROW + 1), "PASS", "WARN"), "rows 11 and 12 red", _
              "row 11 " & IIf(Highlighted(FIRST_ROW), "red", "not red") & ", row 12 " & IIf(Highlighted(FIRST_ROW + 1), "red", "not red"), _
              "Needs the overlap rule on D11:G70 and the helper columns K11:L70 on '" & SH_ENGINE & "'."
    ClearInputs
    PutRow FIRST_ROW, sat, "07:00", "", "", "09:00"
    PutRow FIRST_ROW + 1, sat, "09:00", "", "", "11:00"
    Recalc
    LogResult "L10", cat, "Back-to-back rows (07:00-09:00, 09:00-11:00) are not highlighted", _
              IIf(Highlighted(FIRST_ROW) Or Highlighted(FIRST_ROW + 1), "WARN", "PASS"), "not red", _
              IIf(Highlighted(FIRST_ROW) Or Highlighted(FIRST_ROW + 1), "red", "not red")

    If mQ1 = DEFAULT_Q1 And mQ2 = DEFAULT_Q2 Then
        ClearInputs
        PutRow FIRST_ROW, tue, "22:00", "09:00", "", ""
        Recalc
        g = mForm.Range("H11").Value
        LogResult "L07", cat, "Overnight weekday travel into next morning's office hours (Tue 22:00 - Wed 09:00)", "INFO", _
                  "policy decision", ToText(g) & " h travel", _
                  "Only the start day's office hours are deducted, so 08:00-09:00 on Wednesday counts as overtime (11 h). " & _
                  "If policy says it should not, the answer would be 10 h."
    End If
    ClearInputs
    Exit Sub
Boom:
    LogCrash "L99", cat
End Sub

Private Sub LimitCase(ByVal id As String, ByVal title As String, ByVal dVal As Variant, ByVal sD As String, ByVal sE As String, _
                      ByVal sF As String, ByVal sG As String, ByVal realT As Double, ByVal realW As Double, ByVal why As String)
    Dim g As Variant
    ClearInputs
    PutRow FIRST_ROW, dVal, sD, sE, sF, sG
    Recalc
    g = mForm.Range("H11:I11").Value
    If SameValue(g(1, 1), realT) And SameValue(g(1, 2), realW) Then
        LogResult id, "Limitations", title, "PASS", FmtPair(realT, realW), FmtPair(g(1, 1), g(1, 2)), "Handled correctly."
    Else
        LogResult id, "Limitations", title, "WARN", FmtPair(realT, realW), FmtPair(g(1, 1), g(1, 2)), _
                  why & " Inputs: " & DescribeInput(dVal, sD, sE, sF, sG)
    End If
End Sub

'==============================================================================
'  P - PASTED / BAD INPUT (data validation does not stop paste or VBA)
'==============================================================================
Private Sub TestRobustness()
    Dim cat As String, tue As Date, sat As Date, g As Variant, tot As Variant, longText As String
    cat = "Bad input"
    On Error GoTo Boom
    tue = DateSerial(2026, 9, 22): sat = DateSerial(2026, 9, 19)

    ' P01 text pasted into a time column
    ClearInputs
    PutRow 12, sat, "09:00", "10:00", "", ""       ' a good row so the total has something in it
    PutRow FIRST_ROW, tue, "", "10:00", "17:00", "18:00"
    mForm.Range("D11").Value = "abc"
    Recalc
    ReportBadInput "P01", "Text ('abc') pasted into a time cell"

    ' P02 date pasted as text
    ClearInputs
    PutRow 12, sat, "09:00", "10:00", "", ""
    PutRow FIRST_ROW, Empty, "09:00", "10:00", "17:00", "18:00"
    mForm.Range("A11").Value = "'25/12/2026"
    Recalc
    ReportBadInput "P02", "Date pasted as text ('25/12/2026')"

    ' P03 / P04 out-of-range times
    ClearInputs
    PutRow FIRST_ROW, sat, "09:00", "", "", ""
    mForm.Range("E11").Value = 25# / 24#
    Recalc
    g = mForm.Range("H11").Value
    LogResult "P03", cat, "Time of 25:00 pasted into a time cell", IIf(IsError(g), "PASS", "WARN"), "rejected / error", _
              ToText(g) & " h accepted", "Only the red highlight shows the problem; the hours still reach the total."
    ClearInputs
    PutRow FIRST_ROW, sat, "", "10:00", "", ""
    mForm.Range("D11").Value = -0.1
    Recalc
    g = mForm.Range("H11").Value
    LogResult "P04", cat, "Negative time pasted into a time cell", IIf(IsError(g), "PASS", "WARN"), "rejected / error", _
              ToText(g) & " h accepted", "Only the red highlight shows the problem; the hours still reach the total."

    ' P05 seconds are rounded sensibly
    ClearInputs
    PutRow FIRST_ROW, sat, "", "", "", ""
    mForm.Range("D11").Value = (9# * 3600# + 30#) / 86400#
    mForm.Range("E11").Value = (10# * 3600# + 59#) / 86400#
    Recalc
    g = mForm.Range("H11").Value
    LogResult "P05", cat, "Times with seconds (09:00:30 to 10:00:59)", IIf(SameValue(g, 1.01), "PASS", "FAIL"), "1.01", ToText(g)

    ' P06 long text and Project ID with trailing zero
    ClearInputs
    longText = String$(255, "X")
    If mColLocal > 0 Then mForm.Cells(FIRST_ROW, mColLocal).Value = "OVERSEAS"
    mForm.Cells(FIRST_ROW, mColProject).Value = "100084981.010"
    mForm.Cells(FIRST_ROW, mColVessel).Value = longText
    Recalc
    g = mForm.Cells(FIRST_ROW, mColProject).Value
    tot = mForm.Cells(FIRST_ROW, mColVessel).Value
    If mColActivity > 0 Then
        mForm.Cells(FIRST_ROW + 1, mColActivity).Value = "0010"
        mForm.Cells(FIRST_ROW + 2, mColActivity).Value = "1234567890123456789"
        Recalc
        LogResult "P08", cat, "Activity number keeps what was typed (0010, 19-digit number)", _
                  IIf(SameValue(mForm.Cells(FIRST_ROW + 1, mColActivity).Value, "0010") And _
                      SameValue(mForm.Cells(FIRST_ROW + 2, mColActivity).Value, "1234567890123456789"), "PASS", "WARN"), _
                  "0010 / 1234567890123456789", ToText(mForm.Cells(FIRST_ROW + 1, mColActivity).Text) & " / " & _
                  ToText(mForm.Cells(FIRST_ROW + 2, mColActivity).Text), _
                  "Format the Activity number column as Text (like Project ID) so Excel does not turn entries into numbers."
    End If
    LogResult "P06", cat, "Project ID keeps trailing zeros; 255-character vessel name kept", _
              IIf(SameValue(g, "100084981.010") And Len(ToText(tot)) = 255, "PASS", "FAIL"), _
              "100084981.010 / 255 chars", ToText(g) & " / " & Len(CStr(tot)) & " chars"

    ' P07 dates before the validation minimum still calculate
    ClearInputs
    PutRow FIRST_ROW, DateSerial(2019, 12, 31), "18:00", "19:00", "", ""
    Recalc
    g = mForm.Range("H11").Value
    LogResult "P07", cat, "Date before 1 Jan 2020 (validation minimum) pasted in", "INFO", , ToText(g) & " h", _
              "Calculates normally; only typed entries are blocked."
    ClearInputs
    Exit Sub
Boom:
    LogCrash "P99", cat
End Sub

Private Sub ReportBadInput(ByVal id As String, ByVal title As String)
    Dim g As Variant, tot As Variant
    g = mForm.Range("H11:I11").Value
    tot = mForm.Range("K71").Value
    If IsError(tot) Then
        LogResult id, "Bad input", title, "WARN", "row flagged, total still valid", _
                  "row " & ToText(g(1, 1)) & " / " & ToText(g(1, 2)) & ", TOTAL " & ErrText(tot), _
                  "One bad pasted value turns the whole claim TOTAL into " & ErrText(tot) & "."
    ElseIf IsError(g(1, 1)) Or IsError(g(1, 2)) Then
        LogResult id, "Bad input", title, "PASS", "row flagged, total still valid", "row error, TOTAL " & ToText(tot)
    Else
        LogResult id, "Bad input", title, "WARN", "row flagged", "row " & FmtPair(g(1, 1), g(1, 2)) & ", TOTAL " & ToText(tot), _
                  "The bad value was silently accepted."
    End If
End Sub

'==============================================================================
'  Z / O - RANDOM FULL-FORM FILLS COMPARED WITH THE REFERENCE MODEL
'==============================================================================
Private Sub TestFuzz(ByVal iterations As Long, ByVal idPrefix As String, ByVal label As String)
    Dim cat As String, it As Long, i As Long, c As Long
    Dim aA(1 To NROWS, 1 To 1) As Variant, aT(1 To NROWS, 1 To 4) As Variant, aJ(1 To NROWS, 1 To 4) As Variant
    Dim tm(1 To NROWS, 0 To 3) As Long, ds(1 To NROWS) As Double, t(0 To 3) As Long
    Dim v As Variant, tot As Variant, e As OTResult, vessels As Variant
    Dim rowsChecked As Long, rowsBad As Long, totBad As Long, idBad As Long
    Dim sumT As Double, sumW As Double, t0 As Double, tW As Double, tC As Double, tR As Double, st As Variant
    cat = label
    On Error GoTo Boom
    If iterations <= 0 Then Exit Sub
    vessels = Array("NORD HARMONY", "LMV19", "PAGNA", "LEVERKRUSEN EXPRESS", "TEST VESSEL", "MV A-B/C 'Q'")

    For it = 1 To iterations
        If it Mod 10 = 1 Then Progress label & " " & it & " of " & iterations
        For i = 1 To NROWS
            ds(i) = RandomDate()
            GenTimes t
            For c = 0 To 3
                tm(i, c) = t(c)
                If t(c) >= 0 Then aT(i, c + 1) = t(c) / 1440# Else aT(i, c + 1) = Empty
            Next c
            If ds(i) = 0 Then aA(i, 1) = Empty Else aA(i, 1) = CDate(ds(i))
            aJ(i, 1) = IIf(Rnd < 0.8, "LOCAL", "OVERSEAS")
            aJ(i, 2) = "1000" & Format$(Int(Rnd * 100000), "00000") & "." & Format$(Int(Rnd * 1000), "000")
            aJ(i, 3) = vessels(Int(Rnd * (UBound(vessels) + 1)))
            aJ(i, 4) = "A" & Format$(Int(Rnd * 10000), "0000")
        Next i

        t0 = NowMs()
        mForm.Range(mForm.Cells(FIRST_ROW, 1), mForm.Cells(LAST_ROW, 1)).Value = aA
        mForm.Range(mForm.Cells(FIRST_ROW, 4), mForm.Cells(LAST_ROW, 7)).Value = aT
        WriteJobInfo aJ
        tW = NowMs() - t0
        t0 = NowMs()
        Recalc
        tC = NowMs() - t0
        t0 = NowMs()
        v = mForm.Range(mForm.Cells(FIRST_ROW, 1), mForm.Cells(LAST_ROW, 12)).Value
        tot = mForm.Range(mForm.Cells(TOTAL_ROW, 8), mForm.Cells(TOTAL_ROW, 11)).Value

        sumT = 0: sumW = 0
        For i = 1 To NROWS
            For c = 0 To 3
                t(c) = tm(i, c)
            Next c
            e = Oracle(ds(i), t, mQ1, mQ2)
            rowsChecked = rowsChecked + 1
            If VarType(e.Travel) = vbDouble Then sumT = sumT + e.Travel
            If VarType(e.Work) = vbDouble Then sumW = sumW + e.Work
            If Not RowMatches(e, v(i, 2), v(i, 3), v(i, 8), v(i, 9), ds(i)) Then
                rowsBad = rowsBad + 1
                LogMismatch label, it, i + FIRST_ROW - 1, ds(i), t, e, v(i, 2), v(i, 3), v(i, 8), v(i, 9)
            End If
            If VarType(v(i, mColProject)) <> vbString Then
                idBad = idBad + 1
            ElseIf v(i, mColProject) <> aJ(i, 2) Then
                idBad = idBad + 1
            End If
        Next i
        If Not (SameValue(tot(1, 1), sumT) And SameValue(tot(1, 2), sumW) And SameValue(tot(1, 4), sumT + sumW)) Then totBad = totBad + 1
        tR = NowMs() - t0
        AddTiming label & " - write", it, tW, True, ""
        AddTiming label & " - recalc", it, tC, True, ""
        AddTiming label & " - read + verify", it, tR, True, ""
    Next it
    ClearInputs

    LogResult idPrefix & "01", cat, "Random rows match the reference model (" & iterations & " fills x 60 rows)", _
              IIf(rowsBad = 0, "PASS", "FAIL"), rowsChecked & " rows", (rowsChecked - rowsBad) & " rows match", _
              IIf(rowsBad = 0, "Office hours " & HHMM(mQ1) & "-" & HHMM(mQ2) & ", seed " & RANDOM_SEED, "See " & OUT_MISMATCH & ".")
    LogResult idPrefix & "02", cat, "Totals row equals the sum of the rows", IIf(totBad = 0, "PASS", "FAIL"), _
              iterations & " fills", (iterations - totBad) & " correct"
    LogResult idPrefix & "03", cat, "Project IDs stay text with trailing zeros", IIf(idBad = 0, "PASS", "FAIL"), "0 changed", idBad & " changed"
    st = TimingStats(label & " - recalc")
    LogResult idPrefix & "04", cat, "Average recalculation after filling 60 rows", IIf(st(2) <= SLOW_RECALC_MS, "PASS", "WARN"), _
              "<= " & SLOW_RECALC_MS & " ms", Format$(st(2), "0.0") & " ms", _
              "min " & Format$(st(1), "0.0") & " / p95 " & Format$(st(3), "0.0") & " / max " & Format$(st(4), "0.0") & " ms"
    Exit Sub
Boom:
    LogCrash idPrefix & "99", cat
    ClearInputs
End Sub

'==============================================================================
'  C - CAPACITY AND RECALCULATION PERFORMANCE
'==============================================================================
Private Sub TestCapacity()
    Dim cat As String, i As Long, c As Long, d0 As Double, t(0 To 3) As Long, e As OTResult
    Dim aA(1 To NROWS, 1 To 1) As Variant, aT(1 To NROWS, 1 To 4) As Variant, aJ(1 To NROWS, 1 To 4) As Variant
    Dim v As Variant, tot As Variant, bad As Long, sumT As Double, sumW As Double, ds(1 To NROWS) As Double
    Dim it As Long, t0 As Double, st As Variant
    cat = "Capacity"
    On Error GoTo Boom

    ' C01 all 60 rows used, near-24h days, 60 consecutive dates, maximum-length text
    d0 = DateSerial(mHolYear, 12, 1) - 30
    For i = 1 To NROWS
        ds(i) = d0 + i - 1
        aA(i, 1) = CDate(ds(i))
        aT(i, 1) = 0#: aT(i, 2) = 1# / 1440#: aT(i, 3) = 1438# / 1440#: aT(i, 4) = 1439# / 1440#
        aJ(i, 1) = "OVERSEAS": aJ(i, 2) = "999999999.999": aJ(i, 3) = String$(255, "V"): aJ(i, 4) = String$(50, "9")
    Next i
    mForm.Range(mForm.Cells(FIRST_ROW, 1), mForm.Cells(LAST_ROW, 1)).Value = aA
    mForm.Range(mForm.Cells(FIRST_ROW, 4), mForm.Cells(LAST_ROW, 7)).Value = aT
    WriteJobInfo aJ
    Recalc
    v = mForm.Range(mForm.Cells(FIRST_ROW, 1), mForm.Cells(LAST_ROW, 12)).Value
    tot = mForm.Range(mForm.Cells(TOTAL_ROW, 8), mForm.Cells(TOTAL_ROW, 11)).Value
    t(0) = 0: t(1) = 1: t(2) = 1438: t(3) = 1439
    For i = 1 To NROWS
        e = Oracle(ds(i), t, mQ1, mQ2)
        sumT = sumT + e.Travel: sumW = sumW + e.Work
        If Not RowMatches(e, v(i, 2), v(i, 3), v(i, 8), v(i, 9), ds(i)) Then
            bad = bad + 1
            LogMismatch "Capacity", 0, i + FIRST_ROW - 1, ds(i), t, e, v(i, 2), v(i, 3), v(i, 8), v(i, 9)
        End If
    Next i
    LogResult "C01", cat, "All 60 rows filled with near-24-hour days", IIf(bad = 0, "PASS", "FAIL"), "60 rows correct", (60 - bad) & " rows correct"
    LogResult "C02", cat, "Totals with a full form", _
              IIf(SameValue(tot(1, 1), sumT) And SameValue(tot(1, 2), sumW) And SameValue(tot(1, 4), sumT + sumW), "PASS", "FAIL"), _
              FmtPair(sumT, sumW) & " / " & Format$(sumT + sumW, "0.00"), FmtPair(tot(1, 1), tot(1, 2)) & " / " & ToText(tot(1, 4))
    LogResult "C03", cat, "Only rows 11-70 are totalled", "INFO", , "60 rows", _
              "A month with more than 60 entries needs a second page; row 71 is the total."

    ' C04 forced full recalculation of a full form
    For it = 1 To RECALC_ITERATIONS
        If it Mod 10 = 1 Then Progress "full recalculation " & it & " of " & RECALC_ITERATIONS
        t0 = NowMs()
        ForceRecalc
        AddTiming "Full recalculation", it, NowMs() - t0, True, ""
    Next it
    If RECALC_ITERATIONS > 0 Then
        st = TimingStats("Full recalculation")
        LogResult "C04", cat, "Full recalculation of a full form (" & RECALC_ITERATIONS & " times)", _
                  IIf(st(2) <= SLOW_RECALC_MS, "PASS", "WARN"), "<= " & SLOW_RECALC_MS & " ms", Format$(st(2), "0.0") & " ms", _
                  "min " & Format$(st(1), "0.0") & " / p95 " & Format$(st(3), "0.0") & " / max " & Format$(st(4), "0.0") & " ms"
    End If
    ClearInputs
    Exit Sub
Boom:
    LogCrash "C99", cat
    ClearInputs
End Sub

'==============================================================================
'  O - OTHER OFFICE HOURS (needs the form to be unprotected on the copy)
'==============================================================================
Private Sub TestOfficeHours()
    Dim cat As String, oldQ1 As Variant, oldQ2 As Variant, sets As Variant, k As Long
    Dim keepQ1 As Long, keepQ2 As Long, wasProtected As Boolean
    cat = "Office hours"
    On Error GoTo Boom
    If OFFICE_HOURS_FUZZ <= 0 Then Exit Sub

    wasProtected = mForm.ProtectContents
    If wasProtected Then
        On Error Resume Next
        mForm.Unprotect Password:=FORM_PASSWORD
        On Error GoTo Boom
        If mForm.ProtectContents Then
            LogResult "O00", cat, "Form with different office hours", "SKIP", , , _
                      "'" & SH_FORM & "' has a password. Put it in FORM_PASSWORD at the top of the module to run these tests."
            Exit Sub
        End If
    End If

    oldQ1 = mForm.Range("Q1").Value2: oldQ2 = mForm.Range("Q2").Value2
    keepQ1 = mQ1: keepQ2 = mQ2
    sets = Array(Array(450, 990), Array(540, 1080), Array(420, 1140), Array(0, 1439))
    SeedRandom
    For k = LBound(sets) To UBound(sets)
        mQ1 = sets(k)(0): mQ2 = sets(k)(1)
        mForm.Range("Q1").Value = mQ1 / 1440#
        mForm.Range("Q2").Value = mQ2 / 1440#
        Recalc
        TestFuzz OFFICE_HOURS_FUZZ, "O" & (k + 1), "Office hours " & HHMM(mQ1) & "-" & HHMM(mQ2)
    Next k
    mForm.Range("Q1").Value = oldQ1
    mForm.Range("Q2").Value = oldQ2
    mQ1 = keepQ1: mQ2 = keepQ2
    If wasProtected Then mForm.Protect Password:=FORM_PASSWORD
    Recalc
    Exit Sub
Boom:
    LogCrash "O99", cat
    mQ1 = keepQ1: mQ2 = keepQ2
End Sub

'==============================================================================
'  U - YOUR OWN / COPILOT-GENERATED SCENARIOS (sheet ST_Scenarios)
'==============================================================================
Private Sub TestScenarioSheet()
    Dim ws As Worksheet, r As Long, lastRow As Long, id As String, desc As String
    Dim ds As Double, t(0 To 3) As Long, c As Long, okParse As Boolean, why As String
    Dim g As Variant, e As OTResult, noResult As OTResult, expT As Variant, expW As Variant
    Dim status As String, note As String, gotText As String, nRun As Long
    On Error GoTo Boom
    Set ws = ThisWorkbook.Worksheets(OUT_SCENARIOS)
    ws.Range(ws.Cells(SCEN_FIRST_ROW, 11), ws.Cells(ws.Rows.Count, 16)).ClearContents
    ws.Range(ws.Cells(SCEN_FIRST_ROW, 15), ws.Cells(ws.Rows.Count, 15)).Interior.ColorIndex = xlColorIndexNone
    lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    If ws.Cells(ws.Rows.Count, 3).End(xlUp).Row > lastRow Then lastRow = ws.Cells(ws.Rows.Count, 3).End(xlUp).Row

    For r = SCEN_FIRST_ROW To lastRow
        If Application.WorksheetFunction.CountA(ws.Range(ws.Cells(r, 2), ws.Cells(r, 9))) > 0 Then
            nRun = nRun + 1
            id = Trim$(CStr(ws.Cells(r, 1).Value))
            If Len(id) = 0 Then id = "U" & Format$(nRun, "000"): ws.Cells(r, 1).Value = id
            desc = Trim$(CStr(ws.Cells(r, 2).Value))
            e = noResult
            gotText = "-"
            okParse = True: why = ""
            ds = ParseDateCell(ws.Cells(r, 3).Value, okParse)
            If Not okParse Then why = "Date not understood: '" & ws.Cells(r, 3).Text & "'. "
            For c = 0 To 3
                t(c) = ParseTimeCell(ws.Cells(r, 4 + c).Value)
                If t(c) = -2 Then okParse = False: why = why & "Time not understood: '" & ws.Cells(r, 4 + c).Text & "'. "
            Next c
            expT = ParseNumberCell(ws.Cells(r, 8).Value)
            expW = ParseNumberCell(ws.Cells(r, 9).Value)

            If Not okParse Then
                status = "INVALID": note = why
            Else
                ClearInputs
                If ds = 0 Then mForm.Cells(FIRST_ROW, 1).Value = Empty Else mForm.Cells(FIRST_ROW, 1).Value = CDate(ds)
                For c = 0 To 3
                    If t(c) >= 0 Then mForm.Cells(FIRST_ROW, 4 + c).Value = t(c) / 1440# Else mForm.Cells(FIRST_ROW, 4 + c).Value = Empty
                Next c
                Recalc
                g = mForm.Range(mForm.Cells(FIRST_ROW, 1), mForm.Cells(FIRST_ROW, 9)).Value
                e = Oracle(ds, t, mQ1, mQ2)
                ws.Cells(r, 11).Value = ToText(g(1, 8))
                ws.Cells(r, 12).Value = ToText(g(1, 9))
                ws.Cells(r, 13).Value = ToText(e.Travel)
                ws.Cells(r, 14).Value = ToText(e.Work)
                gotText = FmtPair(g(1, 8), g(1, 9))
                If Not RowMatches(e, g(1, 2), g(1, 3), g(1, 8), g(1, 9), ds) Then
                    status = "FAIL": note = "Form differs from the reference model."
                    LogMismatch "Scenario " & id, 0, FIRST_ROW, ds, t, e, g(1, 2), g(1, 3), g(1, 8), g(1, 9)
                ElseIf Not IsEmpty(expT) And Not SameValue(g(1, 8), expT) Then
                    status = "REVIEW": note = "Form = reference model, but not your expected Travel (" & ToText(expT) & "). Check the expectation."
                ElseIf Not IsEmpty(expW) And Not SameValue(g(1, 9), expW) Then
                    status = "REVIEW": note = "Form = reference model, but not your expected Work (" & ToText(expW) & "). Check the expectation."
                Else
                    status = "PASS": note = IIf(IsEmpty(expT) And IsEmpty(expW), "Matches the reference model.", "Matches expected and reference model.")
                End If
            End If
            ws.Cells(r, 15).Value = status
            ws.Cells(r, 15).Interior.Color = StatusColor(IIf(status = "REVIEW" Or status = "INVALID", "WARN", status))
            ws.Cells(r, 16).Value = note
            LogResult id, "Scenarios", IIf(Len(desc) > 0, desc, "Scenario row " & r), _
                      IIf(status = "PASS", "PASS", IIf(status = "FAIL", "FAIL", "WARN")), _
                      FmtPair(e.Travel, e.Work), gotText, note
        End If
    Next r
    If nRun = 0 Then LogResult "U00", "Scenarios", "Rows on " & OUT_SCENARIOS, "SKIP", , , "No scenario rows."
    ClearInputs
    Exit Sub
Boom:
    LogCrash "U99", "Scenarios"
End Sub

'==============================================================================
'  V - POP-UPS (input messages), ERROR ALERTS AND THE MONTH ENTRY IN K5
'  Validation.Value asks Excel whether the value now in a cell passes that
'  cell's rule - the same answer the engineer gets when typing it.
'==============================================================================
Private Sub TestPopups()
    Dim cat As String, sections As Variant, i As Long, c As Range, bad As String, info As String
    Dim oldK5 As Variant, cases As Variant, wrong As String, loose As String, okNow As Variant
    Dim j6 As Variant, lf As String
    cat = "Pop-ups and validation"
    On Error GoTo Boom
    oldK5 = mForm.Range(MONTH_CELL).Value
    mPopupCount = 0
    ReDim mPopups(1 To 30, 1 To 6)

    ' V01 every section the engineer fills in shows a pop-up
    sections = Array("B5", "Name", "I5", "Employee ID", MONTH_CELL, "Month", "A11", "Date (row 11)", "A70", "Date (row 70)", _
                     "D11", "From", "E11", "Until / From (2nd)", "F11", "Until / From (3rd)", "G11", "Until", _
                     "D70", "From (row 70)", "G70", "Until (row 70)", "C77", "Submitted by")
    bad = ""
    For i = 0 To UBound(sections) Step 2
        Set c = mForm.Range(sections(i))
        If VTrue(VProp(c, "ShowInput")) And Len(NzStr(VProp(c, "InputMessage"))) > 0 Then
            AddPopup CStr(sections(i + 1)), c
        Else
            bad = bad & " " & sections(i + 1) & " (" & sections(i) & ");"
        End If
    Next i
    LogResult "V01", cat, "Input pop-up on every section the engineer fills in", IIf(Len(bad) = 0, "PASS", "WARN"), _
              "name, ID, month, dates, times, submitted by", IIf(Len(bad) = 0, (UBound(sections) + 1) \ 2 & " cells checked", "no pop-up:" & bad)

    ' V02 sections still without a pop-up
    bad = "": info = ""
    For i = 10 To mLastInput
        Set c = mForm.Cells(FIRST_ROW, i)
        If Not (VTrue(VProp(c, "ShowInput")) And Len(NzStr(VProp(c, "InputMessage"))) > 0) Then bad = bad & " " & c.Address(False, False)
    Next i
    If mColLocal > 0 Then
        If Not VTrue(VProp(mForm.Cells(FIRST_ROW, mColLocal), "ShowError")) Then
            info = "LOCAL / OVERSEAS has its error alert switched off, so anything typed there is accepted. "
        End If
    End If
    LogResult "V02", cat, "Project ID / Vessel (and LOCAL/OVERSEAS if present) have pop-ups", _
              IIf(Len(info) > 0, "WARN", IIf(Len(bad) = 0, "PASS", "INFO")), "pop-ups on " & ColRange(10, mLastInput), _
              IIf(Len(bad) = 0, "all present", "no pop-up on" & bad), info & IIf(Len(bad) = 0, "", "Optional: a pop-up such as 'Project ID as on the job sheet, e.g. 100084981.002'.")

    ' V03 the cells with a rule stop wrong entries and say why
    bad = ""
    For Each c In mForm.Range("A11,D11,E11,F11,G11," & MONTH_CELL).Cells
        If Not VTrue(VProp(c, "ShowError")) Then
            bad = bad & " " & c.Address(False, False) & "(alert off)"
        ElseIf NzStr(VProp(c, "AlertStyle")) <> CStr(xlValidAlertStop) Then
            bad = bad & " " & c.Address(False, False) & "(warning only, can be overridden)"
        ElseIf Len(NzStr(VProp(c, "ErrorMessage"))) = 0 Then
            bad = bad & " " & c.Address(False, False) & "(no message)"
        End If
    Next c
    LogResult "V03", cat, "Date, time and month rules block wrong entries with an explanation", IIf(Len(bad) = 0, "PASS", "WARN"), _
              "Stop alert + message", IIf(Len(bad) = 0, "6 cells OK", bad)

    ' V04 month entry accepts only FULL MONTH NAME + 4-digit year in capitals
    cases = Array("OCTOBER 2026", True, "MAY 2026", True, "JANUARY 2027", True, "FEBRUARY 2028", True, _
                  "October 2026", False, "october 2026", False, "Oct 2026", False, "OCT 2026", False, "OCTOBER 26", False, _
                  "2026 OCTOBER", False, "OCTOBER  2026", False, " OCTOBER 2026", False, "OCTOBER 2026 ", False, _
                  "SEPT 2026", False, "OCTOBER-2026", False, "10/2026", False, _
                  "1/10/2026", False, "01-Oct-2026", False, "1 OCTOBER 2026", False, "2026-10-01", False, "Oct-26", False)
    wrong = ""
    For i = 0 To UBound(cases) Step 2
        mForm.Range(MONTH_CELL).Value = cases(i)
        okNow = ValidOK(mForm.Range(MONTH_CELL))
        If IsNull(okNow) Then
            wrong = wrong & " '" & cases(i) & "'=error;"
        ElseIf okNow <> cases(i + 1) Then
            wrong = wrong & " '" & cases(i) & "' " & IIf(okNow, "accepted", "rejected") & ";"
        End If
    Next i
    LogResult "V04", cat, "Month entry (" & MONTH_CELL & ") accepts 'OCTOBER 2026' style only", IIf(Len(wrong) = 0, "PASS", "FAIL"), _
              "4 accepted, 17 rejected (incl. 5 date-style entries)", IIf(Len(wrong) = 0, "4 accepted, 17 rejected", "wrong:" & wrong)

    ' V05 the year must be four plain digits (any 4-digit year is allowed)
    loose = ""
    For Each okNow In Array("OCTOBER -202", "OCTOBER 2.26", "OCTOBER 1E03", "OCTOBER +202", "OCTOBER 2O26", "OCTOBER 20 6")
        mForm.Range(MONTH_CELL).Value = okNow
        If VTrue(ValidOK(mForm.Range(MONTH_CELL))) Then loose = loose & " '" & okNow & "'"
    Next okNow
    LogResult "V05", cat, "Month entry rejects years that are not four plain digits", IIf(Len(loose) = 0, "PASS", "WARN"), "rejected", _
              IIf(Len(loose) = 0, "rejected", "accepted:" & loose), _
              IIf(Len(loose) = 0, "", "The rule reads the year as a number instead of checking four digits. " & _
                  "TEXT(--RIGHT(K5,4),""0000"")=RIGHT(K5,4) only lets plain digits through.")

    ' V06 where the date rule gets the claim month from: straight from K5, or via engine J6
    mForm.Range(MONTH_CELL).Value = "OCTOBER 2026"
    Recalc
    lf = NzStr(VProp(mForm.Range("A11"), "Formula1"))
    If InStr(1, Replace(lf, "$", ""), ENGINE_MONTH_CELL, vbTextCompare) = 0 Then
        LogResult "V06", cat, "Date rule reads the claim month", "PASS", MONTH_CELL & " or " & ENGINE_MONTH_CELL, _
                  "directly from " & MONTH_CELL, "The rule on A11:A70 does not use '" & SH_ENGINE & "'!" & ENGINE_MONTH_CELL & _
                  ", so that cell can stay empty. V07 checks the rule's answers."
    Else
        j6 = mEng.Range(ENGINE_MONTH_CELL).Value
        If IsError(j6) Or IsEmpty(j6) Or Not IsNumeric(j6) Or VarType(j6) = vbString Then
            LogResult "V06", cat, "Date rule reads the claim month", "FAIL", _
                      "1-Oct-2026 in " & ENGINE_MONTH_CELL & " for 'OCTOBER 2026'", ToText(j6) & IIf(mEng.Range(ENGINE_MONTH_CELL).HasFormula, "", " (no formula)"), _
                      "The date rule on A11:A70 reads '" & SH_ENGINE & "'!" & ENGINE_MONTH_CELL & _
                      ", which is empty, so Excel rejects EVERY date an engineer types (see V07). Either fill " & ENGINE_MONTH_CELL & _
                      " from " & MONTH_CELL & ", or change the A11:A70 rule to read " & MONTH_CELL & " directly (see docs/FINDINGS.md)."
        ElseIf Year(CDate(j6)) = 2026 And Month(CDate(j6)) = 10 Then
            LogResult "V06", cat, "Date rule reads the claim month", "PASS", "Oct 2026 in " & ENGINE_MONTH_CELL, Format$(CDate(j6), "d-mmm-yyyy")
        Else
            LogResult "V06", cat, "Date rule reads the claim month", "FAIL", "Oct 2026 in " & ENGINE_MONTH_CELL, Format$(CDate(j6), "d-mmm-yyyy")
        End If
    End If

    ' V07 date rule: only the claim month, plus the last day of the month before (overnight jobs)
    cases = Array("OCTOBER 2026", DateSerial(2026, 10, 1), True, "OCTOBER 2026", DateSerial(2026, 10, 15), True, _
                  "OCTOBER 2026", DateSerial(2026, 10, 31), True, "OCTOBER 2026", DateSerial(2026, 9, 30), True, _
                  "OCTOBER 2026", DateSerial(2026, 9, 29), False, "OCTOBER 2026", DateSerial(2026, 11, 1), DATE_ALLOWS_NEXT_FIRST, _
                  "OCTOBER 2026", DateSerial(2026, 11, 2), False, "DECEMBER 2026", DateSerial(2027, 1, 1), DATE_ALLOWS_NEXT_FIRST, _
                  "OCTOBER 2026", DateSerial(2025, 10, 15), False, _
                  "JANUARY 2027", DateSerial(2026, 12, 31), True, "JANUARY 2027", DateSerial(2027, 1, 1), True, _
                  "JANUARY 2027", DateSerial(2026, 12, 30), False, _
                  "MARCH 2028", DateSerial(2028, 2, 29), True, "MARCH 2027", DateSerial(2027, 2, 28), True, _
                  "MARCH 2027", DateSerial(2027, 2, 27), False)
    wrong = ""
    ClearInputs
    For i = 0 To UBound(cases) Step 3
        mForm.Range(MONTH_CELL).Value = cases(i)
        mForm.Range("A11").Value = cases(i + 1)
        mForm.Range("A70").Value = cases(i + 1)
        Recalc
        okNow = ValidOK(mForm.Range("A11"))
        If IsNull(okNow) Then
            wrong = wrong & " " & cases(i) & "/" & Format$(cases(i + 1), "d-mmm-yy") & "=error;"
        ElseIf okNow <> cases(i + 2) Or VTrue(ValidOK(mForm.Range("A70"))) <> okNow Then
            wrong = wrong & " " & cases(i) & "/" & Format$(cases(i + 1), "d-mmm-yy") & " " & IIf(okNow, "accepted", "rejected") & ";"
        End If
    Next i
    LogResult "V07", cat, "Date column accepts the claim month, the day before it" & IIf(DATE_ALLOWS_NEXT_FIRST, " and the 1st of the next month", "") & ", nothing else", _
              IIf(Len(wrong) = 0, "PASS", "FAIL"), "15 cases (incl. Dec->Jan and leap year)", IIf(Len(wrong) = 0, "all 15 correct", "wrong:" & wrong)

    ' V08 time rules (From, Until/From, Until/From, Until)
    wrong = ""
    ClearInputs
    TimeRule wrong, "D", Array(Empty, Empty, Empty, Empty), "08:00", True
    TimeRule wrong, "D", Array(Empty, Empty, Empty, Empty), "00:00", True
    TimeRule wrong, "D", Array(Empty, Empty, Empty, Empty), 25# / 24#, False
    TimeRule wrong, "D", Array(Empty, Empty, Empty, Empty), -0.1, False
    TimeRule wrong, "D", Array(Empty, Empty, Empty, Empty), "0800", False
    TimeRule wrong, "E", Array("09:00", Empty, Empty, Empty), "10:00", True
    TimeRule wrong, "E", Array("09:00", Empty, Empty, Empty), "08:00", False
    TimeRule wrong, "E", Array("21:00", Empty, Empty, Empty), "00:00", True
    TimeRule wrong, "E", Array(Empty, Empty, Empty, Empty), "05:00", True
    TimeRule wrong, "F", Array("09:00", "10:00", Empty, Empty), "09:30", False
    TimeRule wrong, "F", Array("09:00", "10:00", Empty, Empty), "11:00", True
    TimeRule wrong, "F", Array("20:00", "22:00", Empty, Empty), "00:00", True
    TimeRule wrong, "G", Array("09:00", "10:00", "11:00", Empty), "10:30", False
    TimeRule wrong, "G", Array("09:00", "10:00", "11:00", Empty), "12:00", True
    TimeRule wrong, "G", Array("16:00", "20:30", "22:30", Empty), "00:00", True
    LogResult "V08", cat, "Time columns accept real times in order, 00:00 as midnight, and reject the rest", _
              IIf(Len(wrong) = 0, "PASS", "FAIL"), "15 cases", IIf(Len(wrong) = 0, "all 15 correct", "wrong:" & wrong)
    LogResult "V09", cat, "Times after midnight other than 00:00 (e.g. 21:00 -> 02:00) are blocked when typed", "INFO", , _
              "E=02:00 after D=21:00 rejected", "The rule only lets 00:00 go 'backwards'. An engineer working 21:00-02:00 " & _
              "must type 00:00 and carry on in a new row - the pop-up text says so. The formulas themselves would handle 02:00."

    ClearInputs
    mForm.Range(MONTH_CELL).Value = oldK5
    Recalc
    Exit Sub
Boom:
    LogCrash "V99", cat
    On Error Resume Next
    ClearInputs
    mForm.Range(MONTH_CELL).Value = oldK5
End Sub

' Fills D:G of row 11 with the "before" values, puts testValue in column col and checks the rule's answer.
Private Sub TimeRule(ByRef wrong As String, ByVal col As String, ByVal before As Variant, ByVal testValue As Variant, ByVal expectOk As Boolean)
    Dim i As Long, okNow As Variant, shown As String
    For i = 0 To 3
        If IsEmpty(before(i)) Then mForm.Cells(FIRST_ROW, 4 + i).Value = Empty Else mForm.Cells(FIRST_ROW, 4 + i).Value = HM(CStr(before(i)))
    Next i
    If VarType(testValue) = vbString Then
        If InStr(testValue, ":") > 0 Then
            mForm.Range(col & FIRST_ROW).Value = HM(CStr(testValue))
        Else
            mForm.Range(col & FIRST_ROW).Value = "'" & testValue        ' typed without a colon, e.g. 0800
        End If
        shown = CStr(testValue)
    Else
        mForm.Range(col & FIRST_ROW).Value = testValue
        shown = Format$(testValue, "0.000")
    End If
    okNow = ValidOK(mForm.Range(col & FIRST_ROW))
    If IsNull(okNow) Then
        wrong = wrong & " " & col & "=" & shown & " error;"
    ElseIf okNow <> expectOk Then
        wrong = wrong & " " & col & "=" & shown & IIf(okNow, " accepted;", " rejected;")
    End If
    mForm.Range(mForm.Cells(FIRST_ROW, 4), mForm.Cells(FIRST_ROW, 7)).ClearContents
End Sub

'==============================================================================
'  Y - YEAR ROLLOVER (simulated on the copy; the query itself is not touched)
'  The holiday rows of the copy's table are replaced by next year's and the
'  year after's fixed-date holidays - what a refresh in that year would load -
'  then put back exactly as they were.
'==============================================================================
Private Sub TestYearRollover()
    Dim cat As String, lo As ListObject, orig As Variant, origSig As String, baseYear As Long, y As Long, idx As Long
    Dim oldK5 As Variant, d As Date, g As Variant, n As Long, i As Long, r As Long, k As Long, bad As String
    Dim t(0 To 3) As Long, e As OTResult, v As Variant, dates() As Double, expPH() As String
    cat = "Year rollover"
    On Error GoTo Boom
    Set lo = GetHolidayTable()
    If lo Is Nothing Then
        LogResult "Y00", cat, "Year rollover", "SKIP", , , "No holiday table."
        Exit Sub
    End If
    If lo.DataBodyRange Is Nothing Then
        LogResult "Y00", cat, "Year rollover", "SKIP", , , "Holiday table is empty."
        Exit Sub
    End If
    orig = lo.DataBodyRange.Value
    origSig = TableSignature()
    baseYear = mHolYear
    oldK5 = mForm.Range(MONTH_CELL).Value

    ' Y01 the first days of a new year, before the query has loaded it
    y = baseYear + 1
    d = DateSerial(y, 1, 1)
    ClearInputs
    SetClaimMonth y, 1
    PutRow FIRST_ROW, d, "08:00", "09:00", "17:00", "18:00"
    Recalc
    g = mForm.Range("C11").Value
    LogResult "Y01", cat, "New Year's Day " & y & " while the table still holds " & baseYear, _
              IIf(ToText(g) = "Y", "PASS", "WARN"), "Y", ToText(g), _
              "Until the query refreshes in " & y & " - and MOM has published " & y & " - new-year holidays get normal rates. " & _
              "In early January the query stops with 'No public holiday records were found' and the old year stays loaded."

    t(0) = 480: t(1) = 540: t(2) = 1020: t(3) = 1080
    For y = baseYear + 1 To baseYear + 2
        idx = y - baseYear
        Progress "year rollover " & y
        WriteSyntheticHolidays lo, y
        Recalc
        LoadHolidays

        ' Y?1 each holiday of the new year is Y with full-day hours, the day after is N
        ClearInputs
        ReDim dates(1 To NROWS)
        ReDim expPH(1 To NROWS)
        k = 0
        For i = 1 To mHolCount
            If i = 1 Or mHol(i) <> mHol(MaxL(i - 1, 1)) Then          ' skip the padding repeats
                If k < NROWS - 1 Then
                    k = k + 1: dates(k) = mHol(i): expPH(k) = "Y"
                    If Not IsHoliday(mHol(i) + 1) Then k = k + 1: dates(k) = mHol(i) + 1: expPH(k) = "N"
                End If
            End If
        Next i
        For r = 1 To k
            PutRow FIRST_ROW + r - 1, CDate(dates(r)), "08:00", "09:00", "17:00", "18:00"
        Next r
        SetClaimMonth y, 1
        Recalc
        v = mForm.Range(mForm.Cells(FIRST_ROW, 1), mForm.Cells(LAST_ROW, 12)).Value
        bad = ""
        For r = 1 To k
            e = Oracle(dates(r), t, mQ1, mQ2)
            If ToText(v(r, 3)) <> expPH(r) Or Not RowMatches(e, v(r, 2), v(r, 3), v(r, 8), v(r, 9), dates(r)) Then
                bad = bad & " " & Format$(dates(r), "d-mmm") & "=" & ToText(v(r, 3)) & " " & FmtPair(v(r, 8), v(r, 9)) & ";"
            End If
        Next r
        LogResult "Y" & idx & "1", cat, "After a refresh in " & y & ": its holidays (and in-lieu Mondays) get holiday rates", _
                  IIf(Len(bad) = 0, "PASS", "FAIL"), k & " dates correct", IIf(Len(bad) = 0, k & " dates correct", "wrong:" & bad), _
                  "Simulated with the fixed-date holidays only (1 Jan, 1 May, 9 Aug, 25 Dec + Sunday in-lieu days); " & _
                  "moving holidays depend on what MOM publishes."

        ' Y?2 random full forms in the new year
        TestFuzz ROLLOVER_FUZZ, "Y" & idx & "2", "Year " & y & " random fills"

        ' Y?3 last December's claim, finished after the new year's table is loaded
        d = DateSerial(y - 1, 12, 25)
        ClearInputs
        SetClaimMonth y - 1, 12
        PutRow FIRST_ROW, d, "08:00", "09:00", "17:00", "18:00"
        Recalc
        g = mForm.Range("C11:I11").Value
        LogResult "Y" & idx & "3", cat, "December " & (y - 1) & " claim submitted after the " & y & " holidays load (25 Dec)", _
                  IIf(ToText(g(1, 1)) = "Y", "PASS", "WARN"), "Y", ToText(g(1, 1)) & ", " & FmtPair(g(1, 6), g(1, 7)), _
                  "Last year's holidays drop out on 1 January. Submit December claims before the year ends, " & _
                  "or keep a copy of last year's holiday list."
    Next y

    ' put the real holiday rows back exactly
    lo.DataBodyRange.Value = orig
    ClearInputs
    mForm.Range(MONTH_CELL).Value = oldK5
    Recalc
    LoadHolidays
    LogResult "Y99", cat, "Holiday table restored after the simulation", IIf(TableSignature() = origSig, "PASS", "FAIL"), _
              "identical", IIf(TableSignature() = origSig, "identical", "different")
    Exit Sub
Boom:
    LogCrash "Y98", cat
    On Error Resume Next
    If IsArray(orig) Then lo.DataBodyRange.Value = orig
    ClearInputs
    mForm.Range(MONTH_CELL).Value = oldK5
    Recalc
    LoadHolidays
End Sub

' Next-year holidays that never move, plus the Monday in lieu when one falls on a Sunday.
' The table keeps its size: spare rows repeat the last holiday (harmless for COUNTIF).
Private Sub WriteSyntheticHolidays(ByVal lo As ListObject, ByVal y As Long)
    Dim fixedList As Variant, i As Long, k As Long, cnt As Long, n As Long, d As Date
    Dim ds(1 To 8) As Date, nm(1 To 8) As String, tp(1 To 8) As String, rows() As Variant
    fixedList = Array(1, 1, "New Year's Day", 5, 1, "Labour Day", 8, 9, "National Day", 12, 25, "Christmas Day")
    For i = 0 To UBound(fixedList) Step 3
        d = DateSerial(y, fixedList(i), fixedList(i + 1))
        cnt = cnt + 1: ds(cnt) = d: nm(cnt) = fixedList(i + 2): tp(cnt) = "Public Holiday"
        If Weekday(d, vbMonday) = 7 Then
            cnt = cnt + 1: ds(cnt) = d + 1: nm(cnt) = fixedList(i + 2) & " (In lieu)": tp(cnt) = "Public Holiday in lieu"
        End If
    Next i
    n = lo.DataBodyRange.Rows.Count
    ReDim rows(1 To n, 1 To lo.ListColumns.Count)
    For i = 1 To n
        k = MinL(i, cnt)
        rows(i, 1) = ds(k)
        rows(i, 2) = EnglishDay(CDbl(ds(k)))
        rows(i, 3) = nm(k)
        rows(i, 4) = tp(k)
    Next i
    lo.DataBodyRange.Value = rows
End Sub

Private Sub AddPopup(ByVal section As String, ByVal c As Range)
    If mPopupCount >= UBound(mPopups, 1) Then Exit Sub
    mPopupCount = mPopupCount + 1
    mPopups(mPopupCount, 1) = section
    mPopups(mPopupCount, 2) = c.Address(False, False)
    mPopups(mPopupCount, 3) = NzStr(VProp(c, "InputTitle"))
    mPopups(mPopupCount, 4) = NzStr(VProp(c, "InputMessage"))
    mPopups(mPopupCount, 5) = NzStr(VProp(c, "ErrorTitle"))
    mPopups(mPopupCount, 6) = NzStr(VProp(c, "ErrorMessage"))
End Sub

' A property of a cell's data validation, or Null when the cell has no validation.
Private Function VProp(ByVal c As Range, ByVal prop As String) As Variant
    VProp = Null
    On Error Resume Next
    VProp = CallByName(c.Validation, prop, VbGet)
End Function

Private Function VTrue(ByVal v As Variant) As Boolean
    If VarType(v) = vbBoolean Then VTrue = v
End Function

' True / False = does the cell's current value pass its rule; Null = Excel could not tell.
Private Function ValidOK(ByVal c As Range) As Variant
    ValidOK = Null
    On Error Resume Next
    ValidOK = c.Validation.Value
End Function

' "OCTOBER 2026" -> 1-Oct-2026 as a serial, anything else -> 0
Private Function ParseMonthText(ByVal s As String) As Double
    Dim parts() As String, m As Long
    parts = Split(s, " ")
    If UBound(parts) <> 1 Then Exit Function
    If Len(parts(1)) <> 4 Or Not parts(1) Like "####" Then Exit Function
    For m = 1 To 12
        If parts(0) = EnglishMonth(m) Then
            ParseMonthText = DateSerial(Val(parts(1)), m, 1)
            Exit Function
        End If
    Next m
End Function

' Same rule as the date validation: inside the claim month, or its previous day.
Private Function InClaimMonth(ByVal ds As Double, ByVal monthStart As Double) As Boolean
    InClaimMonth = (Year(CDate(ds)) = Year(CDate(monthStart)) And Month(CDate(ds)) = Month(CDate(monthStart))) Or _
                   (Int(ds) = monthStart - 1) Or _
                   (DATE_ALLOWS_NEXT_FIRST And Int(ds) = DateSerial(Year(CDate(monthStart)), Month(CDate(monthStart)) + 1, 1))
End Function

Private Sub SetClaimMonth(ByVal y As Long, ByVal m As Long)
    mForm.Range(MONTH_CELL).Value = EnglishMonth(m) & " " & y
End Sub

Private Function EnglishMonth(ByVal m As Long) As String
    EnglishMonth = Choose(m, "JANUARY", "FEBRUARY", "MARCH", "APRIL", "MAY", "JUNE", "JULY", "AUGUST", _
                          "SEPTEMBER", "OCTOBER", "NOVEMBER", "DECEMBER")
End Function

Private Function EnglishDay(ByVal ds As Double) As String
    EnglishDay = Choose(Weekday(CDate(ds), vbMonday), "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday")
End Function

' Start and end of a row on one timeline (minutes since 30-Dec-1899), using the form's midnight roll-over.
Private Function RowSpan(ByVal ds As Double, ByRef t() As Long, ByRef startMin As Double, ByRef endMin As Double) As Boolean
    Dim i As Long, prev As Long, u As Long, first As Long, last As Long, n As Long
    first = -1
    prev = 0
    If t(0) >= 0 Then prev = t(0)
    For i = 0 To 3
        If t(i) >= 0 Then
            u = t(i)
            Do While u < prev
                u = u + 1440
            Loop
            prev = u
            If first < 0 Then first = u
            last = u
            n = n + 1
        End If
    Next i
    If n < 2 Or last <= first Then Exit Function
    startMin = Int(ds) * 1440# + first
    endMin = Int(ds) * 1440# + last
    RowSpan = True
End Function

' "rows 11 & 12 (19-Sep 09:00-10:00)" for every pair of rows whose time spans overlap; touching ends are fine.
Private Function OverlapList(ByRef spanS() As Double, ByRef spanE() As Double, ByRef hasSpan() As Boolean) As String
    Dim i As Long, j As Long, a As Double, b As Double, n As Long
    For i = LBound(spanS) To UBound(spanS) - 1
        If hasSpan(i) Then
            For j = i + 1 To UBound(spanS)
                If hasSpan(j) Then
                    a = spanS(i): If spanS(j) > a Then a = spanS(j)
                    b = spanE(i): If spanE(j) < b Then b = spanE(j)
                    If a < b Then
                        n = n + 1
                        If n <= 15 Then OverlapList = OverlapList & " rows " & (i + FIRST_ROW - 1) & " & " & (j + FIRST_ROW - 1) & _
                            " (" & Format$(CDate(Int(a / 1440#)), "d-mmm") & " " & HHMM(CLng(a - Int(a / 1440#) * 1440#)) & "-" & _
                            HHMM(CLng(b - Int(a / 1440#) * 1440#) Mod 1440) & ");"
                    End If
                End If
            Next j
        End If
    Next i
    If n > 15 Then OverlapList = OverlapList & " ... " & n & " pairs in all"
End Function

' Overlap check on two rows currently on the form (used by L08 to prove B11 would catch it).
Private Function OverlapOfFormRows(ByVal r1 As Long, ByVal r2 As Long) As String
    Dim s(1 To 2) As Double, e(1 To 2) As Double, h(1 To 2) As Boolean, t(0 To 3) As Long, k As Long, c As Long, r As Long, v As Variant
    For k = 1 To 2
        r = IIf(k = 1, r1, r2)
        For c = 0 To 3
            v = mForm.Cells(r, 4 + c).Value2
            If VarType(v) = vbDouble Then t(c) = Int(v * 1440 + 0.5) Mod 1440 Else t(c) = -1
        Next c
        v = mForm.Cells(r, 1).Value2
        If VarType(v) = vbDouble Then h(k) = RowSpan(CDbl(v), t, s(k), e(k))
    Next k
    OverlapOfFormRows = OverlapList(s, e, h)
End Function

' True when the From cell of a form row shows the overlap highlight (light red fill).
Private Function Highlighted(ByVal r As Long) As Boolean
    On Error Resume Next
    Highlighted = (mForm.Cells(r, 4).DisplayFormat.Interior.Color = RGB(255, 199, 206))
End Function

' Finds Project ID / Vessel / LOCAL-OVERSEAS columns from the headings in row 9 (the column was removed in rev2).
Private Sub DetectLayout()
    Dim c As Long, h As String
    mColLocal = 0: mColProject = 0: mColVessel = 0: mColActivity = 0
    For c = 10 To 13
        h = UCase$(Squash(CStr(mForm.Cells(9, c).Value)))
        If h = "LOCAL / OVERSEAS" Then mColLocal = c
        If h = "PROJECT ID" Then mColProject = c
        If h = "VESSEL NAME" Then mColVessel = c
        If h = "ACTIVITY NUMBER" Then mColActivity = c
    Next c
    If mColProject = 0 Then mColProject = IIf(mColLocal > 0, mColLocal + 1, 10)
    If mColVessel = 0 Then mColVessel = IIf(mColActivity > 0, mColActivity, mColProject) + 1
    mLastInput = mColVessel
    If mColLocal > mLastInput Then mLastInput = mColLocal
    If mColActivity > mLastInput Then mLastInput = mColActivity
End Sub

Private Sub WriteJobInfo(ByRef aJ() As Variant)
    Dim i As Long, colArr(1 To NROWS, 1 To 1) As Variant, k As Long, cols As Variant
    cols = Array(mColLocal, mColProject, mColVessel, mColActivity)
    For k = 0 To 3
        If cols(k) > 0 Then
            For i = 1 To NROWS
                colArr(i, 1) = aJ(i, k + 1)
            Next i
            mForm.Range(mForm.Cells(FIRST_ROW, cols(k)), mForm.Cells(LAST_ROW, cols(k))).Value = colArr
        End If
    Next k
End Sub

Private Function ColLetter(ByVal c As Long) As String
    ColLetter = Split(mForm.Cells(1, c).Address(True, False), "$")(0)
End Function

Private Function ColRange(ByVal c1 As Long, ByVal c2 As Long) As String
    ColRange = ColLetter(c1) & FIRST_ROW & ":" & ColLetter(c2) & LAST_ROW
End Function

Private Function OtherSheetNames() As String
    Dim ws As Worksheet
    For Each ws In mWb.Worksheets
        If ws.Name <> SH_FORM And ws.Name <> SH_ENGINE And ws.Name <> SH_HOL Then OtherSheetNames = OtherSheetNames & "'" & ws.Name & "' "
    Next ws
End Function

Private Function HelpShapeCount() As Long
    Dim ws As Worksheet
    For Each ws In mWb.Worksheets
        If ws.Name <> SH_FORM And ws.Name <> SH_ENGINE And ws.Name <> SH_HOL Then HelpShapeCount = HelpShapeCount + ws.Shapes.Count
    Next ws
End Function

'==============================================================================
'  REFERENCE MODEL - an independent re-implementation of the form's rules
'  (checked against the real formulas on 540 random rows + 21 hand cases)
'==============================================================================
Private Function Oracle(ByVal dSerial As Double, ByRef t() As Long, ByVal q1 As Long, ByVal q2 As Long) As OTResult
    Dim r As OTResult, u(0 To 3) As Long, i As Long, prev As Long, special As Boolean
    Dim tr As Long, wk As Long
    If dSerial = 0 Then
        r.DayText = " ": r.PH = "": r.Travel = "": r.Work = ""
        Oracle = r
        Exit Function
    End If
    r.DayText = Format$(CDate(dSerial), "ddd")
    r.PH = IIf(IsHoliday(dSerial), "Y", "N")
    If t(0) < 0 And t(1) < 0 And t(2) < 0 And t(3) < 0 Then
        r.Travel = "": r.Work = ""
        Oracle = r
        Exit Function
    End If
    special = (Weekday(CDate(dSerial), vbMonday) >= 6) Or (r.PH = "Y")

    ' put the times on one timeline: each later time is at or after the one before it
    prev = 0
    If t(0) >= 0 Then prev = t(0)
    u(0) = prev
    For i = 1 To 3
        If t(i) >= 0 Then
            u(i) = t(i)
            Do While u(i) < prev
                u(i) = u(i) + 1440
            Loop
            prev = u(i)
        End If
    Next i

    ' travel: From -> 2nd, 3rd -> Until, or From -> Until when the middle is empty
    If t(0) >= 0 And t(1) >= 0 Then tr = tr + SegmentMinutes(u(0), u(1), special, q1, q2)
    If t(2) >= 0 And t(3) >= 0 Then tr = tr + SegmentMinutes(u(2), u(3), special, q1, q2)
    If t(0) >= 0 And t(3) >= 0 And t(1) < 0 And t(2) < 0 Then tr = tr + SegmentMinutes(u(0), u(3), special, q1, q2)
    ' work: 2nd -> 3rd, else 2nd -> Until, else From -> 3rd
    If t(1) >= 0 Then
        If t(2) >= 0 Then
            wk = SegmentMinutes(u(1), u(2), special, q1, q2)
        ElseIf t(3) >= 0 Then
            wk = SegmentMinutes(u(1), u(3), special, q1, q2)
        End If
    ElseIf t(0) >= 0 And t(2) >= 0 Then
        wk = SegmentMinutes(u(0), u(2), special, q1, q2)
    End If
    r.Travel = Application.WorksheetFunction.Round(tr / 60#, 2)
    r.Work = Application.WorksheetFunction.Round(wk / 60#, 2)
    Oracle = r
End Function

' Minutes of overtime in [s, e]. Weekends and public holidays pay everything;
' weekdays pay the part outside office hours of the day the segment starts on.
Private Function SegmentMinutes(ByVal s As Long, ByVal e As Long, ByVal special As Boolean, ByVal q1 As Long, ByVal q2 As Long) As Long
    Dim off As Long, ov As Long
    If e <= s Then Exit Function
    If special Then
        SegmentMinutes = e - s
        Exit Function
    End If
    If s >= 1440 Then off = 1440
    ov = MinL(e, q2 + off) - MaxL(s, q1 + off)
    If ov < 0 Then ov = 0
    SegmentMinutes = (e - s) - ov
End Function

Private Function RowMatches(ByRef e As OTResult, ByVal gB As Variant, ByVal gC As Variant, ByVal gH As Variant, _
                            ByVal gI As Variant, ByVal dSerial As Double) As Boolean
    If IsError(gB) Or IsError(gC) Or IsError(gH) Or IsError(gI) Then Exit Function
    If dSerial = 0 Then
        If Trim$(CStr(gB)) <> "" Then Exit Function
    ElseIf CStr(gB) <> e.DayText Then
        If CStr(gB) <> Application.WorksheetFunction.Text(dSerial, "ddd") Then Exit Function
    End If
    If CStr(gC) <> e.PH Then Exit Function
    If Not SameValue(gH, e.Travel) Then Exit Function
    If Not SameValue(gI, e.Work) Then Exit Function
    RowMatches = True
End Function

'==============================================================================
'  RANDOM INPUT GENERATION
'==============================================================================
Private Sub SeedRandom()
    Rnd -1
    Randomize RANDOM_SEED
End Sub

Private Function RandomDate() As Double
    Dim r As Double, first As Double
    r = Rnd
    If r < 0.05 Then Exit Function                       ' 5% rows without a date
    If r < 0.3 And mHolCount > 0 Then                    ' 25% public holidays
        RandomDate = mHol(1 + Int(Rnd * mHolCount))
        Exit Function
    End If
    first = DateSerial(mHolYear, 1, 1)
    RandomDate = first + Int(Rnd * (DateSerial(mHolYear, 12, 31) - first + 1))
End Function

' Up to four times on one line, spanning less than 24 hours (one midnight roll-over at most).
Private Sub GenTimes(ByRef t() As Long)
    Static masks As Variant
    Dim m As String, n As Long, i As Long, j As Long, k As Long, start As Long, span As Long, tmp As Long
    Dim pts(1 To 4) As Long, snap As Boolean
    If IsEmpty(masks) Then masks = Array("1111", "1111", "1111", "1100", "0011", "1001", "1101", "1011", _
                                         "0111", "1110", "1000", "0000", "1010", "0101")
    m = masks(Int(Rnd * (UBound(masks) + 1)))
    n = Len(Replace(m, "0", ""))
    snap = (Rnd < 0.5)
    If snap Then start = Int(Rnd * 96) * 15 Else start = Int(Rnd * 1440)
    span = Int(Rnd * 1439)
    For i = 1 To n
        pts(i) = Int(Rnd * (span + 1))
        If snap Then pts(i) = (pts(i) \ 15) * 15
    Next i
    For i = 2 To n
        tmp = pts(i)
        j = i - 1
        Do While j >= 1
            If pts(j) <= tmp Then Exit Do
            pts(j + 1) = pts(j)
            j = j - 1
        Loop
        pts(j + 1) = tmp
    Next i
    For i = 0 To 3
        If Mid$(m, i + 1, 1) = "1" Then
            k = k + 1
            t(i) = (start + pts(k)) Mod 1440
        Else
            t(i) = -1
        End If
    Next i
End Sub

'==============================================================================
'  FORM HELPERS
'==============================================================================
Private Function BindSheets() As Boolean
    Dim nm As Variant, missing As String
    For Each nm In Array(SH_FORM, SH_ENGINE, SH_HOL)
        If SheetByName(mWb, CStr(nm)) Is Nothing Then missing = missing & " '" & nm & "'"
    Next nm
    If Len(missing) = 0 Then
        LogResult "S01", "Structure", "Form, engine and holiday sheets present", "PASS", "3 sheets", "3 sheets", _
                  "Other sheets: " & OtherSheetNames()
    Else
        LogResult "S01", "Structure", "Form, engine and holiday sheets present", "FAIL", "3 sheets", "missing:" & missing
    End If
    Set mForm = SheetByName(mWb, SH_FORM)
    Set mEng = SheetByName(mWb, SH_ENGINE)
    BindSheets = Not (mForm Is Nothing Or mEng Is Nothing)
    If BindSheets Then DetectLayout
End Function

Private Sub ClearInputs()
    mForm.Range(mForm.Cells(FIRST_ROW, 1), mForm.Cells(LAST_ROW, 1)).ClearContents
    mForm.Range(mForm.Cells(FIRST_ROW, 4), mForm.Cells(LAST_ROW, 7)).ClearContents
    mForm.Range(mForm.Cells(FIRST_ROW, 10), mForm.Cells(LAST_ROW, mLastInput)).ClearContents
End Sub

Private Sub PutRow(ByVal r As Long, ByVal dVal As Variant, ByVal sD As String, ByVal sE As String, ByVal sF As String, ByVal sG As String)
    mForm.Cells(r, 1).Value = dVal
    mForm.Cells(r, 4).Value = HM(sD)
    mForm.Cells(r, 5).Value = HM(sE)
    mForm.Cells(r, 6).Value = HM(sF)
    mForm.Cells(r, 7).Value = HM(sG)
End Sub

' "hh:mm" -> fraction of a day, "" -> Empty
Private Function HM(ByVal s As String) As Variant
    Dim p As Long
    s = Trim$(s)
    If Len(s) = 0 Then
        HM = Empty
    Else
        p = InStr(s, ":")
        HM = (Val(Left$(s, p - 1)) * 60# + Val(Mid$(s, p + 1))) / 1440#
    End If
End Function

Private Sub Recalc()
    Application.Calculate
End Sub

' Recalculate every formula of the test copy without touching other open workbooks.
Private Sub ForceRecalc()
    On Error GoTo Fallback
    mEng.UsedRange.Dirty
    mForm.UsedRange.Dirty
    Application.Calculate
    Exit Sub
Fallback:
    Application.CalculateFull
End Sub

Private Function GetHolidayTable() As ListObject
    Dim ws As Worksheet
    Set ws = SheetByName(mWb, SH_HOL)
    If ws Is Nothing Then Exit Function
    On Error Resume Next
    Set GetHolidayTable = ws.ListObjects(TABLE_NAME)
End Function

Private Function LoadHolidays() As Boolean
    Dim lo As ListObject, v As Variant, i As Long, n As Long
    mHolCount = 0
    mHolYear = Year(Date)
    Set lo = GetHolidayTable()
    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    v = As2D(lo.ListColumns(1).DataBodyRange.Value2)
    ReDim mHol(1 To UBound(v, 1))
    For i = 1 To UBound(v, 1)
        If VarType(v(i, 1)) = vbDouble Then
            n = n + 1
            mHol(n) = CLng(Int(v(i, 1)))
        End If
    Next i
    mHolCount = n
    If n > 0 Then mHolYear = Year(CDate(mHol(1)))
    LoadHolidays = (n > 0)
End Function

' Same test the form uses: COUNTIF(Holidays_1[Date], date) or COUNTIF(..., TEXT(date,"yyyy-mm-dd")),
' so a date carrying a time of day still matches its calendar day.
Private Function IsHoliday(ByVal dSerial As Double) As Boolean
    Dim i As Long
    For i = 1 To mHolCount
        If CDbl(mHol(i)) = Int(dSerial) Then
            IsHoliday = True
            Exit Function
        End If
    Next i
End Function

Private Function FirstWeekdayHoliday() As Double
    Dim i As Long
    For i = 1 To mHolCount
        If Weekday(CDate(mHol(i)), vbMonday) <= 5 Then
            FirstWeekdayHoliday = mHol(i)
            Exit Function
        End If
    Next i
End Function

Private Function HolidayRowExists(ByRef v As Variant, ByVal d As Double, ByVal nameText As String, ByVal typeText As String, _
                                  Optional ByVal partialName As Boolean = False) As Boolean
    Dim i As Long
    For i = 1 To UBound(v, 1)
        If v(i, 1) = d Then
            If (Len(typeText) = 0 Or v(i, 4) = typeText) Then
                If partialName Then
                    If InStr(1, v(i, 3), nameText, vbTextCompare) > 0 Then HolidayRowExists = True: Exit Function
                ElseIf StrComp(v(i, 3), nameText, vbTextCompare) = 0 Then
                    HolidayRowExists = True: Exit Function
                End If
            End If
        End If
    Next i
End Function

Private Function TableSignature() As String
    Dim lo As ListObject, v As Variant, i As Long, j As Long, s As String
    Set lo = GetHolidayTable()
    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    v = As2D(lo.DataBodyRange.Value2)
    For i = 1 To UBound(v, 1)
        For j = 1 To UBound(v, 2)
            s = s & CStr(v(i, j)) & "|"
        Next j
        s = s & vbLf
    Next i
    TableSignature = s
End Function

Private Function RefreshHolidayQuery(ByRef errMsg As String) As Boolean
    Dim lo As ListObject
    errMsg = ""
    Set lo = GetHolidayTable()
    If lo Is Nothing Then
        errMsg = "table " & TABLE_NAME & " not found"
        Exit Function
    End If
    On Error Resume Next
    lo.QueryTable.Refresh BackgroundQuery:=False
    If Err.Number <> 0 Then
        errMsg = "QueryTable.Refresh: " & Err.Description
        Err.Clear
        mWb.Connections(CONN_NAME).Refresh
        If Err.Number = 0 Then Application.CalculateUntilAsyncQueriesDone
        If Err.Number <> 0 Then
            errMsg = errMsg & " | Connection.Refresh: " & Err.Description
            Exit Function
        End If
        errMsg = ""
    End If
    RefreshHolidayQuery = True
End Function

' Read only: returns the M code of the query.
Private Function ReadQueryFormula(ByRef found As Boolean) As String
    Dim q As Object
    found = False
    On Error Resume Next
    Set q = mWb.Queries(QUERY_NAME)
    If Err.Number = 0 And Not q Is Nothing Then
        ReadQueryFormula = q.Formula
        found = (Err.Number = 0)
    End If
End Function

' Read only: connection string + command text of the query's connection.
Private Function ReadConnectionText() As String
    Dim cn As WorkbookConnection, s As String, cmd As Variant
    On Error Resume Next
    Set cn = mWb.Connections(CONN_NAME)
    If cn Is Nothing Then Exit Function
    s = CStr(cn.OLEDBConnection.Connection)
    cmd = cn.OLEDBConnection.CommandText
    If IsArray(cmd) Then cmd = Join(cmd, "")
    ReadConnectionText = s & " | " & CStr(cmd)
End Function

Private Function MakeWorkingCopy(ByVal src As String) As String
    Dim dst As String, wb As Workbook, ext As String, folder As String
    ext = Mid$(src, InStrRev(src, "."))
    folder = ThisWorkbook.Path
    If Len(folder) = 0 Or LCase$(Left$(folder, 4)) = "http" Then folder = Environ$("TEMP")
    If Len(folder) = 0 Then folder = CurDir$
    dst = folder & Application.PathSeparator & "~STRESSTEST_COPY_" & Format$(Now, "yyyymmdd_hhnnss") & ext
    Set wb = FindOpenWorkbook(src)
    If Not wb Is Nothing Then
        wb.SaveCopyAs dst
    Else
        On Error Resume Next
        FileCopy src, dst
        If Err.Number <> 0 Then
            Err.Clear
            On Error GoTo 0
            Set wb = Workbooks.Open(Filename:=src, UpdateLinks:=0, ReadOnly:=True, AddToMru:=False)
            wb.SaveCopyAs dst
            wb.Close SaveChanges:=False
        End If
        On Error GoTo 0
    End If
    MakeWorkingCopy = dst
End Function

Private Function FindOpenWorkbook(ByVal fullPath As String) As Workbook
    Dim wb As Workbook
    For Each wb In Application.Workbooks
        If StrComp(wb.FullName, fullPath, vbTextCompare) = 0 Then
            Set FindOpenWorkbook = wb
            Exit Function
        End If
    Next wb
End Function

Private Function SheetByName(ByVal wb As Workbook, ByVal nm As String) As Worksheet
    On Error Resume Next
    Set SheetByName = wb.Worksheets(nm)
End Function

Private Function ShapeCount(ByVal sheetName As String) As Long
    On Error Resume Next
    ShapeCount = mWb.Worksheets(sheetName).Shapes.Count
End Function

Private Function ValidationType(ByVal cell As Range) As Long
    ValidationType = -1
    On Error Resume Next
    ValidationType = cell.Validation.Type
End Function

Private Function ValidationFormula(ByVal cell As Range) As String
    On Error Resume Next
    ValidationFormula = cell.Validation.Formula1
End Function

Private Function TimeCellMinutes(ByVal v As Variant, ByVal fallback As Long) As Long
    If VarType(v) = vbDouble Then
        TimeCellMinutes = Int((v - Int(v)) * 1440 + 0.5)
    Else
        TimeCellMinutes = fallback
    End If
End Function

' Scenario sheet parsing ------------------------------------------------------
Private Function ParseDateCell(ByVal v As Variant, ByRef ok As Boolean) As Double
    Dim s As String
    If IsEmpty(v) Then Exit Function
    If VarType(v) = vbDate Or VarType(v) = vbDouble Then
        ParseDateCell = Int(CDbl(v))
        Exit Function
    End If
    s = Trim$(CStr(v))
    If Len(s) = 0 Then Exit Function
    If s Like "####-##-##" Then
        ParseDateCell = DateSerial(Val(Left$(s, 4)), Val(Mid$(s, 6, 2)), Val(Mid$(s, 9, 2)))
    ElseIf IsDate(s) Then
        ParseDateCell = Int(CDbl(CDate(s)))
    Else
        ok = False
    End If
End Function

' -1 = blank, -2 = not a valid time, otherwise minutes after midnight
Private Function ParseTimeCell(ByVal v As Variant) As Long
    Dim s As String, p As Long, h As Long, m As Long
    ParseTimeCell = -1
    If IsEmpty(v) Then Exit Function
    If VarType(v) = vbDouble Or VarType(v) = vbDate Then
        If CDbl(v) >= 0 And CDbl(v) < 1 Then
            ParseTimeCell = Int(CDbl(v) * 1440 + 0.5) Mod 1440
        Else
            ParseTimeCell = -2
        End If
        Exit Function
    End If
    s = Trim$(CStr(v))
    If Len(s) = 0 Or s = "-" Then Exit Function
    p = InStr(s, ":")
    If p < 2 Or Not IsNumeric(Left$(s, p - 1)) Or Not IsNumeric(Mid$(s, p + 1, 2)) Then
        ParseTimeCell = -2
        Exit Function
    End If
    h = Val(Left$(s, p - 1)): m = Val(Mid$(s, p + 1, 2))
    If h < 0 Or h > 23 Or m < 0 Or m > 59 Then ParseTimeCell = -2 Else ParseTimeCell = h * 60 + m
End Function

Private Function ParseNumberCell(ByVal v As Variant) As Variant
    ParseNumberCell = Empty
    If IsEmpty(v) Then Exit Function
    If IsNumeric(v) And Len(Trim$(CStr(v))) > 0 Then ParseNumberCell = CDbl(v)
End Function

'==============================================================================
'  LOGGING AND OUTPUT
'==============================================================================
Private Sub PrepareOutputSheets()
    Set mRes = GetOutSheet(OUT_RESULTS, True)
    Set mTim = GetOutSheet(OUT_TIMINGS, True)
    Set mMis = GetOutSheet(OUT_MISMATCH, True)
    GetOutSheet OUT_SUMMARY, True
    GetOutSheet OUT_COPILOT, True           ' cleared so no =COPILOT() call runs during the test
    EnsureScenarioSheet
    mPass = 0: mFail = 0: mWarn = 0: mSkip = 0: mInfo = 0

    mRes.Range("A1:H1").Value = Array("ID", "Category", "Check", "Status", "Expected", "Actual", "Details", "Logged at")
    mRes.Columns("A:G").NumberFormat = "@"
    mRes.Range("A1:H1").Font.Bold = True
    mResRow = 1
    mTim.Range("A1:E1").Value = Array("Kind", "Iteration", "Milliseconds", "OK", "Note")
    mTim.Range("A1:E1").Font.Bold = True
    mTimRow = 1
    mMis.Range("A1:P1").Value = Array("Source", "Iteration", "Form row", "Date", "From", "Until / From", "Until / From", "Until", _
                                      "Expected day", "Form day", "Expected PH", "Form PH", "Expected travel", "Form travel", _
                                      "Expected work", "Form work")
    mMis.Columns("A:P").NumberFormat = "@"
    mMis.Range("A1:P1").Font.Bold = True
    mMisRow = 1
End Sub

Private Function GetOutSheet(ByVal nm As String, ByVal clearIt As Boolean) As Worksheet
    Dim ws As Worksheet
    Set ws = SheetByName(ThisWorkbook, nm)
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = nm
    ElseIf clearIt Then
        ws.Cells.Clear
    End If
    Set GetOutSheet = ws
End Function

Private Sub EnsureScenarioSheet()
    Dim ws As Worksheet, created As Boolean
    Set ws = SheetByName(ThisWorkbook, OUT_SCENARIOS)
    If ws Is Nothing Then
        Set ws = GetOutSheet(OUT_SCENARIOS, False)
        created = True
    End If
    ws.Range("A1").Value = "Scenarios - your own or Copilot-generated test rows"
    ws.Range("A1").Font.Bold = True
    ws.Range("A2").Value = "Type or paste rows from row " & SCEN_FIRST_ROW & ". Date as a date or yyyy-mm-dd; times as hh:mm (24 h), " & _
                           "blank when unused. Expected hours are optional. Columns K:P are filled in by the macro."
    ws.Range("A4:P4").Value = Array("ID", "Description", "Date", "From", "Until / From", "Until / From", "Until", _
                                    "Expected travel (optional)", "Expected work (optional)", "Source", _
                                    "Form travel", "Form work", "Reference travel", "Reference work", "Status", "Note")
    ws.Range("A4:P4").Font.Bold = True
    If created Then
        AddScenario ws, 5, "S001", "Sample: weekday overnight travel", DateSerial(2026, 9, 18), "21:00", "00:00", "", "", 3, 0
        AddScenario ws, 6, "S002", "Saturday travel-work-travel", DateSerial(2026, 9, 19), "00:00", "01:00", "06:30", "09:30", 4, 5.5
        AddScenario ws, 7, "S003", "Weekday travel before office hours", DateSerial(2026, 9, 22), "06:30", "08:30", "11:30", "12:30", 1.5, 0
        AddScenario ws, 8, "S004", "Weekday work across midnight", DateSerial(2026, 9, 22), "20:00", "22:00", "02:00", "03:00", 3, 4
        AddScenario ws, 9, "S005", "Two minutes across 08:00", DateSerial(2026, 9, 22), "07:59", "08:01", "", "", 0.02, 0
        AddScenario ws, 10, "S006", "Saturday almost 24 h of travel", DateSerial(2026, 9, 19), "00:00", "", "", "23:59", 23.98, 0
        ws.Columns("C").NumberFormat = "yyyy-mm-dd"
        ws.Columns("D:G").NumberFormat = "hh:mm"
        ws.Columns("A:P").AutoFit
        ws.Columns("B").ColumnWidth = 40
    End If
End Sub

Private Sub AddScenario(ByVal ws As Worksheet, ByVal r As Long, ByVal id As String, ByVal desc As String, ByVal d As Date, _
                        ByVal sD As String, ByVal sE As String, ByVal sF As String, ByVal sG As String, ByVal expT As Double, ByVal expW As Double)
    ws.Cells(r, 1).Value = id
    ws.Cells(r, 2).Value = desc
    ws.Cells(r, 3).Value = d
    ws.Cells(r, 4).Value = HM(sD)
    ws.Cells(r, 5).Value = HM(sE)
    ws.Cells(r, 6).Value = HM(sF)
    ws.Cells(r, 7).Value = HM(sG)
    ws.Cells(r, 8).Value = expT
    ws.Cells(r, 9).Value = expW
    ws.Cells(r, 10).Value = "Example"
End Sub

Private Sub LogResult(ByVal id As String, ByVal cat As String, ByVal checkName As String, ByVal status As String, _
                      Optional ByVal expected As Variant = "", Optional ByVal actual As Variant = "", Optional ByVal details As String = "")
    Select Case status
        Case "PASS": mPass = mPass + 1
        Case "FAIL": mFail = mFail + 1
        Case "WARN": mWarn = mWarn + 1
        Case "SKIP": mSkip = mSkip + 1
        Case Else: mInfo = mInfo + 1
    End Select
    If mRes Is Nothing Then Exit Sub
    mResRow = mResRow + 1
    mRes.Range(mRes.Cells(mResRow, 1), mRes.Cells(mResRow, 8)).Value = _
        Array(id, cat, checkName, status, ToText(expected), ToText(actual), Trim$(details), Format$(Now, "hh:nn:ss"))
    mRes.Cells(mResRow, 4).Interior.Color = StatusColor(status)
End Sub

Private Sub LogCrash(ByVal id As String, ByVal cat As String)
    LogResult id, cat, "Checks stopped by an unexpected VBA error", "FAIL", , , "Error " & Err.Number & ": " & Err.Description
End Sub

Private Sub AddTiming(ByVal kind As String, ByVal iter As Long, ByVal ms As Double, ByVal ok As Boolean, ByVal note As String)
    mTimRow = mTimRow + 1
    mTim.Range(mTim.Cells(mTimRow, 1), mTim.Cells(mTimRow, 5)).Value = Array(kind, iter, Round(ms, 3), ok, note)
End Sub

Private Sub LogMismatch(ByVal source As String, ByVal iter As Long, ByVal formRow As Long, ByVal ds As Double, ByRef t() As Long, _
                        ByRef e As OTResult, ByVal gB As Variant, ByVal gC As Variant, ByVal gH As Variant, ByVal gI As Variant)
    If mMis Is Nothing Then Exit Sub
    If mMisRow - 1 >= MAX_MISMATCH_LOG Then Exit Sub
    mMisRow = mMisRow + 1
    mMis.Range(mMis.Cells(mMisRow, 1), mMis.Cells(mMisRow, 16)).Value = Array( _
        source, CStr(iter), CStr(formRow), IIf(ds = 0, "(blank)", Format$(CDate(ds), "yyyy-mm-dd ddd")), _
        HHMM(t(0)), HHMM(t(1)), HHMM(t(2)), HHMM(t(3)), e.DayText, ToText(gB), e.PH, ToText(gC), _
        ToText(e.Travel), ToText(gH), ToText(e.Work), ToText(gI))
End Sub

Private Function TimingStats(ByVal kind As String) As Variant   ' Array(n, min, avg, p95, max)
    Dim v As Variant, i As Long, j As Long, n As Long, arr() As Double, tmp As Double, sum As Double
    TimingStats = Array(0, 0#, 0#, 0#, 0#)
    If mTimRow < 2 Then Exit Function
    v = mTim.Range(mTim.Cells(2, 1), mTim.Cells(mTimRow, 4)).Value
    ReDim arr(1 To UBound(v, 1))
    For i = 1 To UBound(v, 1)
        If CStr(v(i, 1)) = kind And v(i, 4) = True Then
            n = n + 1
            arr(n) = v(i, 3)
            sum = sum + v(i, 3)
        End If
    Next i
    If n = 0 Then Exit Function
    For i = 2 To n
        tmp = arr(i)
        j = i - 1
        Do While j >= 1
            If arr(j) <= tmp Then Exit Do
            arr(j + 1) = arr(j)
            j = j - 1
        Loop
        arr(j + 1) = tmp
    Next i
    TimingStats = Array(n, arr(1), sum / n, arr(-Int(-0.95 * n)), arr(n))
End Function

Private Sub WriteSummary(ByVal seconds As Double)
    Dim ws As Worksheet, r As Long, kinds As Variant, k As Variant, st As Variant, v As Variant, i As Long
    Dim verdict As String, issuesTop As Long
    Set ws = ThisWorkbook.Worksheets(OUT_SUMMARY)
    ws.Cells.Clear
    If mFail > 0 Then
        verdict = "FAIL - " & mFail & " check(s) failed"
    ElseIf mWarn > 0 Then
        verdict = "PASS with " & mWarn & " warning(s)"
    Else
        verdict = "PASS"
    End If
    ws.Range("A1").Value = "Overtime form stress test"
    ws.Range("A1").Font.Size = 14: ws.Range("A1").Font.Bold = True
    ws.Range("A2").Value = verdict
    ws.Range("A2").Font.Bold = True
    ws.Range("A2").Interior.Color = StatusColor(IIf(mFail > 0, "FAIL", IIf(mWarn > 0, "WARN", "PASS")))

    r = 4
    SummaryLine ws, r, "Workbook tested", mSourcePath
    SummaryLine ws, r, "Run finished", Format$(Now, "yyyy-mm-dd hh:nn:ss")
    SummaryLine ws, r, "Duration", Format$(seconds, "0.0") & " s"
    SummaryLine ws, r, "Excel", Application.Version & " on " & Application.OperatingSystem
    SummaryLine ws, r, "Office hours on the form", HHMM(mQ1) & " - " & HHMM(mQ2)
    SummaryLine ws, r, "Holiday table", mHolCount & " rows for " & mHolYear
    SummaryLine ws, r, "Query M code fingerprint", mQueryFingerprint
    SummaryLine ws, r, "Settings", "refreshes " & REFRESH_ITERATIONS & ", random fills " & FUZZ_ITERATIONS & " x 60 rows, " & _
                                   "recalcs " & RECALC_ITERATIONS & ", seed " & RANDOM_SEED
    r = r + 1
    ws.Cells(r, 1).Value = "Status": ws.Cells(r, 2).Value = "Count"
    ws.Range(ws.Cells(r, 1), ws.Cells(r, 2)).Font.Bold = True
    For Each k In Array("PASS", "FAIL", "WARN", "SKIP", "INFO")
        r = r + 1
        ws.Cells(r, 1).Value = k
        ws.Cells(r, 1).Interior.Color = StatusColor(CStr(k))
        ws.Cells(r, 2).Value = Choose(Application.Match(k, Array("PASS", "FAIL", "WARN", "SKIP", "INFO"), 0), mPass, mFail, mWarn, mSkip, mInfo)
    Next k

    r = r + 2
    ws.Cells(r, 1).Value = "PERFORMANCE"
    ws.Cells(r, 1).Font.Bold = True
    r = r + 1
    ws.Range(ws.Cells(r, 1), ws.Cells(r, 6)).Value = Array("Timing (ms)", "Samples", "Min", "Average", "95th percentile", "Max")
    ws.Range(ws.Cells(r, 1), ws.Cells(r, 6)).Font.Bold = True
    ws.Range("H1").Value = r                       ' remembered for the Copilot sheet
    kinds = DistinctTimingKinds()
    If IsArray(kinds) Then
        For i = LBound(kinds) To UBound(kinds)
            st = TimingStats(CStr(kinds(i)))
            r = r + 1
            ws.Range(ws.Cells(r, 1), ws.Cells(r, 6)).Value = Array(kinds(i), st(0), Round(st(1), 1), Round(st(2), 1), Round(st(3), 1), Round(st(4), 1))
        Next i
    End If
    ws.Range("H2").Value = r

    r = r + 2
    ws.Cells(r, 1).Value = "ISSUES (every FAIL and WARN)"
    ws.Cells(r, 1).Font.Bold = True
    r = r + 1
    issuesTop = r
    ws.Range(ws.Cells(r, 1), ws.Cells(r, 5)).Value = Array("ID", "Category", "Check", "Status", "Details")
    ws.Range(ws.Cells(r, 1), ws.Cells(r, 5)).Font.Bold = True
    If mResRow >= 2 Then
        v = mRes.Range(mRes.Cells(2, 1), mRes.Cells(mResRow, 7)).Value
        For i = 1 To UBound(v, 1)
            If v(i, 4) = "FAIL" Or v(i, 4) = "WARN" Then
                r = r + 1
                ws.Range(ws.Cells(r, 1), ws.Cells(r, 5)).Value = Array(v(i, 1), v(i, 2), v(i, 3), v(i, 4), _
                    Trim$(IIf(Len(v(i, 6)) > 0, "Got " & v(i, 6) & IIf(Len(v(i, 5)) > 0, ", expected " & v(i, 5), "") & ". ", "") & v(i, 7)))
                ws.Cells(r, 4).Interior.Color = StatusColor(CStr(v(i, 4)))
            End If
        Next i
    End If
    If r = issuesTop Then
        r = r + 1
        ws.Cells(r, 1).Value = "(none)"
    End If
    ws.Range("H3").Value = issuesTop
    ws.Range("H4").Value = r
    ws.Columns("H").Hidden = True
    ws.Columns("A").ColumnWidth = 28
    ws.Columns("B:D").AutoFit
    ws.Columns("E").ColumnWidth = 100
    ws.Columns("E").WrapText = True
    mRes.Columns("A:H").AutoFit
    If mRes.Columns("G").ColumnWidth > 100 Then mRes.Columns("G").ColumnWidth = 100
    mTim.Columns("A:E").AutoFit
    mMis.Columns("A:P").AutoFit
End Sub

Private Sub SummaryLine(ByVal ws As Worksheet, ByRef r As Long, ByVal label As String, ByVal value As String)
    ws.Cells(r, 1).Value = label
    ws.Cells(r, 2).NumberFormat = "@"
    ws.Cells(r, 2).Value = value
    r = r + 1
End Sub

Private Function DistinctTimingKinds() As Variant
    Dim v As Variant, i As Long, s As String
    If mTimRow < 2 Then Exit Function
    v = mTim.Range(mTim.Cells(2, 1), mTim.Cells(mTimRow, 1)).Value
    If Not IsArray(v) Then v = As2D(v)
    For i = 1 To UBound(v, 1)
        If InStr(s, "|" & v(i, 1) & "|") = 0 Then s = s & "|" & v(i, 1) & "|"
    Next i
    s = Replace(s, "||", "|")
    DistinctTimingKinds = Split(Mid$(s, 2, Len(s) - 2), "|")
End Function

'==============================================================================
'  COPILOT SHEET - =COPILOT() explains the results; VBA decides PASS / FAIL
'==============================================================================
Private Sub WriteCopilotSheet()
    Dim ws As Worksheet, sm As Worksheet, perfTop As Long, perfEnd As Long, issTop As Long, issEnd As Long
    Dim issuesRef As String, perfRef As String, misRef As String, holRef As String, i As Long, r As Long
    Dim p As String, f As String, travelRef As String, workRef As String
    On Error GoTo Boom
    Set ws = ThisWorkbook.Worksheets(OUT_COPILOT)
    Set sm = ThisWorkbook.Worksheets(OUT_SUMMARY)
    ws.Cells.Clear
    perfTop = Val(sm.Range("H1").Value): perfEnd = Val(sm.Range("H2").Value)
    issTop = Val(sm.Range("H3").Value): issEnd = Val(sm.Range("H4").Value)
    issuesRef = "'" & OUT_SUMMARY & "'!$A$" & issTop & ":$E$" & issEnd
    perfRef = "'" & OUT_SUMMARY & "'!$A$" & perfTop & ":$F$" & perfEnd
    misRef = "'" & OUT_MISMATCH & "'!$A$1:$P$" & MinL(mMisRow, 51)

    ws.Range("A1").Value = "Copilot analysis of this stress-test run"
    ws.Range("A1").Font.Size = 14: ws.Range("A1").Font.Bold = True
    ws.Range("A2").Value = "Each block holds a =COPILOT() formula. It needs a Microsoft 365 Copilot licence and the COPILOT " & _
                           "function in Excel. #NAME? means the function is not available to you: copy the prompt above it into " & _
                           "the Copilot chat pane (Home > Copilot) and select the context range it names."
    ws.Range("A3").Value = "Copilot only explains. The PASS / FAIL verdicts come from the VBA checks, never from Copilot. " & _
                           "Every run rebuilds this sheet, and each formula uses your Copilot allowance."
    ws.Range("A2:A3").WrapText = True

    ' context the prompts need, written as plain text on this sheet
    ws.Range("H4").Value = "Context data for the prompts"
    ws.Range("H4").Font.Bold = True
    ws.Range("H5:I5").Value = Array("Public holiday", "Date")
    For i = 1 To mHolCount
        ws.Cells(5 + i, 8).Value = "Holiday " & i
        ws.Cells(5 + i, 9).NumberFormat = "@"
        ws.Cells(5 + i, 9).Value = Format$(mHol(i), "yyyy-mm-dd ddd")
    Next i
    holRef = "$H$5:$I$" & (5 + mHolCount)
    ws.Range("K5").Value = "Office hours"
    ws.Range("K6").NumberFormat = "@"
    ws.Range("K6").Value = HHMM(mQ1) & "-" & HHMM(mQ2)
    ws.Range("M5").Value = "Travel OT formula (engine H11)"
    ws.Range("M7").Value = "Work OT formula (engine I11)"
    ws.Range("M6,M8").NumberFormat = "@"
    ws.Range("M6").Value = IIf(Len(mEngFormulaH) > 0, mEngFormulaH, "(not captured)")
    ws.Range("M8").Value = IIf(Len(mEngFormulaI) > 0, mEngFormulaI, "(not captured)")
    travelRef = "$M$6": workRef = "$M$8"

    r = 5
    p = "You are a QA reviewer. These are the FAIL and WARN results of an automated stress test of an Excel overtime-claim " & _
        "form for field service engineers (travel and work overtime hours, Singapore public holidays loaded by Power Query). " & _
        "Summarise in at most 6 short bullet points for a manager: what is broken, what is risky, what is fine. FAIL items first."
    CopilotBlock ws, r, "1. Executive summary", p, issuesRef, "=COPILOT(" & XlText(p) & "," & issuesRef & ")"

    p = "For each FAIL or WARN below, give the most likely root cause in the workbook and one concrete fix " & _
        "(formula change, data validation rule, or process change). Do not suggest editing the Power Query. Issues:"
    f = "=COPILOT(" & XlText(p) & "," & issuesRef & "," & XlText("Travel OT formula for row 11:") & "," & travelRef & _
        "," & XlText("Work OT formula for row 11:") & "," & workRef & ")"
    CopilotBlock ws, r, "2. Root causes and fixes", p, issuesRef & ", " & travelRef & ", " & workRef, f

    p = "These are timing statistics in milliseconds from repeated Power Query refreshes and recalculations of the overtime " & _
        "form. Say whether performance is acceptable for an engineer filling in a monthly claim, point out outliers, " & _
        "one short sentence per timing."
    CopilotBlock ws, r, "3. Performance", p, perfRef, "=COPILOT(" & XlText(p) & "," & perfRef & ")"

    If mMisRow > 1 Then
        p = "Each row is an input where the form's result differed from an independent reference calculation. Find what the " & _
            "inputs have in common (weekday or weekend, public holiday, midnight crossing, which time columns are blank) and " & _
            "explain the pattern in plain words."
        CopilotBlock ws, r, "4. Mismatch pattern", p, misRef, "=COPILOT(" & XlText(p) & "," & misRef & ")"
    Else
        CopilotBlock ws, r, "4. Mismatch pattern", "(No mismatches in this run - nothing to analyse.)", "", ""
    End If

    p = "Create 15 new edge-case test rows for an overtime form. Return only a table with these 6 columns and no header: " & _
        "Description, Date (yyyy-mm-dd, year " & mHolYear & "), From, Until/From, Until/From, Until (24-hour hh:mm, empty when unused). " & _
        "The four times are travel-work-travel on one line and may cross midnight once. On weekdays only time outside the office " & _
        "hours given counts; weekends and the public holidays listed count in full. Mix: midnight crossings, times on the office-hour " & _
        "boundaries, public holidays, weekends, rows with only some time columns filled."
    f = "=COPILOT(" & XlText(p) & "," & holRef & "," & XlText("Office hours:") & ",$K$6)"
    CopilotBlock ws, r, "5. New test scenarios (copy the result, Paste Values into " & OUT_SCENARIOS & "!B" & SCEN_FIRST_ROW & _
                 " or the next free row, then run RunScenariosOnly)", p, holRef & ", $K$6", f

    If mPopupCount > 0 Then
        ws.Range("O4").Value = "Pop-ups found on the form"
        ws.Range("O4").Font.Bold = True
        ws.Range("O5:T5").Value = Array("Section", "Cell", "Pop-up title", "Pop-up text", "Error title", "Error text")
        ws.Range("O6:T" & (5 + mPopupCount)).NumberFormat = "@"
        For i = 1 To mPopupCount
            ws.Range(ws.Cells(5 + i, 15), ws.Cells(5 + i, 20)).Value = Array(mPopups(i, 1), mPopups(i, 2), mPopups(i, 3), _
                                                                            Replace(Replace(mPopups(i, 4), vbCr, ""), vbLf, " "), mPopups(i, 5), Replace(Replace(mPopups(i, 6), vbCr, ""), vbLf, " "))
        Next i
        p = "These are the input pop-ups and error messages on an Excel overtime form filled in by field service engineers. " & _
            "For each, say whether it is clear, consistent with the others (capitals, examples, date format) and matches what the " & _
            "cell accepts, and suggest a shorter or clearer wording where useful. Keep each suggestion under 255 characters."
        CopilotBlock ws, r, "6. Pop-up wording review", p, "$O$5:$T$" & (5 + mPopupCount), _
                     "=COPILOT(" & XlText(p) & ",$O$5:$T$" & (5 + mPopupCount) & ")"
    End If

    ws.Columns("A").ColumnWidth = 110
    ws.Columns("H:I").AutoFit
    ws.Columns("M").ColumnWidth = 40
    Exit Sub
Boom:
    LogResult "X01", "Runner", "Could not build the Copilot sheet", "WARN", , , "Error " & Err.Number & ": " & Err.Description
End Sub

Private Sub CopilotBlock(ByVal ws As Worksheet, ByRef r As Long, ByVal title As String, ByVal prompt As String, _
                         ByVal contextRef As String, ByVal formula As String)
    ws.Cells(r, 1).Value = title
    ws.Cells(r, 1).Font.Bold = True
    ws.Cells(r, 1).Interior.Color = RGB(221, 235, 247)
    ws.Cells(r + 1, 1).Value = "Prompt: " & prompt
    ws.Cells(r + 1, 1).WrapText = True
    If Len(contextRef) > 0 Then ws.Cells(r + 2, 1).Value = "Context: " & contextRef
    If Len(formula) > 0 Then
        On Error Resume Next
        ws.Cells(r + 3, 1).Formula2 = formula
        If Err.Number <> 0 Then
            Err.Clear
            ws.Cells(r + 3, 1).Formula = formula
        End If
        If Err.Number <> 0 Then
            Err.Clear
            ws.Cells(r + 3, 1).Value = "'" & formula
        End If
        On Error GoTo 0
        ws.Cells(r + 3, 1).WrapText = True
    End If
    r = r + 40                                         ' room for the answer to spill
End Sub

' A formula text literal. Excel limits each literal to 255 characters, so long prompts are joined with &.
Private Function XlText(ByVal s As String) As String
    Dim out As String
    s = Replace(s, """", "'")
    Do While Len(s) > 0
        If Len(out) > 0 Then out = out & "&"
        out = out & """" & Left$(s, 250) & """"
        s = Mid$(s, 251)
    Loop
    If Len(out) = 0 Then out = """"""
    XlText = out
End Function

'==============================================================================
'  SMALL UTILITIES
'==============================================================================
Private Function ToText(ByVal v As Variant) As String
    If IsError(v) Then
        ToText = ErrText(v)
    ElseIf IsEmpty(v) Then
        ToText = "(empty)"
    ElseIf VarType(v) = vbString Then
        If Len(v) = 0 Then ToText = "(blank)" Else ToText = v
    ElseIf VarType(v) = vbDate Then
        ToText = Format$(v, "yyyy-mm-dd hh:nn")
    ElseIf IsNumeric(v) Then
        ToText = Format$(v, "0.00")
    Else
        ToText = CStr(v)
    End If
End Function

Private Function ErrText(ByVal v As Variant) As String
    Select Case Mid$(CStr(v), 7)
        Case "2000": ErrText = "#NULL!"
        Case "2007": ErrText = "#DIV/0!"
        Case "2015": ErrText = "#VALUE!"
        Case "2023": ErrText = "#REF!"
        Case "2029": ErrText = "#NAME?"
        Case "2036": ErrText = "#NUM!"
        Case "2042": ErrText = "#N/A"
        Case Else: ErrText = CStr(v)
    End Select
End Function

Private Function FmtPair(ByVal a As Variant, ByVal b As Variant) As String
    FmtPair = "travel " & ToText(a) & " / work " & ToText(b)
End Function

Private Function DescribeInput(ByVal dVal As Variant, ByVal sD As String, ByVal sE As String, ByVal sF As String, ByVal sG As String) As String
    DescribeInput = IIf(IsEmpty(dVal), "(no date)", Format$(dVal, "ddd d mmm yyyy hh:nn")) & "  " & _
                    IIf(Len(sD) > 0, sD, "-") & " | " & IIf(Len(sE) > 0, sE, "-") & " | " & _
                    IIf(Len(sF) > 0, sF, "-") & " | " & IIf(Len(sG) > 0, sG, "-")
End Function

Private Function HHMM(ByVal minutes As Long) As String
    If minutes < 0 Then
        HHMM = "-"
    Else
        HHMM = Format$(minutes \ 60, "00") & ":" & Format$(minutes Mod 60, "00")
    End If
End Function

Private Function SameValue(ByVal got As Variant, ByVal expd As Variant) As Boolean
    If IsError(got) Then Exit Function
    If VarType(expd) = vbString Then
        If VarType(got) = vbString Then
            SameValue = (CStr(got) = CStr(expd))
        Else
            SameValue = (Len(expd) = 0 And IsEmpty(got))
        End If
    ElseIf VarType(got) = vbDouble Or VarType(got) = vbLong Or VarType(got) = vbInteger Or VarType(got) = vbCurrency Then
        SameValue = (Abs(CDbl(got) - CDbl(expd)) < TOL)
    End If
End Function

Private Function StatusColor(ByVal status As String) As Long
    Select Case status
        Case "PASS": StatusColor = RGB(198, 239, 206)
        Case "FAIL": StatusColor = RGB(255, 199, 206)
        Case "WARN": StatusColor = RGB(255, 235, 156)
        Case "SKIP": StatusColor = RGB(217, 217, 217)
        Case Else: StatusColor = RGB(221, 235, 247)
    End Select
End Function

Private Function Fingerprint(ByVal s As String) As String
    Dim i As Long, h As Double
    h = 5381
    For i = 1 To Len(s)
        h = h * 33 + AscW(Mid$(s, i, 1))
        h = h - Int(h / 4294967296#) * 4294967296#
    Next i
    Fingerprint = Len(s) & " chars, djb2 " & Format$(h, "0")
End Function

Private Function As2D(ByVal v As Variant) As Variant
    Dim a(1 To 1, 1 To 1) As Variant
    If IsArray(v) Then
        As2D = v
    Else
        a(1, 1) = v
        As2D = a
    End If
End Function

Private Function JoinRow(ByVal v As Variant) As String
    Dim j As Long
    For j = LBound(v, 2) To UBound(v, 2)
        JoinRow = JoinRow & IIf(j > LBound(v, 2), " | ", "") & CStr(v(1, j))
    Next j
End Function

Private Function Squash(ByVal s As String) As String
    s = Trim$(Replace(s, Chr$(160), " "))
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    Squash = s
End Function

Private Function NzStr(ByVal v As Variant) As String
    If IsNull(v) Or IsError(v) Then NzStr = "" Else NzStr = CStr(v)
End Function

Private Function MinL(ByVal a As Long, ByVal b As Long) As Long
    If a < b Then MinL = a Else MinL = b
End Function

Private Function MaxL(ByVal a As Long, ByVal b As Long) As Long
    If a > b Then MaxL = a Else MaxL = b
End Function

Private Sub Progress(ByVal what As String)
    Application.StatusBar = "Overtime stress test: " & what & "..."
    DoEvents
End Sub

Private Sub PauseMs(ByVal ms As Double)
    Dim t As Double
    t = NowMs()
    Do While NowMs() - t < ms
        DoEvents
    Loop
End Sub

Private Function NowMs() As Double
#If Mac Then
    NowMs = Timer * 1000#
#Else
    Dim c As Currency, f As Currency
    QueryPerformanceFrequency f
    QueryPerformanceCounter c
    NowMs = CDbl(c) / CDbl(f) * 1000#
#End If
End Function
