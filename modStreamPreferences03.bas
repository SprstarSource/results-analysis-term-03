Attribute VB_Name = "modStreamPreferences03"
Option Explicit

Private m_runSucceeded As Boolean

' ============================================================
' Preferences 03: Tecno D:G -> Data O:R; Litr D:E -> Data01 K:L, in FinalGuidance.xlsx
' Matching: normalized name tokens (order-independent, so surname/name order does not matter)
'           + birth date. If a birth date is empty on either side, a UNIQUE name match is accepted.
' Only matched wishes are replaced. Each track invalidates its own decisions and summary.
' Invalid/ambiguous source records are skipped; valid records continue to import.
' Unknown or unreadable wishes never skip a pupil: they are imported as EMPTY, with a warning.
' One final warning identifies data issues (up to 20 rows); skipped targets are not written.
' ============================================================

Private Const WISH_FILE As String = "FinalStreamPreferences.xlsx"
Private Const WISH_SHEET As String = "Tecno"
Private Const W_FIRST_ROW As Long = 3     ' headers in row 2
Private Const W_COL_NAME As Long = 2      ' B
Private Const W_COL_BIRTH As Long = 3     ' C
Private Const W_COL_W1 As Long = 4        ' D
Private Const W_COL_W2 As Long = 5        ' E
Private Const WISH_COUNT As Long = 4      ' source D:G; target O:R or K:N

Private Const DST_FILE As String = "FinalGuidance.xlsx"
Private Const DST_SHEET As String = "Data"
Private Const DST_TABLE As String = "Data"
Private Const C_NAME As Long = 2
Private Const C_BIRTH As Long = 3
Private Const C_W1 As Long = 15
Private Const C_W2 As Long = 16
Private Const MAX_LIST As Long = 15

Private Type WishesBatch
    Wishes As Variant
    WriteRows As Variant
    Matched As Long
    IsArts As Boolean
    Message As String
    MissingSource As Boolean
End Type

Public Sub Run_StreamPreferences03()
    Dim oldScreen As Boolean, oldAlerts As Boolean, oldEvents As Boolean, stateSaved As Boolean
    Dim scienceTouched As Boolean, artsTouched As Boolean
    Dim wbW As Workbook, wbDst As Workbook, wOpened As Boolean, dOpened As Boolean
    Dim science As WishesBatch, arts As WishesBatch, problems As Object
    Dim ws As Worksheet, lo As ListObject, failed As Boolean, errText As String, msg As String
    m_runSucceeded = False
    On Error GoTo EH
    oldScreen = Application.ScreenUpdating
    oldAlerts = Application.DisplayAlerts
    oldEvents = Application.EnableEvents
    stateSaved = True
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Set wbW = OpenWb(WISH_FILE, True, wOpened)
    Set wbDst = OpenWb(DST_FILE, False, dOpened)
    If wbDst.ReadOnly Then Err.Raise vbObjectError + 2, , DST_FILE & " is open read-only."
    Set problems = CreateObject("Scripting.Dictionary")
    PrepareWishes wbW, wbDst, False, science, problems
    PrepareWishes wbW, wbDst, True, arts, problems
    If science.MissingSource And arts.MissingSource Then Err.Raise vbObjectError + 84, , "Neither 'Tecno' nor 'Litr' was found in " & WISH_FILE & ". No wishes were written."
    If science.Matched > 0 Then
        scienceTouched = True
        CheckTrackWrite wbDst, RequireWorksheet(wbDst, "Data")
    End If
    If arts.Matched > 0 Then
        artsTouched = True
        CheckTrackWrite wbDst, RequireWorksheet(wbDst, "Data01")
    End If
    WriteWishes wbDst, science
    WriteWishes wbDst, arts
    If science.Matched + arts.Matched > 0 Then wbDst.Save
    m_runSucceeded = True
    msg = science.Message & vbLf & arts.Message
    If problems.Count > 0 Then msg = ValidationWarnings(problems, msg)
Done:
    On Error Resume Next
    If wOpened Then wbW.Close SaveChanges:=False
    If failed And dOpened Then
        wbDst.Close SaveChanges:=False
    ElseIf failed And Not wbDst Is Nothing Then
        For Each ws In wbDst.Worksheets
            If (ws.Name = "Data" And scienceTouched) Or (ws.Name = "Data01" And artsTouched) Then
                Set lo = Nothing
                If ws.Name = "Data01" Then Set lo = GuidanceTable(ws, 17) Else Set lo = GuidanceTable(ws, 27)
                If Not lo Is Nothing Then LockAndProtect ws, lo.DataBodyRange
            End If
        Next ws
    End If
    If stateSaved Then
        Application.EnableEvents = oldEvents
        Application.DisplayAlerts = oldAlerts
        Application.ScreenUpdating = oldScreen
    End If
    On Error GoTo 0
    If failed Then
        MsgBox "Wishes import failed:" & vbLf & errText, vbCritical, "Stream preferences 03"
    Else
        MsgBox msg, IIf(problems.Count > 0 Or science.MissingSource Or arts.MissingSource, vbExclamation, vbInformation), "Stream preferences 03"
    End If
    Exit Sub
EH:
    m_runSucceeded = False
    failed = True
    errText = Err.Description
    If (scienceTouched Or artsTouched) And Not dOpened Then
        errText = errText & vbLf & vbLf & ArtsText("failure_unsaved")
    End If
    Resume Done
End Sub

Private Sub PrepareWishes(ByVal wbW As Workbook, ByVal wbDst As Workbook, ByVal isArts As Boolean, _
                          ByRef batch As WishesBatch, ByVal problems As Object)
    Dim wsW As Worksheet, wsDst As Worksheet, lo As ListObject, body As Range
    Dim sourceSheet As String, targetSheet As String, wishColumn As Long, width As Long, wishCount As Long
    Dim tv As Variant, wv As Variant, wishOut As Variant, writeRows() As Boolean
    Dim dFull As Object, dNameCnt As Object, dNameRow As Object, dSource As Object, dWritten As Object
    Dim sourceKey As String, tdk() As String
    Dim nT As Long, i As Long, r As Long, k As Long, lastRow As Long, rowT As Long
    Dim nk As String, dk As String, fk As String, reason As String, nm As String
    Dim matched As Long, unmatched As Long, shown As Long, noWish As Long, lst As String
    Dim born As Date, location As String, firstSourceRow As Long, firstSourceIndex As Long
    Dim sourceNames() As String, sourceDates() As String, sourceOKRows() As Boolean
    Dim sourceTargets() As Long, badTargetDate() As Boolean, nSource As Long, skippedInvalid As Long
    batch.IsArts = isArts
    If isArts Then
        sourceSheet = "Litr": targetSheet = "Data01": wishColumn = 11: width = 17: wishCount = 2
    Else
        sourceSheet = "Tecno": targetSheet = "Data": wishColumn = 15: width = 27: wishCount = 4
    End If
    Set wsW = FindWorksheet(wbW, sourceSheet)
    If wsW Is Nothing Then
        batch.MissingSource = True
        batch.Message = targetSheet & ": source sheet '" & sourceSheet & "' not found in " & wbW.Name & "; skipped, existing data kept."
        Exit Sub
    End If
    Set wsDst = RequireWorksheet(wbDst, targetSheet)
    Set lo = GuidanceTable(wsDst, width)
    wsW.Calculate
    wsDst.Calculate
    If GuidanceTableIsEmpty(lo) Then
        batch.Message = targetSheet & ": no registered pupils."
        Exit Sub
    End If
    Set body = lo.DataBodyRange
    tv = body.Value2
    nT = UBound(tv, 1)
    ReDim writeRows(1 To nT)
    wishOut = body.Columns(wishColumn).Resize(, wishCount).Value2

    Set dSource = CreateObject("Scripting.Dictionary")
    Set dWritten = CreateObject("Scripting.Dictionary")
    Set dFull = CreateObject("Scripting.Dictionary")
    Set dNameCnt = CreateObject("Scripting.Dictionary")
    Set dNameRow = CreateObject("Scripting.Dictionary")
    ReDim tdk(1 To nT)
    ReDim badTargetDate(1 To nT)

    ' ---------- index target students ----------
    For i = 1 To nT
        nk = NameKey(tv(i, C_NAME))
        If Len(nk) > 0 Then
            If Not BlankValue(tv(i, C_BIRTH)) Then
                If TryBirthDate(tv(i, C_BIRTH), wbDst.Date1904, born) Then
                    tdk(i) = DateKey(tv(i, C_BIRTH), wbDst.Date1904)
                Else
                    AddValidationIssue problems, wsDst.Name & "!" & (body.Row + i - 1), "C"
                    badTargetDate(i) = True
                End If
            End If
            If dNameCnt.Exists(nk) Then
                dNameCnt(nk) = dNameCnt(nk) + 1
            Else
                dNameCnt.Add nk, 1
                dNameRow.Add nk, i
            End If
            If Len(tdk(i)) > 0 Then
                fk = nk & "|" & tdk(i)
                If dFull.Exists(fk) Then
                    dFull(fk) = -1            ' duplicate in target
                Else
                    dFull.Add fk, i
                End If
            End If
        End If
    Next i

    ' ---------- validate ALL source records before matching ----------
    ' This catches later duplicates before any record from their group is written.
    lastRow = wsW.Cells(wsW.Rows.Count, W_COL_NAME).End(xlUp).Row
    If lastRow >= W_FIRST_ROW Then
        wv = wsW.Range(wsW.Cells(W_FIRST_ROW, 1), wsW.Cells(lastRow, W_COL_W1 + wishCount - 1)).Value2
        nSource = UBound(wv, 1)
        ReDim sourceNames(1 To nSource)
        ReDim sourceDates(1 To nSource)
        ReDim sourceOKRows(1 To nSource)
        ReDim sourceTargets(1 To nSource)
        For r = 1 To nSource
            nk = NameKey(wv(r, W_COL_NAME))
            sourceNames(r) = nk
            If Len(nk) > 0 Then
                location = sourceSheet & "!" & (W_FIRST_ROW + r - 1)
                sourceOKRows(r) = True
                dk = ""
                If Not BlankValue(wv(r, W_COL_BIRTH)) Then
                    If TryBirthDate(wv(r, W_COL_BIRTH), wbW.Date1904, born) Then
                        dk = DateKey(wv(r, W_COL_BIRTH), wbW.Date1904)
                    Else
                        AddValidationIssue problems, location, "C"
                        sourceOKRows(r) = False
                    End If
                End If
                sourceDates(r) = dk
                If sourceOKRows(r) Then
                    sourceKey = nk & "|" & dk
                    If dSource.Exists(sourceKey) Then
                        firstSourceRow = CLng(dSource(sourceKey))
                        firstSourceIndex = firstSourceRow - W_FIRST_ROW + 1
                        AddValidationIssue problems, location, "DUP(" & firstSourceRow & ")"
                        AddValidationIssue problems, sourceSheet & "!" & firstSourceRow, "DUP(" & (W_FIRST_ROW + r - 1) & ")"
                        sourceOKRows(r) = False
                        sourceOKRows(firstSourceIndex) = False
                    Else
                        dSource.Add sourceKey, W_FIRST_ROW + r - 1
                    End If
                End If
                ' Unknown wishes: warn, keep the pupil, import the unknown wish as empty.
                For k = 0 To wishCount - 1
                    If IsUnknownWish(wv(r, W_COL_W1 + k), isArts) Then AddValidationIssue problems, location, ColumnLetter(W_COL_W1 + k)
                Next k
                ' Compare all six pairs. Repeated blanks are valid; repeated wishes skip only this record.
                If Not CollectWishListIssues(wv, r, W_COL_W1, location, problems, isArts, True) Then sourceOKRows(r) = False
            End If
        Next r
    End If

    ' ---------- match only validated records ----------
    For r = 1 To nSource
        If sourceOKRows(r) Then
            nk = sourceNames(r)
            dk = sourceDates(r)
            location = sourceSheet & "!" & (W_FIRST_ROW + r - 1)
            rowT = 0
            reason = ""
            If Not dNameCnt.Exists(nk) Then
                reason = "name not found"
            ElseIf Len(dk) > 0 Then
                fk = nk & "|" & dk
                If dFull.Exists(fk) Then
                    rowT = dFull(fk)
                    If rowT < 0 Then
                        rowT = 0
                        reason = "duplicate in target"
                    End If
                ElseIf dNameCnt(nk) = 1 Then
                    If Len(tdk(dNameRow(nk))) = 0 Then
                        rowT = dNameRow(nk)
                    Else
                        reason = "birth date mismatch"
                    End If
                Else
                    reason = "birth date mismatch"
                End If
            Else
                If dNameCnt(nk) = 1 Then
                    rowT = dNameRow(nk)
                Else
                    reason = "ambiguous name (no date)"
                End If
            End If

            If rowT > 0 Then
                If badTargetDate(rowT) Then
                    ' A malformed target birth date must never act as a missing date.
                    sourceOKRows(r) = False
                ElseIf dWritten.Exists(CStr(rowT)) Then
                    firstSourceRow = Abs(CLng(dWritten(CStr(rowT))))
                    firstSourceIndex = firstSourceRow - W_FIRST_ROW + 1
                    AddValidationIssue problems, location, "DUP(" & firstSourceRow & ")"
                    AddValidationIssue problems, sourceSheet & "!" & firstSourceRow, "DUP(" & (W_FIRST_ROW + r - 1) & ")"
                    sourceOKRows(r) = False
                    sourceOKRows(firstSourceIndex) = False
                    ' Negative owner keeps this target blocked for any third record.
                    dWritten(CStr(rowT)) = -firstSourceRow
                Else
                    dWritten.Add CStr(rowT), W_FIRST_ROW + r - 1
                    sourceTargets(r) = rowT
                End If
            Else
                unmatched = unmatched + 1
                If shown < MAX_LIST Then
                    nm = CStr(wv(r, W_COL_NAME))
                    lst = lst & vbLf & "  - " & nm & " (" & reason & ")"
                    shown = shown + 1
                End If
            End If
        End If
    Next r

    ' Build only unique valid assignments after all collisions have been excluded.
    For r = 1 To nSource
        If Len(sourceNames(r)) > 0 Then
            If Not sourceOKRows(r) Then
                skippedInvalid = skippedInvalid + 1
            ElseIf sourceTargets(r) > 0 Then
                rowT = sourceTargets(r)
                For k = 1 To wishCount
                    wishOut(rowT, k) = CleanWish(wv(r, W_COL_W1 + k - 1), isArts)
                Next k
                writeRows(rowT) = True
                matched = matched + 1
            End If
        End If
    Next r

    For i = 1 To nT
        Call CollectWishListIssues(wishOut, i, 1, wsDst.Name & "!" & (body.Row + i - 1), problems, isArts, False, wishColumn)
        If BlankValue(wishOut(i, 1)) Then noWish = noWish + 1
    Next i
    ' Data issues are warnings: valid assignments still proceed.


    batch.Wishes = wishOut
    batch.WriteRows = writeRows
    batch.Matched = matched
    batch.Message = targetSheet & ": imported " & matched & "; skipped invalid " & skippedInvalid & _
                    "; unmatched " & unmatched & "; no wish 01: " & noWish & "."
End Sub

Private Sub WriteWishes(ByVal wb As Workbook, ByRef batch As WishesBatch)
    Dim ws As Worksheet, lo As ListObject, body As Range, rows As Variant, wishes As Variant
    Dim width As Long, firstWish As Long, wishCount As Long, i As Long, k As Long, wishRow() As Variant
    If batch.Matched = 0 Then Exit Sub
    If batch.IsArts Then
        Set ws = wb.Worksheets("Data01"): width = 17: firstWish = 11: wishCount = 2
    Else
        Set ws = wb.Worksheets("Data"): width = 27: firstWish = 15: wishCount = 4
    End If
    ReDim wishRow(1 To 1, 1 To wishCount)
    Set lo = GuidanceTable(ws, width)
    Set body = lo.DataBodyRange
    If lo.ShowAutoFilter Then
        If lo.AutoFilter.FilterMode Then lo.AutoFilter.ShowAllData
    End If
    InvalidateOrientation wb, ws
    rows = batch.WriteRows
    wishes = batch.Wishes
    For i = LBound(rows) To UBound(rows)
        If rows(i) Then
            For k = 1 To wishCount
                wishRow(1, k) = wishes(i, k)
            Next k
            ws.Cells(body.Row + i - 1, firstWish).Resize(1, wishCount).Value2 = wishRow
        End If
    Next i
    LockAndProtect ws, body
End Sub

' Name key: normalized, lower-case, tokens sorted alphabetically (surname/name order does not matter)
Private Function NameKey(ByVal x As Variant) As String
    Dim s As String, t() As String, i As Long, j As Long, tmp As String, n As Long
    If IsError(x) Then Exit Function
    s = LCase$(Norm(x))
    s = Replace(s, ".", " ")
    s = Replace(s, "-", " ")
    s = Replace(s, "_", " ")
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    s = Trim$(s)
    If Len(s) = 0 Then Exit Function
    t = Split(s, " ")
    n = UBound(t)
    For i = 1 To n
        tmp = t(i)
        j = i - 1
        Do While j >= 0
            If StrComp(t(j), tmp, vbBinaryCompare) > 0 Then
                t(j + 1) = t(j)
                j = j - 1
            Else
                Exit Do
            End If
        Loop
        t(j + 1) = tmp
    Next i
    NameKey = Join(t, " ")
End Function

' Date key yyyymmdd ("" if empty or unreadable). Accepts date serials, yyyy-mm-dd and dd/mm/yyyy text.
Private Function DateKey(ByVal x As Variant, Optional ByVal date1904 As Boolean = False) As String
    Dim d As Date
    If IsError(x) Then Err.Raise vbObjectError + 44, , "Error value in a birth date."
    If BlankValue(x) Then Exit Function
    If Not TryBirthDate(x, date1904, d) Then
        Err.Raise vbObjectError + 44, , "Invalid birth date: " & CStr(x) & ". No wishes were written."
    End If
    DateKey = Format$(d, "yyyymmdd")
End Function

' Wish text -> canonical branch name; unrecognized text is kept as is; blank -> Empty
Private Function Canon(ByVal x As Variant) As Variant
    Dim t As String
    Canon = Empty
    If IsError(x) Then Exit Function
    t = Norm(x)
    If Len(t) = 0 Then Exit Function
    If InStr(t, Norm(ArtsText("f_lang"))) > 0 Then
        Canon = ArtsText("lang")
    ElseIf InStr(t, Norm(ArtsText("f_arts"))) > 0 Or InStr(t, Norm(ArtsText("f_philo"))) > 0 Then
        Canon = ArtsText("arts")
    ElseIf InStr(t, Norm(L("f_tech"))) > 0 Then
        Canon = L("tech")
    ElseIf InStr(t, Norm(L("f_mgmt"))) > 0 Then
        Canon = L("mgmt")
    ElseIf InStr(t, Norm(L("f_exp"))) > 0 Then
        Canon = L("exp")
    ElseIf InStr(t, Norm(L("f_math"))) > 0 Then
        Canon = L("math")
    Else
        Canon = Trim$(CStr(x))
    End If
End Function

' ---------- Common helpers ----------
Private Function Txt(ByVal codes As String) As String
    Dim parts() As String, i As Long, s As String
    parts = Split(codes, " ")
    For i = 0 To UBound(parts)
        s = s & ChrW(CLng("&H" & parts(i)))
    Next i
    Txt = s
End Function

Private Function L(ByVal key As String) As String
    Select Case key
    Case "math": L = Txt("631 64A 627 636 64A 627 62A")
    Case "tech": L = Txt("62A 642 646 64A 20 631 64A 627 636 64A")
    Case "exp": L = Txt("639 644 648 645 20 62A 62C 631 64A 628 64A 629")
    Case "mgmt": L = Txt("62A 633 64A 64A 631 20 648 627 642 62A 635 627 62F")
    Case "orient": L = Txt("62A 648 62C 64A 647")
    Case "repeat": L = Txt("625 639 627 62F 629 20 627 644 633 646 629")
    Case "yes": L = Txt("646 639 645")
    Case "sheet": L = ArtsText("science_report")
    Case "title": L = Txt("62D 648 635 644 629 20 627 642 62A 631 627 62D 627 62A 20 627 644 62A 648 62C 64A 647")
    Case "c_branch": L = Txt("627 644 634 639 628 629")
    Case "c_w1": L = Txt("627 644 631 63A 628 629 20 30 31")
    Case "c_w2": L = Txt("627 644 631 63A 628 629 20 30 32")
    Case "c_wsum": L = Txt("645 62C 645 648 639 20 627 644 631 63A 628 627 62A")
    Case "c_dec": L = Txt("642 631 627 631 627 62A 20 627 644 62A 648 62C 64A 647")
    Case "c_met": L = Txt("62A 644 628 64A 629 20 627 644 631 63A 628 629 20 30 31")
    Case "c_pct": L = Txt("627 644 646 633 628 629 20 645 646 20 627 644 646 627 62C 62D 64A 646")
    Case "total": L = Txt("627 644 645 62C 645 648 639")
    Case "item": L = Txt("627 644 628 64A 627 646")
    Case "count": L = Txt("627 644 639 62F 62F")
    Case "n_students": L = Txt("639 62F 62F 20 627 644 62A 644 627 645 64A 630")
    Case "n_pass": L = Txt("639 62F 62F 20 627 644 646 627 62C 62D 64A 646")
    Case "n_rep": L = Txt("639 62F 62F 20 642 631 627 631 627 62A 20 625 639 627 62F 629 20 627 644 633 646 629")
    Case "n_ori": L = Txt("639 62F 62F 20 642 631 627 631 627 62A 20 62A 648 62C 64A 647")
    Case "f_math": L = Txt("631 64A 627 636")
    Case "f_tech": L = Txt("62A 642 646 64A")
    Case "f_mgmt": L = Txt("62A 633 64A 64A 631")
    Case "f_exp": L = Txt("639 644 648 645")
    End Select
End Function

Private Function OpenWb(ByVal fileName As String, ByVal asReadOnly As Boolean, ByRef opened As Boolean) As Workbook
    Dim wb As Workbook, fullPath As String
    opened = False
    If Len(ThisWorkbook.Path) = 0 Then
        Err.Raise vbObjectError + 1, , "Save the tool workbook before running this procedure."
    End If
    fullPath = ThisWorkbook.Path & Application.PathSeparator & fileName
    For Each wb In Application.Workbooks
        If StrComp(wb.FullName, fullPath, vbTextCompare) = 0 Then
            Set OpenWb = wb
            Exit Function
        End If
        If StrComp(wb.Name, fileName, vbTextCompare) = 0 Then
            Err.Raise vbObjectError + 1, , "A different file with the same name is open: " & wb.FullName
        End If
    Next wb
    If Dir(fullPath) = "" Then Err.Raise vbObjectError + 1, , "File not found: " & fullPath
    Set OpenWb = Workbooks.Open(Filename:=fullPath, UpdateLinks:=0, ReadOnly:=asReadOnly)
    opened = True
End Function

' Normalizes Arabic text: alef/yaa/taa-marbuta variants, tatweel, NBSP, Arabic-Indic digits, extra spaces
Private Function Norm(ByVal v As Variant) As String
    Dim t As String, i As Long
    If IsError(v) Then Exit Function
    If IsNull(v) Then Exit Function
    t = CStr(v)
    t = Replace(t, ChrW(&H623), ChrW(&H627))
    t = Replace(t, ChrW(&H625), ChrW(&H627))
    t = Replace(t, ChrW(&H622), ChrW(&H627))
    t = Replace(t, ChrW(&H649), ChrW(&H64A))
    t = Replace(t, ChrW(&H629), ChrW(&H647))
    t = Replace(t, ChrW(&H640), "")
    t = Replace(t, ChrW(&HA0), " ")
    t = Replace(t, vbTab, " ")
    t = Replace(t, ChrW(&H200E), "")
    t = Replace(t, ChrW(&H200F), "")
    t = Replace(t, vbLf, " ")
    t = Replace(t, vbCr, " ")
    For i = 0 To 9
        t = Replace(t, ChrW(&H660 + i), CStr(i))
        t = Replace(t, ChrW(&H6F0 + i), CStr(i))
    Next i
    Do While InStr(t, "  ") > 0
        t = Replace(t, "  ", " ")
    Loop
    Norm = Trim$(t)
End Function

Private Function IsNum(ByVal v As Variant) As Boolean
    If VarType(v) <> vbDouble Then Exit Function
    IsNum = (CDbl(v) >= 0 And CDbl(v) <= 20)
End Function

Private Function Rnd2(ByVal x As Double) As Double
    Rnd2 = Application.WorksheetFunction.Round(x, 2)
End Function

Private Function RoundN(ByVal x As Double) As Long
    RoundN = CLng(Application.WorksheetFunction.Round(x, 0))
End Function

' Lock everything except wishes (O:R for science, K:N for arts), then protect the sheet (no password, filtering allowed)

' These are derived outputs: invalidate them whenever their inputs change.

Private Function KnownWish(ByVal x As Variant, Optional ByVal isArts As Boolean = False) As Boolean
    Dim t As String, k As Long, labels As Variant
    If IsError(x) Then Exit Function
    If BlankValue(x) Then
        KnownWish = True
        Exit Function
    End If
    t = Norm(x)
    If isArts Then
        labels = Array(ArtsText("arts"), ArtsText("lang"))
    Else
        labels = Array(L("math"), L("tech"), L("exp"), L("mgmt"))
    End If
    For k = LBound(labels) To UBound(labels)
        If t = Norm(labels(k)) Then
            KnownWish = True
            Exit Function
        End If
    Next k
End Function

' Known wish -> canonical branch name; unknown, unreadable or blank -> Empty.
Private Function CleanWish(ByVal x As Variant, Optional ByVal isArts As Boolean = False) As Variant
    Dim c As Variant
    CleanWish = Empty
    If IsError(x) Then Exit Function
    c = Canon(x)
    If KnownWish(c, isArts) Then CleanWish = c
End Function

Private Function IsUnknownWish(ByVal x As Variant, Optional ByVal isArts As Boolean = False) As Boolean
    If IsError(x) Then
        IsUnknownWish = True
    Else
        IsUnknownWish = Not KnownWish(Canon(x), isArts)
    End If
End Function

Private Function CollectWishListIssues(ByRef values As Variant, ByVal rowNo As Long, ByVal firstCol As Long, _
                                       ByVal location As String, ByVal problems As Object, ByVal isArts As Boolean, _
                                       ByVal cleanUnknown As Boolean, Optional ByVal displayFirst As Long = 0) As Boolean
    Dim one As Variant, two As Variant, a As Long, b As Long, displayCol As Long, wishCount As Long
    If isArts Then wishCount = 2 Else wishCount = 4
    CollectWishListIssues = True
    displayCol = displayFirst
    If displayCol = 0 Then displayCol = firstCol
    For a = 0 To wishCount - 1
        If cleanUnknown Then one = CleanWish(values(rowNo, firstCol + a), isArts) Else one = Canon(values(rowNo, firstCol + a))
        If IsError(values(rowNo, firstCol + a)) Or Not KnownWish(one, isArts) Then
            AddValidationIssue problems, location, ColumnLetter(displayCol + a)
            If Not cleanUnknown Then CollectWishListIssues = False
        End If
        If Not BlankValue(one) Then
            For b = a + 1 To wishCount - 1
                If cleanUnknown Then two = CleanWish(values(rowNo, firstCol + b), isArts) Else two = Canon(values(rowNo, firstCol + b))
                If Not BlankValue(two) Then
                    If Norm(one) = Norm(two) Then
                        AddValidationIssue problems, location, ColumnLetter(displayCol + a) & "=" & ColumnLetter(displayCol + b)
                        CollectWishListIssues = False
                    End If
                End If
            Next b
        End If
    Next a
End Function

' Workbook-aware birth date parsing. Invalid dates never become missing dates.
Private Function BlankValue(ByVal x As Variant) As Boolean
    If IsError(x) Then Exit Function
    If IsNull(x) Then
        BlankValue = True
    ElseIf IsEmpty(x) Then
        BlankValue = True
    Else
        BlankValue = (Len(Norm(x)) = 0)
    End If
End Function

Private Function DigitsOnly(ByVal s As String) As Boolean
    Dim i As Long, ch As String
    If Len(s) = 0 Then Exit Function
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch < "0" Or ch > "9" Then Exit Function
    Next i
    DigitsOnly = True
End Function

Private Function ValidYMD(ByVal ys As String, ByVal ms As String, ByVal ds As String, ByRef d As Date) As Boolean
    Dim y As Long, m As Long, dy As Long
    On Error GoTo InvalidDate
    If Len(ys) <> 4 Or Len(ms) > 2 Or Len(ds) > 2 Then Exit Function
    If Not DigitsOnly(ys) Or Not DigitsOnly(ms) Or Not DigitsOnly(ds) Then Exit Function
    y = CLng(ys): m = CLng(ms): dy = CLng(ds)
    If y < 1950 Or y > Year(Date) Then Exit Function
    If m < 1 Or m > 12 Or dy < 1 Or dy > 31 Then Exit Function
    d = DateSerial(y, m, dy)
    If Year(d) <> y Or Month(d) <> m Or Day(d) <> dy Then Exit Function
    ValidYMD = True
InvalidDate:
End Function

Private Function TryBirthDate(ByVal x As Variant, ByVal date1904 As Boolean, ByRef d As Date) As Boolean
    Dim s As String, parts() As String, serial As Double
    On Error GoTo InvalidDate
    If IsError(x) Then Exit Function
    If BlankValue(x) Then Exit Function
    Select Case VarType(x)
    Case vbDate
        d = DateSerial(Year(x), Month(x), Day(x))
        If Year(d) >= 1950 And Year(d) <= Year(Date) Then TryBirthDate = True
        Exit Function
    Case vbByte, vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal
        serial = CDbl(x)
        If serial < 1 Or serial >= 100000 Then Exit Function
        serial = Fix(serial)
        If date1904 Then serial = serial + 1462
        d = CDate(serial)
        If Year(d) >= 1950 And Year(d) <= Year(Date) Then TryBirthDate = True
        Exit Function
    Case vbString
        s = Norm(x)
    Case Else
        Exit Function
    End Select
    s = Replace(s, "-", "/")
    s = Replace(s, ".", "/")
    s = Replace(s, "\", "/")
    s = Replace(s, " ", "/")
    Do While InStr(s, "//") > 0
        s = Replace(s, "//", "/")
    Loop
    If InStr(s, "/") > 0 Then
        parts = Split(s, "/")
        If UBound(parts) <> 2 Then Exit Function
        If Len(parts(0)) = 4 Then
            TryBirthDate = ValidYMD(parts(0), parts(1), parts(2), d)
        ElseIf Len(parts(2)) = 4 Then
            TryBirthDate = ValidYMD(parts(2), parts(1), parts(0), d)
        End If
    ElseIf Len(s) = 8 And DigitsOnly(s) Then
        If ValidYMD(Left$(s, 4), Mid$(s, 5, 2), Right$(s, 2), d) Then
            TryBirthDate = True
        Else
            TryBirthDate = ValidYMD(Right$(s, 4), Mid$(s, 3, 2), Left$(s, 2), d)
        End If
    End If
InvalidDate:
End Function

' Collect data issues while valid records continue; the final warning lists up to 20.
' Short column codes keep the entire summary within the MsgBox text budget.
Private Sub AddValidationIssue(ByVal problems As Object, ByVal location As String, ByVal detail As String)
    If problems.Exists(location) Then
        If InStr(1, " " & problems(location) & " ", " " & detail & " ", vbBinaryCompare) = 0 Then
            problems(location) = problems(location) & " " & detail
        End If
    Else
        problems.Add location, detail
    End If
End Sub

Private Function ValidationWarnings(ByVal problems As Object, ByVal heading As String) As String
    Const MAX_ROWS As Long = 20
    Const MAX_LINE As Long = 36
    Const MAX_TEXT As Long = 1000
    Dim key As Variant, shown As Long, line As String, msg As String
    msg = heading
    If problems.Count > 0 Then
        msg = msg & vbLf & "Rows with data issues: " & problems.Count & " (up to 20 shown)."
        msg = msg & vbLf & "Letters=invalid columns (unknown D:G wishes left empty); X=Y=repeated non-empty wishes; DUP(row)=duplicate."
        For Each key In problems.Keys
            If shown >= MAX_ROWS Then Exit For
            line = CStr(key) & ": " & problems(key)
            If Len(line) > MAX_LINE Then line = Left$(line, MAX_LINE - 3) & "..."
            If Len(msg) + Len(line) + 40 > MAX_TEXT Then Exit For
            msg = msg & vbLf & line
            shown = shown + 1
        Next key
        If problems.Count > shown Then msg = msg & vbLf & "Additional issue rows: " & (problems.Count - shown)
    End If
    ValidationWarnings = msg
End Function

Private Function ArtsText(ByVal key As String) As String
    Select Case key
    Case "arts": ArtsText = Txt("622 62F 627 628 20 648 641 644 633 641 629")
    Case "lang": ArtsText = Txt("644 63A 627 62A 20 623 62C 646 628 64A 629")
    Case "f_arts": ArtsText = Txt("622 62F 627 628")
    Case "f_philo": ArtsText = Txt("641 644 633 641 629")
    Case "f_lang": ArtsText = Txt("644 63A 627 62A")
    Case "arts_prefix": ArtsText = Txt("623 648 644 649 20 62B 627 646 648 64A 20 62C 630 639 20 645 634 62A 631 643 20 622 62F 627 628")
    Case "science_report": ArtsText = Txt("62D 648 635 644 629 20 2D 20 639 644 648 645")
    Case "arts_report": ArtsText = Txt("62D 648 635 644 629 20 2D 20 622 62F 627 628")
    Case "legacy_report": ArtsText = Txt("627 644 62D 648 635 644 629")
    Case "arts_title": ArtsText = Txt("627 644 62A 648 62C 64A 647 20 627 644 646 647 627 626 64A 20 2D 20 623 648 644 649 20 62B 627 646 648 64A 20 62C 630 639 20 645 634 62A 631 643 20 622 62F 627 628")
    Case "arts_report_title": ArtsText = Txt("62D 635 64A 644 629 20 627 644 62A 648 62C 64A 647 20 627 644 646 647 627 626 64A 20 644 644 62C 630 639 20 627 644 645 634 62A 631 643 20 622 62F 627 628 20 62D 633 628 20 627 644 646 62A 627 626 62C 20 627 644 633 646 648 64A 629")
    Case "number": ArtsText = Txt("627 644 631 642 645")
    Case "name": ArtsText = Txt("627 644 644 642 628 20 648 627 644 627 633 645")
    Case "birth": ArtsText = Txt("62A 627 631 64A 62E 20 627 644 645 64A 644 627 62F")
    Case "sex": ArtsText = Txt("627 644 62C 646 633")
    Case "repeat": ArtsText = Txt("627 644 625 639 627 62F 629")
    Case "arabic": ArtsText = Txt("627 644 644 63A 629 20 627 644 639 631 628 64A 629 20 648 622 62F 627 628 647 627")
    Case "history": ArtsText = Txt("627 644 62A 627 631 64A 62E 20 648 627 644 62C 63A 631 627 641 64A 627")
    Case "french": ArtsText = Txt("627 644 644 63A 629 20 627 644 641 631 646 633 64A 629")
    Case "english": ArtsText = Txt("627 644 644 63A 629 20 627 644 625 646 62C 644 64A 632 64A 629")
    Case "average": ArtsText = Txt("627 644 645 639 62F 644 20 627 644 633 646 648 64A")
    Case "wish1": ArtsText = Txt("627 644 631 63A 628 629 20 30 31")
    Case "wish2": ArtsText = Txt("627 644 631 63A 628 629 20 30 32")
    Case "wish3": ArtsText = Txt("627 644 631 63A 628 629 20 30 33")
    Case "wish4": ArtsText = Txt("627 644 631 63A 628 629 20 30 34")
    Case "fine_arts": ArtsText = Txt("641 646 648 646")
    Case "voc_education": ArtsText = Txt("62A 639 644 64A 645 20 645 647 646 64A")
    Case "voc_training": ArtsText = Txt("62A 643 648 64A 646 20 645 647 646 64A")
    Case "p_lang": ArtsText = Txt("645 644 645 62D 20 644 63A 627 62A 20 623 62C 646 628 64A 629")
    Case "p_arts": ArtsText = Txt("645 644 645 62D 20 622 62F 627 628 20 648 641 644 633 641 629")
    Case "decision": ArtsText = Txt("627 642 62A 631 627 62D 20 645 633 62A 634 627 631 20 627 644 62A 648 62C 64A 647")
    Case "invalidated": ArtsText = Txt("62A 63A 64A 651 631 62A 20 628 64A 627 646 627 62A 20 647 630 627 20 627 644 645 633 627 631 2E 20 64A 62C 628 20 625 639 627 62F 629 20 62A 634 63A 64A 644 20 627 644 62A 648 62C 64A 647 20 644 62A 62D 62F 64A 62B 20 627 644 642 631 627 631 627 62A 20 648 627 644 62D 648 635 644 629 2E")
    Case "failure_unsaved": ArtsText = Txt("642 62F 20 62A 643 648 646 20 628 64A 627 646 627 62A 20 627 644 645 644 641 20 627 644 645 641 62A 648 62D 20 642 62F 20 62A 63A 64A 651 631 62A 20 62C 632 626 64A 64B 627 2E 20 644 627 20 62A 62D 641 638 20 627 644 645 644 641 60C 20 648 623 63A 644 642 647 20 62F 648 646 20 62D 641 638 20 62B 645 20 623 639 62F 20 641 62A 62D 647 2E")
    Case "header_expected": ArtsText = Txt("627 644 639 646 648 627 646 20 627 644 645 62A 648 642 639 20 64A 62A 636 645 646 3A 20")
    Case "h_number": ArtsText = Txt("631 642 645")
    Case "h_name": ArtsText = Txt("627 633 645")
    Case "h_birth": ArtsText = Txt("645 64A 644 627 62F")
    Case "h_sex": ArtsText = Txt("62C 646 633")
    Case "h_repeat": ArtsText = Txt("625 639 627 62F 629")
    Case "h_arabic": ArtsText = Txt("639 631 628 64A")
    Case "h_math": ArtsText = Txt("631 64A 627 636 64A 627 62A")
    Case "h_physics": ArtsText = Txt("641 64A 632 64A 627")
    Case "h_natural": ArtsText = Txt("637 628 64A 639")
    Case "h_history": ArtsText = Txt("62A 627 631 64A 62E")
    Case "h_french": ArtsText = Txt("641 631 646 633 64A")
    Case "h_english": ArtsText = Txt("627 646 62C 644 64A 632")
    Case "h_info": ArtsText = Txt("645 639 644 648 645 627 62A")
    Case "h_info_alt": ArtsText = Txt("625 639 644 627 645")
    Case "h_technology": ArtsText = Txt("62A 643 646 648 644 648 62C")
    Case "h_average": ArtsText = Txt("645 639 62F 644 20 627 644 641 635 644")
    Case "h_avg_word": ArtsText = Txt("645 639 62F 644")
    Case "h_term_word": ArtsText = Txt("641 635 644")
    Case "general_report": ArtsText = Txt("62D 648 635 644 629 20 639 627 645 629")
    Case "general_title": ArtsText = Txt("62D 635 64A 644 629 20 627 644 62A 648 62C 64A 647 20 627 644 646 647 627 626 64A")
    Case "general_subtitle": ArtsText = Txt("627 642 62A 631 627 62D 627 62A 20 627 644 62A 648 62C 64A 647 20 644 644 62A 644 627 645 64A 630 20 627 644 646 627 62C 62D 64A 646")
    Case "track_science": ArtsText = Txt("62C 20 645 20 639 644 648 645 20 648 62A 643 646 648 644 648 62C 64A 627")
    Case "track_arts": ArtsText = Txt("62C 20 645 20 622 62F 627 628")
    Case "rank1": ArtsText = Txt("627 644 62A 631 62A 64A 628")
    Case "rank2": ArtsText = Txt("627 644 62A 631 62A 64A 628 32")
    End Select
End Function

' Find the table by its A2 anchor, preserving an existing table name.
Private Function GuidanceTable(ByVal ws As Worksheet, ByVal width As Long) As ListObject
    Dim lo As ListObject
    For Each lo In ws.ListObjects
        If lo.Range.Column = 1 And lo.HeaderRowRange.Row = 2 Then
            If lo.ListColumns.Count < width Then Err.Raise vbObjectError + 80, , ws.Name & ": too few table columns."
            Set GuidanceTable = lo
            Exit Function
        End If
    Next lo
    Err.Raise vbObjectError + 80, , ws.Name & ": a table starting at A2 is required."
End Function

Private Function SummarySheet(ByVal wb As Workbook, ByVal isArts As Boolean, ByVal create As Boolean) As Worksheet
    Dim ws As Worksheet
    Set ws = FindWorksheet(wb, ArtsText("general_report"))
    If ws Is Nothing And create Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = ArtsText("general_report")
    End If
    Set SummarySheet = ws
End Function

Private Sub CheckTrackWrite(ByVal wb As Workbook, ByVal ws As Worksheet)
    Dim report As Worksheet, nm As Variant
    ws.Unprotect Password:=""
    If ws.ProtectContents Then Err.Raise vbObjectError + 81, , "Unprotect " & ws.Name & " before running."
    For Each nm In Array(ArtsText("general_report"), ArtsText("legacy_report"))
        Set report = FindWorksheet(wb, CStr(nm))
        If Not report Is Nothing Then
            report.Unprotect Password:=""
            If report.ProtectContents Then Err.Raise vbObjectError + 81, , "Unprotect " & report.Name & " before running."
        End If
    Next nm
End Sub

Private Sub LockAndProtect(ByVal ws As Worksheet, ByVal body As Range)
    Dim firstWish As Long, wishCount As Long
    If ws.Name = "Data01" Then firstWish = 11 Else firstWish = 15
    If ws.Name = "Data01" Then wishCount = 2 Else wishCount = 4
    ws.Cells.Locked = True
    If Not body Is Nothing Then body.Columns(firstWish).Resize(, wishCount).Locked = False
    ws.Protect Password:="", DrawingObjects:=True, Contents:=True, Scenarios:=True, _
               AllowFormattingColumns:=True, AllowFormattingRows:=True, AllowFiltering:=True
End Sub

Private Sub InvalidateOrientation(ByVal wb As Workbook, ByVal ws As Worksheet)
    Dim report As Worksheet, lastRow As Long, decisionCol As Long, nm As Variant
    If ws.Name = "Data01" Then decisionCol = 17 Else decisionCol = 27
    For Each nm In Array(ArtsText("general_report"), ArtsText("legacy_report"))
        Set report = FindWorksheet(wb, CStr(nm))
        If Not report Is Nothing Then
            report.Unprotect Password:=""
            If report.ProtectContents Then Err.Raise vbObjectError + 81, , "Unprotect " & report.Name & " before running."
        End If
    Next nm
    lastRow = ws.Cells(ws.Rows.Count, decisionCol).End(xlUp).Row
    If lastRow >= 3 Then ws.Range(ws.Cells(3, decisionCol), ws.Cells(lastRow, decisionCol)).ClearContents
    For Each nm In Array(ArtsText("general_report"), ArtsText("legacy_report"))
        Set report = FindWorksheet(wb, CStr(nm))
        If Not report Is Nothing Then report.Calculate
    Next nm
End Sub

Private Function ColumnLetter(ByVal n As Long) As String
    Do While n > 0
        ColumnLetter = Chr$(65 + (n - 1) Mod 26) & ColumnLetter
        n = (n - 1) \ 26
    Loop
End Function


' A placeholder row, whitespace, or formulas returning empty strings is not a pupil.
' A numeric zero, an error, or any other data makes the table non-empty and validated.
Private Function GuidanceTableIsEmpty(ByVal lo As ListObject) As Boolean
    Dim values As Variant, r As Long, c As Long
    GuidanceTableIsEmpty = True
    If lo.DataBodyRange Is Nothing Then Exit Function
    values = lo.DataBodyRange.Value2
    For r = 1 To UBound(values, 1)
        For c = 1 To UBound(values, 2)
            If IsError(values(r, c)) Then
                GuidanceTableIsEmpty = False
                Exit Function
            End If
            If Not BlankValue(values(r, c)) Then
                GuidanceTableIsEmpty = False
                Exit Function
            End If
        Next c
    Next r
End Function

Private Function FindWorksheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set FindWorksheet = wb.Worksheets(sheetName)
    Err.Clear
    On Error GoTo 0
End Function

Private Function RequireWorksheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    Set ws = FindWorksheet(wb, sheetName)
    If ws Is Nothing Then Err.Raise vbObjectError + 83, , "Required sheet '" & sheetName & "' not found in " & wb.Name & ". Run data transfer first."
    Set RequireWorksheet = ws
End Function

Public Function Preferences03Succeeded() As Boolean
    Preferences03Succeeded = m_runSucceeded
End Function
