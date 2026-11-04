"Attribute VB_Name = ""modAnalysisWorkbookBuilderAnalysis03"""
Option Explicit

'==============================================================================
 Module  : modAnalysisWorkbookBuilderAnalysis03
' Purpose : Adds Analysis01 (Term 3 from Data) and AnnualAnalysis (Annual
'           Averages from Data01) to a workbook that already has the Data
'           and Data01 sheets. One block per section, one under another.
' Design  : - Term 3 is read from the Data table (structured references ColRef).
'           - Annual Averages are read from the Data01 table (ColRef01).
"'           - Subjects are discovered from the Data headers (see DiscoverData);"
'             the labels are the Data headers without the term suffix.
'           - A mark counts only when it is > 0 (zero = not taught / exempt).
"'           - Live formulas only; Office 365 (MAXIFS, MINIFS, MODE.SNGL, dynamic arrays)."
"'           - Layout = union of the sample workbooks and modExamenAnalysis02;"
'             formats, fonts and colours = modExamenAnalysis02.
'==============================================================================

"' Keywords (normalised Arabic, no hamza) of the ""core"" subjects of each file"
Public Function CoreKeysFor(ByVal stem As String) As Variant
    ' Core subjects = the yellow cells of the coefficients workbook (per stream and per year)
    Select Case stem
"        Case ""1LitAnalysis03"", ""1LitAnalysis02"""
"            CoreKeysFor = Array(ArT(""627 644 639 631 628 64A 629""), ArT(""627 644 641 631 646 633 64A 629""), ArT(""627 644 627 646 62C 644 64A 632 64A 629""), ArT(""627 644 62A 627 631 64A 62E""))"
"        Case ""1TecnoGlobalAnalysis03"", ""2SEGlobalAnalysis03"", ""3SEGlobalAnalysis03"", _"
"             ""1TecnoGlobalAnalysis02"", ""2SEGlobalAnalysis02"", ""3SEGlobalAnalysis02"""
"            CoreKeysFor = Array(ArT(""627 644 631 64A 627 636 64A 627 62A""), ArT(""627 644 641 64A 632 64A""), ArT(""627 644 637 628 64A 639""))"
"        Case ""2MathAnalysis03"", ""3MathAnalysis03"", ""2MathAnalysis02"", ""3MathAnalysis02"""
"            CoreKeysFor = Array(ArT(""627 644 631 64A 627 636 64A 627 62A""), ArT(""627 644 641 64A 632 64A""))"
"        Case ""2GEAnalysis03"", ""3GEAnalysis03"", ""2GEAnalysis02"", ""3GEAnalysis02"""
"            CoreKeysFor = Array(ArT(""627 644 62A 627 631 64A 62E""), ArT(""627 644 631 64A 627 636 64A 627 62A""), ArT(""627 644 62A 633 64A 64A 631""), ArT(""627 644 645 62D 627 633 628""), ArT(""627 644 627 642 62A 635 627 62F""))"
"        Case ""2PhiloAnalysis03"", ""3PhiloAnalysis03"", ""2PhiloAnalysis02"", ""3PhiloAnalysis02"""
"            CoreKeysFor = Array(ArT(""627 644 639 631 628 64A 629""), ArT(""627 644 641 631 646 633 64A 629""), ArT(""627 644 627 646 62C 644 64A 632 64A 629""), ArT(""627 644 62A 627 631 64A 62E""), ArT(""627 644 641 644 633 641 629""))"
"        Case ""2LEAnalysis03"", ""2LEAnalysis02"""
"            CoreKeysFor = Array(ArT(""627 644 639 631 628 64A 629""), ArT(""627 644 641 631 646 633 64A 629""), ArT(""627 644 627 646 62C 644 64A 632 64A 629""), ArT(""627 644 627 644 645 627 646 64A 629""), ArT(""627 644 627 633 628 627 646 64A 629""), ArT(""627 644 627 64A 637 627 644 64A 629""), ArT(""627 644 62A 627 631 64A 62E""))"
"        Case ""3LEAnalysis03"", ""3LEAnalysis02"""
"            CoreKeysFor = Array(ArT(""627 644 639 631 628 64A 629""), ArT(""627 644 641 631 646 633 64A 629""), ArT(""627 644 627 646 62C 644 64A 632 64A 629""), ArT(""627 644 627 644 645 627 646 64A 629""), ArT(""627 644 627 633 628 627 646 64A 629""), ArT(""627 644 627 64A 637 627 644 64A 629""))"
        Case Else
            CoreKeysFor = Array()
    End Select
End Function

Private Function LangKeys() As Variant
"    LangKeys = Array(ArT(""627 644 641 631 646 633 64A 629""), ArT(""627 644 627 646 62C 644 64A 632 64A 629""), ArT(""627 644 627 644 645 627 646 64A 629""), ArT(""627 644 627 633 628 627 646 64A 629""), ArT(""627 644 627 64A 637 627 644 64A 629""))"
End Function

Private Function InList(ByVal subjName As String, ByVal keys As Variant) As Boolean
    Dim i As Long, nm As String, kw As String
    nm = NormAr(subjName)
    For i = LBound(keys) To UBound(keys)
        kw = NormAr(CStr(keys(i)))
        If Len(kw) > 0 Then
            If InStr(nm, kw) > 0 Then
                InList = True
                Exit Function
            End If
        End If
    Next i
End Function

'==============================================================================
' MAIN
'==============================================================================
Public Function BuildAnalysisSheets(ByVal wb As Workbook, ByVal coreKeys As Variant, _
                                    ByRef errText As String) As Boolean
    Dim t As Long, s As Long, r0 As Long, rStart As Long, ws As Worksheet
    Dim lastCol As Long

    BuildAnalysisSheets = False
"    errText = """""
    On Error GoTo EH

    If Not DiscoverData(wb, errText) Then Exit Function

    lastCol = FIRST_NUM_COL + gSubjN

    ' t = 1 -> Analysis01 (Term 3 from Data)
    ' t = 2 -> AnnualAnalysis (Annual Averages from Data01)
    For t = 1 To 2
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        If t = 1 Then
            ws.Name = ANALYSIS_T3_SHEET
        Else
            ws.Name = ANALYSIS_ANN_SHEET
        End If
        ws.DisplayRightToLeft = True
        r0 = 1
        For s = 1 To gSecN
            rStart = r0
            r0 = BuildBlock(ws, r0, t, gSecList(s), coreKeys)
            If s > 1 Then
                On Error Resume Next
                ws.HPageBreaks.Add Before:=ws.Cells(rStart, 1)
                Err.Clear
                On Error GoTo EH
            End If
        Next s
        SetColumnWidths ws, lastCol
        SetupPrint ws, r0 - 3, lastCol
    Next t

    BuildAnalysisSheets = True
    Exit Function
EH:
"    errText = ""BuildAnalysisSheets: "" & Err.Number & "" : "" & Err.Description"
    Err.Clear
End Function

'==============================================================================
"' ONE BLOCK  (banner, counters, table 1, four ""passed"" cells, bands, indicators, classes)"
' Returns the first free row after the block (2 blank rows left)
'==============================================================================
Private Function BuildBlock(ByVal ws As Worksheet, ByVal r0 As Long, ByVal t As Long, _
                            ByVal sec As Long, ByVal coreKeys As Variant) As Long
    Dim n As Long, lastCol As Long, k As Long, col As Long, i As Long
    Dim ref() As String, hdrText() As String
    Dim sc As String, sx As String, rp As String, annColHdr As String
    Dim secN As String, c As String, f As String, ttl As String
    Dim pe As String, fem As String, mal As String, yes As String, pm As String
    Dim rT As Long, rC1 As Long, rC2 As Long, rC3 As Long, rH1 As Long
    Dim rAvg As Long, rFem As Long, rMal As Long, rRate As Long, rRep As Long, rMax As Long, rMin As Long
    Dim rQ1 As Long, rT2 As Long, rH2 As Long, rB1 As Long, rH3 As Long, rI1 As Long, rK1 As Long
    Dim prodAll As String, prodCore As String, prodCoreF As String, prodLang As String
    Dim nCore As Long, nLang As Long
    Dim lows As Variant, highs As Variant, bandLbl As Variant
    Dim mAddr As String, sdAddr As String, avgRef As String
    Dim nQ As Long, qLbl(1 To 4) As String, qF(1 To 4) As String
    Dim a1() As Variant, a2() As Variant, a3() As Variant, a4() As Variant, nc As Long, j As Long

    n = gSubjN
    lastCol = FIRST_NUM_COL + n
    ReDim ref(FIRST_NUM_COL To lastCol)
    ReDim hdrText(FIRST_NUM_COL To lastCol)

    If t = 1 Then
        ' Analysis01: Term 3 data directly from Data sheet
        sx = ColRef(gHdr(4))
        rp = ColRef(gHdr(5))
        sc = ColRef(gHdr(gSecCol))
        For k = 1 To n
            ref(FIRST_NUM_COL + k - 1) = ColRef(gHdr(gSubjC3(k)))
            hdrText(FIRST_NUM_COL + k - 1) = gSubjName(k)
        Next k
        ref(lastCol) = ColRef(gHdr(gAvgC3))
        hdrText(lastCol) = gHdr(gAvgC3)
    Else
        ' AnnualAnalysis: Annual Average data directly from Data01 sheet
        sx = ColRef01(gHdr(4))
        rp = ColRef01(gHdr(5))
        sc = ColRef01(gHdr(gSecCol))
        For k = 1 To n
            ref(FIRST_NUM_COL + k - 1) = ColRef01(gSubjName(k))
            hdrText(FIRST_NUM_COL + k - 1) = gSubjName(k)
        Next k
        annColHdr = AnnualAvgHeader()
        ref(lastCol) = ColRef01(annColHdr)
        hdrText(lastCol) = annColHdr
    End If
    avgRef = ref(lastCol)

    secN = CStr(sec)
"    fem = ArT(""623 646 62B 649"")"
"    mal = ArT(""630 643 631"")"
"    yes = ArT(""646 639 645"")"
    pm = CStr(PASS_MARK)

"    ' ---- the ""passed"" cells (built first because their number moves the rows below)"
    '   all subjects : a mark >= 10, or 0 (= not taught / exempt, e.g. physical education)
    '   core subjects: all >= 10 (the term average does not matter)
    '   without foreign languages: only when foreign languages are core subjects
    '   failed all core subjects: no core mark >= 10
    For k = 1 To n
        c = ref(FIRST_NUM_COL + k - 1)
"        prodAll = prodAll & ""*((("" & c & "">="" & pm & "")+("" & c & ""=0))>0)"""
        If InList(gSubjName(k), coreKeys) Then
            nCore = nCore + 1
"            prodCore = prodCore & ""*("" & c & "">="" & pm & "")"""
"            prodCoreF = prodCoreF & ""*("" & c & ""<"" & pm & "")"""
            If InList(gSubjName(k), LangKeys()) Then
                nLang = nLang + 1
"                prodLang = prodLang & ""*("" & c & ""<"" & pm & "")"""
            End If
        End If
    Next k

    nQ = 1
"    qLbl(1) = ArT(""639 62F 62F 20 627 644 62A 644 627 645 64A 630 20 627 644 646 627 62C 62D 64A 646 20 641 64A 20 62C 645 64A 639 20 627 644 645 648 627 62F"")"
"    qF(1) = ""=SUMPRODUCT(("" & sc & ""="" & secN & "")"" & prodAll & "")"""
    If nCore > 0 Then
        nQ = nQ + 1
"        qLbl(nQ) = ArT(""639 62F 62F 20 627 644 62A 644 627 645 64A 630 20 627 644 646 627 62C 62D 64A 646 20 641 64A 20 627 644 645 648 627 62F 20 627 644 623 633 627 633 64A 629"")"
"        qF(nQ) = ""=SUMPRODUCT(("" & sc & ""="" & secN & "")"" & prodCore & "")"""
    End If
    If nLang > 0 Then
        nQ = nQ + 1
"        qLbl(nQ) = ArT(""639 62F 62F 20 627 644 62A 644 627 645 64A 630 20 627 644 646 627 62C 62D 64A 646 20 62F 648 646 20 627 644 644 63A 627 62A 20 627 644 623 62C 646 628 64A 629"")"
"        qF(nQ) = ""=SUMPRODUCT(("" & sc & ""="" & secN & "")*("" & avgRef & "">="" & pm & "")"" & prodLang & "")"""
    End If
    If nCore > 0 Then
        nQ = nQ + 1
"        qLbl(nQ) = ArT(""639 62F 62F 20 627 644 62A 644 627 645 64A 630 20 63A 64A 631 20 627 644 646 627 62C 62D 64A 646 20 641 64A 20 623 64A 20 645 627 62F 629 20 623 633 627 633 64A 629"")"
"        qF(nQ) = ""=SUMPRODUCT(("" & sc & ""="" & secN & "")"" & prodCoreF & "")"""
    End If

    ' ---- rows of the block
    rT = r0
    rC1 = r0 + 1: rC2 = r0 + 2: rC3 = r0 + 3
    rH1 = r0 + 5
    rAvg = rH1 + 1: rFem = rH1 + 2: rMal = rH1 + 3: rRate = rH1 + 4
    rRep = rH1 + 5: rMax = rH1 + 6: rMin = rH1 + 7
    rQ1 = rH1 + 9
    rT2 = rQ1 + nQ + 1: rH2 = rT2 + 1: rB1 = rH2 + 1
    rH3 = rB1 + 9: rI1 = rH3 + 1
    rK1 = rI1 + 9

    ' ---- banner
    If t = 1 Then
"        ' ""تحليل نتائج الفصل الثالث قسم """
"        ttl = ArT(""62A 62D 644 64A 644 20 646 62A 627 626 62C 20 627 644 641 635 644 20 627 644 62B 627 644 62B 20 642 633 645 20"") & gClassTitle"
    Else
"        ' ""تحليل معطيات المعدل السنوي قسم """
"        ttl = ArT(""62A 62D 644 64A 644 20 645 639 637 64A 627 62A 20 627 644 645 639 62F 644 20 627 644 633 646 648 64A 20 642 633 645 20"") & gClassTitle"
    End If
"    If gSecN > 1 Then ttl = ttl & "" "" & Format$(sec, ""00"")"
    MergeText ws, rT, 1, lastCol, ttl
    StyleBanner ws.Range(ws.Cells(rT, 1), ws.Cells(rT, lastCol))
    ws.Rows(rT).RowHeight = TITLE_ROW_HEIGHT

    ' ---- counters
"    PutStrip ws, rC1, 1, 3, 4, ArT(""627 644 62A 639 62F 627 62F""), ""=COUNTIFS("" & sc & "","" & secN & "")"""
"    PutStrip ws, rC1, 5, 7, 8, ArT(""627 644 645 639 64A 62F 648 646""), ""=COUNTIFS("" & sc & "","" & secN & "","" & rp & "","" & QT & yes & QT & "")"""
"    PutStrip ws, rC2, 1, 3, 4, ArT(""639 62F 62F 20 627 644 630 643 648 631""), ""=COUNTIFS("" & sc & "","" & secN & "","" & sx & "","" & QT & mal & QT & "")"""
    If gPEIdx > 0 Then
        pe = ref(FIRST_NUM_COL + gPEIdx - 1)
"        f = ""=COUNTIFS("" & sc & "","" & secN & "","" & pe & "",0)+COUNTIFS("" & sc & "","" & secN & "","" & pe & "","" & QT & QT & "")"""
    Else
"        f = ""=0"""
    End If
"    PutStrip ws, rC2, 5, 7, 8, ArT(""627 644 645 639 641 64A 648 646""), f"
"    PutStrip ws, rC3, 1, 3, 4, ArT(""639 62F 62F 20 627 644 625 646 627 62B""), ""=COUNTIFS("" & sc & "","" & secN & "","" & sx & "","" & QT & fem & QT & "")"""
    ws.Rows(r0 + 4).RowHeight = 6

    ' ---- table 1 : per subject (the whole table is written with ONE assignment)
    nc = lastCol - FIRST_NUM_COL + 1
"    PutHeaderRow ws, rH1, lastCol, ArT(""627 644 645 648 627 62F""), hdrText"
    ReDim a1(1 To 7, 1 To nc)
    For col = FIRST_NUM_COL To lastCol
        j = col - FIRST_NUM_COL + 1
        c = ref(col)
"        a1(1, j) = Subst(""=IFERROR(AVERAGEIFS({C},{S},{N},{C},"""">0""""),"""""""")"", ""C"", c, ""S"", sc, ""N"", secN)"
"        a1(2, j) = Subst(""=COUNTIFS({S},{N},{C},"""">={P}"""",{X},""""{F}"""")"", ""C"", c, ""S"", sc, ""N"", secN, ""P"", pm, ""X"", sx, ""F"", fem)"
"        a1(3, j) = Subst(""=COUNTIFS({S},{N},{C},"""">={P}"""",{X},""""{F}"""")"", ""C"", c, ""S"", sc, ""N"", secN, ""P"", pm, ""X"", sx, ""F"", mal)"
"        a1(4, j) = Subst(""=IFERROR(COUNTIFS({S},{N},{C},"""">={P}"""")/COUNTIFS({S},{N},{C},"""">0""""),"""""""")"", ""C"", c, ""S"", sc, ""N"", secN, ""P"", pm)"
"        a1(5, j) = Subst(""=COUNTIFS({S},{N},{C},"""">={P}"""",{X},""""{F}"""")"", ""C"", c, ""S"", sc, ""N"", secN, ""P"", pm, ""X"", rp, ""F"", yes)"
"        a1(6, j) = Subst(""=IFERROR(MAXIFS({C},{S},{N},{C},"""">0""""),"""""""")"", ""C"", c, ""S"", sc, ""N"", secN)"
"        a1(7, j) = Subst(""=IFERROR(MINIFS({C},{S},{N},{C},"""">0""""),"""""""")"", ""C"", c, ""S"", sc, ""N"", secN)"
    Next col
    ws.Range(ws.Cells(rAvg, FIRST_NUM_COL), ws.Cells(rMin, lastCol)).Formula = a1
    StyleGrid ws.Range(ws.Cells(rAvg, FIRST_NUM_COL), ws.Cells(rMin, lastCol))
    ws.Range(ws.Cells(rAvg, FIRST_NUM_COL), ws.Cells(rAvg, lastCol)).NumberFormat = FMT_DEC
    ws.Range(ws.Cells(rFem, FIRST_NUM_COL), ws.Cells(rMal, lastCol)).NumberFormat = FMT_COUNT
    ws.Range(ws.Cells(rRate, FIRST_NUM_COL), ws.Cells(rRate, lastCol)).NumberFormat = FMT_PCT
    ws.Range(ws.Cells(rRep, FIRST_NUM_COL), ws.Cells(rRep, lastCol)).NumberFormat = FMT_COUNT
    ws.Range(ws.Cells(rMax, FIRST_NUM_COL), ws.Cells(rMin, lastCol)).NumberFormat = FMT_DEC
"    PutLabel ws, rAvg, ArT(""645 639 62F 644 627 62A 20 627 644 645 648 627 62F"")"
"    PutLabel ws, rFem, ArT(""639 62F 62F 20 627 644 62A 644 645 64A 630 627 62A 20 627 644 646 627 62C 62D 627 62A"")"
"    PutLabel ws, rMal, ArT(""639 62F 62F 20 627 644 62A 644 627 645 64A 630 20 627 644 646 627 62C 62D 64A 646 20 28 630 643 648 631 29"")"
"    PutLabel ws, rRate, ArT(""646 633 628 629 20 627 644 646 62C 627 62D 20 641 64A 20 643 644 20 645 627 62F 629"")"
"    PutLabel ws, rRep, ArT(""639 62F 62F 20 627 644 62A 644 627 645 64A 630 20 627 644 645 639 64A 62F 64A 646 20 627 644 646 627 62C 62D 64A 646"")"
"    PutLabel ws, rMax, ArT(""623 639 644 649 20 645 639 62F 644"")"
"    PutLabel ws, rMin, ArT(""623 62F 646 649 20 645 639 62F 644"")"
    ws.Rows(rMin + 1).RowHeight = 6

"    ' ---- the ""passed"" cells"
    For i = 1 To nQ
        PutLabel ws, rQ1 + i - 1, qLbl(i)
        StyleGrid ws.Cells(rQ1 + i - 1, FIRST_NUM_COL)
        ws.Cells(rQ1 + i - 1, FIRST_NUM_COL).NumberFormat = FMT_COUNT
        ws.Cells(rQ1 + i - 1, FIRST_NUM_COL).Formula = qF(i)
    Next i
    ws.Rows(rQ1 + nQ).RowHeight = 6

    ' ---- table 2 : distribution of the marks in bands
"    MergeText ws, rT2, 1, lastCol, ArT(""62A 648 632 64A 639 20 627 644 645 639 62F 644 627 62A 20 62D 633 628 20 627 644 641 626 627 62A"")"
    StyleTableTitle ws.Range(ws.Cells(rT2, 1), ws.Cells(rT2, lastCol))
    ws.Rows(rT2).RowHeight = 36
"    PutHeaderRow ws, rH2, lastCol, ArT(""627 644 641 626 627 62A 20 2F 20 627 644 645 648 627 62F""), hdrText"

    lows = Array(0, 7, 9, 10, 12, 14, 16, 18)
    highs = Array(7, 9, 10, 12, 14, 16, 18, 0)
"    bandLbl = Array(""07 >"", ""08.99 - 07"", ""09.99 - 09"", ""11.99 - 10"", ""13.99 - 12"", ""15.99 - 14"", ""17.99 - 16"", ""18 =<"")"
    ReDim a2(1 To 8, 1 To nc)
    For i = 0 To 7
        For col = FIRST_NUM_COL To lastCol
            j = col - FIRST_NUM_COL + 1
            c = ref(col)
            If i = 0 Then
"                f = Subst(""=COUNTIFS({S},{N},{C},"""">0"""",{C},""""<{H}"""")"", ""S"", sc, ""N"", secN, ""C"", c, ""H"", CStr(highs(i)))"
            ElseIf i = 7 Then
"                f = Subst(""=COUNTIFS({S},{N},{C},"""">={L}"""")"", ""S"", sc, ""N"", secN, ""C"", c, ""L"", CStr(lows(i)))"
            Else
"                f = Subst(""=COUNTIFS({S},{N},{C},"""">={L}"""",{C},""""<{H}"""")"", ""S"", sc, ""N"", secN, ""C"", c, ""L"", CStr(lows(i)), ""H"", CStr(highs(i)))"
            End If
            a2(i + 1, j) = f
        Next col
"        ws.Cells(rB1 + i, 1).NumberFormat = ""@"""
    Next i
    ws.Range(ws.Cells(rB1, FIRST_NUM_COL), ws.Cells(rB1 + 7, lastCol)).Formula = a2
    StyleGrid ws.Range(ws.Cells(rB1, FIRST_NUM_COL), ws.Cells(rB1 + 7, lastCol))
    ws.Range(ws.Cells(rB1, FIRST_NUM_COL), ws.Cells(rB1 + 7, lastCol)).NumberFormat = FMT_COUNT
    For i = 0 To 7
        PutLabel ws, rB1 + i, CStr(bandLbl(i))
    Next i
    ws.Rows(rB1 + 8).RowHeight = 6

    ' ---- table 3 : statistical indicators
"    PutHeaderRow ws, rH3, lastCol, ArT(""627 644 645 624 634 631 20 627 644 625 62D 635 627 626 64A""), hdrText"
    ReDim a3(1 To 8, 1 To nc)
    For col = FIRST_NUM_COL To lastCol
        j = col - FIRST_NUM_COL + 1
        c = ref(col)
        mAddr = ws.Cells(rI1 + 1, col).Address(False, False)
        sdAddr = ws.Cells(rI1 + 5, col).Address(False, False)
"        a3(1, j) = ""="" & ws.Cells(rRate, col).Address(False, False)"
"        a3(2, j) = ""="" & ws.Cells(rAvg, col).Address(False, False)"
"        a3(3, j) = Subst(""=IFERROR(AGGREGATE(17,6,{C}/(({S}={N})*({C}>0)),2),"""""""")"", ""C"", c, ""S"", sc, ""N"", secN)"
"        a3(4, j) = """""
"        a3(5, j) = ""=IFERROR("" & ws.Cells(rMax, col).Address(False, False) & ""-"" & ws.Cells(rMin, col).Address(False, False) & "","" & QT & QT & "")"""
"        a3(6, j) = Subst(""=IFERROR(SQRT(SUMPRODUCT(({S}={N})*({C}>0)*({C}-{M})^2)/COUNTIFS({S},{N},{C},"""">0"""")),"""""""")"", ""C"", c, ""S"", sc, ""N"", secN, ""M"", mAddr)"
"        a3(7, j) = ""=IFERROR("" & sdAddr & ""^2,"" & QT & QT & "")"""
"        a3(8, j) = ""=IFERROR("" & sdAddr & ""/"" & mAddr & "","" & QT & QT & "")"""
    Next col
    ws.Range(ws.Cells(rI1, FIRST_NUM_COL), ws.Cells(rI1 + 7, lastCol)).Formula = a3
    For col = FIRST_NUM_COL To lastCol
"        PutModeFormula ws.Cells(rI1 + 3, col), Subst(""=IFERROR(MODE.SNGL(IF(({S}={N})*({C}>0),{C})),"""""""")"", ""C"", ref(col), ""S"", sc, ""N"", secN)"
    Next col
    StyleGrid ws.Range(ws.Cells(rI1, FIRST_NUM_COL), ws.Cells(rI1 + 7, lastCol))
    ws.Range(ws.Cells(rI1, FIRST_NUM_COL), ws.Cells(rI1, lastCol)).NumberFormat = FMT_PCT
    ws.Range(ws.Cells(rI1 + 1, FIRST_NUM_COL), ws.Cells(rI1 + 6, lastCol)).NumberFormat = FMT_DEC
    ws.Range(ws.Cells(rI1 + 7, FIRST_NUM_COL), ws.Cells(rI1 + 7, lastCol)).NumberFormat = FMT_PCT
"    PutLabel ws, rI1, ArT(""646 633 628 629 20 627 644 646 62C 627 62D"")"
"    PutLabel ws, rI1 + 1, ArT(""627 644 645 62A 648 633 637 20 627 644 62D 633 627 628 64A"")"
"    PutLabel ws, rI1 + 2, ArT(""627 644 648 633 64A 637"")"
"    PutLabel ws, rI1 + 3, ArT(""627 644 645 646 648 627 644"")"
"    PutLabel ws, rI1 + 4, ArT(""627 644 645 62F 649"")"
"    PutLabel ws, rI1 + 5, ArT(""627 644 627 646 62D 631 627 641 20 627 644 645 639 64A 627 631 64A"")"
"    PutLabel ws, rI1 + 6, ArT(""627 644 62A 628 627 64A 646"")"
"    PutLabel ws, rI1 + 7, ArT(""645 639 627 645 644 20 627 644 627 62E 62A 644 627 641"")"
    ws.Rows(rI1 + 8).RowHeight = 6

    ' ---- five classes (mean +/- 0.5 and 1.5 standard deviations)
    ReDim a4(1 To 5, 1 To nc)
    For col = FIRST_NUM_COL To lastCol
        j = col - FIRST_NUM_COL + 1
        c = ref(col)
        mAddr = ws.Cells(rI1 + 1, col).Address(False, False)
        sdAddr = ws.Cells(rI1 + 5, col).Address(False, False)
"        a4(1, j) = Subst(""=IFERROR(COUNTIFS({S},{N},{C},"""">0"""",{C},"""">=""""&({M}+1.5*{D}),{C},""""<=20""""),"""""""")"", ""C"", c, ""S"", sc, ""N"", secN, ""M"", mAddr, ""D"", sdAddr)"
"        a4(2, j) = Subst(""=IFERROR(COUNTIFS({S},{N},{C},"""">0"""",{C},"""">=""""&({M}+0.5*{D}),{C},""""<""""&({M}+1.5*{D})),"""""""")"", ""C"", c, ""S"", sc, ""N"", secN, ""M"", mAddr, ""D"", sdAddr)"
"        a4(3, j) = Subst(""=IFERROR(COUNTIFS({S},{N},{C},"""">0"""",{C},"""">=""""&({M}-0.5*{D}),{C},""""<""""&({M}+0.5*{D})),"""""""")"", ""C"", c, ""S"", sc, ""N"", secN, ""M"", mAddr, ""D"", sdAddr)"
"        a4(4, j) = Subst(""=IFERROR(COUNTIFS({S},{N},{C},"""">0"""",{C},"""">=""""&({M}-1.5*{D}),{C},""""<""""&({M}-0.5*{D})),"""""""")"", ""C"", c, ""S"", sc, ""N"", secN, ""M"", mAddr, ""D"", sdAddr)"
"        a4(5, j) = Subst(""=IFERROR(COUNTIFS({S},{N},{C},"""">0"""",{C},""""<""""&({M}-1.5*{D})),"""""""")"", ""C"", c, ""S"", sc, ""N"", secN, ""M"", mAddr, ""D"", sdAddr)"
    Next col
    ws.Range(ws.Cells(rK1, FIRST_NUM_COL), ws.Cells(rK1 + 4, lastCol)).Formula = a4
    StyleGrid ws.Range(ws.Cells(rK1, FIRST_NUM_COL), ws.Cells(rK1 + 4, lastCol))
    ws.Range(ws.Cells(rK1, FIRST_NUM_COL), ws.Cells(rK1 + 4, lastCol)).NumberFormat = FMT_COUNT
"    PutLabel ws, rK1, ArT(""627 644 641 626 629 20 30 31"")"
"    PutLabel ws, rK1 + 1, ArT(""627 644 641 626 629 20 30 32"")"
"    PutLabel ws, rK1 + 2, ArT(""627 644 641 626 629 20 30 33"")"
"    PutLabel ws, rK1 + 3, ArT(""627 644 641 626 629 20 30 34"")"
"    PutLabel ws, rK1 + 4, ArT(""627 644 641 626 629 20 30 35"")"

    ' ---- display colours: red (white text) below the limit, green (white text) from the limit up
"    ApplyPassColors ws.Range(ws.Cells(rAvg, FIRST_NUM_COL), ws.Cells(rAvg, lastCol)), ""=10"", ""=20"""
"    ApplyPassColors ws.Range(ws.Cells(rRate, FIRST_NUM_COL), ws.Cells(rRate, lastCol)), ""=0.5"", ""=1"""
"    ApplyPassColors ws.Range(ws.Cells(rI1, FIRST_NUM_COL), ws.Cells(rI1, lastCol)), ""=0.5"", ""=1"""
"    ApplyPassColors ws.Range(ws.Cells(rI1 + 1, FIRST_NUM_COL), ws.Cells(rI1 + 1, lastCol)), ""=10"", ""=20"""

    BuildBlock = rK1 + 5 + 2
End Function

'==============================================================================
' SMALL HELPERS
'==============================================================================
' Mode needs an array formula: dynamic-array entry (Office 365) first, legacy array entry as fallback
Private Sub PutModeFormula(ByVal cel As Range, ByVal f As String)
    Dim done As Boolean
    On Error Resume Next
"    CallByName cel, ""Formula2"", VbLet, f"
    If Err.Number = 0 Then done = True
    Err.Clear
    If Not done Then
        If Len(f) <= 255 Then cel.FormulaArray = f
        Err.Clear
    End If
    On Error GoTo 0
End Sub

' Cell-value rules: text / empty cells are never coloured (text counts as greater than any number)
Private Sub ApplyPassColors(ByVal rng As Range, ByVal limitF As String, ByVal topF As String)
    Dim fc As Object
    On Error Resume Next
    rng.FormatConditions.Delete
    Set fc = rng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlLess, Formula1:=limitF)
    fc.Interior.Color = RGB(192, 0, 0)
    fc.Font.Color = RGB(255, 255, 255)
    fc.Font.Bold = True
    Set fc = rng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlBetween, Formula1:=limitF, Formula2:=topF)
    fc.Interior.Color = RGB(0, 128, 64)
    fc.Font.Color = RGB(255, 255, 255)
    fc.Font.Bold = True
    Err.Clear
End Sub

Private Sub PutStrip(ByVal ws As Worksheet, ByVal r As Long, ByVal c1 As Long, ByVal c2 As Long, _
                     ByVal cv As Long, ByVal lbl As String, ByVal fml As String)
    MergeText ws, r, c1, c2, lbl
    StyleHead ws.Range(ws.Cells(r, c1), ws.Cells(r, c2))
    With ws.Cells(r, cv)
        .Formula = fml
        .NumberFormat = FMT_COUNT
        .Font.Name = gFontNum: .Font.Size = 15: .Font.Bold = True
        .Interior.Color = RGB(221, 235, 247)
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous: .Borders.Weight = xlThin
    End With
    ws.Rows(r).RowHeight = DATA_ROW_HEIGHT
End Sub

' header row of an analysis table: merged label in A:C, subject names rotated
Private Sub PutHeaderRow(ByVal ws As Worksheet, ByVal r As Long, ByVal lastCol As Long, _
                         ByVal lbl As String, ByRef hdrText() As String)
    Dim c As Long
"    ws.Range(ws.Cells(r, FIRST_NUM_COL), ws.Cells(r, lastCol)).NumberFormat = ""@"""
    For c = FIRST_NUM_COL To lastCol
        ws.Cells(r, c).Value2 = hdrText(c)
    Next c
    StyleHead ws.Range(ws.Cells(r, 1), ws.Cells(r, lastCol))
    ws.Range(ws.Cells(r, FIRST_NUM_COL), ws.Cells(r, lastCol)).Orientation = 90
    MergeText ws, r, 1, 3, lbl
    ws.Range(ws.Cells(r, 1), ws.Cells(r, 3)).Font.Size = 16
    ws.Rows(r).RowHeight = HDR_ROW_HEIGHT
End Sub

Private Sub SetColumnWidths(ByVal ws As Worksheet, ByVal lastCol As Long)
    ws.Columns(1).ColumnWidth = ID_COL_WIDTH
    ws.Columns(2).ColumnWidth = 22
    ws.Columns(3).ColumnWidth = DATE_COL_MIN_WIDTH
    ws.Range(ws.Columns(FIRST_NUM_COL), ws.Columns(lastCol)).ColumnWidth = NUM_COL_WIDTH
End Sub

Private Sub SetupPrint(ByVal ws As Worksheet, ByVal lastRow As Long, ByVal lastCol As Long)
    On Error Resume Next
    Application.PrintCommunication = False
    With ws.PageSetup
        .PrintArea = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol)).Address
        .Orientation = xlPortrait
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .CenterHorizontally = True
    End With
    Application.PrintCommunication = True
    Err.Clear
End Sub
