"Attribute VB_Name = ""modAnalysisWorkbookBuilderCartTec03"""
Option Explicit

'==============================================================================
 Module  : modAnalysisWorkbookBuilderCartTec03   (Build_CartTec)
' Purpose : Adds the Cart_Tec sheet (one pupil card) to a workbook that already
'           has the Data and Data01 sheets. Designed for a projector (DataShow).
' Design  : - I1 = record number, linked to a spinner placed over I1.
'           - Term 1 (col D), Term 2 (col E), Term 3 (col F) come from Data.
'           - Column G is a 0.5-width white separator column.
'           - Annual Averages (col H) come from Data01.
'           - Row-wise conditional formatting on D4:F[rLast] highlights the
'             unique highest mark in green (white text) and unique lowest mark
'             in red (white text), ignoring ties and excluding column H.
'             The rules use ONLY operators and absolute cell addresses (no function
'             names, no list separators, no dependence on the active cell), so they are
'             accepted by every Excel language / regional setting.
'==============================================================================

"Private Const FMT_MARK As String = ""00.00;-00.00;;@"""
"Private Const FMT_HIDE0 As String = ""General;-General;;@"""
"Private Const FMT_DATE0 As String = ""yyyy-mm-dd;;;@"""
"Private Const FMT_RANK0 As String = ""00;-00;;@"""

Public Function BuildCartTecSheet(ByVal wb As Workbook, ByRef errText As String) As Boolean
    Dim ws As Worksheet, sp As Shape
    Dim n As Long, k As Long, r As Long, rLast As Long, rBottom As Long
    Dim accCol As Long, picCol As Long, notesCol As Long
    Dim multi As Boolean, maxV As Long
    Dim rowRef As String, yes As String, fmtExempt As String, annColHdr As String
    Dim fcs As Object, stopTxt As String, stopSign As String
    Dim clrYel As Long, clrGry As Long, clrCya As Long, clrGrn As Long, clrBan As Long, clrRow As Long
    Dim rI As Range
    Dim stepName As String, cc As Long

    BuildCartTecSheet = False
"    errText = """""
    On Error GoTo EH

"    stepName = ""discover data"""
    If Not DiscoverData(wb, errText) Then Exit Function

    clrYel = RGB(255, 255, 102)
    clrGry = RGB(235, 235, 235)
    clrCya = RGB(102, 255, 255)
    clrGrn = RGB(102, 255, 102)
    clrBan = RGB(255, 255, 0)

    n = gSubjN
    accCol = gSecCol + 2
    picCol = gSecCol + 3
    notesCol = gSecCol + 4
    multi = (gSecN > 1)
"    yes = ArT(""646 639 645"")"
    annColHdr = AnnualAvgHeader()
    stopSign = ChrW(&HD83D) & ChrW(&HDED1)
"    fmtExempt = ""00.00;-00.00;"" & QT & ArT(""625 639 641 640 640 627 621"") & QT & "";@"""
"    rowRef = "",$I$1,1)"""
    rLast = 4 + n
    rBottom = rLast
    If rBottom < 14 Then rBottom = 14

"    stepName = ""create sheet"""
    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    ws.Name = CART_SHEET
    ws.DisplayRightToLeft = True

"    stepName = ""column widths"""
    ' ---- widths (D, E, F = Terms 1..3 | G = 0.5 white separator | H = Annual)
    ws.Columns(1).ColumnWidth = 22
    ws.Columns(2).ColumnWidth = 17
    ws.Columns(3).ColumnWidth = 28
    ws.Columns(4).ColumnWidth = 12
    ws.Columns(5).ColumnWidth = 12
    ws.Columns(6).ColumnWidth = 12
    ws.Columns(7).ColumnWidth = 0.5
    ws.Columns(8).ColumnWidth = 12
    ws.Columns(9).ColumnWidth = 7.1
    ws.Columns(10).ColumnWidth = 4.2
    ws.Columns(11).ColumnWidth = 5.6

"    stepName = ""record number"""
    ' ---- record number (linked to the spinner, which is placed over this cell)
"    ws.Range(""I1"").Value = 1"
"    With ws.Range(""I1"")"
        .Font.Name = gFontNum: .Font.Size = 12: .Font.Color = RGB(128, 128, 128)
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
"        .NumberFormat = ""0"""
    End With

"    stepName = ""title"""
    ' ---- title
"    MergeText ws, 1, 1, 8, ArT(""628 637 627 642 629 20 641 646 64A 629 20 644 644 62A 644 645 64A 630"")"
"    With ws.Range(""A1:H1"")"
        .Interior.Color = clrBan
        .Font.Name = gFontBase: .Font.Size = 28: .Font.Bold = True
        .Font.Color = RGB(0, 0, 0)
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
        .ReadingOrder = xlRTL
    End With
    ws.Rows(1).RowHeight = 44

"    stepName = ""left block"""
    ' ---- left block : name, birth date, repeat, transition average, rank
"    PutCardLabel ws, ""A2"", ArT(""627 644 627 633 645 20 648 627 644 644 642 628""), clrYel"
"    PutCardValue ws, ""A3"", ""=IFERROR(INDEX("" & ColRef(gHdr(2)) & rowRef & "","" & QT & QT & "")"", FMT_HIDE0"
"    PutCardLabel ws, ""A4"", ArT(""62A 627 631 64A 62E 20 627 644 645 64A 644 627 62F""), clrYel"
"    PutCardValue ws, ""A5"", ""=IFERROR(INDEX("" & ColRef(gHdr(3)) & rowRef & "","" & QT & QT & "")"", FMT_DATE0"
"    PutCardLabel ws, ""A6"", gHdr(5), clrYel"
"    PutCardValue ws, ""A7"", ""=IFERROR(INDEX("" & ColRef(gHdr(5)) & rowRef & "","" & QT & QT & "")"", ""General"""
"    PutCardLabel ws, ""A8"", gHdr(accCol), clrYel"
"    PutCardValue ws, ""A9"", ""=IFERROR(IF(A7="" & QT & yes & QT & "","" & QT & QT & "",INDEX("" & ColRef(gHdr(accCol)) & rowRef & ""),"" & QT & QT & "")"", FMT_MARK"
"    PutCardLabel ws, ""A10"", ArT(""627 644 631 62A 628 629 20 641 64A 20 627 644 642 633 645""), clrYel"
"    PutCardValue ws, ""A11"", ""=IFERROR(INDEX("" & ColRef(gHdr(1)) & rowRef & "","" & QT & QT & "")"", FMT_RANK0"

"    stepName = ""picture and notes"""
    ' ---- picture (in-cell picture typed in Data) and notes
"    ws.Range(""B2:B7"").Merge"
    If multi Then
"        ws.Range(""B2"").Formula = ""=IFERROR(IF($A$12<>"" & QT & QT & "","" & QT & QT & "",INDEX("" & ColRef(gHdr(picCol)) & rowRef & ""),"" & QT & QT & "")"""
    Else
"        ws.Range(""B2"").Formula = ""=IFERROR(INDEX("" & ColRef(gHdr(picCol)) & rowRef & "","" & QT & QT & "")"""
    End If
"    With ws.Range(""B2:B7"")"
        .NumberFormat = FMT_HIDE0
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous: .Borders.Weight = xlThin
    End With
"    ws.Range(""B8:B11"").Merge"
"    ws.Range(""B8"").Formula = ""=IFERROR(INDEX("" & ColRef(gHdr(notesCol)) & rowRef & "","" & QT & QT & "")"""
"    With ws.Range(""B8:B11"")"
        .NumberFormat = FMT_HIDE0
        .Font.Name = gFontName: .Font.Size = 16: .Font.Bold = True
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
        .WrapText = True
        .ReadingOrder = xlRTL
        .Borders.LineStyle = xlContinuous: .Borders.Weight = xlThin
    End With

"    stepName = ""class title"""
    ' ---- class title (and, for several sections, the section being browsed: typed by hand)
    If multi Then
"        ws.Range(""C2"").Value2 = gClassTitle"
"        With ws.Range(""C2"")"
            .Interior.Color = clrYel
            .Font.Name = gFontBase: .Font.Size = 20: .Font.Bold = True
            .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
            .ShrinkToFit = True
            .ReadingOrder = xlRTL
            .Borders.LineStyle = xlContinuous: .Borders.Weight = xlThin
        End With
"        With ws.Range(""C3"")"
            .NumberFormat = FMT_COUNT
            .Interior.Color = clrGry
            .Font.Name = gFontNum: .Font.Size = 22: .Font.Bold = True
            .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
            .Borders.LineStyle = xlContinuous: .Borders.Weight = xlThin
        End With
    Else
"        ws.Range(""C2:C3"").Merge"
"        ws.Range(""C2"").Value2 = gClassTitle"
"        With ws.Range(""C2:C3"")"
            .Interior.Color = clrYel
            .Font.Name = gFontBase: .Font.Size = 20: .Font.Bold = True
            .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
            .WrapText = True
            .ReadingOrder = xlRTL
            .Borders.LineStyle = xlContinuous: .Borders.Weight = xlThin
        End With
    End If

"    stepName = ""results headers"""
    ' ---- results block : Term 1 (D) | Term 2 (E) | Term 3 (F) | White Sep (G) | Annual Averages (H)
"    ws.Range(""D2:D3"").Merge"
"    ws.Range(""D2"").Value2 = ArT(""646 62A 627 626 62C 20 627 644 641 635 644 20 627 644 623 648 644"")"
"    ws.Range(""E2:E3"").Merge"
"    ws.Range(""E2"").Value2 = ArT(""646 62A 627 626 62C 20 627 644 641 635 644 20 627 644 62B 627 646 64A"")"
"    ws.Range(""F2:F3"").Merge"
"    ws.Range(""F2"").Value2 = ArT(""646 62A 627 626 62C 20 627 644 641 635 644 20 627 644 62B 627 644 62B"")"
"    ws.Range(""H2:H3"").Merge"
"    ws.Range(""H2"").Value2 = ArT(""627 644 645 639 62F 644 627 62A 20 627 644 633 646 648 64A 629"")"

"    StyleHead ws.Range(""D2:D3"")"
"    StyleHead ws.Range(""E2:E3"")"
"    StyleHead ws.Range(""F2:F3"")"
"    StyleHead ws.Range(""H2:H3"")"
"    For Each rI In ws.Range(""D2:D3,E2:E3,F2:F3"").Areas"
        rI.Interior.Color = clrCya
        rI.Font.Color = RGB(0, 0, 0)
        rI.Font.Size = 15
    Next rI
"    With ws.Range(""H2:H3"")"
        .Interior.Color = clrGrn
        .Font.Color = RGB(0, 0, 0)
        .Font.Size = 15
    End With

"    stepName = ""term averages and subject rows"""
    ' ---- averages (row 4) and one row per subject
"    ws.Range(""C4"").Value2 = ArT(""627 644 645 639 62F 644 20 627 644 641 635 644 64A"")"
"    ws.Range(""D4"").Formula = ""=IFERROR(INDEX("" & ColRef(gHdr(gAvgC1)) & rowRef & "","" & QT & QT & "")"""
"    ws.Range(""E4"").Formula = ""=IFERROR(INDEX("" & ColRef(gHdr(gAvgC2)) & rowRef & "","" & QT & QT & "")"""
"    ws.Range(""F4"").Formula = ""=IFERROR(INDEX("" & ColRef(gHdr(gAvgC3)) & rowRef & "","" & QT & QT & "")"""
"    ws.Range(""H4"").Formula = ""=IFERROR(INDEX("" & ColRef01(annColHdr) & rowRef & "","" & QT & QT & "")"""

    For k = 1 To n
        r = 4 + k
        ws.Cells(r, 3).Value2 = gSubjName(k)
"        ws.Cells(r, 4).Formula = ""=IFERROR(INDEX("" & ColRef(gHdr(gSubjC1(k))) & rowRef & "","" & QT & QT & "")"""
"        ws.Cells(r, 5).Formula = ""=IFERROR(INDEX("" & ColRef(gHdr(gSubjC2(k))) & rowRef & "","" & QT & QT & "")"""
"        ws.Cells(r, 6).Formula = ""=IFERROR(INDEX("" & ColRef(gHdr(gSubjC3(k))) & rowRef & "","" & QT & QT & "")"""
"        ws.Cells(r, 8).Formula = ""=IFERROR(INDEX("" & ColRef01(gSubjName(k)) & rowRef & "","" & QT & QT & "")"""
    Next k

"    stepName = ""styles and number formats"""
    StyleGrid ws.Range(ws.Cells(4, 3), ws.Cells(rLast, 3))
    StyleGrid ws.Range(ws.Cells(4, 4), ws.Cells(rLast, 6))
    StyleGrid ws.Range(ws.Cells(4, 8), ws.Cells(rLast, 8))
    With ws.Range(ws.Cells(4, 3), ws.Cells(rLast, 3))
        .Font.Name = gFontName: .Font.Size = 17: .Font.Bold = True
    End With
    ws.Range(ws.Cells(4, 4), ws.Cells(rLast, 6)).Font.Size = 20
    ws.Range(ws.Cells(4, 8), ws.Cells(rLast, 8)).Font.Size = 20
    ws.Range(ws.Cells(4, 8), ws.Cells(rLast, 8)).Font.Bold = True
    ws.Range(ws.Cells(4, 4), ws.Cells(rLast, 6)).NumberFormat = FMT_MARK
    ws.Range(ws.Cells(4, 8), ws.Cells(rLast, 8)).NumberFormat = FMT_MARK
    If gPEIdx > 0 Then
        ws.Range(ws.Cells(4 + gPEIdx, 4), ws.Cells(4 + gPEIdx, 6)).NumberFormat = fmtExempt
        ws.Cells(4 + gPEIdx, 8).NumberFormat = fmtExempt
    End If

"    stepName = ""row colours"""
"    ' ---- row colours: row 4 = cyan (H4 = light green); subject rows = alternating yellow / grey; col G = white"
"    For Each rI In ws.Range(""C4,D4,E4,F4"").Areas"
        rI.Interior.Color = clrCya
        rI.Font.Bold = True
    Next rI
"    With ws.Range(""H4"")"
        .Interior.Color = clrGrn
        .Font.Bold = True
    End With
    For k = 1 To n
        r = 4 + k
        If k Mod 2 = 1 Then clrRow = clrYel Else clrRow = clrGry
        ws.Cells(r, 3).Interior.Color = clrRow
        ws.Cells(r, 4).Interior.Color = clrRow
        ws.Cells(r, 5).Interior.Color = clrRow
        ws.Cells(r, 6).Interior.Color = clrRow
        ws.Cells(r, 8).Interior.Color = clrRow
    Next k
    ws.Range(ws.Cells(2, 7), ws.Cells(rBottom, 7)).Interior.Color = RGB(255, 255, 255)

"    stepName = ""row heights"""
    ' ---- same row height everywhere (the picture spans six rows)
    For r = 2 To rBottom
        ws.Rows(r).RowHeight = 32
    Next r

"    stepName = ""end-of-section message"""
    ' ---- end-of-section message (several sections only)
    '   shows only when C3 (typed by hand) holds a section number and the selected pupil is in another one
    If multi Then
"        stopTxt = stopSign & "" "" & ArT(""646 640 647 640 627 64A 640 629 20 627 644 640 642 640 633 640 645"") & "" "" & stopSign"
"        ws.Range(""A12:B14"").Merge"
"        ws.Range(""A12"").Formula = ""=IFERROR(IF($C$3="" & QT & QT & "","" & QT & QT & "",IF(INDEX("" & ColRef(gHdr(gSecCol)) & rowRef & _"
"                                  ""<>$C$3,"" & QT & stopTxt & QT & "","" & QT & QT & "")),"" & QT & QT & "")"""
"        With ws.Range(""A12:B14"")"
            .Font.Name = gFontBase: .Font.Size = 26: .Font.Bold = True
            .Font.Color = RGB(255, 0, 0)
            .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
            .WrapText = False
            .ShrinkToFit = True
            .ReadingOrder = xlRTL
        End With
    End If

"    stepName = ""section message formats"""
    ' ---- conditional formats
    If multi Then
        ' bright yellow band behind the message
"        Set fcs = ws.Range(""A12:B14"").FormatConditions.Add(Type:=xlExpression, Formula1:=""=$A$12<>"" & QT & QT)"
        fcs.Interior.Color = RGB(255, 255, 0)
        fcs.Font.Color = RGB(255, 0, 0)
        ' hide the data of every other cell while the message is shown
"        Set fcs = ws.Range(""A3,A5,A7,A9,A11,B8:B11,D4:F"" & rLast & "",H4:H"" & rLast).FormatConditions.Add( _"
"                  Type:=xlExpression, Formula1:=""=$A$12<>"" & QT & QT)"
"        fcs.NumberFormat = "";;;"""
        fcs.Font.Color = RGB(255, 255, 255)
        fcs.StopIfTrue = True
    End If

    ' Row-wise comparison across Term 1, Term 2, Term 3 (D:F, from row 4 = general term average
"    ' down to the last subject). Each row is compared on its own; column H is never compared."
    '   - unique highest positive mark : green background + white bold text
    '   - unique lowest  positive mark : red background + white bold text
"    '   - two equal marks get no format; the third (different) mark gets it"
    '   - three equal marks, or fewer than two positive marks : no format
    '   - empty, zero and non-numeric cells are ignored (no error)
"    stepName = ""row comparison formats"""
    For r = 4 To rLast
        For cc = 4 To 6
            AddTermRankFormat ws, r, cc
        Next cc
    Next r

"    stepName = ""spinner"""
    ' ---- spinner linked to I1, placed OVER I1 (it hides the number below it)
    maxV = gRows
    If maxV < 1 Then maxV = 1
    If maxV > 30000 Then maxV = 30000
"    Set sp = ws.Shapes.AddFormControl(xlSpinner, ws.Range(""I1"").Left + 2, ws.Range(""I1"").Top + 2, _"
"                                      ws.Range(""I1"").Width - 4, ws.Range(""I1"").Height + ws.Range(""I2"").Height - 4)"
"    sp.Name = ""Spinner 1"""
    With sp.ControlFormat
"        .LinkedCell = ""$I$1"""
        .Min = 1
        .Max = maxV
        .SmallChange = 1
        .Value = 1
    End With

"    stepName = ""window"""
    ' ---- window
    On Error Resume Next
    ws.Activate
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.Zoom = 90
"    ws.Range(""A1"").Select"
    Err.Clear
    On Error GoTo EH

    BuildCartTecSheet = True
    Exit Function
EH:
"    errText = ""BuildCartTecSheet ["" & stepName & ""]: "" & Err.Number & "" : "" & Err.Description"
    Err.Clear
End Function

' ------------------------------------------------------------------------------
' Adds the two comparison rules to ONE cell (r, c) of D:F (Terms 1..3):
'   green = this mark is the unique highest of the row, red = the unique lowest.
"' pos(x) = (x>0)*(x<1000) is 1 only for a positive number: an empty text """" or any text"
' is greater than every number in Excel, so (x<1000) is False for it (0 and blanks fail
' (x>0)). Only operators and absolute addresses are used - no function, no separator.
' ------------------------------------------------------------------------------
Private Sub AddTermRankFormat(ByVal ws As Worksheet, ByVal r As Long, ByVal c As Long)
    Dim a As String, b As String, d As String
    Dim pa As String, pb As String, pd As String
    Dim o1 As Long, o2 As Long, k As Long
    Dim fcs As Object

    ' the two other term cells of the same row
    For k = 4 To 6
        If k <> c Then
            If o1 = 0 Then o1 = k Else o2 = k
        End If
    Next k

    a = ws.Cells(r, c).Address(True, True)
    b = ws.Cells(r, o1).Address(True, True)
    d = ws.Cells(r, o2).Address(True, True)
"    pa = ""("" & a & "">0)*("" & a & ""<1000)"""
"    pb = ""("" & b & "">0)*("" & b & ""<1000)"""
"    pd = ""("" & d & "">0)*("" & d & ""<1000)"""

    ' unique highest: at least one other positive mark, and no other positive mark >= this one
    Set fcs = ws.Cells(r, c).FormatConditions.Add(Type:=xlExpression, _
"              Formula1:=""="" & pa & ""*("" & pb & ""+"" & pd & "">=1)*(1-"" & pb & ""*("" & b & "">="" & a & ""))*(1-"" & pd & ""*("" & d & "">="" & a & ""))"")"
    fcs.Interior.Color = RGB(0, 128, 64)
    fcs.Font.Color = RGB(255, 255, 255)
    fcs.Font.Bold = True

    ' unique lowest: at least one other positive mark, and no other positive mark <= this one
    Set fcs = ws.Cells(r, c).FormatConditions.Add(Type:=xlExpression, _
"              Formula1:=""="" & pa & ""*("" & pb & ""+"" & pd & "">=1)*(1-"" & pb & ""*("" & b & ""<="" & a & ""))*(1-"" & pd & ""*("" & d & ""<="" & a & ""))"")"
    fcs.Interior.Color = RGB(192, 0, 0)
    fcs.Font.Color = RGB(255, 255, 255)
    fcs.Font.Bold = True
End Sub

Private Sub PutCardLabel(ByVal ws As Worksheet, ByVal addr As String, ByVal txt As String, ByVal clr As Long)
    ws.Range(addr).Value2 = txt
    With ws.Range(addr)
        .Interior.Color = clr
        .Font.Name = gFontName: .Font.Size = 17: .Font.Bold = True
        .Font.Color = RGB(0, 0, 0)
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
        .ShrinkToFit = True
        .ReadingOrder = xlRTL
        .Borders.LineStyle = xlContinuous: .Borders.Weight = xlThin
    End With
End Sub

Private Sub PutCardValue(ByVal ws As Worksheet, ByVal addr As String, ByVal f As String, ByVal fmt As String)
    With ws.Range(addr)
        .Formula = f
        .NumberFormat = fmt
        .Font.Name = gFontName: .Font.Size = 17: .Font.Bold = True
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
        .ShrinkToFit = True
        .ReadingOrder = xlRTL
        .Borders.LineStyle = xlContinuous: .Borders.Weight = xlThin
    End With
End Sub
