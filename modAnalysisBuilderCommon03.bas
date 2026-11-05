Attribute VB_Name = "modAnalysisBuilderCommon03"
Option Explicit

'==============================================================================
' Module  : modAnalysisBuilderCommon03
' Purpose : Shared constants, text helpers, fonts, cell styles and the
'           "discover what is inside the Data sheet" routine.
'==============================================================================

' ---- sheet / table names
Public Const DATA_SHEET As String = "Data"
Public Const DATA_TABLE As String = "Data"
Public Const DATA01_SHEET As String = "Data01"
Public Const DATA01_TABLE As String = "Data01"
Public Const DATA_HDR_ROW As Long = 6
Public Const DATA_FIRST_ROW As Long = 7
Public Const ANALYSIS_T3_SHEET As String = "Analysis03"
Public Const ANALYSIS_ANN_SHEET As String = "AnnualAnalysis"
Public Const CART_SHEET As String = "Cart_Tec"

' ---- number formats
Public Const FMT_COUNT As String = "00"
Public Const FMT_DEC As String = "00.00"
Public Const FMT_PCT As String = "[=1]%0;%00.00"
Public Const PASS_MARK As Double = 10

' ---- layout
Public Const FIRST_NUM_COL As Long = 4
Public Const TITLE_ROW_HEIGHT As Double = 40
Public Const DATA_ROW_HEIGHT As Double = 24
Public Const HDR_ROW_HEIGHT As Double = 100
Public Const ID_COL_WIDTH As Double = 4.3
Public Const NUM_COL_WIDTH As Double = 6.3
Public Const DATE_COL_MIN_WIDTH As Double = 10.5

' ---- fonts
Public Const CHECK_FONTS As Boolean = True
Public Const FONT_BASE As String = "Dubai"
Public Const FONT_NAME As String = "Dubai Medium"
Public Const FONT_NUM As String = "Aptos Slab"
Public Const FONT_SAFE_TEXT As String = "Arial"
Public Const FONT_SAFE_NUM As String = "Calibri"

Public Const QT As String = """"

Public gFontBase As String
Public gFontName As String
Public gFontNum As String
Public gFontNotes As String

' ---- result of DiscoverData
Public gHdr() As String
Public gNCols As Long
Public gRows As Long
Public gSecCol As Long
Public gSubjN As Long
Public gSubjName() As String
Public gSubjC1() As Long
Public gSubjC2() As Long
Public gSubjC3() As Long
Public gAvgC1 As Long
Public gAvgC2 As Long
Public gAvgC3 As Long
Public gAvgAnn As Long
Public gPEIdx As Long
Public gSecN As Long
Public gSecList() As Long
Public gClassTitle As String

'==============================================================================
' TEXT HELPERS
'==============================================================================
Public Function ArT(ByVal codes As String) As String
    Dim parts() As String, i As Long, s As String
    parts = Split(codes, " ")
    For i = LBound(parts) To UBound(parts)
        If Len(parts(i)) > 0 Then s = s & ChrW(CLng("&H" & parts(i)))
    Next i
    ArT = s
End Function

Public Function NormKey(ByVal s As String) As String
    Dim c As Variant
    s = Replace(s, ChrW(160), " ")
    s = Replace(s, vbTab, " ")
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")
    For Each c In Array(&H200B, &H200E, &H200F, &H202A, &H202B, &H202C, &H640)
        s = Replace(s, ChrW(CLng(c)), "")
    Next c
    Do While InStr(s, "  ") > 0: s = Replace(s, "  ", " "): Loop
    NormKey = Trim$(s)
End Function

Public Function NormAr(ByVal s As String) As String
    s = NormKey(s)
    s = Replace(s, ChrW(&H623), ChrW(&H627))
    s = Replace(s, ChrW(&H625), ChrW(&H627))
    s = Replace(s, ChrW(&H622), ChrW(&H627))
    s = Replace(s, ChrW(&H671), ChrW(&H627))
    s = Replace(s, ChrW(&H649), ChrW(&H64A))
    NormAr = s
End Function

Public Function ColRef(ByVal hdr As String) As String
    Dim t As String
    t = hdr
    t = Replace(t, "'", "''")
    t = Replace(t, "[", "'[")
    t = Replace(t, "]", "']")
    t = Replace(t, "#", "'#")
    ColRef = DATA_TABLE & "[[" & t & "]]"
End Function

Public Function ColRef01(ByVal hdr As String) As String
    Dim t As String
    t = hdr
    t = Replace(t, "'", "''")
    t = Replace(t, "[", "'[")
    t = Replace(t, "]", "']")
    t = Replace(t, "#", "'#")
    ColRef01 = DATA01_TABLE & "[[" & t & "]]"
End Function

Public Function AnnualAvgHeader() As String
    AnnualAvgHeader = ArT("627 644 645 639 62F 644 20 627 644 633 646 648 64A")
End Function

Public Function ColLetter(ByVal c As Long) As String
    Dim a As String
    a = Cells(1, c).Address(False, False)
    ColLetter = Left$(a, Len(a) - 1)
End Function

Public Function Subst(ByVal f As String, ParamArray kv() As Variant) As String
    Dim i As Long
    For i = LBound(kv) To UBound(kv) - 1 Step 2
        f = Replace(f, "{" & CStr(kv(i)) & "}", CStr(kv(i + 1)))
    Next i
    Subst = f
End Function

'==============================================================================
' FILES
'==============================================================================
Public Function FolderExistsP(ByVal p As String) As Boolean
    Dim fso As Object
    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Right$(p, 1) = "\" Then p = Left$(p, Len(p) - 1)
    FolderExistsP = fso.FolderExists(p)
    If Err.Number <> 0 Then FolderExistsP = False: Err.Clear
End Function

Public Function FileExistsP(ByVal p As String) As Boolean
    Dim fso As Object
    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    FileExistsP = fso.FileExists(p)
    If Err.Number <> 0 Then FileExistsP = False: Err.Clear
End Function

Public Function EnsureFolderP(ByVal p As String) As Boolean
    Dim fso As Object
    If Right$(p, 1) = "\" Then p = Left$(p, Len(p) - 1)
    If Not FolderExistsP(p) Then
        On Error Resume Next
        Set fso = CreateObject("Scripting.FileSystemObject")
        fso.CreateFolder p
        Err.Clear
        On Error GoTo 0
    End If
    EnsureFolderP = FolderExistsP(p)
End Function

Public Function IsOpenWb(ByVal fullPath As String) As Boolean
    Dim wb As Workbook
    For Each wb In Application.Workbooks
        If StrComp(wb.FullName, fullPath, vbTextCompare) = 0 Then
            IsOpenWb = True
            Exit Function
        End If
    Next wb
End Function

Public Function NextFreePath(ByVal folder As String, ByVal stem As String, ByVal ext As String) As String
    Dim p As String, n As Long
    p = folder & stem & ext
    Do While FileExistsP(p) Or IsOpenWb(p)
        n = n + 1
        p = folder & stem & "_" & Format$(n, "00") & ext
    Loop
    NextFreePath = p
End Function

'==============================================================================
' FONTS
'==============================================================================
Public Sub ResolveFonts(ByVal ws As Worksheet)
    gFontNotes = ""
    gFontBase = PickFont(ws, FONT_BASE, FONT_SAFE_TEXT)
    gFontName = PickFont(ws, FONT_NAME, FONT_BASE, FONT_SAFE_TEXT)
    gFontNum = PickFont(ws, FONT_NUM, FONT_SAFE_NUM, FONT_SAFE_TEXT)
    ws.Cells.Clear
End Sub

Public Sub EnsureFonts()
    Dim wbTmp As Workbook
    If gFontBase <> "" And gFontName <> "" And gFontNum <> "" Then Exit Sub
    Set wbTmp = Workbooks.Add(xlWBATWorksheet)
    ResolveFonts wbTmp.Worksheets(1)
    wbTmp.Close SaveChanges:=False
End Sub

Public Function OutputFolderName() As String
    OutputFolderName = ArT("62A 62D 644 64A 644 20 646 62A 627 626 62C 20 627 644 641 635 644 20 627 644 62B 627 644 62B")
End Function

Private Function PickFont(ByVal ws As Worksheet, ParamArray names() As Variant) As String
    Dim i As Long, first As String, chosen As String
    first = CStr(names(LBound(names)))
    If Not CHECK_FONTS Then
        PickFont = first
        Exit Function
    End If
    chosen = CStr(names(UBound(names)))
    For i = LBound(names) To UBound(names) - 1
        If FontAvailable(ws, CStr(names(i))) Then
            chosen = CStr(names(i))
            Exit For
        End If
    Next i
    PickFont = chosen
    If StrComp(chosen, first, vbTextCompare) <> 0 Then
        If InStr(gFontNotes, first & " ->") = 0 Then gFontNotes = gFontNotes & vbLf & first & " -> " & chosen
    End If
End Function

Private Function FontAvailable(ByVal ws As Worksheet, ByVal fontName As String) As Boolean
    Dim samples(1 To 2) As String, i As Long, wBase As Double, wTest As Double
    samples(1) = "MWmwIl0123456789 The quick brown fox " & ArT("627 644 639 631 628 64A 629 20 645 62F 631 633 629")
    samples(2) = "00.00 18.75 05.50 Notef11 MoyenneMe"
    On Error GoTo Unknown
    For i = 1 To 2
        wBase = MeasureWidth(ws, "Zz No Such Font Zz", samples(i))
        wTest = MeasureWidth(ws, fontName, samples(i))
        If Abs(wTest - wBase) > 0.3 Then
            FontAvailable = True
            Exit Function
        End If
    Next i
    Exit Function
Unknown:
    FontAvailable = True
End Function

Private Function MeasureWidth(ByVal ws As Worksheet, ByVal fontName As String, ByVal sample As String) As Double
    With ws.Range("A1")
        .Value2 = sample
        .Font.Name = fontName
        .Font.Size = 14
        .WrapText = False
    End With
    ws.Columns(1).AutoFit
    MeasureWidth = ws.Columns(1).ColumnWidth
End Function

'==============================================================================
' CELL STYLES
'==============================================================================
Public Sub StyleBanner(ByVal rng As Range)
    With rng
        .Font.Name = gFontBase: .Font.Size = 16: .Font.Bold = True
        .Interior.Color = RGB(28, 55, 103): .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
        .WrapText = True
        .ReadingOrder = xlRTL
    End With
End Sub

Public Sub StyleTableTitle(ByVal rng As Range)
    With rng
        .Font.Name = gFontName: .Font.Size = 16: .Font.Bold = True
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
        .ReadingOrder = xlRTL
        .Borders.LineStyle = xlContinuous: .Borders.Weight = xlThin
    End With
End Sub

Public Sub StyleGrid(ByVal rng As Range)
    With rng
        .Font.Name = gFontNum: .Font.Size = 15: .Font.Bold = False
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
        .ReadingOrder = xlRTL
        .WrapText = False
        .Borders.LineStyle = xlContinuous: .Borders.Weight = xlThin
        .Rows.RowHeight = DATA_ROW_HEIGHT
    End With
End Sub

Public Sub StyleHead(ByVal rng As Range)
    With rng
        .Interior.Color = RGB(28, 55, 103)
        .Font.Name = gFontBase: .Font.Size = 14: .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
        .ReadingOrder = xlRTL
        .WrapText = True
        .Borders.LineStyle = xlContinuous: .Borders.Weight = xlThin
    End With
End Sub

Public Sub StyleLabel(ByVal rng As Range)
    With rng
        .Interior.Color = RGB(221, 235, 247)
        .Font.Name = gFontName: .Font.Size = 14: .Font.Bold = True
        .Font.Color = RGB(0, 0, 0)
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
        .ReadingOrder = xlRTL
        .WrapText = False
        .ShrinkToFit = True
        .Borders.LineStyle = xlContinuous: .Borders.Weight = xlThin
    End With
End Sub

Public Sub MergeText(ByVal ws As Worksheet, ByVal r As Long, ByVal c1 As Long, ByVal c2 As Long, ByVal txt As String)
    With ws.Range(ws.Cells(r, c1), ws.Cells(r, c2))
        If c2 > c1 Then .Merge
        .Value2 = txt
    End With
End Sub

Public Sub PutLabel(ByVal ws As Worksheet, ByVal r As Long, ByVal txt As String)
    MergeText ws, r, 1, 3, txt
    StyleLabel ws.Range(ws.Cells(r, 1), ws.Cells(r, 3))
End Sub

'==============================================================================
' DISCOVER the Data sheet
'==============================================================================
Private Function HasMarks(ByVal lo As ListObject, ByVal c1 As Long, ByVal c2 As Long, ByVal c3 As Long) As Boolean
    Dim n As Double
    On Error Resume Next
    n = Application.WorksheetFunction.CountIf(lo.ListColumns(c1).DataBodyRange, ">0") + _
        Application.WorksheetFunction.CountIf(lo.ListColumns(c2).DataBodyRange, ">0") + _
        Application.WorksheetFunction.CountIf(lo.ListColumns(c3).DataBodyRange, ">0")
    HasMarks = (n > 0)
End Function

Public Function DiscoverData(ByVal wb As Workbook, ByRef errText As String) As Boolean
    Dim ws As Worksheet, lo As ListObject
    Dim c As Long, k As Long, c2 As Long, c3 As Long, mx As Long
    Dim h As String, b As String, sfx1 As String, sfx2 As String, sfx3 As String, mk As String
    Dim rngSec As Range

    DiscoverData = False: errText = ""
    On Error GoTo EH

    Set ws = wb.Worksheets(DATA_SHEET)
    Set lo = ws.ListObjects(DATA_TABLE)
    gNCols = lo.ListColumns.Count
    gRows = lo.ListRows.Count
    If gRows < 1 Then errText = "Data table is empty": Exit Function

    ReDim gHdr(1 To gNCols)
    gSecCol = 0
    For c = 1 To gNCols
        gHdr(c) = CStr(lo.ListColumns(c).Name)
        If gSecCol = 0 Then
            If StrComp(NormKey(gHdr(c)), ArT("627 644 642 633 645"), vbTextCompare) = 0 Then gSecCol = c
        End If
    Next c
    If gSecCol < 7 Then errText = "Section column not found in Data": Exit Function
    If gSecCol + 4 > gNCols Then errText = "Data table is missing its extra columns": Exit Function

    sfx1 = " " & ChrW(&H641) & " 1"
    sfx2 = " " & ChrW(&H641) & " 2"
    sfx3 = " " & ChrW(&H641) & " 3"
    mk = ArT("645 639 62F 644")

    gSubjN = 0: gAvgC1 = 0: gAvgC2 = 0: gAvgC3 = 0: gPEIdx = 0
    gAvgAnn = gSecCol + 1
    ReDim gSubjName(1 To gSecCol)
    ReDim gSubjC1(1 To gSecCol)
    ReDim gSubjC2(1 To gSecCol)
    ReDim gSubjC3(1 To gSecCol)

    For c = 6 To gSecCol - 1
        h = NormKey(gHdr(c))
        If Len(h) > 4 Then
            If Right$(h, 4) = sfx1 Then
                b = Left$(h, Len(h) - 4)
                c2 = 0: c3 = 0
                For k = c + 1 To gSecCol - 1
                    If c2 = 0 And StrComp(NormKey(gHdr(k)), b & sfx2, vbTextCompare) = 0 Then
                        c2 = k
                    ElseIf c3 = 0 And StrComp(NormKey(gHdr(k)), b & sfx3, vbTextCompare) = 0 Then
                        c3 = k
                    End If
                Next k
                If c2 > 0 And c3 > 0 Then
                    If HasMarks(lo, c, c2, c3) Then
                        gSubjN = gSubjN + 1
                        gSubjName(gSubjN) = b
                        gSubjC1(gSubjN) = c
                        gSubjC2(gSubjN) = c2
                        gSubjC3(gSubjN) = c3
                        If InStr(NormAr(b), ArT("627 644 628 62F 646 64A 629")) > 0 Then gPEIdx = gSubjN
                    End If
                End If
            ElseIf Left$(h, 4) = mk Then
                If Right$(h, 1) = "1" Then gAvgC1 = c
                If Right$(h, 1) = "2" Then gAvgC2 = c
                If Right$(h, 1) = "3" Then gAvgC3 = c
            End If
        End If
    Next c

    If gSubjN < 1 Then errText = "No subject with marks was found in Data": Exit Function
    If gAvgC1 = 0 Or gAvgC2 = 0 Or gAvgC3 = 0 Then errText = "Term average columns (1, 2, 3) were not found in Data": Exit Function

    Set rngSec = lo.ListColumns(gSecCol).DataBodyRange
    mx = CLng(Application.WorksheetFunction.Max(rngSec))
    If mx < 1 Then errText = "Section numbers are missing in Data": Exit Function
    If mx > 60 Then errText = "Section number out of range": Exit Function
    gSecN = 0
    ReDim gSecList(1 To mx)
    For k = 1 To mx
        If Application.WorksheetFunction.CountIf(rngSec, k) > 0 Then
            gSecN = gSecN + 1
            gSecList(gSecN) = k
        End If
    Next k

    gClassTitle = CStr(ws.Range("A5").Value)
    DiscoverData = True
    Exit Function
EH:
    errText = "DiscoverData: " & Err.Number & " : " & Err.Description
    Err.Clear
End Function
