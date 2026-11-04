"Attribute VB_Name = ""modAnalysisWorkbookBuilderPanel"""
Option Explicit

'==============================================================================
 Module  : modAnalysisWorkbookBuilderPanel
' Purpose : Designs the control panel of the tool: a banner, the source / output
"'           folders, the big button ""Term 03 Analysis"" (runs Build_All), a button"
'           that opens the output folder and one status line per file.
' Use     : run Panel_Create ONCE (it can be run again to redraw the panel).
'           Build_All then keeps the status lines up to date.
'==============================================================================

"Private Const PANEL_SHEET As String = ""Term03Panel"""
"Private Const BTN_MAIN As String = ""btnTerm03Analysis"""
"Private Const BTN_OPEN As String = ""btnOpenOutput"""
Private Const ROW_FIRST As Long = 16

Private Const CLR_NAVY As Long = 6764316      ' RGB(28, 55, 103)

Public Sub Panel_Create()
    Dim ws As Worksheet, map As Object, k As Variant, i As Long, n As Long
    Dim shp As Shape, rngBtn As Range, base As String

    On Error GoTo EH
    Application.ScreenUpdating = False
    EnsureFonts

    Set ws = PanelWs()
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Worksheets(1))
        ws.Name = PANEL_SHEET
    Else
        Do While ws.Shapes.Count > 0
            ws.Shapes(1).Delete
        Loop
        ws.Cells.UnMerge
        ws.Cells.Clear
    End If

    ws.DisplayRightToLeft = True
    With ws.Cells
        .Font.Name = gFontBase
        .Font.Size = 12
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(255, 255, 255)
    End With

    ws.Columns(1).ColumnWidth = 2.5
    ws.Columns(2).ColumnWidth = 5
    ws.Columns(3).ColumnWidth = 34
    ws.Columns(4).ColumnWidth = 12
    ws.Columns(5).ColumnWidth = 20
    ws.Columns(6).ColumnWidth = 62
    ws.Columns(7).ColumnWidth = 11
    ws.Columns(8).ColumnWidth = 2.5

"    ' ---- banner: ""تحليل نتائج الفصل الثالث"""
    ws.Rows(1).RowHeight = 10
"    MergeText ws, 2, 2, 7, ArT(""62A 62D 644 64A 644 20 646 62A 627 626 62C 20 627 644 641 635 644 20 627 644 62B 627 644 62B"")"
"    With ws.Range(""B2:G2"")"
        .Interior.Color = RGB(28, 55, 103)
        .Font.Name = gFontBase: .Font.Size = 28: .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
        .ReadingOrder = xlRTL
    End With
    ws.Rows(2).RowHeight = 58

"    ' ---- subtitle: ""إصدار ملفات التحليل (Data - Data01 - Analysis01 - AnnualAnalysis - Cart_Tec) لكل الشعب بضغطة واحدة"""
"    MergeText ws, 3, 2, 7, ArT(""625 635 62F 627 631 20 645 644 641 627 62A 20 627 644 62A 62D 644 64A 644 20"") & _"
"                           ""("" & DATA_SHEET & ""  -  "" & DATA01_SHEET & ""  -  "" & ANALYSIS_T3_SHEET & ""  -  "" & ANALYSIS_ANN_SHEET & ""  -  "" & CART_SHEET & "") "" & _"
"                           ArT(""644 643 644 20 627 644 634 639 628 20 628 636 63A 637 629 20 648 627 62D 62F 629"")"
"    With ws.Range(""B3:G3"")"
        .Interior.Color = RGB(221, 235, 247)
        .Font.Name = gFontName: .Font.Size = 14
        .Font.Color = RGB(28, 55, 103)
        .HorizontalAlignment = xlCenter
        .ReadingOrder = xlRTL
    End With
    ws.Rows(3).RowHeight = 32
    ws.Rows(4).RowHeight = 12

    ' ---- information block
    base = ThisWorkbook.Path
"    PutInfo ws, 5, ArT(""645 62C 644 62F 20 627 644 645 635 62F 631""), base & ""\Source03"""
"    PutInfo ws, 6, ArT(""645 62C 644 62F 20 627 644 625 635 62F 627 631""), base & ""\"" & OutputFolderName()"
"    PutInfo ws, 7, ArT(""622 62E 631 20 62A 634 63A 64A 644""), ""-"""
"    PutInfo ws, 8, ArT(""627 644 646 62A 64A 62C 629""), ""-"""
    ws.Rows(9).RowHeight = 14

    ' ---- buttons (placed on cells, same way as the spinner of Cart_Tec)
    For i = 10 To 12
        ws.Rows(i).RowHeight = 26
    Next i
    ws.Rows(13).RowHeight = 14

"    Set rngBtn = ws.Range(""C10:D12"")"
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, rngBtn.Left + 2, rngBtn.Top + 2, rngBtn.Width - 4, rngBtn.Height - 4)
    shp.Name = BTN_MAIN
"    StyleButton shp, ""Term 03 Analysis"", RGB(0, 112, 192), 24"
"    shp.OnAction = ""'"" & ThisWorkbook.Name & ""'!modBiuldAnalysisfilesBuildAll03.Build_All"""

"    Set rngBtn = ws.Range(""E10:E12"")"
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, rngBtn.Left + 2, rngBtn.Top + 2, rngBtn.Width - 4, rngBtn.Height - 4)
    shp.Name = BTN_OPEN
"    StyleButton shp, ArT(""641 62A 62D 20 645 62C 644 62F 20 627 644 625 635 62F 627 631""), RGB(89, 89, 89), 14"
"    shp.OnAction = ""'"" & ThisWorkbook.Name & ""'!modBiuldAnalysisfilesPanel.Panel_OpenOutput"""

    ' ---- status table
"    MergeText ws, 14, 2, 7, ArT(""62D 627 644 629 20 627 644 645 644 641 627 62A"")"
"    With ws.Range(""B14:G14"")"
        .Font.Name = gFontBase: .Font.Size = 16: .Font.Bold = True
        .Font.Color = RGB(28, 55, 103)
        .HorizontalAlignment = xlRight
        .ReadingOrder = xlRTL
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Color = RGB(28, 55, 103)
        .Borders(xlEdgeBottom).Weight = xlMedium
    End With
    ws.Rows(14).RowHeight = 30

"    ws.Range(""B15"").Value2 = ""#"""
"    ws.Range(""C15"").Value2 = ArT(""627 644 645 644 641"")"
"    ws.Range(""D15"").Value2 = ArT(""627 644 623 642 633 627 645"")"
"    ws.Range(""E15"").Value2 = ArT(""627 644 62D 627 644 629"")"
"    ws.Range(""F15"").Value2 = ArT(""627 644 62A 641 627 635 64A 644"")"
"    ws.Range(""G15"").Value2 = ArT(""627 644 645 62F 629 20 28 62B 29"")"
"    StyleHead ws.Range(""B15:G15"")"
    ws.Rows(15).RowHeight = 28

    Set map = GetFileMap()
    n = 0
    For Each k In map.Keys
        n = n + 1
        With ws.Cells(ROW_FIRST + n - 1, 2)
            .Value2 = n
"            .NumberFormat = ""00"""
        End With
        ws.Cells(ROW_FIRST + n - 1, 3).Value2 = CStr(k)
        ws.Cells(ROW_FIRST + n - 1, 4).Value2 = UBound(map(k)) - LBound(map(k)) + 1
        ws.Rows(ROW_FIRST + n - 1).RowHeight = 26
    Next k
    With ws.Range(ws.Cells(ROW_FIRST, 2), ws.Cells(ROW_FIRST + n - 1, 7))
        .Font.Name = gFontNum: .Font.Size = 13
        .HorizontalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(191, 191, 191)
    End With
    ws.Range(ws.Cells(ROW_FIRST, 6), ws.Cells(ROW_FIRST + n - 1, 6)).HorizontalAlignment = xlRight
    ws.Range(ws.Cells(ROW_FIRST, 6), ws.Cells(ROW_FIRST + n - 1, 6)).WrapText = True
    ws.Range(ws.Cells(ROW_FIRST, 6), ws.Cells(ROW_FIRST + n - 1, 6)).Font.Size = 12
    For i = 1 To n
"        Panel_Update i, 0, """", 0"
    Next i

    ' ---- footer
    ws.Rows(ROW_FIRST + n).RowHeight = 12
"    MergeText ws, ROW_FIRST + n + 1, 2, 7, ArT(""644 627 20 64A 64F 633 62A 628 62F 644 20 623 64A 20 645 644 641 20 645 648 62C 648 62F 3A 20 639 646 62F 20 625 639 627 62F 629 20 627 644 62A 634 63A 64A 644 20 62A 64F 62D 641 638 20 627 644 645 644 641 627 62A 20 627 644 62C 62F 64A 62F 629 20 628 627 644 644 627 62D 642 629 20 5F 30 31 20 648 20 5F 30 32 20 2E 2E 2E 20 648 64A 64F 62D 641 638 20 633 62C 644 20 643 627 645 644 20 644 643 644 20 62A 634 63A 64A 644 20 641 64A 20 645 62C 644 62F 20 627 644 625 635 62F 627 631 2E"")"
    With ws.Range(ws.Cells(ROW_FIRST + n + 1, 2), ws.Cells(ROW_FIRST + n + 1, 7))
        .Font.Name = gFontName: .Font.Size = 12
        .Font.Color = RGB(100, 100, 100)
        .HorizontalAlignment = xlRight
        .WrapText = True
        .ReadingOrder = xlRTL
    End With
    ws.Rows(ROW_FIRST + n + 1).RowHeight = 34

    On Error Resume Next
    ws.Activate
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.Zoom = 100
"    ws.Range(""A1"").Select"
    Err.Clear
    On Error GoTo 0
    Application.ScreenUpdating = True
    Exit Sub
EH:
    Application.ScreenUpdating = True
"    MsgBox ""Panel_Create: "" & Err.Number & "" : "" & Err.Description, vbCritical, ""Term 03 Analysis"""
End Sub

Private Sub PutInfo(ByVal ws As Worksheet, ByVal r As Long, ByVal lbl As String, ByVal txt As String)
    ws.Range(ws.Cells(r, 2), ws.Cells(r, 3)).Merge
    ws.Cells(r, 2).Value2 = lbl
    With ws.Range(ws.Cells(r, 2), ws.Cells(r, 3))
        .Interior.Color = RGB(28, 55, 103)
        .Font.Name = gFontBase: .Font.Size = 14: .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
        .ReadingOrder = xlRTL
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(255, 255, 255)
    End With
    ws.Range(ws.Cells(r, 4), ws.Cells(r, 7)).Merge
"    ws.Cells(r, 4).NumberFormat = ""@"""
    ws.Cells(r, 4).Value2 = txt
    With ws.Range(ws.Cells(r, 4), ws.Cells(r, 7))
        .Interior.Color = RGB(242, 242, 242)
        .Font.Name = gFontNum: .Font.Size = 12
        .HorizontalAlignment = xlLeft
        .ShrinkToFit = True
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(255, 255, 255)
    End With
    ws.Rows(r).RowHeight = 28
End Sub

Private Sub StyleButton(ByVal shp As Shape, ByVal txt As String, ByVal clr As Long, ByVal sz As Long)
    On Error Resume Next
    shp.Fill.ForeColor.RGB = clr
    shp.Line.ForeColor.RGB = RGB(28, 55, 103)
    shp.Line.Weight = 1.5
    shp.Adjustments.Item(1) = 0.25
    shp.Shadow.Visible = msoTrue
    shp.Shadow.OffsetX = 2
    shp.Shadow.OffsetY = 3
    shp.Shadow.Transparency = 0.65
    With shp.TextFrame2
        .TextRange.Text = txt
        .TextRange.Font.Name = gFontBase
        .TextRange.Font.Size = sz
        .TextRange.Font.Bold = msoTrue
        .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .HorizontalAnchor = msoAnchorCenter
        .VerticalAnchor = msoAnchorMiddle
        .WordWrap = msoFalse
    End With
    Err.Clear
End Sub

'==============================================================================
' Called by Build_All (silent when the panel does not exist)
'==============================================================================
Private Function PanelWs() As Worksheet
    On Error Resume Next
    Set PanelWs = ThisWorkbook.Worksheets(PANEL_SHEET)
    Err.Clear
End Function

' kind: 0 waiting, 1 working, 2 done, 3 failed
Public Sub Panel_Update(ByVal idx As Long, ByVal kind As Long, ByVal detail As String, ByVal secs As Double)
    Dim ws As Worksheet, r As Long
    Set ws = PanelWs()
    If ws Is Nothing Then Exit Sub
    EnsureFonts
    r = ROW_FIRST + idx - 1
    On Error Resume Next
    With ws.Cells(r, 5)
        Select Case kind
            Case 0
"                .Value2 = ArT(""641 64A 20 627 644 627 646 62A 638 627 631"")"
                .Interior.Color = RGB(242, 242, 242): .Font.Color = RGB(100, 100, 100)
            Case 1
"                .Value2 = ArT(""642 64A 62F 20 627 644 625 646 62C 627 632"")"
                .Interior.Color = RGB(255, 242, 204): .Font.Color = RGB(156, 101, 0)
            Case 2
"                .Value2 = ArT(""62A 645"")"
                .Interior.Color = RGB(198, 239, 206): .Font.Color = RGB(0, 97, 0)
            Case Else
"                .Value2 = ArT(""641 634 644"")"
                .Interior.Color = RGB(255, 199, 206): .Font.Color = RGB(156, 0, 6)
        End Select
        .Font.Bold = True
        .Font.Name = gFontBase
    End With
"    ws.Cells(r, 6).NumberFormat = ""@"""
    ws.Cells(r, 6).Value2 = detail
    If kind >= 2 Then
        ws.Cells(r, 7).Value2 = Round(secs, 1)
"        ws.Cells(r, 7).NumberFormat = ""0.0"""
    ElseIf kind = 0 Then
        ws.Cells(r, 7).ClearContents
    End If
    If kind = 3 Then
        ws.Rows(r).RowHeight = 54
    Else
        ws.Rows(r).RowHeight = 26
    End If
    Err.Clear
End Sub

Public Sub Panel_Reset(ByVal map As Object)
    Dim ws As Worksheet, i As Long
    Set ws = PanelWs()
    If ws Is Nothing Then Exit Sub
    On Error Resume Next
    For i = 1 To map.Count
"        Panel_Update i, 0, """", 0"
    Next i
"    ws.Cells(7, 4).Value2 = Format$(Now, ""yyyy-mm-dd  hh:nn"")"
"    ws.Cells(8, 4).Value2 = ArT(""642 64A 62F 20 627 644 62A 646 641 64A 630 20 2E 2E 2E"")"
    ws.Cells(8, 4).Font.Color = RGB(156, 101, 0)
    Err.Clear
End Sub

Public Sub Panel_Finish(ByVal okCount As Long, ByVal total As Long, ByVal secs As Double, ByVal note As String)
    Dim ws As Worksheet
    Set ws = PanelWs()
    If ws Is Nothing Then Exit Sub
    On Error Resume Next
"    If note <> """" Then"
        ws.Cells(8, 4).Value2 = note
        ws.Cells(8, 4).Font.Color = RGB(156, 0, 6)
    Else
"        ws.Cells(8, 4).Value2 = ArT(""62A 645 20 625 646 634 627 621 20"") & okCount & ArT(""20 20 645 646 20 20"") & total & ArT(""20 20 645 644 641 627 64B 20 20 2D 20 20"") & Format$(secs, ""0"") & ArT(""20 20 62B 627 646 64A 629"")"
        If okCount = total Then
            ws.Cells(8, 4).Font.Color = RGB(0, 97, 0)
        Else
            ws.Cells(8, 4).Font.Color = RGB(156, 0, 6)
        End If
    End If
    ws.Cells(8, 4).Font.Bold = True
    Err.Clear
End Sub

Public Sub Panel_Show()
    Dim ws As Worksheet
    Set ws = PanelWs()
    If ws Is Nothing Then Exit Sub
    On Error Resume Next
    ThisWorkbook.Activate
    ws.Activate
"    ws.Range(""A1"").Select"
    Err.Clear
End Sub

Public Sub Panel_OpenOutput()
    Dim p As String
"    p = ThisWorkbook.Path & ""\"" & OutputFolderName()"
    If Not FolderExistsP(p) Then
"        MsgBox ""The output folder does not exist yet. Run Term 03 Analysis first."", vbInformation, ""Term 03 Analysis"""
        Exit Sub
    End If
    On Error Resume Next
    ThisWorkbook.FollowHyperlink Address:=p, NewWindow:=True
    Err.Clear
End Sub
