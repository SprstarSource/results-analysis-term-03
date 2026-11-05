Attribute VB_Name = "modBuildPanel03"
Option Explicit

Private m_workflowBusy As Boolean

' Unified Control Panel Sheet
Private Const PANEL_SHEET As String = "ControlPanel"
Private Const OLD_PANEL_SHEET As String = "Term03Panel"
Private Const BUTTON_WIDTH As Double = 164
Private Const BUTTON_HEIGHT As Double = 54
Private Const BUTTON_GAP As Double = 12

' Status Table Row Constants (Shifted up to Row 10)
Private Const ROW_TITLE_STATUS As Long = 10
Private Const ROW_HDR_STATUS As Long = 11
Private Const ROW_FIRST_STATUS As Long = 12

Private Function PanelActions() As Variant
    ' All modules now read directly from \Source03\ (item(3) = False everywhere).
    ' 8 buttons arranged in a clean 2x4 grid next to the status table.
    PanelActions = Array( _
        Array("btnCollect", "Third Term Results", "modCollector.CollectTerm03Results", False, &H60AE27), _
        Array("btnTerm03Analysis", "Term 03 Analysis", "modAnalysisBuilderBuildAll03.Build_All", False, &HC07000), _
        Array("btnAverage", "Annual Averages", "modAverage03.CreateAnnualResultsBy3", False, &HAD448E), _
        Array("btnSplitTerm", "Split Term Results", "modSplitTermResults.SplitTermResults", False, &H85A016), _
        Array("btnExamen", "Examen Analysis", "modExamenAnalysis03.Run_ExamenAnalysis03", False, &H78604D), _
        Array("btnDataTransfer", "Data Transfer", "modTransfer03.Run_Transfer03", False, &H60AE27), _
        Array("btnStreamPreferences", "Stream Preferences", "modStreamPreferences03.Run_StreamPreferences03", False, &HA05568), _
        Array("btnFinalOrientation", "Final Orientation", "modOrientation03.Run_Orientation03", False, &HB98029))
End Function

Public Sub Panel_Create()
    BuildControlPanel
End Sub

Public Sub BuildControlPanel()
    Dim ws As Worksheet, wsOld As Worksheet, btn As Shape, actions As Variant, item As Variant
    Dim map As Object, k As Variant, n As Long
    Dim i As Long, count As Long, columns As Long, gridRow As Long, gridCol As Long
    Dim rowItems As Long, canvasWidth As Double, rowWidth As Double
    Dim originLeft As Double, originTop As Double, x As Double, y As Double
    Dim prevScreen As Boolean, prevAlerts As Boolean, errorText As String, base As String

    prevScreen = Application.ScreenUpdating
    prevAlerts = Application.DisplayAlerts
    On Error GoTo BuildError
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    If ThisWorkbook.ReadOnly Then Err.Raise vbObjectError + 801, , "The tool workbook is read-only."
    EnsureFonts

    On Error Resume Next
    Set wsOld = ThisWorkbook.Worksheets(OLD_PANEL_SHEET)
    If Not wsOld Is Nothing Then
        If ThisWorkbook.Worksheets.Count > 1 Then wsOld.Delete
    End If
    Err.Clear
    On Error GoTo BuildError

    Set ws = PanelSheet()
    If ws Is Nothing Then
        If ThisWorkbook.ProtectStructure Then Err.Raise vbObjectError + 802, , "The workbook structure is protected."
        Set ws = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Worksheets(1))
        ws.Name = PANEL_SHEET
    End If
    If ws.ProtectContents Or ws.ProtectDrawingObjects Then
        Err.Raise vbObjectError + 803, , "Unprotect the control panel sheet before rebuilding it."
    End If

    ws.Activate
    For i = ws.Shapes.Count To 1 Step -1
        ws.Shapes(i).Delete
    Next i

    With ws
        .Cells.UnMerge
        .Cells.Clear
        .DisplayRightToLeft = True
        With .Cells
            .Font.Name = gFontBase
            .Font.Size = 12
            .VerticalAlignment = xlCenter
            .Interior.Color = RGB(255, 255, 255)
        End With

        .Columns(1).ColumnWidth = 2.5
        .Columns(2).ColumnWidth = 5
        .Columns(3).ColumnWidth = 28
        .Columns(4).ColumnWidth = 10
        .Columns(5).ColumnWidth = 16
        .Columns(6).ColumnWidth = 52
        .Columns(7).ColumnWidth = 11
        .Columns(8).ColumnWidth = 3
        .Columns(9).ColumnWidth = 22
        .Columns(10).ColumnWidth = 22
        .Columns(11).ColumnWidth = 22
        .Columns(12).ColumnWidth = 2.5
    End With

    ' ---- Banner & Subtitle (Full width B:K)
    ws.Rows(1).RowHeight = 10
    MergeText ws, 2, 2, 11, ArT("62A 62D 644 64A 644 20 646 62A 627 626 62C 20 627 644 641 635 644 20 627 644 62B 627 644 62B") & "  |  Third Term Management Dashboard"
    With ws.Range("B2:K2")
        .Interior.Color = RGB(28, 55, 103)
        .Font.Name = gFontBase: .Font.Size = 24: .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
        .ReadingOrder = xlRTL
    End With
    ws.Rows(2).RowHeight = 52

    MergeText ws, 3, 2, 11, ArT("625 635 62F 627 631 20 645 644 641 627 62A 20 627 644 62A 62D 644 64A 644 20") & _
                           "(" & DATA_SHEET & "  -  " & DATA01_SHEET & "  -  " & ANALYSIS_T3_SHEET & "  -  " & ANALYSIS_ANN_SHEET & "  -  " & CART_SHEET & ") " & _
                           ArT("644 643 644 20 627 644 634 639 628 20 628 636 63A 637 629 20 648 627 62D 62F 629")
    With ws.Range("B3:K3")
        .Interior.Color = RGB(221, 235, 247)
        .Font.Name = gFontName: .Font.Size = 14
        .Font.Color = RGB(28, 55, 103)
        .HorizontalAlignment = xlCenter
        .ReadingOrder = xlRTL
    End With
    ws.Rows(3).RowHeight = 32
    ws.Rows(4).RowHeight = 12

    ' ---- Information Block (Rows 5 to 8 in Columns B:G)
    base = ThisWorkbook.Path
    PutInfo ws, 5, ArT("645 62C 644 62F 20 627 644 645 635 62F 631"), base & "\Source03"
    PutInfo ws, 6, ArT("645 62C 644 62F 20 627 644 625 635 62F 627 631"), base & "\" & OutputFolderName()
    PutInfo ws, 7, ArT("622 62E 631 20 62A 634 63A 64A 644"), "-"
    PutInfo ws, 8, ArT("627 644 646 62A 64A 62C 629"), "-"
    ws.Rows(9).RowHeight = 14

    ' ---- Status Table (Row 10, Columns B:G)
    MergeText ws, ROW_TITLE_STATUS, 2, 7, ArT("62D 627 644 629 20 627 644 645 644 641 627 62A")
    With ws.Range(ws.Cells(ROW_TITLE_STATUS, 2), ws.Cells(ROW_TITLE_STATUS, 7))
        .Font.Name = gFontBase: .Font.Size = 16: .Font.Bold = True
        .Font.Color = RGB(28, 55, 103)
        .HorizontalAlignment = xlRight
        .ReadingOrder = xlRTL
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Color = RGB(28, 55, 103)
        .Borders(xlEdgeBottom).Weight = xlMedium
    End With
    ws.Rows(ROW_TITLE_STATUS).RowHeight = 30

    ws.Cells(ROW_HDR_STATUS, 2).Value2 = "#"
    ws.Cells(ROW_HDR_STATUS, 3).Value2 = ArT("627 644 645 644 641")
    ws.Cells(ROW_HDR_STATUS, 4).Value2 = ArT("627 644 623 642 633 627 645")
    ws.Cells(ROW_HDR_STATUS, 5).Value2 = ArT("627 644 62D 627 644 629")
    ws.Cells(ROW_HDR_STATUS, 6).Value2 = ArT("627 644 62A 641 627 635 64A 644")
    ws.Cells(ROW_HDR_STATUS, 7).Value2 = ArT("627 644 645 62F 629 20 28 62B 29")
    StyleHead ws.Range(ws.Cells(ROW_HDR_STATUS, 2), ws.Cells(ROW_HDR_STATUS, 7))
    ws.Rows(ROW_HDR_STATUS).RowHeight = 28

    Set map = GetFileMap()
    n = 0
    For Each k In map.Keys
        n = n + 1
        With ws.Cells(ROW_FIRST_STATUS + n - 1, 2)
            .Value2 = n
            .NumberFormat = "00"
        End With
        ws.Cells(ROW_FIRST_STATUS + n - 1, 3).Value2 = CStr(k)
        ws.Cells(ROW_FIRST_STATUS + n - 1, 4).Value2 = UBound(map(k)) - LBound(map(k)) + 1
        ws.Rows(ROW_FIRST_STATUS + n - 1).RowHeight = 26
    Next k

    With ws.Range(ws.Cells(ROW_FIRST_STATUS, 2), ws.Cells(ROW_FIRST_STATUS + n - 1, 7))
        .Font.Name = gFontNum: .Font.Size = 13
        .HorizontalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(191, 191, 191)
    End With
    ws.Range(ws.Cells(ROW_FIRST_STATUS, 6), ws.Cells(ROW_FIRST_STATUS + n - 1, 6)).HorizontalAlignment = xlRight
    ws.Range(ws.Cells(ROW_FIRST_STATUS, 6), ws.Cells(ROW_FIRST_STATUS + n - 1, 6)).WrapText = True
    ws.Range(ws.Cells(ROW_FIRST_STATUS, 6), ws.Cells(ROW_FIRST_STATUS + n - 1, 6)).Font.Size = 12
    For i = 1 To n
        Panel_Update i, 0, "", 0
    Next i

    ' ---- Footer Note under the Status Table
    ws.Rows(ROW_FIRST_STATUS + n).RowHeight = 12
    MergeText ws, ROW_FIRST_STATUS + n + 1, 2, 7, ArT("644 627 20 64A 64F 633 62A 628 62F 644 20 623 64A 20 645 644 641 20 645 648 62C 648 62F 3A 20 639 646 62F 20 625 639 627 62F 629 20 627 644 62A 634 63A 64A 644 20 62A 64F 62D 641 638 20 627 644 645 644 641 627 62A 20 627 644 62C 62F 64A 62F 629 20 628 627 644 644 627 62D 642 629 20 5F 30 31 20 648 20 5F 30 32 20 2E 2E 2E")
    With ws.Range(ws.Cells(ROW_FIRST_STATUS + n + 1, 2), ws.Cells(ROW_FIRST_STATUS + n + 1, 7))
        .Font.Name = gFontName: .Font.Size = 12
        .Font.Color = RGB(100, 100, 100)
        .HorizontalAlignment = xlRight
        .WrapText = True
        .ReadingOrder = xlRTL
    End With
    ws.Rows(ROW_FIRST_STATUS + n + 1).RowHeight = 32

    ' ---- Action Buttons Grid (Columns I:K, starting at Row 5)
    actions = PanelActions()
    count = UBound(actions) - LBound(actions) + 1
    columns = 2
    canvasWidth = columns * BUTTON_WIDTH + (columns - 1) * BUTTON_GAP
    originLeft = ws.Range("I5").Left + (ws.Range("I5:K5").Width - canvasWidth) / 2
    If originLeft < ws.Range("I5").Left Then originLeft = ws.Range("I5").Left
    originTop = ws.Range("I5").Top

    For i = 0 To count - 1
        item = actions(i)
        gridRow = i \ columns
        gridCol = i Mod columns
        rowItems = count - gridRow * columns
        If rowItems > columns Then rowItems = columns
        rowWidth = rowItems * BUTTON_WIDTH + (rowItems - 1) * BUTTON_GAP
        x = originLeft + (canvasWidth - rowWidth) / 2 + gridCol * (BUTTON_WIDTH + BUTTON_GAP)
        y = originTop + gridRow * (BUTTON_HEIGHT + BUTTON_GAP)
        Set btn = ws.Shapes.AddShape(msoShapeRoundedRectangle, x, y, BUTTON_WIDTH, BUTTON_HEIGHT)
        With btn
            .Name = CStr(item(0))
            .Adjustments(1) = 0.12
            .Placement = xlFreeFloating
            .Line.Weight = 1.25
            With .TextFrame2
                .TextRange.Text = CStr(item(1))
                .TextRange.Font.Name = "Segoe UI"
                .TextRange.Font.Size = 12
                .TextRange.Font.Bold = msoTrue
                .TextRange.ParagraphFormat.Alignment = msoAlignCenter
                .VerticalAnchor = msoAnchorMiddle
                .WordWrap = msoTrue
                .MarginLeft = 7
                .MarginRight = 7
                .MarginTop = 4
                .MarginBottom = 4
            End With
        End With
    Next i

    If Not ActiveWindow Is Nothing Then
        ActiveWindow.DisplayGridlines = False
        ActiveWindow.DisplayHeadings = False
        ActiveWindow.Zoom = 100
    End If
    ws.Range("A1").Select

    RefreshPanelState
    CheckPanelMacros actions
    Application.DisplayAlerts = prevAlerts
    Application.ScreenUpdating = prevScreen
    MsgBox "Control panel updated successfully (" & count & " buttons + Status Table).", vbInformation, "Done"
    Exit Sub

BuildError:
    errorText = CStr(Err.Number) & " : " & Err.Description
    Application.DisplayAlerts = prevAlerts
    Application.ScreenUpdating = prevScreen
    MsgBox "Panel build failed: " & errorText, vbExclamation, "Error"
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
    ws.Cells(r, 4).NumberFormat = "@"
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

'==============================================================================
' STATUS TABLE UPDATERS (Called by modAnalysisBuilderBuildAll03)
'==============================================================================
Public Sub Panel_Update(ByVal idx As Long, ByVal kind As Long, ByVal detail As String, ByVal secs As Double)
    Dim ws As Worksheet, r As Long
    Set ws = PanelSheet()
    If ws Is Nothing Then Exit Sub
    EnsureFonts
    r = ROW_FIRST_STATUS + idx - 1
    On Error Resume Next
    With ws.Cells(r, 5)
        Select Case kind
            Case 0
                .Value2 = ArT("641 64A 20 627 644 627 646 62A 638 627 631")
                .Interior.Color = RGB(242, 242, 242): .Font.Color = RGB(100, 100, 100)
            Case 1
                .Value2 = ArT("642 64A 62F 20 627 644 625 646 62C 627 632")
                .Interior.Color = RGB(255, 242, 204): .Font.Color = RGB(156, 101, 0)
            Case 2
                .Value2 = ArT("62A 645")
                .Interior.Color = RGB(198, 239, 206): .Font.Color = RGB(0, 97, 0)
            Case Else
                .Value2 = ArT("641 634 644")
                .Interior.Color = RGB(255, 199, 206): .Font.Color = RGB(156, 0, 6)
        End Select
        .Font.Bold = True
        .Font.Name = gFontBase
    End With
    ws.Cells(r, 6).NumberFormat = "@"
    ws.Cells(r, 6).Value2 = detail
    If kind >= 2 Then
        ws.Cells(r, 7).Value2 = Round(secs, 1)
        ws.Cells(r, 7).NumberFormat = "0.0"
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
    Set ws = PanelSheet()
    If ws Is Nothing Then Exit Sub
    On Error Resume Next
    For i = 1 To map.Count
        Panel_Update i, 0, "", 0
    Next i
    ws.Cells(7, 4).Value2 = Format$(Now, "yyyy-mm-dd  hh:nn")
    ws.Cells(8, 4).Value2 = ArT("642 64A 62F 20 627 644 62A 646 641 64A 630 20 2E 2E 2E")
    ws.Cells(8, 4).Font.Color = RGB(156, 101, 0)
    Err.Clear
End Sub

Public Sub Panel_Finish(ByVal okCount As Long, ByVal total As Long, ByVal secs As Double, ByVal note As String)
    Dim ws As Worksheet
    Set ws = PanelSheet()
    If ws Is Nothing Then Exit Sub
    On Error Resume Next
    If note <> "" Then
        ws.Cells(8, 4).Value2 = note
        ws.Cells(8, 4).Font.Color = RGB(156, 0, 6)
    Else
        ws.Cells(8, 4).Value2 = ArT("62A 645 20 625 646 634 627 621 20") & okCount & ArT("20 20 645 646 20 20") & total & ArT("20 20 645 644 641 627 64B 20 20 2D 20 20") & Format$(secs, "0") & ArT("20 20 62B 627 646 64A 629")
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
    Set ws = PanelSheet()
    If ws Is Nothing Then Exit Sub
    On Error Resume Next
    ThisWorkbook.Activate
    ws.Activate
    ws.Range("A1").Select
    Err.Clear
End Sub

'==============================================================================
' BUTTON STATE & WORKFLOW MANAGEMENT (Decoupled from Term03Results.xlsx)
'==============================================================================
Public Sub RefreshPanelState()
    Dim ws As Worksheet, btn As Shape, actions As Variant, item As Variant
    Dim i As Long, enabled As Boolean, stage As Long, stepNo As Long
    Set ws = PanelSheet()
    If ws Is Nothing Then Exit Sub
    stage = FinalWorkflowStage03()
    actions = PanelActions()
    For i = LBound(actions) To UBound(actions)
        item = actions(i)
        Set btn = Nothing
        On Error Resume Next
        Set btn = ws.Shapes(CStr(item(0)))
        Err.Clear
        On Error GoTo 0
        If Not btn Is Nothing Then
            enabled = True
            stepNo = WorkflowStep03(CStr(item(0)))
            If stepNo > 1 Then enabled = (stage >= stepNo - 1)
            With btn
                .Fill.Solid
                If enabled Then
                    .Fill.ForeColor.RGB = CLng(item(4))
                    .Line.ForeColor.RGB = CLng(item(4))
                    .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
                Else
                    .Fill.ForeColor.RGB = RGB(220, 220, 220)
                    .Line.ForeColor.RGB = RGB(180, 180, 180)
                    .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(120, 120, 120)
                End If
                If stepNo > 0 Then
                    .OnAction = PanelMacroAction("RunFinalStep03")
                Else
                    .OnAction = PanelMacroAction(CStr(item(2)))
                End If
            End With
        End If
    Next i
End Sub

Private Function PanelSheet() As Worksheet
    On Error Resume Next
    Set PanelSheet = ThisWorkbook.Worksheets(PANEL_SHEET)
    On Error GoTo 0
End Function

Private Function PanelMacroAction(ByVal macroName As String) As String
    PanelMacroAction = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!" & macroName
End Function

Private Sub CheckPanelMacros(ByVal actions As Variant)
    Dim i As Long, item As Variant, missing As String
    For i = LBound(actions) To UBound(actions)
        item = actions(i)
        If Not PanelMacroExists(CStr(item(2))) Then missing = missing & "  - " & CStr(item(2)) & vbCrLf
    Next i
    If Len(missing) > 0 Then
        MsgBox "Import the missing modules before using these actions:" & vbCrLf & missing, _
               vbExclamation, "Missing Macros"
    End If
End Sub

Private Function PanelMacroExists(ByVal macroName As String) As Boolean
    Dim comp As Object, codeMod As Object, txt As String
    Dim moduleName As String, procedureName As String, dotPos As Long
    On Error GoTo AccessDenied
    dotPos = InStrRev(macroName, ".")
    procedureName = macroName
    If dotPos > 0 Then
        moduleName = Left$(macroName, dotPos - 1)
        procedureName = Mid$(macroName, dotPos + 1)
    End If
    For Each comp In ThisWorkbook.VBProject.VBComponents
        If comp.Type = 1 Then
            If Len(moduleName) = 0 Or StrComp(comp.Name, moduleName, vbTextCompare) = 0 Then
                Set codeMod = comp.CodeModule
                If codeMod.CountOfLines > 0 Then
                    txt = codeMod.Lines(1, codeMod.CountOfLines)
                    If InStr(1, txt, "Sub " & procedureName & "(", vbTextCompare) > 0 Then
                        PanelMacroExists = True
                        Exit Function
                    End If
                End If
            End If
        End If
    Next comp
    Exit Function
AccessDenied:
    PanelMacroExists = True
    Err.Clear
End Function

' Ordered pipeline now starts directly at Data Transfer (Step 1) -> Stream Preferences (Step 2) -> Final Orientation (Step 3)
Private Function WorkflowStep03(ByVal shapeName As String) As Long
    Select Case shapeName
    Case "btnDataTransfer": WorkflowStep03 = 1
    Case "btnStreamPreferences": WorkflowStep03 = 2
    Case "btnFinalOrientation": WorkflowStep03 = 3
    End Select
End Function

Private Function WorkflowRead03(ByVal key As String) As String
    Dim s As String
    On Error GoTo Missing
    s = ThisWorkbook.Names("_Final03_" & key).RefersTo
    If Left$(s, 2) = "=" & Chr$(34) And Right$(s, 1) = Chr$(34) Then
        WorkflowRead03 = Replace(Mid$(s, 3, Len(s) - 3), Chr$(34) & Chr$(34), Chr$(34))
    End If
Missing:
End Function

Private Sub WorkflowWrite03(ByVal key As String, ByVal value As String)
    On Error Resume Next
    ThisWorkbook.Names("_Final03_" & key).Delete
    Err.Clear
    On Error GoTo 0
    ThisWorkbook.Names.Add Name:="_Final03_" & key, RefersTo:="=" & Chr$(34) & Replace(value, Chr$(34), Chr$(34) & Chr$(34)) & Chr$(34), Visible:=False
End Sub

Private Function WorkflowFile03(ByVal fileName As String) As String
    Dim path As String
    On Error GoTo Missing
    If Len(ThisWorkbook.Path) = 0 Then Exit Function
    path = ThisWorkbook.Path & Application.PathSeparator & fileName
    WorkflowFile03 = CStr(FileLen(path)) & "|" & Format$(FileDateTime(path), "yyyymmddhhnnss")
    Exit Function
Missing:
    WorkflowFile03 = ""
End Function

Private Function FinalWorkflowStage03() As Long
    Dim stage As Long, s As String
    s = WorkflowRead03("Stage")
    If IsNumeric(s) Then
        If Val(s) >= 0 And Val(s) <= 3 Then stage = CLng(s)
    End If
    If stage < 0 Or stage > 3 Then stage = 0
    If stage >= 1 Then
        s = WorkflowFile03("FinalGuidance.xlsx")
        If Len(s) = 0 Or s <> WorkflowRead03("Guidance") Then stage = 0
    End If
    If stage >= 2 Then
        s = WorkflowFile03("FinalStreamPreferences.xlsx")
        If Len(s) = 0 Or s <> WorkflowRead03("Wishes") Then stage = 1
    End If
    FinalWorkflowStage03 = stage
End Function

Private Sub WorkflowSaveStage03(ByVal stage As Long)
    WorkflowWrite03 "Guidance", WorkflowFile03("FinalGuidance.xlsx")
    WorkflowWrite03 "Wishes", WorkflowFile03("FinalStreamPreferences.xlsx")
    WorkflowWrite03 "Stage", CStr(stage)
End Sub

Private Sub WorkflowGuard03(ByVal stepNo As Long)
    Dim previous As String
    Select Case stepNo
    Case 2: previous = "Data Transfer"
    Case 3: previous = "Stream Preferences"
    End Select
    MsgBox "Complete '" & previous & "' successfully before this action.", vbExclamation, "Required previous step"
End Sub

Public Sub RunFinalStep03()
    Dim caller As Variant, stepNo As Long, stage As Long, actions As Variant, item As Variant
    Dim i As Long, macro As String, resultFunction As String, succeeded As Boolean
    Dim errorText As String
    If m_workflowBusy Then Exit Sub
    caller = Application.Caller
    If VarType(caller) <> vbString Then
        MsgBox "Use the ordered buttons on ControlPanel.", vbExclamation, "Final guidance"
        Exit Sub
    End If
    stepNo = WorkflowStep03(CStr(caller))
    If stepNo = 0 Then Exit Sub
    stage = FinalWorkflowStage03()
    If stage < stepNo - 1 Then
        WorkflowGuard03 stepNo
        RefreshPanelState
        Exit Sub
    End If
    On Error GoTo Failed
    If ThisWorkbook.ReadOnly Then Err.Raise vbObjectError + 810, , "The tool workbook is read-only."
    m_workflowBusy = True
    WorkflowSaveStage03 stepNo - 1
    RefreshPanelState
    actions = PanelActions()
    For i = LBound(actions) To UBound(actions)
        item = actions(i)
        If CStr(item(0)) = CStr(caller) Then
            macro = CStr(item(2))
            Exit For
        End If
    Next i
    If Len(macro) = 0 Then Err.Raise vbObjectError + 811, , "Panel action not found."
    Application.Run PanelMacroAction(macro)
    Select Case stepNo
    Case 1: resultFunction = "modTransfer03.Transfer03Succeeded"
    Case 2: resultFunction = "modStreamPreferences03.Preferences03Succeeded"
    Case 3: resultFunction = "modOrientation03.Orientation03Succeeded"
    End Select
    If Len(resultFunction) > 0 Then succeeded = CBool(Application.Run(PanelMacroAction(resultFunction)))
    If succeeded Then WorkflowSaveStage03 stepNo
Finished:
    m_workflowBusy = False
    On Error Resume Next
    RefreshPanelState
    On Error GoTo 0
    If Len(errorText) > 0 Then MsgBox errorText, vbExclamation, "Final guidance"
    Exit Sub
Failed:
    errorText = Err.Description
    Resume Finished
End Sub
