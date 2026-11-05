Attribute VB_Name = "modAnalysisBuilderBuildAll03"
Option Explicit

'==============================================================================
' Module  : modAnalysisBuilderBuildAll03
' Purpose : ONE macro (Build_All) that creates all the Term 03 & Annual analysis files.
'           BuildLog file generation has been removed as requested.
'==============================================================================

Public Sub Build_All()
    RunBuild ""
End Sub

Public Sub Build_One()
    Dim s As String
    s = InputBox("File to build (stem of the name), for example: 3SEGlobalAnalysis03", "Build_One", "3SEGlobalAnalysis03")
    If Len(Trim$(s)) = 0 Then Exit Sub
    RunBuild Trim$(s)
End Sub

Private Sub RunBuild(ByVal onlyStem As String)
    Dim wbGlobal As Workbook
    Dim sheetIndex As Object, dupKeys As Object, map As Object
    Dim key As Variant, stem As String, titles As Variant
    Dim srcReport As String, why As String, errMsg As String, savedPath As String
    Dim outDir As String, failBrief As String, msg As String
    Dim okCount As Long, failCount As Long, total As Long, idx As Long
    Dim oldSU As Boolean, oldEV As Boolean, oldDA As Boolean, oldSB As Variant, oldCalc As Long
    Dim t0 As Double, t1 As Double, secs As Double, secsAll As Double
    Dim retried As Boolean

    If Len(ThisWorkbook.Path) = 0 Then
        MsgBox "Save the tool workbook first: the Source03 folder and the output folder are located next to it.", _
               vbExclamation, "Term 03 Analysis"
        Exit Sub
    End If
    If LCase$(Left$(ThisWorkbook.Path, 4)) = "http" Then
        MsgBox "The tool is on OneDrive/SharePoint (URL path). Move it to a local folder first.", _
               vbExclamation, "Term 03 Analysis"
        Exit Sub
    End If

    oldSU = Application.ScreenUpdating
    oldEV = Application.EnableEvents
    oldDA = Application.DisplayAlerts
    oldSB = Application.StatusBar
    oldCalc = Application.Calculation
    t0 = Timer

    On Error GoTo Fatal
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False

    Set map = GetFileMap()
    If onlyStem <> "" Then
        If Not map.Exists(onlyStem) Then
            MsgBox "Unknown file name: " & onlyStem, vbExclamation, "Term 03 Analysis"
            GoTo Done
        End If
        total = 1
    Else
        total = map.Count
    End If

    Panel_Reset map

    Application.StatusBar = "Term 03 Analysis : checking fonts ..."
    EnsureFonts

    Application.StatusBar = "Term 03 Analysis : reading Source03 ..."
    If Not LoadSource(wbGlobal, sheetIndex, dupKeys, srcReport) Then
        Panel_Finish 0, total, 0, "Source03: " & Replace(Trim$(srcReport), vbCrLf, " | ")
        MsgBox "No usable source file was found." & vbCrLf & vbCrLf & Left$(srcReport, 700), vbCritical, "Term 03 Analysis"
        GoTo Done
    End If

    outDir = ThisWorkbook.Path & "\" & OutputFolderName() & "\"
    If Not EnsureFolderP(outDir) Then
        MsgBox "The output folder could not be created:" & vbCrLf & outDir, vbCritical, "Term 03 Analysis"
        GoTo Done
    End If

    For Each key In map.Keys
        idx = idx + 1
        stem = CStr(key)
        If onlyStem = "" Or StrComp(stem, onlyStem, vbTextCompare) = 0 Then
            titles = map(key)
            Application.StatusBar = "Term 03 Analysis : " & idx & " / " & map.Count & "  " & stem
            Panel_Update idx, 1, "", 0
            t1 = Timer

            savedPath = ""
            retried = False
            If GroupReady(sheetIndex, dupKeys, titles, why) Then
                errMsg = BuildOneFile(stem, titles, sheetIndex, outDir, savedPath)
                If errMsg <> "" Then
                    DoEvents
                    savedPath = ""
                    errMsg = BuildOneFile(stem, titles, sheetIndex, outDir, savedPath)
                    If errMsg = "" Then retried = True
                End If
            Else
                errMsg = "SKIPPED - the source sections are not complete:" & vbCrLf & why
            End If

            secs = Timer - t1
            If secs < 0 Then secs = secs + 86400

            If errMsg = "" Then
                okCount = okCount + 1
                If retried Then
                    Panel_Update idx, 2, Mid$(savedPath, InStrRev(savedPath, "\") + 1) & "   [retry]", secs
                Else
                    Panel_Update idx, 2, Mid$(savedPath, InStrRev(savedPath, "\") + 1), secs
                End If
            Else
                failCount = failCount + 1
                Panel_Update idx, 3, Replace(Trim$(errMsg), vbCrLf, " | "), secs
                failBrief = failBrief & "  - " & stem & " : " & Left$(Replace(Trim$(errMsg), vbCrLf, " "), 90) & vbCrLf
            End If
        End If
    Next key

    secsAll = Timer - t0
    If secsAll < 0 Then secsAll = secsAll + 86400

    Panel_Finish okCount, total, secsAll, ""

    msg = "Created : " & okCount & " / " & total & "   (" & Format$(secsAll, "0") & " s)" & vbCrLf & "Folder  : " & outDir
    If failCount > 0 Then msg = msg & vbCrLf & vbCrLf & "NOT created (" & failCount & "):" & vbCrLf & failBrief
    msg = msg & vbCrLf & "Full details: see status table on ControlPanel."
    MsgBox Left$(msg, 1000), IIf(failCount = 0, vbInformation, vbExclamation), "Term 03 Analysis"

Done:
    On Error Resume Next
    If Not wbGlobal Is Nothing Then wbGlobal.Close SaveChanges:=False
    Application.Calculation = oldCalc
    Application.StatusBar = oldSB
    Application.DisplayAlerts = oldDA
    Application.EnableEvents = oldEV
    Application.ScreenUpdating = oldSU
    Panel_Show
    Exit Sub

Fatal:
    Panel_Finish 0, total, 0, "Stopped: " & Err.Number & " : " & Err.Description
    MsgBox "Build stopped: " & Err.Number & " : " & Err.Description, vbCritical, "Term 03 Analysis"
    Resume Done
End Sub

Private Function BuildOneFile(ByVal stem As String, ByVal titles As Variant, ByVal sheetIndex As Object, _
                              ByVal outDir As String, ByRef savedPath As String) As String
    Dim wbNew As Workbook, errText As String, acc As String

    On Error GoTo EH
    Application.Calculation = xlCalculationManual
    Set wbNew = Workbooks.Add(xlWBATWorksheet)

    If Left$(stem, 4) = "1Lit" Then
        acc = ArT("645 639 62F 644 20 627 644 642 628 648 644")
    Else
        acc = ArT("645 639 62F 644 20 627 644 627 646 62A 642 627 644")
    End If

    If Not BuildDataSheet(wbNew, sheetIndex, titles, acc, errText) Then Err.Raise vbObjectError + 901, "Build_All", errText
    If Not BuildData01Sheet(wbNew, errText) Then Err.Raise vbObjectError + 905, "Build_All", errText
    If Not BuildAnalysisSheets(wbNew, CoreKeysFor(stem), errText) Then Err.Raise vbObjectError + 902, "Build_All", errText
    If Not BuildCartTecSheet(wbNew, errText) Then Err.Raise vbObjectError + 903, "Build_All", errText

    ' Calculate formulas first so AutoFit on Analysis03 & AnnualAnalysis measures real calculated values
    Application.Calculation = xlCalculationAutomatic
    Application.Calculate
    AutoFitAnalysisSheets wbNew

    On Error Resume Next
    wbNew.Worksheets(DATA_SHEET).Activate
    wbNew.Worksheets(DATA_SHEET).Range("A1").Select
    On Error GoTo EH

    savedPath = NextFreePath(outDir, stem, ".xlsx")
    wbNew.SaveAs Filename:=savedPath, FileFormat:=xlOpenXMLWorkbook
    If Not FileExistsP(savedPath) Then Err.Raise vbObjectError + 904, "Build_All", "The file was not found after saving: " & savedPath

    wbNew.Close SaveChanges:=False
    Set wbNew = Nothing
    BuildOneFile = ""
    Exit Function
EH:
    BuildOneFile = Err.Number & " : " & Err.Description
    Err.Clear
    On Error Resume Next
    If Not wbNew Is Nothing Then wbNew.Close SaveChanges:=False
End Function
