Attribute VB_Name = "modSplitTermResults"
Option Explicit

'==============================================================================
' Module  : modSplitTermResults
' Purpose : Reads \Source03\ once in memory and extracts all three terms
'           (FirstTermOnly.xlsx, SecondTermOnly.xlsx, ThirdTermOnly.xlsx)
'           in a single run with full formatting and backup support.
'==============================================================================

Private Const GROUP_SIZE As Long = 3
Private Const SOURCE_FOLDER_NAME As String = "Source03"
Private Const FIRST_SUBJECT_COL As Long = 6

' ---- Embedded Collector Constants ----
Private Const DATA_START_ROW       As Long = 6
Private Const TRAILING_ROWS_SKIP   As Long = 1
Private Const IDENTITY_LAST_COL    As Long = 5
Private Const COL_WIDTH            As Double = 5.7
Private Const FONT_SIZE            As Long = 14
Private Const FONT_ARABIC          As String = "Arabic UI Display Light"
Private Const FONT_NUMERIC         As String = "Bahnschrift SemiBold"
Private Const FONT_FALLBACK        As String = "Calibri"
Private Const PREFERRED_FONT       As String = "Cairo"
Private Const FALLBACK_FONT        As String = "Arial"
Private Const HEADER_FONT_SIZE     As Double = 11
Private Const HEADER_ROW_HEIGHT    As Double = 30

Private g_ActiveArabicFont  As String
Private g_ActiveNumericFont As String

Public Sub SplitTermResults()
    Dim fso As Object, wbSource As Workbook
    Dim folderPath As String, srcFolder As String
    Dim outNames(1 To 3) As String, t As Long
    Dim copiedSrc As Long, skippedSrc As Long, skippedSrcList As String
    Dim summaryReport As String, termReport As String, fatalMsg As String
    Dim prevScreen As Boolean, prevEvents As Boolean, prevAlerts As Boolean, prevCalc As Long
    Dim prevSec As MsoAutomationSecurity, settingsCaptured As Boolean
    Dim anyExisting As Boolean, existingList As String

    On Error GoTo FatalError
    If Len(ThisWorkbook.Path) = 0 Then
        MsgBox "Save this workbook before running the macro.", vbExclamation, "Split Term Results": Exit Sub
    End If
    If LCase$(Left$(ThisWorkbook.Path, 4)) = "http" Then
        MsgBox "Use a local or locally synced folder.", vbExclamation, "Split Term Results": Exit Sub
    End If

    folderPath = ThisWorkbook.Path & "\"
    srcFolder = folderPath & SOURCE_FOLDER_NAME & "\"
    outNames(1) = "FirstTermOnly"
    outNames(2) = "SecondTermOnly"
    outNames(3) = "ThirdTermOnly"

    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(folderPath & SOURCE_FOLDER_NAME) Then
        MsgBox "Source03 folder not found:" & vbCrLf & srcFolder, vbExclamation, "Split Term Results": Exit Sub
    End If

    For t = 1 To 3
        If TermWorkbookOpen(folderPath & outNames(t) & ".xlsx") Then
            MsgBox "Close " & outNames(t) & ".xlsx before running this action.", vbExclamation, "Split Term Results": Exit Sub
        End If
        If fso.FileExists(folderPath & outNames(t) & ".xlsx") Then
            anyExisting = True
            existingList = existingList & "  - " & outNames(t) & ".xlsx" & vbCrLf
        End If
    Next t

    If anyExisting Then
        If MsgBox("Replace the following files after successful processing?" & vbCrLf & existingList & _
                  "Previous files will be backed up in _Backup.", vbYesNo + vbQuestion + vbDefaultButton2, "Split Term Results") <> vbYes Then Exit Sub
    End If

    prevScreen = Application.ScreenUpdating
    prevEvents = Application.EnableEvents
    prevAlerts = Application.DisplayAlerts
    prevCalc = Application.Calculation
    prevSec = Application.AutomationSecurity
    settingsCaptured = True
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    Application.AutomationSecurity = msoAutomationSecurityForceDisable

    ' ---- Collect \Source03\ ONCE in memory ----
    If Not SrcBuildWorkbook(srcFolder, wbSource, copiedSrc, skippedSrc, skippedSrcList) Then
        MsgBox "No usable source files found in " & SOURCE_FOLDER_NAME & ":" & vbCrLf & skippedSrcList, vbExclamation, "Split Term Results"
        GoTo CleanExit
    End If

    summaryReport = "Source03 files read: " & copiedSrc & " (Skipped: " & skippedSrc & ")" & vbCrLf & String$(45, "-")
    For t = 1 To 3
        termReport = BuildSingleTermOutput(wbSource, folderPath, t, outNames(t), fso)
        summaryReport = summaryReport & vbCrLf & termReport & vbCrLf & String$(45, "-")
    Next t

    GoTo CleanExit

FatalError:
    fatalMsg = CStr(Err.Number) & " : " & Err.Description
    Resume CleanExit

CleanExit:
    On Error Resume Next
    If Not wbSource Is Nothing Then wbSource.Close SaveChanges:=False
    If settingsCaptured Then
        Application.AutomationSecurity = prevSec
        Application.Calculation = prevCalc
        Application.EnableEvents = prevEvents
        Application.DisplayAlerts = prevAlerts
        Application.ScreenUpdating = prevScreen
    End If
    On Error GoTo 0

    If Len(skippedSrcList) > 0 Then summaryReport = summaryReport & vbCrLf & "Source03 notes:" & vbCrLf & skippedSrcList
    If Len(fatalMsg) > 0 Then summaryReport = summaryReport & vbCrLf & "ERROR: " & fatalMsg
    MsgBox Left$(summaryReport, 1000), IIf(Len(fatalMsg) = 0, vbInformation, vbExclamation), "Split Term Results Report"
End Sub

Private Function BuildSingleTermOutput(ByVal wbSource As Workbook, ByVal folderPath As String, _
                                       ByVal termNum As Long, ByVal outBasename As String, ByVal fso As Object) As String
    Dim wbResult As Workbook, wsSource As Worksheet, placeholder As Worksheet
    Dim resultFilePath As String, tempPath As String, backupFile As String
    Dim backupRoot As String, backupDir As String, stem As String, suffix As Long
    Dim tableIndex As Long, status As Long
    Dim totalSheets As Long, okSheets As Long, skippedSheets As Long, failedSheets As Long
    Dim reason As String, details As String
    Dim saved As Boolean, previousMoved As Boolean

    On Error GoTo TermErr
    resultFilePath = folderPath & outBasename & ".xlsx"
    Set wbResult = Workbooks.Add(xlWBATWorksheet)
    Set placeholder = wbResult.Worksheets(1)

    For Each wsSource In wbSource.Worksheets
        totalSheets = totalSheets + 1
        status = ConvertTermSheet(wsSource, wbResult, termNum, tableIndex, reason)
        Select Case status
            Case 1: okSheets = okSheets + 1
            Case 0: skippedSheets = skippedSheets + 1
            Case Else
                failedSheets = failedSheets + 1
                details = details & " [" & wsSource.Name & ": " & reason & "]"
        End Select
    Next wsSource

    If okSheets > 0 And failedSheets = 0 Then
        placeholder.Delete
        Set placeholder = Nothing
        tempPath = folderPath & "~" & outBasename & "_" & fso.GetTempName & ".xlsx"
        wbResult.SaveAs Filename:=tempPath, FileFormat:=xlOpenXMLWorkbook
        wbResult.Close SaveChanges:=False
        Set wbResult = Nothing
        If fso.FileExists(resultFilePath) Then
            backupRoot = folderPath & "_Backup"
            If Not fso.FolderExists(backupRoot) Then fso.CreateFolder backupRoot
            stem = backupRoot & "\" & outBasename & "_" & Format(Now, "yyyymmdd_hhnnss")
            backupDir = stem
            Do While fso.FolderExists(backupDir)
                suffix = suffix + 1
                backupDir = stem & "_" & CStr(suffix)
            Loop
            fso.CreateFolder backupDir
            backupFile = backupDir & "\" & outBasename & ".xlsx"
            Name resultFilePath As backupFile
            previousMoved = True
        End If
        Name tempPath As resultFilePath
        saved = True
    End If

    If Not wbResult Is Nothing Then wbResult.Close SaveChanges:=False
    BuildSingleTermOutput = outBasename & ".xlsx -> Converted: " & okSheets & "/" & totalSheets & _
                            IIf(saved, " (Saved)", " (NOT Saved" & details & ")")
    Exit Function

TermErr:
    On Error Resume Next
    If previousMoved And Not saved Then
        If Not fso.FileExists(resultFilePath) Then Name backupFile As resultFilePath
    End If
    If Not wbResult Is Nothing Then wbResult.Close SaveChanges:=False
    If Len(tempPath) > 0 And Not saved Then
        If fso.FileExists(tempPath) Then fso.DeleteFile tempPath, True
    End If
    BuildSingleTermOutput = outBasename & ".xlsx -> ERROR: " & Err.Description
    Err.Clear
End Function

Private Function ConvertTermSheet(ByVal wsSource As Worksheet, ByVal wbResult As Workbook, _
                                  ByVal termNum As Long, ByRef tableIndex As Long, _
                                  ByRef reasonOut As String) As Long
    Dim wsResult As Worksheet, foundRange As Range
    Dim LastRow As Long, LastCol As Long, offsetCol As Long
    Dim srcData As Variant, resData As Variant
    Dim SourceCol As Long, ResultCol As Long, RowNum As Long, c As Long
    Dim resRows As Long, resCols As Long, newCols As Long
    Dim sheetName As String
    Dim tbl As ListObject, tblRange As Range
    Dim seenHeaders As Object, h As Long, headerText As String, genCounter As Long

    reasonOut = ""
    ConvertTermSheet = 0
    sheetName = wsSource.Name
    offsetCol = termNum - 1

    On Error GoTo SheetErr

    LastCol = wsSource.Cells(2, wsSource.Columns.Count).End(xlToLeft).Column
    Set foundRange = wsSource.Range(wsSource.Cells(3, 1), wsSource.Cells(wsSource.Rows.Count, LastCol)).Find( _
        What:="*", After:=wsSource.Cells(3, 1), LookIn:=xlFormulas, LookAt:=xlPart, _
        SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False, SearchFormat:=False)
    If Not foundRange Is Nothing Then LastRow = foundRange.Row

    If LastRow < 3 Or LastCol < FIRST_SUBJECT_COL + offsetCol Then Exit Function

    Set wsResult = wbResult.Worksheets.Add(After:=wbResult.Worksheets(wbResult.Worksheets.Count))

    On Error Resume Next
    wsResult.Name = Left(SanitizeSheetName(sheetName), 31)
    If Err.Number <> 0 Then wsResult.Name = "Sheet" & wbResult.Worksheets.Count: Err.Clear
    On Error GoTo SheetErr

    wsResult.DisplayRightToLeft = True
    srcData = wsSource.Range(wsSource.Cells(2, 1), wsSource.Cells(LastRow, LastCol)).Value

    resRows = LastRow - 1
    newCols = (LastCol - (FIRST_SUBJECT_COL - 1)) \ GROUP_SIZE
    resCols = (FIRST_SUBJECT_COL - 1) + newCols
    ReDim resData(1 To resRows, 1 To resCols)

    For RowNum = 1 To resRows
        For ResultCol = 1 To FIRST_SUBJECT_COL - 1
            resData(RowNum, ResultCol) = srcData(RowNum, ResultCol)
        Next ResultCol
    Next RowNum

    ResultCol = FIRST_SUBJECT_COL
    For SourceCol = FIRST_SUBJECT_COL + offsetCol To LastCol Step GROUP_SIZE
        resData(1, ResultCol) = StripTermSuffixByNum(srcData(1, SourceCol), termNum)
        ResultCol = ResultCol + 1
    Next SourceCol

    For RowNum = 2 To resRows
        ResultCol = FIRST_SUBJECT_COL
        For SourceCol = FIRST_SUBJECT_COL + offsetCol To LastCol Step GROUP_SIZE
            resData(RowNum, ResultCol) = srcData(RowNum, SourceCol)
            ResultCol = ResultCol + 1
        Next SourceCol
    Next RowNum

    With wsResult
        .Columns("B").NumberFormat = "@"
        .Columns("C").NumberFormat = "yyyy-mm-dd"
        .Columns("D:E").NumberFormat = "@"
        .Range(.Columns(6), .Columns(resCols)).NumberFormat = "0.00"
    End With

    wsResult.Range(wsResult.Cells(2, 1), wsResult.Cells(LastRow, resCols)).Value = resData

    Set seenHeaders = CreateObject("Scripting.Dictionary")
    seenHeaders.CompareMode = vbTextCompare
    genCounter = 0
    For h = 1 To resCols
        headerText = Trim(CStr(wsResult.Cells(2, h).Value))
        If headerText = "" Or seenHeaders.Exists(headerText) Then
            genCounter = genCounter + 1
            Do While seenHeaders.Exists("Column" & genCounter): genCounter = genCounter + 1: Loop
            wsResult.Cells(2, h).Value = "Column" & genCounter
            seenHeaders.Add "Column" & genCounter, True
        Else
            seenHeaders.Add headerText, True
        End If
    Next h

    Set tblRange = wsResult.Range(wsResult.Cells(2, 1), wsResult.Cells(LastRow, resCols))
    Set tbl = wsResult.ListObjects.Add(xlSrcRange, tblRange, , xlYes)
    tbl.Name = "Data" & Format(tableIndex + 1, "00")
    tbl.TableStyle = ""
    tbl.ShowTableStyleColumnStripes = False
    tbl.ShowTableStyleRowStripes = False
    tbl.ShowTableStyleFirstColumn = False
    tbl.ShowTableStyleLastColumn = False
    tbl.ShowAutoFilter = True

    For c = 1 To 5
        With wsResult.Range(wsResult.Cells(2, c), wsResult.Cells(LastRow, c))
            .Font.Name = wsSource.Cells(3, c).Font.Name
            .Font.Size = wsSource.Cells(3, c).Font.Size
        End With
    Next c
    ResultCol = FIRST_SUBJECT_COL
    For SourceCol = FIRST_SUBJECT_COL + offsetCol To LastCol Step GROUP_SIZE
        With wsResult.Range(wsResult.Cells(2, ResultCol), wsResult.Cells(LastRow, ResultCol))
            .Font.Name = wsSource.Cells(3, SourceCol).Font.Name
            .Font.Size = wsSource.Cells(3, SourceCol).Font.Size
        End With
        ResultCol = ResultCol + 1
    Next SourceCol
    With wsResult
        With .Range(.Cells(2, 1), .Cells(LastRow, resCols))
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .WrapText = False
        End With
        .Columns("A:E").AutoFit
        .Range(.Columns(FIRST_SUBJECT_COL), .Columns(resCols)).ColumnWidth = 5.7
        With .Range(.Cells(2, 4), .Cells(2, resCols))
            .Orientation = xlUpward
            .Font.Bold = True
        End With
        .Rows(2).RowHeight = wsSource.Rows(2).RowHeight
        .Rows(1).RowHeight = 30
    End With
    WriteTermTitle wsResult.Range("B1:F1"), TermTitleByNum(CStr(wsSource.Range("B1").Value), termNum), wsSource.Range("B1")
    WriteTermTitle wsResult.Range("H1:Q1"), CStr(wsSource.Range("H1").Value), wsSource.Range("H1")

    tableIndex = tableIndex + 1
    ConvertTermSheet = 1
    Exit Function

SheetErr:
    reasonOut = CStr(Err.Number) & " : " & Err.Description
    ConvertTermSheet = -1
    On Error Resume Next
    If Not wsResult Is Nothing Then wsResult.Delete
    Err.Clear
    On Error GoTo 0
End Function

Private Sub WriteTermTitle(ByVal target As Range, ByVal title As String, ByVal source As Range)
    With target
        .UnMerge: .ClearContents: .Merge: .Value = title
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
        .ReadingOrder = xlRTL: .WrapText = True
        .Font.Name = source.Font.Name: .Font.Size = 11: .Font.Bold = True
    End With
End Sub

Private Function StripTermSuffixByNum(ByVal v As Variant, ByVal termNum As Long) As Variant
    Dim re As Object, t As String, r As String, dChar As String
    If IsError(v) Or IsEmpty(v) Then StripTermSuffixByNum = v: Exit Function
    t = CStr(v)
    On Error GoTo Fail
    Set re = CreateObject("VBScript.RegExp")
    dChar = CStr(termNum) & ChrW(&H660 + termNum) & ChrW(&H6F0 + termNum)
    re.Pattern = "[\s\u00A0]*" & ChrW(1601) & "[\s\u00A0]*[" & dChar & "][\s\u00A0]*$"
    re.Global = False
    r = Trim(re.Replace(t, ""))
    If Len(r) = 0 Then r = Trim(t)
    StripTermSuffixByNum = r
    Exit Function
Fail:
    Err.Clear
    StripTermSuffixByNum = v
End Function

Private Function TermTitleByNum(ByVal sourceTitle As String, ByVal termNum As Long) As String
    Dim rx As Object, matches As Object, digits As String, fixedTitle As String
    Set rx = CreateObject("VBScript.RegExp")
    digits = "[0-9" & ChrW(&H660) & "-" & ChrW(&H669) & ChrW(&H6F0) & "-" & ChrW(&H6F9) & "]"
    rx.Pattern = digits & "{4}\s*[-/" & ChrW(&H2013) & ChrW(&H2014) & "]\s*" & digits & "{4}"
    rx.Global = False
    Set matches = rx.Execute(sourceTitle)
    If matches.Count = 0 Then
        Err.Raise vbObjectError + 720, "TermTitleByNum", "School year not found in source B1."
    End If
    Select Case termNum
        Case 1 ' نتائج الفصل الأول
            fixedTitle = ChrW(1606) & ChrW(1578) & ChrW(1575) & ChrW(1574) & ChrW(1580) & ChrW(32) & ChrW(1575) & ChrW(1604) & ChrW(1601) & ChrW(1589) & ChrW(1604) & ChrW(32) & ChrW(1575) & ChrW(1604) & ChrW(1571) & ChrW(1608) & ChrW(1604)
        Case 2 ' نتائج الفصل الثاني
            fixedTitle = ChrW(1606) & ChrW(1578) & ChrW(1575) & ChrW(1574) & ChrW(1580) & ChrW(32) & ChrW(1575) & ChrW(1604) & ChrW(1601) & ChrW(1589) & ChrW(1604) & ChrW(32) & ChrW(1575) & ChrW(1604) & ChrW(1579) & ChrW(1575) & ChrW(1606) & ChrW(1610)
        Case 3 ' نتائج الفصل الثالث
            fixedTitle = ChrW(1606) & ChrW(1578) & ChrW(1575) & ChrW(1574) & ChrW(1580) & ChrW(32) & ChrW(1575) & ChrW(1604) & ChrW(1601) & ChrW(1589) & ChrW(1604) & ChrW(32) & ChrW(1575) & ChrW(1604) & ChrW(1579) & ChrW(1575) & ChrW(1604) & ChrW(1579)
    End Select
    TermTitleByNum = fixedTitle & " " & matches(0).Value
End Function

Private Function SanitizeSheetName(ByVal s As String) As String
    Dim bad As Variant, i As Long
    bad = Array(":", "\", "/", "?", "*", "[", "]")
    For i = LBound(bad) To UBound(bad): s = Replace(s, bad(i), "_"): Next i
    SanitizeSheetName = s
End Function

Private Function TermWorkbookOpen(ByVal fullPath As String) As Boolean
    Dim wb As Workbook
    For Each wb In Application.Workbooks
        If StrComp(wb.FullName, fullPath, vbTextCompare) = 0 Then
            TermWorkbookOpen = True
            Exit Function
        End If
    Next wb
End Function

' ==============================================================================
' EMBEDDED SOURCE03 COLLECTION ENGINE
' ==============================================================================
Private Function FileExistsPath(ByVal p As String) As Boolean
    On Error Resume Next
    FileExistsPath = ((GetAttr(p) And vbDirectory) = 0)
    If Err.Number <> 0 Then FileExistsPath = False: Err.Clear
End Function

Private Function SortedNames(ByVal col As Collection) As Variant
    Dim arr() As String, i As Long, j As Long, tmp As String
    If col.Count = 0 Then SortedNames = Array(): Exit Function
    ReDim arr(1 To col.Count)
    For i = 1 To col.Count: arr(i) = CStr(col(i)): Next i
    For i = 2 To UBound(arr)
        tmp = arr(i): j = i - 1
        Do While j >= 1
            If StrComp(arr(j), tmp, vbTextCompare) <= 0 Then Exit Do
            arr(j + 1) = arr(j): j = j - 1
        Loop
        arr(j + 1) = tmp
    Next i
    SortedNames = arr
End Function

Private Function FontExists(ByVal FontName As String) As Boolean
    Const HKCU As Long = &H80000001
    Const HKLM As Long = &H80000002
    Dim reg As Object, hv As Variant, names As Variant, types As Variant
    Dim i As Long, p As Long, nm As String, enumOK As Boolean
    On Error GoTo Fallback
    Set reg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
    For Each hv In Array(HKLM, HKCU)
        names = Empty
        If reg.EnumValues(CLng(hv), "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts", names, types) = 0 Then
            enumOK = True
            If IsArray(names) Then
                For i = LBound(names) To UBound(names)
                    nm = CStr(names(i))
                    p = InStr(nm, " (")
                    If p > 0 Then nm = Left(nm, p - 1)
                    If StrComp(nm, FontName, vbTextCompare) = 0 Then FontExists = True: Exit Function
                    If StrComp(Left(FontName, Len(nm) + 1), nm & " ", vbTextCompare) = 0 Then FontExists = True: Exit Function
                Next i
            End If
        End If
    Next hv
    FontExists = Not enumOK
    Exit Function
Fallback:
    FontExists = True
End Function

Private Sub SetupFonts()
    If FontExists(FONT_ARABIC) Then g_ActiveArabicFont = FONT_ARABIC Else g_ActiveArabicFont = FONT_FALLBACK
    If FontExists(FONT_NUMERIC) Then g_ActiveNumericFont = FONT_NUMERIC Else g_ActiveNumericFont = FONT_FALLBACK
End Sub

Private Function SrcProcessFile(ByRef wbResult As Workbook, ByVal SourcePath As String, ByVal fileName As String, ByVal SheetNumber As Long, ByRef ReasonOut As String) As Boolean
    Dim wbSource As Workbook, wsSource As Worksheet, wsResult As Worksheet
    Dim LastRow As Long, LastCol As Long, DataArr As Variant
    Dim errNum As Long, errDesc As String, lastDataRow As Long, wrapEndRow As Long
    Dim idLastCol As Long, lastFind As Range
    Dim tbl As ListObject, tblRange As Range, seenHeaders As Object
    Dim h As Long, headerText As String, genCounter As Long

    SrcProcessFile = False: ReasonOut = ""
    On Error GoTo EH

    On Error Resume Next
    Set wbSource = Workbooks.Open(fileName:=SourcePath & fileName, UpdateLinks:=False, ReadOnly:=True, AddToMRU:=False, Password:="")
    errNum = Err.Number: errDesc = Err.Description: Err.Clear
    On Error GoTo EH

    If wbSource Is Nothing Or errNum <> 0 Then ReasonOut = "OPEN_FAILED | " & errDesc: Exit Function
    If wbSource.Sheets.Count = 0 Then ReasonOut = "EMPTY_WORKBOOK": wbSource.Close SaveChanges:=False: Exit Function

    Set wsSource = wbSource.Sheets(1)
    On Error Resume Next
    LastCol = wsSource.Cells(DATA_START_ROW, wsSource.Columns.Count).End(xlToLeft).Column
    Set lastFind = wsSource.Cells.Find(What:="*", LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    If lastFind Is Nothing Then LastRow = 0 Else LastRow = lastFind.Row - TRAILING_ROWS_SKIP
    errNum = Err.Number: errDesc = Err.Description: Err.Clear
    On Error GoTo EH

    If errNum <> 0 Or LastRow < DATA_START_ROW Or LastCol < 1 Then
        ReasonOut = "EMPTY_OR_INVALID"
        wbSource.Close SaveChanges:=False: Exit Function
    End If

    Set wsResult = wbResult.Sheets.Add(After:=wbResult.Sheets(wbResult.Sheets.Count))
    On Error Resume Next
    wsResult.Name = Format(SheetNumber, "00")
    Err.Clear
    On Error GoTo EH
    wsResult.DisplayRightToLeft = True

    DataArr = wsSource.Range(wsSource.Cells(DATA_START_ROW, 1), wsSource.Cells(LastRow, LastCol)).Value
    lastDataRow = 1 + UBound(DataArr, 1)

    With wsResult
        .Columns("B").NumberFormat = "@"
        .Columns("C").NumberFormat = "yyyy-mm-dd"
        If LastCol >= 4 Then .Columns("D").NumberFormat = "@"
        If LastCol >= 5 Then .Columns("E").NumberFormat = "@"
        If LastCol >= 6 Then .Range(.Columns("F"), .Columns(LastCol)).NumberFormat = "0.00"
    End With

    wsResult.Range("A2").Resize(UBound(DataArr, 1), UBound(DataArr, 2)).Value = DataArr

    Set seenHeaders = CreateObject("Scripting.Dictionary")
    seenHeaders.CompareMode = vbTextCompare
    For h = 1 To LastCol
        headerText = Trim(CStr(wsResult.Cells(2, h).Value))
        If headerText = "" Or seenHeaders.Exists(headerText) Then
            genCounter = genCounter + 1
            Do While seenHeaders.Exists("Column" & genCounter): genCounter = genCounter + 1: Loop
            wsResult.Cells(2, h).Value = "Column" & genCounter
            seenHeaders.Add "Column" & genCounter, True
        Else
            seenHeaders.Add headerText, True
        End If
    Next h

    Set tblRange = wsResult.Range(wsResult.Cells(2, 1), wsResult.Cells(lastDataRow, LastCol))
    On Error Resume Next
    Set tbl = wsResult.ListObjects.Add(xlSrcRange, tblRange, , xlYes)
    If Err.Number = 0 Then
        tbl.Name = "Data" & Format(SheetNumber, "00")
        tbl.TableStyle = ""
    End If
    Err.Clear
    On Error GoTo EH

    With wsResult
        .Columns("B").Font.Name = g_ActiveArabicFont: .Columns("B").Font.Size = FONT_SIZE
        .Columns("C").Font.Name = g_ActiveNumericFont: .Columns("C").Font.Size = FONT_SIZE
        If LastCol >= 4 Then .Columns("D").Font.Name = g_ActiveArabicFont: .Columns("D").Font.Size = FONT_SIZE
        If LastCol >= 5 Then .Columns("E").Font.Name = g_ActiveArabicFont: .Columns("E").Font.Size = FONT_SIZE
        If LastCol >= 6 Then
            With .Range(.Columns("F"), .Columns(LastCol))
                .Font.Name = g_ActiveNumericFont
                .Font.Size = FONT_SIZE
            End With
        End If
        wrapEndRow = IIf(lastDataRow < 2, 2, lastDataRow)
        .Range(.Cells(2, 1), .Cells(wrapEndRow, LastCol)).WrapText = False
        With .Range(.Cells(1, 1), .Cells(lastDataRow, LastCol))
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
        idLastCol = IIf(LastCol < IDENTITY_LAST_COL, LastCol, IDENTITY_LAST_COL)
        .Range(.Columns(1), .Columns(idLastCol)).EntireColumn.AutoFit
        If LastCol >= 6 Then .Range(.Columns(IDENTITY_LAST_COL + 1), .Columns(LastCol)).ColumnWidth = COL_WIDTH
        If LastCol >= 4 Then
            .Range(.Cells(2, 4), .Cells(2, LastCol)).Orientation = xlUpward
            .Range(.Cells(2, 4), .Cells(2, LastCol)).Font.Bold = True
        End If
    End With

    PlaceHeaderTitle wsSource, wsResult
    wbSource.Close SaveChanges:=False
    SrcProcessFile = True
    Exit Function
EH:
    ReasonOut = Err.Number & " : " & Err.Description
    Err.Clear
    On Error Resume Next
    If Not wbSource Is Nothing Then wbSource.Close SaveChanges:=False
    If Not wsResult Is Nothing Then wsResult.Delete
    SrcProcessFile = False
End Function

Private Function SrcBuildWorkbook(ByVal srcFolder As String, ByRef wbOut As Workbook, ByRef copied As Long, ByRef skipped As Long, ByRef skippedList As String) As Boolean
    Dim wb As Workbook, fileName As String, SheetNumber As Long, Reason As String
    Dim names As Collection, sorted As Variant, i As Long
    Dim f As String, ext As String, baseName As String, dotPos As Long

    SrcBuildWorkbook = False: Set wbOut = Nothing
    copied = 0: skipped = 0: SheetNumber = 1: skippedList = ""

    Set names = New Collection
    fileName = Dir(srcFolder & "*.xls*")
    Do While fileName <> ""
        If Left(fileName, 2) <> "~$" Then
            dotPos = InStrRev(fileName, ".")
            ext = LCase(Mid(fileName, dotPos + 1))
            If ext = "xls" Or ext = "xlsx" Or ext = "xlsm" Or ext = "xlsb" Then names.Add fileName
        End If
        fileName = Dir
    Loop

    If names.Count = 0 Then
        skippedList = "  - Source03 folder is empty." & vbCrLf
        Exit Function
    End If

    sorted = SortedNames(names)
    SetupFonts

    Set wb = Workbooks.Add(xlWBATWorksheet)
    For i = LBound(sorted) To UBound(sorted)
        f = CStr(sorted(i))
        dotPos = InStrRev(f, ".")
        ext = LCase(Mid(f, dotPos + 1))
        baseName = Left(f, dotPos - 1)

        If StrComp(srcFolder & f, ThisWorkbook.FullName, vbTextCompare) = 0 Then
            ' skip self
        ElseIf TermWorkbookOpen(srcFolder & f) Then
            skipped = skipped + 1
            skippedList = skippedList & "  - " & f & " : ALREADY OPEN in Excel" & vbCrLf
        ElseIf ext <> "xlsx" And FileExistsPath(srcFolder & baseName & ".xlsx") Then
            skipped = skipped + 1
            skippedList = skippedList & "  - " & f & " : DUPLICATE (.xlsx exists)" & vbCrLf
        ElseIf SrcProcessFile(wb, srcFolder, f, SheetNumber, Reason) Then
            copied = copied + 1: SheetNumber = SheetNumber + 1
        Else
            skipped = skipped + 1
            skippedList = skippedList & "  - " & f & " : " & Reason & vbCrLf
        End If
    Next i

    If copied > 0 Then
        If wb.Sheets.Count > 1 Then wb.Sheets(1).Delete
        Set wbOut = wb
        SrcBuildWorkbook = True
    Else
        wb.Close SaveChanges:=False
    End If
End Function

Private Sub PlaceHeaderTitle(ByVal wsSource As Worksheet, ByVal wsTarget As Worksheet)
    Dim rawText As String, partTitle As String, partGroup As String
    rawText = CleanText(CStr(wsSource.Range("A5").Value))
    SplitAtSchoolYear rawText, partTitle, partGroup
    WriteMergedCell wsTarget.Range("B1:F1"), partTitle
    WriteMergedCell wsTarget.Range("H1:Q1"), partGroup
    wsTarget.Rows(1).RowHeight = HEADER_ROW_HEIGHT
End Sub

Private Function CleanText(ByVal s As String) As String
    s = Replace(s, vbCr, " "): s = Replace(s, vbLf, " "): s = Replace(s, Chr(160), " ")
    s = Application.WorksheetFunction.Clean(s)
    CleanText = Application.WorksheetFunction.Trim(s)
End Function

Private Sub SplitAtSchoolYear(ByVal txt As String, ByRef part1 As String, ByRef part2 As String)
    Dim rx As Object, m As Object, splitPos As Long, digit As String, seps As String
    part1 = txt: part2 = ""
    If Len(txt) = 0 Then Exit Sub
    digit = "[0-9\u0660-\u0669]": seps = "[/\\\-\u2013\u2014]"
    Set rx = CreateObject("VBScript.RegExp")
    rx.Global = False: rx.IgnoreCase = True
    rx.Pattern = digit & "{4}\s*" & seps & "\s*" & digit & "{4}"
    If Not rx.Test(txt) Then
        rx.Pattern = digit & "{4}"
        If Not rx.Test(txt) Then Exit Sub
    End If
    Set m = rx.Execute(txt)
    splitPos = m(0).FirstIndex + m(0).Length
    part1 = Trim(Left(txt, splitPos))
    If splitPos < Len(txt) Then
        part2 = Trim(Mid(txt, splitPos + 1))
        Do While Len(part2) > 0
            If InStr(1, "-:/" & ChrW(&H2013) & ChrW(&H2014), Left(part2, 1)) = 0 Then Exit Do
            part2 = Trim(Mid(part2, 2))
        Loop
    End If
End Sub

Private Sub WriteMergedCell(ByVal rng As Range, ByVal txt As String)
    With rng
        .UnMerge: .ClearContents: .Merge: .Value = txt
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
        .WrapText = True: .ReadingOrder = xlRTL
        .Font.Name = GetAvailableFont(PREFERRED_FONT, FALLBACK_FONT)
        .Font.Size = HEADER_FONT_SIZE: .Font.Bold = True
    End With
End Sub

Private Function GetAvailableFont(ByVal wanted As String, ByVal fallback As String) As String
    If IsFontInstalled(wanted) Then GetAvailableFont = wanted Else GetAvailableFont = fallback
End Function

Private Function IsFontInstalled(ByVal fontName As String) As Boolean
    Dim folders(1 To 2) As String, i As Long
    folders(1) = Environ("WINDIR") & "\Fonts\"
    folders(2) = Environ("LOCALAPPDATA") & "\Microsoft\Windows\Fonts\"
    On Error Resume Next
    For i = 1 To 2
        If Len(Dir(folders(i) & fontName & "*.*")) > 0 Then IsFontInstalled = True: Exit Function
    Next i
    On Error GoTo 0
End Function
