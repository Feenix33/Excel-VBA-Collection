Attribute VB_Name = "SimpleMacros"
Option Explicit
Public iHeatMapG As Double
Public iHeatMapY As Double
Public iHeatMapR As Double
Public iHeatClrG As Long
Public iHeatClrY As Long
Public iHeatClrR As Long
Public giFrmRYGReturn As Long
Sub cmeSetAutoCalc()
    Application.Calculation = xlAutomatic
End Sub
Function cmeDate2Iter(dateIn As Date) As String
' Convert input date into an iteration number
    Dim dayMagic As Date
    Dim dayDiff, outYear As Integer
    Dim itNum As Integer
    
    dayMagic = #1/3/2018#
    outYear = 2018
    If dateIn < dayMagic Then
        dayMagic = #1/4/2017#
        outYear = 2017
    End If
    dayDiff = DateDiff("d", dayMagic, dateIn)
    itNum = Int((dayDiff / 14) + 0.5)
    cmeDate2Iter = Str(outYear) & "#S" & Format(itNum, "00")
End Function
Sub cmePasteValues()
Attribute cmePasteValues.VB_ProcData.VB_Invoke_Func = "V\n14"
' Keyboard Shortcut: Ctrl+Shift+V
    On Error GoTo Fini
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
Fini:
End Sub
Sub cmeFilterOnValue()
Attribute cmeFilterOnValue.VB_ProcData.VB_Invoke_Func = "l\n14"
    If Selection.Rows.Count > 1 Then Exit Sub
    If Selection.Columns.Count > 1 Then Exit Sub
    Range(ActiveCell.CurrentRegion.Address).AutoFilter Field:=ActiveCell.Column, Criteria1:=ActiveCell.Value
End Sub
Sub cmeFilterToggle()
Attribute cmeFilterToggle.VB_ProcData.VB_Invoke_Func = "O\n14"
' Keyboard Shortcut: Ctrl+Shift+O
    On Error GoTo errorFilterToggle
    Selection.AutoFilter
errorFilterToggle:
End Sub
Sub cmeNameSheet()
Attribute cmeNameSheet.VB_ProcData.VB_Invoke_Func = "N\n14"
' Keyboard Shortcut: Ctrl+Shift+N
    Dim newSheetName As String
    Dim oldSheetName As String
     
    oldSheetName = ActiveSheet.Name
    newSheetName = oldSheetName
    On Error GoTo errorNameSheet
    newSheetName = InputBox("Enter the new sheet name", "New Sheet Name", newSheetName)
    
    If Len(newSheetName) = 0 Then GoTo errorNameSheet
    newSheetName = ProcessSheetName(newSheetName)
    If Len(newSheetName) > 0 Then
        ActiveSheet.Name = newSheetName
        Exit Sub
    End If
    
errorNameSheet:
    ActiveSheet.Name = oldSheetName
End Sub
Function ProcessSheetName(inputSheetName As String) As String
    Dim newSheetName As String
    Dim pos As Integer
    newSheetName = inputSheetName
    
    pos = InStr(LCase(newSheetName), ">d")
    If pos > 0 Then
       newSheetName = Left(newSheetName, pos - 1) + Format(Date, "yyyy.mm.dd") + Right(newSheetName, Len(newSheetName) - pos - 1)
    End If
    
    pos = InStr(LCase(newSheetName), ">p")
    If pos > 0 Then
       newSheetName = Left(newSheetName, pos - 1) + "Pivot" + Right(newSheetName, Len(newSheetName) - pos - 1)
    End If
    
    pos = InStr(LCase(newSheetName), ">iter")
    If pos > 0 Then
       newSheetName = Left(newSheetName, pos - 1) + "Iteration" + Right(newSheetName, Len(newSheetName) - pos - 4)
    End If
    
    pos = InStr(LCase(newSheetName), ">it")
    If pos > 0 Then
       newSheetName = Left(newSheetName, pos - 1) + "Iteration" + Right(newSheetName, Len(newSheetName) - pos - 2)
    End If
    
    pos = InStr(LCase(newSheetName), ">i")
    If pos > 0 Then
       newSheetName = Left(newSheetName, pos - 1) + "Iter" + Right(newSheetName, Len(newSheetName) - pos - 1)
    End If
    
    pos = InStr(LCase(newSheetName), ">lh")
    If pos > 0 Then
       newSheetName = Left(newSheetName, pos - 1) + "Labor Hours" + Right(newSheetName, Len(newSheetName) - pos - 2)
    End If
    
    pos = InStr(LCase(newSheetName), ">h")
    If pos > 0 Then
       newSheetName = Left(newSheetName, pos - 1) + "Hierarchy" + Right(newSheetName, Len(newSheetName) - pos - 1)
    End If
    
    pos = InStr(LCase(newSheetName), ">an")
    If pos > 0 Then
       newSheetName = Left(newSheetName, pos - 1) + "Analysis" + Right(newSheetName, Len(newSheetName) - pos - 2)
    End If
    
    pos = InStr(LCase(newSheetName), ">m")
    If pos > 0 Then
       newSheetName = Left(newSheetName, pos - 1) + "Milestone" + Right(newSheetName, Len(newSheetName) - pos - 2)
    End If
    
    ProcessSheetName = CheckSheetName(newSheetName)
End Function
Function CheckSheetName(newSheetName As String)
    Dim wks As Worksheet
    Dim strName As String
    Dim iSuffix As Integer
    
    iSuffix = 65 ' 65 = A
    CheckSheetName = newSheetName
RestartCheck:
    For Each wks In Worksheets
        strName = wks.Name
        If CheckSheetName = strName Then
           CheckSheetName = newSheetName & Chr(iSuffix)
           iSuffix = iSuffix + 1
           GoTo RestartCheck
        End If
    Next wks
    
    'CheckSheetName = newSheetName
End Function
Sub cmeWrapToggle()
Attribute cmeWrapToggle.VB_ProcData.VB_Invoke_Func = "W\n14"
' cmeWrapToggle Macro
' Keyboard Shortcut: Ctrl+Shift+W
    Selection.WrapText = Not Selection.Cells(1, 1).WrapText
End Sub
Sub cmdAutoSize()
Attribute cmdAutoSize.VB_ProcData.VB_Invoke_Func = "j\n14"
' cmdAutoSize Macro
' Keyboard Shortcut: Ctrl+Shift+J
    Dim rngOrigSelect As Range
    Dim rngOrigCell As Range
     
    Set rngOrigSelect = Selection
    Set rngOrigCell = ActiveCell
    Cells.Select
    Cells.EntireColumn.AutoFit
    Cells.EntireRow.AutoFit
    rngOrigSelect.Select
    rngOrigCell.Activate
End Sub

Sub cmeAutoLimit()
Attribute cmeAutoLimit.VB_ProcData.VB_Invoke_Func = "L\n14"
' Auto size the cells, but then loop through and limit the width
' only makes stuff smaller, doesn't make them bigger
    Dim myWidth As Integer
    On Error GoTo Fini
    myWidth = 80
    myWidth = Int(InputBox("Maximum width", "User Input", 60))
    cmeAutoLimitProcess (myWidth)
Fini:
End Sub
Sub cmeAutoLimitProcess(myWidth As Integer)
    Dim rngOrigSelect As Range ' reset the original selection
    Dim rngOrigCell As Range
         
    Set rngOrigSelect = Selection
    Set rngOrigCell = ActiveCell
    On Error GoTo Fini
    
    Cells.Select
    ' this part format cells to the top because i like it that way
    With Selection
        .VerticalAlignment = xlTop
        .WrapText = False ' if not turned off, won't autofit bigger
    End With

    Cells.EntireColumn.AutoFit
    Cells.EntireRow.AutoFit
    
    If (myWidth <= 0) Then
        GoTo Fini
    End If
    
    Dim lastColumn As Long
    With ActiveSheet.UsedRange
        lastColumn = .Columns(.Columns.Count).Column
    End With

    Dim c As Integer
    For c = 1 To lastColumn
        If (Cells(1, c).ColumnWidth > myWidth) Then
            Cells(1, c).ColumnWidth = myWidth
            Columns(c).WrapText = True
        End If
    Next c
Fini:
    rngOrigSelect.Select
    rngOrigCell.Activate
End Sub
Sub cmeFreeze()
Attribute cmeFreeze.VB_ProcData.VB_Invoke_Func = "F\n14"
'
' cmeFreeze Macro
'
' Keyboard Shortcut: Ctrl+Shift+W
'
    'ActiveWindow.FreezePanes = True
    ActiveWindow.FreezePanes = Not ActiveWindow.FreezePanes
End Sub
Sub PivotToggleCountSum()
'PURPOSE: Toggles between Counting and Summing Pivot Table data columns from current cell selection
'SOURCE: www.TheSpreadsheetGuru.com
    
Dim pf As PivotField
Dim AnyPFs As Boolean
Dim cell As Range

AnyPFs = False

'Optimize Code
  Application.ScreenUpdating = False
  Application.Calculation = xlCalculationManual

'Cycle through first row of selected cells
  For Each cell In Selection.Rows(1).Cells
    On Error Resume Next
      Set pf = cell.PivotField
    On Error GoTo 0
    
    If Not pf Is Nothing Then
      'Toggle between Counting and Summing
        pf.Function = xlCount + xlSum - pf.Function
      
      'No need for error message
        AnyPFs = True
      
      'Reset pf variable value
        Set pf = Nothing
    End If
  Next cell

'Did user select cells inside a Pivot Field?
  If AnyPFs = False Then MsgBox "There were no cells inside a Pivot Field selected."

'Optimize Code
  Application.Calculation = xlCalculationAutomatic
  Application.ScreenUpdating = True

End Sub
Sub combinationFilter()
Attribute combinationFilter.VB_ProcData.VB_Invoke_Func = "l\n14"
    Dim cell As Range, tableObj As ListObject, subSelection As Range
    Dim filterCriteria() As String, filterFields() As Integer
    Dim i As Integer
    
    'If the selection is in a table and one row height
        
    If Not Selection.ListObject Is Nothing And Selection.Rows.Count = 1 Then
        Set tableObj = ActiveSheet.ListObjects(Selection.ListObject.Name)
        
        i = 1
        ReDim filterCriteria(1 To Selection.Cells.Count) As String
        ReDim filterFields(1 To Selection.Cells.Count) As Integer
        
        ' handle multi-selects
        
        For Each subSelection In Selection.Areas
            For Each cell In subSelection
                filterCriteria(i) = cell.Text
                filterFields(i) = cell.Column - tableObj.Range.Cells(1, 1).Column + 1
                i = i + 1
            Next cell
        Next subSelection
        
        With tableObj.Range
            For i = 1 To UBound(filterCriteria)
                .AutoFilter Field:=filterFields(i), Criteria1:=filterCriteria(i)
            Next i
        End With
        Set tableObj = Nothing
    End If
End Sub
'Sub cmeCalendar()
'    Dim rows As Integer
'    Dim cols As Integer
'    Dim aDay As Date
'    Dim c
'
'
'    rows = Selection.rows.Count
'    cols = Selection.Columns.Count
'    Selection.Resize(rows, 7).Select
'    aDay = Date - Weekday(Date) + 1
'    ActiveCell.Value = aDay
'    For Each c In Selection.Cells
'        c.Value = aDay
'        aDay = aDay + 1
'    Next
'End Sub
Sub cmeTableFormat()
    Dim PvtTbl As PivotTable
    Dim pvtFld As PivotField
    
    Set PvtTbl = ActiveSheet.PivotTables(1)
    
    'hide Subtotals for all fields in the PivotTable .
    With PvtTbl
     For Each pvtFld In .PivotFields
        pvtFld.Subtotals(1) = True
        pvtFld.Subtotals(1) = False
        Next pvtFld
        .ColumnGrand = False
        .RowGrand = False
    End With
    
    ' format in report format
    PvtTbl.RowAxisLayout xlTabularRow
    ' random fun on the style, changes every day
    PvtTbl.TableStyle2 = cmeMagicPivotStyle '"PivotStyleMedium" & Weekday(Date)
End Sub
Sub SaveRallyExport()
' Saves file to a specific filename in a specific directory
    Dim saveDir As String, saveBaseName As String
    Dim savePath As String, saveExt As String
        
    saveDir = "C:\Users\sg0213341\Documents\Rally Exports\"
    saveBaseName = "Rally.Export."
    saveExt = ".xlsx"
    
    Call GenericVersionSave(saveDir, saveBaseName, saveExt)
End Sub
Sub SaveRallyLGSExport()
' Saves file to a specific filename in a specific directory
    Dim saveDir As String, saveBaseName As String
    Dim savePath As String, saveExt As String
        
    saveDir = "C:\Users\sg0213341\Documents\Rally LGS\"
    saveBaseName = "RallyLGS.Export."
    saveExt = ".xlsx"
    
    Call GenericVersionSave(saveDir, saveBaseName, saveExt)
End Sub
Sub GenericVersionSave(saveDir As String, saveBaseName As String, saveExt As String)
    Dim saveVer As String
    Dim bSaved As Boolean
    Dim savePath As String
    
    If Right(saveDir, 1) <> "\" Then saveDir = saveDir + "\"
    If Right(saveBaseName, 1) <> "." Then saveBaseName = saveBaseName + "."
    Call RevisionVersionSave(saveDir + saveBaseName + Format(Date, "yyyy.mm.dd"), saveExt)
End Sub
Sub RevisionVersionSave(saveFile As String, saveExt As String)
    'pass in a base filename w/path and the ext and auto add the letter
     Dim saveVer As String
    Dim savePath As String
    Dim bSaved As Boolean
    
    bSaved = False
    saveVer = ""
    Do While Not bSaved
        savePath = saveFile + saveVer + saveExt
        If Dir(savePath) <> "" Then
            'file exists, update version
            If saveVer = "" Then
                saveVer = "A"
            Else
                saveVer = Chr(Asc(saveVer) + 1)    ' fails at Z
            End If
        Else
            ActiveWorkbook.SaveAs savePath, FileFormat:=xlOpenXMLWorkbook
            bSaved = True
        End If
    Loop
End Sub
Sub UpdateRev()
    Dim saveDir As String
    Dim saveBaseName As String
    Dim saveExt As String
    
    On Error GoTo Fini
    saveDir = ActiveWorkbook.Path
    saveBaseName = StripRev(ActiveWorkbook.Name)
    saveExt = ".xlsx"
    Call RevisionVersionSave(saveDir + "\" + saveBaseName, saveExt)
Fini:
End Sub
'Sub VersionSave(savePathFile As String)
'    Dim saveVer As String
'    Dim bSaved As Boolean
'    Dim savePath As String
'    Dim saveExt As String
'
'    saveExt = ".xlsx"
'    bSaved = False
'    saveVer = ""
'    Do While Not bSaved
'        savePath = savePathFile + Format(Date, "yyyy.mm.dd") + saveVer + saveExt
'        If Dir(savePath) <> "" Then
'            'file exists, update version
'            If saveVer = "" Then
'                saveVer = "A"
'            Else
'                saveVer = Chr(Asc(saveVer) + 1)    ' fails at Z
'            End If
'        Else
'            ActiveWorkbook.SaveAs savePath, FileFormat:=xlOpenXMLWorkbook
'            bSaved = True
'        End If
'    Loop
'End Sub
'Sub SaveBusinessObjectsReport()
'' Saves file to a specific filename in a specific directory
'    Dim saveDir As String, saveBaseName As String
'    Dim savePath As String, saveExt As String
'    Dim reportName As String
'
'    reportName = ActiveWorkbook.Name
'    If Left(reportName, Len("BusObj.")) = "BusObj." Then
'        reportName = Mid(reportName, Len("BusObj.") + 1, 255)
'        If InStr(reportName, ".") > 0 Then
'            reportName = Left(reportName, InStr(reportName, ".") - 1)
'        End If
'    ElseIf InStr(reportName, ".") > 0 Then 'use the part before the dot
'    'guess the first part and remove the date
'        reportName = Left(reportName, InStr(reportName, ".") - 1)
'    Else
'        reportName = "BObj"
'    End If
'    If InStr(reportName, " ") > 0 Then ' check if space in name
'        reportName = Left(reportName, InStr(reportName, " ") - 1)
'    End If
'    saveBaseName = InputBox("Which Report?", "Business Object Save", reportName)
'    If Len(saveBaseName) = 0 Then
'        Exit Sub
'    End If
'
'    saveDir = "C:\Users\sg0213341\Documents\BusObjReports\"
'    saveBaseName = "BusObj." + saveBaseName
'    saveExt = ".xlsx"
'
'    Call GenericVersionSave(saveDir, saveBaseName, saveExt)
'End Sub
'Sub SaveJiraReport()
'' Saves file to a specific filename in a specific directory
'    Dim saveDir As String, saveBaseName As String
'    Dim savePath As String, saveExt As String
'
'    saveDir = "C:\Users\sg0213341\Documents\Exports\"
'    saveBaseName = "Jira.Export."
'    saveExt = ".xlsx"
'
'    Call GenericVersionSave(saveDir, saveBaseName, saveExt)
'End Sub
Sub SaveAsWithDate()
' Saves file to a user specified filename in my documents folder
    Dim saveDir As String, saveBaseName As String
    Dim savePath As String, saveExt As String
        
    'improve by getting the current workbook name
    'saveDir = "C:\Users\sg0213341\Documents\"
    saveDir = GetMyDirectory() + "\"
    If Len(saveDir) = 1 Then GoTo Fini
    Dim strCurrentName As String
    strCurrentName = StripDate(ActiveWorkbook.Name)
    saveBaseName = InputBox("Base Filename", "File Plus Date", strCurrentName)
    saveExt = ".xlsx"
    
    On Error GoTo Fini
    If Len(saveBaseName) = 0 Then
        GoTo Fini
    End If
    If Right(saveBaseName, 1) <> "." Then
        saveBaseName = saveBaseName + "."
    End If
    Call GenericVersionSave(saveDir, saveBaseName, saveExt)
Fini:
    'Debug.Print (saveBaseName)
End Sub
Function StripRev(inName As String) As String
'   Strip off the ext and if the name has a date, strip off any rev number from the date
'   first the ext
    Dim strDate As String
    
    On Error GoTo Fini
    inName = Left(inName, Len(inName) - 5) ' assume end is .xlsx=5 char
    StripRev = inName
    If Len(inName) <= 10 Then
        Exit Function
    Else
        strDate = Right(inName, 10)
        If strDate Like "####?##?##" Then
            Exit Function
        Else
            strDate = Right(inName, 11)
            If strDate Like "####?##?##?" Then
                StripRev = Left(StripRev, Len(StripRev) - 1)
                Exit Function
            End If
        End If
    End If
Fini:
End Function
Function StripDate(inName As String) As String
    If Len(inName) < 15 Then 'can't hold the date if too short (date + .xlsx ext)
        StripDate = inName
        Exit Function
    End If
    Dim btest As Boolean
    Dim strDate As String
    
    strDate = Left(inName, Len(inName) - 5) 'get rid of the ext
    strDate = Right(strDate, 10)
    'Debug.Print (strDate)
    If strDate Like "####?##?##" Then
        StripDate = Left(inName, Len(inName) - 5 - 10)
        If Right(StripDate, 1) = "." Then
            StripDate = Left(StripDate, Len(StripDate) - 1)
        End If
        Exit Function
    End If
    'try again for if there is an date letter extension
    strDate = Left(inName, Len(inName) - 6) 'get rid of the ext plus the date ext
    strDate = Right(strDate, 10)
    'Debug.Print (strDate)
    If strDate Like "####?##?##" Then
        StripDate = Left(inName, Len(inName) - 6 - 10)
        If Right(StripDate, 1) = "." Then
            StripDate = Left(StripDate, Len(StripDate) - 1)
        End If
        Exit Function
    End If
    
End Function
Function GetMyDirectory()
    Dim fDialogue As FileDialog
    'Set fDialogue = Application.FileDialog(msoFileDialogFilePicker)
    Set fDialogue = Application.FileDialog(msoFileDialogFolderPicker)
    
    'fDialogue.Filters.Add "Excel files", "*.xlsx"
    'fDialogue.Filters.Add "All files", "*.*"
    'fDialogue.InitialFileName = "C:\"
    'Debug.Print ("Path = " + ActiveWorkbook.Path)
    'Debug.Print ("Cur = " + CurDir())
    If Len(ActiveWorkbook.Path) = 0 Then
        fDialogue.InitialFileName = CurDir()
    Else
        fDialogue.InitialFileName = ActiveWorkbook.Path
        fDialogue.InitialFileName = "C:\Users\sg0213341\Documents\"
    End If
    If fDialogue.Show = -1 Then
       GetMyDirectory = fDialogue.SelectedItems(1)
    Else
       GetMyDirectory = ""
    End If
End Function
'Sub cmeBObjLaborReportPreparation()
'    Dim tbl As ListObject
'    Dim rng As Range
'    Dim iStyle As Integer
'
'    Sheets("Labor Details").Select
'    Range("B2").Select
'    Range(Selection, Selection.End(xlDown)).Select
'    Range(Selection, Selection.End(xlToRight)).Select
'    Selection.Copy
'    Sheets.Add After:=ActiveSheet
'    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
'        :=False, Transpose:=False
'    Range("A1").Select
'    ActiveWorkbook.ActiveSheet.rows(1).Find("Week Ending").Select
'    ActiveCell.EntireColumn.Select
'    Selection.NumberFormat = "dd-mmm-yy"
'
'    ActiveSheet.Name = ProcessSheetName(">L >D")
'    Set rng = Range(Range("A1"), Range("A1").SpecialCells(xlLastCell))
'    Set tbl = ActiveSheet.ListObjects.Add(xlSrcRange, rng, , xlYes)
'    'tbl.TableStyle = "TableStyleMedium" & Weekday(Date)
''    iStyle = Day(Date) '28 styles available
''    If iStyle > 28 Then iStyle = iStyle - 28
''    tbl.TableStyle = "TableStyleMedium" & Day(Date)
'    tbl.TableStyle = cmeMagicTableStyle
'    cmeAutoLimitProcess (60)
'    Range("A1").Select
'End Sub
Function cmeMagicTableStyle() As String
    Dim iStyle As Integer
    iStyle = Day(Date) '28 styles available
    If iStyle <= 28 Then
        cmeMagicTableStyle = "TableStyleMedium" & iStyle
    Else
        cmeMagicTableStyle = "TableStyleDark" & (iStyle - 28)
    End If
End Function
Function cmeMagicPivotStyle() As String
    Dim iStyle As Integer
    iStyle = Day(Date) '28 styles available
    If iStyle <= 28 Then
        cmeMagicPivotStyle = "PivotStyleMedium" & iStyle
    Else
        cmeMagicPivotStyle = "PivotStyleDark" & (iStyle - 28)
    End If
End Function
Sub MagicTablePreparation()
    Dim tbl As ListObject
    Dim rng As Range
    Dim arrMagicChange As Variant
    
    arrMagicChange = Array("now", "outpu", "sheet", "expor")
  
    Set rng = Range(Range("A1"), Range("A1").SpecialCells(xlLastCell))
    Set tbl = ActiveSheet.ListObjects.Add(xlSrcRange, rng, , xlYes)

    ' cme magic table formatter
    tbl.TableStyle = cmeMagicTableStyle
    
    Range("A1").Select
    cmeAutoLimitProcess (60) ' resize the sheet
'    If IsNumeric(Application.Match(LCase(Left(ActiveSheet.Name, 5)), arrMagicChange, 0)) Then
'        ActiveSheet.Name = ProcessSheetName(">d") ' name the sheet for today
'    End If
End Sub
Sub cmeWIPTable()
    Dim objTable As PivotTable
    Dim objfield As PivotField
    
    ActiveSheet.Range("b2").Select
    Set objTable = ActiveSheet.PivotTableWizard
    
    Set objfield = objTable.PivotFields("Schedule State")
    objfield.Orientation = xlColumnField
    
    Set objfield = objTable.PivotFields("Project")
    objfield.Orientation = xlRowField
    
    Set objfield = objTable.PivotFields("Plan Estimate")
    objfield.Orientation = xlDataField
    objfield.Function = xlSum
    
    Set objfield = objTable.PivotFields("Plan Estimate")
    objfield.Orientation = xlPageField
    objfield.PivotItems("0").Visible = False
    objfield.PivotItems("(blank)").Visible = False
    objTable.TableStyle2 = cmeMagicPivotStyle ' "PivotStyleMedium" & Weekday(Date)
    
    ActiveSheet.Name = ProcessSheetName(">P WIP") ' name the sheet Pivot WIP
End Sub
Sub cmeAddRallyType()
    Dim strActiveTable As String
    Dim ptrTable As ListObject
    
    strActiveTable = ActiveCell.ListObject.Name
    
    If HeaderExists(strActiveTable, "Type") = True Then Exit Sub
    
    Set ptrTable = ActiveSheet.ListObjects(strActiveTable)
    ptrTable.ListColumns.Add
    ptrTable.ListColumns(ptrTable.ListColumns.Count).Name = "Type"
    On Error GoTo AltName
        ptrTable.ListColumns("Type").DataBodyRange.FormulaR1C1 = "=LEFT([@[Formatted ID]],2)"
        Exit Sub
AltName:
    On Error GoTo Fini
        ptrTable.ListColumns("Type").DataBodyRange.FormulaR1C1 = "=LEFT([@[FormattedID]],2)"
Fini:
End Sub
Sub cmeAddRallyDone()
    Dim strActiveTable As String
    Dim ptrTable As ListObject
    
    strActiveTable = ActiveCell.ListObject.Name
    
    If HeaderExists(strActiveTable, "Done") = True Then Exit Sub
    
    Set ptrTable = ActiveSheet.ListObjects(strActiveTable)
    ptrTable.ListColumns.Add
    ptrTable.ListColumns(ptrTable.ListColumns.Count).Name = "Done"
    On Error GoTo AltName
        ptrTable.ListColumns("Done").DataBodyRange.FormulaR1C1 = "=OR([Schedule State]=""Accepted"",[Schedule State]=""Completed"",[Schedule State]=""Released-to-Production"")"
        Exit Sub
AltName:
    On Error GoTo Fini
        ptrTable.ListColumns("Done").DataBodyRange.FormulaR1C1 = "=OR([ScheduleState]=""Accepted"",[ScheduleState]=""Completed"",[ScheduleState]=""Released-to-Production"")"
Fini:
End Sub
Sub cmeAddRallyIterSort()
    Dim strActiveTable As String
    Dim ptrTable As ListObject
    
    strActiveTable = ActiveCell.ListObject.Name
    
    If HeaderExists(strActiveTable, "Iteration.Sortable") = True Then Exit Sub
    
    Set ptrTable = ActiveSheet.ListObjects(strActiveTable)
    ptrTable.ListColumns.Add
    ptrTable.ListColumns(ptrTable.ListColumns.Count).Name = "Iteration.Sortable"
    On Error GoTo AltName
        ptrTable.ListColumns("Iteration.Sortable").DataBodyRange.FormulaR1C1 = "=IF(LEN([Iteration])>9,MID([@[Iteration]],5,4)&""#""&LEFT([@[Iteration]],3),"""")"
        Exit Sub
AltName:
    On Error GoTo Fini
        ptrTable.ListColumns("Iteration.Sortable").DataBodyRange.FormulaR1C1 = "=IF(LEN([Iteration.Name])>9,MID([@[Iteration.Name]],5,4)&""#""&LEFT([@[Iteration.Name]],3),"""")"
Fini:
End Sub
Sub cmeAddTFID()
    Dim strActiveTable As String
    Dim ptrTable As ListObject
    
    strActiveTable = ActiveCell.ListObject.Name
    
    If HeaderExists(strActiveTable, "TF.ID") = True Then Exit Sub
    
    Set ptrTable = ActiveSheet.ListObjects(strActiveTable)
    ptrTable.ListColumns.Add
    ptrTable.ListColumns(ptrTable.ListColumns.Count).Name = "TF.ID"
    On Error GoTo AltName
        ptrTable.ListColumns("TF.ID").DataBodyRange.FormulaR1C1 = "=MID([@[Team Feature]],14,FIND("":"",[@[Team Feature]])-14)"
        Exit Sub
AltName:
    On Error GoTo Fini
Fini:
End Sub
Sub cmeAddRallyExtras()
    cmeAddRallyType
    cmeAddRallyDone
    cmeAddRallyIterSort
End Sub
Sub cmeUniqueCount()
'
' cmeUniqueCount Macro
'

'
    With ActiveSheet.PivotTables("PivotTable1").PivotFields( _
        "[Measures].[Count of FEA.ID]")
        .Caption = "Distinct Count of FEA.ID"
        .Function = xlDistinctCount
    End With
End Sub
Public Function HeaderExists(TableName As String, HeaderName As String) As Boolean
'PURPOSE: Output a true value if column name exists in specified table
'SOURCE: www.TheSpreadsheetGuru.com
Dim tbl As ListObject
Dim hdr As ListColumn

On Error GoTo DoesNotExist
  Set tbl = ActiveSheet.ListObjects(TableName)
  Set hdr = tbl.ListColumns(HeaderName)
On Error GoTo 0

HeaderExists = True

Exit Function

'Error Handler
DoesNotExist:
  Err.Clear
  HeaderExists = False

End Function
Sub Combine()
    Dim J As Integer

    On Error Resume Next
    Sheets(1).Select
    Worksheets.Add ' add a sheet in first place
    Sheets(1).Name = "Combined"

    ' copy headings
    Sheets(2).Activate
    Range("A1").EntireRow.Select
    Selection.Copy Destination:=Sheets(1).Range("A1")

    ' work through sheets
    For J = 2 To Sheets.Count ' from sheet 2 to last sheet
        Sheets(J).Activate ' make the sheet active
        Range("A1").Select
        Selection.CurrentRegion.Select ' select all cells in this sheets

        ' select all lines except title
        Selection.Offset(1, 0).Resize(Selection.Rows.Count - 1).Select

        ' copy cells selected in the new sheet on last line
        Selection.Copy Destination:=Sheets(1).Range("A65536").End(xlUp)(2)
    Next
End Sub

Sub cmeFmtFocus()
    Dim rngGrid As Range
    With ActiveSheet.UsedRange
        Set rngGrid = Range(.Cells(2, 2), .Cells(1, 1).Offset(.Rows.Count - 1, .Columns.Count - 1))
        rngGrid.Select
    End With
    
    ' RYG Rule
    Selection.FormatConditions.AddColorScale ColorScaleType:=3
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    Selection.FormatConditions(1).ColorScaleCriteria(1).Type = xlConditionValueNumber
    Selection.FormatConditions(1).ColorScaleCriteria(1).Value = 2
    With Selection.FormatConditions(1).ColorScaleCriteria(1).FormatColor
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).ColorScaleCriteria(2).Type = xlConditionValueNumber
    Selection.FormatConditions(1).ColorScaleCriteria(2).Value = 3
    With Selection.FormatConditions(1).ColorScaleCriteria(2).FormatColor
        .ThemeColor = xlThemeColorAccent4
        .TintAndShade = 0.599993896298105
    End With
    Selection.FormatConditions(1).ColorScaleCriteria(3).Type = xlConditionValueNumber
    Selection.FormatConditions(1).ColorScaleCriteria(3).Value = 4
    With Selection.FormatConditions(1).ColorScaleCriteria(3).FormatColor
        .Color = 6579450
        .TintAndShade = 0
    End With

    ' Header and first col
    Range("A1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Style = "40% - Accent1"
    Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Style = "40% - Accent3"
    Range("A1").Select
    cmeAutoLimitProcess (40)
End Sub
Sub cmeTabulateImportData()
' Convert imported data to columns
    Range("A1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=True, Tab:=True, _
        Semicolon:=True, Comma:=True, Space:=True, Other:=False, FieldInfo:= _
        Array(Array(1, 1), Array(2, 1), Array(3, 1)), TrailingMinusNumbers:=True
    Rows("1:1").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "TF.ID"
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "FEA.ID"
    Range("C1").Select
    Selection.FormulaR1C1 = "FEA.Name"
End Sub
Public Sub PivotFieldsToDistinctCount()
    Dim pf As PivotField
    With Selection.PivotTable
       For Each pf In .DataFields
        With pf
            .Function = xlDistinctCount
        End With
       Next pf
    End With
End Sub
Sub cmeHeatMapPivot()
    'MsgBox "HeatG = [" + Str(iHeatMapG) + "]"
    If (iHeatClrG = 0) Or (iHeatClrY = 0) Or (iHeatClrR = 0) Then
        iHeatMapG = 2
        iHeatMapY = 3
        iHeatMapR = 4
        iHeatClrG = RGB(99, 190, 123)
        iHeatClrY = RGB(255, 235, 132)
        iHeatClrR = RGB(248, 105, 107)
    End If
    frmRYG.Show

    If giFrmRYGReturn = 0 Then
        GoTo Fini
    End If
    
    Dim pPvtTbl As PivotTable
    Set pPvtTbl = ActiveSheet.PivotTables(1)
    
    pPvtTbl.DataBodyRange.Select
    Selection.FormatConditions.AddColorScale ColorScaleType:=3
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    Selection.FormatConditions(1).ColorScaleCriteria(1).Type = xlConditionValueNumber
    Selection.FormatConditions(1).ColorScaleCriteria(1).Value = iHeatMapG
    Selection.FormatConditions(1).ColorScaleCriteria(1).FormatColor.Color = iHeatClrG
    Selection.FormatConditions(1).ColorScaleCriteria(2).Type = xlConditionValueNumber
    Selection.FormatConditions(1).ColorScaleCriteria(2).Value = iHeatMapY
    Selection.FormatConditions(1).ColorScaleCriteria(2).FormatColor.Color = iHeatClrY
    Selection.FormatConditions(1).ColorScaleCriteria(3).Type = xlConditionValueNumber
    Selection.FormatConditions(1).ColorScaleCriteria(3).Value = iHeatMapR
    Selection.FormatConditions(1).ColorScaleCriteria(3).FormatColor.Color = iHeatClrR
    Range("A1").Select
Fini:
End Sub
Sub cmeFmtThreeColor()
    If (iHeatClrG = 0) Or (iHeatClrY = 0) Or (iHeatClrR = 0) Then
        iHeatMapG = 2
        iHeatMapY = 3
        iHeatMapR = 4
        iHeatClrG = RGB(99, 190, 123)
        iHeatClrY = RGB(255, 235, 132)
        iHeatClrR = RGB(248, 105, 107)
    End If
    Dim rngGrid As Range
    With ActiveSheet.UsedRange
        Set rngGrid = Range(.Cells(2, 2), .Cells(1, 1).Offset(.Rows.Count - 1, .Columns.Count - 1))
        rngGrid.Select
    End With
    frmRYG.Show
    If giFrmRYGReturn = 0 Then GoTo Fini
    
    Selection.FormatConditions.AddColorScale ColorScaleType:=3
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    Selection.FormatConditions(1).ColorScaleCriteria(1).Type = xlConditionValueNumber
    Selection.FormatConditions(1).ColorScaleCriteria(1).Value = iHeatMapG
    Selection.FormatConditions(1).ColorScaleCriteria(1).FormatColor.Color = iHeatClrG
    Selection.FormatConditions(1).ColorScaleCriteria(2).Type = xlConditionValueNumber
    Selection.FormatConditions(1).ColorScaleCriteria(2).Value = iHeatMapY
    Selection.FormatConditions(1).ColorScaleCriteria(2).FormatColor.Color = iHeatClrY
    Selection.FormatConditions(1).ColorScaleCriteria(3).Type = xlConditionValueNumber
    Selection.FormatConditions(1).ColorScaleCriteria(3).Value = iHeatMapR
    Selection.FormatConditions(1).ColorScaleCriteria(3).FormatColor.Color = iHeatClrR
    Range("A1").Select
Fini:
End Sub
Sub CombineWithNames()
    Dim J As Integer

    On Error Resume Next
    Sheets(1).Select
    Worksheets.Add ' add a sheet in first place
    Sheets(1).Name = "Combined"

    ' copy headings
    Sheets(2).Activate
    Range("A1").EntireRow.Select
    Selection.Copy Destination:=Sheets(1).Range("A1")

    ' work through sheets
    For J = 2 To Sheets.Count ' from sheet 2 to last sheet
        Sheets(J).Activate ' make the sheet active
        Range("A1").Select
        Selection.CurrentRegion.Select ' select all cells in this sheets

        ' select all lines except title
        Selection.Offset(1, 0).Resize(Selection.Rows.Count - 1).Select

        ' copy cells selected in the new sheet on last line
        Selection.Copy Destination:=Sheets(1).Range("A65536").End(xlUp)(2)
    Next
End Sub
Sub cmeVerticalText()
' cmeVerticalText Macro
    If Selection.Orientation = -4128 Then
        Selection.Orientation = 90
    Else
        Selection.Orientation = 0
    End If
End Sub
Sub cmeOneDigit()
' cmeOneDigit Macro
    On Error GoTo Fini
    Dim myDigit As Integer
    
    myDigit = 1
    myDigit = Int(InputBox("Number of Digits", "User Input", 1))
    If myDigit = 1 Then
        Selection.NumberFormat = "0.0"
    ElseIf myDigit = 3 Then
        Selection.NumberFormat = "0.000"
    ElseIf myDigit = 0 Then
        Selection.NumberFormat = "0"
    Else
        Selection.NumberFormat = "0.00"
    End If
Fini:
End Sub

