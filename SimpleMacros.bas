Attribute VB_Name = "SimpleMacros"
Option Explicit
Sub cmePasteValues()
Attribute cmePasteValues.VB_ProcData.VB_Invoke_Func = "V\n14"
' Keyboard Shortcut: Ctrl+Shift+V
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
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
    
    pos = InStr(LCase(newSheetName), ">gt")
    If pos > 0 Then
       newSheetName = Left(newSheetName, pos - 1) + "GetThere" + Right(newSheetName, Len(newSheetName) - pos - 2)
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
    
    pos = InStr(LCase(newSheetName), ">l")
    If pos > 0 Then
       newSheetName = Left(newSheetName, pos - 1) + "Labor" + Right(newSheetName, Len(newSheetName) - pos - 1)
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
Sub UsedRange_Example_Column()
    Dim LastColumn As Long
    With ActiveSheet.UsedRange
        LastColumn = .Columns(.Columns.Count).Column
    End With
    MsgBox LastColumn
End Sub
Sub cmeAutoLimit()
Attribute cmeAutoLimit.VB_ProcData.VB_Invoke_Func = "L\n14"
' Auto size the cells, but then loop through and limit the width
' only makes stuff smaller, doesn't make them bigger
    Dim myWidth As Integer
    myWidth = 80
    myWidth = Int(InputBox("Maximum width", "User Input", 60))
    cmeAutoLimitProcess (myWidth)
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
    
    Dim LastColumn As Long
    With ActiveSheet.UsedRange
        LastColumn = .Columns(.Columns.Count).Column
    End With

    Dim c As Integer
    For c = 1 To LastColumn
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
  For Each cell In Selection.rows(1).Cells
    On Error Resume Next
      Set pf = cell.PivotField
    On Error GoTo 0
    
    If Not pf Is Nothing Then
      'Toggle between Counting and Summing
        pf.Function = xlCount + xlSum - pf.Function
      
      'Format Numbers with Custom Rule
        pf.NumberFormat = "#,##0_);(#,##0);-"
      
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
        
    If Not Selection.ListObject Is Nothing And Selection.rows.Count = 1 Then
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
                .AutoFilter field:=filterFields(i), Criteria1:=filterCriteria(i)
            Next i
        End With
        Set tableObj = Nothing
    End If
End Sub
Sub cmeCalendar()
    Dim rows As Integer
    Dim cols As Integer
    Dim aDay As Date
    Dim c
    
    
    rows = Selection.rows.Count
    cols = Selection.Columns.Count
    Selection.Resize(rows, 7).Select
    aDay = Date - Weekday(Date) + 1
    ActiveCell.Value = aDay
    For Each c In Selection.Cells
        c.Value = aDay
        aDay = aDay + 1
    Next
End Sub
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
Sub GenericVersionSave(saveDir As String, saveBaseName As String, saveExt As String)
    Dim saveVer As String
    Dim bSaved As Boolean
    Dim savePath As String
    
    If Right(saveDir, 1) <> "\" Then saveDir = saveDir + "\"
    If Right(saveBaseName, 1) <> "." Then saveBaseName = saveBaseName + "."
    bSaved = False
    saveVer = ""
    Do While Not bSaved
        savePath = saveDir + saveBaseName + Format(Date, "yyyy.mm.dd") + saveVer + saveExt
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
Sub VersionSave(savePathFile As String)
    Dim saveVer As String
    Dim bSaved As Boolean
    Dim savePath As String
    Dim saveExt As String
    
    saveExt = ".xlsx"
    bSaved = False
    saveVer = ""
    Do While Not bSaved
        savePath = savePathFile + Format(Date, "yyyy.mm.dd") + saveVer + saveExt
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
Sub SaveBusinessObjectsReport()
' Saves file to a specific filename in a specific directory
    Dim saveDir As String, saveBaseName As String
    Dim savePath As String, saveExt As String
    Dim reportName As String
    
    reportName = ActiveWorkbook.Name
    If Left(reportName, Len("BusObj.")) = "BusObj." Then
        reportName = Mid(reportName, Len("BusObj.") + 1, 255)
        If InStr(reportName, ".") > 0 Then
            reportName = Left(reportName, InStr(reportName, ".") - 1)
        End If
    ElseIf InStr(reportName, ".") > 0 Then 'use the part before the dot
    'guess the first part and remove the date
        reportName = Left(reportName, InStr(reportName, ".") - 1)
    Else
        reportName = "BObj"
    End If
    If InStr(reportName, " ") > 0 Then ' check if space in name
        reportName = Left(reportName, InStr(reportName, " ") - 1)
    End If
    saveBaseName = InputBox("Which Report?", "Business Object Save", reportName)
    If Len(saveBaseName) = 0 Then
        Exit Sub
    End If
    
    saveDir = "C:\Users\sg0213341\Documents\BusObjReports\"
    saveBaseName = "BusObj." + saveBaseName
    saveExt = ".xlsx"
    
    Call GenericVersionSave(saveDir, saveBaseName, saveExt)
End Sub
Sub SaveJiraReport()
' Saves file to a specific filename in a specific directory
    Dim saveDir As String, saveBaseName As String
    Dim savePath As String, saveExt As String
        
    saveDir = "C:\Users\sg0213341\Documents\Exports\"
    saveBaseName = "Jira.Export."
    saveExt = ".xlsx"
    
    Call GenericVersionSave(saveDir, saveBaseName, saveExt)
End Sub
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
    End If
    If fDialogue.Show = -1 Then
       GetMyDirectory = fDialogue.SelectedItems(1)
    Else
       GetMyDirectory = ""
    End If
End Function
Sub cmeBObjLaborReportPreparation()
    Dim tbl As ListObject
    Dim rng As Range
    Dim iStyle As Integer
    
    Sheets("Labor Details").Select
    Range("B2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Copy
    Sheets.Add After:=ActiveSheet
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("A1").Select
    ActiveWorkbook.ActiveSheet.rows(1).Find("Week Ending").Select
    ActiveCell.EntireColumn.Select
    Selection.NumberFormat = "dd-mmm-yy"
    
    ActiveSheet.Name = ProcessSheetName(">L >D")
    Set rng = Range(Range("A1"), Range("A1").SpecialCells(xlLastCell))
    Set tbl = ActiveSheet.ListObjects.Add(xlSrcRange, rng, , xlYes)
    'tbl.TableStyle = "TableStyleMedium" & Weekday(Date)
'    iStyle = Day(Date) '28 styles available
'    If iStyle > 28 Then iStyle = iStyle - 28
'    tbl.TableStyle = "TableStyleMedium" & Day(Date)
    tbl.TableStyle = cmeMagicTableStyle
    cmeAutoLimitProcess (60)
    Range("A1").Select
End Sub
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
Sub RallyReportPreparation()
    Dim tbl As ListObject
    Dim rng As Range
    
    Set rng = Range(Range("A1"), Range("A1").SpecialCells(xlLastCell))
    Set tbl = ActiveSheet.ListObjects.Add(xlSrcRange, rng, , xlYes)

    ' cme magic table formatter
    tbl.TableStyle = cmeMagicTableStyle
    
    Range("A1").Select
    cmeAutoLimitProcess (60) ' resize the sheet
    ActiveSheet.Name = ProcessSheetName(">D") ' name the sheet for today
End Sub
Sub cmeWIPTable()
    Dim objtable As PivotTable
    Dim objfield As PivotField
    
    ActiveSheet.Range("b2").Select
    Set objtable = ActiveSheet.PivotTableWizard
    
    Set objfield = objtable.PivotFields("Schedule State")
    objfield.Orientation = xlColumnField
    
    Set objfield = objtable.PivotFields("Project")
    objfield.Orientation = xlRowField
    
    Set objfield = objtable.PivotFields("Plan Estimate")
    objfield.Orientation = xlDataField
    objfield.Function = xlSum
    
    Set objfield = objtable.PivotFields("Plan Estimate")
    objfield.Orientation = xlPageField
    objfield.PivotItems("0").Visible = False
    objfield.PivotItems("(blank)").Visible = False
    objtable.TableStyle2 = cmeMagicPivotStyle ' "PivotStyleMedium" & Weekday(Date)
    
    ActiveSheet.Name = ProcessSheetName(">P WIP") ' name the sheet Pivot WIP
End Sub

