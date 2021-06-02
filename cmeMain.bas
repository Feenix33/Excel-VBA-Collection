Attribute VB_Name = "cmeMain"
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
Sub cmePasteValues()
Attribute cmePasteValues.VB_ProcData.VB_Invoke_Func = "V\n14"
    On Error GoTo Fini
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
Fini:
End Sub
Sub cmeFilterToggle()
Attribute cmeFilterToggle.VB_ProcData.VB_Invoke_Func = "O\n14"
' Keyboard Shortcut: Ctrl+Shift+O
    On Error GoTo errorFilterToggle
    Selection.AutoFilter
errorFilterToggle:
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
'    cmeAutoLimitProcess (60) ' resize the sheet
'    If IsNumeric(Application.Match(LCase(Left(ActiveSheet.Name, 5)), arrMagicChange, 0)) Then
'        ActiveSheet.Name = ProcessSheetName(">d") ' name the sheet for today
'    End If
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
Function xxStripRev(inName As String) As String
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
        fDialogue.InitialFileName = """C:\Users\christophere\OneDrive - Magenic\"""
    End If
    If fDialogue.Show = -1 Then
       GetMyDirectory = fDialogue.SelectedItems(1)
    Else
       GetMyDirectory = ""
    End If
End Function
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
'Sub UpdateRev()
'    Dim saveDir As String
'    Dim saveBaseName As String
'    Dim saveExt As String
'
'    On Error GoTo Fini
'    saveDir = ActiveWorkbook.Path
'    saveBaseName = StripRev(ActiveWorkbook.Name)
'    saveExt = ".xlsx"
'    Call RevisionVersionSave(saveDir + "\" + saveBaseName, saveExt)
'Fini:
'End Sub
