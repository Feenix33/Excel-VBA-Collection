Attribute VB_Name = "QlikView"
Sub cmePrepQView()
Attribute cmePrepQView.VB_ProcData.VB_Invoke_Func = " \n14"
' cmePrepQView Macro
    rows("2:2").Select
    Selection.Delete Shift:=xlUp
    cmeTextColumnToNumber ("Approved Overtime Hours")
    cmeTextColumnToNumber ("Approved Labor Hours")
    cmeTextColumnToNumber ("Approved Straight Hours")
    
    cmeTextColumnToNumber ("Labor Hours")
    cmeTextColumnToNumber ("Standard Hours")
    cmeTextColumnToNumber ("Overtime Hours")
    
    cmeTextColumnToDate ("Weekend")
    cmeTextColumnToDate ("Approval Date")
    
    ActiveSheet.UsedRange.Select
    Selection.ClearFormats
    
'    cmeTextColumnToMonth ("Month")
    ActiveSheet.UsedRange.Select
    
    Dim objTable As ListObject
    Set objTable = ActiveSheet.ListObjects.Add(xlSrcRange, Selection, , xlYes)
    objTable.TableStyle = cmeMagicTableStyle
    'Format again because the above clears our the date
    cmeFormatAsDate ("Weekend")
    cmeFormatAsDate ("Approval Date")
    cmeReFormatDate ("Month")
    Range("A1").Select
End Sub

Sub cmeTextColumnToNumber(strColName As String)
    On Error GoTo DoesNotExist
    ActiveWorkbook.ActiveSheet.rows(1).Find(strColName).Select
    ActiveCell.EntireColumn.Select
    'Destination:=Range("P1"),
    Selection.TextToColumns DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True
    Selection.NumberFormat = "General"
DoesNotExist:
End Sub
Sub cmeTextColumnToDate(strColName As String)
    On Error GoTo DoesNotExist
    ActiveWorkbook.ActiveSheet.rows(1).Find(strColName).Select
    ActiveCell.EntireColumn.Select
    Selection.TextToColumns DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 3), TrailingMinusNumbers:=True
    Selection.NumberFormat = "mm/dd/yy;@"
DoesNotExist:
End Sub
Sub cmeTextColumnToMonth(strColName As String)
    On Error GoTo DoesNotExist
    ActiveWorkbook.ActiveSheet.rows(1).Find(strColName).Select
    ActiveCell.EntireColumn.Select
    Selection.TextToColumns DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 3), TrailingMinusNumbers:=True
    Selection.NumberFormat = "mmm"
DoesNotExist:
End Sub
Sub cmeReFormatDate(strColName As String)
    On Error GoTo DoesNotExist
    ActiveWorkbook.ActiveSheet.rows(1).Find(strColName).Select
    ActiveCell.EntireColumn.Select
    Selection.NumberFormat = "mmm-yy"
DoesNotExist:
End Sub
Sub cmeFormatAsDate(strColName As String)
    On Error GoTo DoesNotExist
    ActiveWorkbook.ActiveSheet.rows(1).Find(strColName).Select
    ActiveCell.EntireColumn.Select
    Selection.NumberFormat = "mm/dd/yy;@"
DoesNotExist:
End Sub
Sub SaveQlikviewReport()
' Saves file to a specific filename in a specific directory
    Dim saveDir As String, saveBaseName As String
    Dim savePath As String, saveExt As String
    Dim strCurrentName As String
    
    saveDir = "C:\Users\sg0213341\Documents\QlikView Reports\"
    strCurrentName = StripDate(ActiveWorkbook.Name)
    saveBaseName = InputBox("Base Filename", "File Plus Date", strCurrentName)
    If Left(saveBaseName, 3) <> "QV." Then
        saveBaseName = "QV." + saveBaseName
    End If
    saveExt = ".xlsx"
    
    Call GenericVersionSave(saveDir, saveBaseName, saveExt)
End Sub
