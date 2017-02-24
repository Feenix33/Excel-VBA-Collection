Sub cmePrepQView()
' cmePrepQView Macro
    rows("2:2").Select
    Selection.Delete Shift:=xlUp
    cmeTextColumnToNumber ("Labor Hours")
    cmeTextColumnToNumber ("Standard Hours")
    cmeTextColumnToNumber ("Overtime Hours")
    cmeTextColumnToDate ("Weekend")
    
    Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearFormats
    Dim objTable As ListObject
    Set objTable = ActiveSheet.ListObjects.Add(xlSrcRange, Selection, , xlYes)
    objTable.TableStyle = cmeMagicTableStyle
    'Format again because the above clears our the date
    cmeFormatAsDate ("Weekend")
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
    saveBaseName = "QV." + saveBaseName
    saveExt = ".xlsx"
    
    Call GenericVersionSave(saveDir, saveBaseName, saveExt)
End Sub
Sub aaToDate()
'
' aaToDate Macro
'

'
    Columns("O:O").Select
End Sub


