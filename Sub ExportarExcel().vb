Sub ExportarExcel()
'
' Macro Exportar Excel
'

'
    Range("C8").Select
    Selection.End(xlToLeft).Select
    Selection.End(xlToLeft).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Range("Tabla1[#All]").Select
    Selection.Copy
    Workbooks.Add
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Cells.Select
    Cells.EntireColumn.AutoFit
    Range("H2").Select
    Columns("H:H").ColumnWidth = 12.71
    Range("A1").Select
    Application.CutCopyMode = False
    Windows("Requisiciones Timbres v2.xlsm").Activate
End Sub