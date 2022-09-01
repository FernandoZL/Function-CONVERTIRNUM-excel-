Sub ImprimirPDF()
'
' ImprimirPDF Macro
'

'Establecer Hoja
        Application.ScreenUpdating = False
        Dim WB_Extract As Worksheet
        Sheets("Impresion").Activate

    Range("F2").Select
    ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True, _
        IgnorePrintAreas:=False
        
        
End Sub
