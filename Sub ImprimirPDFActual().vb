Sub ImprimirActualPDF()
'
' ImprimirPDF Macro
'

'Establecer Hoja
        Application.ScreenUpdating = False
        Dim WB_Extract As Worksheet
        Sheets("Reporte").Activate
fila = Application.WorksheetFunction.CountA(Range("A:A")) + 1
NoReqAnterior = Range("H" & fila - 1)
NoReqActual = NoReqAnterior

'NoReqActual = 5
Sheets("Impresion").Activate
Range("F2").Formula = NoReqActual




   ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True, _
        IgnorePrintAreas:=False
End Sub
