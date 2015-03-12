'botones

Sub botonProcesarTabla()
    Application.ScreenUpdating = False
    Call copiarLote
    Call ordenarLote
    If MsgBox("Desea modificar las cantidades", vbYesNo) = vbNo Then
        Call botonGenerarDistribución
    Else
        Sheets("Lote").Select
        Range("I1").Select
    End If
    Application.ScreenUpdating = True
End Sub

Sub botonGenerarDistribución()
    Application.ScreenUpdating = False
    Call ordenarPorLocal
    Call limpiarDistribucion
    Call llenarDistribucion
    Call agregarATS
    Call agregarLineas
    Call formatoDistribucion
    Call guardarDocumento
    If MsgBox("Desea imprimir la distribución?", vbYesNo) = vbYes Then
        Call botonImprimirDistribucion
    Else
        Sheets("Menu").Select
    End If
    Application.ScreenUpdating = True
End Sub

Sub botonImprimirDistribucion()
    Application.ScreenUpdating = False
    Sheets("Distrib").Select
    ActiveWindow.SelectedSheets.PrintOut
    Sheets("Menu").Select
    Application.ScreenUpdating = True
End Sub

Sub botonSeleccionarBultos()
    Application.ScreenUpdating = False
    If MsgBox("¿Pegó el packing list?", vbYesNo) = vbYes Then
        Call agregarIDBulto
        Call identificarBultos
        Application.CutCopyMode = False
        MsgBox ("Agregue un 1 al inicio de cada bulto que no se haya identificado.")
        Sheets("Lote").Select
        Range("N2").Select
    Else
        MsgBox ("Primero debe bajar el packing list desde el B2B.")
    End If
    Application.ScreenUpdating = True
End Sub

Sub botonGenerarPacking()
    Application.ScreenUpdating = False
    If MsgBox("Está facturada?", vbYesNo) = vbYes Then
        Call extraeNumeroBulto
        Call generarNumeroNuevoBulto
        Call llenarFactura
        Sheets("Lote").Select
        Range("A2", Range("M2").End(xlDown)).Copy
        MsgBox ("Pegue esta planilla en el B2B de Paris.")
    Else
        MsgBox ("Primero debe facturar el pedido.")
    End If
End Sub

Sub botonGenerarRotulos()
    Application.ScreenUpdating = False
    Call copiarPackingListRotulo
    Call ordenarRotulos
    Call llenarNotaDeVentaRotulo
    Call llenarNumeroCajasRotulo
    Call llenarCajasPorLocalRotulo
    Call totalBultosRotulo
    Call llenarCitaRotulo
    Call generarArchivoRotulo
    Columns("A:K").AutoFit
    Sheets("menu").Select
    MsgBox ("Listo para imprimir.")
    Application.ScreenUpdating = True
End Sub

'procesos

Sub copiarPackingListRotulo()
    Sheets("Lote").Select
    Columns("A:M").Copy
    Sheets("Rotulos").Select
    Range("A1").PasteSpecial Paste:=xlValues
End Sub

Sub ordenarRotulos()
    Sheets("Rotulos").Select
    Columns("A:A").Delete
    Columns("B:D").Delete
    Columns("C:C").Delete
    Columns("D:D").Delete
    Columns("F:F").Delete
    Range("A1").FormulaR1C1 = "BULTO"
    Range("B1").FormulaR1C1 = "DEPTO"
    Range("C1").FormulaR1C1 = "LOCAL"
    Range("D1").FormulaR1C1 = "ATSPV"
    Range("E1").FormulaR1C1 = "ORDEN"
    Range("F1").FormulaR1C1 = "FACTR"
    Range("G1").FormulaR1C1 = "NVENT"
    Range("H1").FormulaR1C1 = "NCAJA"
    Range("I1").FormulaR1C1 = "CAJAS"
    Range("J1").FormulaR1C1 = "NOBLT"
    Range("K1").FormulaR1C1 = "NCITA"
    Columns("A:J").Select
    ActiveSheet.Range("A1", Range("M1").End(xlDown)).RemoveDuplicates Columns:=1, Header:=xlYes
End Sub

Sub llenarNotaDeVentaRotulo()
    Sheets("Rotulos").Select
    Range("G2").FormulaR1C1 = Sheets("Distrib").Range("G1").Value
    If Range("A3").Value <> "" Then
        Range("G2").AutoFill Destination:=Range("G2", Range("F2").End(xlDown).Offset(0, 1))
    End If
End Sub

Sub llenarNumeroCajasRotulo()
    Sheets("Rotulos").Select
    Range("H2").FormulaR1C1 = "=IF(RC[-5]<>R[-1]C[-5],1,R[-1]C+1)"
    If Range("A3").Value <> "" Then
        Range("H2").AutoFill Destination:=Range("H2", Range("G2").End(xlDown).Offset(0, 1))
        Range("H2", Range("G2").End(xlDown).Offset(0, 1)).Copy
        Range("H2").PasteSpecial Paste:=xlPasteValues
    Else
        Range("H2").Copy
        Range("H2").PasteSpecial Paste:=xlPasteValues
    End If
End Sub

Sub llenarCajasPorLocalRotulo()
    Sheets("Rotulos").Select
    Range("I2").FormulaR1C1 = "=COUNTIF(C[-6],RC[-6])"
    If Range("A3").Value <> "" Then
        Range("I2").AutoFill Destination:=Range("I2", Range("H1").End(xlDown).Offset(0, 1))
        Range("I2", Range("H1").End(xlDown).Offset(0, 1)).Copy
        Range("I2").PasteSpecial Paste:=xlValues
    Else
        Range("I2").Copy
        Range("I2").PasteSpecial Paste:=xlPasteValues
    End If
End Sub

Sub totalBultosRotulo()
    Sheets("Rotulos").Select
    Range("J2").FormulaR1C1 = "=COUNTA(C[-9])-1"
    If Range("A3").Value <> "" Then
        Range("J2").AutoFill Destination:=Range("J2", Range("I1").End(xlDown).Offset(0, 1))
        Range("J2", Range("I1").End(xlDown).Offset(0, 1)).Copy
        Range("J2").PasteSpecial Paste:=xlValues
    Else
        Range("J2").Copy
        Range("J2").PasteSpecial Paste:=xlPasteValues
    End If
End Sub

Sub llenarCitaRotulo()
    Range("K2").FormulaR1C1 = InputBox("Ingrese el número de cita.")
    If Range("A3").Value <> "" Then
        Range("K2").AutoFill Destination:=Range("K2", Range("J1").End(xlDown).Offset(0, 1))
    End If
End Sub

Sub generarArchivoRotulo()
    If Dir(ThisWorkbook.Path & "\bJohnsons\eJohnsons.xls", vbNormal) <> "" Then
        Kill ThisWorkbook.Path & "\bJohnsons\eJohnsons.xls"
    End If
    Columns("A:K").Copy
    Workbooks.Add
    Range("A1").PasteSpecial Paste:=xlPasteValues
    ActiveWorkbook.SaveAs Filename:=ThisWorkbook.Path & "\bJohnsons\eJohnsons.xls", FileFormat:=xlExcel8
    ActiveWindow.Close
End Sub

Sub copiarLote()
    Sheets("Lote").Select
    Columns("A:AA").ClearContents
    Sheets("bJohnsons").Select
    Columns("A:M").Copy
    Sheets("Lote").Select
    Range("A1").PasteSpecial Paste:=xlPasteValues
    Rows(1).Delete
    Columns(13).Delete
    Columns(12).Delete
    Columns(9).Delete
    Columns("A:J").AutoFit
End Sub

Sub ordenarLote()
    Sheets("Lote").Select
    Columns("H:I").Cut
    Columns(1).Insert
    Columns(3).Cut
    Columns(10).Insert
    Columns("E:F").Cut
    Columns(4).Insert
    Columns(2).Insert
    Columns(2).Insert
    Range("A1").FormulaR1C1 = "TIPO FLUJO"
    Range("B1").FormulaR1C1 = "ID BULTO"
    Range("C1").FormulaR1C1 = "ID PALLET"
    Range("D1").FormulaR1C1 = "DESC. PRODUCTO"
    Range("E1").FormulaR1C1 = "COD. LOCAL"
    Range("F1").FormulaR1C1 = "COD. DEPTO"
    Range("G1").FormulaR1C1 = "DEPT."
    Range("H1").FormulaR1C1 = "DESC. LOCAL"
    Range("I1").FormulaR1C1 = "COD. JOHNSONS"
    Range("J1").FormulaR1C1 = "COD. PROV."
    Range("K1").FormulaR1C1 = "COD. ORDEN"
    Range("L1").FormulaR1C1 = "CANTIDAD"
    Range("M1").FormulaR1C1 = "DOCTO LEGAL"
End Sub

Sub ordenarPorLocal()
    Sheets("Lote").Select
    With ActiveWorkbook.Worksheets("Lote").Sort
        .SortFields.Clear
        .SortFields.Add Key:=Columns(8), Order:=xlAscending
        .SetRange Columns("A:M")
        .Header = xlYes
        .Apply
    End With
End Sub

Sub limpiarDistribucion()
    Sheets("Distrib").Select
    Range("A4", Range("G3").End(xlDown)).ClearContents
    Range("A4", Range("G3").End(xlDown)).ClearFormats
End Sub

Sub llenarDistribucion()
    Sheets("Lote").Select
    Range("E2", Range("E1").End(xlDown)).Copy
    Sheets("Distrib").Select
    Range("A4").PasteSpecial Paste:=xlPasteValues
    Sheets("Lote").Select
    Range("H2", Range("H1").End(xlDown)).Copy
    Sheets("Distrib").Select
    Range("B4").PasteSpecial Paste:=xlPasteValues
    Sheets("Lote").Select
    Range("I2", Range("J1").End(xlDown)).Copy
    Sheets("Distrib").Select
    Range("D4").PasteSpecial Paste:=xlPasteValues
    Sheets("Lote").Select
    Range("D2", Range("D1").End(xlDown)).Copy
    Sheets("Distrib").Select
    Range("F4").PasteSpecial Paste:=xlPasteValues
    Sheets("Lote").Select
    Range("L2", Range("L1").End(xlDown)).Copy
    Sheets("Distrib").Select
    Range("G4").PasteSpecial Paste:=xlPasteValues
    Range("G2").FormulaR1C1 = Sheets("Lote").Range("K2").Value
    Range("G1").FormulaR1C1 = InputBox("Ingrese la nota de venta: ")
End Sub

Sub agregarATS()
    Sheets("Distrib").Select
    Range("E4").FormulaR1C1 = "=IFERROR(VLOOKUP(VALUE(RC[-1]),Maestras!C[-4]:C[-3],2,0),1)"
    If Range("A5") <> "" Then
        Range("E4").AutoFill Destination:=Range("E4", Range("D3").End(xlDown).Offset(0, 1))
        Range("E4", Range("D3").End(xlDown).Offset(0, 1)).Copy
        Range("E4").PasteSpecial Paste:=xlPasteValues
    Else
        Range("E4").Copy
        Range("E4").PasteSpecial Paste:=xlPasteValues
    End If
End Sub

Sub agregarLineas()
    Sheets("Distrib").Select
    Range("C4").FormulaR1C1 = "=IF(RC[-1]<>R[-1]C[-1],1,R[-1]C+1)"
    If Range("A5") <> "" Then
        Range("C4").AutoFill Destination:=Range("C4", Range("D3").End(xlDown).Offset(0, -1))
        Range("C4", Range("D3").End(xlDown).Offset(0, -1)).Copy
        Range("C4").PasteSpecial Paste:=xlPasteValues
    Else
        Range("C4").Copy
        Range("C4").PasteSpecial Paste:=xlPasteValues
    End If
End Sub

Sub formatoDistribucion()
    Sheets("Distrib").Select
    Range("A4", Range("G3").End(xlDown)).Select
    Selection.Borders(xlInsideVertical).LineStyle = xlContinuous
    Selection.Borders(xlInsideVertical).Weight = xlHairline
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeLeft).Weight = xlHairline
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).Weight = xlHairline
    Range("A4").Select
    While ActiveCell.Value <> ""
        If ActiveCell.Value = ActiveCell.Offset(1, 0) Then
            Range(ActiveCell, ActiveCell.End(xlToRight)).Borders(xlEdgeBottom).LineStyle = xlContinuous
            Range(ActiveCell, ActiveCell.End(xlToRight)).Borders(xlEdgeBottom).Weight = xlHairline
        Else
            Range(ActiveCell, ActiveCell.End(xlToRight)).Borders(xlEdgeBottom).LineStyle = xlContinuous
            Range(ActiveCell, ActiveCell.End(xlToRight)).Borders(xlEdgeBottom).Weight = xlThick
        End If
        ActiveCell.Offset(1, 0).Select
    Wend
End Sub

Sub guardarDocumento()
    notaVenta = Sheets("Distrib").Range("G1").Value
    ActiveWorkbook.SaveAs Filename:=ThisWorkbook.Path & "\" & notaVenta & ".xlsm"
    If Dir(ThisWorkbook.Path & "\" & notaVenta & ".bat", vbNormal) <> "" Then
        Kill ThisWorkbook.Path & "\" & notaVenta & ".bat"
    End If
    Open ThisWorkbook.Path & "\" & notaVenta & ".bat" For Output As #1
        Print #1, "start " & notaVenta & ".xlsm"
    Close #1
End Sub

Sub agregarIDBulto()
    Sheets("PList").Select
    Columns(8).Cut
    Columns(2).Insert
    Sheets("Lote").Select
    Range("B2").FormulaR1C1 = "=VLOOKUP(RC[6],PList!C:C[1],2,0)"
    If Range("A3") <> "" Then
        Range("B2").AutoFill Destination:=Range("B2", Range("A1").End(xlDown).Offset(0, 1))
        Range("B2", Range("A1").End(xlDown).Offset(0, 1)).Copy
        Range("B2").PasteSpecial Paste:=xlPasteValues
    Else
        Range("B2").Copy
        Range("B2").PasteSpecial Paste:=xlPasteValues
    End If
    Sheets("PList").Select
    Columns(2).Cut
    Columns(9).Insert
End Sub

Sub identificarBultos()
    Sheets("Lote").Select
    Range("N2").FormulaR1C1 = "=IF(RC[-6]<>R[-1]C[-6],1,"""")"
    If Range("A3") <> "" Then
        Range("N2").AutoFill Destination:=Range("N2", Range("L1").End(xlDown).Offset(0, 2))
        Range("N2", Range("L1").End(xlDown).Offset(0, 2)).Copy
        Range("N2").PasteSpecial Paste:=xlPasteValues
    Else
        Range("N2").Copy
        Range("N2").PasteSpecial Paste:=xlPasteValues
    End If
End Sub

Sub extraeNumeroBulto()
    Sheets("Lote").Select
    Range("M2").FormulaR1C1 = "=VALUE(RIGHT(RC[-11],5))"
    If Range("A3").Value <> "" Then
        Range("M2").AutoFill Destination:=Range("M2", Range("L1").End(xlDown).Offset(0, 1))
        Range("M2", Range("L1").End(xlDown).Offset(0, 1)).Copy
        Range("M2").PasteSpecial Paste:=xlPasteValues
    Else
        Range("M2").Copy
        Range("M2").PasteSpecial Paste:=xlPasteValues
    End If
    Range("M2").Select
    While ActiveCell.Value <> ""
        If ActiveCell.Offset(0, -5).Value = ActiveCell.Offset(-1, -5).Value Then
            If ActiveCell.Offset(0, 1).Value = 1 Then
                ActiveCell.FormulaR1C1 = ActiveCell.Offset(-1, 0).Value + 1
            End If
        End If
        ActiveCell.Offset(1, 0).Select
    Wend
End Sub

Sub generarNumeroNuevoBulto()
    Sheets("Lote").Select
    Range("O2").FormulaR1C1 = "=CONCATENATE(LEFT(RC[-13],6),TEXT(RC[-2],""00000""))"
    If Range("A3").Value <> "" Then
        Range("O2").AutoFill Destination:=Range("O2", Range("M1").End(xlDown).Offset(0, 2))
        Range("O2", Range("M1").End(xlDown).Offset(0, 2)).Copy
        Range("O2").PasteSpecial Paste:=xlPasteValues
    Else
        Range("O2").Copy
        Range("O2").PasteSpecial Paste:=xlPasteValues
    End If
    Range("O1").FormulaR1C1 = "ID BULTO"
    Columns(15).Copy
    Range("B1").PasteSpecial Paste:=xlPasteValues
    Columns(15).ClearContents
End Sub

Sub llenarFactura()
    Sheets("Lote").Select
    Range("M2").FormulaR1C1 = InputBox("Ingrese el número de la factura: ")
    Range("M2").Copy
    Range("M2", Range("M1").End(xlDown)).PasteSpecial Paste:=xlPasteValues
End Sub

'especial

Sub separarEspecial()
    Range("M2").Select
    While ActiveCell.Value <> ""
        i = 1
        While i < ActiveCell.Value
            ActiveCell.EntireRow.Copy
            ActiveCell.EntireRow.Insert
            i = i + 1
            ActiveCell.Offset(1, 0).Select
        Wend
        ActiveCell.Offset(1, 0).Select
    Wend
End Sub
