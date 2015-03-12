'botones

Sub botonPrepararTabla()
    Application.ScreenUpdating = False
    Call copiarTabla
    Call limpiarTabla
    Call ordenarTabla2
    Call completarDatos
    If MsgBox("¿Desea modificar la distribución?", vbYesNo, "Antes de continuar") = vbYes Then
        Sheets("PreDist").Select
    Else
        Call botonGenerarDistribucion
    End If
    Application.ScreenUpdating = True
End Sub

Sub botonGenerarDistribucion()
    Application.ScreenUpdating = False
    Call ordenarTabla
    Call limpiarDistrib
    Call llenarDistrib
    Call FormatoDistrib
    If MsgBox("¿Desea imprimir la distribución?", vbYesNo, "Antes de continuar") = vbNo Then
        Sheets("Distrib").Select
    Else
        Call botonImprimirDistribucion
    End If
    Application.ScreenUpdating = True
    Call guardarDocumento
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
    Sheets("Distrib").Select
    Range("H4").Select
    Fila = 0
    While ActiveCell.Offset(Fila, -6).Value <> ""
        If ActiveCell.Offset(Fila, -6) <> ActiveCell.Offset(Fila - 1, -6) Then
            ActiveCell.Offset(Fila, 0).FormulaR1C1 = 1
        End If
        Fila = Fila + 1
    Wend
    MsgBox ("Agregue un 1 al inicio de cada bulto no identificado automaticamente.")
    Application.ScreenUpdating = True
End Sub

Sub botonGenerarASN()
    Application.ScreenUpdating = False
    If MsgBox("¿Esta facturada la distribución?", vbYesNo, "Antes de continuar...") = vbNo Then
        MsgBox ("Debe facturar antes de continuar.")
    Else
        Call copiarASN
        Call leerFolioASN
        Call llenarLPNASN
        Call completarFolioASN
        Call obtenerFechasASN
        Call generarFoliosASN
        Call nuevoASN
        Sheets("Menu").Select
        MsgBox ("Listo para cargar el archivo en el sistema de HITES.")
    End If
    Application.ScreenUpdating = True
End Sub

Sub botonGenerarEtiqueta()
    Application.ScreenUpdating = False
    Sheets("ASN").Select
    Columns("A:J").Copy
    Sheets("Rotulo").Select
    Range("A1").PasteSpecial Paste:=xlValues
    Range("K1").FormulaR1C1 = "DESTINO"
    Range("L1").FormulaR1C1 = "NVENTA"
    Range("M1").FormulaR1C1 = "CAJA"
    Range("N1").FormulaR1C1 = "CAJAS"
    Columns(10).Delete
    Columns("F:H").Delete
    Sheets("Distrib").Select
    Range("B4", Range("B4").End(xlDown)).Copy
    Sheets("Rotulo").Select
    Range("G2").PasteSpecial Paste:=xlValues
    Range("H2").FormulaR1C1 = "=Distrib!R1C7"
    Columns(1).Delete
    Columns(2).Delete
    Columns("A:E").RemoveDuplicates Columns:=3, Header:=xlYes
    Range("G2").FormulaR1C1 = "=IF(RC[-5]<>R[-1]C[-5],1,R[-1]C+1)"
    Range("H2").FormulaR1C1 = "=COUNTIF(C[-6],RC[-6])"
    If Range("A2").Value <> "" Then
        Range("F2:H2").AutoFill Destination:=Range("F2", Range("E1").End(xlDown).Offset(0, 3))
    End If
    Kill ThisWorkbook.Path & "\bHites\EtiquetaHites.xls"
    Columns("F:H").Copy
    Range("F1").PasteSpecial Paste:=xlValues
    Columns("A:H").AutoFit
    Columns("A:H").Copy
    Workbooks.Add
    Range("A1").PasteSpecial Paste:=xlValues
    ActiveWorkbook.SaveAs Filename:=ThisWorkbook.Path & "\bHites\EtiquetaHites.xls", FileFormat:=xlExcel8
    ActiveWindow.Close
    Sheets("Menu").Select
    MsgBox ("Listo para imprimir etiquetas.")
    Application.ScreenUpdating = True
End Sub

'procesos

Sub copiarTabla()
    Sheets("PreDist").Columns("A:AA").ClearContents
    Sheets("bHites").Select
    Columns("A:AA").Copy
    Sheets("PreDist").Select
    Range("A1").PasteSpecial Paste:=xlPasteValues
End Sub

Sub limpiarTabla()
    Sheets("PreDist").Select
    Sheets("ASN").Range("A2", Sheets("ASN").Range("J1").End(xlDown)).ClearContents
    Sheets("ASN").Range("C2").FormulaR1C1 = Range("D2").Value
    Sheets("ASN").Range("A2").FormulaR1C1 = Range("B2").Value
    Sheets("ASN").Range("B2").FormulaR1C1 = Range("A2").Value
    Columns("O:Q").Delete
    Columns("L:M").Delete
    Columns("J:J").Delete
    Columns("B:E").Delete
    Range("A1").FormulaR1C1 = "OC"
    Range("B1").FormulaR1C1 = "COD"
    Range("C1").FormulaR1C1 = "LOCAL"
    Range("D1").FormulaR1C1 = "SKU"
    Range("E1").FormulaR1C1 = "ATS"
    Range("F1").FormulaR1C1 = "DESCRIP"
    Range("G1").FormulaR1C1 = "CANT"
End Sub

Sub ordenarTabla2()
    Sheets("PreDist").Select
    Sheets("PreDist").Sort.SortFields.Clear
    Sheets("PreDist").Sort.SortFields.Add Key:=Columns(4), Order:=xlAscending
    Sheets("PreDist").Sort.SetRange Columns("A:G")
    Sheets("PreDist").Sort.Header = xlYes
    Sheets("PreDist").Sort.Apply
End Sub

Sub ordenarTabla()
    Sheets("PreDist").Select
    Sheets("PreDist").Sort.SortFields.Clear
    Sheets("PreDist").Sort.SortFields.Add Key:=Columns(2), Order:=xlAscending
    Sheets("PreDist").Sort.SetRange Columns("A:H")
    Sheets("PreDist").Sort.Header = xlYes
    Sheets("PreDist").Sort.Apply
End Sub

Sub completarDatos()
    Sheets("PreDist").Select
    Range("E2").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(VALUE(RC[-1]),Maestras!C[-4]:C[-2],3,0)"
    If Range("A3").Value <> "" Then
        Selection.AutoFill Destination:=Range("E2", Range("D1").End(xlDown).Offset(0, 1))
    End If
    Columns("E:E").Copy
    Range("E1").PasteSpecial Paste:=xlValues
    Columns("D:D").Insert
    Range("D1").FormulaR1C1 = "LIN"
    Range("D2").FormulaR1C1 = "=IF(RC[-1]<>R[-1]C[-1],1,R[-1]C+1)"
    If Range("A3").Value <> "" Then
        Range("D2").AutoFill Destination:=Range("D2", Range("C1").End(xlDown).Offset(0, 1))
    End If
    Columns("D:D").Copy
    Range("D1").PasteSpecial Paste:=xlValues
    Columns("A:H").AutoFit
    Application.CutCopyMode = False
End Sub

' Boton Preparar Distribucion

Sub limpiarDistrib()
    Sheets("Distrib").Range("A4", Sheets("Distrib").Range("G3").End(xlDown)).ClearContents
    Sheets("Distrib").Range("A4", Sheets("Distrib").Range("G3").End(xlDown)).ClearFormats
    Sheets("Distrib").Columns(8).ClearContents
End Sub

Sub llenarDistrib()
    Sheets("Predist").Select
    If Range("H1").End(xlDown).Offset(0, -1) = "" Then
        Range("H1").End(xlDown).EntireRow.Delete
    End If
    Range("B2", Range("H1").End(xlDown)).Copy
    Sheets("Distrib").Select
    Range("A4").PasteSpecial Paste:=xlValues
    Range("G2").FormulaR1C1 = Sheets("Predist").Range("A2").Value
    Range("G1").FormulaR1C1 = InputBox("Ingrese número de Nota de Venta:")
    Range("A4", Range("G3").End(xlDown)).Select
End Sub

Sub FormatoDistrib()
    Sheets("Distrib").Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeLeft).Weight = xlHairline
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).Weight = xlHairline
    Selection.Borders(xlInsideVertical).LineStyle = xlContinuous
    Selection.Borders(xlInsideVertical).Weight = xlHairline
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

'Boton Generar ASN

Sub copiarASN()
    Sheets("Distrib").Select
    Range("A4", Range("A3").End(xlDown)).Copy
    Sheets("ASN").Select
    Range("D2").PasteSpecial Paste:=xlValues
    Sheets("Distrib").Select
    Range("D4", Range("E3").End(xlDown)).Copy
    Sheets("ASN").Select
    Range("F2").PasteSpecial Paste:=xlValues
    Sheets("Distrib").Select
    Range("G4", Range("G3").End(xlDown)).Copy
    Sheets("ASN").Select
    Range("H2").PasteSpecial Paste:=xlValues
    Sheets("Distrib").Select
    Range("H4", Range("G3").End(xlDown).Offset(0, 1)).Copy
    Sheets("ASN").Select
    Range("E2").PasteSpecial Paste:=xlValues
    Range("A2:C2").Select
    Selection.AutoFill Destination:=Range("A2", Range("D1").End(xlDown).Offset(0, -1))
    Range("J2").FormulaR1C1 = "EFAC"
    Range("I2").FormulaR1C1 = InputBox("Ingrese el número de la factura.")
    Range("I2:J2").Copy
    Range("I2", Range("H1").End(xlDown).Offset(0, 2)).PasteSpecial Paste:=xlValues
End Sub

Sub leerFolioASN()
    Open ThisWorkbook.Path & "\bHites\FolioLPN.txt" For Input As #1
        Line Input #1, FolioLPN
        Sheets("Maestras").Range("E2").FormulaR1C1 = FolioLPN
    Close #1
End Sub

Sub llenarLPNASN()
    Sheets("ASN").Select
    Range("E2").Select
    Folio = Sheets("Maestras").Range("E2").Value - 1
    Linea = 0
    While ActiveCell.Offset(Linea, -1).Value <> ""
        If ActiveCell.Offset(Linea, 0).Value = 1 Then
            Folio = Folio + 1
            ActiveCell.Offset(Linea, 0).FormulaR1C1 = Folio
        Else
            ActiveCell.Offset(Linea, 0).FormulaR1C1 = Folio
        End If
        Linea = Linea + 1
    Wend
    Sheets("Maestras").Range("E4").FormulaR1C1 = Sheets("Maestras").Range("E2").Value
    Sheets("Maestras").Range("E2").FormulaR1C1 = Folio + 1
End Sub

Sub completarFolioASN()
    Sheets("ASN").Select
    Range("K2").Select
    ActiveCell.FormulaR1C1 = "=CONCATENATE(""55"",""90914000"",TEXT(RC[-6],""00000000""))"
    If Range("A3").Value <> "" Then
        Selection.AutoFill Destination:=Range("K2", Range("J1").End(xlDown).Offset(0, 1))
    End If
    Range("K1").FormulaR1C1 = "LPN"
    Columns("K:K").Copy
    Range("E1").PasteSpecial Paste:=xlValues
    Columns("K:K").Delete
    Columns("A:J").AutoFit
End Sub

Sub obtenerFechasASN()
    Sheets("Maestras").Range("E8").FormulaR1C1 = "=TODAY()"
    Sheets("Maestras").Range("E8").Copy
    Sheets("Maestras").Range("E8").PasteSpecial Paste:=xlValues
    Open ThisWorkbook.Path & "\bHites\Fecha.txt" For Input As #1
        Line Input #1, AUX
        Sheets("Maestras").Range("E6").FormulaR1C1 = AUX
    Close #1
End Sub

Sub generarFoliosASN()
    Open ThisWorkbook.Path & "\bHites\FolioDiario.txt" For Input As #1
        Line Input #1, AUX
        Sheets("Maestras").Range("E10").FormulaR1C1 = AUX
    Close #1
    If Sheets("Maestras").Range("E8").Value > Sheets("Maestras").Range("E6").Value Then
        Sheets("Maestras").Range("E6").FormulaR1C1 = Sheets("Maestras").Range("E8").Value
        Sheets("Maestras").Range("E10").FormulaR1C1 = 1
        Open ThisWorkbook.Path & "\bHites\Fecha.txt" For Output As #1
            AUX = Sheets("Maestras").Range("E6").Value
            Print #1, AUX
        Close #1
        Open ThisWorkbook.Path & "\bHites\FolioDiario.txt" For Output As #1
            AUX = Sheets("Maestras").Range("E10").Value
            Print #1, AUX
        Close #1
    Else
        Sheets("Maestras").Range("E10").FormulaR1C1 = Sheets("Maestras").Range("E10").Value + 1
        Open ThisWorkbook.Path & "\bHites\Fecha.txt" For Output As #1
            AUX = Sheets("Maestras").Range("E6").Value
            Print #1, AUX
        Close #1
        Open ThisWorkbook.Path & "\bHites\FolioDiario.txt" For Output As #1
            AUX = Sheets("Maestras").Range("E10").Value
            Print #1, AUX
        Close #1
    End If
    Sheets("Maestras").Range("E11").FormulaR1C1 = "=TEXT(R[-1]C,""00"")"
    Sheets("Maestras").Range("E11").Copy
    Sheets("Maestras").Range("E11").PasteSpecial Paste:=xlValues
    Folio = Sheets("Maestras").Range("E2").Value
    Open ThisWorkbook.Path & "\bHites\FolioLPN.txt" For Output As #1
        Print #1, Folio
    Close #1
End Sub

Sub nuevoASN()
    Sheets("Distrib").Columns(8).ClearContents
    Sheets("ASN").Select
    Range("K1").FormulaR1C1 = "=CONCATENATE(""_"",TEXT(TODAY(),""dd""),""_"",TEXT(TODAY(),""mm""),""_"",TEXT(TODAY(),""yyyy""),""_"")"
    Fecha = Range("K1").Value
    Range("K1").ClearContents
    ASNFolio = Sheets("Maestras").Range("E11").Value
    Columns("A:J").Copy
    Workbooks.Add
    Range("A1").PasteSpecial Paste:=xlValues
    Rows("1:1").Delete
    Columns("A:J").AutoFit
    ActiveWorkbook.SaveAs Filename:=ThisWorkbook.Path & "\bHites\ASN_90914000" & Fecha & ASNFolio & ".xls", FileFormat:=xlExcel8
    ActiveWindow.Close
End Sub

Sub guardarDocumento()
    notaVenta = Sheets("Distrib").Range("G1").Value
    ActiveWorkbook.SaveAs Filename:=ThisWorkbook.Path & "\" & notaVenta & ".xlsm"
    If Dir(ThisWorkbook.Path & "\" & notaVenta & ".bat", vbNormal) = "" Then
        Open ThisWorkbook.Path & "\" & notaVenta & ".bat" For Output As #1
            Print #1, "start " & notaVenta & ".xlsm"
        Close #1
    End If
End Sub
