'botones

Sub botonProcesarTabla()
    Application.ScreenUpdating = False
    On Error Resume Next
    Call copiarEOD
    Call separarColumnas
    Call limpiarEOD
    Call llenarDatos
    If MsgBox("¿Desea modificar las cantidades?", vbYesNo, "Antes de continuar...") = vbYes Then
        Sheets("eOD").Select
        Range("A1").Select
        Columns("A:K").AutoFit
        Application.CutCopyMode = False
    Else
        Call botonGenerarDistribucion
    End If
    Application.ScreenUpdating = True
End Sub

Sub botonGenerarDistribucion()
    Application.ScreenUpdating = False
    On Error Resume Next
    Call ordenarPorLocal
    Call seleccionarBultos
    Call numeroBulto
    Call limpiarDistribucion
    Call llenarTipoYNotaVenta
    Call completarEOD
    Call copiarDistribucion
    Call formatoDistribucion
    Call guardarDocumento
    Sheets("Menu").Select
    If MsgBox("¿Desea imprimir la distribución?", vbYesNo, "Antes de continuar...") = vbYes Then
        Call botonImprimirDistribucion
    Else
        MsgBox ("Proceso finalizado.")
    End If
    Application.ScreenUpdating = True
End Sub

Sub botonImprimirDistribucion()
    Application.ScreenUpdating = False
    On Error Resume Next
    Sheets("Distrib").Select
    Application.PrintCommunication = True
    ActiveWindow.SelectedSheets.PrintOut
    Sheets("Menu").Select
    MsgBox ("Enviado a impresora.")
    Application.ScreenUpdating = True
End Sub

Sub botonGenerarRotulo()
    Application.ScreenUpdating = False
    On Error Resume Next
    Call llenarDatosDeRotulo
    Call quitarRotulosDuplicados
    Call crearNuevoDocumento
    Application.ScreenUpdating = True
End Sub

Sub botonCrearEPIR()
    Application.ScreenUpdating = False
    On Error Resume Next
    Call limpiarEPIR
    Call llenarEPIR
    Call crearArchivoEPIR
    Sheets("Distrib").Columns(8).ClearContents
    Sheets("Menu").Select
    MsgBox ("Archivo creado correctamente.")
    Application.ScreenUpdating = True
End Sub

'procesos

Sub copiarEOD()
    Sheets("eOD").Select
    Columns("A:AA").ClearContents
    Sheets("bTottus").Select
    Columns(1).Copy
    Sheets("eOD").Select
    Range("A1").PasteSpecial Paste:=xlPasteValues
End Sub

Sub separarColumnas()
    Sheets("eOD").Select
    Columns(1).TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
            TextQualifier:=xlDoubleQuote, Other:=True, OtherChar:="|"
End Sub

Sub limpiarEOD()
    Sheets("eOD").Select
    Columns(14).Delete
    Columns("G:J").Delete
    Columns("B:D").Delete
    Columns(5).Cut
    Columns(2).Insert
    Columns(5).Cut
    Columns(3).Insert
    Columns(5).Cut
    Columns(4).Insert
    Columns(5).Insert
    Columns(5).Insert
    Columns(8).Cut
    Columns(7).Insert
    Columns(8).Insert
    Range("A1").FormulaR1C1 = "NRO_OD"
    Range("B1").FormulaR1C1 = "LOCAL"
    Range("C1").FormulaR1C1 = "NRO_LOCAL"
    Range("D1").FormulaR1C1 = "SKU"
    Range("E1").FormulaR1C1 = "ITEM"
    Range("F1").FormulaR1C1 = "ATS"
    Range("G1").FormulaR1C1 = "UNIDADES"
    Range("H1").FormulaR1C1 = "NRO_BULTO"
    Range("I1").FormulaR1C1 = "UPC"
    Range("J1").FormulaR1C1 = "TIPO"
    Range("K1").FormulaR1C1 = "NVENTA"
End Sub

Sub llenarDatos()
    Sheets("eOD").Select
    Range("B2").FormulaR1C1 = "=VLOOKUP(RC[1],Maestras!C[-1]:C,2,0)"
    Range("B2").AutoFill Destination:=Range("B2", Range("B2").End(xlDown))
    Columns(2).Copy
    Columns(2).PasteSpecial Paste:=xlValues
    Range("F2").FormulaR1C1 = "=VLOOKUP(RC[-2],Maestras!C[1]:C[2],2,0)"
    Range("F2").AutoFill Destination:=Range("F2", Range("G2").End(xlDown).Offset(0, -1))
    Columns(6).Copy
    Columns(6).PasteSpecial Paste:=xlValues
End Sub

Sub ordenarPlanilla()
    Sheets("eOD").Select
    Columns(5).Copy
    Columns(8).PasteSpecial Paste:=xlValues
    Columns(4).Copy
    Columns(9).PasteSpecial Paste:=xlValues
    Columns(3).Copy
    Columns(10).PasteSpecial Paste:=xlValues
    Columns(7).Copy
    Columns(12).PasteSpecial Paste:=xlValues
    Columns(6).Copy
    Columns(13).PasteSpecial Paste:=xlValues
    Columns(2).Copy
    Columns(15).PasteSpecial Paste:=xlValues
    Range("K1").FormulaR1C1 = "ITEM"
    Range("N1").FormulaR1C1 = "NRO_BULTO"
    Range("P1").FormulaR1C1 = "TIPO"
    Range("Q1").FormulaR1C1 = "NVENTA"
    Columns("B:G").Delete
    Columns("A:G").AutoFit
End Sub

'Boton Generar Distribución

Sub ordenarPorLocal()
    Sheets("eOD").Select
    With ActiveWorkbook.Worksheets("eOD").Sort
        .SortFields.Clear
        .SortFields.Add _
            Key:=Columns(2), _
            SortOn:=xlSortOnValues, _
            Order:=xlAscending, _
            DataOption:=xlSortNormal
        .SetRange Columns("A:K")
        .Header = xlYes
        .Apply
    End With
End Sub

Sub seleccionarBultos()
    Sheets("eOD").Select
    Range("H2").Select
    Linea = 0
    While ActiveCell.Offset(Linea, -1) <> ""
        If ActiveCell.Offset(Linea, -5).Value <> ActiveCell.Offset(Linea - 1, -5).Value Then
            ActiveCell.Offset(Linea, 0).FormulaR1C1 = 1
            SumaCant = ActiveCell.Offset(Linea, -1).Value
        Else
            SumaCant = SumaCant + ActiveCell.Offset(Linea, -1).Value
            'If SumaCant >= 96 Then
            '    ActiveCell.Offset(Linea + 1, 0).FormulaR1C1 = 1
            '    SumaCant = 0
            'End If
        End If
        Linea = Linea + 1
    Wend
End Sub

Sub numeroBulto()
    Sheets("eOD").Select
    Open ThisWorkbook.Path & "\bfoliost.txt" For Input As #1
        Line Input #1, FolioBulto
    Close #1
    FolioBulto = FolioBulto - 1
    Linea = 0
    Range("H2").Select
    While ActiveCell.Offset(Linea, -1) <> ""
        If ActiveCell.Offset(Linea, 0).Value = 1 Then
            FolioBulto = FolioBulto + 1
        End If
        ActiveCell.Offset(Linea, 0).FormulaR1C1 = FolioBulto
        Linea = Linea + 1
    Wend
    FolioBulto = FolioBulto + 1
    Open ThisWorkbook.Path & "\bfoliost.txt" For Output As #1
        Print #1, FolioBulto
    Close #1
    Range("L2").FormulaR1C1 = "=CONCATENATE(Maestras!R2C3,TEXT(RC[-4],""00000000""))"
    Range("L2").AutoFill Destination:=Range("L2", Range("I1").End(xlDown).Offset(0, 3))
    Range("L2", Range("H1").End(xlDown).Offset(0, 4)).Copy
    Range("H2").PasteSpecial Paste:=xlValues
    Columns(12).ClearContents
    Range("E2").FormulaR1C1 = "=IF(RC[-2]=R[-1]C[-2],R[-1]C+1,1)"
    Range("E2").AutoFill Destination:=Range("E2", Range("D2").End(xlDown).Offset(0, 1))
    Columns(5).Copy
    Range("E1").PasteSpecial Paste:=xlValues
    Columns("A:K").AutoFit
End Sub

Sub limpiarDistribucion()
    Sheets("Distrib").Select
    Range("A4", Range("G3").End(xlDown)).ClearContents
    Range("A4", Range("G3").End(xlDown)).ClearFormats
End Sub

Sub llenarTipoYNotaVenta()
    Sheets("Distrib").Select
    Range("A2").FormulaR1C1 = InputBox("Departamento:")
    Range("G1").FormulaR1C1 = InputBox("Nota de Venta:")
    Range("G2").FormulaR1C1 = Sheets("eOD").Range("A2").Value
End Sub

Sub completarEOD()
    Sheets("Distrib").Select
    Range("A2").Copy
    Sheets("eOD").Select
    Range("J2", Range("I2").End(xlDown).Offset(0, 1)).PasteSpecial Paste:=xlValues
    Sheets("Distrib").Select
    Range("G1").Copy
    Sheets("eOD").Select
    Range("K2", Range("I2").End(xlDown).Offset(0, 2)).PasteSpecial Paste:=xlValues
    Columns("J:K").AutoFit
End Sub

Sub copiarDistribucion()
    Sheets("eOD").Select
    Range(Range("B2:H2"), Range("B2:H2").End(xlDown)).Copy
    Sheets("Distrib").Select
    Range("A4").PasteSpecial Paste:=xlValues
End Sub

Sub formatoDistribucion()
    Sheets("Distrib").Select
    Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
    Selection.Borders(xlEdgeLeft).Weight = xlHairline
    Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
    Selection.Borders(xlEdgeRight).Weight = xlHairline
    Selection.Borders(xlInsideVertical).LineStyle = xlContinuous
    Selection.Borders(xlInsideVertical).Weight = xlHairline
    Range("A4").Select
    While ActiveCell.Value <> ""
        If ActiveCell.Offset(0, 6).Value = ActiveCell.Offset(1, 6).Value Then
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
    If Dir(ThisWorkbook.Path & "\" & notaVenta & ".bat", vbNormal) <> "" Then
        Kill ThisWorkbook.Path & ThisWorkbook.Path & "\" & notaVenta & ".bat"
    End If
    Open ThisWorkbook.Path & "\" & notaVenta & ".bat" For Output As #1
        Print #1, "start " & notaVenta & ".xlsm"
    Close #1
    ActiveWorkbook.Save
    ActiveWorkbook.SaveAs FileName:=ThisWorkbook.Path & "\" & notaVenta & ".xlsm"
End Sub

'Boton Generar Rotulo

Sub llenarDatosDeRotulo()
    Sheets("eOD").Select
    Columns(2).Copy
    Sheets("EtiquetaBulto").Select
    Range("A1").PasteSpecial Paste:=xlValues
    Sheets("eOD").Select
    Columns(8).Copy
    Sheets("EtiquetaBulto").Select
    Range("B1").PasteSpecial Paste:=xlValues
    Sheets("eOD").Select
    Columns("J:K").Copy
    Sheets("EtiquetaBulto").Select
    Range("C1").PasteSpecial Paste:=xlValues
End Sub

Sub quitarRotulosDuplicados()
    Sheets("EtiquetaBulto").Select
    Columns("A:D").EntireColumn.AutoFit
    Columns("A:D").RemoveDuplicates _
        Columns:=2, _
        Header:=xlYes
End Sub

Sub crearNuevoDocumento()
    Sheets("EtiquetaBulto").Select
    If Dir(ThisWorkbook.Path & "\bTottus\eTottus.xls", vbNormal) <> "" Then
        Kill ThisWorkbook.Path & "\bTottus\eTottus.xls"
    End If
    Columns("A:D").Copy
    Workbooks.Add
    Range("A1").PasteSpecial Paste:=xlPasteValues
    ActiveWorkbook.SaveAs FileName:=ThisWorkbook.Path & "\bTottus\eTottus.xls", FileFormat:=xlExcel8
    ActiveWindow.Close
    Sheets("menu").Select
    MsgBox ("Listo para imprimir los rotulos.")
End Sub

'Boton Crear ePIR

Sub limpiarEPIR()
    Sheets("eASN").Select
    Columns("A:AA").ClearContents
End Sub

Sub llenarEPIR()
    Sheets("eASN").Select
    Range("B1").FormulaR1C1 = Sheets("Distrib").Range("G2").Value
    Sheets("Distrib").Select
    Range("H4").FormulaR1C1 = "=IF(RC[-1]=R[-1]C[-1],R[-1]C,R[-1]C+1)"
    Range("H4").Select
    Selection.AutoFill Destination:=Range("H4", Range("G4").End(xlDown).Offset(0, 1))
    Sheets("eASN").Select
    Range("E1").FormulaR1C1 = Sheets("Distrib").Range("H4").End(xlDown).Value
    Sheets("eOD").Select
    Range(Range("I2"), Range("I2").End(xlDown)).Copy
    Sheets("eASN").Select
    Range("B2").PasteSpecial Paste:=xlValues
    Sheets("eOD").Select
    Range(Range("C2"), Range("C2").End(xlDown)).Copy
    Sheets("eASN").Select
    Range("C2").PasteSpecial Paste:=xlValues
    Sheets("eOD").Select
    Range(Range("B2"), Range("B2").End(xlDown)).Copy
    Sheets("eASN").Select
    Range("D2").PasteSpecial Paste:=xlValues
    Sheets("eOD").Select
    Range(Range("G2"), Range("G2").End(xlDown)).Copy
    Sheets("eASN").Select
    Range("E2").PasteSpecial Paste:=xlValues
    Sheets("Distrib").Select
    Range(Range("G4"), Range("G4").End(xlDown)).Copy
    Sheets("eASN").Select
    Range("G2").PasteSpecial Paste:=xlValues
    Range("C1").NumberFormat = "@"
    Range("C1").FormulaR1C1 = InputBox("Fecha de la cita dia-mes-año:")
    Range("D1").FormulaR1C1 = InputBox("Hora de la cita hh:mm:")
    Factura = InputBox("Ingrese el número de factura:")
    Range("I1").FormulaR1C1 = Factura
    Range("J1").FormulaR1C1 = "=CONCATENATE(""1"",""|"",RC[-8],""|"",TEXT(RC[-7],""dd-mm-yyyy""),""|"",TEXT(RC[-6],""hh:mm""),""|"",RC[-5],""|0|0|0|"",RC[-1],""|412"")"
    Range("J2").FormulaR1C1 = "=CONCATENATE(""2"",""|"",RC[-8],""|"",RC[-7],""|"",RC[-6],""|"",RC[-5],""|CJ|"",RC[-3])"
    Range("J2").AutoFill Destination:=Range("J2", Range("G2").End(xlDown).Offset(0, 3))
    Range("J2").End(xlDown).Offset(1, 0).FormulaR1C1 = "3|" & Factura
    Columns(10).Copy
    Columns(10).PasteSpecial Paste:=xlValues
End Sub

Sub crearArchivoEPIR()
    Sheets("eASN").Select
    If Dir(ThisWorkbook.Path & "\bTottus\eASN-" & _
                                Sheets("eASN").Range("B1").Value & "-" & _
                                Sheets("eASN").Range("I1").Value & "-" & _
                                Sheets("Distrib").Range("G1").Value & ".txt", vbNormal) <> "" Then
        Kill ThisWorkbook.Path & "\bTottus\eASN-" & _
                                Sheets("eASN").Range("B1").Value & "-" & _
                                Sheets("eASN").Range("I1").Value & "-" & _
                                Sheets("Distrib").Range("G1").Value & ".txt"
    End If
    Open ThisWorkbook.Path & "\bTottus\eASN-" & _
                                Sheets("eASN").Range("B1").Value & "-" & _
                                Sheets("eASN").Range("I1").Value & "-" & _
                                Sheets("Distrib").Range("G1").Value & ".txt" For Output As #1
    Print #1, Range("J1").Value
    Range("J2").Select
    While ActiveCell.FormulaR1C1 <> ""
        Print #1, ActiveCell.Value
        ActiveCell.Offset(1, 0).Select
    Wend
    Close #1
End Sub

'especial

Sub separar()
    Range("H2").Select
    While ActiveCell.Value <> ""
        i = 1
        fin = ActiveCell.Value
        While i < fin
            ActiveCell.EntireRow.Copy
            ActiveCell.Offset(1, 0).Select
            ActiveCell.EntireRow.Insert
            i = i + 1
        Wend
        ActiveCell.Offset(1, 0).Select
    Wend
    Range("G2").Select
    While ActiveCell.Value <> ""
        ActiveCell.FormulaR1C1 = 18
        ActiveCell.Offset(0, 1).FormulaR1C1 = ""
        ActiveCell.Offset(1, 0).Select
    Wend
End Sub
