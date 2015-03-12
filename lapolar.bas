'procesos

Sub copiarTabla()
    Sheets("ord").Select
    Columns("A:AA").ClearContents
    Sheets("loteLaPolar").Select
    Columns("A:K").Copy
    Sheets("ord").Select
    Range("A1").PasteSpecial Paste:=xlPasteValues
End Sub

Sub ordenarTabla()
    Sheets("ord").Select
    If Range("A1").Value = "General" Then
        Rows(1).Delete
    End If
    If Range("K1").End(xlDown).Offset(0, -10).Value = "" Then
        Range("K1").End(xlDown).EntireRow.ClearContents
    End If
    Columns("J:K").ClearContents
    With ActiveWorkbook.Worksheets("ord").Sort
        .SortFields.Clear
        .SortFields.Add Key:=Columns(6), SortOn:=xlSortOnValues, Order:=xlAscending
        .SetRange Columns("A:I")
        .Header = xlYes
        .Orientation = xlTopToBottom
        .Apply
    End With
End Sub

Sub botonTransformarTabla()
    Application.ScreenUpdating = False
    If MsgBox("¿Ha copiado previamente el lote descargado del sistem B2B de La Polar?", _
    vbYesNo, "Antes de continuar...") = vbYes Then
        Call copiarTabla
        Call ordenarTabla
        If MsgBox("¿Desea modificar las cantidades antes de generar la distribución?", _
        vbYesNo, "Antes de continuar...") = vbYes Then
            Sheets("ord").Select
            Range("A1").Select
        Else
            Call botonGenerarDistribucion
        End If
    Else
        Sheets("ord").Select
        MsgBox ("Copie y pegue la información del lote en la hoja loteLaPolar.")
    End If
    Application.ScreenUpdating = True
End Sub

Sub ordenarPorLocal()
    Sheets("ord").Select
    With ActiveWorkbook.Worksheets("ord").Sort
        .SortFields.Clear
        .SortFields.Add Key:=Columns(4), SortOn:=xlSortOnValues, Order:=xlAscending
        .SetRange Columns("A:I")
        .Header = xlYes
        .Orientation = xlTopToBottom
        .Apply
    End With
End Sub

Sub agregarATS()
    Sheets("ord").Select
    Range("G1").FormulaR1C1 = "ATS"
    Range("G2").FormulaR1C1 = "=VLOOKUP(VALUE(RC[-1]),mae!C[-6]:C[-5],2,0)"
    If Range("F3").Value <> "" Then
        Range("G2").AutoFill Destination:=Range("G2", Range("F1").End(xlDown).Offset(0, 1))
    End If
    Columns(7).Copy
    Range("G1").PasteSpecial Paste:=xlPasteValues
End Sub

Sub agregarBLK()
    Sheets("ord").Select
    Range("J1").FormulaR1C1 = "BLK"
    Range("J2").FormulaR1C1 = "=VLOOKUP(VALUE(RC[-7]),mae!C[-5]:C[-4],2,0)"
    If Range("I3").Value <> "" Then
        Range("J2").AutoFill Destination:=Range("J2", Range("I1").End(xlDown).Offset(0, 1))
    End If
    Columns(10).Copy
    Range("J1").PasteSpecial Paste:=xlPasteValues
End Sub

Sub copiarDistribucion()
    Sheets("dis").Select
    Columns("A:AA").ClearContents
    Columns("A:AA").ClearFormats
    Sheets("ord").Select
    Columns("A:I").Copy
    Sheets("dis").Select
    Range("A1").PasteSpecial Paste:=xlPasteValues
End Sub

Sub formatoTituloDistribucion()
    Sheets("dis").Select
    Rows(1).Insert
    Rows(1).Insert
    Columns(5).Delete
    Columns(2).Delete
    Range("F1").FormulaR1C1 = "Nota Venta"
    Range("F2").FormulaR1C1 = "Orden de Compra"
    Range("G2").FormulaR1C1 = Range("A4").Value
    Range("G1").FormulaR1C1 = InputBox("Ingrese el numero de la nota de venta", "Antes de continuar", "123456")
    Range("F1").HorizontalAlignment = xlRight
    Range("F2").HorizontalAlignment = xlRight
    Range("B1:E2").Merge
    Range("B1:E2").HorizontalAlignment = xlCenter
    Range("B1:E2").VerticalAlignment = xlCenter
    Range("B1").FormulaR1C1 = "DISTRIBUCION LA POLAR"
    Range("B1").Font.Bold = True
    Range("G1").Font.Bold = True
    Range("G2").Font.Bold = True
    Range("B3:G3,F1:F2").Interior.ThemeColor = xlThemeColorLight1
    Range("B3:G3,F1:F2").Font.ThemeColor = xlThemeColorDark1
    Range("B1:G3").Borders(xlEdgeLeft).LineStyle = xlContinuous
    Range("B1:G3").Borders(xlEdgeLeft).Weight = xlThin
    Range("B1:G3").Borders(xlEdgeTop).LineStyle = xlContinuous
    Range("B1:G3").Borders(xlEdgeTop).Weight = xlThin
    Range("B1:G3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B1:G3").Borders(xlEdgeBottom).Weight = xlThin
    Range("B1:G3").Borders(xlEdgeRight).LineStyle = xlContinuous
    Range("B1:G3").Borders(xlEdgeRight).Weight = xlThin
    Range("B1:G3").Borders(xlInsideVertical).LineStyle = xlContinuous
    Range("B1:G3").Borders(xlInsideVertical).Weight = xlThin
    Range("B1:G3").Borders(xlInsideHorizontal).LineStyle = xlContinuous
    Range("B1:G3").Borders(xlInsideHorizontal).Weight = xlThin
    Columns(1).Delete
    Range("A3").FormulaR1C1 = "LOC"
    Columns("A:F").AutoFit
End Sub

Sub separacionPorTienda()
    Sheets("dis").Select
    Range("A4", Range("F4").End(xlDown)).Borders(xlInsideVertical).Weight = xlHairline
    Range("A4").Select
    While ActiveCell <> ""
        If ActiveCell.Offset(1, 0).Value <> ActiveCell.Value Then
            Range(ActiveCell, ActiveCell.Offset(0, 5)).Borders(xlEdgeBottom).Weight = xlThick
        Else
            Range(ActiveCell, ActiveCell.Offset(0, 5)).Borders(xlEdgeBottom).Weight = xlHairline
        End If
        ActiveCell.Offset(1, 0).Select
    Wend
    Range("A1").Select
End Sub

Sub guardarDocumento()
    On Error Resume Next
    notaVenta = Sheets("dis").Range("F1").Value
    If Dir(ThisWorkbook.Path & "\" & notaVenta & ".bat", vbNormal) <> "" Then
        Kill ThisWorkbook.Path & ThisWorkbook.Path & "\" & notaVenta & ".bat"
    End If
    Open ThisWorkbook.Path & "\" & notaVenta & ".bat" For Output As #1
        Print #1, "start " & notaVenta & ".xlsm"
    Close #1
    ActiveWorkbook.SaveAs Filename:=ThisWorkbook.Path & "\" & notaVenta & ".xlsm"
End Sub

Sub imprimirDistribucion()
    Sheets("dis").Select
    Range("A1").Select
    ActiveWindow.SelectedSheets.PrintOut
End Sub

Sub botonGenerarDistribucion()
    Application.ScreenUpdating = False
    Call ordenarPorLocal
    Call agregarATS
    Call agregarBLK
    Call copiarDistribucion
    Call formatoTituloDistribucion
    Call separacionPorTienda
    Call guardarDocumento
    If MsgBox("¿Desea imprimir la distribución?", vbYesNo, "Antes de continuar...") = vbYes Then
        Call imprimirDistribucion
    End If
    Sheets("Menu").Select
    Application.ScreenUpdating = True
End Sub

Sub botonSeleccionarBultos()
    Application.ScreenUpdating = False
    Sheets("dis").Select
    Range("G4").FormulaR1C1 = "=IF(RC[-6]<>R[-1]C[-6],1,"""")"
    If Range("F5").Value <> "" Then
        Range("G4").AutoFill Destination:=Range("G4", Range("F4").End(xlDown).Offset(0, 1))
    End If
    Columns(7).Copy
    Range("G1").PasteSpecial Paste:=xlPasteValues
    Range("G4").Select
    MsgBox ("Agrege un 1 al inicio de cada bulto que no se haya identificado automaticamente.")
    Application.ScreenUpdating = True
End Sub

Sub copiarDistribucion()
    Sheets("plist").Select
    Columns("A:AA").ClearContents
    Sheets("dis").Select
    Columns("A:G").Copy
    Sheets("plist").Select
    Range("A1").PasteSpecial Paste:=xlPasteValues
    Rows("1:3").Delete
End Sub

Sub completarColumnas()
    Sheets("plist").Select
    Columns("A:B").Insert
    Sheets("ord").Select
    Range("A2:B2").Copy
    Sheets("plist").Select
    Range("A1", Range("C1").End(xlDown).Offset(0, -1)).PasteSpecial Paste:=xlPasteValues
    Columns(6).ClearContents
    Range("F1").FormulaR1C1 = "=IFERROR(VLOOKUP(VALUE(RC[-1]),packingLaPolar!C[-1]:C,2,0),"""")"
    If Range("E2").Value <> "" Then
        Range("F1").AutoFill Destination:=Range("F1", Range("E1").End(xlDown).Offset(0, 1))
    End If
    Columns(6).Copy
    Range("F1").PasteSpecial Paste:=xlPasteValues
    Range("J1").FormulaR1C1 = "=VALUE(RIGHT(packingLaPolar!R[1]C[-1],8))"
    Range("J2").FormulaR1C1 = "=IF(RC[-1]=1,IF(RC[-6]<>R[-1]C[-6],VALUE(RIGHT(packingLaPolar!R[1]C[-1],8)),R[-1]C+1),R[-1]C)"
    If Range("H3").Value <> "" Then
        Range("J2").AutoFill Destination:=Range("J2", Range("H2").End(xlDown).Offset(0, 2))
    End If
    Columns(10).Copy
    Range("J1").PasteSpecial Paste:=xlPasteValues
    Range("I1").FormulaR1C1 = "=CONCATENATE(""1090914000"",TEXT(RC[1],""00000000""))"
    If Range("H2").Value <> "" Then
        Range("I1").AutoFill Destination:=Range("I1", Range("H1").End(xlDown).Offset(0, 1))
    End If
    Columns(9).Copy
    Range("I1").PasteSpecial Paste:=xlPasteValues
    Range("J1").FormulaR1C1 = "-1"
    Range("J1").Copy
    Range("J1", Range("J1").End(xlDown)).PasteSpecial Paste:=xlPasteValues
    Columns("A:J").AutoFit
End Sub

Sub copiarInformacion()
    Sheets("plist").Select
    Range("A1", Range("J1").End(xlDown)).Copy
    MsgBox ("Suba este packing list al sistema La Polar")
End Sub

Sub botonGenerarPackingList()
    Application.ScreenUpdating = False
    Call copiarDistribucion
    Call completarColumnas
    Call copiarInformacion
    Application.ScreenUpdating = True
End Sub

Sub copiarPacking()
    Sheets("etq").Select
    Columns("A:AC").ClearContents
    Sheets("plist").Select
    Columns("A:J").Copy
    Sheets("etq").Select
    Range("A1").PasteSpecial Paste:=xlValues
End Sub

Sub ordenarRotulo()
    Sheets("etq").Select
    Rows(1).Insert
    Columns(10).Delete
    Columns(6).Delete
    Columns("B:C").Delete
    Range("A1").FormulaR1C1 = "OCOMPRA"
    Range("B1").FormulaR1C1 = "LOCAL"
    Range("C1").FormulaR1C1 = "PLU"
    Range("D1").FormulaR1C1 = "DESCRIP"
    Range("E1").FormulaR1C1 = "CANT"
    Range("F1").FormulaR1C1 = "CODBARRA"
    Range("G1").FormulaR1C1 = "NVENTA"
End Sub

Sub completarRotulo()
    Sheets("etq").Select
    Range("G2").FormulaR1C1 = InputBox("Ingrese el número de Nota de Venta.")
    Range("G2").Copy
    If Range("A3").Value <> "" Then
        Range("G2", Range("F1").End(xlDown).Offset(0, 1)).PasteSpecial Paste:=xlValues
    End If
    Range("H1").FormulaR1C1 = "LOTE"
    Range("H2").FormulaR1C1 = InputBox("Ingrese el número de Lote:")
    Range("H2").Copy
    If Range("A3").Value <> "" Then
    Range("H2", Range("G1").End(xlDown).Offset(0, 1)).PasteSpecial Paste:=xlValues
    End If
    Range("J1").FormulaR1C1 = "NLINEA"
    Range("I2").Select
    While ActiveCell.Offset(0, -1).Value <> ""
        If ActiveCell.Offset(0, -3).Value <> ActiveCell.Offset(-1, -3).Value Then
            ActiveCell.FormulaR1C1 = 1
        Else
            ActiveCell.FormulaR1C1 = ActiveCell.Offset(-1, 0).Value + 1
        End If
        ActiveCell.Offset(1, 0).Select
    Wend
    Range("J2").Select
    Range("J2").FormulaR1C1 = "=TEXT(RC[-1],""00"")"
    Range("J2").AutoFill Destination:=Range("J2", Range("H1").End(xlDown).Offset(0, 2))
    Columns("J:J").Copy
    Range("I1").PasteSpecial Paste:=xlValues
    Columns("J:J").Delete
    Columns("A:I").AutoFit
    Range("J1").FormulaR1C1 = "DESCRIP"
    Range("J2").Select
    Range("J2").FormulaR1C1 = _
    "=LEFT(CONCATENATE(RC[-6],""000000000000000000000000000000000000000000000""),40)"
    Range("J2").AutoFill Destination:=Range("J2", Range("H1").End(xlDown).Offset(0, 2))
    Columns("J:J").Copy
    Range("D1").PasteSpecial Paste:=xlValues
    Columns("J:J").Delete
    Range("J1").FormulaR1C1 = "CANT"
    Range("J2").Select
    Range("J2").FormulaR1C1 = "=RIGHT(CONCATENATE(""0000"",RC[-5]),4)"
    Range("J2").AutoFill Destination:=Range("J2", Range("H1").End(xlDown).Offset(0, 2))
    Columns("J:J").Copy
    Range("E1").PasteSpecial Paste:=xlValues
    Columns("J:J").Delete
    Range("J1").FormulaR1C1 = "LINEA01"
    Range("J1").EntireColumn.AutoFit
    Range("J2").Select
    Range("J2").FormulaR1C1 = "=CONCATENATE(RC[-1],""  "",RC[-6],""  "",RC[-7],""  "",RC[-5])"
    Range("J2").AutoFill Destination:=Range("J2", Range("H1").End(xlDown).Offset(0, 2))
    Columns("J:J").Copy
    Range("J1").PasteSpecial Paste:=xlValues
    Columns("I:I").Delete
    Columns("C:D").Delete
    Range("H1").FormulaR1C1 = "CANT"
    Range("H2").Select
    Range("H2").FormulaR1C1 = "=VALUE(RC[-5])"
    Range("H2").AutoFill Destination:=Range("H2", Range("F1").End(xlDown).Offset(0, 2))
    Columns("H:H").Copy
    Range("C1").PasteSpecial Paste:=xlValues
    Columns("H:H").Delete
    Columns("G:G").Insert
    Range("I1").FormulaR1C1 = "LINEA02"
    Range("J1").FormulaR1C1 = "LINEA03"
    Range("K1").FormulaR1C1 = "LINEA04"
    Range("L1").FormulaR1C1 = "LINEA05"
    Range("M1").FormulaR1C1 = "LINEA06"
    Range("N1").FormulaR1C1 = "LINEA07"
    Range("O1").FormulaR1C1 = "LINEA08"
    Range("P1").FormulaR1C1 = "LINEA09"
    Range("Q1").FormulaR1C1 = "LINEA10"
    Range("R1").FormulaR1C1 = "LINEA11"
    Range("S1").FormulaR1C1 = "LINEA12"
    Range("T1").FormulaR1C1 = "LINEA13"
    Range("U1").FormulaR1C1 = "LINEA14"
    Range("V1").FormulaR1C1 = "LINEA15"
    Range("W1").FormulaR1C1 = "LINEA16"
    Range("X1").FormulaR1C1 = "LINEA17"
    Range("Y1").FormulaR1C1 = "LINEA18"
    Range("Z1").FormulaR1C1 = "LINEA19"
    Range("AA1").FormulaR1C1 = "LINEA20"
    Columns("I:AA").AutoFit
    Range("H2").Select
    tUnidades = 0
    While ActiveCell.Value <> ""
        If ActiveCell.Offset(0, -4).Value <> ActiveCell.Offset(-1, -4).Value Then
            ActiveCell.Offset(-1, -1).FormulaR1C1 = tUnidades
            tUnidades = 0
            tUnidades = tUnidades + ActiveCell.Offset(0, -5).Value
            ActiveCell.Offset(1, 0).Select
            Columna = 1
        Else
            tUnidades = tUnidades + ActiveCell.Offset(0, -5).Value
            ActiveCell.Offset(-1, Columna).FormulaR1C1 = ActiveCell.Value
            Columna = Columna + 1
            ActiveCell.EntireRow.Delete
        End If
    Wend
    ActiveCell.Offset(-1, -1).FormulaR1C1 = tUnidades
    Range("G1").FormulaR1C1 = "TUNIDS"
    Range("G1").EntireColumn.AutoFit
    Columns("C:C").Delete
    Columns("G:G").Insert
    Range("G1").FormulaR1C1 = "CAJA"
    Range("G2").Select
    While ActiveCell.Offset(0, -1).Value <> ""
        If ActiveCell.Offset(0, -5).Value <> ActiveCell.Offset(-1, -5).Value Then
            ActiveCell.FormulaR1C1 = 1
        Else
            ActiveCell.FormulaR1C1 = ActiveCell.Offset(-1, 0).Value + 1
        End If
        ActiveCell.Offset(1, 0).Select
    Wend
    Columns("H:H").Insert
    Range("H1").FormulaR1C1 = "TOTCAJAS"
    Range("H2").Select
    ActiveCell.FormulaR1C1 = "=COUNTIF(C[-6],RC[-6])"
    Range("H2").AutoFill Destination:=Range("H2", Range("F1").End(xlDown).Offset(0, 2))
    Columns("H:H").Copy
    Range("H1").PasteSpecial Paste:=xlValues
    Columns("I:I").Insert
    Range("I1").FormulaR1C1 = "DESPAC"
    Range("I2").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-7],mae!C[-5]:C[-3],3,0)"
    Range("I2").AutoFill Destination:=Range("I2", Range("G1").End(xlDown).Offset(0, 2))
    Columns("I:I").Copy
    Range("I1").PasteSpecial Paste:=xlValues
    Range("C2").Select
    While ActiveCell.Value <> ""
        ActiveCell.FormulaR1C1 = "<FNC1>" & ActiveCell.Value
        ActiveCell.Offset(1, 0).Select
    Wend
    Columns("G:I").AutoFit
    Application.CutCopyMode = False
End Sub

Sub botonCrearArchivoEtiqueta()
    Application.ScreenUpdating = False
    Call copiarPacking
    Call ordenarRotulo
    Call completarRotulo
    If Dir(ThisWorkbook.Path & "\bLaPolar\eLaPolar.xls", vbNormal) <> "" Then
        Kill ThisWorkbook.Path & "\bLaPolar\eLaPolar.xls"
    End If
    Columns("A:AC").Copy
    Workbooks.Add
    Range("A1").PasteSpecial Paste:=xlPasteValues
    ActiveWorkbook.SaveAs Filename:=ThisWorkbook.Path & "\bLaPolar\eLaPolar.xls", FileFormat:=xlExcel8
    ActiveWindow.Close
    Sheets("menu").Select
    MsgBox ("Listo para imprimir los rotulos.")
    Application.ScreenUpdating = True
End Sub
