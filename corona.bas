'procesos

Sub ConvertirTabla()
    Application.ScreenUpdating = False
    Sheets("PreDist").Select
    Columns("A:AA").ClearContents
    Sheets("bCorona").Select
    Columns("A:AA").Copy
    Sheets("PreDist").Select
    Range("A1").PasteSpecial xlPasteValues
    Sheets("PreDist").Select
    If Range("C1").End(xlDown).Value = "Total General" Then
        Range("C1").End(xlDown).EntireRow.Delete
    End If
    Range("A1").End(xlToRight).EntireColumn.Delete
    Columns("C:C").Insert
    Range("C1").FormulaR1C1 = "Codigo SKU"
    Range("C2").Select
    ActiveCell.FormulaR1C1 = "=RIGHT(LEFT(RC[-1],20),14)"
    If Range("A3").Value <> "" Then
    Selection.AutoFill Destination:=Range("C2", _
        Range("B1").End(xlDown).Offset(0, 1))
    End If
    Columns("C:C").Copy
    Range("B1").PasteSpecial Paste:=xlValues
    Columns("C:C").Delete
    Columns("D:E").Insert
    Range("D1").Select
    ActiveCell.FormulaR1C1 = "LOCAL"
    Range("E1").FormulaR1C1 = "CANT"
    ActiveCell.Offset(1, 0).Select
    Solicitados = Range("A2", Range("C1").End(xlDown)).AddressLocal
    Lineas = Range("A2", Range("A1").End(xlDown)).Count
    i = 0
    While Range("F1").Value <> ""
        While ActiveCell.Offset(0, -1).Value <> ""
            ActiveCell.FormulaR1C1 = Range("F1").Value
            ActiveCell.Offset(0, 1).FormulaR1C1 = ActiveCell.Offset(0 - i, 2).Value
            ActiveCell.Offset(1, 0).Select
        Wend
        If Range("G1").Value <> "" Then
            Range(Solicitados).Copy
            Range("A1").End(xlDown).Offset(1, 0).PasteSpecial _
                Paste:=xlPasteValues, _
                Operation:=xlNone, _
                SkipBlanks:=False, _
                Transpose:=False
            Range("D1").End(xlDown).Offset(1, 0).Select
        End If
        Columns("F:F").EntireColumn.Delete
        i = i + Lineas
    Wend
    Range("A1").FormulaR1C1 = "OCOMPRA"
    Range("B1").FormulaR1C1 = "SKU"
    Range("C1").FormulaR1C1 = "DESCRIP"
    Range("F1").FormulaR1C1 = "ATS"
    Range("F2").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-4],Maestras!C[-5]:C[-3],2,0)"
    Selection.AutoFill Destination:=Range("F2", Range("E1").End(xlDown).Offset(0, 1))
    Columns("F:F").Select
    Selection.Copy
    Selection.PasteSpecial _
        Paste:=xlPasteValues, _
        Operation:=xlNone, _
        SkipBlanks:=False, _
        Transpose:=False
    Columns("D:D").Cut
    Columns("B:B").Insert Shift:=xlToRight
    Columns("F:F").Cut
    Columns("D:D").Insert Shift:=xlToRight
    Columns("A:F").AutoFit
    Range("F2").Select
    While ActiveCell.Value <> ""
        If ActiveCell.Value = 0 Then
            ActiveCell.EntireRow.Delete
        Else
        ActiveCell.Offset(1, 0).Select
        End If
    Wend
    
    With ActiveWorkbook.Worksheets("PreDist").Sort
        .SortFields.Clear
        .SortFields.Add Key:=Columns("C:C"), SortOn:=xlSortOnValues, Order:=xlAscending
        .SetRange Columns("A:F")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    If MsgBox("Desea modificar las cantidades?", vbYesNo) = vbYes Then
        Sheets("Predist").Select
        Range("A1").Select
    Else
        Call LlenarDistribucion
    End If
    Application.ScreenUpdating = True
End Sub

Sub LlenarDistribucion()
    Application.ScreenUpdating = False
    Sheets("Distrib").Select
    Range("A4", Range("A4").End(xlDown).Offset(0, 17)).ClearContents
    Range("A4", Range("A4").End(xlDown).Offset(0, 17)).ClearFormats
    Range("F1").FormulaR1C1 = InputBox("Ingrese el número de Nota de Venta.")
    Range("F2").FormulaR1C1 = Sheets("PreDist").Range("A2").Value
    Sheets("PreDist").Select
    With ActiveWorkbook.Worksheets("PreDist").Sort
        .SortFields.Clear
        .SortFields.Add Key:=Columns("B:B"), SortOn:=xlSortOnValues, Order:=xlAscending
        .SetRange Columns("A:F")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("B2", Range("B2").End(xlDown)).Copy
    Sheets("Distrib").Select
    Range("A4").PasteSpecial _
        Paste:=xlPasteValues, _
        Operation:=xlNone, _
        SkipBlanks:=False, _
        Transpose:=False
    Range("B4").Select
    While ActiveCell.Offset(0, -1).Value <> ""
        If ActiveCell.Offset(0, -1).Value <> ActiveCell.Offset(-1, -1).Value Then
            ActiveCell.FormulaR1C1 = 1
            ActiveCell.Offset(1, 0).Select
        Else
            ActiveCell.FormulaR1C1 = ActiveCell.Offset(-1, 0).Value + 1
            ActiveCell.Offset(1, 0).Select
        End If
    Wend
    Sheets("PreDist").Select
    Range("C2", Range("F2").End(xlDown)).Copy
    Sheets("Distrib").Select
    Range("C4").PasteSpecial _
        Paste:=xlPasteValues, _
        Operation:=xlNone, _
        SkipBlanks:=False, _
        Transpose:=False
    Range("A4", Range("F4").End(xlDown)).Select
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
    On Error Resume Next
    notaVenta = Sheets("Distrib").Range("F1").Value
    ActiveWorkbook.SaveAs Filename:=ThisWorkbook.Path & "\" & notaVenta & ".xlsm"
    If Dir(ThisWorkbook.Path & "\" & notaVenta & ".bat", vbNormal) <> "" Then
        Kill ThisWorkbook.Path & "\" & notaVenta & ".bat"
    End If
    Open ThisWorkbook.Path & "\" & notaVenta & ".bat" For Output As #1
        Print #1, "start " & notaVenta & ".xlsm"
    Close #1
    If MsgBox("Desea imprimir la distribucion?", vbYesNo) = vbYes Then
        Call Imprimir
    Else
        Sheets("Menu").Select
        MsgBox ("Finalizado")
    End If
    Application.ScreenUpdating = True
End Sub

Sub Imprimir()
    Application.ScreenUpdating = False
    Sheets("Distrib").Select
    ActiveWindow.SelectedSheets.PrintOut
    Sheets("Menu").Select
    MsgBox ("Enviando a impresora.")
    Application.ScreenUpdating = True
End Sub

Sub Rotulo()
    Application.ScreenUpdating = False
    Range("A9").Select
    Sheets("Predist").Select
    Range("G1").FormulaR1C1 = "T"
    Range("H1").FormulaR1C1 = "C"
    Range("I1").FormulaR1C1 = "D"
    Range("J1").FormulaR1C1 = "U"
    Range("G2").Select
    While ActiveCell.Offset(0, -1).Value <> ""
        ActiveCell.FormulaR1C1 = "=VLOOKUP(MID(RC[-3],8,3),Maestras!C[-2]:C[-1],2,0)"
        ActiveCell.Offset(0, 1).FormulaR1C1 = "=VLOOKUP(MID(RC[-4],12,3),Maestras!C[0]:C[1],2,0)"
        ActiveCell.Offset(0, 2).FormulaR1C1 = "=MID(CONCATENATE(RC[-4],""                    ""),1,20)"
        ActiveCell.Offset(0, 3).FormulaR1C1 = "=MID(CONCATENATE(TEXT(RC[-4],""0""),""   ""),1,4)"
        ActiveCell.Offset(1, 0).Select
    Wend
    Range("G:J").Copy
    Range("G:J").PasteSpecial _
        Paste:=xlPasteValues, _
        Operation:=xlNone, _
        SkipBlanks:=False, _
        Transpose:=False
    Range("K1").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "LINEA"
    ActiveCell.Offset(1, 0).Select
    While ActiveCell.Offset(0, -1).Value <> ""
        ActiveCell.FormulaR1C1 = _
            ActiveCell.Offset(0, -7).Value & " " & _
            ActiveCell.Offset(0, -8).Value & " " & _
            ActiveCell.Offset(0, -1).Value & " " & _
            ActiveCell.Offset(0, -2).Value & " " & _
            ActiveCell.Offset(0, -4).Value & " " & _
            ActiveCell.Offset(0, -3).Value
        ActiveCell.Offset(1, 0).Select
    Wend
    Columns("G:J").Delete
    Range("H2").Select
    While ActiveCell.Offset(0, -1).Value <> ""
        If ActiveCell.Offset(0, -6).Value = ActiveCell.Offset(-1, -6).Value Then
            ActiveCell.FormulaR1C1 = ""
            ActiveCell.Offset(1, 0).Select
        Else
            ActiveCell.FormulaR1C1 = 1
            ActiveCell.Offset(1, 0).Select
        End If
    Wend
    Columns("A:H").EntireColumn.AutoFit
    Range("H2").Select
    MsgBox ("Agrege un 1 al inicio de cada bulto que no haya sido identificado")
    Application.ScreenUpdating = True
End Sub

Sub GenerarLabel()
    Application.ScreenUpdating = False
    Range("A11").Select
    Sheets("Rotulo").Select
    Columns("A:AB").EntireColumn.Delete
    Sheets("PreDist").Select
    Columns("B:C").Copy
    Sheets("Rotulo").Select
    Range("A1").PasteSpecial Paste:=xlPasteValues
    Sheets("PreDist").Select
    Columns("E:F").Copy
    Sheets("Rotulo").Select
    Range("C1").PasteSpecial Paste:=xlPasteValues
    
    Range("E1").FormulaR1C1 = "DESC"
    Range("E2").FormulaR1C1 = "=CONCATENATE(TEXT(RC[-3],0),""      "",LEFT(TEXT(RC[-2],0),16),""      "",TEXT(RC[-1],0))"
    If Range("A3").Value <> "" Then
        Range("E2").AutoFill Destination:=Range("E2", Range("D1").End(xlDown).Offset(0, 1))
    End If
    Columns("E:E").Copy
    Range("E1").PasteSpecial xlPasteValues
    Columns("B:D").Delete
    Sheets("PreDist").Select
    Columns("H:H").Copy
    Sheets("Rotulo").Select
    Range("C1").PasteSpecial xlPasteValues
    
    Range("C1").FormulaR1C1 = "BULTOS"
    Range("D1").FormulaR1C1 = "LINE01"
    Range("E1").FormulaR1C1 = "LINE02"
    Range("F1").FormulaR1C1 = "LINE03"
    Range("G1").FormulaR1C1 = "LINE04"
    Range("H1").FormulaR1C1 = "LINE05"
    Range("I1").FormulaR1C1 = "LINE06"
    Range("J1").FormulaR1C1 = "LINE07"
    Range("K1").FormulaR1C1 = "LINE08"
    Range("L1").FormulaR1C1 = "LINE09"
    Range("M1").FormulaR1C1 = "LINE10"
    Range("N1").FormulaR1C1 = "LINE11"
    Range("O1").FormulaR1C1 = "LINE12"
    Range("P1").FormulaR1C1 = "LINE13"
    Range("Q1").FormulaR1C1 = "LINE14"
    Range("R1").FormulaR1C1 = "LINE15"
    Range("S1").FormulaR1C1 = "LINE16"
    Range("T1").FormulaR1C1 = "LINE17"
    Range("U1").FormulaR1C1 = "LINE18"
    Range("V1").FormulaR1C1 = "LINE19"
    Range("W1").FormulaR1C1 = "LINE20"
    Columns("A:C").EntireColumn.AutoFit
    Range("D2").Select
    tBultos = 0
    While ActiveCell.Offset(0, -2).Value <> ""
        If ActiveCell.Offset(0, -1).Value = 1 Then
            ActiveCell.FormulaR1C1 = ActiveCell.Offset(0, -2).Value
            ActiveCell.Offset(1, 0).Select
            Columna = 1
            tBultos = tBultos + 1
        Else
            ActiveCell.Offset(-1, Columna).FormulaR1C1 = ActiveCell.Offset(0, -2).Value
            Columna = Columna + 1
            ActiveCell.EntireRow.Delete
        End If
    Wend
    Range("B1").Select
    Selection.EntireColumn.Delete
    Selection.EntireColumn.Insert
    Selection.EntireColumn.Insert
    Selection.EntireColumn.Insert
    Range("B1").FormulaR1C1 = "NVENTA"
    Range("C1").FormulaR1C1 = "CAJANRO"
    Range("D1").FormulaR1C1 = "TOTCAJAS"
    Range("D2").Select
    Factura = 0
    While ActiveCell.Offset(0, -3).Value <> ""
        ActiveCell.FormulaR1C1 = "=COUNTIF(C[-3],RC[-3])"
        ActiveCell.Offset(0, -1).FormulaR1C1 = "=IF(RC[-2]=R[-1]C[-2],R[-1]C+1,1)"
        ActiveCell.Offset(0, -2).FormulaR1C1 = Sheets("Distrib").Range("F1").Value
        ActiveCell.Offset(1, 0).Select
    Wend
    Columns("B:D").Copy
    Columns("B:D").PasteSpecial Paste:=xlPasteValues
    Columns("E:E").Delete
    Columns("B:C").Insert
    
    Sheets("Subir").Select
    Columns("H:H").Copy
    Sheets("Rotulo").Select
    Range("B1").PasteSpecial xlPasteValues
    Sheets("Subir").Select
    Columns("C:C").Copy
    Sheets("Rotulo").Select
    Range("C1").PasteSpecial xlPasteValues
    Range("B1").FormulaR1C1 = "FECHA"
    Range("C1").FormulaR1C1 = "LPN"
    Columns("C:C").NumberFormat = "0"
    Columns("B:B").NumberFormat = "dd-mm-yyyy"
    Range("C2").Select
    While ActiveCell.Value <> ""
        If ActiveCell.Value = ActiveCell.Offset(-1, 0).Value Then
            Range(ActiveCell.Offset(0, -1), ActiveCell).Delete Shift:=xlUp
            ActiveCell.Offset(-1, 0).Select
        End If
        ActiveCell.Offset(1, 0).Select
    Wend
    'Crea Archivo Etiqueta
    Sheets("Rotulo").Select
    Columns("A:Z").Copy
    Workbooks.Add
    Range("A1").PasteSpecial Paste:=xlPasteValues
    ActiveWorkbook.SaveAs Filename:=ThisWorkbook.Path & "\bCorona\eCorona.xls" _
        , FileFormat:=xlExcel8, Password:="", WriteResPassword:="", _
        ReadOnlyRecommended:=False, CreateBackup:=False
    ActiveWindow.Close
    'Crea PL Corona
    Sheets("Subir").Select
    Columns("A:H").Copy
    Workbooks.Add
    Range("A1").PasteSpecial Paste:=xlPasteValues
    ActiveWorkbook.SaveAs Filename:=ThisWorkbook.Path & "\bCorona\PLCorona.xls" _
        , FileFormat:=xlExcel8, Password:="", WriteResPassword:="", _
        ReadOnlyRecommended:=False, CreateBackup:=False
    ActiveWindow.Close
    Sheets("Menu").Select
    MsgBox ("Archivo de etiqueta y Packing List creados exitosamente")
    Application.ScreenUpdating = True
End Sub

Sub newPlist()
    Application.ScreenUpdating = False
    If MsgBox("Primero debe pegar los numeros de bultos bajados desde el portal B2B, ¿Desea continuar?", vbYesNo, "Antes de continuar...") = vbYes Then
        Call setMaestra
        Call setTitle
        Call setPreDist
        Call fillUp
        Call GenerarLabel
    End If
    MsgBox ("Ingrese al portal y baje el archivo de etiquetas necesarias para realizar este proceso.")
    Application.ScreenUpdating = True
End Sub

Sub setMaestra()
    Application.ScreenUpdating = False
    Sheets("Maestras").Select
    Range("O:Q").ClearContents
    Sheets("bCorona").Select
    Range("B2", Range("B1").End(xlDown)).Copy
    Sheets("Maestras").Select
    Range("P1").PasteSpecial xlPasteValues
    Range("O1").Select
    While ActiveCell.Offset(0, 1).Value <> ""
        ActiveCell.FormulaR1C1 = Right(Left(ActiveCell.Offset(0, 1).Value, 20), 14)
        ActiveCell.Offset(0, 2).FormulaR1C1 = Right(Left(ActiveCell.Offset(0, 1).Value, 20), 14)
        ActiveCell.Offset(1, 0).Select
    Wend
    Application.ScreenUpdating = True
End Sub

Sub setTitle()
    Application.ScreenUpdating = False
    Sheets("Subir").Select
    Columns("A:AA").ClearContents
    Range("A1").FormulaR1C1 = "N DE OC"
    Range("B1").FormulaR1C1 = "COD LOCAL DESTINO"
    Range("C1").FormulaR1C1 = "LPN"
    Range("D1").FormulaR1C1 = "COD DEL PRODUCTO"
    Range("E1").FormulaR1C1 = "CANTIDAD"
    Range("F1").FormulaR1C1 = "TIPO DE DOCUMENTO"
    Range("G1").FormulaR1C1 = "N DE DOCUMENTO"
    Range("H1").FormulaR1C1 = "FECHA DEL DOCUMENTO"
    Application.ScreenUpdating = True
End Sub

Sub setPreDist()
    Application.ScreenUpdating = False
    Sheets("PreDist").Select
    Columns("J:J").ClearContents
    Columns("J:J").NumberFormat = "0"
    Range("J2").Select
    i = 1
    While ActiveCell.Offset(0, -3).Value <> ""
        If ActiveCell.Offset(0, -2).Value = 1 Then
            ActiveCell.FormulaR1C1 = Str(Sheets("bLPNCorona").Cells(i, 1).Value)
            i = i + 1
        Else
            ActiveCell.FormulaR1C1 = Str(ActiveCell.Offset(-1, 0).Value)
        End If
        ActiveCell.Offset(1, 0).Select
    Wend
    Application.ScreenUpdating = True
End Sub

Sub fillUp()
    Application.ScreenUpdating = False
    Sheets("PreDist").Select
    Range("A2", Range("B1").End(xlDown)).Copy
    Sheets("Subir").Select
    Range("A2").PasteSpecial xlPasteValues
    Sheets("PreDist").Select
    If Range("J3").Value = "" Then
        Range("J2").Copy
    Else
        Range("J2", Range("J2").End(xlDown)).Copy
    End If
    Sheets("Subir").Select
    Range("C2").PasteSpecial xlPasteValues
    Columns("C:C").NumberFormat = "0"
    Columns("C:C").AutoFit
    Sheets("PreDist").Select
    Range("C2", Range("C1").End(xlDown)).Copy
    Sheets("Subir").Select
    Range("D2").PasteSpecial xlPasteValues
    Sheets("PreDist").Select
    Range("F2", Range("F1").End(xlDown)).Copy
    Sheets("Subir").Select
    Range("E2").PasteSpecial xlPasteValues
    Range("F2").FormulaR1C1 = "FE"
    Range("G2").FormulaR1C1 = InputBox("Ingrese el numero de la factura: ")
    Range("H2").FormulaR1C1 = InputBox("Ingrese la fecha de la factura: ")
    If Range("E3").Value <> "" Then
        Range("F2:H2").Copy
        Range("F2", Range("E1").End(xlDown).Offset(0, 3)).PasteSpecial xlPasteValues
    End If
    Range("I1").FormulaR1C1 = Range("B1").Value
    Range("I2").FormulaR1C1 = "=VLOOKUP(RC[-7],Maestras!C[2]:C[3],2,0)"
    If Range("A3").Value <> "" Then
        Range("I2").AutoFill Destination:=Range("I2", Range("H1").End(xlDown).Offset(0, 1))
    End If
    Columns("I:I").Copy
    Range("B1").PasteSpecial xlPasteValues
    Columns("I:I").ClearContents
    Range("I1").FormulaR1C1 = Range("D1").Value
    Range("I2").FormulaR1C1 = "=VLOOKUP(RC[-5],Maestras!C[6]:C[7],2,0)"
    If Range("A3").Value <> "" Then
        Range("I2").AutoFill Destination:=Range("I2", Range("H1").End(xlDown).Offset(0, 1))
    End If
    Columns("I:I").Copy
    Range("D1").PasteSpecial xlPasteValues
    Columns("I:I").ClearContents
    Application.ScreenUpdating = True
End Sub

Sub Macro2()
    Sheets("Rotulo").Select
    Columns("A:Z").Copy
    Workbooks.Add
    Range("A1").PasteSpecial Paste:=xlPasteValues
    ActiveWorkbook.SaveAs Filename:=ThisWorkbook.Path & "\bCorona\eCorona.xls" _
        , FileFormat:=xlExcel8, Password:="", WriteResPassword:="", _
        ReadOnlyRecommended:=False, CreateBackup:=False
    ActiveWindow.Close
    Sheets("Menu").Select
End Sub
