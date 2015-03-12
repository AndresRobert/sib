'botones

Sub botonProcesarTabla()
    Application.ScreenUpdating = False
    If MsgBox("Primero debe copiar en la hoja archivob2b la planilla descargada " & _
        "desde el sitio de proveedores de Ripley. ¿Esta seguro de que ha realizado este proceso?", _
        vbYesNo, "Antes de continuar...") = vbYes Then
        Call copiarTabla
        Call separarTablaOrd
        Call eliminarColumnasOrd
        Call nombrarColumnas
        If MsgBox("¿Desea modificar las cantidades informadas por la orden de compra?", vbYesNo) = vbYes Then
            Sheets("ord").Select
        Else
            Call botonGenerarDistribucion
        End If
    Else
        MsgBox ("Ingrese al sitio de proveedores de Ripley y descargue el archivo correspondiente " & _
            "a la orden compra, luego copie y pegue la primera columna del archivo Ripley en la celda A1 " & _
            "de la hoja archivob2b de esta planilla.")
    End If
    Application.ScreenUpdating = True
End Sub

Sub botonGenerarDistribucion()
    Application.ScreenUpdating = False
    Call copiarColumnasDis
    Call codigoProveedorDis
    Call insertarTitulos
    Call formatoTabla
    Call botonImprimirDistribucion
    Call guardarDocumento
    Application.ScreenUpdating = True
End Sub

Sub botonImprimirDistribucion()
    Sheets("dis").Select
    Range("A1").Select
    If MsgBox("Desea imprimir la distribución?", vbYesNo) = vbYes Then
        ActiveWindow.SelectedSheets.PrintOut
    End If
    Sheets("menu").Select
End Sub

Sub botonSeleccionarBultos()
    Application.ScreenUpdating = False
    Sheets("dis").Select
    Columns(8).ClearContents
    Range("H4").FormulaR1C1 = "=IF(RC[-7]=R[-1]C[-7],"""",1)"
    If Range("G5").Value <> "" Then
        Range("H4").AutoFill Destination:=Range("H4", Range("G4").End(xlDown).Offset(0, 1))
    End If
    Columns(8).Copy
    Range("H1").PasteSpecial Paste:=xlPasteValues
    Range("H4").Select
    Application.CutCopyMode = False
    MsgBox ("Agregue un 1 al inicio de cada bulto que no haya sido reconocido automáticamente. " & _
            "Modifique las cantidades si corresponde. " & _
            "Agregue o elimine filas hasta hacer coincidir esta planilla con el documento físico.")
    Application.ScreenUpdating = True
End Sub

Sub botonGenerarEtiqueta()
    Application.ScreenUpdating = False
    If MsgBox("Está facturada la distribución?", vbYesNo, "Antes de continuar...") = vbYes Then
        i = 1
        Call probarFactura
        While (i < 3) And (Sheets("afc").Range("A2").Value = "")
            Call probarFactura
            i = i + 1
        Wend
        If i <> 3 Then
            Call modificarMaestra
            Call bultosSeleccionados
            Call leerFolios
            Call generarFolios
            Call devolverFolios
            Call crearArchivoFolios
            Call baseArchivoFactura
            Call baseArchivoBulto
            Call baseEtiqueta
        Else
            MsgBox ("Solicite la facturación del pedido e intentelo nuevamente.")
        End If
    Else
        MsgBox ("Solicite la facturación del pedido e intentelo nuevamente.")
    End If
    Application.ScreenUpdating = True
End Sub

Sub botonCrearArchivos()
    Application.ScreenUpdating = False
    Call modificarMaestra
    Call baseArchivoFactura
    Call baseArchivoBulto
    Call crearArchivoDeBulto
    Call crearArchivoDeFactura
    Application.ScreenUpdating = True
End Sub

Sub botonCrearArchivoEtiqueta()
    Application.ScreenUpdating = False
    Sheets("etq").Select
    If Dir(ThisWorkbook.Path & "\bRipley\eRipley.xls", vbNormal) <> "" Then
        Kill ThisWorkbook.Path & "\bRipley\eRipley.xls"
    End If
    Columns("A:K").Copy
    Workbooks.Add
    Range("A1").PasteSpecial Paste:=xlPasteValues
    ActiveWorkbook.SaveAs FileName:=ThisWorkbook.Path & "\bRipley\eRipley.xls", FileFormat:=xlExcel8
    ActiveWindow.Close
    Sheets("menu").Select
    MsgBox ("Listo para imprimir los rotulos.")
    Application.ScreenUpdating = True
End Sub

'procesos

Sub copiarTabla()
    Sheets("ord").Select
    Columns("A:AA").ClearContents
    Columns("A:AA").ClearFormats
    Sheets("archivob2b").Select
    Columns(1).Copy
    Sheets("ord").Select
    Range("A1").PasteSpecial Paste:=xlPasteValues
End Sub

Sub separarTablaOrd()
    Sheets("ord").Select
    Columns(1).TextToColumns _
        Destination:=Range("A1"), _
        DataType:=xlDelimited, _
        Comma:=True, _
        TrailingMinusNumbers:=True, _
        ConsecutiveDelimiter:=False, _
        Tab:=False, _
        Semicolon:=False, _
        Space:=False, _
        Other:=False
End Sub

Sub eliminarColumnasOrd()
    Sheets("ord").Select
    Columns("AD:AK").Delete
    Columns("V:AB").Delete
    Columns("T:T").Delete
    Columns("J:Q").Delete
    Columns("B:G").Delete
End Sub

Sub nombrarColumnas()
    Sheets("ord").Select
    Range("A1").FormulaR1C1 = "OCOMP"
    Range("B1").FormulaR1C1 = "CODEPTO"
    Range("C1").FormulaR1C1 = "DEPTO"
    Range("D1").FormulaR1C1 = "NROLOC"
    Range("E1").FormulaR1C1 = "LOCAL"
    Range("F1").FormulaR1C1 = "SKU"
    Range("G1").FormulaR1C1 = "CANT"
    Columns("A:G").AutoFit
End Sub

Sub copiarColumnasDis()
    Sheets("dis").Select
    Columns("A:AA").ClearContents
    Columns("A:AA").ClearFormats
    Sheets("ord").Select
    Columns("D:G").Copy
    Sheets("dis").Select
    Range("A1").PasteSpecial Paste:=xlPasteValues
    With ActiveWorkbook.Worksheets("dis").Sort
        .SortFields.Clear
        .SortFields.Add Key:=Columns(1), SortOn:=xlSortOnValues, Order:=xlAscending
        .SortFields.Add Key:=Columns(3), SortOn:=xlSortOnValues, Order:=xlAscending
        .SetRange Columns("A:D")
        .Header = xlYes
        .Orientation = xlTopToBottom
        .Apply
    End With
End Sub

Sub codigoProveedorDis()
    Sheets("dis").Select
    Columns("D:F").Insert
    Range("D1").FormulaR1C1 = "ITEM"
    Range("E1").FormulaR1C1 = "CODPROV"
    Range("F1").FormulaR1C1 = "UM"
    Range("D2").FormulaR1C1 = "=IF(RC[-3]<>R[-1]C[-3],1,R[-1]C+1)"
    Range("E2").FormulaR1C1 = "=VLOOKUP(RC[-2],mae!C[-4]:C[-1],2,0)"
    Range("F2").FormulaR1C1 = "=VLOOKUP(RC[-3],mae!C[-5]:C[-3],3,0)"
    Range("D2:F2").AutoFill Destination:=Range("D2", Range("C1").End(xlDown).Offset(0, 3))
    Columns("D:F").Copy
    Columns("D:F").PasteSpecial Paste:=xlPasteValues
End Sub

Sub insertarTitulos()
    Sheets("dis").Select
    notaVenta = InputBox("Ingrese la nota de venta (Nro. de pedido):", "Sistema Integrado B2B", "11111")
    Rows("1:2").Insert
    Range("A1").FormulaR1C1 = "DISTRIBUCION RIPLEY"
    Range("E1").FormulaR1C1 = "NOTA DE VENTA"
    Range("E2").FormulaR1C1 = "ORDEN DE COMPRA"
    Range("F2").FormulaR1C1 = Sheets("ord").Range("A2").Value
    Range("A2").FormulaR1C1 = Sheets("ord").Range("C2").Value
    Range("F1").FormulaR1C1 = notaVenta
End Sub

Sub formatoTabla()
    Sheets("dis").Select
    Columns("A:AA").ClearFormats
    Range("A1:C1").Merge
    Range("D2:E2").Merge
    Range("D1:E1").Merge
    Range("A2:C2").Merge
    Range("F1:G1").Merge
    Range("F2:G2").Merge
    Range("A3:G3").Interior.ThemeColor = xlThemeColorLight1
    Range("A3:G3").Font.ThemeColor = xlThemeColorDark1
    Range("D1:E2").Interior.ThemeColor = xlThemeColorLight1
    Range("D1:E2").Font.ThemeColor = xlThemeColorDark1
    Range("A1:D2").Borders(xlEdgeLeft).Weight = xlThin
    Range("A1:D2").Borders(xlEdgeTop).Weight = xlThin
    Range("A1:D2").Borders(xlEdgeBottom).Weight = xlThin
    Range("A1:D2").Borders(xlEdgeRight).Weight = xlThin
    Range("A1:D2").Borders(xlInsideHorizontal).Weight = xlThin
    Range("F1:G2").Borders(xlEdgeLeft).Weight = xlThin
    Range("F1:G2").Borders(xlEdgeTop).Weight = xlThin
    Range("F1:G2").Borders(xlEdgeBottom).Weight = xlThin
    Range("F1:G2").Borders(xlEdgeRight).Weight = xlThin
    Range("F1:G2").Borders(xlInsideHorizontal).Weight = xlThin
    Range("A4", Range("G3").End(xlDown)).Borders(xlEdgeLeft).Weight = xlHairline
    Range("A4", Range("G3").End(xlDown)).Borders(xlEdgeRight).Weight = xlHairline
    Range("A4", Range("G3").End(xlDown)).Borders(xlInsideVertical).Weight = xlHairline
    Range("A4").Select
    While ActiveCell <> ""
        If ActiveCell.Offset(1, 0).Value <> ActiveCell.Value Then
            Range(ActiveCell, ActiveCell.Offset(0, 6)).Borders(xlEdgeBottom).Weight = xlThick
        Else
            Range(ActiveCell, ActiveCell.Offset(0, 6)).Borders(xlEdgeBottom).Weight = xlHairline
        End If
        ActiveCell.Offset(1, 0).Select
    Wend
    Range("A1").Select
    Columns("A:G").EntireColumn.AutoFit
End Sub

Sub bultosSeleccionados()
    Sheets("dis").Select
    Range("I4").FormulaR1C1 = "=IF(RC[-1]=1,IF(RC[-8]=R[-1]C[-8],R[-1]C+1,1),R[-1]C)"
    If Range("G5").Value <> "" Then
        Range("I4").AutoFill Destination:=Range("I4", Range("G4").End(xlDown).Offset(0, 2))
    End If
    Columns(9).Copy
    Range("I1").PasteSpecial Paste:=xlPasteValues
End Sub

Sub guardarDocumento()
    notaVenta = Sheets("dis").Range("F1").Value
    Open ThisWorkbook.Path & "\" & notaVenta & ".bat" For Output As #1
        Print #1, "start " & notaVenta & ".xlsm"
    Close #1
    ActiveWorkbook.SaveAs FileName:=ThisWorkbook.Path & "\" & notaVenta & ".xlsm"
End Sub

Sub leerFolios()
    Sheets("mae").Select
    Open ThisWorkbook.Path & "\bfoliosr.txt" For Input As #1
    Range("H2").Select
    i = 1
    Lineas = Split(Input(LOF(1), #1), Chr(13))
    While i <= UBound(Lineas)
        ActiveCell.Value = "=TEXT(VALUE(" & Lineas(i - 1) & "),""000000000"")"
        ActiveCell.Offset(1, 0).Select
        i = i + 1
    Wend
    Close #1
    Columns("H:H").Copy
    Columns("H:H").PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
End Sub

Sub generarFolios()
    Sheets("abl").Select
    Columns("A:AA").ClearContents
    Sheets("dis").Select
    Columns("A:I").Copy
    Sheets("abl").Select
    Range("A1").PasteSpecial Paste:=xlPasteValues
    Rows("1:2").Delete
    Columns(8).Delete
    Columns(4).Delete
    Range("H2").FormulaR1C1 = "=VLOOKUP(RC[-7],mae!C[-2]:C,3,0)+RC[-1]"
    If Range("G3").Value <> "" Then
        Range("H2").AutoFill Destination:=Range("H2", Range("G2").End(xlDown).Offset(0, 1))
    End If
    Columns(8).Copy
    Range("H1").PasteSpecial Paste:=xlPasteValues
    Columns(7).Delete
    Range("H2").FormulaR1C1 = "=CONCATENATE(""5055"",RC[-7],TEXT(RC[-1],""000000000""))"
    If Range("G3").Value <> "" Then
        Range("H2").AutoFill Destination:=Range("H2", Range("G2").End(xlDown).Offset(0, 1))
    End If
    Columns(8).Copy
    Range("H1").PasteSpecial Paste:=xlPasteValues
    Range("G1").FormulaR1C1 = "NVOFOLIO"
    Range("H1").FormulaR1C1 = "FOLIOBTO"
    Columns("A:H").AutoFit
End Sub

Sub devolverFolios()
    Sheets("abl").Select
    Columns(1).Copy
    Sheets("mae").Select
    Range("J1").PasteSpecial Paste:=xlPasteValues
    Sheets("abl").Select
    Columns(7).Copy
    Sheets("mae").Select
    Range("K1").PasteSpecial Paste:=xlPasteValues
    Columns("J:K").Select
    With ActiveWorkbook.Worksheets("mae").Sort
        .SortFields.Clear
        .SortFields.Add _
            Key:=Columns(11), _
            SortOn:=xlSortOnValues, _
            Order:=xlDescending, _
            DataOption:=xlSortNormal
        .SetRange Columns("J:K")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Selection.RemoveDuplicates Columns:=1, Header:=xlYes
    Columns(8).Copy
    Range("L1").PasteSpecial Paste:=xlPasteValues
    Range("H2").FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-2],C[2]:C[3],2,0)+1,RC[4])"
    Range("H2").AutoFill Destination:=Range("H2", Range("H2").End(xlDown))
    Columns(8).Copy
    Range("H1").PasteSpecial Paste:=xlPasteValues
End Sub

Sub crearArchivoFolios()
    Sheets("mae").Select
    If Dir(ThisWorkbook.Path & "\bfoliosr.txt", vbNormal) <> "" Then
        Kill ThisWorkbook.Path & "\bfoliosr.txt"
    End If
    Open ThisWorkbook.Path & "\bfoliosr.txt" For Output As #1
    Range("H2").Select
    While ActiveCell.Value <> ""
        Print #1, ActiveCell.Value
        ActiveCell.Offset(1, 0).Select
    Wend
    Close #1
    Sheets("menu").Select
End Sub

Sub baseArchivoFactura()
    Sheets("afc").Select
    Columns(17).ClearContents
    Range("Q1").FormulaR1C1 = "RUT90914000-5"
    Range("Q2").FormulaR1C1 = "DOCEFAC"
    Range("Q3").FormulaR1C1 = "=CONCATENATE(""NUM"",TEXT(R[-1]C[-7],""0000000000""))"
    Range("Q4").FormulaR1C1 = "=CONCATENATE(""ODI"",TEXT(ord!R[-2]C[-16],""00000000""))"
    Range("Q5").FormulaR1C1 = "=CONCATENATE(""FEC"",TEXT(R[-3]C[-16],""ddmmyyyy""))"
    Range("Q6").FormulaR1C1 = "=CONCATENATE(""NET"",TEXT(R[-4]C[-15],""0000000000""))"
    Range("Q7").FormulaR1C1 = "=CONCATENATE(""IVA"",TEXT(R[-5]C[-14],""0000000000""))"
    Range("Q8").FormulaR1C1 = "=CONCATENATE(""TOT"",TEXT(R[-6]C[-13],""0000000000""))"
    Escribir = Range("Q9").AddressLocal
    Leer = Range("K2").AddressLocal
    While Range(Leer).Value <> ""
        Range(Escribir).FormulaR1C1 = "ARP" & Range(Leer).Value
        Escribir = Range(Escribir).Offset(1, 0).AddressLocal
        Range(Escribir).FormulaR1C1 = "ARR" & Range(Leer).Offset(0, 1).Value
        Escribir = Range(Escribir).Offset(1, 0).AddressLocal
        Range(Escribir).FormulaR1C1 = "CAN" & Range(Leer).Offset(0, 2).Value
        Escribir = Range(Escribir).Offset(1, 0).AddressLocal
        Range(Escribir).FormulaR1C1 = "PRU" & Range(Leer).Offset(0, 3).Value
        Escribir = Range(Escribir).Offset(1, 0).AddressLocal
        Range(Escribir).FormulaR1C1 = "PRT" & Range(Leer).Offset(0, 4).Value
        Escribir = Range(Escribir).Offset(1, 0).AddressLocal
        Range(Escribir).FormulaR1C1 = "FIA"
        Escribir = Range(Escribir).Offset(1, 0).AddressLocal
        Leer = Range(Leer).Offset(1, 0).AddressLocal
    Wend
    Range(Escribir).FormulaR1C1 = "FID"
    Escribir = Range(Escribir).Offset(1, 0).AddressLocal
    Range(Escribir).FormulaR1C1 = "FIT"
    Columns(17).Copy
    Range("Q1").PasteSpecial Paste:=xlPasteValues
End Sub

Sub baseArchivoBulto()
    Sheets("abl").Select
    Range("I1").FormulaR1C1 = "Factura"
    Range("J1").FormulaR1C1 = "oCompra"
    Range("K1").FormulaR1C1 = "FECHA"
    Range("L1").FormulaR1C1 = "NVenta"
    Range("M1").FormulaR1C1 = "nDepto"
    Range("I2").FormulaR1C1 = Sheets("afc").Range("J2").Value
    Range("J2").FormulaR1C1 = Sheets("ord").Range("A2").Value
    Range("K2").FormulaR1C1 = "=TEXT(TODAY(),""dd-mm-yyyy"")"
    Range("K2").Copy
    Range("K2").PasteSpecial Paste:=xlPasteValues
    Range("L2").FormulaR1C1 = Sheets("dis").Range("F1").Value
    Range("M2").FormulaR1C1 = Sheets("ord").Range("B2").Value
    Range("N1").FormulaR1C1 = "Depto"
    Range("N2").FormulaR1C1 = Sheets("ord").Range("C2").Value
    Range("I2:N2").Copy
    Range("I2", Range("H1").End(xlDown).Offset(0, 6)).PasteSpecial Paste:=xlPasteValues
    Range("H1").FormulaR1C1 = "Folio2"
    Range("B1").FormulaR1C1 = "Nombre Local"
    Columns("A:M").AutoFit
End Sub

Sub baseEtiqueta()
    Sheets("etq").Select
    Columns("A:AA").ClearContents
    Sheets("abl").Select
    Columns("A:N").Copy
    Sheets("etq").Select
    Range("A1").PasteSpecial Paste:=xlPasteValues
    Columns(11).Delete
    Columns("C:F").Delete
    Columns(1).Delete
    Columns("A:H").AutoFit
    Columns("A:H").RemoveDuplicates Columns:=3, Header:=xlYes
    Range("I1").FormulaR1C1 = "NBulto"
    Range("J1").FormulaR1C1 = "TotBultos"
    Range("K1").FormulaR1C1 = "Peso"
    Range("I2").FormulaR1C1 = "=IF(RC[-8]<>R[-1]C[-8],1,R[-1]C+1)"
    Range("J2").FormulaR1C1 = "=COUNTIF(C[-9],RC[-9])"
    If Range("H3").Value <> "" Then
        Range("I2:J2").AutoFill Destination:=Range("I2", Range("H2").End(xlDown).Offset(0, 2))
    End If
    Columns("K:K").NumberFormat = "0.00 ""Kg"""
    MsgBox ("Complete la tabla agregando el peso de los bultos.")
    Range("K2").Select
End Sub

Sub probarFactura()
    Sheets("afc").Select
    Range("A1").Select
    nroFactura = InputBox("Ingrese un número válido de factura.")
    With ActiveWorkbook.Connections("iFactura").ODBCConnection
        .CommandText = Array("SELECT VenT_DoctoLegalCar.DlcFecDocto AS 'Fecha', VenT_DoctoLegalCar.DlcFolioDocto, ", _
        "VenT_DoctoLegalCar.DlcMtoNeto AS 'Neto', VenT_DoctoLegalCar.DlcMtoIva AS 'Iva', VenT_DoctoLegalCar.DlcMtoTotal AS 'Total', ", _
        "VenT_DoctoLegalDet.DldItem AS 'Linea', VenT_DoctoLegalDet.CodigoArticulo AS 'ATS', VenT_DoctoLegalDet.DldCantDoctoFac AS 'Cant', ", _
        "VenT_DoctoLegalDet.DvdPrecioUnitario AS 'VUni', VenT_DoctoLegalDet.DldValorFinal AS 'VTot'" & Chr(13) & "" & Chr(10) & _
        "FROM fin700v60.dbo.ExiT_Articulos ExiT_Articulos, fin700v60.dbo.VenT_DoctoLegalCar VenT_DoctoLegalCar, ", _
        "fin700v60.dbo.VenT_DoctoLegalDet VenT_DoctoLegalDet" & Chr(13) & "" & Chr(10) & _
        "WHERE VenT_DoctoLegalCar.EmpId = ExiT_Articulos.EmpId AND VenT_DoctoLegalDet.CodigoArticulo = ExiT_Articulos.CodigoArticulo AND ", _
        "VenT_DoctoLegalDet.EmpId = ExiT_Articulos.EmpId AND VenT_DoctoLegalDet.DivCodigo = VenT_DoctoLegalCar.DivCodigo AND ", _
        "VenT_DoctoLegalDet.DlcNumDocto = VenT_DoctoLegalCar.DlcNumDocto AND VenT_DoctoLegalDet.EmpId = VenT_DoctoLegalCar.EmpId AND ", _
        "VenT_DoctoLegalDet.PerId = VenT_DoctoLegalCar.PerId AND VenT_DoctoLegalDet.UniCodigo = VenT_DoctoLegalCar.UniCodigo AND ", _
        "((ExiT_Articulos.EmpId=1) AND (VenT_DoctoLegalCar.CliRut='0083382700-6') AND (VenT_DoctoLegalCar.TdoId=33) AND ", _
        "(VenT_DoctoLegalCar.DlcFolioDocto=" & nroFactura & "))")
        .CommandType = xlCmdSql
        .Connection = Array("ODBC;DSN=fin700;Description=base;UID=CONSULTA_moletto;PWD=123edc;APP=Microsoft Office 2010;", _
        "WSID=BROTULADO;DATABASE=Fin700V60;LANGUAGE=Español;Network=DBMSSOCN;Address=100.1.20.10,1433")
        .ServerCredentialsMethod = xlCredentialsMethodIntegrated
    End With
    Selection.ListObject.QueryTable.Refresh
End Sub

Sub crearArchivoDeBulto()
    Sheets("abl").Select
    factura = Sheets("afc").Range("J2").Value
    oCompra = Sheets("abl").Range("J2").Value
    If Dir(ThisWorkbook.Path & "\bRipley\Bulto de Factura Nro " & factura & ".txt", vbNormal) <> "" Then
        Kill ThisWorkbook.Path & "\bRipley\Bulto de Factura Nro " & factura & ".txt"
    End If
    Open ThisWorkbook.Path & "\bRipley\Bulto de Factura Nro " & factura & ".txt" For Output As #1
        Print #1, "<?xml version=""1.0""  encoding=""ISO-8859-1"" standalone=""no""?>"
        Print #1, "<!DOCTYPE DOC_BULTOS_DEM_ODI SYSTEM  ""dtd/odi_asignada.dtd"">"
        Print #1, "<DOC_BULTOS_DEM_ODI>"
        Print #1, "<RUT>90914000-5</RUT>"
        Print #1, "<FACTURA_GUIA>"
        Print #1, "<NUM_FAC_GUI>" & factura & "</NUM_FAC_GUI>"
        Print #1, "<FACTURA_O_GUIA>FACTURA_ELECTRONICA</FACTURA_O_GUIA>"
        Print #1, "<NUM_ODI>" & oCompra & "</NUM_ODI>"
    Range("H2").Select
    BultoActivo = Range("H2").Value
    While ActiveCell.Value <> ""
        Print #1, "<BULTO_DEM_ODI>"
        Print #1, "<NUM_BULTO>" & ActiveCell.Value & "</NUM_BULTO>"
        Print #1, "<FECHA_BULTO>" & ActiveCell.Offset(0, 3).Value & "</FECHA_BULTO>"
        While ActiveCell.Value = BultoActivo
            Print #1, "<DETALLE_BULTO_DEM_ODI>"
            Print #1, "<NUM_PROD_RIPLEY>" & ActiveCell.Offset(0, -5).Value & "</NUM_PROD_RIPLEY>"
            Print #1, "<PROV_CASEPACK>" & ActiveCell.Offset(0, -4).Value & "</PROV_CASEPACK>"
            Print #1, "<CANTIDAD>" & ActiveCell.Offset(0, -2).Value & "</CANTIDAD>"
            Print #1, "</DETALLE_BULTO_DEM_ODI>"
            ActiveCell.Offset(1, 0).Select
        Wend
        Print #1, "</BULTO_DEM_ODI>"
        BultoActivo = ActiveCell.Value
    Wend
    Print #1, "</FACTURA_GUIA>"
    Print #1, "</DOC_BULTOS_DEM_ODI>"
    Close #1
    MsgBox ("Archivo de Bulto creado correctamente.")
End Sub

Sub crearArchivoDeFactura()
    Sheets("afc").Select
    factura = Sheets("afc").Range("J2").Value
    If Dir(ThisWorkbook.Path & "\bRipley\Factura " & factura & ".txt", vbNormal) <> "" Then
        Kill ThisWorkbook.Path & "\bRipley\Factura " & factura & ".txt"
    End If
    Open ThisWorkbook.Path & "\bRipley\Factura " & factura & ".txt" For Output As #1
        Range("Q1").Select
        While ActiveCell.Value <> ""
            Print #1, ActiveCell.Value
            ActiveCell.Offset(1, 0).Select
        Wend
    Close #1
    Sheets("menu").Select
    MsgBox ("Archivo creado de Factura correctamente.")
End Sub

Sub modificarMaestra()
    Sheets("dis").Select
    If Range("A5").Value <> "" Then
        Range("C4", Range("C4").End(xlDown)).Copy
    Else
        Range("C4").Copy
    End If
    Sheets("mae").Select
    Range("A2").PasteSpecial Paste:=xlPasteValues
    Range("D2").PasteSpecial Paste:=xlPasteValues
    Sheets("dis").Select
    If Range("A5").Value <> "" Then
        Range("E4", Range("F4").End(xlDown)).Copy
    Else
        Range("E4:F4").Copy
    End If
    Sheets("mae").Select
    Range("B2").PasteSpecial Paste:=xlPasteValues
End Sub

'extras

Sub separarEspecial()
    Range("H2").Select
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
