Attribute VB_Name = "ModTMP"
Option Explicit

'MODULO PARA LA CARGA Y DESCARGA DE TABLAS TEMPORALES


'================================================================================
'================================================================================

'================================================================================
'TMPnseries: Temporal para introducir los Nº de Serie de los Articulos en compras o en ventas
'USO: frmFacEntAlbaran, frmRepEntAlbaran
'================================================================================

Public Function CargarDatosTMPNumSeries(NomTabla As String, codArtic As String, cant As Integer, NumLinAlb As String) As Boolean
'IN -> NomTabla: Nombre de la tabla temporal
'      CodArtic: Codigo Articulo del que se van a Introducir los Nº de Serie
'      Cant: cantidad de Articulo (tantas filas como articulos)
'      Mostrar: si true se cargar los Nº de serie sino en blanco
Dim SQL As String
Dim i As Integer
Dim numlinea As String, vWhere As String

    On Error GoTo ECargaDatosTMP

    'Insertar tantos registros como cantidad de Articulo Introducida
    vWhere = "(codusu=" & vUsu.Codigo & " and codartic=" & DBSet(codArtic, "T") & " and numlinealb=" & DBSet(NumLinAlb, "N") & ")"

    'insertamos tantos num.serie como cantidad
    For i = 0 To cant - 1
        'Obtener Num Linea
        numlinea = SugerirCodigoSiguienteStr(NomTabla, "numlinea", vWhere)
        'Insertar en la temporal para Nº Series
        SQL = "INSERT INTO " & NomTabla & " (codusu, codartic, numlinealb, numlinea, numserie) VALUES ("
        SQL = SQL & vUsu.Codigo & ", " & DBSet(codArtic, "T") & ", " & NumLinAlb & ", " & numlinea & ", ' ')"
        Conn.Execute SQL
    Next i

ECargaDatosTMP:
    If Err.Number <> 0 Then
        CargarDatosTMPNumSeries = False
        MuestraError Err.Number, "Numeros Serie", Err.Description
    Else
        CargarDatosTMPNumSeries = True
    End If
End Function


Public Function DescargarDatosTMPNumSeries(NomTabla As String)
'Al salir de la aplicacion se borran los datos de la tabla temporal
Dim SQL As String

    On Error GoTo EDescargaDatos

     '------------- AHORA
    SQL = "DELETE from " & NomTabla & " where codusu= " & vUsu.Codigo
    Conn.Execute SQL
    
    Exit Function
EDescargaDatos:
        MuestraError Err.Number, "Descargar Tabla Temporal (Nº Serie).", Err.Description
End Function



Public Function InsertarNSeries(codArtic As String, CadValuesI As String, cadValuesU As String, DeVenta As Boolean) As Boolean
'Insertar un registro en la tabla "sserie" por cada uno de los
'Nº de Serie introducidos en la Tabla Temporal
Dim RS As ADODB.Recordset
Dim SQL As String, devuelve As String
Dim codTipar As String, numalbar As String
    
    On Error GoTo EInsertar

    'Seleccionar los nº de serie cargados en la temporal: tmpnseries
    SQL = "SELECT * FROM tmpnseries WHERE codusu=" & vUsu.Codigo & " AND codartic=" & DBSet(codArtic, "T")
    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Not RS.EOF Then RS.MoveFirst
    
    While Not RS.EOF
        'Comprobar si existe en la tabla sserie
        If DeVenta Then
            numalbar = "numalbar" 'Nº albaran de Venta
        Else
            numalbar = "numalbpr" 'Nº albaran de Compras
        End If
        devuelve = DevuelveDesdeBDNew(cAgro, "sserie", "numserie", "numserie", RS!numserie, "T", numalbar, "codartic", RS!codArtic, "T")
        If devuelve <> "" Then 'Existe en tabla sserie
            If numalbar = "" Then
                SQL = "UPDATE sserie SET " & cadValuesU
                '=== Laura 17/01/2007
                SQL = SQL & " WHERE numserie=" & DBSet(RS!numserie, "T") & " AND codartic=" & DBSet(RS!codArtic, "T")
                '===
            End If
        Else
            'Obtener el tipo de Articulo
            codTipar = DevuelveDesdeBDNew(cAgro, "sartic", "codtipar", "codartic", RS!codArtic, "T")
        
            'Insertar en la tabla sserie
            SQL = "INSERT INTO sserie (numserie, codartic, codtipar, codclien, coddirec,tieneman, nummante, ultrepar, fingaran, "
            SQL = SQL & " codtipom, numfactu, fechavta, numalbar, numline1, codprove, numalbpr, fechacom, numline2) "
            SQL = SQL & " VALUES ( " & DBSet(RS!numserie, "T") & ", " & DBSet(RS!codArtic, "T") & ", " & DBSet(codTipar, "T") & ","
            SQL = SQL & CadValuesI
            SQL = SQL & ") "
        End If
        Conn.Execute SQL
        RS.MoveNext
    Wend
    RS.Close
    Set RS = Nothing
EInsertar:
    If Err.Number <> 0 Then MuestraError Err.Number, "Insertando Nº Serie", Err.Description
End Function





Public Sub PedirNSeriesGnral(ByRef RS As ADODB.Recordset, Men As Boolean)
Dim SQL As String
Dim b As Boolean

    On Error GoTo EPedirNSeries

        If Men Then
            SQL = "Hay artículos que tienen control de Nº de Serie." & vbCrLf & vbCrLf
            SQL = SQL & "Introduzca los Nº De Serie." & vbCrLf
            MsgBox SQL, vbInformation
        End If
        
        'Cargar la tabla temporal con tantas filas como cantidad de Articulo
        'Para introducir el Nº de Serie
        DescargarDatosTMPNumSeries ("tmpnseries")
        b = True
        
        While Not RS.EOF
            If Not CargarDatosTMPNumSeries("tmpnseries", RS!codArtic, RS!Cantidad, RS!numlinea) Then
                b = False
            End If
            RS.MoveNext
        Wend
        
        'Visualizar en pantalla el Grid, y rellenar los Nº Serie
        If Not b Then MsgBox "No se han podido mostrar todos los Artículos con Nº de Serie.", vbInformation
        
EPedirNSeries:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub



Public Function MostrarNSeriesGnral(ByRef RSLineas As ADODB.Recordset, vCampos As String, Optional Rectifica As Boolean) As String
'Si los Nº de serie se introdujeron en ALBARAN COMPRAS se muestran
'los Nº de serie de los articulos comprados y se seleccionan tantos como cantidad de la linea
'IN -> RSLineas: lineas del Albaran generado
'OUT -> vCampos: concatena la cantidad requerida de Nº series de cada articulo
'RETURN -> cadena SQL con la Select que se pasara para mostrar los Nº series
Dim SQL As String
Dim cadArtic As String
Dim Campos As String
Dim totArtic As Integer

    On Error GoTo EMostrar

    'Concatenamos los codigos de Articulo que tenemos que seleccionar se "sseries"
    cadArtic = ""
    totArtic = 0
    Campos = ""
    While Not RSLineas.EOF
        Campos = Campos & RSLineas!codArtic & "|" & RSLineas!Cantidad & "·"
        If cadArtic = "" Then
            cadArtic = DBSet(RSLineas!codArtic, "T")
        Else
            cadArtic = cadArtic & ", " & DBSet(RSLineas!codArtic, "T")
        End If
        totArtic = totArtic + 1
        RSLineas.MoveNext
    Wend
    RSLineas.MoveFirst
    cadArtic = "(" & cadArtic & ")"
    vCampos = Campos
   
    'Se puede seleccionar todos los Nº de Serie que se necesitan
     'se introdujo los Nº de Serie en COMPRAS y ahora
    'mostramos los Nº de Serie para seleccionar cual vamos a
    'vender al Cliente
    Screen.MousePointer = vbDefault
    If Rectifica Then
        'viene de una factura rectificativa los nº de serie que seleccionemos sera para quitar
        SQL = "Hay Artículos que tienen control de Nº de Serie." & vbCrLf
        SQL = SQL & "Seleccione los nº de Serie que desea rectificar."
    Else
        If totArtic > 1 Then
            SQL = "Hay Artículos que tienen control de Nº de Serie." & vbCrLf
            SQL = SQL & vbCrLf & "Seleccione un Nº de Serie para cada Artículo."
        Else
            SQL = "El Artículo tienen control de Nº de Serie." & vbCrLf
            SQL = SQL & vbCrLf & "Seleccione los Nº de Serie para el Artículo."
        End If
    End If
    MsgBox SQL, vbInformation
    SQL = " WHERE sserie.codartic IN " & cadArtic
    MostrarNSeriesGnral = SQL
    
EMostrar:
    If Err.Number <> 0 Then MuestraError Err.Number, "Mostrar Nº series", Err.Description
End Function


'================================================================================
'TMPStockFec: Temporal para obtener el Stock que habia en una determinada fecha
'(Para Listado de Almacenes: "Inf. Stock a una Fecha")
'USO: frmListados
'================================================================================

Public Function CargarTMPStockFecha(vSQL As String, cadFecha As String, cadHora As String) As Boolean
'Carga la tabla temporal con el Stock del almacen seleccionado
'de los articulos seleccionados que habia a una determinada FECHA, HORA
Dim RS As ADODB.Recordset
Dim vStock As Single
Dim cadSQL As String

    On Error GoTo ECargarTMPStock

    CargarTMPStockFecha = False
    
    Set RS = New ADODB.Recordset
    RS.Open vSQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not RS.EOF
        'Para cada articulo obtener el stock en esa fecha e insertarlo en la temporal
        cadSQL = "SELECT sum(cantidad) FROM smoval WHERE "
        cadSQL = cadSQL & " codartic=" & DBSet(RS!codArtic, "T") & " AND codalmac=" & RS!codAlmac
        'Ver si fecha de Stock es posterior a Fecha Inventario
        If Not IsNull(RS!fechainv) Then
            'Partir del stock que habia en la fecha que se hizo inventario
            'David
            'vStock = DBLet(RS!stockinv, "N")
            vStock = TransformaComasPuntos(CStr(DBLet(RS!stockinv, "N")))
            'Obtener los movimientos entre la fecha de inventario y la fecha stock
            If CDate(cadFecha) > RS!fechainv Then
                If cadHora = "" Then
                    cadSQL = cadSQL & " AND (fechamov between '" & Format(RS!fechainv, FormatoFecha) & "' and '" & Format(cadFecha, FormatoFecha) & "')"
                Else
                    cadSQL = cadSQL & " AND (horamovi between '" & Format(RS!fechainv, FormatoFechaHora) & "' and '" & Format(cadFecha & " " & cadHora, FormatoFechaHora) & "')"
                End If
                'Movimientos de ENTRADA (tipomovi=1)
                vStock = vStock + TotMovimientosStock(cadSQL, 1)
                'Movimientos de SALIDA (tipomovi=0)
                vStock = vStock - TotMovimientosStock(cadSQL, 0)
            Else 'Fecha Inv posterior a donde queremos obtener el Stock
                'Deshacer los movimientos entre esas fecha alreves
                If cadHora = "" Then
                    cadSQL = cadSQL & " AND (fechamov between '" & Format(cadFecha, FormatoFecha) & "' and '" & Format(RS!fechainv, FormatoFecha) & "')"
                Else
                    cadSQL = cadSQL & " AND (horamovi between '" & Format(cadFecha & " " & cadHora, FormatoFechaHora) & "' and '" & Format(RS!fechainv & " " & "23:59:59", FormatoFechaHora) & "')"
                End If
                'Movimientos de ENTRADA (tipomovi=1)
                vStock = vStock - TotMovimientosStock(cadSQL, 1)
                'Movimientos de SALIDA (tipomovi=0)
                vStock = vStock + TotMovimientosStock(cadSQL, 0)
            End If
            
                 
        Else
            'Partir del stock actual
            vStock = RS!CanStock
            'Deshacer los movimientos entre esas fecha alreves
            If cadHora = "" Then
                cadSQL = cadSQL & " AND fechamov> '" & Format(cadFecha, FormatoFecha) & "' "
            Else
                cadSQL = cadSQL & " AND horamovi> '" & Format(cadFecha & " " & cadHora, FormatoFechaHora) & "' "
            End If
            'Movimientos de ENTRADA (tipomovi=1)
            vStock = vStock - TotMovimientosStock(cadSQL, 1)
            'Movimientos de SALIDA (tipomovi=0)
            vStock = vStock + TotMovimientosStock(cadSQL, 0)
        End If
        'Insertar en la Tabla TMP el stock en esa Fecha del codartic,codalmac
        cadSQL = "INSERT INTO tmpstockfec (codusu,codartic,codalmac,stock)"
        cadSQL = cadSQL & " VALUES (" & vUsu.Codigo & ", " & DBSet(RS!codArtic, "T") & ", "
        cadSQL = cadSQL & RS!codAlmac & ", " & DBSet(vStock, "N") & ")"
        Conn.Execute cadSQL
        RS.MoveNext
    Wend
    RS.Close
    Set RS = Nothing
    CargarTMPStockFecha = True
    
ECargarTMPStock:
    If Err.Number <> 0 Then
'        RS.Close
        Set RS = Nothing
        MsgBox " No se ha podido cargar la Tabla Temporal correctamente", vbInformation
    End If
End Function



Public Function DescargarDatosTMPStockFecha()
'Al salir de la aplicacion se borran los datos de la tabla temporal
Dim SQL As String

    On Error GoTo EDescargaDatos

    '------------- AHORA
    SQL = "DELETE from tmpstockfec" & " where codusu= " & vUsu.Codigo
    Conn.Execute SQL
    Exit Function
    
EDescargaDatos:
        MuestraError Err.Number, "Descargar Tabla Temporal (Stock a Fecha).", Err.Description
End Function



Private Function TotMovimientosStock(cadSQL As String, vTipomovi As Byte) As Single
'Para un tipo de Movimiento vtipomovi(0=Salida, 1=Entrada) devolver
'la cantidad de stock para esos registros de la select
Dim RSmov As ADODB.Recordset

        TotMovimientosStock = 0
        cadSQL = cadSQL & " AND tipomovi=" & vTipomovi
        Set RSmov = New ADODB.Recordset
        RSmov.Open cadSQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        If Not RSmov.EOF Then
            If Not IsNull(RSmov.Fields(0).Value) Then _
                TotMovimientosStock = RSmov.Fields(0).Value
        End If
        
        RSmov.Close
        Set RSmov = Nothing

End Function

'============ Temporales de INFORMES ====================================

Public Function TempVentasClientes(cadSel As String, cadSelPeriodo As String, cadSelAnte As String) As Boolean
'Inserta en la temporal TMPINFORMES
Dim SQL As String, Sql2 As String
Dim SQLinsert As String
Dim RS As ADODB.Recordset
Dim Cliente As String
Dim totVentas As Currency
Dim totMante As Currency
Dim totRepar As Currency
Dim totRectif As Currency
Dim totServi As Currency

Dim total As String 'Total del periodo seleccionado
Dim totalAnt As String 'total del periodo anterior

Dim i As Integer

    On Error GoTo ETmpVentas
    
    'Obtenemos el TOTAL de ventas en ese PERIODO, de todos los clientes.
    'para obtener el % de ventas de cada cliente
    '---------------------------------------------------------------------
    SQL = "select sum(baseimp1 + if(isnull(baseimp2),0,baseimp2) + if(isnull(baseimp3),0,baseimp3)) as BaseImp "
    SQL = SQL & " FROM scafac "
    If cadSelPeriodo <> "" Then SQL = SQL & " WHERE " & cadSelPeriodo
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not RS.EOF Then
        total = CStr(RS.Fields(0))
    End If
    RS.Close
    Set RS = Nothing
     
    
    
    'Obtenemos el TOTAL de ventas en el PERIODO ANTERIOR, de todos los clientes.
    'para obtener el % de ventas de cada cliente
    '---------------------------------------------------------------------------
    If cadSelAnte <> "" Then
        SQL = "select sum(baseimp1 + if(isnull(baseimp2),0,baseimp2) + if(isnull(baseimp3),0,baseimp3)) as BaseImp "
        SQL = SQL & " FROM scafac "
        If cadSelAnte <> "" Then SQL = SQL & " WHERE " & cadSelAnte
        
        Set RS = New ADODB.Recordset
        RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not RS.EOF Then
            totalAnt = CStr(RS.Fields(0))
        End If
        RS.Close
        Set RS = Nothing
    End If
    
    
    'Seleccion del PERIODO seleccionado
    'ventas por cliente y tipo de movimiento
    '----------------------------------------
    SQL = "select codtipom,codclien,nomclien,sum(baseimp1),sum(baseimp2),sum(baseimp3),sum(baseimp1 + if(isnull(baseimp2),0,baseimp2) + if(isnull(baseimp3),0,baseimp3)) as BaseImp "
    SQL = SQL & " from scafac "
    If cadSel <> "" Then SQL = SQL & " where " & cadSel
    SQL = SQL & " group by codclien,codtipom "
    SQL = SQL & " order by codclien"
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    SQLinsert = "INSERT INTO tmpinformes (codusu, codigo1,nombre1,importe1,importe2,importe3,importe4,importe5,porcen1,importeb1,importeb2,importeb3,importeb4,importeb5) "
    SQLinsert = SQLinsert & " VALUES "
    
    SQL = ""
    Sql2 = ""
    i = 0
    While Not RS.EOF
        If Cliente <> RS!CodClien Then
            If SQL <> "" Then
                SQL = SQL & DBSet(totVentas, "N") & "," & DBSet(totMante, "N") & "," & DBSet(totRepar, "N") & "," & DBSet(totRectif, "N") & ","
                '---- Laura: modificado 26/09/2006
'                totVentas = CStr(CCur(ComprobarCero(totVentas)) + CCur(ComprobarCero(totMante)) + CCur(ComprobarCero(totRepar)) + CCur(ComprobarCero(totRectif)))
                totVentas = totVentas + totMante + totRepar + totRectif + totServi
                '----
                SQL = SQL & DBSet(totVentas, "N") & ","
                '% sobre el total de ventas
                totVentas = Round((totVentas * 100) / CCur(total), 2)
                SQL = SQL & DBSet(totVentas, "N") & ","
                'Obtener ventas del cliente para el periodo anterior
                SQL = SQL & VentasPeriodoAnterior(Cliente, cadSelAnte) & ")"
                Sql2 = Sql2 & SQL & ","
            End If
            'Insertamos por bloques de 500
            If i = 30 Then
                'Insertamos en la tabla temporal
                If Sql2 <> "" Then
                    Sql2 = Mid(Sql2, 1, Len(Sql2) - 1)
                    SQL = SQLinsert & Sql2
                    Conn.Execute SQL
                End If
                
                'Reiniciamos los valores
                SQL = ""
                Sql2 = ""
                i = 0
            End If
            
            'Empezamos el registro para el siguiente cliente
            SQL = "(" & vUsu.Codigo & "," & RS!CodClien & "," & DBSet(RS!nomclien, "T") & ","
            totVentas = 0
            totMante = 0
            totRepar = 0
            totRectif = 0
            totServi = 0
            i = i + 1
        End If
        Select Case RS!codTipoM
            Case "FAV", "FTI": totVentas = totVentas + RS!BaseImp
            Case "FAM": totMante = RS!BaseImp
            Case "FAR": totRepar = RS!BaseImp
            Case "FRT": totRectif = RS!BaseImp
            Case "FAS": totServi = RS!BaseImp
        End Select
    
        Cliente = RS!CodClien
        RS.MoveNext
    Wend
    
    RS.Close
    Set RS = Nothing
    
    If SQL <> "" Then 'para el ultimo registro
        SQL = SQL & DBSet(totVentas, "N") & "," & DBSet(totMante, "N") & "," & DBSet(totRepar, "N") & "," & DBSet(totRectif, "N") & ","
        '---- Laura: Modificado 26/09/2006
        'totVentas = CStr(CCur(ComprobarCero(totVentas)) + CCur(ComprobarCero(totMante)) + CCur(ComprobarCero(totRepar)) + CCur(ComprobarCero(totRectif)))
        totVentas = CStr(totVentas + totMante + totRepar + totRectif + totServi)
        '----
        SQL = SQL & DBSet(totVentas, "N") & ","
        totVentas = CStr(Round((CCur(totVentas) * 100) / CCur(total), 2))
        SQL = SQL & DBSet(totVentas, "N") & ","
        'Obtener ventas del cliente para el periodo anterior
        SQL = SQL & VentasPeriodoAnterior(Cliente, cadSelAnte) & ")"
        Sql2 = Sql2 & SQL & ","
    End If
    
    If Sql2 <> "" Then
        Sql2 = Mid(Sql2, 1, Len(Sql2) - 1)
        SQL = SQLinsert & Sql2
        Conn.Execute SQL
    End If
    
    cadSelPeriodo = DBSet(total, "N")
    cadSelAnte = DBSet(totalAnt, "N")
    
ETmpVentas:
    If Err.Number <> 0 Then
        TempVentasClientes = False
        MuestraError Err.Number, "Ventas del periodo", Err.Description
    Else
        TempVentasClientes = True
    End If
End Function


Private Function VentasPeriodoAnterior(Cliente, cadSel) As String
Dim SQL As String
Dim RS As ADODB.Recordset
Dim totVentas As Currency
Dim totMante As Currency
Dim totRepar As Currency
Dim totRectif As Currency


    On Error GoTo EVentas
    
    totVentas = 0
    totMante = 0
    totRepar = 0
    totRectif = 0
    
    If cadSel <> "" Then
        '---- Laura: Modificaco 26/09/2006
        'SQL = "select codclien,codtipom,sum(baseimp1)+sum(baseimp2)+sum(baseimp3) as BaseImp "
        SQL = "SELECT codclien,codtipom, sum(baseimp1 + if(isnull(baseimp2),0,baseimp2) + if(isnull(baseimp3),0,baseimp3)) as BaseImp "
        '----
        SQL = SQL & " from scafac where " & cadSel
        If cadSel <> "" Then SQL = SQL & " AND "
        SQL = SQL & "(scafac.codclien = " & Cliente & ")"
        SQL = SQL & " group by codclien,codtipom "
        
        Set RS = New ADODB.Recordset
        RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not RS.EOF
             Select Case RS!codTipoM
                Case "FAV", "FTI": totVentas = totVentas + RS!BaseImp
                Case "FAM": totMante = RS!BaseImp
                Case "FAR": totRepar = RS!BaseImp
                Case "FRT": totRectif = RS!BaseImp
            End Select
            RS.MoveNext
        Wend
        RS.Close
        Set RS = Nothing
    End If
    
    SQL = DBSet(totVentas, "N") & "," & DBSet(totMante, "N") & "," & DBSet(totRepar, "N") & "," & DBSet(totRectif, "N") & ","
    '---- Laura: Modificado 26/09/2006
    'totVentas = CStr(CCur(ComprobarCero(totVentas)) + CCur(ComprobarCero(totMante)) + CCur(ComprobarCero(totRepar)) + CCur(ComprobarCero(totRectif)))
    totVentas = totVentas + totMante + totRepar + totRectif
    '----
    SQL = SQL & DBSet(totVentas, "N")
    VentasPeriodoAnterior = SQL
    
EVentas:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Ventas periodo anterior", Err.Description
    End If
End Function





Public Function TempVentasMeses(cadSel As String, Anyo As String) As Boolean
'Inseta en la tabla temporal TMPINFORMES
Dim SQL As String
Dim RS As ADODB.Recordset

Dim Cliente As Integer
Dim MesAnt As Integer
Dim i As Integer

Dim llis As Collection
Dim TotClien(12) As Currency
Dim TotAnyo(12) As Currency
Dim Porce As Single

Dim Izquierda As String
Dim Derecha As String


    On Error GoTo ETmpVentas
    
    Set llis = New Collection
    
    'Inicializamos las listas
    For i = 1 To 12
        TotClien(i) = 0
        TotAnyo(i) = 0
    Next i
    
   
    i = InStr(cadSel, "codclien")
    If i > 0 Then 'Se ha seleccionado un cliente
        SQL = "SELECT  codclien , year(fecfactu) AnyoFac,month(fecfactu) as MesFac, sum(baseimp1+if(isnull(baseimp2),0,baseimp2)+If(isnull(baseimp3),0,baseimp3)) as BaseImp "
        SQL = SQL & " FROM scafac "
        SQL = SQL & " WHERE " & cadSel '& " AND month(fecfactu)=1 "
        SQL = SQL & " GROUP BY codclien,year(fecfactu),month(fecfactu)"
        SQL = SQL & " order by codclien,month(fecfactu) asc,year(fecfactu) asc"
    Else
        'Se seleccionara el total del anyo anterior
        SQL = "SELECT  year(fecfactu) AnyoFac,month(fecfactu) as MesFac, sum(baseimp1+if(isnull(baseimp2),0,baseimp2)+If(isnull(baseimp3),0,baseimp3)) as BaseImp "
        SQL = SQL & " FROM scafac "
        SQL = SQL & " WHERE year(fecfactu)=" & Anyo - 1
        SQL = SQL & " GROUP BY year(fecfactu),month(fecfactu)"
        SQL = SQL & " order by month(fecfactu) asc,year(fecfactu) asc"
    End If
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    'Actualizar la lista con el total del cliente o del anyo anterior
    If i > 0 Then Cliente = RS!CodClien
    While Not RS.EOF
        i = CInt(RS!mesfac)
        TotClien(i) = RS!BaseImp
        RS.MoveNext
    Wend
    RS.Close
    Set RS = Nothing
    
    
    'Obtener el total del AÑO solicitado
    '-------------------------------------------------------------------
    SQL = "SELECT   year(fecfactu) AnyoFac,month(fecfactu) as MesFac, sum(baseimp1+if(isnull(baseimp2),0,baseimp2)+If(isnull(baseimp3),0,baseimp3)) as BaseImp "
    SQL = SQL & " FROM scafac "
    SQL = SQL & " WHERE  year(scafac.fecfactu) = " & Anyo
    SQL = SQL & " GROUP BY year(fecfactu),month(fecfactu)"
    SQL = SQL & " order by year(fecfactu) asc,month(fecfactu) asc"
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    'Actualizar la lista con el total del Anyo solicitado
    While Not RS.EOF
        i = CInt(RS!mesfac)
        TotAnyo(i) = RS!BaseImp
        RS.MoveNext
    Wend
    RS.Close
    Set RS = Nothing
    
    
    'Insertamos en la lista todos los registros que vamos a insertar en la temporal
    'Un registro para cada mes
    For i = 1 To 12
        If TotAnyo(i) <> 0 Then
            If Cliente <> 0 Then
                'porcentaje del cliente respecto al total del año (por mes)
                Porce = Round((TotClien(i) * 100) / TotAnyo(i), 2)
            Else
                'Incremento/decremento respecto al anyo anterior (por mes)
                'en TotClien en este caso se ha almacenado el total del año anterior de cada mes
                If TotClien(i) <> 0 Then
                    Porce = Round(((TotAnyo(i) - TotClien(i)) / TotClien(i)) * 100, 2)
                Else
                    Porce = 0
                End If
            End If
        Else
            Porce = 0
        End If
        Derecha = "(" & vUsu.Codigo & "," & Cliente & "," & Anyo & "," & i & "," & DBSet(TotClien(i), "N") & "," & DBSet(Porce, "N") & "," & DBSet(TotAnyo(i), "N") & ")"
        llis.Add Derecha
    Next i
    
    
    Izquierda = "INSERT INTO tmpinformes (codusu,codigo1,campo1,campo2,importe1,porcen1,importeb1) VALUES "
    
    
    'Insertamos en la temporal todos los registros insertados en la lista
    'recorremos toda las lista
    SQL = ""
    For i = 1 To llis.Count
        SQL = SQL & llis.Item(i) & ","
        MesAnt = MesAnt + 1
    Next i
    Set llis = Nothing
    
    
    SQL = Mid(SQL, 1, Len(SQL) - 1)
    SQL = Izquierda & SQL
    Conn.Execute SQL
 
    
ETmpVentas:
    If Err.Number <> 0 Then
        TempVentasMeses = False
        MuestraError Err.Number, "Ventas por meses", Err.Description
    Else
        TempVentasMeses = True
    End If
End Function




Public Sub BorrarTempInformes()
Dim SQL As String

    On Error GoTo EBorrar
    
    SQL = "DELETE FROM tmpinformes WHERE codusu=" & vUsu.Codigo
    Conn.Execute SQL
    
EBorrar:
    If Err.Number <> 0 Then Err.Clear
End Sub








'================================================================================
'================================================================================

'================================================================================
'TMPnlotes: Temporal para introducir los Nº de lote de los Articulos en compras
'USO:
'================================================================================


Public Function DescargarDatosTMPNumLotes(NomTabla As String, cadWHERE As String)
'Al salir de la aplicacion se borran los datos de la tabla temporal
Dim SQL As String

    On Error GoTo EDescargaDatos

     '------------- AHORA
    SQL = "DELETE from " & NomTabla & " where codusu= " & vUsu.Codigo
    If cadWHERE <> "" Then SQL = SQL & " AND " & cadWHERE
    Conn.Execute SQL
    
    Exit Function
EDescargaDatos:
        MuestraError Err.Number, "Descargar Tabla Temporal (Nº Lotes).", Err.Description
End Function



'Public Function CargarDatosTMPNumLotes(NomTabla As String, cadWhere As String) As Boolean
''IN -> NomTabla: Nombre de la tabla temporal
'Dim SQL As String
'Dim i As Integer
'Dim numlinea As String, vWhere As String
'
'    On Error GoTo ECargaDatosTMP
'
'    'Insertar tantos registros como cantidad de Articulo Introducida
''    vWhere = "(codusu=" & vUsu.Codigo & " and codartic=" & DBSet(codArtic, "T") & ")"
'
'
'
''    'insertamos tantos num.serie como cantidad
''    For i = 0 To cant - 1
''        'Obtener Num Linea
''        numlinea = SugerirCodigoSiguienteStr(NomTabla, "numlinea", vWhere)
''        'Insertar en la temporal para Nº Series
''        SQL = "INSERT INTO " & NomTabla & " (codusu, codartic, numlinealb, numlinea, numserie) VALUES ("
''        SQL = SQL & vUsu.Codigo & ", " & DBSet(codArtic, "T") & ", " & NumLinAlb & ", " & numlinea & ", ' ')"
''        Conn.Execute SQL
''    Next i
'
'ECargaDatosTMP:
'    If Err.Number <> 0 Then
'        CargarDatosTMPNumLotes = False
'        MuestraError Err.Number, "Numeros Serie", Err.Description
'    Else
'        CargarDatosTMPNumSeries = True
'    End If
'End Function



Public Function PedirNLotesGnral(ByRef RS As ADODB.Recordset, Men As Boolean) As Boolean
Dim SQL As String
'Dim b As Boolean

    On Error GoTo EPedirNLotes

    If Men Then
        SQL = "Hay artículos que tienen control de Nº de Lote." & vbCrLf & vbCrLf
        SQL = SQL & "Introduzca los Nº De Lote." & vbCrLf
        MsgBox SQL, vbInformation
    End If
    
    'Cargar la tabla temporal con tantas filas como cantidad de Articulos
    'Para introducir el Nº de lote
    SQL = "numalbar=" & DBSet(RS!numalbar, "T") & " AND fechaalb=" & DBSet(RS!FechaAlb, "F") & " AND codprove=" & DBSet(RS!codProve, "N")
    DescargarDatosTMPNumLotes "tmpnlotes", SQL
'    b = True
    
    While Not RS.EOF
'        If Not CargarDatosTMPNumSeries("tmpnseries", RS!codArtic, RS!Cantidad, RS!numlinea) Then
'            b = False
'        End If
        SQL = "INSERT INTO tmpnlotes (codusu, numalbar, fechaalb, codprove, numlinea, codartic, codalmac, nomartic, cantidad, numlotes) VALUES ("
        SQL = SQL & vUsu.Codigo & "," & DBSet(RS!numalbar, "T") & "," & DBSet(RS!FechaAlb, "F") & "," & RS!codProve & "," & RS!numlinea & "," & DBSet(RS!codArtic, "T")
        SQL = SQL & "," & DBSet(RS!codAlmac, "N") & "," & DBSet(RS!NomArtic, "T") & "," & DBSet(RS!Cantidad, "N") & "," & DBSet(RS!numlotes, "T", "S") & ")"
        Conn.Execute SQL
        RS.MoveNext
    Wend
    PedirNLotesGnral = True
    Exit Function
    'Visualizar en pantalla el Grid, y rellenar los Nº Serie
'    If Not b Then MsgBox "No se han podido mostrar todos los Artículos con Nº de Serie.", vbInformation
    
EPedirNLotes:
    If Err.Number <> 0 Then
        MsgBox Err.Number & ": " & Err.Description, vbExclamation
        PedirNLotesGnral = False
    End If
End Function
