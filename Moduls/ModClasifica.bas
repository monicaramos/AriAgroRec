Attribute VB_Name = "ModClasifica"
Option Explicit


Public Function InsertarClasificacion(ByRef Rs As ADODB.Recordset, cadErr As String, vCalidad As String) As Boolean
'Insertando en tabla conta.cabfact
Dim Sql As String
Dim Sql1 As String
Dim RS1 As ADODB.Recordset
Dim Cad As String
Dim KilosMuestra As Currency
Dim TotalKilos As Currency
Dim Calidad As Currency
Dim Diferencia As Currency
Dim HayReg As Byte
Dim TipoClasif As Byte
Dim vTipoClasif As String
Dim vCalidDest As String
Dim CalidadClasif As String
Dim CalidadVC As String

    On Error GoTo EInsertar
    
    Sql = "insert into rclasifica_clasif (numnotac,codvarie, codcalid, muestra, kilosnet) values "

    If vCalidad <> "" Then
        Sql1 = "insert into rclasifica_clasif (numnotac, codvarie, codcalid, muestra, kilosnet) "
        Sql1 = Sql1 & "values (" & DBSet(Rs!Numnotac, "N") & "," & DBSet(Rs!codvarie, "N") & ","
        Sql1 = Sql1 & DBSet(vCalidad, "N") & ",100," & DBSet(Rs!KilosNet, "N") & ")"
        
        conn.Execute Sql1
        InsertarClasificacion = True
        Exit Function
    End If
        
        

    vTipoClasif = ""
    vTipoClasif = DevuelveDesdeBDNew(cAgro, "variedades", "tipoclasifica", "codvarie", Rs!codvarie, "N")

    If CByte(vTipoClasif) = 0 Then ' clasificacion por campo
    
        Sql1 = "select rcampos_clasif.* from rcampos_clasif where codcampo = " & DBLet(Rs!codcampo, "N")
        
        Set RS1 = New ADODB.Recordset
        RS1.Open Sql1, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        If Not RS1.EOF Then
            Cad = ""
            
            TotalKilos = 0
            HayReg = 0
            
            While Not RS1.EOF
                HayReg = 1
                
                KilosMuestra = Round2(DBLet(Rs!KilosNet, "N") * DBLet(RS1!Muestra, "N") / 100, 0)
                TotalKilos = TotalKilos + KilosMuestra
                
                Cad = Cad & "(" & DBSet(Rs!Numnotac, "N") & "," & DBSet(Rs!codvarie, "N") & ","
                Cad = Cad & DBSet(RS1!codcalid, "N") & "," & DBSet(RS1!Muestra, "N") & ","
                Cad = Cad & DBSet(KilosMuestra, "N") & "),"
                
                Calidad = DBLet(RS1!codcalid, "N")
                
                RS1.MoveNext
            Wend
        
            Set RS1 = Nothing
            
            If HayReg = 1 Then
                ' quitamos la ultima coma de la cadena
                If Cad <> "" Then
                    Cad = Mid(Cad, 1, Len(Cad) - 1)
                End If
                
                Sql = Sql & Cad
                
                conn.Execute Sql
                
                ' si el kilosneto es diferente a la suma de totalkilos actualizamos la ultima linea
                If TotalKilos <> DBLet(Rs!KilosNet, "N") Then
                    Diferencia = DBLet(Rs!KilosNet, "N") - TotalKilos
                    
                    vCalidDest = CalidadDestrioenClasificacion(CStr(Rs!codvarie), CStr(Rs!Numnotac))
                    If vCalidDest <> "" Then Calidad = vCalidDest
                    
                    Sql = "update rclasifica_clasif set kilosnet = kilosnet + (" & DBSet(Diferencia, "N") & ")"
                    Sql = Sql & " where numnotac = " & DBSet(Rs!Numnotac, "N")
                    Sql = Sql & " and codvarie = " & DBSet(Rs!codvarie, "N")
                    Sql = Sql & " and codcalid = " & DBSet(Calidad, "N")
                    
                    conn.Execute Sql
                End If
            End If
        Else
            ' el campo no tiene la clasificacion
            cadErr = "El campo " & DBLet(Rs!codcampo, "N") & " no tiene clasificación. Revise."
            InsertarClasificacion = False
            Exit Function
            
        End If
    Else
        ' la clasificacion es en almacen luego insertamos tantos registros como calidades
        ' tenga la variedad
        Sql1 = "insert into rclasifica_clasif (numnotac, codvarie, codcalid, muestra, kilosnet) "
        Sql1 = Sql1 & "select " & DBSet(Rs!Numnotac, "N") & ",rcalidad.codvarie, rcalidad.codcalid, " & ValorNulo & "," & ValorNulo & " from rcalidad where codvarie = " & DBLet(Rs!codvarie, "N")
        
        conn.Execute Sql1
    
    End If
EInsertar:
    If Err.Number <> 0 Then
        InsertarClasificacion = False
        cadErr = Err.Description
    Else
        InsertarClasificacion = True
    End If
End Function



Public Function InsertarClasificacionConDestrio(ByRef Rs As ADODB.Recordset, cadErr As String, vCalidad As String, CalDestrio As String, Porcen As String) As Boolean
'Insertando en tabla conta.cabfact
Dim Sql As String
Dim Sql1 As String
Dim RS1 As ADODB.Recordset
Dim Cad As String
Dim KilosMuestra As Currency
Dim TotalKilos As Currency
Dim Calidad As Currency
Dim Diferencia As Currency
Dim HayReg As Byte
Dim TipoClasif As Byte
Dim vTipoClasif As String
Dim vCalidDest As String
Dim CalidadClasif As String
Dim CalidadVC As String

Dim KilosDest As Long
Dim KilosRet As Long

Dim vPorcen As Currency

    On Error GoTo EInsertar
    
    
    vPorcen = CCur(Porcen)
    KilosDest = Round2(DBLet(Rs!KilosNet, "N") * vPorcen / 100, 0)
    KilosRet = DBLet(Rs!KilosNet, "N") - KilosDest
    
    ' calidad de destrio
    Sql = "insert into rclasifica_clasif (numnotac,codvarie, codcalid, muestra, kilosnet) values "
    Sql = Sql & "(" & DBSet(Rs!Numnotac, "N") & "," & DBSet(Rs!codvarie, "N") & ","
    Sql = Sql & DBSet(CalDestrio, "N") & "," & DBSet(vPorcen, "N") & "," & DBSet(KilosDest, "N") & ")"

    conn.Execute Sql


    ' calidad de retirada
    Sql = "insert into rclasifica_clasif (numnotac,codvarie, codcalid, muestra, kilosnet) values "
    Sql = Sql & "(" & DBSet(Rs!Numnotac, "N") & "," & DBSet(Rs!codvarie, "N") & ","
    Sql = Sql & DBSet(vCalidad, "N") & "," & DBSet(100 - vPorcen, "N") & "," & DBSet(KilosRet, "N") & ")"
    
    conn.Execute Sql
    InsertarClasificacionConDestrio = True
    Exit Function

    
EInsertar:
    If Err.Number <> 0 Then
        InsertarClasificacionConDestrio = False
        cadErr = Err.Description
    Else
        InsertarClasificacionConDestrio = True
    End If
End Function





Public Function ActualizarTransporte(ByRef Rs As ADODB.Recordset, cadErr As String) As Boolean
Dim Sql1 As String
Dim Rs2 As ADODB.Recordset
Dim KilosDestrio As Currency
Dim Precio As Currency
Dim Transporte As Currency
Dim GasRecol As Currency
Dim Kilos As Currency


    On Error GoTo eActualizarTransporte

    If Not Rs.EOF Then

        '[Monica] 27/04/2010: si el gasto de transporte se calcula segun los portes poblacion
        '                     como se hacia inicialmente
        If vParamAplic.TipoPortesTRA = 0 Then

            Sql1 = "select imptrans from rportespobla, rpartida, rcampos, variedades "
            Sql1 = Sql1 & " where rpartida.codparti = rcampos.codparti and "
            Sql1 = Sql1 & " variedades.codprodu = rportespobla.codprodu and "
            Sql1 = Sql1 & " rpartida.codpobla = rportespobla.codpobla and "
            Sql1 = Sql1 & " variedades.codvarie = " & DBSet(Rs!codvarie, "N") & " and "
            Sql1 = Sql1 & " rcampos.codcampo = " & DBSet(Rs!codcampo, "N") & " and "
            Sql1 = Sql1 & " rcampos.codvarie = variedades.codvarie "
            
            Set Rs2 = New ADODB.Recordset
            Rs2.Open Sql1, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            
            Precio = 0
            If Not Rs2.EOF Then
                Precio = DBLet(Rs2.Fields(0).Value, "N")
            End If
            
            Set Rs2 = Nothing
            
            ' cogemos los kilos de la clasificacion que sean de destrio
            Sql1 = "select kilosnet from rclasifica_clasif, rcalidad where numnotac = " & DBSet(Rs!Numnotac, "N")
            Sql1 = Sql1 & " and rclasifica_clasif.codvarie = rcalidad.codvarie "
            Sql1 = Sql1 & " and rclasifica_clasif.codcalid = rcalidad.codcalid "
            Sql1 = Sql1 & " and rcalidad.tipcalid = 1 "
            KilosDestrio = DevuelveValor(Sql1)
            
            ' los gastos de transporte se calculan sobre los kilosnetos - los de destrio
            Kilos = DBLet(Rs!KilosNet, "N") - KilosDestrio
            Transporte = Round2(Kilos * Precio, 2)
            
            Sql1 = "update rclasifica set imptrans = " & DBSet(Transporte, "N")
            Sql1 = Sql1 & " where numnotac = " & DBSet(Rs!Numnotac, "N")
            conn.Execute Sql1
    
'        Else
'        ' [Monica] 27/04/2010 esto no haria falta aqui
'        ' el calculo de gasto de transporte se hace segun la tarifa
'            Precio = DevuelveValor("select preciokg from rtarifatra where codtarif = " & DBSet(Rs!codtarif, "N"))
'            Transporte = Round2(Rs!KilosBru * ImporteSinFormato(Precio), 2)
'
'            Sql1 = "update rclasifica set imptrans = " & DBSet(Transporte, "N")
'            Sql1 = Sql1 & " where numnotac = " & DBSet(Rs!numnotac, "N")
'            conn.Execute Sql1
'
        '[Monica]14/10/2010 : a Picassent le ponemos el transporte
        Else
            If vParamAplic.Cooperativa = 2 Or vParamAplic.Cooperativa = 16 Then
                Precio = DevuelveValor("select preciokg from rtarifatra where codtarif = " & DBSet(Rs!Codtarif, "N"))
                Transporte = Round2(DBLet(Rs!KilosTra, "N") * Precio, 2)
            
                ' metemos tambien los gastos de recoleccion
                Precio = DevuelveValor("select eurdesta from variedades where codvarie = " & DBSet(Rs!codvarie, "N"))
                GasRecol = Round2(Rs!KilosTra * Precio, 2)
            
                Sql1 = "update rclasifica set impacarr = " & DBSet(Transporte, "N")
                Sql1 = Sql1 & ", imprecol = " & DBSet(GasRecol, "N")
                Sql1 = Sql1 & " where numnotac = " & DBSet(Rs!Numnotac, "N")
                conn.Execute Sql1
            End If
        End If
    End If
eActualizarTransporte:
    If Err.Number <> 0 Then
        ActualizarTransporte = False
        cadErr = Err.Description
    Else
        ActualizarTransporte = True
    End If
    

End Function



Public Function ActualizarGastos(ByRef Rs As ADODB.Recordset, cadErr As String) As Boolean
Dim GasRecol As Currency
Dim GasAcarreo As Currency
Dim KilosTria As Long
Dim KilosNet As Long
Dim EurDesta As Currency
Dim EurRecol As Currency
Dim PrecAcarreo As Currency
Dim i As Integer
Dim Sql As String
Dim Rs2 As ADODB.Recordset



    On Error GoTo eActualizarGastos
    
    ActualizarGastos = False
    
    GasRecol = 0
    GasAcarreo = 0
    
    Sql = "select eurdesta, eurecole from variedades where codvarie = " & DBSet(Rs!codvarie, "N")
    
    Set Rs2 = New ADODB.Recordset
    Rs2.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not Rs2.EOF Then
        EurDesta = DBLet(Rs2.Fields(0).Value, "N")
        EurRecol = DBLet(Rs2.Fields(1).Value, "N")
    End If

    Set Rs2 = Nothing


    KilosNet = CLng(DBLet(Rs!KilosNet, "N"))

    'recolecta socio
    If DBLet(Rs!Recolect, "N") = 1 Then
        Sql = "select sum(kilosnet) from rclasifica_clasif, rcalidad  where numnotac = " & DBSet(Rs!Numnotac, "N")
        Sql = Sql & " and rclasifica_clasif.codvarie = rcalidad.codvarie "
        Sql = Sql & " and rclasifica_clasif.codcalid = rcalidad.codcalid "
        Sql = Sql & " and rcalidad.gastosrec = 1"
        
        KilosTria = TotalRegistros(Sql)
        
        GasRecol = Round2(KilosTria * EurRecol, 2)
    Else
    'recolecta cooperativa
        If DBLet(Rs!tiporecol, "N") = 0 Then
            'horas
            'gastosrecol = horas * personas * rparam.(costeshora + costesegso)
            GasRecol = Round2(HorasDecimal(DBLet(Rs!horastra, "N")) * CCur(DBLet(Rs!numtraba, "N")) * (vParamAplic.CosteHora + vParamAplic.CosteSegSo), 2)
        Else
            'destajo
            GasRecol = Round2(KilosNet * EurDesta, 2)
        End If
    End If
    
    PrecAcarreo = 0
    Sql = ""
    Sql = DevuelveDesdeBDNew(cAgro, "rtarifatra", "preciokg", "codtarif", CStr(DBLet(Rs!Codtarif, "N")), "N")
    If Sql <> "" Then
        PrecAcarreo = CCur(Sql)
    End If
    
    If vParamAplic.Cooperativa = 4 Then
        Sql = "select sum(kilosnet) from rclasifica_clasif, rcalidad  where numnotac = " & DBSet(Rs!Numnotac, "N")
        Sql = Sql & " and rclasifica_clasif.codvarie = rcalidad.codvarie "
        Sql = Sql & " and rclasifica_clasif.codcalid = rcalidad.codcalid "
        Sql = Sql & " and rcalidad.gastosrec = 1"
        
        KilosTria = TotalRegistros(Sql)
        
        If DBLet(Rs!transportadopor, "N") = 1 Then ' transportado por socio
            GasAcarreo = Round2(PrecAcarreo * KilosTria, 2)
        Else
            GasAcarreo = Round2(PrecAcarreo * KilosNet, 2)
        End If
    
        Sql = "update rclasifica set kilostra = " & DBSet(KilosTria, "N")
        Sql = Sql & " where numnotac = " & DBSet(Rs!Numnotac, "N")
    
        conn.Execute Sql
    Else
        GasAcarreo = Round2(PrecAcarreo * KilosNet, 2)
    End If
        
    Sql = "update rclasifica set impacarr = " & DBSet(GasAcarreo, "N")
    Sql = Sql & ", imprecol = " & DBSet(GasRecol, "N")
    Sql = Sql & " where numnotac = " & DBSet(Rs!Numnotac, "N")
    
    conn.Execute Sql
    
    ActualizarGastos = True
    Exit Function
    
eActualizarGastos:
    cadErr = cadErr & " " & Err.Description
End Function


Public Function CalculoGastosCorrectos(NumNota As String) As Boolean
Dim Rs As ADODB.Recordset
Dim RS1 As ADODB.Recordset
Dim Sql As String
Dim TotalEnvases As String
Dim TotalCostes As String
Dim Valor As Currency
Dim GasRecol As Currency
Dim GasAcarreo As Currency
Dim KilosTria As Long
Dim KilosNet As Long
Dim EurDesta As Currency
Dim EurRecol As Currency
Dim PrecAcarreo As Currency
Dim i As Integer

    On Error Resume Next
    
    
    Sql = "select * from rclasifica where numnotac = " & DBSet(NumNota, "N")
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not Rs.EOF Then
        
        GasRecol = 0
        GasAcarreo = 0
        
        If DBLet(Rs!TipoEntr, "N") = 1 Then ' es venta campo
            CalculoGastosCorrectos = True
            Exit Function
        End If
        
        Sql = "select eurdesta, eurecole from variedades where codvarie = " & DBSet(Rs!codvarie, "N")
        
        Set RS1 = New ADODB.Recordset
        RS1.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        If Not Rs.EOF Then
            EurDesta = DBLet(RS1.Fields(0).Value, "N")
            EurRecol = DBLet(RS1.Fields(1).Value, "N")
        End If
    
        Set RS1 = Nothing
    
    '    Sql = "select sum(kilosnet) from rclasifica_clasif where numnotac = " & DBSet(Text1(0).Text, "N")
    '    KilosNet = TotalRegistros(Sql)
    
        KilosNet = DBLet(Rs!KilosNet, "N")
        
        If vParamAplic.Cooperativa = 2 Or vParamAplic.Cooperativa = 16 Then KilosNet = DBLet(Rs!KilosTra, "N")
    
        'recolecta socio
        If DBLet(Rs!Recolect, "N") = 1 Then
            Sql = "select sum(kilosnet) from rclasifica_clasif, rcalidad  where numnotac = " & DBSet(NumNota, "N")
            Sql = Sql & " and rclasifica_clasif.codvarie = rcalidad.codvarie "
            Sql = Sql & " and rclasifica_clasif.codcalid = rcalidad.codcalid "
            Sql = Sql & " and rcalidad.gastosrec = 1"
            
            KilosTria = TotalRegistros(Sql)
            
            GasRecol = Round2(KilosTria * EurRecol, 2)
        Else
        'recolecta cooperativa
            If DBLet(Rs!tiporecol, "N") = 0 Then
                'horas
                'gastosrecol = horas * personas * rparam.(costeshora + costesegso)
                GasRecol = Round2(HorasDecimal(Format(DBLet(Rs!horastra, "N"), "###,##0.00")) * DBLet(Rs!numtraba, "N") * (vParamAplic.CosteHora + vParamAplic.CosteSegSo), 2)
            Else
                'destajo
                GasRecol = Round2(KilosNet * EurDesta, 2)
            End If
        End If
        
        If vParamAplic.Cooperativa = 2 Or vParamAplic.Cooperativa = 16 Then GasRecol = Round2(KilosNet * EurDesta, 2)

'12/05/2009
'        If DBLet(Rs!codtarif, "N") <> 0 Then
'            Sql = DevuelveDesdeBDNew(cAgro, "rtarifatra", "preciokg", "codtarif", Rs!codtarif, "N")
'            PrecAcarreo = CCur(Sql)
'        Else
'            PrecAcarreo = 0
'        End If
'12/05/2009 cambiado por esto pq si que hay tarifa 0
        
        PrecAcarreo = 0
        Sql = ""
        Sql = DevuelveDesdeBDNew(cAgro, "rtarifatra", "preciokg", "codtarif", DBLet(Rs!Codtarif, "N"), "N")
        If Sql <> "" Then
            PrecAcarreo = CCur(Sql)
        End If
        
        GasAcarreo = Round2(PrecAcarreo * KilosNet, 2)
        
        CalculoGastosCorrectos = Not (((DBLet(Rs!imprecol, "N") <> GasRecol) Or (DBLet(Rs!impacarr, "N") <> GasAcarreo)))
    End If
    
End Function


