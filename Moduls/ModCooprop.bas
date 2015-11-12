Attribute VB_Name = "ModCooprop"
Option Explicit

'MODULO PARA EL REPARTO ENTRE COOPROPIETARIOS


'================================================================================
'================================================================================

'================================================================================

' Variables de reparto
Dim vKilosBru As Long
Dim vNumcajon As Long
Dim vKilosNet As Single
Dim vImpTrans As Single
Dim vImpAcarr As Single
Dim vImpRecol As Single
Dim vImppenal As Single
Dim vImpEntrada As Single
Dim vTaraBodega As Single

Dim vMuestra As Single


'****************************************************************************
'******************************REPARTO A COOPROPIETARIOS*********************
'****************************************************************************



Public Function RepartoAlbaranes(NumAlbarOrigen As Long, cadErr As String) As Boolean
'Comprobar si hay registros a Mostrar antes de abrir el Informe
Dim SQL As String
Dim Sql2 As String
Dim Sql4 As String
Dim RS As ADODB.Recordset
Dim Rs2 As ADODB.Recordset
Dim rs3 As ADODB.Recordset

Dim Numalbar As Long
Dim vTipoMov As CTiposMov

Dim Albaranes As String

Dim tKilosBru As Long
Dim tNumcajon As Long
Dim tKilosNet As Single
Dim tImpTrans As Single
Dim tImpAcarr As Single
Dim tImpRecol As Single
Dim tImppenal As Single
Dim tImpEntrada As Single
Dim tTarabodega As Single
Dim CodTipoMov As String
Dim b As Boolean
Dim devuelve As String
Dim Existe As Boolean
Dim Mens As String
Dim NroPropiedad As String
Dim Numreg As Long
Dim Campo As Long
Dim cTabla As String
Dim NumF As Long



    On Error GoTo eRepartoAlbaranes

    RepartoAlbaranes = False
    
    cadErr = ""
    
    b = True
    
    SQL = "select * from rhisfruta where numalbar = " & DBSet(NumAlbarOrigen, "N")
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not RS.EOF Then
        
        If TieneCopropietarios(CStr(RS!CodCampo), CStr(RS!Codsocio)) Then
        
            '[Monica]31/08/2012: reparto de entradas en bodega
            If (vParamAplic.Cooperativa = 3 Or vParamAplic.Cooperativa = 7) And EsVariedadGrupo6(RS!codvarie) Then
                cTabla = "(rhisfruta inner join variedades on rhisfruta.codvarie = variedades.codvarie) " & _
                         " inner join productos on variedades.codprodu = productos.codprodu and codgrupo = 6 "
                
                NumF = SugerirCodigoSiguienteStr(cTabla, "numalbar")
                NumF = NumF - 1
                
                Numalbar = NumF
                
                If NumF < 0 Then b = False
            Else
        
                CodTipoMov = "ALF"
            
                Set vTipoMov = New CTiposMov
                If Not vTipoMov.Leer(CodTipoMov) Then Exit Function

            End If

                Albaranes = ""

                tKilosBru = DBLet(RS!KilosBru, "N")
                tNumcajon = DBLet(RS!Numcajon, "N")
                tKilosNet = DBLet(RS!KilosNet, "N")
                tImpTrans = DBLet(RS!ImpTrans, "N")
                tImpAcarr = DBLet(RS!impacarr, "N")
                tImpRecol = DBLet(RS!imprecol, "N")
                tImppenal = DBLet(RS!ImpPenal, "N")
                tImpEntrada = DBLet(RS!ImpEntrada, "N")
                tTarabodega = DBLet(RS!tarabodega, "N")

                Sql2 = "select * from rcampos_cooprop where codcampo = " & DBSet(RS!CodCampo, "N")
                Sql2 = Sql2 & " and rcampos_cooprop.codsocio <> " & DBSet(RS!Codsocio, "N")
                Sql2 = Sql2 & " order by numlinea "
                
                Set Rs2 = New ADODB.Recordset
                Rs2.Open Sql2, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                
                While Not Rs2.EOF And b
                    
                    If (vParamAplic.Cooperativa = 3 Or vParamAplic.Cooperativa = 7) And EsVariedadGrupo6(RS!codvarie) Then
                        
                        Numalbar = Numalbar + 1
                    
                    Else
                    
                        Numalbar = vTipoMov.ConseguirContador(CodTipoMov)
                        Do
                            devuelve = DevuelveDesdeBDNew(cAgro, "rhisfruta", "numalbar", "numalbar", CStr(Numalbar), "N")
                            If devuelve <> "" Then
                                'Ya existe el contador incrementarlo
                                Existe = True
                                vTipoMov.IncrementarContador (CodTipoMov)
                                Numalbar = vTipoMov.ConseguirContador(CodTipoMov)
                            Else
                                Existe = False
                            End If
                        Loop Until Not Existe
                    End If
                    
                    vKilosBru = Round2(DBLet(RS!KilosBru, "N") * DBLet(Rs2!Porcentaje, "N") / 100)
                    vNumcajon = Round2(DBLet(RS!Numcajon, "N") * DBLet(Rs2!Porcentaje, "N") / 100)
                    vKilosNet = Round2(DBLet(RS!KilosNet, "N") * DBLet(Rs2!Porcentaje, "N") / 100)
                    vImpTrans = Round2(DBLet(RS!ImpTrans, "N") * DBLet(Rs2!Porcentaje, "N") / 100, 2)
                    vImpAcarr = Round2(DBLet(RS!impacarr, "N") * DBLet(Rs2!Porcentaje, "N") / 100, 2)
                    vImpRecol = Round2(DBLet(RS!imprecol, "N") * DBLet(Rs2!Porcentaje, "N") / 100, 2)
                    vImppenal = Round2(DBLet(RS!ImpPenal, "N") * DBLet(Rs2!Porcentaje, "N") / 100, 2)
                    vImpEntrada = Round2(DBLet(RS!ImpEntrada, "N") * DBLet(Rs2!Porcentaje, "N") / 100, 2)
                    vTaraBodega = Round2(DBLet(RS!tarabodega, "N") * DBLet(Rs2!Porcentaje, "N") / 100, 2)
                    
                    tKilosBru = tKilosBru - vKilosBru
                    tNumcajon = tNumcajon - vNumcajon
                    tKilosNet = tKilosNet - vKilosNet
                    tImpTrans = tImpTrans - vImpTrans
                    tImpAcarr = tImpAcarr - vImpAcarr
                    tImpRecol = tImpRecol - vImpRecol
                    tImppenal = tImppenal - vImppenal
                    tImpEntrada = tImpEntrada - vImpEntrada
                    tTarabodega = tTarabodega - vTaraBodega
                    
                    Sql4 = "insert into rhisfruta (numalbar,fecalbar,codvarie,codsocio,codcampo,tipoentr,recolect,"
                    Sql4 = Sql4 & "kilosbru,numcajon,kilosnet,imptrans,impacarr,imprecol,imppenal,impreso,impentrada,"
                    Sql4 = Sql4 & "cobradosn,prestimado,coddeposito,codpobla,transportadopor,albarorigen, tarabodega,codtraba,tolva,kgradobonif,esbonifespecial) values ("
                    Sql4 = Sql4 & DBSet(Numalbar, "N") & ","
                    Sql4 = Sql4 & DBSet(RS!Fecalbar, "F") & ","
                    Sql4 = Sql4 & DBSet(RS!codvarie, "N") & ","
                    Sql4 = Sql4 & DBSet(Rs2!Codsocio, "N") & ","
                    Sql4 = Sql4 & DBSet(RS!CodCampo, "N") & ","
                    Sql4 = Sql4 & DBSet(RS!TipoEntr, "N") & ","
                    Sql4 = Sql4 & DBSet(RS!Recolect, "N") & ","
                    Sql4 = Sql4 & DBSet(vKilosBru, "N") & ","
                    Sql4 = Sql4 & DBSet(vNumcajon, "N") & ","
                    Sql4 = Sql4 & DBSet(vKilosNet, "N") & ","
                    Sql4 = Sql4 & DBSet(vImpTrans, "N", "S") & ","
                    Sql4 = Sql4 & DBSet(vImpAcarr, "N", "S") & ","
                    Sql4 = Sql4 & DBSet(vImpRecol, "N", "S") & ","
                    Sql4 = Sql4 & DBSet(vImppenal, "N", "S") & ","
                    Sql4 = Sql4 & DBSet(RS!impreso, "N") & ","
                    Sql4 = Sql4 & DBSet(vImpEntrada, "N") & ","
                    Sql4 = Sql4 & DBSet(RS!cobradosn, "N") & ","
                    Sql4 = Sql4 & DBSet(RS!PrEstimado, "N", "S") & ","
                    
                    '[Monica]31/08/2012: si venimos  de una entrada de bodega
                    If EsVariedadGrupo6(RS!codvarie) Then
                        Sql4 = Sql4 & DBSet(RS!coddeposito, "N") & ","
                    Else
                        Sql4 = Sql4 & DBSet(RS!coddeposito, "N", "S") & ","
                    End If
                    
                    Sql4 = Sql4 & DBSet(RS!CodPobla, "N", "S") & ","
                    Sql4 = Sql4 & DBSet(RS!transportadopor, "N") & ","
                    Sql4 = Sql4 & DBSet(NumAlbarOrigen, "N") & ","
                    Sql4 = Sql4 & DBSet(vTaraBodega, "N") & ","
                    Sql4 = Sql4 & DBSet(RS!CodTraba, "N") & ","
                    Sql4 = Sql4 & DBSet(RS!tolva, "N") & ","
                    Sql4 = Sql4 & DBSet(RS!kgradobonif, "N") & ","
                    Sql4 = Sql4 & DBSet(RS!esbonifespecial, "N") & ")"
                    
                    conn.Execute Sql4
                    
                    Mens = "Reparto de Entradas."
                    If b Then b = RepartoEntradas(NumAlbarOrigen, Numalbar, DBLet(Rs2!Porcentaje, "N"), Mens)
                
                    Mens = "Reparto de Clasificación."
                    If b Then b = RepartoClasificacion(NumAlbarOrigen, Numalbar, DBLet(Rs2!Porcentaje, "N"), Mens)
                    
                    Mens = "Reparto de Gastos."
                    If b Then b = RepartoGastos(NumAlbarOrigen, Numalbar, DBLet(Rs2!Porcentaje, "N"), Mens)
                    
                    Mens = "Grabar Incidencias."
                    If b Then b = GrabarIncidencias(NumAlbarOrigen, Numalbar, Mens)
                
                    Albaranes = Albaranes & Numalbar & ","
                    
                
                    Rs2.MoveNext
                Wend
                
                Set Rs2 = Nothing
                
                If b Then
                    ' ultimo registro la diferencia ( se updatean las tablas del registro de rhisfruta origen )
                    Sql4 = "update rhisfruta set kilosbru = " & DBSet(tKilosBru, "N") & ","
                    Sql4 = Sql4 & "numcajon = " & DBSet(tNumcajon, "N") & ","
                    Sql4 = Sql4 & "kilosnet = " & DBSet(tKilosNet, "N") & ","
                    Sql4 = Sql4 & "imptrans = " & DBSet(tImpTrans, "N") & ","
                    Sql4 = Sql4 & "impacarr = " & DBSet(tImpAcarr, "N") & ","
                    Sql4 = Sql4 & "imprecol = " & DBSet(tImpRecol, "N") & ","
                    Sql4 = Sql4 & "Imppenal = " & DBSet(tImppenal, "N") & ","
                    Sql4 = Sql4 & "Impentrada = " & DBSet(tImpEntrada, "N") & ","
                    Sql4 = Sql4 & "tarabodega = " & DBSet(tTarabodega, "N") & ","
                    '[Monica]09/01/2013: marcamos la entrada como que ha sido repartida entre los coopropietarios
                    Sql4 = Sql4 & "estarepcooprop = 1 "
                    Sql4 = Sql4 & " where numalbar = " & DBSet(NumAlbarOrigen, "N")
                    
                    conn.Execute Sql4
                
                    Albaranes = "(" & Mid(Albaranes, 1, Len(Albaranes) - 1) & ")"
                
                    Mens = "Actualizar Entradas."
                    If b Then b = ActualizarEntradas(NumAlbarOrigen, Albaranes, Mens)
                
                    Mens = "Actualizar Clasificación."
                    If b Then b = ActualizarClasificacion(NumAlbarOrigen, Albaranes, Mens)
                    
                    Mens = "Actualizara Gastos."
                    If b Then b = ActualizarGastosAlbaranes(NumAlbarOrigen, Albaranes, Mens)
                
                    If Not ((vParamAplic.Cooperativa = 3 Or vParamAplic.Cooperativa = 7) And EsVariedadGrupo6(RS!codvarie)) Then
                
                        vTipoMov.IncrementarContador (CodTipoMov)
                    
                    End If
                    
                ' fin de ultimo registro
                End If
'            Else
'                b = False
'            End If
        
        End If
    
    End If
    
    Set RS = Nothing

eRepartoAlbaranes:
    If Err.Number <> 0 Or Not b Then
        cadErr = cadErr & Err.Description
    Else
        RepartoAlbaranes = True
    End If
End Function



'****************************************************************************
'******************************REPARTO A COOPROPIETARIOS*********************
'****************************************************************************

' Funcion utilizada en el reparto de entradas a coopropietarios en la actualizacion de entradas bascula


Public Function RepartoAlbaranesBascula(NumAlbarOrigen As Long, cadErr As String) As Boolean
'Comprobar si hay registros a Mostrar antes de abrir el Informe
Dim SQL As String
Dim Sql1 As String
Dim Sql2 As String
Dim Sql4 As String
Dim RS As ADODB.Recordset
Dim Rs2 As ADODB.Recordset
Dim rs3 As ADODB.Recordset

Dim Numalbar As Long
Dim vTipoMov As CTiposMov

Dim Albaranes As String

Dim tKilosBru As Long
Dim tNumcajon As Long
Dim tKilosNet As Single
Dim tImpTrans As Single
Dim tImpAcarr As Single
Dim tImpRecol As Single
Dim tImppenal As Single
Dim tImpEntrada As Single
Dim CodTipoMov As String
Dim b As Boolean
Dim devuelve As String
Dim Existe As Boolean
Dim Mens As String
Dim NroPropiedad As String
Dim Numreg As Long
Dim Campo As Long
Dim Porcentaje As Single

    On Error GoTo eRepartoAlbaranesBascula

    RepartoAlbaranesBascula = False
    
    cadErr = ""
    
    b = True
    
    SQL = "select * from rclasifica where numnotac = " & DBSet(NumAlbarOrigen, "N")
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not RS.EOF Then
        
        If TieneCopropietarios(CStr(RS!CodCampo), CStr(RS!Codsocio)) Then
            CodTipoMov = "NOC"
        
            Set vTipoMov = New CTiposMov
            If vTipoMov.Leer(CodTipoMov) Then

                Albaranes = ""

                tKilosBru = DBLet(RS!KilosBru, "N")
                tNumcajon = DBLet(RS!Numcajon, "N")
                tKilosNet = DBLet(RS!KilosNet, "N")
                tImpTrans = DBLet(RS!ImpTrans, "N")
                tImpAcarr = DBLet(RS!impacarr, "N")
                tImpRecol = DBLet(RS!imprecol, "N")
                tImppenal = DBLet(RS!ImpPenal, "N")

                Sql2 = "select * from rcampos_cooprop where codcampo = " & DBSet(RS!CodCampo, "N")
                Sql2 = Sql2 & " and rcampos_cooprop.codsocio <> " & DBSet(RS!Codsocio, "N")
                Sql2 = Sql2 & " order by numlinea "
                
                Set Rs2 = New ADODB.Recordset
                Rs2.Open Sql2, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                
                While Not Rs2.EOF And b
                    Numalbar = vTipoMov.ConseguirContador(CodTipoMov)
                    Do
                        devuelve = DevuelveDesdeBDNew(cAgro, "rclasifica", "numnotac", "numnotac", CStr(Numalbar), "N")
                        If devuelve <> "" Then
                            'Ya existe el contador incrementarlo
                            Existe = True
                            vTipoMov.IncrementarContador (CodTipoMov)
                            Numalbar = vTipoMov.ConseguirContador(CodTipoMov)
                        Else
                            Existe = False
                        End If
                    Loop Until Not Existe
                    
                    vKilosBru = Round2(DBLet(RS!KilosBru, "N") * DBLet(Rs2!Porcentaje, "N") / 100)
                    vNumcajon = Round2(DBLet(RS!Numcajon, "N") * DBLet(Rs2!Porcentaje, "N") / 100)
                    vKilosNet = Round2(DBLet(RS!KilosNet, "N") * DBLet(Rs2!Porcentaje, "N") / 100)
                    vImpTrans = Round2(DBLet(RS!ImpTrans, "N") * DBLet(Rs2!Porcentaje, "N") / 100, 2)
                    vImpAcarr = Round2(DBLet(RS!impacarr, "N") * DBLet(Rs2!Porcentaje, "N") / 100, 2)
                    vImpRecol = Round2(DBLet(RS!imprecol, "N") * DBLet(Rs2!Porcentaje, "N") / 100, 2)
                    vImppenal = Round2(DBLet(RS!ImpPenal, "N") * DBLet(Rs2!Porcentaje, "N") / 100, 2)
                    
                    tKilosBru = tKilosBru - vKilosBru
                    tNumcajon = tNumcajon - vNumcajon
                    tKilosNet = tKilosNet - vKilosNet
                    tImpTrans = tImpTrans - vImpTrans
                    tImpAcarr = tImpAcarr - vImpAcarr
                    tImpRecol = tImpRecol - vImpRecol
                    tImppenal = tImppenal - vImppenal
                    
                    Sql4 = "insert into rclasifica (numnotac,fechaent,horaentr,codvarie,codsocio,codcampo,tipoentr,recolect,"
                    Sql4 = Sql4 & "codtrans,codcapat,codtarif,kilosbru,numcajon,kilosnet,observac,imptrans,impacarr,"
                    Sql4 = Sql4 & "imprecol,imppenal,tiporecol,horastra,numtraba,numalbar,fecalbar,impreso,prestimado,"
                    Sql4 = Sql4 & "transportadopor,kilostra)  values ("
                    Sql4 = Sql4 & DBSet(Numalbar, "N") & ","
                    Sql4 = Sql4 & DBSet(RS!FechaEnt, "F") & ","
                    Sql4 = Sql4 & DBSet(RS!horaentr, "FH") & ","
                    Sql4 = Sql4 & DBSet(RS!codvarie, "N") & ","
                    Sql4 = Sql4 & DBSet(Rs2!Codsocio, "N") & ","
                    Sql4 = Sql4 & DBSet(RS!CodCampo, "N") & ","
                    Sql4 = Sql4 & DBSet(RS!TipoEntr, "N") & ","
                    Sql4 = Sql4 & DBSet(RS!Recolect, "N") & ","
                    Sql4 = Sql4 & DBSet(RS!codTrans, "N") & ","
                    Sql4 = Sql4 & DBSet(RS!codcapat, "N") & ","
                    Sql4 = Sql4 & DBSet(RS!Codtarif, "N") & ","
                    Sql4 = Sql4 & DBSet(vKilosBru, "N") & ","
                    Sql4 = Sql4 & DBSet(vNumcajon, "N") & ","
                    Sql4 = Sql4 & DBSet(vKilosNet, "N") & ","
                    Sql4 = Sql4 & DBSet(RS!Observac, "T") & ","
                    Sql4 = Sql4 & DBSet(vImpTrans, "N", "S") & ","
                    Sql4 = Sql4 & DBSet(vImpAcarr, "N", "S") & ","
                    Sql4 = Sql4 & DBSet(vImpRecol, "N", "S") & ","
                    Sql4 = Sql4 & DBSet(vImppenal, "N", "S") & ","
                    Sql4 = Sql4 & DBSet(RS!tiporecol, "N") & ","
                    Sql4 = Sql4 & DBSet(RS!horastra, "N", "S") & ","
                    Sql4 = Sql4 & DBSet(RS!numtraba, "N", "S") & ","
                    Sql4 = Sql4 & DBSet(RS!Numalbar, "N", "S") & ","
                    Sql4 = Sql4 & DBSet(RS!Fecalbar, "F", "S") & ","
                    Sql4 = Sql4 & DBSet(RS!impreso, "N") & ","
                    Sql4 = Sql4 & DBSet(RS!PrEstimado, "N", "S") & ","
                    Sql4 = Sql4 & DBSet(RS!transportadopor, "N") & ","
                    Sql4 = Sql4 & DBSet(RS!KilosTra, "N") & ")"
                    
                    conn.Execute Sql4
                    
                    '[Monica]25/03/2014: en el caso de que en la origen hayan plagas la copiamos
                    '                    solo puede ocurrir en el caso de quatretonda pq en el mto de entradas bascula indicamos sin ausencia de plagas
                    SQL = "select * from rclasifica_incidencia where numnotac = " & DBSet(NumAlbarOrigen, "N")
                    If TotalRegistros(SQL) <> 0 Then
                        Sql4 = "insert into rclasifica_incidencia (numnotac, codincid) select " & DBSet(Numalbar, "N") & ", codincid "
                        Sql4 = Sql4 & " from rclasifica_incidencia where numnotac = " & DBSet(NumAlbarOrigen, "N")
                    
                        conn.Execute Sql4
                    End If
                    
                    Mens = "Reparto de Clasificación Báscula."
                    If b Then b = RepartoClasificacionBascula(NumAlbarOrigen, Numalbar, DBLet(Rs2!Porcentaje, "N"), Mens)
                    
                    Albaranes = Albaranes & Numalbar & ","
                
                    Rs2.MoveNext
                Wend
                
                Set Rs2 = Nothing
                
                If b Then
                    ' ultimo registro la diferencia ( se updatean las tablas del registro de rclasifica origen )
                    Sql4 = "update rclasifica set kilosbru = " & DBSet(tKilosBru, "N") & ","
                    Sql4 = Sql4 & "numcajon = " & DBSet(tNumcajon, "N") & ","
                    Sql4 = Sql4 & "kilosnet = " & DBSet(tKilosNet, "N") & ","
                    Sql4 = Sql4 & "imptrans = " & DBSet(tImpTrans, "N") & ","
                    Sql4 = Sql4 & "impacarr = " & DBSet(tImpAcarr, "N") & ","
                    Sql4 = Sql4 & "imprecol = " & DBSet(tImpRecol, "N") & ","
                    Sql4 = Sql4 & "Imppenal = " & DBSet(tImppenal, "N")
                    Sql4 = Sql4 & " where numnotac = " & DBSet(NumAlbarOrigen, "N")
                    
                    conn.Execute Sql4
                    
                    vKilosNet = tKilosNet
                    Sql2 = "select porcentaje from rcampos_cooprop where codcampo = " & DBSet(RS!CodCampo, "N")
                    Sql2 = Sql2 & " and rcampos_cooprop.codsocio = " & DBSet(RS!Codsocio, "N")
                    Porcentaje = DevuelveValor(Sql2)

                    Mens = "Reparto de Clasificación Báscula."
                    If b Then b = RepartoClasificacionBascula(NumAlbarOrigen, NumAlbarOrigen, DBLet(Porcentaje, "N"), Mens)
                
                    Albaranes = "(" & Mid(Albaranes, 1, Len(Albaranes) - 1) & ")"
                
                    vTipoMov.IncrementarContador (CodTipoMov)
                ' fin de ultimo registro
                End If
            Else
                b = False
            End If
        
        End If
    
    End If
    
    Set RS = Nothing

eRepartoAlbaranesBascula:
    If Err.Number <> 0 Or Not b Then
        cadErr = cadErr & Err.Description
    Else
        RepartoAlbaranesBascula = True
    End If
End Function


Private Function TieneCopropietarios(Campo As String, Propietario As String) As Boolean
Dim NroCampo As String
Dim SQL As String

    SQL = "select count(*) from rcampos_cooprop where codcampo = " & DBSet(Campo, "N") & " and codsocio <> " & DBSet(Propietario, "N")
    
    TieneCopropietarios = TotalRegistros(SQL) > 0

End Function


Private Function RepartoEntradas(AlbAnt As Long, Numalbar As Long, Porcentaje As Currency, Mens As String) As Boolean
Dim SQL As String
Dim RS As ADODB.Recordset
Dim Sql2 As String

Dim lKilosBru As Long
Dim lNumcajon As Long
Dim lKilosNet As Single
Dim lImpTrans As Single
Dim lImpAcarr As Single
Dim lImpRecol As Single
Dim lImppenal As Single
Dim lImpEntrada As Single
    
Dim tKilosBru As Long
Dim tNumcajon As Long
Dim tKilosNet As Single
Dim tImpTrans As Single
Dim tImpAcarr As Single
Dim tImpRecol As Single
Dim tImppenal As Single
Dim tImpEntrada As Single

Dim NumNota As Long

    On Error GoTo eRepartoEntradas

    RepartoEntradas = False


    tKilosBru = vKilosBru
    tNumcajon = vNumcajon
    tKilosNet = vKilosNet
    tImpTrans = vImpTrans
    tImpAcarr = vImpAcarr
    tImpRecol = vImpRecol
    tImppenal = vImppenal

    SQL = "select * from rhisfruta_entradas where numalbar = " & DBSet(AlbAnt, "N")
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not RS.EOF
        lKilosBru = Round2(DBLet(RS!KilosBru, "N") * Porcentaje / 100)
        lNumcajon = Round2(DBLet(RS!Numcajon, "N") * Porcentaje / 100)
        lKilosNet = Round2(DBLet(RS!KilosNet, "N") * Porcentaje / 100)
        lImpTrans = Round2(DBLet(RS!ImpTrans, "N") * Porcentaje / 100, 2)
        lImpAcarr = Round2(DBLet(RS!impacarr, "N") * Porcentaje / 100, 2)
        lImpRecol = Round2(DBLet(RS!imprecol, "N") * Porcentaje / 100, 2)
        lImppenal = Round2(DBLet(RS!ImpPenal, "N") * Porcentaje / 100, 2)
        
        tKilosBru = tKilosBru - lKilosBru
        tNumcajon = tNumcajon - lNumcajon
        tKilosNet = tKilosNet - lKilosNet
        tImpTrans = tImpTrans - lImpTrans
        tImpAcarr = tImpAcarr - lImpAcarr
        tImpRecol = tImpRecol - lImpRecol
        tImppenal = tImppenal - lImppenal
       
        NumNota = RS!numnotac
        
        Sql2 = "insert into rhisfruta_entradas (numalbar,numnotac,fechaent,horaentr,kilosbru,numcajon,kilosnet,observac,imptrans,"
        Sql2 = Sql2 & "impacarr,imprecol,imppenal,prestimado,codtrans,codtarif,codcapat) values ("
        Sql2 = Sql2 & DBSet(Numalbar, "N") & ","
        Sql2 = Sql2 & DBSet(RS!numnotac, "N") & ","
        Sql2 = Sql2 & DBSet(RS!FechaEnt, "F") & ","
        Sql2 = Sql2 & DBSet(RS!horaentr, "FH") & ","
        Sql2 = Sql2 & DBSet(lKilosBru, "N") & ","
        Sql2 = Sql2 & DBSet(lNumcajon, "N") & ","
        Sql2 = Sql2 & DBSet(lKilosNet, "N") & ","
        Sql2 = Sql2 & DBSet(RS!Observac, "T", "S") & ","
        Sql2 = Sql2 & DBSet(lImpTrans, "N", "S") & ","
        Sql2 = Sql2 & DBSet(lImpAcarr, "N", "S") & ","
        Sql2 = Sql2 & DBSet(lImpRecol, "N", "S") & ","
        Sql2 = Sql2 & DBSet(lImppenal, "N", "S") & ","
        Sql2 = Sql2 & DBSet(RS!PrEstimado, "N", "S") & ","
        Sql2 = Sql2 & DBSet(RS!codTrans, "T", "S") & ","
        Sql2 = Sql2 & DBSet(RS!Codtarif, "N", "S") & ","
        Sql2 = Sql2 & DBSet(RS!codcapat, "N", "S") & ")"
        
        conn.Execute Sql2
    
        RS.MoveNext
    Wend
    
    Sql2 = "update rhisfruta_entradas set kilosbru = kilosbru + " & DBSet(tKilosBru, "N")
    Sql2 = Sql2 & ", numcajon = numcajon + " & DBSet(tNumcajon, "N")
    Sql2 = Sql2 & ", kilosnet = kilosnet + " & DBSet(tKilosNet, "N")
    Sql2 = Sql2 & ", imptrans = imptrans + " & DBSet(tImpTrans, "N")
    Sql2 = Sql2 & ", impacarr = impacarr + " & DBSet(tImpAcarr, "N")
    Sql2 = Sql2 & ", imprecol = imprecol + " & DBSet(tImpAcarr, "N")
    Sql2 = Sql2 & ", imppenal = imppenal + " & DBSet(tImppenal, "N")
    Sql2 = Sql2 & " where numalbar = " & DBSet(Numalbar, "N") & " and numnotac = " & DBSet(NumNota, "N")
    
    conn.Execute Sql2
    
    Set RS = Nothing
    
    RepartoEntradas = True
    Exit Function
    
    
eRepartoEntradas:
    Mens = Mens & vbCrLf & "Reparto Entradas. " & Err.Description
End Function



Private Function RepartoClasificacion(AlbAnt As Long, Numalbar As Long, Porcentaje As Currency, Mens As String) As Boolean
Dim SQL As String
Dim RS As ADODB.Recordset
Dim Sql2 As String

Dim lKilosNet As Single
Dim tKilosNet As Single
Dim Calid As Long
Dim rs3 As ADODB.Recordset
Dim Sql3 As String

    On Error GoTo eRepartoClasificacion

    RepartoClasificacion = False


    tKilosNet = vKilosNet
    Calid = 0
    
    SQL = "select * from rhisfruta_clasif where numalbar = " & DBSet(AlbAnt, "N")
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not RS.EOF
        lKilosNet = Round2(DBLet(RS!KilosNet, "N") * Porcentaje / 100)
        
        If lKilosNet <> 0 And Calid = 0 Then
            Calid = DBLet(RS!codcalid, "N")
        End If
        
        tKilosNet = tKilosNet - lKilosNet
        
        Sql2 = "insert into rhisfruta_clasif (numalbar,codvarie,codcalid,kilosnet) values ("
        Sql2 = Sql2 & DBSet(Numalbar, "N") & ","
        Sql2 = Sql2 & DBSet(RS!codvarie, "N") & ","
        Sql2 = Sql2 & DBSet(RS!codcalid, "N") & ","
        Sql2 = Sql2 & DBSet(lKilosNet, "N", "S") & ")"
        
        conn.Execute Sql2
    
        RS.MoveNext
    Wend
    
    '[Monica]24/05/2013: si hay alguna calidad con kilos superior a la diferencia de redondeo
    Sql3 = "select min(codcalid) from rhisfruta_clasif where numalbar = " & DBSet(Numalbar, "N")
    Sql3 = Sql3 & " and kilosnet >= " & DBSet(tKilosNet * (-1), "N")
    Set rs3 = New ADODB.Recordset
    rs3.Open Sql3, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not rs3.EOF Then
        Calid = DBLet(rs3.Fields(0).Value, "N")
    End If
    Set rs3 = Nothing
    '24/05/2013: hasta aqui
    
    
    
    ' si no está correcto el totalkilonet el redondeo  va sobre la primera calidad con kilos
    Sql2 = "update rhisfruta_clasif set kilosnet = kilosnet + " & DBSet(tKilosNet, "N")
    Sql2 = Sql2 & " where numalbar = " & DBSet(Numalbar, "N") & " and codcalid = " & DBSet(Calid, "N")
    
    conn.Execute Sql2
    
    Set RS = Nothing
    
    RepartoClasificacion = True
    Exit Function
    
eRepartoClasificacion:
    Mens = Mens & vbCrLf & "Reparto Clasificación. " & Err.Description
End Function


Private Function RepartoClasificacionBascula(AlbAnt As Long, Numalbar As Long, Porcentaje As Currency, Mens As String) As Boolean
Dim SQL As String
Dim RS As ADODB.Recordset
Dim Sql2 As String

Dim lKilosNet As Single
Dim tKilosNet As Single
Dim tMuestra As Single
Dim lMuestra As Single
Dim Calid As Long
Dim Variedad As Long
Dim Sql3 As String
Dim rs3 As ADODB.Recordset

    On Error GoTo eRepartoClasificacionBascula

    RepartoClasificacionBascula = False


    tKilosNet = vKilosNet
    tMuestra = DevuelveValor("select sum(muestra) from rclasifica_clasif where numnotac = " & DBSet(AlbAnt, "N"))
    Calid = 0
    
    SQL = "select * from rclasifica_clasif where numnotac = " & DBSet(AlbAnt, "N")
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not RS.EOF
        Variedad = DBLet(RS!codvarie, "N")
        
        lKilosNet = Round2(DBLet(RS!KilosNet, "N") * Porcentaje / 100)
        lMuestra = Round2(DBLet(RS!Muestra, "N") * Porcentaje / 100)
        If lKilosNet <> 0 And Calid = 0 Then
            Calid = DBLet(RS!codcalid, "N")
        End If
        
        tKilosNet = tKilosNet - lKilosNet
        tMuestra = tMuestra - lMuestra
        
        
        SQL = "select count(*) from rclasifica_clasif where numnotac = " & DBSet(Numalbar, "N")
        SQL = SQL & " and codvarie = " & DBSet(RS!codvarie, "N")
        SQL = SQL & " and codcalid = " & DBSet(RS!codcalid, "N")
        
        If TotalRegistros(SQL) = 0 Then
        
            Sql2 = "insert into rclasifica_clasif (numnotac,codvarie,codcalid,muestra,kilosnet) values ("
            Sql2 = Sql2 & DBSet(Numalbar, "N") & ","
            Sql2 = Sql2 & DBSet(RS!codvarie, "N") & ","
            Sql2 = Sql2 & DBSet(RS!codcalid, "N") & ","
            Sql2 = Sql2 & DBSet(lMuestra, "N", "S") & ","
            Sql2 = Sql2 & DBSet(lKilosNet, "N", "S") & ")"
            
        Else
        
            Sql2 = "update rclasifica_clasif set kilosnet = " & DBSet(lKilosNet, "N", "S")
            Sql2 = Sql2 & ",muestra = " & DBSet(lMuestra, "N", "S")
            Sql2 = Sql2 & " where numnotac = " & DBSet(Numalbar, "N")
            Sql2 = Sql2 & " and codvarie = " & DBSet(RS!codvarie, "N")
            Sql2 = Sql2 & " and codcalid = " & DBSet(RS!codcalid, "N")
        
        End If
            
            
        conn.Execute Sql2
    
        RS.MoveNext
    Wend
    
    '[Monica]24/05/2013: si hay alguna calidad con kilos superior a la diferencia de redondeo
    Sql3 = "select min(codcalid) from rclasifica_clasif where numnotac = " & DBSet(Numalbar, "N")
    Sql3 = Sql3 & " and kilosnet >= " & DBSet(tKilosNet * (-1), "N")

    Set rs3 = New ADODB.Recordset
    rs3.Open Sql3, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not rs3.EOF Then
        Calid = DBLet(rs3.Fields(0).Value, "N")
    End If
    Set rs3 = Nothing
    '24/05/2013: hasta aqui
    
    
    ' si no está correcto el totalkilonet el redondeo  va sobre la primera calidad con kilos
    Sql2 = "update rclasifica_clasif set kilosnet = kilosnet + " & DBSet(tKilosNet, "N")
    Sql2 = Sql2 & ",muestra = muestra + " & DBSet(tMuestra, "N")
    Sql2 = Sql2 & " where numnotac = " & DBSet(Numalbar, "N")
    Sql2 = Sql2 & " and codvarie = " & DBSet(Variedad, "N")
    Sql2 = Sql2 & " and codcalid = " & DBSet(Calid, "N")
    
    conn.Execute Sql2
    
    Set RS = Nothing
    
    RepartoClasificacionBascula = True
    Exit Function
    
eRepartoClasificacionBascula:
    Mens = Mens & vbCrLf & "Reparto Clasificación Báscula. " & Err.Description
End Function





Private Function RepartoGastos(AlbAnt As Long, Numalbar As Long, Porcentaje As Currency, Mens As String) As Boolean
Dim SQL As String
Dim RS As ADODB.Recordset
Dim Sql2 As String

Dim lGastos As Single

    On Error GoTo eRepartogastos

    RepartoGastos = False

    SQL = "select * from rhisfruta_gastos where numalbar = " & DBSet(AlbAnt, "N")
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not RS.EOF
        lGastos = Round2(RS!Importe * Porcentaje / 100, 2)
        
        Sql2 = "insert into rhisfruta_gastos (numalbar,numlinea,codgasto,importe) values ("
        Sql2 = Sql2 & DBSet(Numalbar, "N") & ","
        Sql2 = Sql2 & DBSet(RS!numlinea, "N") & ","
        Sql2 = Sql2 & DBSet(RS!Codgasto, "N") & ","
        Sql2 = Sql2 & DBSet(lGastos, "N", "S") & ")"
        
        conn.Execute Sql2
        
        RS.MoveNext
    Wend
    
    Set RS = Nothing
    
    RepartoGastos = True
    Exit Function
    
eRepartogastos:
    Mens = Mens & vbCrLf & "Reparto Gastos. " & Err.Description
End Function



Private Function GrabarIncidencias(AlbAnt As Long, Numalbar As Long, Mens As String) As Boolean
Dim SQL As String
Dim RS As ADODB.Recordset
Dim Sql2 As String

Dim lGastos As Single

    On Error GoTo eGrabarIncidencias

    GrabarIncidencias = False

    SQL = "insert into rhisfruta_incidencia (numalbar,numnotac,codincid) "
    SQL = "select " & DBSet(Numalbar, "N") & ",numnotac, codincid from rhisfruta_incidencia where numalbar = " & DBSet(AlbAnt, "N")
    conn.Execute SQL
        
    GrabarIncidencias = True
    Exit Function
    
eGrabarIncidencias:
    Mens = Mens & vbCrLf & "Grabar Incidencias. " & Err.Description
End Function



Private Function ActualizarEntradas(AlbAnt As Long, cadAlbaran As String, Mens As String) As Boolean
Dim SQL As String
Dim RS As ADODB.Recordset
Dim Sql2 As String

    On Error GoTo eActualizarEntradas

    ActualizarEntradas = False

    SQL = "select numnotac, sum(kilosbru) kilbru, sum(numcajon) numcaj, sum(kilosnet) kilnet, sum(imptrans) imptra, "
    SQL = SQL & " sum(impacarr) impaca, sum(imprecol) imprec, sum(imppenal) imppen "
    SQL = SQL & " from rhisfruta_entradas where numalbar in " & cadAlbaran
    SQL = SQL & " group by 1 order by 1"
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not RS.EOF
        Sql2 = "update rhisfruta_entradas set kilosbru = kilosbru - " & DBSet(RS!kilbru, "N")
        Sql2 = Sql2 & ", numcajon = numcajon - " & DBSet(RS!numcaj, "N")
        Sql2 = Sql2 & ", kilosnet = kilosnet - " & DBSet(RS!kilnet, "N")
        Sql2 = Sql2 & ", imptrans = imptrans - " & DBSet(RS!imptra, "N")
        Sql2 = Sql2 & ", impacarr = impacarr - " & DBSet(RS!impaca, "N")
        Sql2 = Sql2 & ", imprecol = imprecol - " & DBSet(RS!ImpREC, "N")
        Sql2 = Sql2 & ", imppenal = imppenal - " & DBSet(RS!imppen, "N")
        Sql2 = Sql2 & " where numalbar = " & DBSet(AlbAnt, "N")
        Sql2 = Sql2 & " and numnotac = " & DBSet(RS!numnotac, "N")
    
        conn.Execute Sql2
        RS.MoveNext
    Wend
    
    Set RS = Nothing
    
    ActualizarEntradas = True
    Exit Function
    
eActualizarEntradas:
    Mens = Mens & vbCrLf & "Actualizar Entradas. " & Err.Description
End Function



Private Function ActualizarClasificacion(AlbAnt As Long, cadAlbaran As String, Mens As String) As Boolean
Dim SQL As String
Dim RS As ADODB.Recordset
Dim Sql2 As String

Dim lKilosNet As Single
Dim tKilosNet As Single
Dim Calid As Long

    On Error GoTo eActualizarClasificacion

    ActualizarClasificacion = False

    SQL = "select codvarie, codcalid, sum(kilosnet) as kilosnet from rhisfruta_clasif where numalbar in " & cadAlbaran
    SQL = SQL & " group by 1,2 order by 1,2 "
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not RS.EOF
        Sql2 = "update rhisfruta_clasif set kilosnet = kilosnet - " & DBSet(RS!KilosNet, "N")
        Sql2 = Sql2 & " where numalbar = " & DBSet(AlbAnt, "N") & " and codvarie = " & DBSet(RS!codvarie, "N")
        Sql2 = Sql2 & " and codcalid = " & DBSet(RS!codcalid, "N")
        
        conn.Execute Sql2
        
        RS.MoveNext
    Wend
    ' si no está correcto el totalkilonet el redondeo  va sobre la primera calidad con kilos
    Set RS = Nothing
    
    ActualizarClasificacion = True
    Exit Function
    
eActualizarClasificacion:
    Mens = Mens & vbCrLf & "Actualizar Clasificación. " & Err.Description
End Function



Private Function ActualizarGastosAlbaranes(AlbAnt As Long, cadAlbaran As String, Mens As String) As Boolean
Dim SQL As String
Dim RS As ADODB.Recordset
Dim Sql2 As String

Dim lKilosNet As Single
Dim tKilosNet As Single
Dim Calid As Long

    On Error GoTo eActualizarGastosAlbaranes

    ActualizarGastosAlbaranes = False

    SQL = "select numlinea, sum(importe) as importe from rhisfruta_gastos where numalbar in " & cadAlbaran
    SQL = SQL & " group by 1 order by 1 "
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not RS.EOF
        Sql2 = "update rhisfruta_gastos set importe = importe - " & DBSet(RS!Importe, "N")
        Sql2 = Sql2 & " where numalbar = " & DBSet(AlbAnt, "N") & " and numlinea = " & DBSet(RS!numlinea, "N")
        
        conn.Execute Sql2
        
        RS.MoveNext
    Wend
    ' si no está correcto el totalkilonet el redondeo  va sobre la primera calidad con kilos
    Set RS = Nothing
    
    ActualizarGastosAlbaranes = True
    Exit Function
    
eActualizarGastosAlbaranes:
    Mens = Mens & vbCrLf & "Actualizar Gastos Albaranes. " & Err.Description
End Function




'*************************FIN DE REOARTO A COOPROPIETARIOS*******************




Public Function PorCoopropiedadCampo(Campo As String, Socio As String) As Currency
Dim SQL As String
Dim Porcen As Currency

    Porcen = 0
    
    SQL = "select porcentaje from rcampos_cooprop where codcampo = " & DBSet(Campo, "N") & " and codsocio = " & DBSet(Socio, "N")
    
    Porcen = DevuelveValor(SQL)

    PorCoopropiedadCampo = Porcen

End Function
