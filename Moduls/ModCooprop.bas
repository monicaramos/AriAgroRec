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
Dim vKilosTra As Single
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
Dim Rs As ADODB.Recordset
Dim Rs2 As ADODB.Recordset
Dim rs3 As ADODB.Recordset

Dim NumAlbar As Long
Dim vTipoMov As CTiposMov

Dim albaranes As String

Dim tKilosBru As Long
Dim tNumcajon As Long
Dim tKilosNet As Single
Dim tKilosTra As Single
Dim tImpTrans As Single
Dim tImpAcarr As Single
Dim tImpRecol As Single
Dim tImppenal As Single
Dim tImpEntrada As Single
Dim tTarabodega As Single
Dim CodTipoMov As String
Dim B As Boolean
Dim devuelve As String
Dim Existe As Boolean
Dim Mens As String
Dim NroPropiedad As String
Dim NumReg As Long
Dim campo As Long
Dim cTabla As String
Dim NumF As Long



    On Error GoTo eRepartoAlbaranes

    RepartoAlbaranes = False
    
    cadErr = ""
    
    B = True
    
    SQL = "select * from rhisfruta where numalbar = " & DBSet(NumAlbarOrigen, "N")
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not Rs.EOF Then
        
        If TieneCopropietarios(CStr(Rs!codCampo), CStr(Rs!Codsocio)) Then
        
            '[Monica]31/08/2012: reparto de entradas en bodega
            If (vParamAplic.Cooperativa = 3 Or vParamAplic.Cooperativa = 7) And EsVariedadGrupo6(Rs!Codvarie) Then
                cTabla = "(rhisfruta inner join variedades on rhisfruta.codvarie = variedades.codvarie) " & _
                         " inner join productos on variedades.codprodu = productos.codprodu and codgrupo = 6 "
                
                NumF = SugerirCodigoSiguienteStr(cTabla, "numalbar")
                NumF = NumF - 1
                
                NumAlbar = NumF
                
                If NumF < 0 Then B = False
            Else
        
                CodTipoMov = "ALF"
            
                Set vTipoMov = New CTiposMov
                If Not vTipoMov.Leer(CodTipoMov) Then Exit Function

            End If

                albaranes = ""

                tKilosBru = DBLet(Rs!KilosBru, "N")
                tNumcajon = DBLet(Rs!Numcajon, "N")
                tKilosNet = DBLet(Rs!KilosNet, "N")
                tImpTrans = DBLet(Rs!ImpTrans, "N")
                tImpAcarr = DBLet(Rs!impacarr, "N")
                tImpRecol = DBLet(Rs!imprecol, "N")
                tImppenal = DBLet(Rs!ImpPenal, "N")
                tImpEntrada = DBLet(Rs!ImpEntrada, "N")
                tTarabodega = DBLet(Rs!tarabodega, "N")
                '[Monica]29/05/2019: faltaban los kilostra
                tKilosTra = DBLet(Rs!KilosTra, "N")

                Sql2 = "select * from rcampos_cooprop where codcampo = " & DBSet(Rs!codCampo, "N")
                Sql2 = Sql2 & " and rcampos_cooprop.codsocio <> " & DBSet(Rs!Codsocio, "N")
                Sql2 = Sql2 & " order by numlinea "
                
                Set Rs2 = New ADODB.Recordset
                Rs2.Open Sql2, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                
                While Not Rs2.EOF And B
                    
                    If (vParamAplic.Cooperativa = 3 Or vParamAplic.Cooperativa = 7) And EsVariedadGrupo6(Rs!Codvarie) Then
                        
                        NumAlbar = NumAlbar + 1
                    
                    Else
                    
                        NumAlbar = vTipoMov.ConseguirContador(CodTipoMov)
                        Do
                            devuelve = DevuelveDesdeBDNew(cAgro, "rhisfruta", "numalbar", "numalbar", CStr(NumAlbar), "N")
                            If devuelve <> "" Then
                                'Ya existe el contador incrementarlo
                                Existe = True
                                vTipoMov.IncrementarContador (CodTipoMov)
                                NumAlbar = vTipoMov.ConseguirContador(CodTipoMov)
                            Else
                                Existe = False
                            End If
                        Loop Until Not Existe
                    End If
                    
                    vKilosBru = Round2(DBLet(Rs!KilosBru, "N") * DBLet(Rs2!Porcentaje, "N") / 100)
                    vNumcajon = Round2(DBLet(Rs!Numcajon, "N") * DBLet(Rs2!Porcentaje, "N") / 100)
                    vKilosNet = Round2(DBLet(Rs!KilosNet, "N") * DBLet(Rs2!Porcentaje, "N") / 100)
                    '[Monica]29/05/2019: faltaban
                    vKilosTra = Round2(DBLet(Rs!KilosTra, "N") * DBLet(Rs2!Porcentaje, "N") / 100)
                    
                    vImpTrans = Round2(DBLet(Rs!ImpTrans, "N") * DBLet(Rs2!Porcentaje, "N") / 100, 2)
                    vImpAcarr = Round2(DBLet(Rs!impacarr, "N") * DBLet(Rs2!Porcentaje, "N") / 100, 2)
                    vImpRecol = Round2(DBLet(Rs!imprecol, "N") * DBLet(Rs2!Porcentaje, "N") / 100, 2)
                    vImppenal = Round2(DBLet(Rs!ImpPenal, "N") * DBLet(Rs2!Porcentaje, "N") / 100, 2)
                    vImpEntrada = Round2(DBLet(Rs!ImpEntrada, "N") * DBLet(Rs2!Porcentaje, "N") / 100, 2)
                    vTaraBodega = Round2(DBLet(Rs!tarabodega, "N") * DBLet(Rs2!Porcentaje, "N") / 100, 2)
                    
                    tKilosBru = tKilosBru - vKilosBru
                    tNumcajon = tNumcajon - vNumcajon
                    tKilosNet = tKilosNet - vKilosNet
                    '[Monica]29/05/2019: faltaban
                    tKilosTra = tKilosTra - vKilosTra
                    
                    tImpTrans = tImpTrans - vImpTrans
                    tImpAcarr = tImpAcarr - vImpAcarr
                    tImpRecol = tImpRecol - vImpRecol
                    tImppenal = tImppenal - vImppenal
                    tImpEntrada = tImpEntrada - vImpEntrada
                    tTarabodega = tTarabodega - vTaraBodega
                    
                    Sql4 = "insert into rhisfruta (numalbar,fecalbar,codvarie,codsocio,codcampo,tipoentr,recolect,"
                    Sql4 = Sql4 & "kilosbru,numcajon,kilosnet,imptrans,impacarr,imprecol,imppenal,impreso,impentrada,"
                    Sql4 = Sql4 & "cobradosn,prestimado,coddeposito,codpobla,transportadopor,albarorigen, tarabodega,codtraba,tolva,kgradobonif,esbonifespecial, kilostra) values ("
                    Sql4 = Sql4 & DBSet(NumAlbar, "N") & ","
                    Sql4 = Sql4 & DBSet(Rs!Fecalbar, "F") & ","
                    Sql4 = Sql4 & DBSet(Rs!Codvarie, "N") & ","
                    Sql4 = Sql4 & DBSet(Rs2!Codsocio, "N") & ","
                    Sql4 = Sql4 & DBSet(Rs!codCampo, "N") & ","
                    Sql4 = Sql4 & DBSet(Rs!TipoEntr, "N") & ","
                    Sql4 = Sql4 & DBSet(Rs!Recolect, "N") & ","
                    Sql4 = Sql4 & DBSet(vKilosBru, "N") & ","
                    Sql4 = Sql4 & DBSet(vNumcajon, "N") & ","
                    Sql4 = Sql4 & DBSet(vKilosNet, "N") & ","
                    Sql4 = Sql4 & DBSet(vImpTrans, "N", "S") & ","
                    Sql4 = Sql4 & DBSet(vImpAcarr, "N", "S") & ","
                    Sql4 = Sql4 & DBSet(vImpRecol, "N", "S") & ","
                    Sql4 = Sql4 & DBSet(vImppenal, "N", "S") & ","
                    Sql4 = Sql4 & DBSet(Rs!impreso, "N") & ","
                    Sql4 = Sql4 & DBSet(vImpEntrada, "N") & ","
                    Sql4 = Sql4 & DBSet(Rs!cobradosn, "N") & ","
                    Sql4 = Sql4 & DBSet(Rs!PrEstimado, "N", "S") & ","
                    
                    '[Monica]31/08/2012: si venimos  de una entrada de bodega
                    If EsVariedadGrupo6(Rs!Codvarie) Then
                        Sql4 = Sql4 & DBSet(Rs!coddeposito, "N") & ","
                    Else
                        Sql4 = Sql4 & DBSet(Rs!coddeposito, "N", "S") & ","
                    End If
                    
                    Sql4 = Sql4 & DBSet(Rs!CodPobla, "N", "S") & ","
                    Sql4 = Sql4 & DBSet(Rs!transportadopor, "N") & ","
                    Sql4 = Sql4 & DBSet(NumAlbarOrigen, "N") & ","
                    Sql4 = Sql4 & DBSet(vTaraBodega, "N") & ","
                    Sql4 = Sql4 & DBSet(Rs!CodTraba, "N") & ","
                    Sql4 = Sql4 & DBSet(Rs!tolva, "N") & ","
                    Sql4 = Sql4 & DBSet(Rs!kgradobonif, "N") & ","
                    Sql4 = Sql4 & DBSet(Rs!esbonifespecial, "N") & "," ' ")"
                    '[Monica]29/05/2019: faltan kilostra
                    Sql4 = Sql4 & DBSet(vKilosTra, "N") & ")"
                    
                    conn.Execute Sql4
                    
                    Mens = "Reparto de Entradas."
                    If B Then B = RepartoEntradas(NumAlbarOrigen, NumAlbar, DBLet(Rs2!Porcentaje, "N"), Mens)
                
                    Mens = "Reparto de Clasificación."
                    If B Then B = RepartoClasificacion(NumAlbarOrigen, NumAlbar, DBLet(Rs2!Porcentaje, "N"), Mens)
                    
                    Mens = "Reparto de Gastos."
                    If B Then B = RepartoGastos(NumAlbarOrigen, NumAlbar, DBLet(Rs2!Porcentaje, "N"), Mens)
                    
                    Mens = "Grabar Incidencias."
                    If B Then B = GrabarIncidencias(NumAlbarOrigen, NumAlbar, Mens)
                
                    albaranes = albaranes & NumAlbar & ","
                    
                
                    Rs2.MoveNext
                Wend
                
                Set Rs2 = Nothing
                
                If B Then
                    ' ultimo registro la diferencia ( se updatean las tablas del registro de rhisfruta origen )
                    Sql4 = "update rhisfruta set kilosbru = " & DBSet(tKilosBru, "N") & ","
                    Sql4 = Sql4 & "numcajon = " & DBSet(tNumcajon, "N") & ","
                    Sql4 = Sql4 & "kilosnet = " & DBSet(tKilosNet, "N") & ","
                    '[Monica]29/05/2019
                    Sql4 = Sql4 & "kilostra = " & DBSet(tKilosTra, "N") & ","
                    
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
                
                    albaranes = "(" & Mid(albaranes, 1, Len(albaranes) - 1) & ")"
                
                    Mens = "Actualizar Entradas."
                    If B Then B = ActualizarEntradas(NumAlbarOrigen, albaranes, Mens)
                
                    Mens = "Actualizar Clasificación."
                    If B Then B = ActualizarClasificacion(NumAlbarOrigen, albaranes, Mens)
                    
                    Mens = "Actualizara Gastos."
                    If B Then B = ActualizarGastosAlbaranes(NumAlbarOrigen, albaranes, Mens)
                
                    If Not ((vParamAplic.Cooperativa = 3 Or vParamAplic.Cooperativa = 7) And EsVariedadGrupo6(Rs!Codvarie)) Then
                
                        vTipoMov.IncrementarContador (CodTipoMov)
                    
                    End If
                    
                ' fin de ultimo registro
                End If
'            Else
'                b = False
'            End If
        
        End If
    
    End If
    
    Set Rs = Nothing

eRepartoAlbaranes:
    If Err.Number <> 0 Or Not B Then
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
Dim SQL1 As String
Dim Sql2 As String
Dim Sql4 As String
Dim Rs As ADODB.Recordset
Dim Rs2 As ADODB.Recordset
Dim rs3 As ADODB.Recordset

Dim NumAlbar As Long
Dim vTipoMov As CTiposMov

Dim albaranes As String

Dim tKilosBru As Long
Dim tNumcajon As Long
Dim tKilosNet As Single
Dim tImpTrans As Single
Dim tImpAcarr As Single
Dim tImpRecol As Single
Dim tImppenal As Single
Dim tImpEntrada As Single
Dim CodTipoMov As String
Dim B As Boolean
Dim devuelve As String
Dim Existe As Boolean
Dim Mens As String
Dim NroPropiedad As String
Dim NumReg As Long
Dim campo As Long
Dim Porcentaje As Single

    On Error GoTo eRepartoAlbaranesBascula

    RepartoAlbaranesBascula = False
    
    cadErr = ""
    
    B = True
    
    SQL = "select * from rclasifica where numnotac = " & DBSet(NumAlbarOrigen, "N")
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not Rs.EOF Then
        
        If TieneCopropietarios(CStr(Rs!codCampo), CStr(Rs!Codsocio)) Then
            CodTipoMov = "NOC"
        
            Set vTipoMov = New CTiposMov
            If vTipoMov.Leer(CodTipoMov) Then

                albaranes = ""

                tKilosBru = DBLet(Rs!KilosBru, "N")
                tNumcajon = DBLet(Rs!Numcajon, "N")
                tKilosNet = DBLet(Rs!KilosNet, "N")
                tImpTrans = DBLet(Rs!ImpTrans, "N")
                tImpAcarr = DBLet(Rs!impacarr, "N")
                tImpRecol = DBLet(Rs!imprecol, "N")
                tImppenal = DBLet(Rs!ImpPenal, "N")

                Sql2 = "select * from rcampos_cooprop where codcampo = " & DBSet(Rs!codCampo, "N")
                Sql2 = Sql2 & " and rcampos_cooprop.codsocio <> " & DBSet(Rs!Codsocio, "N")
                Sql2 = Sql2 & " order by numlinea "
                
                Set Rs2 = New ADODB.Recordset
                Rs2.Open Sql2, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                
                While Not Rs2.EOF And B
                    NumAlbar = vTipoMov.ConseguirContador(CodTipoMov)
                    Do
                        devuelve = DevuelveDesdeBDNew(cAgro, "rclasifica", "numnotac", "numnotac", CStr(NumAlbar), "N")
                        If devuelve <> "" Then
                            'Ya existe el contador incrementarlo
                            Existe = True
                            vTipoMov.IncrementarContador (CodTipoMov)
                            NumAlbar = vTipoMov.ConseguirContador(CodTipoMov)
                        Else
                            Existe = False
                        End If
                    Loop Until Not Existe
                    
                    vKilosBru = Round2(DBLet(Rs!KilosBru, "N") * DBLet(Rs2!Porcentaje, "N") / 100)
                    vNumcajon = Round2(DBLet(Rs!Numcajon, "N") * DBLet(Rs2!Porcentaje, "N") / 100)
                    vKilosNet = Round2(DBLet(Rs!KilosNet, "N") * DBLet(Rs2!Porcentaje, "N") / 100)
                    vImpTrans = Round2(DBLet(Rs!ImpTrans, "N") * DBLet(Rs2!Porcentaje, "N") / 100, 2)
                    vImpAcarr = Round2(DBLet(Rs!impacarr, "N") * DBLet(Rs2!Porcentaje, "N") / 100, 2)
                    vImpRecol = Round2(DBLet(Rs!imprecol, "N") * DBLet(Rs2!Porcentaje, "N") / 100, 2)
                    vImppenal = Round2(DBLet(Rs!ImpPenal, "N") * DBLet(Rs2!Porcentaje, "N") / 100, 2)
                    
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
                    Sql4 = Sql4 & DBSet(NumAlbar, "N") & ","
                    Sql4 = Sql4 & DBSet(Rs!FechaEnt, "F") & ","
                    Sql4 = Sql4 & DBSet(Rs!horaentr, "FH") & ","
                    Sql4 = Sql4 & DBSet(Rs!Codvarie, "N") & ","
                    Sql4 = Sql4 & DBSet(Rs2!Codsocio, "N") & ","
                    Sql4 = Sql4 & DBSet(Rs!codCampo, "N") & ","
                    Sql4 = Sql4 & DBSet(Rs!TipoEntr, "N") & ","
                    Sql4 = Sql4 & DBSet(Rs!Recolect, "N") & ","
                    Sql4 = Sql4 & DBSet(Rs!codTrans, "N") & ","
                    Sql4 = Sql4 & DBSet(Rs!codcapat, "N") & ","
                    Sql4 = Sql4 & DBSet(Rs!Codtarif, "N") & ","
                    Sql4 = Sql4 & DBSet(vKilosBru, "N") & ","
                    Sql4 = Sql4 & DBSet(vNumcajon, "N") & ","
                    Sql4 = Sql4 & DBSet(vKilosNet, "N") & ","
                    Sql4 = Sql4 & DBSet(Rs!Observac, "T") & ","
                    Sql4 = Sql4 & DBSet(vImpTrans, "N", "S") & ","
                    Sql4 = Sql4 & DBSet(vImpAcarr, "N", "S") & ","
                    Sql4 = Sql4 & DBSet(vImpRecol, "N", "S") & ","
                    Sql4 = Sql4 & DBSet(vImppenal, "N", "S") & ","
                    Sql4 = Sql4 & DBSet(Rs!tiporecol, "N") & ","
                    Sql4 = Sql4 & DBSet(Rs!horastra, "N", "S") & ","
                    Sql4 = Sql4 & DBSet(Rs!numtraba, "N", "S") & ","
                    Sql4 = Sql4 & DBSet(Rs!NumAlbar, "N", "S") & ","
                    Sql4 = Sql4 & DBSet(Rs!Fecalbar, "F", "S") & ","
                    Sql4 = Sql4 & DBSet(Rs!impreso, "N") & ","
                    Sql4 = Sql4 & DBSet(Rs!PrEstimado, "N", "S") & ","
                    Sql4 = Sql4 & DBSet(Rs!transportadopor, "N") & ","
                    Sql4 = Sql4 & DBSet(Rs!KilosTra, "N") & ")"
                    
                    conn.Execute Sql4
                    
                    '[Monica]25/03/2014: en el caso de que en la origen hayan plagas la copiamos
                    '                    solo puede ocurrir en el caso de quatretonda pq en el mto de entradas bascula indicamos sin ausencia de plagas
                    SQL = "select * from rclasifica_incidencia where numnotac = " & DBSet(NumAlbarOrigen, "N")
                    If TotalRegistros(SQL) <> 0 Then
                        Sql4 = "insert into rclasifica_incidencia (numnotac, codincid) select " & DBSet(NumAlbar, "N") & ", codincid "
                        Sql4 = Sql4 & " from rclasifica_incidencia where numnotac = " & DBSet(NumAlbarOrigen, "N")
                    
                        conn.Execute Sql4
                    End If
                    
                    Mens = "Reparto de Clasificación Báscula."
                    If B Then B = RepartoClasificacionBascula(NumAlbarOrigen, NumAlbar, DBLet(Rs2!Porcentaje, "N"), Mens)
                    
                    albaranes = albaranes & NumAlbar & ","
                
                    Rs2.MoveNext
                Wend
                
                Set Rs2 = Nothing
                
                If B Then
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
                    Sql2 = "select porcentaje from rcampos_cooprop where codcampo = " & DBSet(Rs!codCampo, "N")
                    Sql2 = Sql2 & " and rcampos_cooprop.codsocio = " & DBSet(Rs!Codsocio, "N")
                    Porcentaje = DevuelveValor(Sql2)

                    Mens = "Reparto de Clasificación Báscula."
                    If B Then B = RepartoClasificacionBascula(NumAlbarOrigen, NumAlbarOrigen, DBLet(Porcentaje, "N"), Mens)
                
                    albaranes = "(" & Mid(albaranes, 1, Len(albaranes) - 1) & ")"
                
                    vTipoMov.IncrementarContador (CodTipoMov)
                ' fin de ultimo registro
                End If
            Else
                B = False
            End If
        
        End If
    
    End If
    
    Set Rs = Nothing

eRepartoAlbaranesBascula:
    If Err.Number <> 0 Or Not B Then
        cadErr = cadErr & Err.Description
    Else
        RepartoAlbaranesBascula = True
    End If
End Function


Private Function TieneCopropietarios(campo As String, Propietario As String) As Boolean
Dim NroCampo As String
Dim SQL As String

    SQL = "select count(*) from rcampos_cooprop where codcampo = " & DBSet(campo, "N") & " and codsocio <> " & DBSet(Propietario, "N")
    
    TieneCopropietarios = TotalRegistros(SQL) > 0

End Function


Private Function RepartoEntradas(AlbAnt As Long, NumAlbar As Long, Porcentaje As Currency, Mens As String) As Boolean
Dim SQL As String
Dim Rs As ADODB.Recordset
Dim Sql2 As String

Dim lKilosBru As Long
Dim lNumcajon As Long
Dim lKilosNet As Single
Dim lKilosTra As Single

Dim lImpTrans As Single
Dim lImpAcarr As Single
Dim lImpRecol As Single
Dim lImppenal As Single
Dim lImpEntrada As Single
    
Dim tKilosBru As Long
Dim tNumcajon As Long
Dim tKilosNet As Single
Dim tKilosTra As Single
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
    '[Monica]29/05/2019
    tKilosTra = vKilosTra

    SQL = "select * from rhisfruta_entradas where numalbar = " & DBSet(AlbAnt, "N")
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not Rs.EOF
        lKilosBru = Round2(DBLet(Rs!KilosBru, "N") * Porcentaje / 100)
        lNumcajon = Round2(DBLet(Rs!Numcajon, "N") * Porcentaje / 100)
        lKilosNet = Round2(DBLet(Rs!KilosNet, "N") * Porcentaje / 100)
        '[Monica]29/05/2019
        lKilosTra = Round2(DBLet(Rs!KilosTra, "N") * Porcentaje / 100)
        
        lImpTrans = Round2(DBLet(Rs!ImpTrans, "N") * Porcentaje / 100, 2)
        lImpAcarr = Round2(DBLet(Rs!impacarr, "N") * Porcentaje / 100, 2)
        lImpRecol = Round2(DBLet(Rs!imprecol, "N") * Porcentaje / 100, 2)
        lImppenal = Round2(DBLet(Rs!ImpPenal, "N") * Porcentaje / 100, 2)
        
        tKilosBru = tKilosBru - lKilosBru
        tNumcajon = tNumcajon - lNumcajon
        tKilosNet = tKilosNet - lKilosNet
        '[Monica]29/05/2019
        tKilosTra = tKilosTra - lKilosTra
        tImpTrans = tImpTrans - lImpTrans
        tImpAcarr = tImpAcarr - lImpAcarr
        tImpRecol = tImpRecol - lImpRecol
        tImppenal = tImppenal - lImppenal
       
        NumNota = Rs!NumNotac
        
        Sql2 = "insert into rhisfruta_entradas (numalbar,numnotac,fechaent,horaentr,kilosbru,numcajon,kilosnet,observac,imptrans,"
        Sql2 = Sql2 & "impacarr,imprecol,imppenal,prestimado,codtrans,codtarif,codcapat,kilostra) values ("
        Sql2 = Sql2 & DBSet(NumAlbar, "N") & ","
        Sql2 = Sql2 & DBSet(Rs!NumNotac, "N") & ","
        Sql2 = Sql2 & DBSet(Rs!FechaEnt, "F") & ","
        Sql2 = Sql2 & DBSet(Rs!horaentr, "FH") & ","
        Sql2 = Sql2 & DBSet(lKilosBru, "N") & ","
        Sql2 = Sql2 & DBSet(lNumcajon, "N") & ","
        Sql2 = Sql2 & DBSet(lKilosNet, "N") & ","
        Sql2 = Sql2 & DBSet(Rs!Observac, "T", "S") & ","
        Sql2 = Sql2 & DBSet(lImpTrans, "N", "S") & ","
        Sql2 = Sql2 & DBSet(lImpAcarr, "N", "S") & ","
        Sql2 = Sql2 & DBSet(lImpRecol, "N", "S") & ","
        Sql2 = Sql2 & DBSet(lImppenal, "N", "S") & ","
        Sql2 = Sql2 & DBSet(Rs!PrEstimado, "N", "S") & ","
        Sql2 = Sql2 & DBSet(Rs!codTrans, "T", "S") & ","
        Sql2 = Sql2 & DBSet(Rs!Codtarif, "N", "S") & ","
        Sql2 = Sql2 & DBSet(Rs!codcapat, "N", "S") & "," '")"
        Sql2 = Sql2 & DBSet(lKilosTra, "N") & ")"
        
        conn.Execute Sql2
    
        Rs.MoveNext
    Wend
    
    Sql2 = "update rhisfruta_entradas set kilosbru = kilosbru + " & DBSet(tKilosBru, "N")
    Sql2 = Sql2 & ", numcajon = numcajon + " & DBSet(tNumcajon, "N")
    Sql2 = Sql2 & ", kilosnet = kilosnet + " & DBSet(tKilosNet, "N")
    '[Monica]29/05/2019
    Sql2 = Sql2 & ", kilostra = kilostra + " & DBSet(tKilosTra, "N")
    
    Sql2 = Sql2 & ", imptrans = imptrans + " & DBSet(tImpTrans, "N")
    Sql2 = Sql2 & ", impacarr = impacarr + " & DBSet(tImpAcarr, "N")
    Sql2 = Sql2 & ", imprecol = imprecol + " & DBSet(tImpAcarr, "N")
    Sql2 = Sql2 & ", imppenal = imppenal + " & DBSet(tImppenal, "N")
    Sql2 = Sql2 & " where numalbar = " & DBSet(NumAlbar, "N") & " and numnotac = " & DBSet(NumNota, "N")
    
    conn.Execute Sql2
    
    Set Rs = Nothing
    
    RepartoEntradas = True
    Exit Function
    
    
eRepartoEntradas:
    Mens = Mens & vbCrLf & "Reparto Entradas. " & Err.Description
End Function



Private Function RepartoClasificacion(AlbAnt As Long, NumAlbar As Long, Porcentaje As Currency, Mens As String) As Boolean
Dim SQL As String
Dim Rs As ADODB.Recordset
Dim Sql2 As String

Dim lKilosNet As Single
Dim tKilosNet As Single

Dim lKilosTra As Single
Dim tKilosTra As Single

Dim Calid As Long
Dim rs3 As ADODB.Recordset
Dim Sql3 As String

    On Error GoTo eRepartoClasificacion

    RepartoClasificacion = False


    tKilosNet = vKilosNet
    '[Monica]29/05/2019: faltaban los kilostra
    tKilosTra = vKilosTra
    
    Calid = 0
    
    SQL = "select * from rhisfruta_clasif where numalbar = " & DBSet(AlbAnt, "N")
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not Rs.EOF
        lKilosNet = Round2(DBLet(Rs!KilosNet, "N") * Porcentaje / 100)
        '[Monica]29/05/2019
        lKilosTra = Round2(DBLet(Rs!KilosTra, "N") * Porcentaje / 100)
        
        If (lKilosNet <> 0 Or lKilosTra <> 0) And Calid = 0 Then
            Calid = DBLet(Rs!codcalid, "N")
        End If
        
        tKilosNet = tKilosNet - lKilosNet
        '[Monica]29/05/2019
        tKilosTra = tKilosTra - lKilosTra
        
        
        Sql2 = "insert into rhisfruta_clasif (numalbar,codvarie,codcalid,kilosnet, kilostra) values ("
        Sql2 = Sql2 & DBSet(NumAlbar, "N") & ","
        Sql2 = Sql2 & DBSet(Rs!Codvarie, "N") & ","
        Sql2 = Sql2 & DBSet(Rs!codcalid, "N") & ","
        Sql2 = Sql2 & DBSet(lKilosNet, "N", "S") & "," '")"
        Sql2 = Sql2 & DBSet(lKilosTra, "N", "S") & ")"
        
        conn.Execute Sql2
    
        Rs.MoveNext
    Wend
    
    '[Monica]24/05/2013: si hay alguna calidad con kilos superior a la diferencia de redondeo
    Sql3 = "select min(codcalid) from rhisfruta_clasif where numalbar = " & DBSet(NumAlbar, "N")
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
    Sql2 = Sql2 & " where numalbar = " & DBSet(NumAlbar, "N") & " and codcalid = " & DBSet(Calid, "N")
    
    conn.Execute Sql2
    
    
    
    '[Monica]29/05/2019: idem para los kilostra
    Sql3 = "select min(codcalid) from rhisfruta_clasif where numalbar = " & DBSet(NumAlbar, "N")
    Sql3 = Sql3 & " and kilostra >= " & DBSet(tKilosTra * (-1), "N")
    Set rs3 = New ADODB.Recordset
    rs3.Open Sql3, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not rs3.EOF Then
        Calid = DBLet(rs3.Fields(0).Value, "N")
    End If
    Set rs3 = Nothing
    
    ' si no está correcto el totalkilonet el redondeo  va sobre la primera calidad con kilos
    Sql2 = "update rhisfruta_clasif set kilostra = kilostra + " & DBSet(tKilosTra, "N")
    Sql2 = Sql2 & " where numalbar = " & DBSet(NumAlbar, "N") & " and codcalid = " & DBSet(Calid, "N")
    
    conn.Execute Sql2
    
    
    
    Set Rs = Nothing
    
    RepartoClasificacion = True
    Exit Function
    
eRepartoClasificacion:
    Mens = Mens & vbCrLf & "Reparto Clasificación. " & Err.Description
End Function


Private Function RepartoClasificacionBascula(AlbAnt As Long, NumAlbar As Long, Porcentaje As Currency, Mens As String) As Boolean
Dim SQL As String
Dim Rs As ADODB.Recordset
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
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not Rs.EOF
        Variedad = DBLet(Rs!Codvarie, "N")
        
        lKilosNet = Round2(DBLet(Rs!KilosNet, "N") * Porcentaje / 100)
        lMuestra = Round2(DBLet(Rs!Muestra, "N") * Porcentaje / 100)
        If lKilosNet <> 0 And Calid = 0 Then
            Calid = DBLet(Rs!codcalid, "N")
        End If
        
        tKilosNet = tKilosNet - lKilosNet
        tMuestra = tMuestra - lMuestra
        
        
        SQL = "select count(*) from rclasifica_clasif where numnotac = " & DBSet(NumAlbar, "N")
        SQL = SQL & " and codvarie = " & DBSet(Rs!Codvarie, "N")
        SQL = SQL & " and codcalid = " & DBSet(Rs!codcalid, "N")
        
        If TotalRegistros(SQL) = 0 Then
        
            Sql2 = "insert into rclasifica_clasif (numnotac,codvarie,codcalid,muestra,kilosnet) values ("
            Sql2 = Sql2 & DBSet(NumAlbar, "N") & ","
            Sql2 = Sql2 & DBSet(Rs!Codvarie, "N") & ","
            Sql2 = Sql2 & DBSet(Rs!codcalid, "N") & ","
            Sql2 = Sql2 & DBSet(lMuestra, "N", "S") & ","
            Sql2 = Sql2 & DBSet(lKilosNet, "N", "S") & ")"
            
        Else
        
            Sql2 = "update rclasifica_clasif set kilosnet = " & DBSet(lKilosNet, "N", "S")
            Sql2 = Sql2 & ",muestra = " & DBSet(lMuestra, "N", "S")
            Sql2 = Sql2 & " where numnotac = " & DBSet(NumAlbar, "N")
            Sql2 = Sql2 & " and codvarie = " & DBSet(Rs!Codvarie, "N")
            Sql2 = Sql2 & " and codcalid = " & DBSet(Rs!codcalid, "N")
        
        End If
            
            
        conn.Execute Sql2
    
        Rs.MoveNext
    Wend
    
    '[Monica]24/05/2013: si hay alguna calidad con kilos superior a la diferencia de redondeo
    Sql3 = "select min(codcalid) from rclasifica_clasif where numnotac = " & DBSet(NumAlbar, "N")
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
    Sql2 = Sql2 & " where numnotac = " & DBSet(NumAlbar, "N")
    Sql2 = Sql2 & " and codvarie = " & DBSet(Variedad, "N")
    Sql2 = Sql2 & " and codcalid = " & DBSet(Calid, "N")
    
    conn.Execute Sql2
    
    Set Rs = Nothing
    
    RepartoClasificacionBascula = True
    Exit Function
    
eRepartoClasificacionBascula:
    Mens = Mens & vbCrLf & "Reparto Clasificación Báscula. " & Err.Description
End Function





Private Function RepartoGastos(AlbAnt As Long, NumAlbar As Long, Porcentaje As Currency, Mens As String) As Boolean
Dim SQL As String
Dim Rs As ADODB.Recordset
Dim Sql2 As String

Dim lGastos As Single

    On Error GoTo eRepartogastos

    RepartoGastos = False

    SQL = "select * from rhisfruta_gastos where numalbar = " & DBSet(AlbAnt, "N")
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not Rs.EOF
        lGastos = Round2(Rs!Importe * Porcentaje / 100, 2)
        
        Sql2 = "insert into rhisfruta_gastos (numalbar,numlinea,codgasto,importe) values ("
        Sql2 = Sql2 & DBSet(NumAlbar, "N") & ","
        Sql2 = Sql2 & DBSet(Rs!NumLinea, "N") & ","
        Sql2 = Sql2 & DBSet(Rs!Codgasto, "N") & ","
        Sql2 = Sql2 & DBSet(lGastos, "N", "S") & ")"
        
        conn.Execute Sql2
        
        Rs.MoveNext
    Wend
    
    Set Rs = Nothing
    
    RepartoGastos = True
    Exit Function
    
eRepartogastos:
    Mens = Mens & vbCrLf & "Reparto Gastos. " & Err.Description
End Function



Private Function GrabarIncidencias(AlbAnt As Long, NumAlbar As Long, Mens As String) As Boolean
Dim SQL As String
Dim Rs As ADODB.Recordset
Dim Sql2 As String

Dim lGastos As Single

    On Error GoTo eGrabarIncidencias

    GrabarIncidencias = False

    SQL = "insert into rhisfruta_incidencia (numalbar,numnotac,codincid) "
    SQL = "select " & DBSet(NumAlbar, "N") & ",numnotac, codincid from rhisfruta_incidencia where numalbar = " & DBSet(AlbAnt, "N")
    conn.Execute SQL
        
    GrabarIncidencias = True
    Exit Function
    
eGrabarIncidencias:
    Mens = Mens & vbCrLf & "Grabar Incidencias. " & Err.Description
End Function



Private Function ActualizarEntradas(AlbAnt As Long, cadAlbaran As String, Mens As String) As Boolean
Dim SQL As String
Dim Rs As ADODB.Recordset
Dim Sql2 As String

    On Error GoTo eActualizarEntradas

    ActualizarEntradas = False

    SQL = "select numnotac, sum(kilosbru) kilbru, sum(numcajon) numcaj, sum(kilosnet) kilnet, sum(imptrans) imptra, "
    SQL = SQL & " sum(impacarr) impaca, sum(imprecol) imprec, sum(imppenal) imppen, sum(kilostra) kiltra "
    SQL = SQL & " from rhisfruta_entradas where numalbar in " & cadAlbaran
    SQL = SQL & " group by 1 order by 1"
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not Rs.EOF
        Sql2 = "update rhisfruta_entradas set kilosbru = kilosbru - " & DBSet(Rs!kilbru, "N")
        Sql2 = Sql2 & ", numcajon = numcajon - " & DBSet(Rs!numcaj, "N")
        Sql2 = Sql2 & ", kilosnet = kilosnet - " & DBSet(Rs!kilnet, "N")
        '[Monica]29/05/2019:faltaba
        Sql2 = Sql2 & ", kilostra = kilostra - " & DBSet(Rs!kiltra, "N")
        Sql2 = Sql2 & ", imptrans = imptrans - " & DBSet(Rs!imptra, "N")
        Sql2 = Sql2 & ", impacarr = impacarr - " & DBSet(Rs!impaca, "N")
        Sql2 = Sql2 & ", imprecol = imprecol - " & DBSet(Rs!ImpREC, "N")
        Sql2 = Sql2 & ", imppenal = imppenal - " & DBSet(Rs!imppen, "N")
        Sql2 = Sql2 & " where numalbar = " & DBSet(AlbAnt, "N")
        Sql2 = Sql2 & " and numnotac = " & DBSet(Rs!NumNotac, "N")
    
        conn.Execute Sql2
        Rs.MoveNext
    Wend
    
    Set Rs = Nothing
    
    ActualizarEntradas = True
    Exit Function
    
eActualizarEntradas:
    Mens = Mens & vbCrLf & "Actualizar Entradas. " & Err.Description
End Function



Private Function ActualizarClasificacion(AlbAnt As Long, cadAlbaran As String, Mens As String) As Boolean
Dim SQL As String
Dim Rs As ADODB.Recordset
Dim Sql2 As String

Dim lKilosNet As Single
Dim tKilosNet As Single
Dim Calid As Long

    On Error GoTo eActualizarClasificacion

    ActualizarClasificacion = False

    SQL = "select codvarie, codcalid, sum(kilosnet) as kilosnet, sum(kilostra) as kilostra from rhisfruta_clasif where numalbar in " & cadAlbaran
    SQL = SQL & " group by 1,2 order by 1,2 "
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not Rs.EOF
        Sql2 = "update rhisfruta_clasif set kilosnet = kilosnet - " & DBSet(Rs!KilosNet, "N")
        '[Monica]29/05/2019: faltabla
        Sql2 = Sql2 & ", kilostra = kilostra - " & DBSet(Rs!KilosTra, "N")
        Sql2 = Sql2 & " where numalbar = " & DBSet(AlbAnt, "N") & " and codvarie = " & DBSet(Rs!Codvarie, "N")
        Sql2 = Sql2 & " and codcalid = " & DBSet(Rs!codcalid, "N")
        
        conn.Execute Sql2
        
        Rs.MoveNext
    Wend
    ' si no está correcto el totalkilonet el redondeo  va sobre la primera calidad con kilos
    Set Rs = Nothing
    
    ActualizarClasificacion = True
    Exit Function
    
eActualizarClasificacion:
    Mens = Mens & vbCrLf & "Actualizar Clasificación. " & Err.Description
End Function



Private Function ActualizarGastosAlbaranes(AlbAnt As Long, cadAlbaran As String, Mens As String) As Boolean
Dim SQL As String
Dim Rs As ADODB.Recordset
Dim Sql2 As String

Dim lKilosNet As Single
Dim tKilosNet As Single
Dim Calid As Long

    On Error GoTo eActualizarGastosAlbaranes

    ActualizarGastosAlbaranes = False

    SQL = "select numlinea, sum(importe) as importe from rhisfruta_gastos where numalbar in " & cadAlbaran
    SQL = SQL & " group by 1 order by 1 "
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not Rs.EOF
        Sql2 = "update rhisfruta_gastos set importe = importe - " & DBSet(Rs!Importe, "N")
        Sql2 = Sql2 & " where numalbar = " & DBSet(AlbAnt, "N") & " and numlinea = " & DBSet(Rs!NumLinea, "N")
        
        conn.Execute Sql2
        
        Rs.MoveNext
    Wend
    ' si no está correcto el totalkilonet el redondeo  va sobre la primera calidad con kilos
    Set Rs = Nothing
    
    ActualizarGastosAlbaranes = True
    Exit Function
    
eActualizarGastosAlbaranes:
    Mens = Mens & vbCrLf & "Actualizar Gastos Albaranes. " & Err.Description
End Function




'*************************FIN DE REOARTO A COOPROPIETARIOS*******************




Public Function PorCoopropiedadCampo(campo As String, Socio As String) As Currency
Dim SQL As String
Dim Porcen As Currency

    Porcen = 0
    
    SQL = "select porcentaje from rcampos_cooprop where codcampo = " & DBSet(campo, "N") & " and codsocio = " & DBSet(Socio, "N")
    
    Porcen = DevuelveValor(SQL)

    PorCoopropiedadCampo = Porcen

End Function
