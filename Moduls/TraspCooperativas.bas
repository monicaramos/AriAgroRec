Attribute VB_Name = "TraspCooperativas"
Option Explicit

'#########################################################################################################
'
'################### MODULO CON LAS FUNCIONES NECESARIAS PARA COMUNICACION ENTRE COOPIC Y PICASSENT
'
'#########################################################################################################


Public Function ComunicaCooperativa(vtabla As String, vSQL As String, vOperacion As String, Optional vObservaciones As String) As Boolean
' vOperacion: I insercion
'             U modificacion
Dim Sql As String
Dim vInsert As String
Dim vValues As String
Dim vCad As String

    On Error GoTo eComunicaCooperativa
    
    ComunicaCooperativa = False
        
    vCad = SaltosDeLinea(vSQL)
    vSQL = vCad
        
    Sql = "INSERT INTO comunica_env (fechacreacion,usuariocreacion,tipo,tabla,sqlaejecutar,  "
    Sql = Sql & "observaciones,fechadescarga,usuariodescarga) VALUES ("
    Sql = Sql & DBSet(Now, "FH") & "," & DBSet(vUsu.Login, "T") & "," & DBSet(vOperacion, "T") & "," & DBSet(vtabla, "T") & ","
    Sql = Sql & DBSet(vSQL, "T") & "," & DBSet(vObservaciones, "T", "S") & "," & ValorNulo & "," & ValorNulo & ")"
    
    conn.Execute Sql
    
    
    ComunicaCooperativa = True
    Exit Function
    
eComunicaCooperativa:
    MuestraError Err.Number, "Comunica cooperativa", Err.Description
End Function


Public Function EsSocioCooperativa(vSoc As String) As Boolean
Dim Sql As String

    EsSocioCooperativa = True
    If vParamAplic.Cooperativa <> 2 And vParamAplic.Cooperativa <> 16 Then Exit Function
    
    EsSocioCooperativa = (CLng(ComprobarCero(vSoc)) < cMaxSocio)

End Function

Public Function EsCampoCooperativa(vCam As String) As Boolean
Dim Sql As String

    EsCampoCooperativa = True
    If vParamAplic.Cooperativa <> 2 And vParamAplic.Cooperativa <> 16 Then Exit Function
    
    EsCampoCooperativa = (CLng(ComprobarCero(vCam)) < cMaxCampo)

End Function

Public Function EsCapatazCooperativa(vCap As String) As Boolean
Dim Sql As String

    EsCapatazCooperativa = True
    If vParamAplic.Cooperativa <> 2 And vParamAplic.Cooperativa <> 16 Then Exit Function
    
    EsCapatazCooperativa = (CLng(ComprobarCero(vCap)) < cMaxCapa)

End Function

Public Function EsTransportistaCooperativa(vTra As String) As Boolean
Dim Sql As String
Dim vCar As String

    EsTransportistaCooperativa = True
    If vParamAplic.Cooperativa <> 2 And vParamAplic.Cooperativa <> 16 Then Exit Function
    
    vCar = Mid(vTra, 1, 1)
    
    EsTransportistaCooperativa = ((vParamAplic.Cooperativa = 2 And vCar <> "C") Or (vParamAplic.Cooperativa = 16 And vCar <> "A"))

End Function


Public Function EsVariedadComercializada(vCodvarie As String) As Boolean
Dim Sql As String

    Sql = "select count(*) from variedades  where variedades.codvarie = " & DBSet(vCodvarie, "N")
    Sql = Sql & " and variedades.comerciocomun = 1"
    
    EsVariedadComercializada = (TotalRegistros(Sql) <> 0)

End Function



Public Function EsDeVariedadComercializada(vCodcampo As String) As Boolean
Dim Sql As String

    Sql = "select count(*) from variedades inner join rcampos on variedades.codvarie = rcampos.codvarie where rcampos.codcampo = " & DBSet(vCodcampo, "N")
    Sql = Sql & " and variedades.comerciocomun = 1"
    
    EsDeVariedadComercializada = (TotalRegistros(Sql) <> 0)

End Function

Public Function TieneCamposVariedadComercializada(vSocio As String) As Boolean
Dim Sql As String

    Sql = "select count(*) from rcampos inner join variedades on rcampos.codvarie = variedades.codvarie where codsocio = " & DBSet(vSocio, "N")
    Sql = Sql & " and variedades.comerciocomun = 1"
    
    TieneCamposVariedadComercializada = (TotalRegistros(Sql) <> 0)
    
End Function


Public Function CargarFicheroCsv(vDesFecEnt As String, vHasFecEnt As String, vDesFecAlb As String, vHasFecAlb As String, Entradas As Boolean, albaranes As Boolean, vCd1 As CommonDialog) As Boolean
Dim Sql As String
Dim Sql2 As String
Dim Mens As String

Dim B As Boolean
Dim Rs As ADODB.Recordset
Dim NFic As Integer
Dim Regs As Long
Dim v_cadena As String

    On Error GoTo eCargarFicheroCsv

    conn.BeginTrans

    CargarFicheroCsv = False
    
    B = True
    
    If Entradas Then
        ' cargamos el fichero para luego copiarlo
        B = CargarEntradasClasificadas(vDesFecEnt, vHasFecEnt)
    End If
    
    If albaranes Then
        If B Then B = CargarAlbaranesVenta(vDesFecAlb, vHasFecAlb)
    End If
    
    If B Then
        
        Sql = "select * from comunica_env where fechadescarga is null order by fechacreacion, id"
        
        Set Rs = New ADODB.Recordset
        Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        NFic = FreeFile
        Open App.Path & "\comunica.txt" For Output As NFic
        
        Regs = 0
        v_cadena = ""
        While Not Rs.EOF
            Regs = Regs + 1
            
            v_cadena = DBLet(Rs!Id, "N") & ";" & DBLet(Rs!fechacreacion, "FH") & ";"
            v_cadena = v_cadena & DBLet(Rs!usuariocreacion, "N") & ";" & DBLet(Rs!Tipo, "T") & ";"
            v_cadena = v_cadena & DBLet(Rs!tabla, "T") & ";" & DBLet(Rs!SQLAEJECUTAR, "T") & ";"
            v_cadena = v_cadena & DBLet(Rs!Observaciones, "T") & ";" & Now() & ";" & vUsu.Login & ";"
            
            Print #NFic, v_cadena
            
            Sql2 = "update comunica_env set fechadescarga = " & DBSet(Now(), "FH")
            Sql2 = Sql2 & ", usuariodescarga = " & DBSet(vUsu.Login, "T")
            Sql2 = Sql2 & " where id = " & DBSet(Rs!Id, "N")
            
            conn.Execute Sql2
            
            Rs.MoveNext
        Wend
        
        Close (NFic)
        NFic = -1
        
        If Regs = 0 Then
            MsgBox "No existen datos a comunicar", vbExclamation
            conn.RollbackTrans
            Exit Function
        End If
        
        If CopiarFichero(vCd1) Then
            MsgBox "Proceso realizado correctamente", vbExclamation
        
            CargarFicheroCsv = True
            conn.CommitTrans
            Exit Function
        End If
        
    End If
    
eCargarFicheroCsv:
    MuestraError Err.Number, "Cargar Fichero Csv", Err.Description
    conn.RollbackTrans
End Function

Private Function CopiarFichero(vCd1 As CommonDialog) As Boolean
Dim nomFich As String

On Error GoTo ecopiarfichero

    CopiarFichero = False
    ' abrimos el commondialog para indicar donde guardarlo
'    Me.CommonDialog1.InitDir = App.path

    vCd1.DefaultExt = "csv"
    
    vCd1.Filter = "Archivos csv|csv|"
    vCd1.FilterIndex = 1
    
    ' copiamos el primer fichero
    vCd1.FileName = "comunica.csv"
    vCd1.ShowSave
    
    If vCd1.FileName <> "" Then
        FileCopy App.Path & "\comunica.txt", vCd1.FileName
    End If
    
    CopiarFichero = True
    Exit Function

ecopiarfichero:
    If Err.Number <> 0 Then
        MuestraError Err.Number, Err.Description
    End If
    Err.Clear
End Function



Private Function CargarEntradasClasificadas(vDFec As String, vHFec As String) As Boolean
'[Monica]18/10/2018: antes 2 situaciones ( comunicada y no comunicada )
' Vamos a tener ahora 3 situaciones: 0 = sin comunicar
'                                    1 = comunicada entrada sin clasificacion
'                                    2 = comunicada entrada con clasificacion
Dim Sql As String
Dim Sql2 As String
Dim CadInsert As String
Dim CadValues As String
Dim CadIns2 As String
Dim CadVal2 As String
Dim CadIns3 As String
Dim CadVal3 As String
Dim CadIns4 As String
Dim CadVal4 As String
Dim CadIns5 As String
Dim CadVal5 As String
Dim CadIns6 As String
Dim CadVal6 As String

Dim Rs As ADODB.Recordset
Dim Rs2 As ADODB.Recordset
Dim NumNotac As Long
Dim Transpor As String

    On Error GoTo eCargarEntradasClasificadas

    CargarEntradasClasificadas = False

    ' metemos en comunica las entradas entre fechas a comunicar que sean de las variedades comunes
    Sql = "select * from (rclasifica inner join variedades on rclasifica.codvarie = variedades.codvarie) "
    Sql = Sql & " where variedades.comerciocomun = 1 and estacomunicada in (0,1) " ' [Monica]17/10/2018: ahora pasaremos tanto entradas con clasificacion como sin clasificar
    
'[Monica]06/11/2018: se deben de comunicar todas las entradas
'    ' entradas del socio de la cooperativa de Picassent
'    Sql = Sql & " and codsocio >= " & DBSet(cMaxSocio, "N")
    
    If vDFec <> "" Then Sql = Sql & " and fechaent >= " & DBSet(vDFec, "F")
    If vHFec <> "" Then Sql = Sql & " and fechaent <= " & DBSet(vHFec, "F")
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    ' rclasifica
    CadInsert = "replace into rclasifica (numnotac,fechaent,horaentr,codvarie,codsocio,codcampo,tipoentr,recolect,codtrans,codcapat,codtarif,kilosbru,"
    CadInsert = CadInsert & "numcajon,kilosnet,observac,imptrans,impacarr,imprecol,imppenal,tiporecol,horastra,numtraba,numalbar,fecalbar,"
    CadInsert = CadInsert & "impreso , PrEstimado, transportadopor, KilosTra, contrato) values ("
    
    ' rclasifica_clasif
    CadIns2 = "replace into rclasifica_clasif (numnotac,codvarie,codcalid,muestra,kilosnet) values ("
    
    ' rclasifica_incidencia
    CadIns3 = "replace into rclasifica_incidencia (numnotac,codincid) values ("
    
    
    While Not Rs.EOF
    
        If EntradaClasificada(DBLet(Rs!NumNotac, "N")) Then
    
' tienen su talonario
            NumNotac = DBSet(Rs!NumNotac, "N") ' + 1000000
    
            '[Monica]06/11/2018: ponemos el socio correspondiente dependiendo de es un socio de la cooperativa o no
            If DBLet(Rs!Codsocio, "N") >= cMaxSocio Then
                CadValues = DBSet(NumNotac, "N") & "," & DBSet(Rs!FechaEnt, "F") & "," & DBSet(Rs!horaentr, "FH") & ","
                CadValues = CadValues & DBSet(Rs!codvarie, "N") & "," & DBSet(Rs!Codsocio - cMaxSocio, "N") & "," & DBSet(Rs!codCampo - cMaxCampo, "N") & ","
                CadValues = CadValues & DBSet(Rs!TipoEntr, "N") & "," & DBSet(Rs!Recolect, "N") & ","
            Else
                CadValues = DBSet(NumNotac, "N") & "," & DBSet(Rs!FechaEnt, "F") & "," & DBSet(Rs!horaentr, "FH") & ","
                CadValues = CadValues & DBSet(Rs!codvarie, "N") & "," & DBSet(Rs!Codsocio + cMaxSocio, "N") & "," & DBSet(Rs!codCampo + cMaxCampo, "N") & ","
                CadValues = CadValues & DBSet(Rs!TipoEntr, "N") & "," & DBSet(Rs!Recolect, "N") & ","
            End If
            
            ' transportista
            Transpor = DBLet(Rs!codTrans, "T")
            If Transpor <> "" Then
                If Mid(Transpor, 1, 1) = "A" Then
                    Transpor = Mid(Transpor, 2)
                Else
                    Transpor = "C" & Transpor
                End If
            End If
            CadValues = CadValues & DBSet(Transpor, "T") & ","
            
            ' capataz
            If DBLet(Rs!codcapat, "N") >= cMaxCapa Then
                CadValues = CadValues & DBSet(Rs!codcapat - cMaxCapa, "N") & ","
            Else
                CadValues = CadValues & DBSet(Rs!codcapat + cMaxCapa, "N") & ","
            End If
            
            CadValues = CadValues & DBSet(Rs!Codtarif, "N") & "," & DBSet(Rs!KilosBru, "N") & ","
            CadValues = CadValues & DBSet(Rs!Numcajon, "N") & "," & DBSet(Rs!KilosNet, "N") & "," & DBSet(Rs!Observac, "T") & ","
            CadValues = CadValues & DBSet(Rs!ImpTrans, "N") & "," & DBSet(Rs!impacarr, "N") & "," & DBSet(Rs!imprecol, "N") & ","
            CadValues = CadValues & DBSet(Rs!ImpPenal, "N") & "," & DBSet(Rs!tiporecol, "N") & "," & DBSet(Rs!horastra, "N") & ","
            CadValues = CadValues & DBSet(Rs!numtraba, "N") & "," & DBSet(Rs!NumAlbar, "N") & "," & DBSet(Rs!Fecalbar, "F") & ","
            CadValues = CadValues & DBSet(Rs!impreso, "N") & "," & DBSet(Rs!PrEstimado, "N") & "," & DBSet(Rs!transportadopor, "N") & ","
            CadValues = CadValues & DBSet(Rs!KilosTra, "N") & "," & DBSet(Rs!contrato, "N") & ")"
        
            CadValues = CadInsert & CadValues
        
'[Monica]05/11/2018: borramos las subtablas pq hacemos un replace
            Dim CadValuesDel As String
            
            CadValuesDel = "delete from rclasifica_clasif where numnotac = " & DBSet(NumNotac, "N")
            ComunicaCooperativa "rclasifica", CadValuesDel, "I"
            CadValuesDel = "delete from rclasifica_incidencia where numnotac = " & DBSet(NumNotac, "N")
            ComunicaCooperativa "rclasifica", CadValuesDel, "I"
            
        
            ComunicaCooperativa "rclasifica", CadValues, "I"
        
            ' rclasifica_clasif
            Sql2 = "select * from rclasifica_clasif where numnotac = " & DBSet(NumNotac, "N")
            Set Rs2 = New ADODB.Recordset
            Rs2.Open Sql2, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            
            While Not Rs2.EOF
                CadVal2 = DBSet(NumNotac, "N") & "," & DBSet(Rs!codvarie, "N") & "," & DBSet(Rs2!codcalid, "N") & "," & DBSet(Rs2!Muestra, "N") & ","
                CadVal2 = CadVal2 & DBSet(Rs2!KilosNet, "N") & ")"
            
                CadVal2 = CadIns2 & CadVal2
            
                ComunicaCooperativa "rclasifica_clasif", CadVal2, "I"
                
                Rs2.MoveNext
            Wend
            Set Rs2 = Nothing
            
            ' rclasifica_incidencia
            Sql2 = "select * from rclasifica_incidencia where numnotac = " & DBSet(NumNotac, "N")
            Set Rs2 = New ADODB.Recordset
            Rs2.Open Sql2, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            
            While Not Rs2.EOF
                CadVal3 = DBSet(NumNotac, "N") & "," & DBSet(Rs2!codincid, "N") & ")"
            
                CadVal3 = CadIns3 & CadVal3
            
                ComunicaCooperativa "rclasifica_incidencia", CadVal3, "I"
                
                Rs2.MoveNext
            Wend
            Set Rs2 = Nothing
            
            '[Monica]17/10/2018: ahora las entradas que estan comunicadas y clasificadas le ponemos la situacion 2
            Sql = "update rclasifica set estacomunicada = 2 where numnotac = " & DBSet(NumNotac, "N")
            conn.Execute Sql
            
        Else
        
            NumNotac = DBSet(Rs!NumNotac, "N") ' + 1000000
    
            '[Monica]06/11/2018: se a�ade la condicion de si el socio es o no de la cooperativa que comunica
            If DBLet(Rs!Codsocio, "N") >= cMaxSocio Then
                CadValues = DBSet(NumNotac, "N") & "," & DBSet(Rs!FechaEnt, "F") & "," & DBSet(Rs!horaentr, "FH") & ","
                CadValues = CadValues & DBSet(Rs!codvarie, "N") & "," & DBSet(Rs!Codsocio - cMaxSocio, "N") & "," & DBSet(Rs!codCampo - cMaxCampo, "N") & ","
                CadValues = CadValues & DBSet(Rs!TipoEntr, "N") & "," & DBSet(Rs!Recolect, "N") & ","
            Else
                CadValues = DBSet(NumNotac, "N") & "," & DBSet(Rs!FechaEnt, "F") & "," & DBSet(Rs!horaentr, "FH") & ","
                CadValues = CadValues & DBSet(Rs!codvarie, "N") & "," & DBSet(Rs!Codsocio + cMaxSocio, "N") & "," & DBSet(Rs!codCampo + cMaxCampo, "N") & ","
                CadValues = CadValues & DBSet(Rs!TipoEntr, "N") & "," & DBSet(Rs!Recolect, "N") & ","
            End If
            
            ' transportista
            Transpor = DBLet(Rs!codTrans, "T")
            If Transpor <> "" Then
                If Mid(Transpor, 1, 1) = "A" Then
                    Transpor = Mid(Transpor, 2)
                Else
                    Transpor = "C" & Transpor
                End If
            End If
            CadValues = CadValues & DBSet(Transpor, "T") & ","
            
            ' capataz
            If DBLet(Rs!codcapat, "N") >= cMaxCapa Then
                CadValues = CadValues & DBSet(Rs!codcapat - cMaxCapa, "N") & ","
            Else
                CadValues = CadValues & DBSet(Rs!codcapat + cMaxCapa, "N") & ","
            End If
            
            CadValues = CadValues & DBSet(Rs!Codtarif, "N") & "," & DBSet(Rs!KilosBru, "N") & ","
            CadValues = CadValues & DBSet(Rs!Numcajon, "N") & "," & DBSet(Rs!KilosNet, "N") & "," & DBSet(Rs!Observac, "T") & ","
            CadValues = CadValues & DBSet(Rs!ImpTrans, "N") & "," & DBSet(Rs!impacarr, "N") & "," & DBSet(Rs!imprecol, "N") & ","
            CadValues = CadValues & DBSet(Rs!ImpPenal, "N") & "," & DBSet(Rs!tiporecol, "N") & "," & DBSet(Rs!horastra, "N") & ","
            CadValues = CadValues & DBSet(Rs!numtraba, "N") & "," & DBSet(Rs!NumAlbar, "N") & "," & DBSet(Rs!Fecalbar, "F") & ","
            CadValues = CadValues & DBSet(Rs!impreso, "N") & "," & DBSet(Rs!PrEstimado, "N") & "," & DBSet(Rs!transportadopor, "N") & ","
            CadValues = CadValues & DBSet(Rs!KilosTra, "N") & "," & DBSet(Rs!contrato, "N") & ")"
        
            CadValues = CadInsert & CadValues
        
            ComunicaCooperativa "rclasifica", CadValues, "I"
            
        
            '[Monica]17/10/2018: entrada que esta comunicada y sin clasificar situacion 1
            Sql = "update rclasifica set estacomunicada = 1 where numnotac = " & DBSet(NumNotac, "N")
            conn.Execute Sql
        
        End If
        
        Rs.MoveNext
    Wend
    
    CargarEntradasClasificadas = True
    Exit Function
    
eCargarEntradasClasificadas:
    MuestraError Err.Number, "Cargar Entradas Clasificadas", Err.Description
End Function

Private Function CargarAlbaranesVenta(vDFec As String, vHFec As String) As String
Dim Sql As String
Dim Sql2 As String
Dim CadInsert As String
Dim CadValues As String
Dim CadIns2 As String
Dim CadVal2 As String
Dim CadIns3 As String
Dim CadVal3 As String
Dim CadIns4 As String
Dim CadVal4 As String
Dim CadIns5 As String
Dim CadVal5 As String
Dim CadIns6 As String
Dim CadVal6 As String

Dim Rs As ADODB.Recordset
Dim Rs2 As ADODB.Recordset
Dim RsCab As ADODB.Recordset
Dim Albaran As Long


    On Error GoTo eCargarAlbaranesVenta
    
    
    ' metemos los albaranes de venta
    Sql = "select distinct albaran.numalbar from (albaran inner join albaran_variedad on albaran.numalbar = albaran_variedad.numalbar) "
    Sql = Sql & " inner join variedades on albaran_variedad.codvarie = variedades.codvarie "
    Sql = Sql & " where variedades.comerciocomun = 1 and albaran.estacomunicada = 0 "
    If vDFec <> "" Then Sql = Sql & " and fechaalb >= " & DBSet(vDFec, "F")
    If vHFec <> "" Then Sql = Sql & " and fechaalb <= " & DBSet(vHFec, "F")
    
    ' albaran
    CadInsert = "replace into albaran (numalbar,fechaalb,codclien,coddesti,codtrans,matriveh,matrirem,refclien,codtimer,totpalet,"
    CadInsert = CadInsert & "portespre,nrocontra,nroactas,numpedid,fechaped,observac,pasaridoc,codalmac,portespag,paletspag,"
    CadInsert = CadInsert & "numerocmr,comisionespre,comisionespag,codcomis,codsocio,precnodef, codtipom) values ("
    
    ' albaran_variedad
    CadIns2 = "replace into albaran_variedad (numalbar,numlinea,codvarie,codvarco,codforfait,codmarca,categori,totpalet,numcajas,pesobrut,"
    CadIns2 = CadIns2 & "pesoneto,preciopro,preciodef,codincid,impcomis,observac,unidades,referencia,codpalet,nrotraza,"
    CadIns2 = CadIns2 & "codtipo,sefactura,codcomis,nrotraza1,nrotraza2,nrotraza3,nrotraza4,nrotraza5,nrotraza6,expediente) values ("
    
    ' albaran_calibre
    CadIns3 = "replace into albaran_calibre (numalbar,numlinea,numline1,codvarie,codcalib,numcajas,pesobrut,pesoneto,unidades,preciopro"
    CadIns3 = CadIns3 & ") values ("
    
    ' albaran_palets
    CadIns4 = "replace into albaran_palets (numalbar,numlinea,numpalet) values ("
    
    ' albaran_envase
    CadIns5 = "replace into albaran_envase (numalbar,numlinea,fechamov,codartic,tipomovi,cantidad,codclien,impfianza,factura,fecfactu) values ("
    
    
    Set RsCab = New ADODB.Recordset
    RsCab.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    
    
    While Not RsCab.EOF
        Albaran = DBLet(RsCab!NumAlbar, "N")
    
        Sql = "select * from albaran where numalbar = " & DBSet(Albaran, "N")
        Set Rs = New ADODB.Recordset
        Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        If Not Rs.EOF Then
        
            CadValues = DBSet(Albaran, "N") & "," & DBSet(Rs!FechaAlb, "F") & "," & DBSet(Rs!CodClien, "N") & "," & DBSet(Rs!coddesti, "N") & ","
            CadValues = CadValues & DBSet(Rs!codTrans, "T") & "," & DBSet(Rs!matriveh, "T", "S") & "," & DBSet(Rs!matrirem, "T", "S") & ","
            CadValues = CadValues & DBSet(Rs!refclien, "T", "S") & "," & DBSet(Rs!codtimer, "N") & "," & DBSet(Rs!totpalet, "N", "S") & ","
            CadValues = CadValues & DBSet(Rs!portespre, "N") & "," & DBSet(Rs!nrocontra, "T", "S") & "," & DBSet(Rs!nroactas, "N", "S") & ","
            CadValues = CadValues & DBSet(Rs!numpedid, "N", "S") & "," & DBSet(Rs!fechaped, "F", "S") & "," & DBSet(Rs!Observac, "T", "S") & ","
            CadValues = CadValues & DBSet(Rs!pasaridoc, "N") & "," & DBSet(Rs!codAlmac, "N") & "," & DBSet(Rs!portespag, "N") & ","
            CadValues = CadValues & DBSet(Rs!paletspag, "N") & "," & DBSet(Rs!numerocmr, "N", "S") & "," & DBSet(Rs!comisionespre, "N") & ","
            CadValues = CadValues & DBSet(Rs!comisionespag, "N", "S") & "," & DBSet(Rs!codcomis, "N", "S") & "," & DBSet(Rs!Codsocio, "N", "S") & ","
            CadValues = CadValues & DBSet(Rs!precnodef, "N") & "," & DBSet(Rs!CodTipom, "T") & ")"
        
            CadValues = CadInsert & CadValues
            
                
            ComunicaCooperativa "albaran", CadValues, "I"
        
            ' albaran_variedad
            Sql2 = "select * from albaran_variedad where numalbar = " & DBSet(Albaran, "N")
            Set Rs2 = New ADODB.Recordset
            Rs2.Open Sql2, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            
            While Not Rs2.EOF
                CadVal2 = DBSet(Albaran, "N") & "," & DBSet(Rs2!NumLinea, "N") & "," & DBSet(Rs2!codvarie, "N") & "," & DBSet(Rs2!codvarco, "N") & ","
                CadVal2 = CadVal2 & DBSet(Rs2!codforfait, "T") & "," & DBSet(Rs2!codmarca, "N") & "," & DBSet(Rs2!categori, "T") & "," & DBSet(Rs2!totpalet, "N") & ","
                CadVal2 = CadVal2 & DBSet(Rs2!NumCajas, "N") & "," & DBSet(Rs2!pesobrut, "N") & ","
                CadVal2 = CadVal2 & DBSet(Rs2!PesoNeto, "N") & "," & DBSet(Rs2!preciopro, "N") & "," & DBSet(Rs2!preciodef, "N") & ","
                CadVal2 = CadVal2 & DBSet(Rs2!codincid, "N") & "," & DBSet(Rs2!impcomis, "N") & "," & DBSet(Rs2!Observac, "T") & ","
                CadVal2 = CadVal2 & DBSet(Rs2!Unidades, "N") & "," & DBSet(Rs2!Referencia, "T") & "," & DBSet(Rs2!codpalet, "N") & ","
                CadVal2 = CadVal2 & DBSet(Rs2!nrotraza, "T") & "," & DBSet(Rs2!codtipo, "N") & "," & DBSet(Rs2!sefactura, "N") & ","
                CadVal2 = CadVal2 & DBSet(Rs2!codcomis, "N") & "," & DBSet(Rs2!nrotraza1, "T", "S") & "," & DBSet(Rs2!nrotraza2, "T", "S") & ","
                CadVal2 = CadVal2 & DBSet(Rs2!nrotraza3, "T", "S") & "," & DBSet(Rs2!nrotraza4, "T", "S") & "," & DBSet(Rs2!nrotraza5, "T", "S") & ","
                CadVal2 = CadVal2 & DBSet(Rs2!nrotraza6, "T", "S") & "," & DBSet(Rs2!expediente, "T", "S") & ")"
            
                CadVal2 = CadIns2 & CadVal2
        
                ComunicaCooperativa "albaran_variedad", CadVal2, "I"
        
                Rs2.MoveNext
            Wend
            Set Rs2 = Nothing
            
            ' albaran_calibre
            Sql2 = "select * from albaran_calibre where numalbar = " & DBSet(Albaran, "N")
            Set Rs2 = New ADODB.Recordset
            Rs2.Open Sql2, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            
            While Not Rs2.EOF
                CadVal3 = DBSet(Albaran, "N") & "," & DBSet(Rs2!NumLinea, "N") & "," & DBSet(Rs2!numline1, "N") & "," & DBSet(Rs2!codvarie, "N") & ","
                CadVal3 = CadVal3 & DBSet(Rs2!codcalib, "N") & "," & DBSet(Rs2!NumCajas, "N") & "," & DBSet(Rs2!pesobrut, "N") & ","
                CadVal3 = CadVal3 & DBSet(Rs2!PesoNeto, "N") & "," & DBSet(Rs2!Unidades, "N") & "," & DBSet(Rs2!preciopro, "N") & ")"
            
                CadVal3 = CadIns3 & CadVal3
            
                ComunicaCooperativa "albaran_calibre", CadVal3, "I"
                
                Rs2.MoveNext
            Wend
            Set Rs2 = Nothing
            
            
            ' albaran_palets
            Sql2 = "select * from albaran_palets where numalbar = " & DBSet(Albaran, "N")
            Set Rs2 = New ADODB.Recordset
            Rs2.Open Sql2, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            
            While Not Rs2.EOF
                CadVal4 = DBSet(Albaran, "N") & "," & DBSet(Rs2!NumLinea, "N") & "," & DBSet(Rs2!NumPalet, "N") & ")"
            
                CadVal4 = CadIns4 & CadVal4
            
                ComunicaCooperativa "albaran_palets", CadVal4, "I"
                
                Rs2.MoveNext
            Wend
            Set Rs2 = Nothing
        
            ' albaran_envase
            CadIns5 = "replace into albaran_envase (numalbar,numlinea,fechamov,codartic,tipomovi,cantidad,codclien,impfianza,factura,fecfactu) values ("
            
            Sql2 = "select * from albaran_palets where numalbar = " & DBSet(Albaran, "N")
            Set Rs2 = New ADODB.Recordset
            Rs2.Open Sql2, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            
            While Not Rs2.EOF
                CadVal5 = DBSet(Albaran, "N") & "," & DBSet(Rs2!NumLinea, "N") & "," & DBSet(Rs2!Fechamov, "F") & ","
                CadVal5 = CadVal5 & DBSet(Rs2!codArtic, "T") & "," & DBSet(Rs2!tipomovi, "N") & "," & DBSet(Rs2!cantidad, "N") & ","
                CadVal5 = CadVal5 & DBSet(Rs2!CodClien, "N") & "," & DBSet(Rs2!impfianza, "N") & "," & DBSet(Rs2!Factura, "T") & ","
                CadVal5 = CadVal5 & DBSet(Rs2!fecfactu, "F") & ")"
            
                CadVal5 = CadIns5 & CadVal5
            
                ComunicaCooperativa "albaran_palets", CadVal5, "I"
                
                Rs2.MoveNext
            Wend
            Set Rs2 = Nothing
            
            'lo marcamos como que est� comunicado
            Sql = "update albaran set estacomunicada = 1 where numalbar = " & DBSet(Albaran, "N")
            conn.Execute Sql
            
        End If
        Set Rs = Nothing
            
        RsCab.MoveNext
    Wend
    Set Rs = Nothing

    CargarAlbaranesVenta = True
    Exit Function

eCargarAlbaranesVenta:
    MuestraError Err.Number, "Cargar Albaranes Venta", Err.Description
End Function


Public Function EntradaComunicada(Nota As String) As Boolean
Dim Sql As String

    Sql = "select estacomunicada from rclasifica where numnotac = " & DBSet(Nota, "N")
    Sql = DevuelveValor(Sql)
    '[Monica]18/10/2018: ahora la entrada puede estar comunicada sin clasificacion (1) o con la clasificacion (2)
    EntradaComunicada = ((Sql = 1) Or (Sql = 2))

End Function

Public Function EntradaComunicadaConClasificacion(Nota As String) As Boolean
Dim Sql As String

    Sql = "select estacomunicada from rclasifica where numnotac = " & DBSet(Nota, "N")
    Sql = DevuelveValor(Sql)
    '[Monica]18/10/2018: ahora la entrada puede estar comunicada sin clasificacion (1) o con la clasificacion (2)
    EntradaComunicadaConClasificacion = (Sql = 2)

End Function


