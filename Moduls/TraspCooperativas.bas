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

    On Error GoTo eComunicaCooperativa
    
    ComunicaCooperativa = False
        
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


Public Function CargarFicheroCsv(vDesFecEnt As String, vHasFecEnt As String, vDesFecAlb As String, vHasFecAlb As String, Entradas As Boolean, Albaranes As Boolean) As Boolean
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
    
    If Albaranes Then
        If B Then B = CargarAlbaranesVenta(vDesFecAlb, vHasFecAlb)
    End If
    
    If B Then
        
        Sql = "select * from comunica_env where fechadescarga is null"
'        and  (( true "
'        If vDesFec <> "" Then Sql = Sql & " and fechacreacion >= " & DBSet(vDesFec, "F")
'        If vHasFec <> "" Then Sql = Sql & " and fechacreacion <= " & DBSet(vHasFec, "F")
'        Sql = Sql & ") or tabla in ('rclasifica','rclasifica_clasif','rclasifica_incidencia','albaran','albaran_variedad',"
'        Sql = Sql & "'albaran_calibre','albaran_envase','albaran_palets'))"
        
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
            v_cadena = v_cadena & DBLet(Rs!tabla, "T") & ";" & DBLet(Rs!sqlaejecutar, "T") & ";"
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
        
        If CopiarFichero Then
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

Private Function CopiarFichero() As Boolean
Dim nomFich As String

On Error GoTo ecopiarfichero

    CopiarFichero = False
    ' abrimos el commondialog para indicar donde guardarlo
'    Me.CommonDialog1.InitDir = App.path

    cd1.DefaultExt = "csv"
    
    cd1.Filter = "Archivos csv|csv|"
    cd1.FilterIndex = 1
    
    ' copiamos el primer fichero
    cd1.FileName = "comunica.csv"
    Me.cd1.ShowSave
    
    If cd1.FileName <> "" Then
        FileCopy App.Path & "\comunica.csv", cd1.FileName
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

    On Error GoTo eCargarEntradasClasificadas

    CargarEntradasClasificadas = False

    ' metemos en comunica las entradas entre fechas a comunicar que sean de las variedades comunes
    Sql = "select * from (rclasifica inner join variedades on rclasifica.codvarie = variedades.codvarie) inner join rclasifica_clasif "
    Sql = Sql & " on rclasifica.numnotac = rclasifica_clasif.numnotac "
    Sql = Sql & " where variedades.comerciocomun = 1 and estacomunicada = 0 "
    If vDFec <> "" Then Sql = Sql & " and fechaent >= " & DBSet(vDFec, "F")
    If vHFec <> "" Then Sql = Sql & " and fechaent <= " & DBSet(vHFec, "F")
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    ' rclasifica
    CadInsert = "insert into rclasifica (numnotac,fechaent,horaentr,codvarie,codsocio,codcampo,tipoentr,recolect,codtrans,codcapat,codtarif,kilosbru,"
    CadInsert = CadInsert & "numcajon,kilosnet,observac,imptrans,impacarr,imprecol,imppenal,tiporecol,horastra,numtraba,numalbar,fecalbar,"
    CadInsert = CadInsert & "impreso , PrEstimado, transportadopor, KilosTra, contrato) values ("
    
    ' rclasifica_clasif
    CadIns2 = "insert into rclasifica_clasif (numnotac,codvarie,codcalid,muestra,kilosnet) values ("
    
    ' rclasifica_incidencia
    CadIns3 = "insert into rclasifica_incidencia (numnotac,codincid) values ("
    
    
    While Not Rs.EOF
    
        If EntradaClasificada(DBLet(Rs!NumNotac, "N")) Then
    
    
            NumNotac = DBSet(Rs!NumNotac, "N") + 1000000
    
    
            CadValues = DBSet(NumNotac, "N") & "," & DBSet(Rs!FechaEnt, "F") & "," & DBSet(Rs!horaentr, "H") & ","
            CadValues = CadValues & DBSet(Rs!Codvarie, "N") & "," & DBSet(Rs!Codsocio, "N") & "," & DBSet(Rs!codcampo, "N") & ","
            CadValues = CadValues & DBSet(Rs!TipoEntr, "N") & "," & DBSet(Rs!Recolect, "N") & "," & DBSet(Rs!codTrans, "T") & ","
            CadValues = CadValues & DBSet(Rs!codcapat, "N") & "," & DBSet(Rs!Codtarif, "N") & "," & DBSet(Rs!KilosBru, "N") & ","
            CadValues = CadValues & DBSet(Rs!Numcajon, "N") & "," & DBSet(Rs!KilosNet, "N") & "," & DBSet(Rs!Observac, "T") & ","
            CadValues = CadValues & DBSet(Rs!ImpTrans, "N") & "," & DBSet(Rs!impacarr, "N") & "," & DBSet(Rs!imprecol, "N") & ","
            CadValues = CadValues & DBSet(Rs!ImpPenal, "N") & "," & DBSet(Rs!tiporecol, "N") & "," & DBSet(Rs!horastra, "N") & ","
            CadValues = CadValues & DBSet(Rs!numtraba, "N") & "," & DBSet(Rs!numalbar, "N") & "," & DBSet(Rs!Fecalbar, "F") & ","
            CadValues = CadValues & DBSet(Rs!impreso, "N") & "," & DBSet(Rs!PrEstimado, "N") & "," & DBSet(Rs!transportadopor, "N") & ","
            CadValues = CadValues & DBSet(Rs!KilosTra, "N") & "," & DBSet(Rs!contrato, "N") & ")"
        
            CadValues = CadInsert & CadValues
        
            ComunicaCooperativa "rclasifica", CadValues, "I"
        
            ' rclasifica_clasif
            Sql2 = "select * from rclasifica_clasif where numnotac = " & DBSet(Rs!NumNotac, "N")
            Set Rs2 = New ADODB.Recordset
            Rs2.Open Sql2, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            
            While Not Rs2.EOF
                CadVal2 = DBSet(NumNotac, "N") & "," & DBSet(Rs!Codvarie, "N") & "," & DBSet(Rs!codcalid, "N") & "," & DBSet(Rs!Muestra, "N") & ","
                CadVal2 = CadVal2 & DBSet(Rs!KilosNet, "N") & ")"
            
                CadVal2 = CadIns2 & CadVal2
            
                ComunicaCooperativa "rclasifica_clasif", CadVal2, "I"
                
                Rs2.MoveNext
            Wend
            Set Rs2 = Nothing
            
            ' rclasifica_incidencia
            Sql2 = "select * from rclasifica_incidencia where numnotac = " & DBSet(Rs!NumNotac, "N")
            Set Rs2 = New ADODB.Recordset
            Rs2.Open Sql2, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            
            While Not Rs2.EOF
                CadVal3 = DBSet(NumNotac, "N") & "," & DBSet(Rs!codincid, "N") & ")"
            
                CadVal3 = CadIns3 & CadVal3
            
                ComunicaCooperativa "rclasifica_incidencia", CadVal3, "I"
                
                Rs2.MoveNext
            Wend
            Set Rs2 = Nothing
            
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
Dim Albaran As Long


    On Error GoTo eCargarAlbaranesVenta
    
    
    ' metemos los albaranes de venta
    Sql = "select * from (albaran inner join albaran_variedad on albaran.numalbar = albaran_variedad.numalbar) "
    Sql = Sql & " inner join variedades on albaran_variedad.codvarie = variedades.codvarie "
    Sql = Sql & " where variedades.comerciocomun = 1 and albaran.estacomunicada = 0 "
    If vDFec <> "" Then Sql = Sql & " and fechaalb >= " & DBSet(vDFec, "F")
    If vHFec <> "" Then Sql = Sql & " and fechaalb <= " & DBSet(vHFec, "F")
    
    ' albaran
    CadInsert = "insert into albaran (numalbar,fechaalb,codclien,coddesti,codtrans,matriveh,matrirem,refclien,codtimer,totpalet,"
    CadInsert = CadInsert & "portespre,nrocontra,nroactas,numpedid,fechaped,observac,pasaridoc,codalmac,portespag,paletspag,"
    CadInsert = CadInsert & "numerocmr,comisionespre,comisionespag,codcomis,codsocio,precnodef) values ("
    
    ' albaran_variedad
    CadIns2 = "insert into albaran_variedad (numalbar,numlinea,codvarie,codvarco,codforfait,codmarca,categori,totpalet,numcajas,pesobrut,"
    CadIns2 = CadIns2 & "pesoneto,preciopro,preciodef,codincid,impcomis,observac,unidades,referencia,codpalet,nrotraza,"
    CadIns2 = CadIns2 & "codtipo,sefactura,codcomis,nrotraza1,nrotraza2,nrotraza3,nrotraza4,nrotraza5,nrotraza6,expediente) values ("
    
    ' albaran_calibre
    CadIns3 = "insert into albaran_calibre (numalbar,numlinea,numline1,codvarie,codcalib,numcajas,pesobrut,pesoneto,unidades,preciopro"
    CadIns3 = CadIns3 & ") values ("
    
    ' albaran_palets
    CadIns4 = "insert into albaran_palets (numalbar,numlinea,numpalet) values ("
    
    ' albaran_envase
    CadIns5 = "insert into albaran_envase (numalbar,numlinea,fechamov,codartic,tipomovi,cantidad,codclien,impfianza,factura,fecfactu) values ("
    
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    
    
    While Not Rs.EOF
        Albaran = DBLet(Rs!numalbar, "N")
    
        CadValues = DBSet(Albaran, "N") & "," & DBSet(Rs!FechaAlb, "F") & "," & DBSet(Rs!CodClien, "N") & "," & DBSet(Rs!coddesti, "N") & ","
        CadValues = CadValues & DBSet(Rs!codTrans, "T") & "," & DBSet(Rs!matriveh, "T") & "," & DBSet(Rs!matrirem, "T") & ","
        CadValues = CadValues & DBSet(Rs!refclien, "T") & "," & DBSet(Rs!codtimer, "N") & "," & DBSet(Rs!totpalet, "N") & ","
        CadValues = CadValues & DBSet(Rs!portespre, "N") & "," & DBSet(Rs!nrocontra, "T") & "," & DBSet(Rs!nroactas, "N") & ","
        CadValues = CadValues & DBSet(Rs!numpedid, "N") & "," & DBSet(Rs!fechaped, "F") & "," & DBSet(Rs!Observac, "T") & ","
        CadValues = CadValues & DBSet(Rs!pasaridoc, "N") & "," & DBSet(Rs!codAlmac, "N") & "," & DBSet(Rs!portespag, "N") & ","
        CadValues = CadValues & DBSet(Rs!paletspag, "N") & "," & DBSet(Rs!numerocmr, "N") & "," & DBSet(Rs!comisionespre, "N") & ","
        CadValues = CadValues & DBSet(Rs!comisionespag, "N") & "," & DBSet(Rs!codcomis, "N") & "," & DBSet(Rs!Codsocio, "N") & ","
        CadValues = CadValues & DBSet(Rs!precnodef, "N") & ")"
    
        CadValues = CadInsert & CadValues
    
        ComunicaCooperativa "albaran", CadValues, "I"
    
        ' albaran_variedad
        Sql2 = "select * from albaran_variedad where numalbar = " & DBSet(Albaran, "N")
        Set Rs2 = New ADODB.Recordset
        Rs2.Open Sql2, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        While Not Rs2.EOF
            CadVal2 = DBSet(Albaran, "N") & "," & DBSet(Rs2!numlinea, "N") & "," & DBSet(Rs2!Codvarie, "N") & "," & DBSet(Rs2!codvarco, "N") & ","
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
            CadVal3 = DBSet(Albaran, "N") & "," & DBSet(Rs2!numlinea, "N") & "," & DBSet(Rs2!numline1, "N") & DBSet(Rs2!Codvarie, "N") & ","
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
            CadVal4 = DBSet(Albaran, "N") & "," & DBSet(Rs2!numlinea, "N") & "," & DBSet(Rs2!NumPalet, "N") & ")"
        
            CadVal4 = CadIns4 & CadVal4
        
            ComunicaCooperativa "albaran_palets", CadVal4, "I"
            
            Rs2.MoveNext
        Wend
        Set Rs2 = Nothing
    
        ' albaran_envase
        CadIns5 = "insert into albaran_envase (numalbar,numlinea,fechamov,codartic,tipomovi,cantidad,codclien,impfianza,factura,fecfactu) values ("
        
        Sql2 = "select * from albaran_palets where numalbar = " & DBSet(Albaran, "N")
        Set Rs2 = New ADODB.Recordset
        Rs2.Open Sql2, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        While Not Rs2.EOF
            CadVal5 = DBSet(Albaran, "N") & "," & DBSet(Rs2!numlinea, "N") & "," & DBSet(Rs2!Fechamov, "F") & ","
            CadVal5 = CadVal5 & DBSet(Rs2!codArtic, "T") & "," & DBSet(Rs2!tipomovi, "N") & "," & DBSet(Rs2!cantidad, "N") & ","
            CadVal5 = CadVal5 & DBSet(Rs2!CodClien, "N") & "," & DBSet(Rs2!impfianza, "N") & "," & DBSet(Rs2!Factura, "T") & ","
            CadVal5 = CadVal5 & DBSet(Rs2!fecfactu, "F") & ")"
        
            CadVal5 = CadIns5 & CadVal5
        
            ComunicaCooperativa "albaran_palets", CadVal5, "I"
            
            Rs2.MoveNext
        Wend
        Set Rs2 = Nothing
        
        Rs.MoveNext
    Wend
    Set Rs = Nothing

    CargarAlbaranesVenta = True
    Exit Function

eCargarAlbaranesVenta:
    MuestraError Err.Number, "Cargar Albaranes Venta", Err.Description
End Function



Private Function ProcesarFicheroComunicacion2(nomFich As String) As Boolean
Dim NF As Long
Dim cad As String
Dim i As Integer
Dim longitud As Long
Dim Rs As ADODB.Recordset
Dim RS1 As ADODB.Recordset
Dim NumReg As Long
Dim Sql As String
Dim Sql1 As String
Dim Total As Long
Dim v_cant As Currency
Dim v_impo As Currency
Dim v_prec As Currency
Dim B As Boolean

    On Error GoTo eProcesarFichero2
    
    ProcesarFicheroComunicacion2 = False
    
    NF = FreeFile
    Open nomFich For Input As #NF ' & "\BV" & Format(CDate(txtcodigo(0).Text), "ddmmyy") & "." & Format(txtcodigo(1).Text, "000") For Input As #NF
    
    Line Input #NF, cad
    i = 0
    
    lblProgres(0).Caption = "Insertando en Tabla temporal: " & nomFich
    longitud = FileLen(nomFich)
    
    Pb1.visible = True
    Me.Pb1.Max = longitud
    Me.Refresh
    Me.Pb1.Value = 0
    DoEvents
    ' PROCESO DEL FICHERO VENTAS.TXT

    B = True

    While Not EOF(NF) And B
        i = i + 1
        
        Me.Pb1.Value = Me.Pb1.Value + Len(cad)
        lblProgres(1).Caption = "Linea " & i
        Me.Refresh
        DoEvents
        B = ComprobarRegistro(cad)
        
        Line Input #NF, cad
    Wend
    Close #NF
    
    If cad <> "" Then
        i = i + 1
        
        Me.Pb1.Value = Me.Pb1.Value + Len(cad)
        lblProgres(1).Caption = "Linea " & i
        Me.Refresh
        DoEvents
        B = ComprobarRegistro(cad)
    End If
    
    Pb1.visible = False
    lblProgres(0).Caption = ""
    lblProgres(1).Caption = ""

    ProcesarFicheroComunicacion2 = B
    Exit Function

eProcesarFicheroComunicacion2:
    ProcesarFicheroComunicacion2 = False
End Function


Private Function ComprobarRegistro(cad As String) As Boolean
Dim Sql As String

    On Error GoTo eComprobarRegistro

    ComprobarRegistro = True

    Fecha = RecuperaValorNew(cad, ";", 1)
    codsoc = RecuperaValor(cad, 4)
    numfactu = RecuperaValor(cad, 2)
    numfactu = Replace(numfactu, "-", "|") & "|"
    Digito = RecuperaValor(numfactu, 1)
    numfactu = RecuperaValor(numfactu, 4)
    numfactu = Format((CInt(Digito) * 1000000) + CLng(numfactu), "0000000")
    
    
    baseimpo = RecuperaValor(cad, 6)
    CuotaIva = RecuperaValor(cad, 7)
    TotalFac = RecuperaValor(cad, 8)
    
    c_BaseImpo = CCur(TransformaPuntosComas(baseimpo))
    c_CuotaIva = CCur(TransformaPuntosComas(CuotaIva))
    c_TotalFac = CCur(TransformaPuntosComas(TotalFac))
    
    
    'Comprobamos fechas
    If Not EsFechaOK(Fecha) Then
        Mens = "Fecha incorrecta"
        Sql = "insert into tmpinformes (codusu, fecha1, importe1, importe2, importe3, " & _
              "importe4, importe5, nombre1) values (" & _
              vUsu.Codigo & "," & DBSet(Fecha, "F") & _
              "," & DBSet(codsoc, "N") & "," & _
              DBSet(numfactu, "N") & "," & _
              DBSet(c_BaseImpo, "N") & "," & _
              DBSet(c_CuotaIva, "N") & "," & _
              DBSet(c_TotalFac, "N") & "," & DBSet(Mens, "T") & ")"
        
        conn.Execute Sql
    End If

End Function
