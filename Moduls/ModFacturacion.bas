Attribute VB_Name = "ModFacturacion"
' Modulo en donde se encuentran los procedimintos para la facturacion
'
'Dim db As BaseDatos
Option Explicit

Dim Rs As ADODB.Recordset
Dim ImpFactu As Currency
Dim TotalImp As Currency
Dim numser As String
Dim DC As Dictionary
            
Dim baseimpo As Currency
Dim BaseReten As Currency
Dim ImpoIva As Currency
Dim ImpoReten As Currency
Dim PorcIva As Currency
Dim PorcReten As Currency
Dim ImpoAFO As Currency
Dim PorcAFO As Currency
Dim BaseAFO As Currency

Dim Gastos As Currency

Dim Anticipos As Currency

Dim TotalFac As Currency

Dim vSocio As cSocio
Dim vTipoMov As CTiposMov

Dim numfactu As Long

Private cadW As String
Private TipoAlb As String
Private TipoFac As String

Dim Errores As String
Dim ErroresAux As String


'Insertar Cabecera de factura
Public Function InsertCabecera(tipoMov As String, numfactu As String, FecFac As String, Optional EsAnticipoGasto As Boolean, Optional EsAnticipoRetirada As Boolean, Optional EsLiqComplementaria As Boolean) As Boolean

    Dim SQL As String
    
    On Error GoTo eInsertCabe
    
    MensError = ""
    InsertCabecera = False
    
    SQL = "insert into rfactsoc (codtipom, numfactu, fecfactu, codsocio, baseimpo, tipoiva, porc_iva,"
    SQL = SQL & "imporiva, tipoirpf, basereten, porc_ret, impreten, baseaport, porc_apo, impapor, totalfac, impreso, contabilizado, pasaridoc, esanticipogasto, esretirada, esliqcomplem, observaciones) "
    SQL = SQL & " values ('" & tipoMov & "'," & DBSet(numfactu, "N") & "," & DBSet(FecFac, "F") & "," & DBSet(vSocio.Codigo, "N") & ","
    
    '[Monica]29/04/2011: INTERNAS
    If vSocio.EsFactADVInt Then
        SQL = SQL & DBSet(baseimpo, "N") & "," & vParamAplic.CodIvaExeADV & "," & DBSet(PorcIva, "N") & ","
    Else
        SQL = SQL & DBSet(baseimpo, "N") & "," & vSocio.CodIva & "," & DBSet(PorcIva, "N") & ","
    End If
    
    SQL = SQL & DBSet(ImpoIva, "N") & "," & DBSet(vSocio.TipoIRPF, "N") & "," & DBSet(BaseReten, "N") & ","
    SQL = SQL & DBSet(PorcReten, "N") & "," & DBSet(ImpoReten, "N") & "," & DBSet(BaseAFO, "N", "S") & "," & DBSet(PorcAFO, "N", "S") & "," & DBSet(ImpoAFO, "N", "S") & "," & DBSet(TotalFac, "N") & ","
    SQL = SQL & "0,0,0,"
    If EsAnticipoGasto Then
        SQL = SQL & "1"
    Else
        SQL = SQL & "0"
    End If
    
    If EsAnticipoRetirada Then
        SQL = SQL & ",1"
    Else
        SQL = SQL & ",0"
    End If
    
    If EsLiqComplementaria Then
        SQL = SQL & ",1"
    Else
        SQL = SQL & ",0"
    End If
    '[Monica]11/03/2015: añadidas observaciones de la factura
    SQL = SQL & "," & DBSet(ObsFactura, "T")
    
    SQL = SQL & ")"
    
    conn.Execute SQL
    
    InsertCabecera = True
    
    Exit Function

eInsertCabe:
    MensError = "Error en la inserción en rfactsoc de la factura " & numfactu & " del socio " & vSocio.Codigo
    MuestraError Err.Number, MensError
End Function

'Insertar Linea de factura (variedad)
Public Function InsertLinea(tipoMov As String, numfactu As String, FecFac As String, Variedad As String, campo As String, Kilos As String, Importe As String, Gastos As String, Optional KiloGrado As String) As Boolean
Dim Precio As Currency

    Dim SQL As String
    Dim ImpLinea As Currency
    
    On Error GoTo eInsertLinea
    
    InsertLinea = False
    
    MensError = ""
    Precio = 0
    If CCur(ImporteSinFormato(Kilos)) <> 0 Then
        Precio = Round2(CCur(ImporteSinFormato(Importe)) / CCur(ImporteSinFormato(Kilos)), 4)
    End If
    
    SQL = "insert into tmpFact_variedad (codtipom, numfactu, fecfactu, codvarie, codcampo, "
    SQL = SQL & "kilosnet, preciomed, imporvar, imporgasto, kilogrado) values ("
    SQL = SQL & "'" & tipoMov & "'," & DBSet(numfactu, "N") & "," & DBSet(FecFac, "F") & ","
    SQL = SQL & DBSet(Variedad, "N") & "," & DBSet(campo, "N") & ","
    SQL = SQL & DBSet(ImporteSinFormato(Kilos), "N") & "," & DBSet(Precio, "N") & ","
    SQL = SQL & DBSet(ImporteSinFormato(Importe), "N")
    SQL = SQL & "," & DBSet(ImporteSinFormato(Gastos), "N")
    SQL = SQL & "," & DBSet(ImporteSinFormato(KiloGrado), "N") & ")"
    
    conn.Execute SQL

    InsertLinea = True
    Exit Function
    
eInsertLinea:
    If Err.Number <> 0 Then
        MensError = "Se ha producido un error en la inserción de la linea de factura"
        MuestraError Err.Number, MensError, Err.Descripc
    End If
    
End Function


'Insertar Linea de factura (variedad)
Public Function InsertLineasIngresos(tipoMov As String, numfactu As String, FecFac As String, Variedad As String, Socio As String) As Boolean
Dim SQL As String
Dim ImpLinea As Currency
    
    On Error GoTo eInsertLinea
    
    InsertLineasIngresos = False
    
    MensError = ""
    
    SQL = "insert into tmpFact_ingresos (codtipom, numfactu, fecfactu, codvarie, concepto, importe) "
    SQL = SQL & " select '" & tipoMov & "'," & DBSet(numfactu, "N") & "," & DBSet(FecFac, "F") & ", codvarie, concepto, importe "
    SQL = SQL & " from ringresos where codsocio = " & DBSet(Socio, "N") & " and codvarie = " & DBSet(Variedad, "N")
    
    conn.Execute SQL

    InsertLineasIngresos = True
    Exit Function
    
eInsertLinea:
    If Err.Number <> 0 Then
        MensError = "Se ha producido un error en la inserción de lineas de ingresos"
        MuestraError Err.Number, MensError, Err.Descripc
    End If
    
End Function




'Insertar Linea de factura (calidad)
Public Function InsertLineaCalidad(tipoMov As String, numfactu As String, FecFac As String, Variedad As String, campo As String, Calidad As String, Kilos As String, Importe As String, Optional Bonificacion As String) As Boolean
'(rfacsoc_calidad)
Dim Precio As Currency
Dim SQL As String
Dim ImpLinea As Currency
Dim ImporteSinBonif As String
Dim PrecioSinBonif As String
    
    On Error GoTo eInsertLinea
    
    MensError = ""
    
    InsertLineaCalidad = False
    
    Precio = 0
    If CCur(ImporteSinFormato(Kilos)) <> 0 Then
        Precio = Round2(CCur(ImporteSinFormato(Importe)) / CCur(ImporteSinFormato(Kilos)), 4)
    End If
    
    '[Monica] 27/01/2010 : la bonificacion me viene de anticipos liquidaciones de Picassent
    If Bonificacion <> "" Then
        If CCur(ImporteSinFormato(Bonificacion)) <> 0 Then
            ImporteSinBonif = CCur(ImporteSinFormato(Importe)) - CCur(ImporteSinFormato(Bonificacion))
            PrecioSinBonif = Round2(CCur(ImporteSinFormato(ImporteSinBonif)) / CCur(ImporteSinFormato(Kilos)), 4)
        End If
    End If
    
    'insertamos la calidad
    SQL = "insert into tmpfact_calidad (codtipom, numfactu, fecfactu, codvarie, codcampo, "
    SQL = SQL & "codcalid, kilosnet, precio, imporcal, preciocalidad, imporcalidad) values ("
    SQL = SQL & "'" & tipoMov & "'," & DBSet(numfactu, "N") & "," & DBSet(FecFac, "F") & ","
    SQL = SQL & DBSet(Variedad, "N") & "," & DBSet(campo, "N") & ","
    SQL = SQL & DBSet(Calidad, "N") & "," & DBSet(Kilos, "N") & ","
    SQL = SQL & DBSet(Precio, "N") & "," & DBSet(Importe, "N") & ","
    
    If Bonificacion <> "" Then
        If CCur(ImporteSinFormato(Bonificacion)) <> 0 Then
            SQL = SQL & DBSet(PrecioSinBonif, "N") & "," & DBSet(ImporteSinBonif, "N") & ")"
        Else
            SQL = SQL & DBSet(Precio, "N") & "," & DBSet(Importe, "N") & ")"
        End If
    Else
        SQL = SQL & DBSet(Precio, "N") & "," & DBSet(Importe, "N") & ")"
    End If
        
    conn.Execute SQL
    InsertLineaCalidad = True
    Exit Function
    
eInsertLinea:
    If Err.Number <> 0 Then
        MensError = "Se ha producido un error en la inserción de la linea de calidad de factura "
        MuestraError Err.Number, MensError, Err.Description
    End If
End Function


'Insertar Linea de factura (albaran)
Public Function InsertLineaAlbaran(tipoMov As String, numfactu As String, FecFac As String, ByRef Rs As ADODB.Recordset, Precio As String, Importe As String, Optional codCampo As String) As Boolean
'(rfactsoc_albaran)
'codcampo tiene valor cuando venimos de almazara que hemos tenido que buscarlo porque en el cursor Rs no lo tenemos
Dim GastosAlb As Currency
Dim Tipo As String

    Dim SQL As String
    Dim ImpLinea As Currency
    
    On Error GoTo eInsertLinea
    
    MensError = ""
    
    InsertLineaAlbaran = False
    
    Tipo = DevuelveDesdeBDNew(cAgro, "usuarios.stipom", "tipodocu", "codtipom", tipoMov, "T")
    If CInt(Tipo) = 7 Then ' si se trata de un anticipo de almazara no descontamos gastos
        GastosAlb = 0
    Else
        GastosAlb = DevuelveValor("select sum(importe) from rhisfruta_gastos where numalbar = " & DBSet(Rs!numalbar, "N"))
    End If
    
    'insertamos el albaran
    SQL = "insert into tmpfact_albaran (codtipom, numfactu, fecfactu, numalbar, fecalbar, "
    SQL = SQL & "codvarie, codcampo, kilosbru, kilosnet, grado, precio, importe, imporgasto) values ("
    SQL = SQL & "'" & tipoMov & "'," & DBSet(numfactu, "N") & "," & DBSet(FecFac, "F") & ","
    SQL = SQL & DBSet(Rs!numalbar, "N") & "," & DBSet(Rs!Fecalbar, "F") & "," & DBSet(Rs!Codvarie, "N") & ","
    
    If Not IsNull(codCampo) And codCampo <> "" Then
        SQL = SQL & DBSet(codCampo, "N") & ","
    Else
        SQL = SQL & DBSet(Rs!codCampo, "N") & ","
    End If
    SQL = SQL & DBSet(Rs!KilosBru, "N") & "," & DBSet(Rs!KilosNet, "N") & ","
    SQL = SQL & DBSet(Rs!PrEstimado, "N") & "," & DBSet(Precio, "N") & "," & DBSet(Importe, "N") & ","
    SQL = SQL & DBSet(GastosAlb, "N") & ")"
    
    conn.Execute SQL
    InsertLineaAlbaran = True
    Exit Function
    
eInsertLinea:
    If Err.Number <> 0 Then
        MensError = "Se ha producido un error en la inserción de la linea de albaran de factura "
        MuestraError Err.Number, MensError, Err.Description
    End If
End Function


'Insertar facturas varias del socio
Public Function InsertFacturasVarias(tipoMov As String, numfactu As String, FecFac As String, AntLiq As Byte, EsVC As Byte) As Boolean
'AntLiq: 0 = anticipo
'        1 = liquidacion
'EsVC  : 0 = entrada normal ( no de vc)
'        1 = entrada venta campo

    Dim SQL As String
    
    On Error GoTo eInsertCabe
    
    MensError = ""
    InsertFacturasVarias = False
    
    'insertamos las facturas varias
    SQL = "insert into tmpfact_fvarias (codtipom,numfactu,fecfactu,codtipomfvar,numfactufvar,fecfactufvar,codsecci) select "
    SQL = SQL & "'" & tipoMov & "'," & DBSet(numfactu, "N") & "," & DBSet(FecFac, "F") & ","
    SQL = SQL & "codtipom,numfactu,fecfactu,codsecci from fvarcabfact where codsocio = " & DBSet(vSocio.Codigo, "N")
    
    If AntLiq = 0 Then
        ' en anticipo
        SQL = SQL & " and ((enliquidacion = 2 "
    Else
        ' en liquidacion
        SQL = SQL & " and ((enliquidacion = 1 "
    End If
    
    If EsVC = 0 Then
        ' entrada normal
        SQL = SQL & " and envtacampo = 0) "
    Else
        ' entrada venta campo
        SQL = SQL & " and envtacampo = 1) "
    End If
    
    '[Monica]30/11/2017: añado parentesis y lo de encualquier fra
    SQL = SQL & " or enliquidacion = 3) "
    
    SQL = SQL & " and intliqui = 0"
    
    conn.Execute SQL
    
    ' las marcamos como que han sido descontadas
    SQL = "update fvarcabfact set intliqui = 1 where codsocio = " & DBSet(vSocio.Codigo, "N")
    
    If AntLiq = 0 Then
        ' en anticipo
        SQL = SQL & " and ((enliquidacion = 2 "
    Else
        ' en liquidacion
        SQL = SQL & " and ((enliquidacion = 1 "
    End If
    
    If EsVC = 0 Then
        ' entrada normal
        SQL = SQL & " and envtacampo = 0) "
    Else
        ' entrada venta campo
        SQL = SQL & " and envtacampo = 1) "
    End If
    
    '[Monica]30/11/2017: añado parentesis y lo de encualquier fra
    SQL = SQL & " or enliquidacion = 3) "
    
    
    SQL = SQL & " and intliqui = 0"
    
    conn.Execute SQL
    
    InsertFacturasVarias = True
    
    Exit Function

eInsertCabe:
    MensError = "Se ha producido un error en la inserción de la linea de fvarias del socio " & vSocio.Codigo
    MuestraError Err.Number, MensError
End Function


'Insertar Resumen
Public Function InsertResumen(Tipo As String, numfactu As String, Optional Trans As String) As Boolean

    Dim SQL As String
    
    On Error GoTo eInsertResumen
    
    MensError = ""
    InsertResumen = False
    
    If Trans = "" Then
                                            ' codtipom, numfactu
        SQL = "insert into tmpinformes (codusu, nombre1, importe1) values ( " & vUsu.Codigo
        SQL = SQL & ",'" & Tipo & "'," & DBSet(numfactu, "N") & ")"
    
    Else
                                            ' codtipom, numfactu, codtrans
        SQL = "insert into tmpinformes (codusu, nombre1, importe1, nombre2) values ( " & vUsu.Codigo
        SQL = SQL & ",'" & Tipo & "'," & DBSet(numfactu, "N") & "," & DBSet(Trans, "T") & ")"
    
    End If
    
    conn.Execute SQL
    
    InsertResumen = True
    
    Exit Function

eInsertResumen:
    MensError = "Error en la inserción de la factura " & numfactu & " en el Resumen "
    MuestraError Err.Number, MensError
End Function


Public Function ExisteEnHistorico(cDesde As String, cHasta As String, ctipo As String) As Boolean
Dim SQL As String
Dim Tipo As String

    ExisteEnHistorico = False
    
    SQL = "select count(*) from slhfac, scaalb where letraser = " & DBSet(Tipo, "T") & " and " & _
          " slhfac.numfactu= scaalb.numfactu and slhfac.numlinea = scaalb.numlinea "
    
    If cDesde <> "" Then SQL = SQL & " and scaalb.fecalbar >= '" & Format(cDesde, FormatoFecha) & "' "
    If cHasta <> "" Then SQL = SQL & " and scaalb.fecalbar <= '" & Format(cHasta, FormatoFecha) & "' "

    ExisteEnHistorico = (TotalRegistros(SQL) <> 0)
    
End Function


Public Function FacturacionAnticiposValsur(cTabla As String, cWhere As String, FecFac As String, Pb1 As ProgressBar) As Boolean
Dim SQL As String
Dim Rs As ADODB.Recordset

Dim AntSocio As String
Dim AntVarie As String
Dim ActSocio As String
Dim ActVarie As String
Dim actCampo As String
Dim AntCampo As String
Dim ActCalid As String
Dim AntCalid As String

Dim HayReg As Boolean

Dim NumError As Long
Dim vImporte As Currency
Dim vPorcIva As String

Dim PrimFac As String
Dim UltFac As String

Dim tipoMov As String
Dim B As Boolean
Dim vSeccion As CSeccion
Dim Kilos As Currency
Dim KilosCal As Currency
Dim Importe As Currency

Dim devuelve As String
Dim Existe As Boolean

'Dim baseimpo As Currency
'Dim BaseReten As Currency
Dim Neto As Currency
'Dim ImpoIva As Currency
'Dim ImpoReten As Currency
'Dim TotalFac As Currency
Dim Recolect As Byte
Dim vPrecio As Currency

    On Error GoTo eFacturacion

    FacturacionAnticiposValsur = False
    
    tipoMov = "FAA"
    
    BorrarTMPs
    B = CrearTMPs()
    If Not B Then
         Exit Function
    End If
    
    conn.BeginTrans
    
    
    SQL = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute SQL

    SQL = "SELECT  rhisfruta.codsocio, rhisfruta.codvarie,"
    SQL = SQL & "rhisfruta.codcampo, rhisfruta.recolect, rhisfruta_clasif.codcalid, "
'[Monica]01/09/2010 : sustituida la siguiente linea por
'    Sql = Sql & "rprecios_calidad.precoop, rprecios_calidad.presocio, rprecios_calidad.tipofact,sum(rhisfruta_clasif.kilosnet) as kilosnet "
     SQL = SQL & "rprecios.fechaini, rprecios.fechafin, rprecios_calidad.tipofact,max(rprecios.contador) contador, sum(rhisfruta_clasif.kilosnet) as kilosnet "
    SQL = SQL & " FROM  " & cTabla

    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        SQL = SQL & " WHERE " & cWhere
    End If
    ' ordenado por socio, variedad, campo, calidad
    SQL = SQL & " group by rhisfruta.codsocio, rhisfruta.codvarie, rhisfruta.codcampo, rhisfruta_clasif.codcalid, rhisfruta.recolect "
    SQL = SQL & " order by rhisfruta.codsocio, rhisfruta.codvarie, rhisfruta.codcampo, rhisfruta_clasif.codcalid, rhisfruta.recolect "
    
    Set vSeccion = New CSeccion
    
    If vSeccion.LeerDatos(vParamAplic.Seccionhorto) Then
        If Not vSeccion.AbrirConta Then
            Exit Function
        End If
    End If
    
    HayReg = False
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    If Not Rs.EOF Then
        AntSocio = CStr(DBLet(Rs!Codsocio, "N"))
        AntVarie = CStr(DBLet(Rs!Codvarie, "N"))
        AntCampo = CStr(DBLet(Rs!codCampo, "N"))
        AntCalid = CStr(DBLet(Rs!codcalid, "N"))
        
        ActSocio = CStr(DBLet(Rs!Codsocio, "N"))
        ActVarie = CStr(DBLet(Rs!Codvarie, "N"))
        actCampo = CStr(DBLet(Rs!codCampo, "N"))
        ActCalid = CStr(DBLet(Rs!codcalid, "N"))
    
        Set vSocio = New cSocio
        If vSocio.LeerDatos(ActSocio) Then
            If vSocio.LeerDatosSeccion(ActSocio, vParamAplic.Seccionhorto) Then
                baseimpo = 0
                BaseReten = 0
                ImpoIva = 0
                ImpoReten = 0
                TotalFac = 0
                BaseAFO = 0
                ImpoAFO = 0
                PorcAFO = 0
                
                Kilos = 0
                Importe = 0
                
                KilosCal = 0
                
                vPorcIva = ""
                vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vSocio.CodIva, "N")
                PorcIva = CCur(ImporteSinFormato(vPorcIva))
                
                tipoMov = vSocio.CodTipomAnt
                
                Set vTipoMov = New CTiposMov
                
                numfactu = vTipoMov.ConseguirContador(tipoMov)
                Do
                    numfactu = vTipoMov.ConseguirContador(tipoMov)
                    devuelve = DevuelveDesdeBDNew(cAgro, "rfactsoc", "numfactu", "codtipom", tipoMov, "T", , "numfactu", CStr(numfactu), "N", "fecfactu", FecFac, "F")
                    If devuelve <> "" Then
                        'Ya existe el contador incrementarlo
                        Existe = True
                        vTipoMov.IncrementarContador (tipoMov)
                        numfactu = vTipoMov.ConseguirContador(tipoMov)
                    Else
                        Existe = False
                    End If
                Loop Until Not Existe
                
                vParamAplic.PrimFactAnt = numfactu
                
            End If
        End If
    End If
    
    While Not Rs.EOF And B
        ActCalid = DBLet(Rs!codcalid, "N")
        ActVarie = DBLet(Rs!Codvarie, "N")
        actCampo = DBSet(Rs!codCampo, "N")
        ActSocio = DBSet(Rs!Codsocio, "N")
        
        If (ActCalid <> AntCalid Or AntCampo <> actCampo Or AntVarie <> ActVarie Or AntSocio <> ActSocio) Then
            ' kilos e importe por variedad campo
            Kilos = Kilos + KilosCal
            Importe = Importe + vImporte
            
            baseimpo = baseimpo + vImporte
            
            B = InsertLineaCalidad(tipoMov, CStr(numfactu), FecFac, AntVarie, AntCampo, CStr(AntCalid), CStr(DBLet(KilosCal, "N")), CStr(vImporte))
            KilosCal = 0
            vImporte = 0
            
            AntCalid = ActCalid
        End If
        
        If (ActVarie <> AntVarie Or actCampo <> AntCampo Or ActSocio <> AntSocio) Then
            ' insertar linea de variedad, campo
            B = InsertLinea(tipoMov, CStr(numfactu), FecFac, AntVarie, AntCampo, CStr(Kilos), CStr(Importe), "0")
            
            If B Then
                AntVarie = ActVarie
                AntCampo = actCampo
                Kilos = 0
                Importe = 0
            End If
        End If
        
        If ActSocio <> AntSocio Then
            ImpoIva = Round2(baseimpo * PorcIva / 100, 2)
        
            Select Case DBLet(vSocio.TipoIRPF, "N")
                Case 0
                    ImpoReten = Round2((baseimpo + ImpoIva) * vParamAplic.PorcreteFacSoc / 100, 2)
                    BaseReten = (baseimpo + ImpoIva)
                    PorcReten = vParamAplic.PorcreteFacSoc
                Case 1
                    ImpoReten = Round2(baseimpo * vParamAplic.PorcreteFacSoc / 100, 2)
                    BaseReten = baseimpo
                    PorcReten = vParamAplic.PorcreteFacSoc
                Case 2
                    ImpoReten = 0
                    BaseReten = 0
                    PorcReten = 0
            End Select
        
            TotalFac = baseimpo + ImpoIva - ImpoReten - ImpoAFO
            
            IncrementarProgresNew Pb1, 1
            
            'insertar cabecera de factura
            B = InsertCabecera(tipoMov, CStr(numfactu), FecFac)
            
            If B Then B = InsertResumen(tipoMov, CStr(numfactu))
            
            If B Then B = vTipoMov.IncrementarContador(tipoMov)
            
            If B Then
                AntSocio = ActSocio
                
                Set vSocio = Nothing
                Set vSocio = New cSocio
                If vSocio.LeerDatos(ActSocio) Then
                    If vSocio.LeerDatosSeccion(AntSocio, vParamAplic.Seccionhorto) Then
                        vPorcIva = ""
                        vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vSocio.CodIva, "N")
                        PorcIva = CCur(ImporteSinFormato(vPorcIva))
                    End If
                    
                    tipoMov = vSocio.CodTipomAnt
                End If
                baseimpo = 0
                BaseReten = 0
                Neto = 0
                ImpoIva = 0
                ImpoReten = 0
                TotalFac = 0
                BaseAFO = 0
                ImpoAFO = 0
                
                numfactu = vTipoMov.ConseguirContador(tipoMov)
                Do
                    numfactu = vTipoMov.ConseguirContador(tipoMov)
                    devuelve = DevuelveDesdeBDNew(cAgro, "rfactsoc", "numfactu", "codtipom", tipoMov, "T", , "numfactu", CStr(numfactu), "N", "fecfactu", FecFac, "F")
                    If devuelve <> "" Then
                        'Ya existe el contador incrementarlo
                        Existe = True
                        vTipoMov.IncrementarContador (tipoMov)
                        numfactu = vTipoMov.ConseguirContador(tipoMov)
                    Else
                        Existe = False
                    End If
                Loop Until Not Existe
           End If
        End If
        
        Recolect = DBLet(Rs!Recolect, "N")
        
        '[Monica]01/09/2010: añadido ésto, antes los precios los sacabamos en el propio select
        Dim Sql9 As String
        Dim Rs9 As ADODB.Recordset
        Dim PreCoop As Currency
        Dim PreSocio As Currency
        
        Sql9 = "select precoop, presocio from rprecios_calidad where codvarie = " & DBSet(Rs!Codvarie, "N")
        Sql9 = Sql9 & " and tipofact = " & DBSet(Rs!TipoFact, "N")
        Sql9 = Sql9 & " and contador = " & DBSet(Rs!Contador, "N")
        Sql9 = Sql9 & " and codcalid = " & DBSet(Rs!codcalid, "N")
        
        Set Rs9 = New ADODB.Recordset
        Rs9.Open Sql9, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
        
        If Not Rs9.EOF Then
            PreCoop = DBLet(Rs9.Fields(0).Value, "N")
            PreSocio = DBLet(Rs9.Fields(1).Value, "N")
        
            Select Case Recolect
                Case 0
                    vPrecio = DBLet(PreCoop, "N")
                    vImporte = vImporte + Round2(DBLet(Rs!KilosNet, "N") * PreCoop, 2)
                Case 1
                    vPrecio = DBLet(PreSocio, "N")
                    vImporte = vImporte + Round2(DBLet(Rs!KilosNet, "N") * PreSocio, 2)
            End Select
            
            KilosCal = KilosCal + DBLet(Rs!KilosNet, "N")
            
        End If
        
        Set Rs9 = Nothing
        
        'hasta aqui
        HayReg = True
        
        Rs.MoveNext
    Wend
    
    ' ultimo registro si ha entrado
    If B And HayReg Then
        ' insertar linea de calidad
        Kilos = Kilos + KilosCal
        Importe = Importe + vImporte
        
        baseimpo = baseimpo + vImporte
        
        If B Then B = InsertLineaCalidad(tipoMov, CStr(numfactu), FecFac, ActVarie, actCampo, CStr(ActCalid), CStr(DBLet(KilosCal, "N")), CStr(vImporte))
        
        ' insertar linea de variedad
        If B Then B = InsertLinea(tipoMov, CStr(numfactu), FecFac, CStr(ActVarie), CStr(actCampo), CStr(Kilos), CStr(Importe), "0")
        
        ImpoIva = Round2(baseimpo * PorcIva / 100, 2)

        Select Case DBLet(vSocio.TipoIRPF, "N")
            Case 0
                ImpoReten = Round2((baseimpo + ImpoIva) * vParamAplic.PorcreteFacSoc / 100, 2)
                BaseReten = (baseimpo + ImpoIva)
                PorcReten = vParamAplic.PorcreteFacSoc
            Case 1
                ImpoReten = Round2(baseimpo * vParamAplic.PorcreteFacSoc / 100, 2)
                BaseReten = baseimpo
                PorcReten = vParamAplic.PorcreteFacSoc
            Case 2
                ImpoReten = 0
                BaseReten = 0
                PorcReten = 0
        End Select
        
'        BaseAFO = baseimpo
'        PorcAFO = vParamAplic.PorcenAFO
'        ImpoAFO = Round2(BaseAFO * PorcAFO / 100, 2)

        TotalFac = baseimpo + ImpoIva - ImpoReten - ImpoAFO
        
        IncrementarProgresNew Pb1, 1
        
        vParamAplic.UltFactAnt = numfactu
        
        'insertar cabecera de factura
        B = InsertCabecera(tipoMov, CStr(numfactu), FecFac)
        
        If B Then B = InsertResumen(tipoMov, CStr(numfactu))
        
        If B Then B = vTipoMov.IncrementarContador(tipoMov)
        
        'pasamos las temporales a las tablas
        If B Then B = PasarTemporales()
        
        If B Then B = (vParamAplic.Modificar = 1)
    End If
    
'    BorrarTMPs
    
    vSeccion.CerrarConta
    Set vSeccion = Nothing
    Set vSocio = Nothing
    
eFacturacion:
    If Err.Number <> 0 Or Not B Then
        conn.RollbackTrans
        FacturacionAnticiposValsur = False
    Else
        conn.CommitTrans
        FacturacionAnticiposValsur = True
    End If
End Function


'[Monica]20/01/2012: alzira no ha hecho hasta el momento anticipos
'                    Nueva funcion de anticipos para alzira
Public Function FacturacionAnticiposAlzira(cTabla As String, cWhere As String, FecFac As String, Pb1 As ProgressBar) As Boolean
Dim SQL As String
Dim Rs As ADODB.Recordset

Dim AntSocio As String
Dim AntVarie As String
Dim ActSocio As String
Dim ActVarie As String
Dim actCampo As String
Dim AntCampo As String
Dim ActCalid As String
Dim AntCalid As String

Dim HayReg As Boolean

Dim NumError As Long
Dim vImporte As Currency
Dim vPorcIva As String

Dim PrimFac As String
Dim UltFac As String

Dim tipoMov As String
Dim B As Boolean
Dim vSeccion As CSeccion
Dim Kilos As Currency
Dim KilosCal As Currency
Dim Importe As Currency

Dim devuelve As String
Dim Existe As Boolean

'Dim baseimpo As Currency
'Dim BaseReten As Currency
Dim Neto As Currency
'Dim ImpoIva As Currency
'Dim ImpoReten As Currency
'Dim TotalFac As Currency
Dim Recolect As Byte
Dim vPrecio As Currency

    On Error GoTo eFacturacion

    FacturacionAnticiposAlzira = False
    
    tipoMov = "FAA"
    
    BorrarTMPs
    B = CrearTMPs()
    If Not B Then
         Exit Function
    End If
    
    conn.BeginTrans
    
    
    SQL = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute SQL

    SQL = "SELECT  rhisfruta.codsocio, rhisfruta.codvarie,"
    SQL = SQL & "rhisfruta.codcampo, rhisfruta.recolect, rhisfruta_clasif.codcalid, "
'[Monica]01/09/2010 : sustituida la siguiente linea por
'    Sql = Sql & "rprecios_calidad.precoop, rprecios_calidad.presocio, rprecios_calidad.tipofact,sum(rhisfruta_clasif.kilosnet) as kilosnet "
     SQL = SQL & "rprecios.fechaini, rprecios.fechafin, rprecios_calidad.tipofact,max(rprecios.contador) contador, sum(rhisfruta_clasif.kilosnet) as kilosnet "
    SQL = SQL & " FROM  " & cTabla

    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        SQL = SQL & " WHERE " & cWhere
    End If
    ' ordenado por socio, variedad, campo, calidad
    SQL = SQL & " group by rhisfruta.codsocio, rhisfruta.codvarie, rhisfruta.codcampo, rhisfruta_clasif.codcalid, rhisfruta.recolect "
    SQL = SQL & " order by rhisfruta.codsocio, rhisfruta.codvarie, rhisfruta.codcampo, rhisfruta_clasif.codcalid, rhisfruta.recolect "
    
    Set vSeccion = New CSeccion
    
    If vSeccion.LeerDatos(vParamAplic.Seccionhorto) Then
        If Not vSeccion.AbrirConta Then
            Exit Function
        End If
    End If
    
    HayReg = False
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    If Not Rs.EOF Then
        AntSocio = CStr(DBLet(Rs!Codsocio, "N"))
        AntVarie = CStr(DBLet(Rs!Codvarie, "N"))
        AntCampo = CStr(DBLet(Rs!codCampo, "N"))
        AntCalid = CStr(DBLet(Rs!codcalid, "N"))
        
        ActSocio = CStr(DBLet(Rs!Codsocio, "N"))
        ActVarie = CStr(DBLet(Rs!Codvarie, "N"))
        actCampo = CStr(DBLet(Rs!codCampo, "N"))
        ActCalid = CStr(DBLet(Rs!codcalid, "N"))
    
        Set vSocio = New cSocio
        If vSocio.LeerDatos(ActSocio) Then
            If vSocio.LeerDatosSeccion(ActSocio, vParamAplic.Seccionhorto) Then
                baseimpo = 0
                BaseReten = 0
                ImpoIva = 0
                ImpoReten = 0
                TotalFac = 0
                BaseAFO = 0
                ImpoAFO = 0
                PorcAFO = 0
                
                Kilos = 0
                Importe = 0
                
                KilosCal = 0
                
                vPorcIva = ""
                '[Monica]29/04/2011: INTERNAS
                If vSocio.EsFactADVInt Then
                    vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vParamAplic.CodIvaExeADV, "N")
                Else
                    vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vSocio.CodIva, "N")
                End If
                PorcIva = CCur(ImporteSinFormato(vPorcIva))
                
                tipoMov = vSocio.CodTipomAnt
                
                Set vTipoMov = New CTiposMov
                
                numfactu = vTipoMov.ConseguirContador(tipoMov)
                Do
                    numfactu = vTipoMov.ConseguirContador(tipoMov)
                    devuelve = DevuelveDesdeBDNew(cAgro, "rfactsoc", "numfactu", "codtipom", tipoMov, "T", , "numfactu", CStr(numfactu), "N", "fecfactu", FecFac, "F")
                    If devuelve <> "" Then
                        'Ya existe el contador incrementarlo
                        Existe = True
                        vTipoMov.IncrementarContador (tipoMov)
                        numfactu = vTipoMov.ConseguirContador(tipoMov)
                    Else
                        Existe = False
                    End If
                Loop Until Not Existe
                
                vParamAplic.PrimFactAnt = numfactu
                
            End If
        End If
    End If
    
    While Not Rs.EOF And B
        ActCalid = DBLet(Rs!codcalid, "N")
        ActVarie = DBLet(Rs!Codvarie, "N")
        actCampo = DBSet(Rs!codCampo, "N")
        ActSocio = DBSet(Rs!Codsocio, "N")
        
        If (ActCalid <> AntCalid Or AntCampo <> actCampo Or AntVarie <> ActVarie Or AntSocio <> ActSocio) Then
            ' kilos e importe por variedad campo
            Kilos = Kilos + KilosCal
            Importe = Importe + vImporte
            
            baseimpo = baseimpo + vImporte
            
            B = InsertLineaCalidad(tipoMov, CStr(numfactu), FecFac, AntVarie, AntCampo, CStr(AntCalid), CStr(DBLet(KilosCal, "N")), CStr(vImporte))
            KilosCal = 0
            vImporte = 0
            
            AntCalid = ActCalid
        End If
        
        If (ActVarie <> AntVarie Or actCampo <> AntCampo Or ActSocio <> AntSocio) Then
            '[Monica]08/04/2010: grabamos los albaranes que intervienen en la linea de factura
            If B Then
                B = InsertarAlbaranesFactura(tipoMov, CStr(numfactu), FecFac, AntSocio, AntVarie, AntCampo, cTabla, cWhere, 0)
            End If
            
            ' insertar linea de variedad, campo
            B = InsertLinea(tipoMov, CStr(numfactu), FecFac, AntVarie, AntCampo, CStr(Kilos), CStr(Importe), "0")
            
            If B Then
                AntVarie = ActVarie
                AntCampo = actCampo
                Kilos = 0
                Importe = 0
            End If
        End If
        
        If ActSocio <> AntSocio Then
            ImpoIva = Round2(baseimpo * PorcIva / 100, 2)
        
            Select Case DBLet(vSocio.TipoIRPF, "N")
                Case 0
                    ImpoReten = Round2((baseimpo + ImpoIva) * vParamAplic.PorcreteFacSoc / 100, 2)
                    BaseReten = (baseimpo + ImpoIva)
                    PorcReten = vParamAplic.PorcreteFacSoc
                Case 1
                    ImpoReten = Round2(baseimpo * vParamAplic.PorcreteFacSoc / 100, 2)
                    BaseReten = baseimpo
                    PorcReten = vParamAplic.PorcreteFacSoc
                Case 2
                    ImpoReten = 0
                    BaseReten = 0
                    PorcReten = 0
            End Select
        
            TotalFac = baseimpo + ImpoIva - ImpoReten - ImpoAFO
            
            IncrementarProgresNew Pb1, 1
            
            'insertar cabecera de factura
            B = InsertCabecera(tipoMov, CStr(numfactu), FecFac)
            
            '[Monica]04/01/2012: marcamos la factura como contabilizada y como pdte de recibir el nro de factura
            If B And vSocio.EmiteFact Then B = MarcarFactura(tipoMov, CStr(numfactu), FecFac, True)
            
            If B Then B = InsertResumen(tipoMov, CStr(numfactu))
            
'Mirar si quito lo de reclacular calidades
            If B Then B = RecalcularCalidades(tipoMov, CStr(numfactu), FecFac)
            
'Recalculo de todos los importes de tmpfact_calidades y tmpfact_variedades para que cuadre con la base de cabecera
            If B Then B = CuadrarBasesFactura(tipoMov, CStr(numfactu), FecFac, baseimpo)
            
            If B Then B = vTipoMov.IncrementarContador(tipoMov)
            
            If B Then
                AntSocio = ActSocio
                
                Set vSocio = Nothing
                Set vSocio = New cSocio
                If vSocio.LeerDatos(ActSocio) Then
                    If vSocio.LeerDatosSeccion(AntSocio, vParamAplic.Seccionhorto) Then
                        vPorcIva = ""
                        '[Monica]29/04/2011: INTERNAS
                        If vSocio.EsFactADVInt Then
                            vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vParamAplic.CodIvaExeADV, "N")
                        Else
                            vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vSocio.CodIva, "N")
                        End If
                        
                        PorcIva = CCur(ImporteSinFormato(vPorcIva))
                    
                    End If
                    
                    tipoMov = vSocio.CodTipomAnt
                End If
                baseimpo = 0
                BaseReten = 0
                Neto = 0
                ImpoIva = 0
                ImpoReten = 0
                TotalFac = 0
                BaseAFO = 0
                ImpoAFO = 0
                
                numfactu = vTipoMov.ConseguirContador(tipoMov)
                Do
                    numfactu = vTipoMov.ConseguirContador(tipoMov)
                    devuelve = DevuelveDesdeBDNew(cAgro, "rfactsoc", "numfactu", "codtipom", tipoMov, "T", , "numfactu", CStr(numfactu), "N", "fecfactu", FecFac, "F")
                    If devuelve <> "" Then
                        'Ya existe el contador incrementarlo
                        Existe = True
                        vTipoMov.IncrementarContador (tipoMov)
                        numfactu = vTipoMov.ConseguirContador(tipoMov)
                    Else
                        Existe = False
                    End If
                Loop Until Not Existe
           End If
        End If
        
        Recolect = DBLet(Rs!Recolect, "N")
        
        '[Monica]01/09/2010: añadido ésto, antes los precios los sacabamos en el propio select
        Dim Sql9 As String
        Dim Rs9 As ADODB.Recordset
        Dim PreCoop As Currency
        Dim PreSocio As Currency
        
        Sql9 = "select precoop, presocio from rprecios_calidad where codvarie = " & DBSet(Rs!Codvarie, "N")
        Sql9 = Sql9 & " and tipofact = " & DBSet(Rs!TipoFact, "N")
        Sql9 = Sql9 & " and contador = " & DBSet(Rs!Contador, "N")
        Sql9 = Sql9 & " and codcalid = " & DBSet(Rs!codcalid, "N")
        
        Set Rs9 = New ADODB.Recordset
        Rs9.Open Sql9, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
        
        If Not Rs9.EOF Then
            PreCoop = DBLet(Rs9.Fields(0).Value, "N")
            PreSocio = DBLet(Rs9.Fields(1).Value, "N")
        
            Select Case Recolect
'                Case 0
'                    vPrecio = DBLet(PreCoop, "N")
'                    vImporte = vImporte + Round2(DBLet(Rs!KilosNet, "N") * PreCoop, 2)
'                Case 1
'                    vPrecio = DBLet(PreSocio, "N")
'                    vImporte = vImporte + Round2(DBLet(Rs!KilosNet, "N") * PreSocio, 2)
            
                Case 0
                    vPrecio = DBLet(PreCoop, "N")
                    vImporte = vImporte + (DBLet(Rs!KilosNet, "N") * PreCoop)
                Case 1
                    vPrecio = DBLet(PreSocio, "N")
                    vImporte = vImporte + (DBLet(Rs!KilosNet, "N") * PreSocio)
            End Select
            
            KilosCal = KilosCal + DBLet(Rs!KilosNet, "N")
            
        End If
        
        Set Rs9 = Nothing
        
        'hasta aqui
        HayReg = True
        
        Rs.MoveNext
    Wend
    
    ' ultimo registro si ha entrado
    If B And HayReg Then
        ' insertar linea de calidad
        Kilos = Kilos + KilosCal
        Importe = Importe + vImporte
        
        baseimpo = baseimpo + vImporte
        
        If B Then B = InsertLineaCalidad(tipoMov, CStr(numfactu), FecFac, ActVarie, actCampo, CStr(ActCalid), CStr(DBLet(KilosCal, "N")), CStr(vImporte))
        
        '[Monica]08/04/2010: grabamos los albaranes que intervienen en la linea de factura
        If B Then
            B = InsertarAlbaranesFactura(tipoMov, CStr(numfactu), FecFac, AntSocio, AntVarie, AntCampo, cTabla, cWhere, 0)
        End If
        
        ' insertar linea de variedad
        If B Then B = InsertLinea(tipoMov, CStr(numfactu), FecFac, CStr(ActVarie), CStr(actCampo), CStr(Kilos), CStr(Importe), "0")
        
        ImpoIva = Round2(baseimpo * PorcIva / 100, 2)

        Select Case DBLet(vSocio.TipoIRPF, "N")
            Case 0
                ImpoReten = Round2((baseimpo + ImpoIva) * vParamAplic.PorcreteFacSoc / 100, 2)
                BaseReten = (baseimpo + ImpoIva)
                PorcReten = vParamAplic.PorcreteFacSoc
            Case 1
                ImpoReten = Round2(baseimpo * vParamAplic.PorcreteFacSoc / 100, 2)
                BaseReten = baseimpo
                PorcReten = vParamAplic.PorcreteFacSoc
            Case 2
                ImpoReten = 0
                BaseReten = 0
                PorcReten = 0
        End Select
        
'        BaseAFO = baseimpo
'        PorcAFO = vParamAplic.PorcenAFO
'        ImpoAFO = Round2(BaseAFO * PorcAFO / 100, 2)

        TotalFac = baseimpo + ImpoIva - ImpoReten - ImpoAFO
        
        IncrementarProgresNew Pb1, 1
        
        vParamAplic.UltFactAnt = numfactu
        
        'insertar cabecera de factura
        B = InsertCabecera(tipoMov, CStr(numfactu), FecFac)
        
        '[Monica]04/01/2012: marcamos la factura como contabilizada y como pdte de recibir el nro de factura
        If B And vSocio.EmiteFact Then B = MarcarFactura(tipoMov, CStr(numfactu), FecFac, True)
        
        If B Then B = InsertResumen(tipoMov, CStr(numfactu))
        
'Mirar si quito lo de reclacular calidades
        If B Then B = RecalcularCalidades(tipoMov, CStr(numfactu), FecFac)
        
'Recalculo de todos los importes de rfactsoc_calidades y rfactsoc_variedades para que cuadre con la base de cabecera
        If B Then B = CuadrarBasesFactura(tipoMov, CStr(numfactu), FecFac, baseimpo)
        
        
        If B Then B = vTipoMov.IncrementarContador(tipoMov)
        
        'pasamos las temporales a las tablas
        If B Then B = PasarTemporales()
        
        If B Then B = (vParamAplic.Modificar = 1)
    End If
    
'    BorrarTMPs
    
    vSeccion.CerrarConta
    Set vSeccion = Nothing
    Set vSocio = Nothing
    
eFacturacion:
    If Err.Number <> 0 Or Not B Then
        conn.RollbackTrans
        FacturacionAnticiposAlzira = False
    Else
        conn.CommitTrans
        FacturacionAnticiposAlzira = True
    End If
End Function




Public Function FacturacionAnticiposPicassent(cTabla As String, cWhere As String, FecFac As String, Pb1 As ProgressBar, DescontarFVarias As Boolean) As Boolean
Dim SQL As String
Dim Rs As ADODB.Recordset

Dim AntSocio As String
Dim AntVarie As String
Dim ActSocio As String
Dim ActVarie As String
Dim actCampo As String
Dim AntCampo As String
Dim ActCalid As String
Dim AntCalid As String

Dim HayReg As Boolean

Dim NumError As Long
Dim vImporte As Currency
Dim vPorcIva As String

Dim PrimFac As String
Dim UltFac As String

Dim tipoMov As String
Dim B As Boolean
Dim vSeccion As CSeccion
Dim Kilos As Currency
Dim KilosCal As Currency
Dim Importe As Currency
Dim Bonifica As Currency

Dim devuelve As String
Dim Existe As Boolean

'Dim baseimpo As Currency
'Dim BaseReten As Currency
Dim Neto As Currency
'Dim ImpoIva As Currency
'Dim ImpoReten As Currency
'Dim TotalFac As Currency
Dim Recolect As Byte
Dim vPrecio As Currency
Dim vBonifica As Currency
Dim PorcBoni As Currency
Dim PorcComi As Currency



    On Error GoTo eFacturacion

    FacturacionAnticiposPicassent = False
    
    tipoMov = "FAA"
    
    BorrarTMPs
    B = CrearTMPs()
    If Not B Then
         Exit Function
    End If
    
    conn.BeginTrans
    
    
    SQL = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute SQL



    SQL = "SELECT  rhisfruta.codsocio, rhisfruta.codvarie,"
    SQL = SQL & "rhisfruta.codcampo, rhisfruta.recolect, rhisfruta_clasif.codcalid, rhisfruta.fecalbar, "
'[Monica]01/09/2010 : sustituida la siguiente linea por
'    Sql = Sql & "rprecios_calidad.precoop, rprecios_calidad.presocio, rprecios_calidad.tipofact,sum(rhisfruta_clasif.kilosnet) as kilosnet "
     SQL = SQL & "rprecios.fechaini, rprecios.fechafin, rprecios_calidad.tipofact,max(rprecios.contador) contador, sum(rhisfruta_clasif.kilosnet) as kilosnet "
    SQL = SQL & " FROM  " & cTabla

    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        SQL = SQL & " WHERE " & cWhere
    End If
    ' ordenado por socio, variedad, campo, calidad
    SQL = SQL & " group by rhisfruta.codsocio, rhisfruta.codvarie, rhisfruta.codcampo, rhisfruta_clasif.codcalid, rhisfruta.fecalbar, rhisfruta.recolect  "
    SQL = SQL & " order by rhisfruta.codsocio, rhisfruta.codvarie, rhisfruta.codcampo, rhisfruta_clasif.codcalid, rhisfruta.fecalbar, rhisfruta.recolect "
    
    Set vSeccion = New CSeccion
    
    If vSeccion.LeerDatos(vParamAplic.Seccionhorto) Then
        If Not vSeccion.AbrirConta Then
            Exit Function
        End If
    End If
    
    HayReg = False
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    If Not Rs.EOF Then
        AntSocio = CStr(DBLet(Rs!Codsocio, "N"))
        AntVarie = CStr(DBLet(Rs!Codvarie, "N"))
        AntCampo = CStr(DBLet(Rs!codCampo, "N"))
        AntCalid = CStr(DBLet(Rs!codcalid, "N"))
        
        ActSocio = CStr(DBLet(Rs!Codsocio, "N"))
        ActVarie = CStr(DBLet(Rs!Codvarie, "N"))
        actCampo = CStr(DBLet(Rs!codCampo, "N"))
        ActCalid = CStr(DBLet(Rs!codcalid, "N"))
    
        Set vSocio = New cSocio
        If vSocio.LeerDatos(ActSocio) Then
            If vSocio.LeerDatosSeccion(ActSocio, vParamAplic.Seccionhorto) Then
                baseimpo = 0
                BaseReten = 0
                ImpoIva = 0
                ImpoReten = 0
                TotalFac = 0
                BaseAFO = 0
                ImpoAFO = 0
                PorcAFO = 0
                
                Kilos = 0
                Importe = 0
                Bonifica = 0
                
                KilosCal = 0
                
                vPorcIva = ""
                vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vSocio.CodIva, "N")
                PorcIva = CCur(ImporteSinFormato(vPorcIva))
                
                tipoMov = vSocio.CodTipomAnt
                
                Set vTipoMov = New CTiposMov
                
                numfactu = vTipoMov.ConseguirContador(tipoMov)
                Do
                    numfactu = vTipoMov.ConseguirContador(tipoMov)
                    devuelve = DevuelveDesdeBDNew(cAgro, "rfactsoc", "numfactu", "codtipom", tipoMov, "T", , "numfactu", CStr(numfactu), "N", "fecfactu", FecFac, "F")
                    If devuelve <> "" Then
                        'Ya existe el contador incrementarlo
                        Existe = True
                        vTipoMov.IncrementarContador (tipoMov)
                        numfactu = vTipoMov.ConseguirContador(tipoMov)
                    Else
                        Existe = False
                    End If
                Loop Until Not Existe
                
                vParamAplic.PrimFactAnt = numfactu
                
            End If
        End If
    End If
    
    While Not Rs.EOF And B
        ActCalid = DBLet(Rs!codcalid, "N")
        ActVarie = DBLet(Rs!Codvarie, "N")
        actCampo = DBSet(Rs!codCampo, "N")
        ActSocio = DBSet(Rs!Codsocio, "N")
        
        If (ActCalid <> AntCalid Or AntCampo <> actCampo Or AntVarie <> ActVarie Or AntSocio <> ActSocio) Then
            ' kilos e importe por variedad campo
            Kilos = Kilos + KilosCal
            Importe = Importe + vImporte
            Bonifica = Bonifica + vBonifica
            
            baseimpo = baseimpo + vImporte
            
            B = InsertLineaCalidad(tipoMov, CStr(numfactu), FecFac, AntVarie, AntCampo, CStr(AntCalid), CStr(DBLet(KilosCal, "N")), CStr(vImporte), CStr(vBonifica))
            KilosCal = 0
            vImporte = 0
            vBonifica = 0
            
            AntCalid = ActCalid
        End If
        
        If (ActVarie <> AntVarie Or actCampo <> AntCampo Or ActSocio <> AntSocio) Then
            ' insertar linea de variedad, campo
            B = InsertLinea(tipoMov, CStr(numfactu), FecFac, AntVarie, AntCampo, CStr(Kilos), CStr(Importe), "0", CStr(Bonifica))
            
            If B Then
                AntVarie = ActVarie
                AntCampo = actCampo
                Kilos = 0
                Importe = 0
                Bonifica = 0
            End If
        End If
        
        If ActSocio <> AntSocio Then
            ImpoIva = Round2(baseimpo * PorcIva / 100, 2)
        
            Select Case DBLet(vSocio.TipoIRPF, "N")
                Case 0
                    ImpoReten = Round2((baseimpo + ImpoIva) * vParamAplic.PorcreteFacSoc / 100, 2)
                    BaseReten = (baseimpo + ImpoIva)
                    PorcReten = vParamAplic.PorcreteFacSoc
                Case 1
                    ImpoReten = Round2(baseimpo * vParamAplic.PorcreteFacSoc / 100, 2)
                    BaseReten = baseimpo
                    PorcReten = vParamAplic.PorcreteFacSoc
                Case 2
                    ImpoReten = 0
                    BaseReten = 0
                    PorcReten = 0
            End Select
        
            TotalFac = baseimpo + ImpoIva - ImpoReten - ImpoAFO
            
            IncrementarProgresNew Pb1, 1
            
            'insertar cabecera de factura
            B = InsertCabecera(tipoMov, CStr(numfactu), FecFac)
            
            If B Then B = InsertResumen(tipoMov, CStr(numfactu))
            
            '[Monica]15/04/2013: Introducimos las facturas varias a descontar
            If DescontarFVarias Then
                If B Then B = InsertFacturasVarias(tipoMov, CStr(numfactu), FecFac, 0, 0)
            End If
            
            If B Then B = vTipoMov.IncrementarContador(tipoMov)
            
            If B Then
                AntSocio = ActSocio
                
                Set vSocio = Nothing
                Set vSocio = New cSocio
                If vSocio.LeerDatos(ActSocio) Then
                    If vSocio.LeerDatosSeccion(AntSocio, vParamAplic.Seccionhorto) Then
                        vPorcIva = ""
                        vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vSocio.CodIva, "N")
                        PorcIva = CCur(ImporteSinFormato(vPorcIva))
                    End If
                    
                    tipoMov = vSocio.CodTipomAnt
                End If
                baseimpo = 0
                BaseReten = 0
                Neto = 0
                ImpoIva = 0
                ImpoReten = 0
                TotalFac = 0
                BaseAFO = 0
                ImpoAFO = 0
                
                numfactu = vTipoMov.ConseguirContador(tipoMov)
                Do
                    numfactu = vTipoMov.ConseguirContador(tipoMov)
                    devuelve = DevuelveDesdeBDNew(cAgro, "rfactsoc", "numfactu", "codtipom", tipoMov, "T", , "numfactu", CStr(numfactu), "N", "fecfactu", FecFac, "F")
                    If devuelve <> "" Then
                        'Ya existe el contador incrementarlo
                        Existe = True
                        vTipoMov.IncrementarContador (tipoMov)
                        numfactu = vTipoMov.ConseguirContador(tipoMov)
                    Else
                        Existe = False
                    End If
                Loop Until Not Existe
           End If
        End If
        
        Recolect = DBLet(Rs!Recolect, "N")
        
        '[Monica]01/09/2010: añadido ésto, antes los precios los sacabamos en el propio select
        Dim Sql9 As String
        Dim Rs9 As ADODB.Recordset
        Dim PreCoop As Currency
        Dim PreSocio As Currency
        
        Sql9 = "select precoop, presocio from rprecios_calidad where codvarie = " & DBSet(Rs!Codvarie, "N")
        Sql9 = Sql9 & " and tipofact = " & DBSet(Rs!TipoFact, "N")
        Sql9 = Sql9 & " and contador = " & DBSet(Rs!Contador, "N")
        Sql9 = Sql9 & " and codcalid = " & DBSet(Rs!codcalid, "N")
        
        Set Rs9 = New ADODB.Recordset
        Rs9.Open Sql9, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
        
        If Not Rs9.EOF Then
            PreCoop = DBLet(Rs9.Fields(0).Value, "N")
            PreSocio = DBLet(Rs9.Fields(1).Value, "N")
            PorcBoni = 0
            PorcComi = 0
            Select Case Recolect
                Case 0
                    ' si el precio es positivo miramos si hay porcentaje de bonificacion para esa fecha
                    If PreCoop > 0 Then
                        PorcBoni = DevuelveValor("select porcbonif from rbonifentradas where codvarie = " & DBSet(Rs!Codvarie, "N") & " and fechaent = " & DBSet(Rs!Fecalbar, "F"))
                        
                        '[Monica]03/02/2012: Si el precio es positivo vemos si tiene comision el campo y se lo descontamos si es positivo
                        PorcComi = DevuelveValor("select dtoprecio from rcampos where codcampo = " & DBSet(Rs!codCampo, "N"))
                        If CCur(PorcComi) <> 0 Then
                            PreCoop = PreCoop - Round2(PreCoop * PorcComi / 100, 4)
                        End If
                    End If
                
                    vPrecio = DBLet(PreCoop, "N")
                    vImporte = vImporte + Round2(DBLet(Rs!KilosNet, "N") * PreCoop * (1 + (PorcBoni / 100)), 2)
                    vBonifica = vBonifica + Round2(DBLet(Rs!KilosNet, "N") * PreCoop * (1 + (PorcBoni / 100)), 2) - Round2(DBLet(Rs!KilosNet, "N") * PreCoop, 2)
                Case 1
                    ' si el precio es positivo miramos si hay porcentaje de bonificacion para esa fecha
                    If PreSocio > 0 Then
                        PorcBoni = DevuelveValor("select porcbonif from rbonifentradas where codvarie = " & DBSet(Rs!Codvarie, "N") & " and fechaent = " & DBSet(Rs!Fecalbar, "F"))
                        
                        '[Monica]03/02/2012: Si el precio es positivo vemos si tiene comision el campo y se lo descontamos si es positivo
                        PorcComi = DevuelveValor("select dtoprecio from rcampos where codcampo = " & DBSet(Rs!codCampo, "N"))
                        If CCur(PorcComi) <> 0 Then
                            PreSocio = PreSocio - Round2(PreSocio * PorcComi / 100, 4)
                        End If
                    End If
                    
                    vPrecio = DBLet(PreSocio, "N")
                    vImporte = vImporte + Round2(DBLet(Rs!KilosNet, "N") * PreSocio * (1 + (PorcBoni / 100)), 2)
            
                    vBonifica = vBonifica + Round2(DBLet(Rs!KilosNet, "N") * PreSocio * (1 + (PorcBoni / 100)), 2) - Round2(DBLet(Rs!KilosNet, "N") * PreSocio, 2)
            End Select
            
            KilosCal = KilosCal + DBLet(Rs!KilosNet, "N")
            
        End If
        
        Set Rs9 = Nothing
        
        'hasta aqui
        HayReg = True
        
        Rs.MoveNext
    Wend
    
    ' ultimo registro si ha entrado
    If B And HayReg Then
        ' insertar linea de calidad
        Kilos = Kilos + KilosCal
        Importe = Importe + vImporte
        Bonifica = Bonifica + vBonifica
        
        baseimpo = baseimpo + vImporte
        
        If B Then B = InsertLineaCalidad(tipoMov, CStr(numfactu), FecFac, ActVarie, actCampo, CStr(ActCalid), CStr(DBLet(KilosCal, "N")), CStr(vImporte), CStr(vBonifica))
        
        ' insertar linea de variedad
        If B Then B = InsertLinea(tipoMov, CStr(numfactu), FecFac, CStr(ActVarie), CStr(actCampo), CStr(Kilos), CStr(Importe), "0", CStr(Bonifica))
        
        ImpoIva = Round2(baseimpo * PorcIva / 100, 2)

        Select Case DBLet(vSocio.TipoIRPF, "N")
            Case 0
                ImpoReten = Round2((baseimpo + ImpoIva) * vParamAplic.PorcreteFacSoc / 100, 2)
                BaseReten = (baseimpo + ImpoIva)
                PorcReten = vParamAplic.PorcreteFacSoc
            Case 1
                ImpoReten = Round2(baseimpo * vParamAplic.PorcreteFacSoc / 100, 2)
                BaseReten = baseimpo
                PorcReten = vParamAplic.PorcreteFacSoc
            Case 2
                ImpoReten = 0
                BaseReten = 0
                PorcReten = 0
        End Select
        
'        BaseAFO = baseimpo
'        PorcAFO = vParamAplic.PorcenAFO
'        ImpoAFO = Round2(BaseAFO * PorcAFO / 100, 2)

        TotalFac = baseimpo + ImpoIva - ImpoReten - ImpoAFO
        
        IncrementarProgresNew Pb1, 1
        
        vParamAplic.UltFactAnt = numfactu
        
        'insertar cabecera de factura
        B = InsertCabecera(tipoMov, CStr(numfactu), FecFac)
        
        '[Monica]15/04/2013: Introducimos las facturas varias a descontar
        If DescontarFVarias Then
            If B Then B = InsertFacturasVarias(tipoMov, CStr(numfactu), FecFac, 0, 0)
        End If
        
        If B Then B = InsertResumen(tipoMov, CStr(numfactu))
        
        If B Then B = vTipoMov.IncrementarContador(tipoMov)
        
        'pasamos las temporales a las tablas
        If B Then B = PasarTemporales()
        
        If B Then B = (vParamAplic.Modificar = 1)
    End If
    
'    BorrarTMPs
    
    vSeccion.CerrarConta
    Set vSeccion = Nothing
    Set vSocio = Nothing
    
eFacturacion:
    If Err.Number <> 0 Or Not B Then
        conn.RollbackTrans
        FacturacionAnticiposPicassent = False
    Else
        conn.CommitTrans
        FacturacionAnticiposPicassent = True
    End If
End Function





Public Function FechaSuperiorUltimaLiquidacion(fec As Date) As Boolean
Dim SQL As String
Dim Rs As ADODB.Recordset
Dim Mensual As Boolean
Dim Anofactu As Integer
Dim PeriodoFactu As Integer
Dim FechaDesde As Date

    On Error GoTo eFechaSuperiorUltimaLiquidacion

    FechaSuperiorUltimaLiquidacion = False

    ' en caso de que haya contabilidad comprobamos que la fecha de factura introducida
    ' no sea inferior a la ultima liquidacion de iva.
    If vParamAplic.NumeroConta <> 0 Then
        SQL = "select periodos, anofactu, perfactu from parametros"
        Set Rs = New ADODB.Recordset
        Rs.Open SQL, ConnConta, adOpenDynamic, adLockOptimistic
        
        If Not Rs.EOF Then
            Mensual = (Rs.Fields(0).Value = 1)
            Anofactu = Rs.Fields(1).Value
            PeriodoFactu = Rs.Fields(2).Value
            
            If Mensual Then ' facturacion mensual
                If PeriodoFactu = 12 Then
                    FechaDesde = CDate("01/01/" & Format(Anofactu + 1, "0000"))
                Else
                    FechaDesde = CDate("01/" & Format(PeriodoFactu + 1, "00") & "/" & Format(Anofactu, "0000"))
                End If
            Else ' facturacion trimestral
                If PeriodoFactu = 4 Then
                    FechaDesde = CDate("01/01/" & Format(Anofactu + 1, "0000"))
                Else
                    FechaDesde = CDate("01/" & Format((PeriodoFactu * 3) + 1, "00") & "/" & Format(Anofactu, "0000"))
                End If
            End If
            
            FechaSuperiorUltimaLiquidacion = (fec >= FechaDesde)
        End If
    End If

eFechaSuperiorUltimaLiquidacion:
    If Err.Number <> 0 Then
         MuestraError Err.Number, Err.Description
    End If
End Function


Public Function FechaDentroPeriodoContable(fec As Date) As Boolean
Dim SQL As String
Dim Rs As ADODB.Recordset
Dim Mensual As Boolean
Dim Anofactu As Integer
Dim PeriodoFactu As Integer
Dim FechaDesde As Date

    On Error GoTo eFechaDentroPeriodoContable

    FechaDentroPeriodoContable = (CDate(FIni) <= fec) And (fec <= (CDate(FFin) + 365))

eFechaDentroPeriodoContable:
    If Err.Number <> 0 Then
         MuestraError Err.Number, Err.Description
    End If
End Function

Public Function FechaFacturaInferiorUltimaFacturaSerieHco(Fecha As Date, numfactu As Long, Serie As String, Tipo As Byte) As Boolean
' tipo = 0 indica schfac
' tipo = 1 indica schfac2 hco.de ajenas del Regaixo
Dim SQL As String
Dim Rs As ADODB.Recordset

    FechaFacturaInferiorUltimaFacturaSerieHco = False

    SQL = "select fecfactu "
    If Tipo = 0 Then
        SQL = SQL & "from schfac "
    Else
        SQL = SQL & "from schfacr "
    End If
    SQL = SQL & " where numfactu = " & DBSet(numfactu, "N") & " and letraser = " & DBSet(Serie, "T")
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenDynamic, adLockOptimistic
    If Not Rs.EOF Then
        If Rs.Fields(0).Value > Fecha Then
            FechaFacturaInferiorUltimaFacturaSerieHco = True
        End If
    End If

End Function


Public Function DeshacerFacturacion(Tipo As Byte, DesFac As String, HasFac As String, FecFac As String, Pb1 As ProgressBar) As Boolean
' Tipo : 0 --> factura de anticipo
'        1 --> factura de anticipo venta campo
'        2 --> factura de liquidacion venta campo
'        3 --> factura de liquidacion
'        4 --> factura de anticipo almazara
'        5 --> factura de liquidacion almazara
'        6 --> factura de anticipo bodega
'        7 --> factura de liquidacion bodega
Dim SQL As String
Dim Sql2 As String
Dim Rs As ADODB.Recordset
Dim Rs2 As ADODB.Recordset
Dim vTipoMov As CTiposMov
Dim Nregs As Long
Dim tipoMov As String
Dim HayReg As Boolean

Dim B As Boolean
Dim vWhere As String
Dim Sql3 As String


    On Error GoTo eDeshacerFactAnt

    DeshacerFacturacion = False
    
    SQL = "select rfactsoc.* from rfactsoc, usuarios.stipom stipom  where fecfactu = " & DBSet(FecFac, "F")
    SQL = SQL & " and numfactu >= " & DBSet(DesFac, "N")
    SQL = SQL & " and numfactu <= " & DBSet(HasFac, "N")
    
    Select Case Tipo
        Case 0 ' factura de anticipos
            SQL = SQL & " and stipom.tipodocu = 1"
        Case 1 ' factura de anticipo de ventas campo
            SQL = SQL & " and stipom.tipodocu = 3"
        Case 2 ' factura de liquidacion de ventas campo
            SQL = SQL & " and stipom.tipodocu = 4"
        Case 3 ' factura de liquidacion
            SQL = SQL & " and stipom.tipodocu = 2"
        Case 4 ' factura de anticipo de almazara
            SQL = SQL & " and stipom.tipodocu = 7"
        Case 5 ' factura de liquidacion de almazara
            SQL = SQL & " and stipom.tipodocu = 8"
        Case 6 ' factura de anticipo de bodega
            SQL = SQL & " and stipom.tipodocu = 9"
        Case 7 ' factura de liquidacion de bodega
            SQL = SQL & " and stipom.tipodocu = 10"
    End Select
    SQL = SQL & " and rfactsoc.codtipom = stipom.codtipom "
    SQL = SQL & " order by numfactu desc "
    
    Nregs = TotalRegistrosConsulta(SQL)
    If Nregs = 0 Then
        Pb1.visible = False
        MsgBox "No se corresponde con la última facturación", vbExclamation
        Exit Function
    End If
    
    CargarProgres Pb1, CInt(Nregs)
    
    conn.BeginTrans
    
    B = True
    HayReg = False
    
    Set vTipoMov = New CTiposMov
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText

    While Not Rs.EOF And B
        IncrementarProgres Pb1, 1
        
        B = (vTipoMov.DevolverContador(DBLet(Rs!CodTipom, "T"), DBLet(Rs!numfactu, "N")) = 1)
    
        If B Then
            HayReg = True
            
            vWhere = "codtipom = " & DBSet(Rs!CodTipom, "T") & " and numfactu = " & DBLet(Rs!numfactu, "N") & " and fecfactu = " & DBSet(FecFac, "F")
            
            SQL = "delete from rfactsoc_calidad where " & vWhere
            conn.Execute SQL
            
            ' si deshacemos la factura de liquidacion de venta campo (tipo = 2) o de liquidacion (tipo=3)
            ' o de liquidacion almazara (tipo = 5) o de liquidacion bodega (tipo = 7)
            ' hemos de desmarcar los anticipos
            ' descontados y borrarlos de la tabla de anticipos descontados en factura
            If Tipo = 2 Or Tipo = 3 Or Tipo = 5 Or Tipo = 7 Then
                Sql2 = "select codtipomanti, numfactuanti, fecfactuanti, codvarieanti, codcampoanti from rfactsoc_anticipos "
                Sql2 = Sql2 & " where " & vWhere
                
                Set Rs2 = New ADODB.Recordset
                Rs2.Open Sql2, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
            
                While Not Rs2.EOF And B
                    ' desmarcar los anticipos como que no estan descontados
                    Sql3 = "update rfactsoc_variedad set descontado = 0 where "
                    Sql3 = Sql3 & " codtipom = " & DBSet(Rs2.Fields(0).Value, "T")
                    Sql3 = Sql3 & " and numfactu = " & DBSet(Rs2.Fields(1).Value, "N")
                    Sql3 = Sql3 & " and fecfactu = " & DBSet(Rs2.Fields(2).Value, "F")
                    Sql3 = Sql3 & " and codvarie = " & DBSet(Rs2.Fields(3).Value, "N")
                    Sql3 = Sql3 & " and codcampo = " & DBSet(Rs2.Fields(4).Value, "N")
                    
                    conn.Execute Sql3
                    
                    Rs2.MoveNext
                Wend
                ' borrar de la tabla de anticipos descontados
                Sql3 = "delete from rfactsoc_anticipos where " & vWhere
                
                conn.Execute Sql3
            
                
                '   ANTICIPOS DE RETIRADA
                '[Monica]05/12/2011: desmarcamos los anticipos de retirada si los hay (solo Quatretonda)
                Sql2 = "select codtipomanti, numfactuanti, fecfactuanti, codvarieanti, codcampoanti from rfactsoc_retirada "
                Sql2 = Sql2 & " where " & vWhere
                
                Set Rs2 = New ADODB.Recordset
                Rs2.Open Sql2, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
            
                While Not Rs2.EOF And B
                    ' desmarcar los anticipos como que no estan descontados
                    Sql3 = "update rfactsoc_variedad set descontado = 0 where "
                    Sql3 = Sql3 & " codtipom = " & DBSet(Rs2.Fields(0).Value, "T")
                    Sql3 = Sql3 & " and numfactu = " & DBSet(Rs2.Fields(1).Value, "N")
                    Sql3 = Sql3 & " and fecfactu = " & DBSet(Rs2.Fields(2).Value, "F")
                    Sql3 = Sql3 & " and codvarie = " & DBSet(Rs2.Fields(3).Value, "N")
                    Sql3 = Sql3 & " and codcampo = " & DBSet(Rs2.Fields(4).Value, "N")
                    
                    conn.Execute Sql3
                    
                    Rs2.MoveNext
                Wend
                ' borrar de la tabla de anticipos de retirada descontados
                Sql3 = "delete from rfactsoc_retirada where " & vWhere
                
                conn.Execute Sql3
            
                Set Rs2 = Nothing
            
            End If
            
            ' FACTURAS VARIAS para anticipos y para liquidaciones
            '[Monica]16/04/2013: Desmarcar las facturas varias si las hay
            If Tipo = 0 Or Tipo = 2 Or Tipo = 3 Or Tipo = 5 Or Tipo = 7 Then
                Sql2 = "select codtipomfvar, numfactufvar, fecfactufvar, codsecci from rfactsoc_fvarias "
                Sql2 = Sql2 & " where " & vWhere
                
                Set Rs2 = New ADODB.Recordset
                Rs2.Open Sql2, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
            
                While Not Rs2.EOF And B
                    Sql3 = "update fvarcabfact set intliqui = 0 where "
                    Sql3 = Sql3 & " codtipom = " & DBSet(Rs2.Fields(0).Value, "T")
                    Sql3 = Sql3 & " and numfactu = " & DBSet(Rs2.Fields(1).Value, "N")
                    Sql3 = Sql3 & " and fecfactu = " & DBSet(Rs2.Fields(2).Value, "F")
                    Sql3 = Sql3 & " and codsecci = " & DBSet(Rs2.Fields(3).Value, "N")
                    
                    conn.Execute Sql3
                    
                    Rs2.MoveNext
                Wend
                ' borrar de la tabla de facturas varias
                Sql3 = "delete from rfactsoc_fvarias where " & vWhere
                conn.Execute Sql3
                
                Set Rs2 = Nothing
            End If
            
            '[Monica]10/04/2017: Incrementos de factura Picassent y Coopic
            If Tipo = 3 Then
                Sql3 = "insert into ringresos (codsocio, codvarie, concepto, importe) "
                Sql3 = Sql3 & " select aa.codsocio, rfactsoc_ingresos.codvarie, rfactsoc_ingresos.concepto, rfactsoc_ingresos.importe  from rfactsoc_ingresos , (select codsocio from rfactsoc where " & vWhere & ") aa "
                Sql3 = Sql3 & " where " & vWhere
                
                conn.Execute Sql3
            End If
            SQL = "delete from rfactsoc_ingresos where " & vWhere
            conn.Execute SQL
            
           
            SQL = "delete from rfactsoc_variedad where " & vWhere
            conn.Execute SQL
            
            SQL = "delete from rfactsoc_gastos where " & vWhere
            conn.Execute SQL
            
            SQL = "delete from rfactsoc_albaran where " & vWhere
            conn.Execute SQL
            
            SQL = "delete from rfactsoc where " & vWhere
            conn.Execute SQL
            
            
            Select Case Tipo
                Case 0 ' factura de anticipo
                    vParamAplic.UltFactAnt = DBLet(Rs!numfactu, "N")
                    vParamAplic.PrimFactAnt = vParamAplic.UltFactAnt
                Case 1 ' factura de anticipo de venta campo
                    vParamAplic.UltFactAntVC = DBLet(Rs!numfactu, "N")
                    vParamAplic.PrimFactAntVC = vParamAplic.UltFactAntVC
                Case 2 ' factura de liquidacion de venta campo
                    vParamAplic.UltFactLiqVC = DBLet(Rs!numfactu, "N")
                    vParamAplic.PrimFactLiqVC = vParamAplic.UltFactLiqVC
                Case 3 ' factura de liquidacion
                    vParamAplic.UltFactLiq = DBLet(Rs!numfactu, "N")
                    vParamAplic.PrimFactLiq = vParamAplic.UltFactLiq
                Case 4 ' factura de anticipo de almazara
                    vParamAplic.UltFactAntAlmz = DBLet(Rs!numfactu, "N")
                    vParamAplic.PrimFactAntAlmz = vParamAplic.UltFactAntAlmz
                Case 5 ' factura de liquidacion de almazara
                    vParamAplic.UltFactLiqAlmz = DBLet(Rs!numfactu, "N")
                    vParamAplic.PrimFactLiqAlmz = vParamAplic.UltFactLiqAlmz
                Case 6 ' factura de anticipo de bodega
                    vParamAplic.UltFactAntBOD = DBLet(Rs!numfactu, "N")
                    vParamAplic.PrimFactAntBOD = vParamAplic.UltFactAntBOD
                Case 7 ' factura de liquidacion de bodega
                    vParamAplic.UltFactLiqBOD = DBLet(Rs!numfactu, "N")
                    vParamAplic.PrimFactLiqBOD = vParamAplic.UltFactLiqBOD
            End Select
            
        End If
        
        Rs.MoveNext
    Wend
    
    Set Rs = Nothing
    
    If HayReg Then
        Select Case Tipo
            Case 0 ' factura de anticipo
                vParamAplic.UltFactAnt = vParamAplic.UltFactAnt - 1
                vParamAplic.PrimFactAnt = vParamAplic.UltFactAnt
            Case 1 ' factura de anticipo de venta campo
                vParamAplic.UltFactAntVC = vParamAplic.UltFactAntVC - 1
                vParamAplic.PrimFactAntVC = vParamAplic.UltFactAntVC
            Case 2 ' factura de liquidacion de venta campo
                vParamAplic.UltFactLiqVC = vParamAplic.UltFactLiqVC - 1
                vParamAplic.PrimFactLiqVC = vParamAplic.UltFactLiqVC
            Case 3 ' factura de liquidacion
                vParamAplic.UltFactLiq = vParamAplic.UltFactLiq - 1
                vParamAplic.PrimFactLiq = vParamAplic.UltFactLiq
            Case 4 ' factura de anticipo de almazara
                vParamAplic.UltFactAntAlmz = vParamAplic.UltFactAntAlmz - 1
                vParamAplic.PrimFactAntAlmz = vParamAplic.UltFactAntAlmz
            Case 5 ' factura de liquidacion de almazara
                vParamAplic.UltFactLiqAlmz = vParamAplic.UltFactLiqAlmz - 1
                vParamAplic.PrimFactLiqAlmz = vParamAplic.UltFactLiqAlmz
            Case 6 ' factura de anticipo de bodega
                vParamAplic.UltFactAntBOD = vParamAplic.UltFactAntBOD - 1
                vParamAplic.PrimFactAntBOD = vParamAplic.UltFactAntBOD
            Case 7 ' factura de liquidacion de bodega
                vParamAplic.UltFactLiqBOD = vParamAplic.UltFactLiqBOD - 1
                vParamAplic.PrimFactLiqBOD = vParamAplic.UltFactLiqBOD
        End Select

        B = (vParamAplic.Modificar = 1)
    End If
    
    If B Then
        conn.CommitTrans
        DeshacerFacturacion = True
        Set vTipoMov = Nothing
        Exit Function
    End If
    
eDeshacerFactAnt:
    If Err.Number <> 0 Or Not B Then
        conn.RollbackTrans
        Set vTipoMov = Nothing
        MuestraError Err.Number, "Error deshaciendo Facturacion", Err.Description
    End If
End Function

Public Function FacturacionVentaCampo(Tipo As Byte, cTabla As String, cWhere As String, FecFac As String, Pb1 As ProgressBar, ConAFO As Byte, DescontarFVarias As Boolean, EsTerceros As Boolean) As Long
'Tipo: 0 -- factura de venta campo ANTICIPO
'Tipo: 1 -- factura de venta campo LIQUIDACION
Dim SQL As String
Dim Sql2 As String
Dim Rs As ADODB.Recordset
Dim RS1 As ADODB.Recordset

Dim AntSocio As String
Dim ActSocio As String

Dim HayReg As Boolean

Dim NumError As Long
Dim vImporte As Currency
Dim vPorcIva As String

Dim PrimFac As String
Dim UltFac As String
Dim tipoMov As String

Dim Sql3 As String
Dim devuelve As String
Dim B As Boolean
Dim vSeccion As CSeccion
Dim Existe As Boolean

    On Error GoTo eFacturacion
    
'08052009 antes dentro de transaccion
    BorrarTMPs
    B = CrearTMPs()
    If Not B Then
         Exit Function
    End If
'08052009

    conn.BeginTrans

    If EsTerceros Then
        Select Case Tipo
            Case 0 'Anticipo
                tipoMov = "CAT"
            Case 1 'Liquidacion
                tipoMov = "CLT"
        End Select
    Else
        Select Case Tipo
            Case 0 'Anticipo
                tipoMov = "FAC"
            Case 1 'Liquidacion
                tipoMov = "FLC"
        End Select
    End If


    SQL = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute SQL

    SQL = "SELECT  rhisfruta.codsocio, rhisfruta.codvarie,"
    SQL = SQL & " rhisfruta.codcampo, sum(rhisfruta.impentrada) as importe, "
    SQL = SQL & " sum(rhisfruta.kilosnet) as kilosnet "
    SQL = SQL & " FROM  " & cTabla

    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        SQL = SQL & " WHERE " & cWhere
    End If
    
    ' ordenado por socio, variedad, campo, calidad
    SQL = SQL & " group by rhisfruta.codsocio, rhisfruta.codvarie, rhisfruta.codcampo "
    SQL = SQL & " order by rhisfruta.codsocio, rhisfruta.codvarie, rhisfruta.codcampo "
    
    Set vSeccion = New CSeccion
    
    If vSeccion.LeerDatos(vParamAplic.Seccionhorto) Then
        If Not vSeccion.AbrirConta Then
            Exit Function
        End If
    End If
    
    HayReg = False
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    B = False
    
    If Not Rs.EOF Then
        AntSocio = CStr(DBLet(Rs!Codsocio, "N"))
        ActSocio = CStr(DBLet(Rs!Codsocio, "N"))
    
        Set vSocio = New cSocio
        If vSocio.LeerDatos(ActSocio) Then
            If vSocio.LeerDatosSeccion(ActSocio, vParamAplic.Seccionhorto) Then
                baseimpo = 0
                BaseReten = 0
                ImpoIva = 0
                ImpoReten = 0
                TotalFac = 0
                BaseAFO = 0
                ImpoAFO = 0
                PorcAFO = 0
                
                Anticipos = 0
                
                vPorcIva = ""
                vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vSocio.CodIva, "N")
                PorcIva = CCur(ImporteSinFormato(vPorcIva))
                
                '[Monica]29/05/2017: solo si es picassent seleccionamos socios terceros
                If EsTerceros Then
                    Select Case Tipo
                        Case 0 ' anticipos
                            tipoMov = "CAT"
                        Case 1 ' liquidacion
                            tipoMov = "CLT"
                    End Select
                Else
                    Select Case Tipo
                        Case 0 ' anticipos
                            tipoMov = vSocio.CodTipomAntVC
                        Case 1 ' liquidacion
                            tipoMov = vSocio.CodTipomLiqVC
                    End Select
                End If
                Set vTipoMov = New CTiposMov
                numfactu = vTipoMov.ConseguirContador(tipoMov)
                Do
                    numfactu = vTipoMov.ConseguirContador(tipoMov)
                    devuelve = DevuelveDesdeBDNew(cAgro, "rfactsoc", "numfactu", "codtipom", tipoMov, "T", , "numfactu", CStr(numfactu), "N", "fecfactu", FecFac, "F")
                    If devuelve <> "" Then
                        'Ya existe el contador incrementarlo
                        Existe = True
                        vTipoMov.IncrementarContador (tipoMov)
                        numfactu = vTipoMov.ConseguirContador(tipoMov)
                    Else
                        Existe = False
                    End If
                Loop Until Not Existe
                
                Select Case Tipo
                    Case 0  ' anticipo venta campo
                        vParamAplic.PrimFactAntVC = numfactu
                    Case 1  ' liquidacion venta campo
                        vParamAplic.PrimFactLiqVC = numfactu
                End Select
                    
                B = True
            End If
        End If
    End If
    
    While Not Rs.EOF And B
        ActSocio = DBSet(Rs!Codsocio, "N")
        If ActSocio <> AntSocio Then
        
            ImpoIva = Round2(baseimpo * PorcIva / 100, 2)
        
            Select Case DBLet(vSocio.TipoIRPF, "N")
                Case 0
                    ImpoReten = Round2((baseimpo + ImpoIva) * vParamAplic.PorcreteFacSoc / 100, 2)
                    BaseReten = (baseimpo + ImpoIva)
                    PorcReten = vParamAplic.PorcreteFacSoc
                Case 1
                    ImpoReten = Round2(baseimpo * vParamAplic.PorcreteFacSoc / 100, 2)
                    BaseReten = baseimpo
                    PorcReten = vParamAplic.PorcreteFacSoc
                Case 2
                    ImpoReten = 0
                    BaseReten = 0
                    PorcReten = 0
            End Select
        
            ' solo si es liquidacion tiene aportacion
            If Tipo = 1 Then
                ' si es Picassent el importe de Fondo Operativo no va por porcentaje sino por importe global
                ' proceso de Calculo de FO dentro de Pago Socios/Liquidacion
                If vParamAplic.Cooperativa = 2 Or vParamAplic.Cooperativa = 16 Then
                    If ConAFO = 1 Then
                        ImpoAFO = DevuelveValor("select sum(importe) from raporreparto where codsocio = " & DBSet(vSocio.Codigo, "N") & " and tipoentr = 1")
                    Else
                        ImpoAFO = 0
                    End If
                    BaseAFO = 0
                    PorcAFO = 0
                Else
                'cualquier otra cooperativa tiene un porcentaje de fondo operativo
                    ImpoAFO = Round2((baseimpo + Anticipos) * vParamAplic.PorcenAFO / 100, 2)
                    BaseAFO = baseimpo + Anticipos
                    PorcAFO = vParamAplic.PorcenAFO
                End If
            End If
        
            TotalFac = baseimpo + ImpoIva - ImpoReten - ImpoAFO
            
            IncrementarProgresNew Pb1, 1
            
            'insertar cabecera de factura
            B = InsertCabecera(tipoMov, CStr(numfactu), FecFac)
            
            '[Monica]04/01/2012: marcamos la factura como contabilizada y como pdte de recibir el nro de factura
            If B And vSocio.EmiteFact And vParamAplic.Cooperativa = 4 Then B = MarcarFactura(tipoMov, CStr(numfactu), FecFac, True)
            
            
            '[Monica]29/05/2017: para el caso de picassent si es tercero se marca como contabilizada pdte de recibir factura
            If EsTerceros Then
                If B Then B = MarcarFactura(tipoMov, CStr(numfactu), FecFac, True)
            End If
            
            
            If B Then B = InsertResumen(tipoMov, CStr(numfactu))
            
            '[Monica]15/04/2013: Introducimos las facturas varias a descontar
            If DescontarFVarias Then
                If B Then B = InsertFacturasVarias(tipoMov, CStr(numfactu), FecFac, Tipo, 1)
            End If
            
            If B Then B = vTipoMov.IncrementarContador(tipoMov)
            
            If B Then
                AntSocio = ActSocio
                
                Set vSocio = Nothing
                Set vSocio = New cSocio
                If vSocio.LeerDatos(ActSocio) Then
                    If vSocio.LeerDatosSeccion(AntSocio, vParamAplic.Seccionhorto) Then
                        vPorcIva = ""
                        vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vSocio.CodIva, "N")
                        PorcIva = CCur(ImporteSinFormato(vPorcIva))
                    End If
                End If
                baseimpo = 0
                BaseReten = 0
                ImpoIva = 0
                ImpoReten = 0
                TotalFac = 0
                BaseAFO = 0
                ImpoAFO = 0
                PorcAFO = 0
                Anticipos = 0
                
                If EsTerceros Then
                    Select Case Tipo
                        Case 0 ' anticipos
                            tipoMov = "CAT"
                        Case 1 ' liquidacion
                            tipoMov = "CLT"
                    End Select
                Else
                    Select Case Tipo
                        Case 0 ' anticipos
                            tipoMov = vSocio.CodTipomAntVC
                        Case 1 ' liquidacion
                            tipoMov = vSocio.CodTipomLiqVC
                    End Select
                End If
                
                If B Then
                    numfactu = vTipoMov.ConseguirContador(tipoMov)
                    Do
                        numfactu = vTipoMov.ConseguirContador(tipoMov)
                        devuelve = DevuelveDesdeBDNew(cAgro, "rfactsoc", "numfactu", "codtipom", tipoMov, "T", , "numfactu", CStr(numfactu), "N", "fecfactu", FecFac, "F")
                        If devuelve <> "" Then
                            'Ya existe el contador incrementarlo
                            Existe = True
                            vTipoMov.IncrementarContador (tipoMov)
                            numfactu = vTipoMov.ConseguirContador(tipoMov)
                        Else
                            Existe = False
                        End If
                    Loop Until Not Existe
                End If
           End If
        End If
        
        baseimpo = baseimpo + DBLet(Rs!Importe, "N")
        
        ' insertar linea de variedad, campo
        B = InsertLinea(tipoMov, CStr(numfactu), FecFac, CStr(DBLet(Rs!Codvarie, "N")), CStr(DBLet(Rs!codCampo, "N")), CStr(DBLet(Rs!KilosNet, "N")), CStr(DBLet(Rs!Importe, "N")), 0)
        
        '[Monica]08/04/2010: grabamos los albaranes que intervienen en la linea de factura
        If B Then
            B = InsertarAlbaranesFactura(tipoMov, CStr(numfactu), FecFac, CStr(DBLet(Rs!Codsocio, "N")), CStr(DBLet(Rs!Codvarie, "N")), CStr(DBLet(Rs!codCampo, "N")), cTabla, cWhere, 2)
        End If
        
        If B Then
            ' insertamos los totales en la calidad venta campo de la variedad (rfactsoc_calidad)
            Sql2 = "select codcalid from rcalidad where codvarie = " & DBSet(Rs!Codvarie, "N")
            Sql2 = Sql2 & " and tipcalid = 2 " ' calidad de venta campo
            
            Set RS1 = New ADODB.Recordset
            RS1.Open Sql2, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
            
            If Not RS1.EOF Then
                B = InsertLineaCalidad(tipoMov, CStr(numfactu), FecFac, CStr(DBLet(Rs!Codvarie, "N")), CStr(DBLet(Rs!codCampo, "N")), CStr(DBLet(RS1!codcalid, "N")), CStr(DBLet(Rs!KilosNet, "N")), CStr(DBLet(Rs!Importe, "N")))
            End If
            Set RS1 = Nothing
        End If
        
        
        
        
        ' si es una factura de liquidacion hemos de descontar los anticipos de las variedades
        If B And Tipo = 1 Then
            Sql2 = "select rfactsoc_variedad.imporvar, rfactsoc_variedad.numfactu, rfactsoc_variedad.fecfactu "
            Sql2 = Sql2 & " from rfactsoc_variedad INNER JOIN rfactsoc ON rfactsoc_variedad.codtipom = rfactsoc.codtipom and "
            Sql2 = Sql2 & " rfactsoc_variedad.numfactu = rfactsoc.numfactu and rfactsoc_variedad.fecfactu = rfactsoc.fecfactu "
            
            
            '[Monica]29/05/2017: si estercero es solo para picassent
            
            If EsTerceros Then
                Sql2 = Sql2 & " where rfactsoc_variedad.codtipom = 'CAT'"
            Else
                Sql2 = Sql2 & " where rfactsoc_variedad.codtipom = " & DBSet(vSocio.CodTipomAntVC, "T") ' antes era 'FAC' "
            End If
            
            Sql2 = Sql2 & " and rfactsoc.codsocio = " & DBSet(Rs!Codsocio, "N")
            Sql2 = Sql2 & " and codvarie = " & DBSet(Rs!Codvarie, "N")
            Sql2 = Sql2 & " and codcampo = " & DBSet(Rs!codCampo, "N")
            Sql2 = Sql2 & " and rfactsoc_variedad.descontado = 0"
            
            Set RS1 = New ADODB.Recordset
            RS1.Open Sql2, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
            
            While Not RS1.EOF
                baseimpo = baseimpo - DBLet(RS1.Fields(0).Value, "N")
                Anticipos = Anticipos + DBLet(RS1.Fields(0).Value, "N")
                
                ' indicamos que los anticipos ya han sido descontados
                Sql3 = "update rfactsoc_variedad set descontado = 1 where codtipom = "
                
                '[Monica]29/05/2017: en el caso solo de picassent
                If EsTerceros Then
                    Sql3 = Sql3 & "'CAT' "  ' antes era 'FAC'
                Else
                    Sql3 = Sql3 & DBSet(vSocio.CodTipomAntVC, "T")  ' antes era 'FAC'
                End If
                
                Sql3 = Sql3 & " and numfactu = " & DBSet(RS1!numfactu, "N")
                Sql3 = Sql3 & " and fecfactu = " & DBSet(RS1!fecfactu, "F") & " and codvarie = " & DBSet(Rs!Codvarie, "N")
                Sql3 = Sql3 & " and codcampo = " & DBSet(Rs!codCampo, "N")
                
                conn.Execute Sql3
                
                ' insertamos en la tabla de anticipos de liquidacion venta campo
                Sql3 = "insert into tmpFact_anticipos (codtipom, numfactu, fecfactu, codtipomanti, numfactuanti, fecfactuanti, codvarieanti, codcampoanti, baseimpo) values ("
                If EsTerceros Then
                    Sql3 = Sql3 & "'CLT',"  ' liquidacion de terceros
                Else
                    Sql3 = Sql3 & DBSet(vSocio.CodTipomLiqVC, "T") & "," ' antes era 'FLC'
                End If
                
                Sql3 = Sql3 & DBSet(numfactu, "N") & "," & DBSet(FecFac, "F") & ","
                If EsTerceros Then
                    Sql3 = Sql3 & DBSet("CAT", "T") & "," ' antes era 'FAC'
                Else
                    Sql3 = Sql3 & DBSet(vSocio.CodTipomAntVC, "T") & "," ' antes era 'FAC'
                End If
                Sql3 = Sql3 & DBSet(RS1!numfactu, "N") & "," & DBSet(RS1!fecfactu, "F") & ","
                Sql3 = Sql3 & DBSet(Rs!Codvarie, "N") & "," & DBSet(Rs!codCampo, "N") & "," & DBSet(RS1!imporvar, "N") & ")"
                
                conn.Execute Sql3
                
                RS1.MoveNext
            Wend
            
            Set RS1 = Nothing
            
        End If
        
        HayReg = True
        
        Rs.MoveNext
    Wend
    
    ' ultimo registro si ha entrado
    If B And HayReg Then
        ImpoIva = Round2(baseimpo * PorcIva / 100, 2)

        Select Case DBLet(vSocio.TipoIRPF, "N")
            Case 0
                ImpoReten = Round2((baseimpo + ImpoIva) * vParamAplic.PorcreteFacSoc / 100, 2)
                BaseReten = (baseimpo + ImpoIva)
                PorcReten = vParamAplic.PorcreteFacSoc
            Case 1
                ImpoReten = Round2(baseimpo * vParamAplic.PorcreteFacSoc / 100, 2)
                BaseReten = baseimpo
                PorcReten = vParamAplic.PorcreteFacSoc
            Case 2
                ImpoReten = 0
                BaseReten = 0
                PorcReten = 0
        End Select
        
        ' solo si es liquidacion tiene aportacion
        If Tipo = 1 Then
            ' si es Picassent el importe de Fondo Operativo no va por porcentaje sino por importe global
            ' proceso de Calculo de FO dentro de Pago Socios/Liquidacion
            If vParamAplic.Cooperativa = 2 Or vParamAplic.Cooperativa = 16 Then
                If ConAFO = 1 Then
                    ImpoAFO = DevuelveValor("select sum(importe) from raporreparto where codsocio = " & DBSet(vSocio.Codigo, "N") & " and tipoentr = 1")
                Else
                    ImpoAFO = 0
                End If
                BaseAFO = 0
                PorcAFO = 0
            Else
            'cualquier otra cooperativa tiene un porcentaje de fondo operativo
                ImpoAFO = Round2((baseimpo + Anticipos) * vParamAplic.PorcenAFO / 100, 2)
                BaseAFO = baseimpo + Anticipos
                PorcAFO = vParamAplic.PorcenAFO
            End If
        End If
        
        TotalFac = baseimpo + ImpoIva - ImpoReten - ImpoAFO
        
        IncrementarProgresNew Pb1, 1
        
        'insertar cabecera de factura
        B = InsertCabecera(tipoMov, CStr(numfactu), FecFac)
        
        '[Monica]04/01/2012: marcamos la factura como contabilizada y como pdte de recibir el nro de factura
        If B And vSocio.EmiteFact And vParamAplic.Cooperativa = 4 Then B = MarcarFactura(tipoMov, CStr(numfactu), FecFac, True)
        
        '[Monica]29/05/2017: marcamos la factura esperando para ser recibida
        If EsTerceros Then
            If B Then B = MarcarFactura(tipoMov, CStr(numfactu), FecFac, True)
        End If
        
        If B Then B = InsertResumen(tipoMov, CStr(numfactu))
        
        '[Monica]15/04/2013: Introducimos las facturas varias a descontar
        If DescontarFVarias Then
            If B Then B = InsertFacturasVarias(tipoMov, CStr(numfactu), FecFac, Tipo, 1)
        End If
        
        If B Then B = vTipoMov.IncrementarContador(tipoMov)
        
        Select Case Tipo
            Case 0
                vParamAplic.UltFactAntVC = numfactu
            Case 1
                vParamAplic.UltFactLiqVC = numfactu
        End Select
        
        'pasamos las temporales a las tablas
        If B Then B = PasarTemporales()
        
        If B Then B = (vParamAplic.Modificar = 1)
        
        ' si es anticipo o liquidacion de venta campo se vuelven los importes a null
        If B Then
            SQL = "update " & cTabla & " set rhisfruta.impentrada = 0 where " & cWhere
            B = ActualizaRegistros(SQL)
        End If
    End If
    
    BorrarTMPs
    
    vSeccion.CerrarConta
    Set vSeccion = Nothing
    Set vSocio = Nothing
    
eFacturacion:
    If Err.Number <> 0 Or Not B Then
        conn.RollbackTrans
        FacturacionVentaCampo = False
    Else
        conn.CommitTrans
        FacturacionVentaCampo = True
    End If
End Function



'Actualizar registros
Private Function ActualizaRegistros(cad As String) As Boolean
Dim Precio As Currency

    Dim SQL As String
    Dim ImpLinea As Currency
    
    On Error GoTo eActualiza
    
    ActualizaRegistros = False
    
    MensError = ""
    
    conn.Execute cad

    ActualizaRegistros = True
    Exit Function
    
eActualiza:
    If Err.Number <> 0 Then
        MensError = "Se ha producido un error en la actualización de registros."
        MuestraError Err.Number, MensError
    End If
End Function

Public Function FacturasGeneradas(Tipo As String) As String
Dim SQL As String
Dim RS1 As ADODB.Recordset
Dim cad As String
    
    On Error GoTo eFacturasGeneradas

    FacturasGeneradas = ""

    SQL = "select nombre1, importe1 from tmpinformes, usuarios.stipom stipom where codusu = " & vUsu.Codigo
    SQL = SQL & " and stipom.codtipom = nombre1 "
    Select Case Tipo
        Case 0 ' anticipos venta campo
            SQL = SQL & " and stipom.tipodocu = 3 "
        Case 1 ' liquidacion venta campo
            SQL = SQL & " and stipom.tipodocu = 4 "
        Case 2 ' anticipos
            SQL = SQL & " and stipom.tipodocu = 1 "
        Case 3 ' liquidacion
            SQL = SQL & " and stipom.tipodocu = 2 "
        Case 4 ' anticipos almazara
            SQL = SQL & " and stipom.tipodocu = 7 "
        Case 5 ' liquidacion almazara
            SQL = SQL & " and stipom.tipodocu = 8 "
        Case 6 ' anticipos bodega
            SQL = SQL & " and stipom.tipodocu = 9 "
        Case 7 ' liquidacion bodega
            SQL = SQL & " and stipom.tipodocu = 10 "
    End Select
    
'    SQL = SQL & " and nombre1 = " & DBSet(Tipo, "T")
    
    Set RS1 = New ADODB.Recordset
    RS1.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    cad = ""
    While Not RS1.EOF
        cad = cad & DBLet(RS1.Fields(1).Value, "N") & ","
    
        RS1.MoveNext
    Wend
    Set RS1 = Nothing
    
    'si hay facturas quitamos la ultima coma
    If cad <> "" Then cad = Mid(cad, 1, Len(cad) - 1)
    
    FacturasGeneradas = cad
    Exit Function
    
eFacturasGeneradas:
    MuestraError Err.Number, "Cadena de Facturas Generadas", Err.Description
End Function


Public Function ListaFacturasGeneradas(Tipo As String) As String
Dim SQL As String
Dim RS1 As ADODB.Recordset
Dim cad As String
    
    On Error GoTo eFacturasGeneradas

    ListaFacturasGeneradas = ""

    SQL = "select nombre1, importe1 from tmpinformes where codusu = " & vUsu.Codigo
    SQL = SQL & " and nombre1 = " & DBSet(Trim(Tipo), "T")
    
    Set RS1 = New ADODB.Recordset
    RS1.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    cad = ""
    While Not RS1.EOF
        cad = cad & DBLet(RS1.Fields(1).Value, "N") & ","
    
        RS1.MoveNext
    Wend
    Set RS1 = Nothing
    
    'si hay facturas quitamos la ultima coma
    If cad <> "" Then cad = Mid(cad, 1, Len(cad) - 1)
    
    ListaFacturasGeneradas = cad
    Exit Function
    
eFacturasGeneradas:
    MuestraError Err.Number, "Cadena de Facturas Generadas", Err.Description
End Function




Public Function AnticiposLiquidados(tipoMov As String, DesNumFac As String, HasNumFac As String, fecfactu As String) As Boolean
Dim SQL As String

    AnticiposLiquidados = True

    SQL = "select count(*) from rfactsoc_anticipos where codtipomanti = " & DBSet(tipoMov, "T")
    SQL = SQL & " and numfactuanti >= " & DBSet(DesNumFac, "N")
    SQL = SQL & " and numfactuanti <= " & DBSet(HasNumFac, "N")
    SQL = SQL & " and fecfactuanti = " & DBSet(fecfactu, "F")
    
    AnticiposLiquidados = (TotalRegistros(SQL) <> 0)


End Function


Public Function CrearTMPs() As Boolean
' temporales de lineas para insertar posteriormente en rfactsoc_variedad y rfactsoc_calidad
Dim SQL As String
    
    On Error GoTo ECrear
    
    CrearTMPs = False
    
    'rfactsoc_variedad
    SQL = "CREATE TEMPORARY TABLE tmpFact_variedad ( "
    SQL = SQL & "`codtipom` char(3) NOT NULL ,"
    SQL = SQL & "`numfactu` int(7) unsigned NOT NULL,"
    SQL = SQL & "`fecfactu` date NOT NULL,"
    SQL = SQL & "`codvarie` int(6) NOT NULL,"
    SQL = SQL & "`codcampo` int(8) unsigned NOT NULL,"
    SQL = SQL & "`kilosnet` int(6) NOT NULL,"
    SQL = SQL & "`preciomed` decimal(6,4) NOT NULL,"
    SQL = SQL & "`imporvar` decimal(8,2) NOT NULL,"
    SQL = SQL & "`descontado` tinyint(1) NOT NULL default '0',"
    SQL = SQL & "`imporgasto` decimal(8,2) NOT NULL default '0',"
    SQL = SQL & "`kilogrado` decimal(10,2) NOT NULL default '0',"
    SQL = SQL & "`preciorea` decimal(6,4) NOT NULL default '0.0000')"
    
    conn.Execute SQL
    
    '[Monica]10/04/2017: añadimos las lineas de ingreso de coopic
    'rfactsoc_ingresos
    SQL = "CREATE TEMPORARY TABLE tmpFact_ingresos ( "
    SQL = SQL & "`codtipom` char(3) NOT NULL ,"
    SQL = SQL & "`numfactu` int(7) unsigned NOT NULL,"
    SQL = SQL & "`fecfactu` date NOT NULL,"
    SQL = SQL & "`codvarie` int(6) NOT NULL,"
    SQL = SQL & "`concepto` varchar(80),"
    SQL = SQL & "`importe` decimal(10,2) NOT NULL)"
    
    conn.Execute SQL
    
    'rfactsoc_calidad
    SQL = "CREATE TEMPORARY  TABLE tmpFact_calidad ( "
    SQL = SQL & "`codtipom` char(3),"
    SQL = SQL & "`numfactu` int(7) unsigned NOT NULL,"
    SQL = SQL & "`fecfactu` date NOT NULL,"
    SQL = SQL & "`codvarie` int(6) NOT NULL,"
    SQL = SQL & "`codcampo` int(8) unsigned NOT NULL,"
    SQL = SQL & "`codcalid` smallint(2) NOT NULL,"
    SQL = SQL & "`kilosnet` int(6) NOT NULL,"
    SQL = SQL & "`precio` decimal(6,4) NOT NULL,"
    SQL = SQL & "`imporcal` decimal(8,2) NOT NULL,"
    SQL = SQL & "`preciocalidad` decimal(6,4) NOT NULL,"
    SQL = SQL & "`imporcalidad` decimal(8,2) NOT NULL)"
    
    conn.Execute SQL
     
    ' si es liquidacion venta campo o no se insertaran en los anticipos
    SQL = "CREATE TEMPORARY  TABLE tmpFact_anticipos ( "
    SQL = SQL & "`codtipom` char(3) NOT NULL,"
    SQL = SQL & "`numfactu` int(7) unsigned NOT NULL,"
    SQL = SQL & "`fecfactu` date NOT NULL,"
    SQL = SQL & "`codtipomanti` char(3) NOT NULL,"
    SQL = SQL & "`numfactuanti` int(7) unsigned NOT NULL,"
    SQL = SQL & "`fecfactuanti` date NOT NULL,"
    SQL = SQL & "`codvarieanti` int(6) NOT NULL,"
    SQL = SQL & "`codcampoanti` int(8) unsigned NOT NULL,"
    SQL = SQL & "`baseimpo` decimal(8,2) NOT NULL) "

    conn.Execute SQL
     
    ' solo si es de bodega se insertaran los albaranes
    '[Monica] 08/04/2010: tambien si es la liquidacion de alzira
    SQL = "CREATE TEMPORARY TABLE `tmpFact_albaran` ("
    SQL = SQL & "`codtipom` char(3) NOT NULL,"
    SQL = SQL & "`numfactu` int(7) unsigned NOT NULL,"
    SQL = SQL & "`fecfactu` date NOT NULL,"
    SQL = SQL & "`numalbar` int(7) NOT NULL,"
    SQL = SQL & "`fecalbar` date NOT NULL,"
    SQL = SQL & "`codvarie` int(6) NOT NULL,"
    SQL = SQL & "`codcampo` int(8) unsigned NOT NULL,"
    SQL = SQL & "`kilosbru` int(6) NOT NULL,"
    SQL = SQL & "`kilosnet` int(6) NOT NULL,"
    SQL = SQL & "`grado` decimal(10,4) NOT NULL,"
    SQL = SQL & "`precio` decimal(6,4) NOT NULL,"
    SQL = SQL & "`importe` decimal(8,2) NOT NULL,"
    SQL = SQL & "`imporgasto` decimal(8,2) NOT NULL default '0.00',"
    SQL = SQL & "`prretirada` decimal(8,4) NOT NULL,"
    SQL = SQL & "`prmoltura` decimal(8,4) NOT NULL,"
    SQL = SQL & "`prenvasado` decimal(8,4) NOT NULL,"
    SQL = SQL & "`imppenal` decimal(8,2) NOT NULL default '0.00',"
    '[Monica]11/03/2016: nuevo campo de litros consumidos de aceite
    SQL = SQL & "`litrosconsu` int(6) NOT NULL default '0.00'"
    SQL = SQL & ")"
     
    conn.Execute SQL
     
    '
    '[Monica] 18/10/2010: por culpa de transporte de picassent
    SQL = "CREATE TEMPORARY TABLE `tmpFact_albarantra` ("
    SQL = SQL & "`codtipom` char(3) NOT NULL,"
    SQL = SQL & "`numfactu` int(7) unsigned NOT NULL,"
    SQL = SQL & "`fecfactu` date NOT NULL,"
    SQL = SQL & "`numalbar` int(7) NOT NULL,"
    SQL = SQL & "`fecalbar` date NOT NULL,"
    SQL = SQL & "`codvarie` int(6) NOT NULL,"
    SQL = SQL & "`codcampo` int(8) unsigned NOT NULL,"
    SQL = SQL & "`kilosbru` int(6) NOT NULL,"
    SQL = SQL & "`kilosnet` int(6) NOT NULL,"
    SQL = SQL & "`grado` decimal(10,4) NOT NULL,"
    SQL = SQL & "`precio` decimal(6,4) NOT NULL,"
    SQL = SQL & "`importe` decimal(8,2) NOT NULL,"
    SQL = SQL & "`imporgasto` decimal(8,2) NOT NULL default '0.00',"
    SQL = SQL & "`codtrans` varchar(10),"
    '[Monica]21/05/2013: añadimos la fecha de la nota de entrada
    SQL = SQL & "`fechaent` date NOT NULL)"
     
    conn.Execute SQL
    
    '[Monica] 05/12/2011: si es liquidacion de quatretonda hay retirada insertamos los anticipos de retirada que han intervenido
    SQL = "CREATE TEMPORARY  TABLE tmpFact_retirada ( "
    SQL = SQL & "`codtipom` char(3) NOT NULL,"
    SQL = SQL & "`numfactu` int(7) unsigned NOT NULL,"
    SQL = SQL & "`fecfactu` date NOT NULL,"
    SQL = SQL & "`codtipomanti` char(3) NOT NULL,"
    SQL = SQL & "`numfactuanti` int(7) unsigned NOT NULL,"
    SQL = SQL & "`fecfactuanti` date NOT NULL,"
    SQL = SQL & "`codvarieanti` int(6) NOT NULL,"
    SQL = SQL & "`codcampoanti` int(8) unsigned NOT NULL,"
    SQL = SQL & "`kilosret` int(8) unsigned NOT NULL,"
    SQL = SQL & "`imporret` decimal(8,2) NOT NULL) "

    conn.Execute SQL
    
    '[Monica] 15/04/2013: insertamos en la temporal de facturas varias
    SQL = "CREATE TEMPORARY  TABLE tmpFact_fvarias ( "
    SQL = SQL & "`codtipom` char(3) NOT NULL,"
    SQL = SQL & "`numfactu` int(7) unsigned NOT NULL,"
    SQL = SQL & "`fecfactu` date NOT NULL,"
    SQL = SQL & "`codtipomfvar` char(3) NOT NULL,"
    SQL = SQL & "`numfactufvar` int(7) unsigned NOT NULL,"
    SQL = SQL & "`fecfactufvar` date NOT NULL,"
    SQL = SQL & "`codsecci` smallint(3) unsigned NOT NULL)"

    conn.Execute SQL
    
    
'    'rfactsoc_incrementos
'    Sql = "CREATE TEMPORARY TABLE tmpFact_ingresos ( "
'    Sql = Sql & "`codtipom` char(3) NOT NULL ,"
'    Sql = Sql & "`numfactu` int(7) unsigned NOT NULL,"
'    Sql = Sql & "`fecfactu` date NOT NULL,"
'    Sql = Sql & "`codvarie` int(6) NOT NULL,"
'    Sql = Sql & "`concepto` varchar(80)"
'
'    conn.Execute Sql
    
    
    
    CrearTMPs = True
    
ECrear:
     If Err.Number <> 0 Then
        CrearTMPs = False
        'Borrar la tabla temporal
        SQL = " DROP TABLE IF EXISTS tmpFact_variedad;"
        conn.Execute SQL
        SQL = " DROP TABLE IF EXISTS tmpFact_calidad;"
        conn.Execute SQL
        SQL = " DROP TABLE IF EXISTS tmpFact_anticipos;"
        conn.Execute SQL
        SQL = " DROP TABLE IF EXISTS tmpFact_albaran;"
        conn.Execute SQL
        SQL = " DROP TABLE IF EXISTS tmpFact_albarantra;"
        conn.Execute SQL
        SQL = " DROP TABLE IF EXISTS tmpFact_fvarias;"
        conn.Execute SQL
        SQL = " DROP TABLE IF EXISTS tmpFact_ingresos;"
        conn.Execute SQL
    
    
    
    End If
End Function

Public Sub BorrarTMPs()
On Error Resume Next

    conn.Execute " DROP TABLE IF EXISTS tmpFact_variedad;"
    conn.Execute " DROP TABLE IF EXISTS tmpFact_calidad;"
    conn.Execute " DROP TABLE IF EXISTS tmpFact_anticipos;"
    conn.Execute " DROP TABLE IF EXISTS tmpFact_albaran;"
    conn.Execute " DROP TABLE IF EXISTS tmpFact_albarantra;"
    conn.Execute " DROP TABLE IF EXISTS tmpFact_retirada;"
    conn.Execute " DROP TABLE IF EXISTS tmpFact_fvarias;"
    conn.Execute " DROP TABLE IF EXISTS tmpFact_ingresos;"
    
    If Err.Number <> 0 Then Err.Clear
End Sub

Public Function PasarTemporales() As Boolean
On Error GoTo ePasar

    PasarTemporales = False
    '07/07/2014: añado el where
    conn.Execute " INSERT INTO rfactsoc_variedad SELECT * FROM tmpfact_variedad ;" 'where (codtipom, numfactu, fecfactu) in (select codtipom, numfactu, fecfactu from rfactsoc); "
    conn.Execute " INSERT INTO rfactsoc_calidad  SELECT * FROM tmpfact_calidad ;" ' where (codtipom, numfactu, fecfactu) in (select codtipom, numfactu, fecfactu from rfactsoc); "
    conn.Execute " INSERT INTO rfactsoc_anticipos  SELECT * FROM tmpfact_anticipos ;" ' where (codtipom, numfactu, fecfactu) in (select codtipom, numfactu, fecfactu from rfactsoc); "
    conn.Execute " INSERT INTO rfactsoc_albaran  SELECT * FROM tmpfact_albaran ;" ' where (codtipom, numfactu, fecfactu) in (select codtipom, numfactu, fecfactu from rfactsoc); "
    conn.Execute " INSERT INTO rfactsoc_retirada  SELECT * FROM tmpfact_retirada ;" ' where (codtipom, numfactu, fecfactu) in (select codtipom, numfactu, fecfactu from rfactsoc); "
    conn.Execute " INSERT INTO rfactsoc_fvarias  SELECT * FROM tmpfact_fvarias ;" ' where (codtipom, numfactu, fecfactu) in (select codtipom, numfactu, fecfactu from rfactsoc); "
    
    '10/04/2017: ingresos de la factura
    conn.Execute " INSERT INTO rfactsoc_ingresos  SELECT * FROM tmpfact_ingresos ;" ' where (codtipom, numfactu, fecfactu) in (select codtipom, numfactu, fecfactu from rfactsoc); "
    
    PasarTemporales = True
    Exit Function
ePasar:
    MuestraError "Pasar Temporales", Err.Description
End Function

Public Function PasarTemporalesTrans() As Boolean
On Error GoTo ePasar

    PasarTemporalesTrans = False

    '[Monica]30/04/2014: faltaba la columna fechaent (añadida)
    conn.Execute " INSERT INTO rfacttra_albaran (codtipom,numfactu,fecfactu,numalbar,fecalbar,codvarie,codcampo,numnotac,kilosnet,importe,codtrans,fechaent) SELECT codtipom,numfactu,fecfactu,numalbar,fecalbar,codvarie,codcampo,kilosbru,kilosnet,importe,codtrans,fechaent FROM tmpFact_albarantra; "
    
    PasarTemporalesTrans = True
    Exit Function
ePasar:
    MuestraError "Pasar Temporales", Err.Descripc
End Function

'Marcar Factura como contabilizada y como pendiente de recibir nro de factura
Public Function MarcarFactura(tipoMov As String, numfactu As String, FecFac As String, MarcarPdteRecibirFra As Boolean) As Boolean

    Dim SQL As String
    
    On Error GoTo eMarcarFactura
    
    MensError = ""
    MarcarFactura = False
    
    SQL = "update rfactsoc set contabilizado = 1 "
    
    '[Monica]22/06/2017: en el caso de natural de montaña se marca como contabilizada pq no integran en la contabilidad, y no esperan nro de fra
    If MarcarPdteRecibirFra Then SQL = SQL & ", pdtenrofact = 1 "
    
    SQL = SQL & " where codtipom = " & DBSet(tipoMov, "T")
    SQL = SQL & " and numfactu = " & DBSet(numfactu, "N") & " and fecfactu = " & DBSet(FecFac, "F")
    
    conn.Execute SQL
    
    MarcarFactura = True
    
    Exit Function

eMarcarFactura:
    MensError = "Error en el proceso de marcar la factura " & numfactu & " del socio " & vSocio.Codigo
    MuestraError Err.Number, MensError
End Function





Public Function ActualizarRegistros(cTabla As String, cWhere As String) As Boolean
'Actualizar la marca de impreso
Dim SQL As String

    On Error GoTo eActualizarRegistros

    ActualizarRegistros = False
    cTabla = QuitarCaracterACadena(cTabla, "{")
    cTabla = QuitarCaracterACadena(cTabla, "}")
    SQL = "update " & QuitarCaracterACadena(cTabla, "_1") & " set impreso = 1 "
    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        SQL = SQL & " WHERE " & cWhere
    End If
    
    conn.Execute SQL
    
    ActualizarRegistros = True
    Exit Function
    
eActualizarRegistros:
    MuestraError Err.Number, "Actualizando registros", Err.Description
End Function



Public Function ComprobarTiposMovimiento(Tipo As Byte, cTabla As String, cWhere As String, Optional EsVetoRuso As Boolean, Optional EsTerceros As Boolean) As Boolean
Dim SQL As String
Dim Rs As ADODB.Recordset

Dim NumError As Long
Dim TipoMovim As String
Dim Encontrado As Boolean
Dim HayReg As Byte

    On Error GoTo eComprobarTiposMovimiento

    ComprobarTiposMovimiento = False
    
    '[Monica]23/12/2014: si es veto ruso es otro contador de anticipo
    If EsVetoRuso Then
        SQL = "select count(*) from usuarios.stipom where codtipom = 'VAA'" ' anticipo de veto ruso
        If TotalRegistros(SQL) = 0 Then
            MsgBox "No existe el Tipo de Movimiento : VAA", vbExclamation
            Exit Function
        End If
    End If
    
    '[Monica]29/05/2017: si esterceros solo en el caso de picassent
    If EsTerceros Then
        SQL = "select count(*) from usuarios.stipom where codtipom = 'CAT'" ' anticipo v.campo tercero
        If TotalRegistros(SQL) = 0 Then
            MsgBox "No existe el Tipo de Movimiento : CAT", vbExclamation
            Exit Function
        End If
        SQL = "select count(*) from usuarios.stipom where codtipom = 'CLT'" ' liquidacion v.campo tercero
        If TotalRegistros(SQL) = 0 Then
            MsgBox "No existe el Tipo de Movimiento : CLT", vbExclamation
            Exit Function
        End If
    End If
    
    Select Case Tipo
        Case 0 ' anticipos
            SQL = "SELECT distinct rcoope.codtipomant as codtipom "
        Case 1 ' liquidaciones
            SQL = "SELECT distinct rcoope.codtipomliq as codtipom  "
        Case 2 ' anticipos venta campos
            SQL = "SELECT distinct rcoope.codtipomantvc as codtipom  "
        Case 3 ' liquidaciones venta campos
            SQL = "SELECT distinct rcoope.codtipomliqvc as codtipom  "
        Case 7 ' anticipos almazara
            SQL = "SELECT distinct rcoope.codtipomantalmz as codtipom  "
        Case 8 ' liquidacion almazara
            SQL = "SELECT distinct rcoope.codtipomliqalmz as codtipom  "
        Case 9 ' anticipos bodega
            SQL = "SELECT distinct rcoope.codtipomantbod as codtipom  "
        Case 10 ' liquidacion bodega
            SQL = "SELECT distinct rcoope.codtipomliqbod as codtipom  "
    End Select

    SQL = SQL & " FROM  (" & cTabla & ") INNER JOIN rcoope On rsocios.codcoope = rcoope.codcoope "

    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        SQL = SQL & " WHERE " & cWhere
    End If
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    HayReg = 0
    Encontrado = False
    While Not Rs.EOF And Not Encontrado
        HayReg = 1
        TipoMovim = DBLet(Rs!CodTipom, "T")
        SQL = "select count(*) from usuarios.stipom where codtipom = " & DBSet(Rs!CodTipom, "T")
        Select Case Tipo
            Case 0 ' anticipos
                SQL = SQL & " and tipodocu = 1 "
            Case 1 ' liquidaciones
                SQL = SQL & " and tipodocu = 2  "
            Case 2 ' anticipos venta campos
                SQL = SQL & " and tipodocu = 3  "
            Case 3 ' liquidaciones venta campos
                SQL = SQL & " and tipodocu = 4  "
            Case 7 ' anticipos almazara
                SQL = SQL & " and tipodocu = 7  "
            Case 8 ' liquidacion almazara
                SQL = SQL & " and tipodocu = 8  "
            Case 9 ' anticipo bodega
                SQL = SQL & " and tipodocu = 9  "
            Case 10 ' liquidacion bodega
                SQL = SQL & " and tipodocu = 10  "
        End Select
    
        Encontrado = (TotalRegistros(SQL) = 0)
    
        Rs.MoveNext
    Wend
    Set Rs = Nothing
    If HayReg = 1 Then
        If Encontrado Then
            MsgBox "No existe el Tipo de Movimiento : " & TipoMovim, vbExclamation
        Else
            ComprobarTiposMovimiento = True
        End If
    Else
        MsgBox "No se han encontrado movimientos. Revise.", vbExclamation
    End If
    
    Exit Function
    
    
eComprobarTiposMovimiento:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Comprobar Tipos de Movimiento", Err.Description
        ComprobarTiposMovimiento = False
    End If
End Function


Public Function FacturacionLiquidacionesValsur(cTabla As String, cWhere As String, FecFac As String, Pb1 As ProgressBar, Complementaria As Byte) As Boolean
Dim SQL As String
Dim Rs As ADODB.Recordset
Dim RS1 As ADODB.Recordset

Dim AntSocio As String
Dim AntVarie As String
Dim ActSocio As String
Dim ActVarie As String
Dim actCampo As String
Dim AntCampo As String
Dim ActCalid As String
Dim AntCalid As String

Dim HayReg As Boolean

Dim NumError As Long
Dim vImporte As Currency
Dim vPorcIva As String

Dim PrimFac As String
Dim UltFac As String

Dim tipoMov As String
Dim B As Boolean
Dim vSeccion As CSeccion
Dim Kilos As Currency
Dim KilosCal As Currency
Dim Importe As Currency

Dim devuelve As String
Dim Existe As Boolean

'Dim baseimpo As Currency
'Dim BaseReten As Currency
'Dim Neto As Currency
'Dim ImpoIva As Currency
'Dim ImpoReten As Currency
'Dim TotalFac As Currency
Dim Recolect As Byte
Dim vPrecio As Currency

Dim Sql2 As String
Dim Sql3 As String

Dim GastosCoop As Currency
Dim GastosAlb As Currency
Dim vPorcGasto As String


    On Error GoTo eFacturacion

    FacturacionLiquidacionesValsur = False
    
'08052009 antes dentro de transacciones
    BorrarTMPs
    B = CrearTMPs()
    If Not B Then
         Exit Function
    End If
'08052009
    
    conn.BeginTrans
    
    tipoMov = "FAL"
    
    
    SQL = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute SQL

    SQL = "SELECT  rhisfruta.codsocio, rhisfruta.codvarie,"
    SQL = SQL & "rhisfruta.codcampo, rhisfruta.recolect, rhisfruta_clasif.codcalid, "
'[Monica]01/09/2010 : sustituida la siguiente linea por
'    Sql = Sql & "rprecios_calidad.precoop, rprecios_calidad.presocio, rprecios_calidad.tipofact,sum(rhisfruta_clasif.kilosnet) as kilosnet "
     SQL = SQL & "rprecios.fechaini, rprecios.fechafin, rprecios_calidad.tipofact,max(rprecios.contador) contador, sum(rhisfruta_clasif.kilosnet) as kilosnet "
    
    SQL = SQL & " FROM  " & cTabla

    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        SQL = SQL & " WHERE " & cWhere
    End If
    ' ordenado por socio, variedad, campo, calidad
    SQL = SQL & " group by rhisfruta.codsocio, rhisfruta.codvarie, rhisfruta.codcampo, rhisfruta_clasif.codcalid, rhisfruta.recolect "
    SQL = SQL & " order by rhisfruta.codsocio, rhisfruta.codvarie, rhisfruta.codcampo, rhisfruta_clasif.codcalid, rhisfruta.recolect "
    
    Set vSeccion = New CSeccion
    
    If vSeccion.LeerDatos(vParamAplic.Seccionhorto) Then
        If Not vSeccion.AbrirConta Then
            Exit Function
        End If
    End If
    
    HayReg = False
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    If Not Rs.EOF Then
        AntSocio = CStr(DBLet(Rs!Codsocio, "N"))
        AntVarie = CStr(DBLet(Rs!Codvarie, "N"))
        AntCampo = CStr(DBLet(Rs!codCampo, "N"))
        AntCalid = CStr(DBLet(Rs!codcalid, "N"))
        
        ActSocio = CStr(DBLet(Rs!Codsocio, "N"))
        ActVarie = CStr(DBLet(Rs!Codvarie, "N"))
        actCampo = CStr(DBLet(Rs!codCampo, "N"))
        ActCalid = CStr(DBLet(Rs!codcalid, "N"))
    
        Set vSocio = New cSocio
        If vSocio.LeerDatos(ActSocio) Then
            If vSocio.LeerDatosSeccion(ActSocio, vParamAplic.Seccionhorto) Then
                baseimpo = 0
                BaseReten = 0
                ImpoIva = 0
                ImpoReten = 0
                TotalFac = 0
                ImpoAFO = 0
                BaseAFO = 0
                
                Anticipos = 0
                
                Kilos = 0
                Importe = 0
                
                KilosCal = 0
                
                vPorcIva = ""
                vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vSocio.CodIva, "N")
                PorcIva = CCur(ImporteSinFormato(vPorcIva))
                
                tipoMov = vSocio.CodTipomLiq
                
                Set vTipoMov = New CTiposMov
                
                numfactu = vTipoMov.ConseguirContador(tipoMov)
                Do
                    numfactu = vTipoMov.ConseguirContador(tipoMov)
                    devuelve = DevuelveDesdeBDNew(cAgro, "rfactsoc", "numfactu", "codtipom", tipoMov, "T", , "numfactu", CStr(numfactu), "N", "fecfactu", FecFac, "F")
                    If devuelve <> "" Then
                        'Ya existe el contador incrementarlo
                        Existe = True
                        vTipoMov.IncrementarContador (tipoMov)
                        numfactu = vTipoMov.ConseguirContador(tipoMov)
                    Else
                        Existe = False
                    End If
                Loop Until Not Existe
                
                vParamAplic.PrimFactLiq = numfactu
                
            End If
        End If
    End If
    
    While Not Rs.EOF And B
        ActCalid = DBLet(Rs!codcalid, "N")
        ActVarie = DBLet(Rs!Codvarie, "N")
        actCampo = DBSet(Rs!codCampo, "N")
        ActSocio = DBSet(Rs!Codsocio, "N")
        
        If (ActCalid <> AntCalid Or AntCampo <> actCampo Or AntVarie <> ActVarie Or AntSocio <> ActSocio) Then
            ' kilos e importe por variedad campo
            Kilos = Kilos + KilosCal
            Importe = Importe + vImporte
            
            baseimpo = baseimpo + vImporte
            
            B = InsertLineaCalidad(tipoMov, CStr(numfactu), FecFac, AntVarie, AntCampo, CStr(AntCalid), CStr(DBLet(KilosCal, "N")), CStr(vImporte))
            KilosCal = 0
            vImporte = 0
            
            AntCalid = ActCalid
        End If
        
        If (ActVarie <> AntVarie Or actCampo <> AntCampo Or ActSocio <> AntSocio) Then
            If B Then ' descontamos el porcentaje de gastos de cooperativa
                GastosCoop = 0
                
                vPorcGasto = ""
                vPorcGasto = DevuelveDesdeBDNew(cAgro, "rcoope", "porcgast", "codcoope", vSocio.Cooperativa, "N")
                If vPorcGasto = "" Then vPorcGasto = "0"
                
                GastosCoop = Round2(Importe * CCur(ImporteSinFormato(vPorcGasto)) / 100, 2)
                
                Importe = Importe - GastosCoop
                baseimpo = baseimpo - GastosCoop
                
            End If
            
            If B Then ' descontamos los gastos de los albaranes
'[MONICA] 08/09/2009 : los gastos de transporte se suman en ObtenerGastosAlbaranes, quito lo de David
'                '17 AGOSTO 2009
'                ' David###   Para VALSUR los gastos se suman
'                If vParamAplic.Cooperativa = 1 Then
'                    GastosAlb = ObtenerGastosAlbaranes(AntSocio, AntVarie, AntCampo, cTabla, cWhere)
'                    Importe = Importe + GastosAlb
'                    baseimpo = baseimpo + GastosAlb
'
'                Else
'                    'Para el resto sigue como estaba
                    If vParamAplic.Cooperativa = 3 Then
                        GastosAlb = ObtenerGastosAlbaranes(AntSocio, AntVarie, AntCampo, cTabla, cWhere, 1, , Complementaria)
                    Else
                        GastosAlb = ObtenerGastosAlbaranes(AntSocio, AntVarie, AntCampo, cTabla, cWhere, , , Complementaria)
                    End If
                    
                    Importe = Importe - GastosAlb
                    baseimpo = baseimpo - GastosAlb
'                End If
            End If
                        
            ' insertar linea de variedad, campo
            If B Then
                B = InsertLinea(tipoMov, CStr(numfactu), FecFac, AntVarie, AntCampo, CStr(Kilos), CStr(Importe), "0")
            End If
            
            If B Then
                ' tenemos que descontar los anticipos que tengamos para ello
                Sql2 = "select rfactsoc_variedad.imporvar, rfactsoc_variedad.numfactu, rfactsoc_variedad.fecfactu "
                Sql2 = Sql2 & " from rfactsoc_variedad INNER JOIN rfactsoc ON rfactsoc_variedad.codtipom = rfactsoc.codtipom and "
                Sql2 = Sql2 & " rfactsoc_variedad.numfactu = rfactsoc.numfactu and rfactsoc_variedad.fecfactu = rfactsoc.fecfactu "
                Sql2 = Sql2 & " where rfactsoc_variedad.codtipom = " & DBSet(vSocio.CodTipomAnt, "T") ' antes era 'FAA' "
                Sql2 = Sql2 & " and rfactsoc.codsocio = " & DBSet(AntSocio, "N")
                Sql2 = Sql2 & " and codvarie = " & DBSet(AntVarie, "N")
                Sql2 = Sql2 & " and codcampo = " & DBSet(AntCampo, "N")
                Sql2 = Sql2 & " and rfactsoc_variedad.descontado = 0"
                
                Set RS1 = New ADODB.Recordset
                RS1.Open Sql2, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
                
                While Not RS1.EOF
                    baseimpo = baseimpo - DBLet(RS1.Fields(0).Value, "N")
                    Anticipos = Anticipos + DBLet(RS1.Fields(0).Value, "N")
                    
                    ' indicamos que los anticipos ya han sido descontados
                    Sql3 = "update rfactsoc_variedad set descontado = 1 where codtipom = " & DBSet(vSocio.CodTipomAnt, "T") ' antes era 'FAC'
                    Sql3 = Sql3 & " and numfactu = " & DBSet(RS1!numfactu, "N")
                    Sql3 = Sql3 & " and fecfactu = " & DBSet(RS1!fecfactu, "F") & " and codvarie = " & DBSet(AntVarie, "N")
                    Sql3 = Sql3 & " and codcampo = " & DBSet(AntCampo, "N")
                    
                    conn.Execute Sql3
                    
                    ' insertamos en la tabla de anticipos de liquidacion venta campo
                    Sql3 = "insert into tmpFact_anticipos (codtipom, numfactu, fecfactu, codtipomanti, numfactuanti, fecfactuanti, codvarieanti, codcampoanti, baseimpo) values ("
                    Sql3 = Sql3 & DBSet(vSocio.CodTipomLiq, "T") & "," ' antes era 'FAL'
                    Sql3 = Sql3 & DBSet(numfactu, "N") & "," & DBSet(FecFac, "F") & ","
                    Sql3 = Sql3 & DBSet(vSocio.CodTipomAnt, "T") & "," ' antes era 'FAA'
                    Sql3 = Sql3 & DBSet(RS1!numfactu, "N") & "," & DBSet(RS1!fecfactu, "F") & ","
                    Sql3 = Sql3 & DBSet(AntVarie, "N") & "," & DBSet(AntCampo, "N") & "," & DBSet(RS1!imporvar, "N") & ")"
                    
                    conn.Execute Sql3
                    
                    RS1.MoveNext
                Wend
                
                Set RS1 = Nothing
                ' fin descontar anticipos
            
            End If
            
            
            If B Then
                AntVarie = ActVarie
                AntCampo = actCampo
                Kilos = 0
                Importe = 0
            End If
        End If
        
        If ActSocio <> AntSocio Then
            
            ImpoIva = Round2(baseimpo * PorcIva / 100, 2)
        
            Select Case DBLet(vSocio.TipoIRPF, "N")
                Case 0
                    ImpoReten = Round2((baseimpo + ImpoIva) * vParamAplic.PorcreteFacSoc / 100, 2)
                    BaseReten = (baseimpo + ImpoIva)
                    PorcReten = vParamAplic.PorcreteFacSoc
                Case 1
                    ImpoReten = Round2(baseimpo * vParamAplic.PorcreteFacSoc / 100, 2)
                    BaseReten = baseimpo
                    PorcReten = vParamAplic.PorcreteFacSoc
                Case 2
                    ImpoReten = 0
                    BaseReten = 0
                    PorcReten = 0
            End Select
            
            ImpoAFO = Round2((baseimpo + Anticipos) * vParamAplic.PorcenAFO / 100, 2)
            BaseAFO = baseimpo + Anticipos
            PorcAFO = vParamAplic.PorcenAFO
        
            TotalFac = baseimpo + ImpoIva - ImpoReten - ImpoAFO
            
            IncrementarProgresNew Pb1, 1
            
            
            '[Monica]07/02/2012: indicamos si es una factura de liquidacion complementaria
            'insertar cabecera de factura
            B = InsertCabecera(tipoMov, CStr(numfactu), FecFac, , , (Complementaria = 1))
            
            If B Then B = InsertResumen(tipoMov, CStr(numfactu))
            
            If B Then B = RecalcularCalidades(tipoMov, CStr(numfactu), FecFac)
            
            If B Then B = vTipoMov.IncrementarContador(tipoMov)
            
            If B Then
                AntSocio = ActSocio
                
                Set vSocio = Nothing
                Set vSocio = New cSocio
                If vSocio.LeerDatos(ActSocio) Then
                    If vSocio.LeerDatosSeccion(AntSocio, vParamAplic.Seccionhorto) Then
                        vPorcIva = ""
                        vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vSocio.CodIva, "N")
                        PorcIva = CCur(ImporteSinFormato(vPorcIva))
                    End If
                    
                    tipoMov = vSocio.CodTipomLiq
                End If
                baseimpo = 0
                BaseReten = 0
                ImpoIva = 0
                ImpoReten = 0
                TotalFac = 0
                BaseAFO = 0
                ImpoAFO = 0
                
                Anticipos = 0
                
                numfactu = vTipoMov.ConseguirContador(tipoMov)
                Do
                    numfactu = vTipoMov.ConseguirContador(tipoMov)
                    devuelve = DevuelveDesdeBDNew(cAgro, "rfactsoc", "numfactu", "codtipom", tipoMov, "T", , "numfactu", CStr(numfactu), "N", "fecfactu", FecFac, "F")
                    If devuelve <> "" Then
                        'Ya existe el contador incrementarlo
                        Existe = True
                        vTipoMov.IncrementarContador (tipoMov)
                        numfactu = vTipoMov.ConseguirContador(tipoMov)
                    Else
                        Existe = False
                    End If
                Loop Until Not Existe
           End If
        End If
        
        Recolect = DBLet(Rs!Recolect, "N")
        
        '[Monica]01/09/2010: añadido ésto, antes los precios los sacabamos en el propio select
        Dim Sql9 As String
        Dim Rs9 As ADODB.Recordset
        Dim PreCoop As Currency
        Dim PreSocio As Currency
        
        Sql9 = "select precoop, presocio from rprecios_calidad where codvarie = " & DBSet(Rs!Codvarie, "N")
        Sql9 = Sql9 & " and tipofact = " & DBSet(Rs!TipoFact, "N")
        Sql9 = Sql9 & " and contador = " & DBSet(Rs!Contador, "N")
        Sql9 = Sql9 & " and codcalid = " & DBSet(Rs!codcalid, "N")
        
        Set Rs9 = New ADODB.Recordset
        Rs9.Open Sql9, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
        
        If Not Rs9.EOF Then
            PreCoop = DBLet(Rs9.Fields(0).Value, "N")
            PreSocio = DBLet(Rs9.Fields(1).Value, "N")
        
            Select Case Recolect
                Case 0
                    vPrecio = DBLet(PreCoop, "N")
                    vImporte = vImporte + Round2(DBLet(Rs!KilosNet, "N") * PreCoop, 2)
                Case 1
                    vPrecio = DBLet(PreSocio, "N")
                    vImporte = vImporte + Round2(DBLet(Rs!KilosNet, "N") * PreSocio, 2)
            End Select
            
            KilosCal = KilosCal + DBLet(Rs!KilosNet, "N")
        End If
        'hasta aqui
        
        HayReg = True
        
        Rs.MoveNext
    Wend
    
    ' ultimo registro si ha entrado
    If B And HayReg Then
        ' insertar linea de calidad
        Kilos = Kilos + KilosCal
        Importe = Importe + vImporte
        
        baseimpo = baseimpo + vImporte
        
        
        If B Then B = InsertLineaCalidad(tipoMov, CStr(numfactu), FecFac, ActVarie, actCampo, CStr(ActCalid), CStr(DBLet(KilosCal, "N")), CStr(vImporte))
        
        
        If B Then ' descontamos el porcentaje de gastos de cooperativa
            GastosCoop = 0
            
            vPorcGasto = ""
            vPorcGasto = DevuelveDesdeBDNew(cAgro, "rcoope", "porcgast", "codcoope", vSocio.Cooperativa, "N")
            If vPorcGasto = "" Then vPorcGasto = "0"
            
            GastosCoop = Round2(Importe * CCur(ImporteSinFormato(vPorcGasto)) / 100, 2)
            
            Importe = Importe - GastosCoop
            baseimpo = baseimpo - GastosCoop
        End If
        
        If B Then ' descontamos los gastos de los albaranes
'[MONICA] 08/09/2009 : los gastos de transporte se suman en ObtenerGastosAlbaranes, quito lo de David
'            '17 AGOSTO 2009
'            ' David###   Para VALSUR los gastos se suman
'            GastosAlb = ObtenerGastosAlbaranes(AntSocio, AntVarie, AntCampo, cTabla, cWhere)
'            If vParamAplic.Cooperativa = 1 Then
'                Importe = Importe + GastosAlb
'                baseimpo = baseimpo + GastosAlb
'            Else
'                Importe = Importe - GastosAlb
'                baseimpo = baseimpo - GastosAlb
'            End If
        
            '[Monica]29/09/2017: Para el resto sigue como estaba
            If vParamAplic.Cooperativa = 3 Then
                GastosAlb = ObtenerGastosAlbaranes(AntSocio, AntVarie, AntCampo, cTabla, cWhere, 1, , Complementaria)
            Else
                GastosAlb = ObtenerGastosAlbaranes(AntSocio, AntVarie, AntCampo, cTabla, cWhere, , , Complementaria)
            End If
            
            Importe = Importe - GastosAlb
            baseimpo = baseimpo - GastosAlb
        
        
        End If
        
                    
        ' insertar linea de variedad
        If B Then B = InsertLinea(tipoMov, CStr(numfactu), FecFac, CStr(ActVarie), CStr(actCampo), CStr(Kilos), CStr(Importe), "0")
        
        ' tenemos que descontar los anticipos que tengamos para ello
        If B Then
            Sql2 = "select rfactsoc_variedad.imporvar, rfactsoc_variedad.numfactu, rfactsoc_variedad.fecfactu "
            Sql2 = Sql2 & " from rfactsoc_variedad INNER JOIN rfactsoc ON rfactsoc_variedad.codtipom = rfactsoc.codtipom and "
            Sql2 = Sql2 & " rfactsoc_variedad.numfactu = rfactsoc.numfactu and rfactsoc_variedad.fecfactu = rfactsoc.fecfactu "
            Sql2 = Sql2 & " where rfactsoc_variedad.codtipom = " & DBSet(vSocio.CodTipomAnt, "T") ' antes era 'FAA' "
            Sql2 = Sql2 & " and rfactsoc.codsocio = " & DBSet(ActSocio, "N")
            Sql2 = Sql2 & " and codvarie = " & DBSet(ActVarie, "N")
            Sql2 = Sql2 & " and codcampo = " & DBSet(actCampo, "N")
            Sql2 = Sql2 & " and rfactsoc_variedad.descontado = 0"
            
            Set RS1 = New ADODB.Recordset
            RS1.Open Sql2, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
            
            While Not RS1.EOF
                baseimpo = baseimpo - DBLet(RS1.Fields(0).Value, "N")
                Anticipos = Anticipos + DBLet(RS1.Fields(0).Value, "N")
                
                ' indicamos que los anticipos ya han sido descontados
                Sql3 = "update rfactsoc_variedad set descontado = 1 where codtipom = " & DBSet(vSocio.CodTipomAnt, "T") ' antes era 'FAC'
                Sql3 = Sql3 & " and numfactu = " & DBSet(RS1!numfactu, "N")
                Sql3 = Sql3 & " and fecfactu = " & DBSet(RS1!fecfactu, "F") & " and codvarie = " & DBSet(ActVarie, "N")
                Sql3 = Sql3 & " and codcampo = " & DBSet(actCampo, "N")
                
                conn.Execute Sql3
                
                ' insertamos en la tabla de anticipos de liquidacion venta campo
                Sql3 = "insert into tmpFact_anticipos (codtipom, numfactu, fecfactu, codtipomanti, numfactuanti, fecfactuanti, codvarieanti, codcampoanti, baseimpo) values ("
                Sql3 = Sql3 & DBSet(vSocio.CodTipomLiq, "T") & "," ' antes era 'FAL'
                Sql3 = Sql3 & DBSet(numfactu, "N") & "," & DBSet(FecFac, "F") & ","
                Sql3 = Sql3 & DBSet(vSocio.CodTipomAnt, "T") & "," ' antes era 'FAA'
                Sql3 = Sql3 & DBSet(RS1!numfactu, "N") & "," & DBSet(RS1!fecfactu, "F") & ","
                Sql3 = Sql3 & DBSet(ActVarie, "N") & "," & DBSet(actCampo, "N") & "," & DBSet(RS1!imporvar, "N") & ")"
                
                conn.Execute Sql3
                
                RS1.MoveNext
            Wend
            
            Set RS1 = Nothing
            ' fin descontar anticipos
        
        End If
        
        
        ImpoIva = Round2(baseimpo * PorcIva / 100, 2)

        Select Case DBLet(vSocio.TipoIRPF, "N")
            Case 0
                ImpoReten = Round2((baseimpo + ImpoIva) * vParamAplic.PorcreteFacSoc / 100, 2)
                BaseReten = (baseimpo + ImpoIva)
                PorcReten = vParamAplic.PorcreteFacSoc
            Case 1
                ImpoReten = Round2(baseimpo * vParamAplic.PorcreteFacSoc / 100, 2)
                BaseReten = baseimpo
                PorcReten = vParamAplic.PorcreteFacSoc
            Case 2
                ImpoReten = 0
                BaseReten = 0
                PorcReten = 0
        End Select
        
        ImpoAFO = Round2((baseimpo + Anticipos) * vParamAplic.PorcenAFO / 100, 2)
        BaseAFO = baseimpo + Anticipos
        PorcAFO = vParamAplic.PorcenAFO

        TotalFac = baseimpo + ImpoIva - ImpoReten - ImpoAFO
        
        IncrementarProgresNew Pb1, 1
        
        vParamAplic.UltFactLiq = numfactu
        
        '[Monica]07/02/2012: indicamos si es una factura de liquidacion complementaria
        'insertar cabecera de factura
        B = InsertCabecera(tipoMov, CStr(numfactu), FecFac, , , (Complementaria = 1))
        
        If B Then B = InsertResumen(tipoMov, CStr(numfactu))
        
        If B Then B = RecalcularCalidades(tipoMov, CStr(numfactu), FecFac)
        
        If B Then B = vTipoMov.IncrementarContador(tipoMov)
        
        'pasamos las temporales a las tablas
        If B Then B = PasarTemporales()
        
        If B Then B = (vParamAplic.Modificar = 1)
    End If
    
'    BorrarTMPs
    
    vSeccion.CerrarConta
    Set vSeccion = Nothing
    Set vSocio = Nothing
    
eFacturacion:
    If Err.Number <> 0 Or Not B Then
        conn.RollbackTrans
        FacturacionLiquidacionesValsur = False
    Else
        conn.CommitTrans
        FacturacionLiquidacionesValsur = True
    End If
End Function



Public Function FacturacionLiquidacionesAlzira(cTabla As String, cWhere As String, FecFac As String, Pb1 As ProgressBar, TipoPrec As Byte) As Boolean
Dim SQL As String
Dim Rs As ADODB.Recordset
Dim RS1 As ADODB.Recordset

Dim AntSocio As String
Dim AntVarie As String
Dim ActSocio As String
Dim ActVarie As String
Dim actCampo As String
Dim AntCampo As String
Dim ActCalid As String
Dim AntCalid As String

Dim HayReg As Boolean

Dim NumError As Long
Dim vImporte As Currency
Dim vPorcIva As String

Dim PrimFac As String
Dim UltFac As String

Dim tipoMov As String
Dim B As Boolean
Dim vSeccion As CSeccion
Dim Kilos As Currency
Dim KilosCal As Currency
Dim Importe As Currency

Dim devuelve As String
Dim Existe As Boolean

'Dim baseimpo As Currency
'Dim BaseReten As Currency
'Dim Neto As Currency
'Dim ImpoIva As Currency
'Dim ImpoReten As Currency
'Dim TotalFac As Currency
Dim Recolect As Byte
Dim vPrecio As Currency

Dim Sql2 As String
Dim Sql3 As String
Dim SqlAlbaranes As String

Dim GastosCoop As Currency
Dim GastosAlb As Currency
Dim vPorcGasto As String


    On Error GoTo eFacturacion

    FacturacionLiquidacionesAlzira = False
    
'08052009 antes dentro de transacciones
    BorrarTMPs
    B = CrearTMPs()
    If Not B Then
         Exit Function
    End If
'08052009
    
    conn.BeginTrans
    
    tipoMov = "FAL"
    
    
    SQL = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute SQL

    SQL = "SELECT  rhisfruta.codsocio, rhisfruta.codvarie,"
    SQL = SQL & "rhisfruta.codcampo, rhisfruta.recolect, rhisfruta_clasif.codcalid, "               '[Monica]28/03/2013: Añadido el if dentro del sum
    SQL = SQL & "rprecios_calidad.precoop, rprecios_calidad.presocio, rprecios_calidad.tipofact, sum(if(rhisfruta_clasif.kilosnet is null,0, rhisfruta_clasif.kilosnet)) as kilosnet "
    SQL = SQL & " FROM  " & cTabla


    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        SQL = SQL & " WHERE " & cWhere
    End If
     
    
    ' ordenado por socio, variedad, campo, calidad
    SQL = SQL & " group by rhisfruta.codsocio, rhisfruta.codvarie, rhisfruta.codcampo, rhisfruta_clasif.codcalid, rhisfruta.recolect, rprecios_calidad.precoop, rprecios_calidad.presocio, rprecios_calidad.tipofact "
    '[Monica]28/03/2013: Añadido el having
    SQL = SQL & " having  sum(if(rhisfruta_clasif.kilosnet is null,0, rhisfruta_clasif.kilosnet)) <> 0"
    SQL = SQL & " order by rhisfruta.codsocio, rhisfruta.codvarie, rhisfruta.codcampo, rhisfruta_clasif.codcalid, rhisfruta.recolect, rprecios_calidad.precoop, rprecios_calidad.presocio, rprecios_calidad.tipofact "
    
    Set vSeccion = New CSeccion
    
    If vSeccion.LeerDatos(vParamAplic.Seccionhorto) Then
        If Not vSeccion.AbrirConta Then
            Exit Function
        End If
    End If
    
    HayReg = False
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    If Not Rs.EOF Then
        AntSocio = CStr(DBLet(Rs!Codsocio, "N"))
        AntVarie = CStr(DBLet(Rs!Codvarie, "N"))
        AntCampo = CStr(DBLet(Rs!codCampo, "N"))
        AntCalid = CStr(DBLet(Rs!codcalid, "N"))
        
        ActSocio = CStr(DBLet(Rs!Codsocio, "N"))
        ActVarie = CStr(DBLet(Rs!Codvarie, "N"))
        actCampo = CStr(DBLet(Rs!codCampo, "N"))
        ActCalid = CStr(DBLet(Rs!codcalid, "N"))
    
        Set vSocio = New cSocio
        If vSocio.LeerDatos(ActSocio) Then
            If vSocio.LeerDatosSeccion(ActSocio, vParamAplic.Seccionhorto) Then
                baseimpo = 0
                BaseReten = 0
                ImpoIva = 0
                ImpoReten = 0
                TotalFac = 0
                ImpoAFO = 0
                BaseAFO = 0
                
                Anticipos = 0
                
                Kilos = 0
                Importe = 0
                
                KilosCal = 0
                
                vPorcIva = ""
                '[Monica]29/04/2011: INTERNAS
                If vSocio.EsFactADVInt Then
                    vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vParamAplic.CodIvaExeADV, "N")
                Else
                    vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vSocio.CodIva, "N")
                End If
                PorcIva = CCur(ImporteSinFormato(vPorcIva))
                
                tipoMov = vSocio.CodTipomLiq
                
                Set vTipoMov = New CTiposMov
                
                numfactu = vTipoMov.ConseguirContador(tipoMov)
                Do
                    numfactu = vTipoMov.ConseguirContador(tipoMov)
                    devuelve = DevuelveDesdeBDNew(cAgro, "rfactsoc", "numfactu", "codtipom", tipoMov, "T", , "numfactu", CStr(numfactu), "N", "fecfactu", FecFac, "F")
                    If devuelve <> "" Then
                        'Ya existe el contador incrementarlo
                        Existe = True
                        vTipoMov.IncrementarContador (tipoMov)
                        numfactu = vTipoMov.ConseguirContador(tipoMov)
                    Else
                        Existe = False
                    End If
                Loop Until Not Existe
                
                vParamAplic.PrimFactLiq = numfactu
                
            End If
        End If
    End If
    
    While Not Rs.EOF And B
        ActCalid = DBLet(Rs!codcalid, "N")
        ActVarie = DBLet(Rs!Codvarie, "N")
        actCampo = DBSet(Rs!codCampo, "N")
        ActSocio = DBSet(Rs!Codsocio, "N")
        
        If (ActCalid <> AntCalid Or AntCampo <> actCampo Or AntVarie <> ActVarie Or AntSocio <> ActSocio) Then
            ' kilos e importe por variedad campo
            Kilos = Kilos + KilosCal
            Importe = Importe + vImporte
            
            baseimpo = baseimpo + vImporte
            
            B = InsertLineaCalidad(tipoMov, CStr(numfactu), FecFac, AntVarie, AntCampo, CStr(AntCalid), CStr(DBLet(KilosCal, "N")), CStr(vImporte))
            KilosCal = 0
            vImporte = 0
            
            AntCalid = ActCalid
        End If
        
        If (ActVarie <> AntVarie Or actCampo <> AntCampo Or ActSocio <> AntSocio) Then
            If B Then ' descontamos el porcentaje de gastos de cooperativa
                GastosCoop = 0
                
                vPorcGasto = ""
                vPorcGasto = DevuelveDesdeBDNew(cAgro, "rcoope", "porcgast", "codcoope", vSocio.Cooperativa, "N")
                If vPorcGasto = "" Then vPorcGasto = "0"
                
                '[Monica]25/02/2011: Sólo hay gastos si no es complementaria ( Añadido el if )
                If TipoPrec <> 3 Then
                    GastosCoop = Round2(Importe * CCur(ImporteSinFormato(vPorcGasto)) / 100, 2)
                    Importe = Importe - GastosCoop
                    baseimpo = baseimpo - GastosCoop
                End If
            End If
            
            If B Then ' descontamos los gastos de los albaranes
'[MONICA] 08/09/2009 : los gastos de transporte se suman en ObtenerGastosAlbaranes, quito lo de David
'                '17 AGOSTO 2009
'                ' David###   Para VALSUR los gastos se suman
'                If vParamAplic.Cooperativa = 1 Then
'                    GastosAlb = ObtenerGastosAlbaranes(AntSocio, AntVarie, AntCampo, cTabla, cWhere)
'                    Importe = Importe + GastosAlb
'                    baseimpo = baseimpo + GastosAlb
'
'                Else
'                    'Para el resto sigue como estaba
                '[Monica]25/02/2011: Sólo hay gastos si no es complementaria ( Añadido el if )
                If TipoPrec <> 3 Then
                    GastosAlb = ObtenerGastosAlbaranes(AntSocio, AntVarie, AntCampo, cTabla, cWhere, 1)
                    Importe = Importe - GastosAlb
                    baseimpo = baseimpo - GastosAlb
                End If
            End If
            
            '[Monica]08/04/2010: grabamos los albaranes que intervienen en la linea de factura
            If B Then
                B = InsertarAlbaranesFactura(tipoMov, CStr(numfactu), FecFac, AntSocio, AntVarie, AntCampo, cTabla, cWhere, 0)
            End If
                        
            ' insertar linea de variedad, campo
            If B Then
                B = InsertLinea(tipoMov, CStr(numfactu), FecFac, AntVarie, AntCampo, CStr(Kilos), CStr(Importe), CStr(GastosAlb))
            End If
            
            If B Then
                ' tenemos que descontar los anticipos que tengamos para ello
                Sql2 = "select rfactsoc_variedad.imporvar, rfactsoc_variedad.numfactu, rfactsoc_variedad.fecfactu "
                Sql2 = Sql2 & " from rfactsoc_variedad INNER JOIN rfactsoc ON rfactsoc_variedad.codtipom = rfactsoc.codtipom and "
                Sql2 = Sql2 & " rfactsoc_variedad.numfactu = rfactsoc.numfactu and rfactsoc_variedad.fecfactu = rfactsoc.fecfactu "
                Sql2 = Sql2 & " where rfactsoc_variedad.codtipom = " & DBSet(vSocio.CodTipomAnt, "T") ' antes era 'FAA' "
                Sql2 = Sql2 & " and rfactsoc.codsocio = " & DBSet(AntSocio, "N")
                Sql2 = Sql2 & " and codvarie = " & DBSet(AntVarie, "N")
                Sql2 = Sql2 & " and codcampo = " & DBSet(AntCampo, "N")
                Sql2 = Sql2 & " and rfactsoc_variedad.descontado = 0"
                
                Set RS1 = New ADODB.Recordset
                RS1.Open Sql2, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
                
                While Not RS1.EOF
                    baseimpo = baseimpo - DBLet(RS1.Fields(0).Value, "N")
                    Anticipos = Anticipos + DBLet(RS1.Fields(0).Value, "N")
                    
                    ' indicamos que los anticipos ya han sido descontados
                    Sql3 = "update rfactsoc_variedad set descontado = 1 where codtipom = " & DBSet(vSocio.CodTipomAnt, "T") ' antes era 'FAC'
                    Sql3 = Sql3 & " and numfactu = " & DBSet(RS1!numfactu, "N")
                    Sql3 = Sql3 & " and fecfactu = " & DBSet(RS1!fecfactu, "F") & " and codvarie = " & DBSet(AntVarie, "N")
                    Sql3 = Sql3 & " and codcampo = " & DBSet(AntCampo, "N")
                    
                    conn.Execute Sql3
                    
                    ' insertamos en la tabla de anticipos de liquidacion venta campo
                    Sql3 = "insert into tmpFact_anticipos (codtipom, numfactu, fecfactu, codtipomanti, numfactuanti, fecfactuanti, codvarieanti, codcampoanti, baseimpo) values ("
                    Sql3 = Sql3 & DBSet(vSocio.CodTipomLiq, "T") & "," ' antes era 'FAL'
                    Sql3 = Sql3 & DBSet(numfactu, "N") & "," & DBSet(FecFac, "F") & ","
                    Sql3 = Sql3 & DBSet(vSocio.CodTipomAnt, "T") & "," ' antes era 'FAA'
                    Sql3 = Sql3 & DBSet(RS1!numfactu, "N") & "," & DBSet(RS1!fecfactu, "F") & ","
                    Sql3 = Sql3 & DBSet(AntVarie, "N") & "," & DBSet(AntCampo, "N") & "," & DBSet(RS1!imporvar, "N") & ")"
                    
                    conn.Execute Sql3
                    
                    RS1.MoveNext
                Wend
                
                Set RS1 = Nothing
                ' fin descontar anticipos
            
            End If
            
            
            If B Then
                AntVarie = ActVarie
                AntCampo = actCampo
                Kilos = 0
                Importe = 0
            End If
        End If
        
        If ActSocio <> AntSocio Then
            
            ImpoIva = Round2(baseimpo * PorcIva / 100, 2)
        
            Select Case DBLet(vSocio.TipoIRPF, "N")
                Case 0
                    ImpoReten = Round2((baseimpo + ImpoIva) * vParamAplic.PorcreteFacSoc / 100, 2)
                    BaseReten = (baseimpo + ImpoIva)
                    PorcReten = vParamAplic.PorcreteFacSoc
                Case 1
                    ImpoReten = Round2(baseimpo * vParamAplic.PorcreteFacSoc / 100, 2)
                    BaseReten = baseimpo
                    PorcReten = vParamAplic.PorcreteFacSoc
                Case 2
                    ImpoReten = 0
                    BaseReten = 0
                    PorcReten = 0
            End Select
            
            ImpoAFO = Round2((baseimpo + Anticipos) * vParamAplic.PorcenAFO / 100, 2)
            BaseAFO = baseimpo + Anticipos
            PorcAFO = vParamAplic.PorcenAFO
        
            TotalFac = baseimpo + ImpoIva - ImpoReten - ImpoAFO
            
            IncrementarProgresNew Pb1, 1
            
            '[Monica]07/02/2012: indicamos si es una factura de liquidacion complementaria
            'insertar cabecera de factura
            B = InsertCabecera(tipoMov, CStr(numfactu), FecFac, , , (TipoPrec = 3))
            
            '[Monica]04/01/2012: marcamos la factura como contabilizada y como pdte de recibir el nro de factura
            If B And vSocio.EmiteFact Then B = MarcarFactura(tipoMov, CStr(numfactu), FecFac, True)
            
            If B Then B = InsertResumen(tipoMov, CStr(numfactu))
            
'Mirar si quito lo de reclacular calidades
            If B Then B = RecalcularCalidades(tipoMov, CStr(numfactu), FecFac)
            
'Recalculo de todos los importes de tmpfact_calidades y tmpfact_variedades para que cuadre con la base de cabecera
            If B Then B = CuadrarBasesFactura(tipoMov, CStr(numfactu), FecFac, baseimpo)
            
            
            If B Then B = vTipoMov.IncrementarContador(tipoMov)
            
            If B Then
                AntSocio = ActSocio
                
                Set vSocio = Nothing
                Set vSocio = New cSocio
                If vSocio.LeerDatos(ActSocio) Then
                    If vSocio.LeerDatosSeccion(AntSocio, vParamAplic.Seccionhorto) Then
                        vPorcIva = ""
                        '[Monica]29/04/2011: INTERNAS
                        If vSocio.EsFactADVInt Then
                            vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vParamAplic.CodIvaExeADV, "N")
                        Else
                            vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vSocio.CodIva, "N")
                        End If
                        
                        PorcIva = CCur(ImporteSinFormato(vPorcIva))
                    End If
                    tipoMov = vSocio.CodTipomLiq
                                        
                End If
                baseimpo = 0
                BaseReten = 0
                ImpoIva = 0
                ImpoReten = 0
                TotalFac = 0
                BaseAFO = 0
                ImpoAFO = 0
                
                Anticipos = 0
                
                numfactu = vTipoMov.ConseguirContador(tipoMov)
                Do
                    numfactu = vTipoMov.ConseguirContador(tipoMov)
                    devuelve = DevuelveDesdeBDNew(cAgro, "rfactsoc", "numfactu", "codtipom", tipoMov, "T", , "numfactu", CStr(numfactu), "N", "fecfactu", FecFac, "F")
                    If devuelve <> "" Then
                        'Ya existe el contador incrementarlo
                        Existe = True
                        vTipoMov.IncrementarContador (tipoMov)
                        numfactu = vTipoMov.ConseguirContador(tipoMov)
                    Else
                        Existe = False
                    End If
                Loop Until Not Existe
           End If
        End If
        
        Recolect = DBLet(Rs!Recolect, "N")
        
        Select Case Recolect
'            Case 0
'                vPrecio = DBLet(RS!precoop, "N")
'                vImporte = vImporte + Round2(DBLet(RS!KilosNet, "N") * RS!precoop, 2)
'            Case 1
'                vPrecio = DBLet(RS!presocio, "N")
'                vImporte = vImporte + Round2(DBLet(RS!KilosNet, "N") * RS!presocio, 2)
            Case 0
                vPrecio = DBLet(Rs!PreCoop, "N")
                vImporte = vImporte + (DBLet(Rs!KilosNet, "N") * Rs!PreCoop)
            Case 1
                vPrecio = DBLet(Rs!PreSocio, "N")
                vImporte = vImporte + (DBLet(Rs!KilosNet, "N") * Rs!PreSocio)
        End Select
        
        KilosCal = KilosCal + DBLet(Rs!KilosNet, "N")
        
        HayReg = True
        
        Rs.MoveNext
    Wend
    
    ' ultimo registro si ha entrado
    If B And HayReg Then
        ' insertar linea de calidad
        Kilos = Kilos + KilosCal
        Importe = Importe + vImporte
        
        baseimpo = baseimpo + vImporte
        
        
        If B Then B = InsertLineaCalidad(tipoMov, CStr(numfactu), FecFac, ActVarie, actCampo, CStr(ActCalid), CStr(DBLet(KilosCal, "N")), CStr(vImporte))
        
        
        If B Then ' descontamos el porcentaje de gastos de cooperativa
            GastosCoop = 0
            
            vPorcGasto = ""
            vPorcGasto = DevuelveDesdeBDNew(cAgro, "rcoope", "porcgast", "codcoope", vSocio.Cooperativa, "N")
            If vPorcGasto = "" Then vPorcGasto = "0"
            
            '[Monica]25/02/2011: Sólo hay gastos si no es complementaria ( Añadido el if )
            If TipoPrec <> 3 Then
                GastosCoop = Round2(Importe * CCur(ImporteSinFormato(vPorcGasto)) / 100, 2)
                Importe = Importe - GastosCoop
                baseimpo = baseimpo - GastosCoop
            End If
        End If
        
        If B Then ' descontamos los gastos de los albaranes
'[MONICA] 08/09/2009 : los gastos de transporte se suman en ObtenerGastosAlbaranes, quito lo de David
'            '17 AGOSTO 2009
'            ' David###   Para VALSUR los gastos se suman
'            GastosAlb = ObtenerGastosAlbaranes(AntSocio, AntVarie, AntCampo, cTabla, cWhere)
'            If vParamAplic.Cooperativa = 1 Then
'                Importe = Importe + GastosAlb
'                baseimpo = baseimpo + GastosAlb
'            Else
'                Importe = Importe - GastosAlb
'                baseimpo = baseimpo - GastosAlb
'            End If
            
            '[Monica]25/02/2011: Sólo hay gastos si no es complementaria ( Añadido el if )
            If TipoPrec <> 3 Then
                GastosAlb = ObtenerGastosAlbaranes(AntSocio, AntVarie, AntCampo, cTabla, cWhere, 1)
                Importe = Importe - GastosAlb
                baseimpo = baseimpo - GastosAlb
            End If
        
        End If
        
        '[Monica]08/04/2010: grabamos los albaranes que intervienen en la linea de factura
        If B Then
            B = InsertarAlbaranesFactura(tipoMov, CStr(numfactu), FecFac, AntSocio, AntVarie, AntCampo, cTabla, cWhere, 0)
        End If
                    
                    
        ' insertar linea de variedad
        If B Then B = InsertLinea(tipoMov, CStr(numfactu), FecFac, CStr(ActVarie), CStr(actCampo), CStr(Kilos), CStr(Importe), CStr(GastosAlb))
        
        ' tenemos que descontar los anticipos que tengamos para ello
        If B Then
            Sql2 = "select rfactsoc_variedad.imporvar, rfactsoc_variedad.numfactu, rfactsoc_variedad.fecfactu "
            Sql2 = Sql2 & " from rfactsoc_variedad INNER JOIN rfactsoc ON rfactsoc_variedad.codtipom = rfactsoc.codtipom and "
            Sql2 = Sql2 & " rfactsoc_variedad.numfactu = rfactsoc.numfactu and rfactsoc_variedad.fecfactu = rfactsoc.fecfactu "
            Sql2 = Sql2 & " where rfactsoc_variedad.codtipom = " & DBSet(vSocio.CodTipomAnt, "T") ' antes era 'FAA' "
            Sql2 = Sql2 & " and rfactsoc.codsocio = " & DBSet(ActSocio, "N")
            Sql2 = Sql2 & " and codvarie = " & DBSet(ActVarie, "N")
            Sql2 = Sql2 & " and codcampo = " & DBSet(actCampo, "N")
            Sql2 = Sql2 & " and rfactsoc_variedad.descontado = 0"
            
            Set RS1 = New ADODB.Recordset
            RS1.Open Sql2, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
            
            While Not RS1.EOF
                baseimpo = baseimpo - DBLet(RS1.Fields(0).Value, "N")
                Anticipos = Anticipos + DBLet(RS1.Fields(0).Value, "N")
                
                ' indicamos que los anticipos ya han sido descontados
                Sql3 = "update rfactsoc_variedad set descontado = 1 where codtipom = " & DBSet(vSocio.CodTipomAnt, "T") ' antes era 'FAC'
                Sql3 = Sql3 & " and numfactu = " & DBSet(RS1!numfactu, "N")
                Sql3 = Sql3 & " and fecfactu = " & DBSet(RS1!fecfactu, "F") & " and codvarie = " & DBSet(ActVarie, "N")
                Sql3 = Sql3 & " and codcampo = " & DBSet(actCampo, "N")
                
                conn.Execute Sql3
                
                ' insertamos en la tabla de anticipos de liquidacion venta campo
                Sql3 = "insert into tmpFact_anticipos (codtipom, numfactu, fecfactu, codtipomanti, numfactuanti, fecfactuanti, codvarieanti, codcampoanti, baseimpo) values ("
                Sql3 = Sql3 & DBSet(vSocio.CodTipomLiq, "T") & "," ' antes era 'FAL'
                Sql3 = Sql3 & DBSet(numfactu, "N") & "," & DBSet(FecFac, "F") & ","
                Sql3 = Sql3 & DBSet(vSocio.CodTipomAnt, "T") & "," ' antes era 'FAA'
                Sql3 = Sql3 & DBSet(RS1!numfactu, "N") & "," & DBSet(RS1!fecfactu, "F") & ","
                Sql3 = Sql3 & DBSet(ActVarie, "N") & "," & DBSet(actCampo, "N") & "," & DBSet(RS1!imporvar, "N") & ")"
                
                conn.Execute Sql3
                
                RS1.MoveNext
            Wend
            
            Set RS1 = Nothing
            ' fin descontar anticipos
        
        End If
        
        
        ImpoIva = Round2(baseimpo * PorcIva / 100, 2)

        Select Case DBLet(vSocio.TipoIRPF, "N")
            Case 0
                ImpoReten = Round2((baseimpo + ImpoIva) * vParamAplic.PorcreteFacSoc / 100, 2)
                BaseReten = (baseimpo + ImpoIva)
                PorcReten = vParamAplic.PorcreteFacSoc
            Case 1
                ImpoReten = Round2(baseimpo * vParamAplic.PorcreteFacSoc / 100, 2)
                BaseReten = baseimpo
                PorcReten = vParamAplic.PorcreteFacSoc
            Case 2
                ImpoReten = 0
                BaseReten = 0
                PorcReten = 0
        End Select
        
        ImpoAFO = Round2((baseimpo + Anticipos) * vParamAplic.PorcenAFO / 100, 2)
        BaseAFO = baseimpo + Anticipos
        PorcAFO = vParamAplic.PorcenAFO

        TotalFac = baseimpo + ImpoIva - ImpoReten - ImpoAFO
        
        IncrementarProgresNew Pb1, 1
        
        vParamAplic.UltFactLiq = numfactu
        
        '[Monica]07/02/2012: indicamos si es una factura de liquidacion complementaria
        'insertar cabecera de factura
        B = InsertCabecera(tipoMov, CStr(numfactu), FecFac, , , (TipoPrec = 3))
        
        '[Monica]04/01/2012: marcamos la factura como contabilizada y como pdte de recibir el nro de factura
        If B And vSocio.EmiteFact Then B = MarcarFactura(tipoMov, CStr(numfactu), FecFac, True)
        
        If B Then B = InsertResumen(tipoMov, CStr(numfactu))

'Mirar si quito lo de reclacular calidades
        If B Then B = RecalcularCalidades(tipoMov, CStr(numfactu), FecFac)
        
'Recalculo de todos los importes de rfactsoc_calidades y rfactsoc_variedades para que cuadre con la base de cabecera
        If B Then B = CuadrarBasesFactura(tipoMov, CStr(numfactu), FecFac, baseimpo)
        
        If B Then B = vTipoMov.IncrementarContador(tipoMov)
        
        'pasamos las temporales a las tablas
        If B Then B = PasarTemporales()
        
        If B Then B = (vParamAplic.Modificar = 1)
    End If
    
'    BorrarTMPs
    
    vSeccion.CerrarConta
    Set vSeccion = Nothing
    Set vSocio = Nothing
    
eFacturacion:
    If Err.Number <> 0 Or Not B Then
        conn.RollbackTrans
        FacturacionLiquidacionesAlzira = False
    Else
        conn.CommitTrans
        FacturacionLiquidacionesAlzira = True
    End If
End Function



Public Function EsFacturaLiquidacion(CodTipom As String) As Boolean
Dim SQL As String

    If CodTipom = "" Then
        EsFacturaLiquidacion = False
        Exit Function
    End If

    SQL = ""
    SQL = DevuelveDesdeBDNew(cAgro, "usuarios.stipom", "tipodocu", "codtipom", CodTipom, "T")
    
    EsFacturaLiquidacion = (CInt(SQL) = 2 Or CInt(SQL) = 4 Or CInt(SQL) = 6)

End Function


Public Function ObtenerGastosAlbaranes(Socio As String, Varie As String, campo As String, cTabla As String, cWhere As String, Optional deTablaGastos As Byte, Optional deAlmazara As Byte, Optional Complementaria As Byte) As Currency
' deTablaGastos = 0 indica que cogemos unicamente los gastos que tenemos en rhisfruta
'               = 1 indica que cogemos los gastos que tenemos en la tabla rhisfruta_gastos
' deAlmazara = 0 indica que no viene de almazara :tenemos el campo
'            = 1 indica que viene de almazara: el codigo de campo es el minimo campo
' complementaria = 0 indica que no es complementaria
'                = 1 indica que es complementaria
Dim SQL As String
Dim RS1 As ADODB.Recordset

    On Error Resume Next
    
    ObtenerGastosAlbaranes = 0
    
    
    Select Case deTablaGastos
        Case 0
            ' 08/09/2009 : los gastos de tranporte para valsur son como una bonificacion luego se restan
            '              del resto de gastos ( aunque en ppio valsur no gasta impacarr, imprecol, imppenal
            '              Cambiado esto por lo de abajo
            '    SQL = "select sum(if(isnull(imptrans),0,imptrans)) + sum(if(isnull(impacarr),0,impacarr)) + sum(if(isnull(imprecol),0,imprecol)) + sum(if(isnull(imppenal),0,imppenal))  from rhisfruta "
                
            If Complementaria = 1 Then
            ' 15/03/2010
            '   SQL = "select sum(if(isnull(impacarr),0,impacarr)) + sum(if(isnull(imprecol),0,imprecol)) + sum(if(isnull(imppenal),0,imppenal)) - sum(if(isnull(imptrans),0,imptrans)) from rhisfruta "
            ' sustituido por esta otra donde no se dan las bonificaciones
                SQL = "select sum(if(isnull(impacarr),0,impacarr)) + sum(if(isnull(imprecol),0,imprecol)) + sum(if(isnull(imppenal),0,imppenal))  from rhisfruta "
            Else
                SQL = "select sum(if(isnull(impacarr),0,impacarr)) + sum(if(isnull(imprecol),0,imprecol)) + sum(if(isnull(imppenal),0,imppenal)) - sum(if(isnull(imptrans),0,imptrans)) from rhisfruta "
            End If
                SQL = SQL & " where numalbar in (select rhisfruta.numalbar from " & cTabla
                SQL = SQL & " where " & cWhere
                SQL = SQL & " and rhisfruta.codsocio = " & DBSet(Socio, "N")
                SQL = SQL & " and rhisfruta.codvarie = " & DBSet(Varie, "N")
                SQL = SQL & " and rhisfruta.codcampo = " & DBSet(campo, "N") & ")"
        Case 1
                SQL = "select sum(if(isnull(importe),0,importe)) from rhisfruta_gastos "
                SQL = SQL & " where numalbar in (select rhisfruta.numalbar from " & cTabla
                SQL = SQL & " where " & cWhere
                SQL = SQL & " and rhisfruta.codsocio = " & DBSet(Socio, "N")
                SQL = SQL & " and rhisfruta.codvarie = " & DBSet(Varie, "N")
                
                If deAlmazara = 0 Then
                    SQL = SQL & " and rhisfruta.codcampo = " & DBSet(campo, "N") & ")"
                Else
                    SQL = SQL & ")"
                End If
        
    End Select
    
    Set RS1 = New ADODB.Recordset
    RS1.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    If Not RS1.EOF Then ObtenerGastosAlbaranes = DBLet(RS1.Fields(0).Value, "N")

    Set RS1 = Nothing
    

End Function


Private Function RecalcularCalidades(TMov As String, Factu As String, FecFac As String) As Boolean
Dim SQL As String
Dim Sql2 As String
Dim Sql3 As String
Dim RS1 As ADODB.Recordset
Dim Rs2 As ADODB.Recordset
Dim ImpCalidad As Currency
Dim TotalVar As Currency
Dim TotCalidad As Currency
Dim UltCalid As Integer
Dim UltKilos As Currency
Dim Precio As Currency

    On Error GoTo eRecalcularCalidades

    RecalcularCalidades = False
    
    SQL = "select codvarie, codcampo, kilosnet, imporvar from tmpFact_variedad "
    SQL = SQL & " where codtipom = " & DBSet(TMov, "T")
    SQL = SQL & " and numfactu = " & DBSet(Factu, "N")
    SQL = SQL & " and fecfactu = " & DBSet(FecFac, "F")
    
    Set RS1 = New ADODB.Recordset
    RS1.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    While Not RS1.EOF
        Sql2 = "select codcalid, kilosnet from tmpFact_calidad "
        Sql2 = Sql2 & " where codtipom = " & DBSet(TMov, "T")
        Sql2 = Sql2 & " and numfactu = " & DBSet(Factu, "N")
        Sql2 = Sql2 & " and fecfactu = " & DBSet(FecFac, "F")
        Sql2 = Sql2 & " and codvarie = " & DBSet(RS1!Codvarie, "N")
        Sql2 = Sql2 & " and codcampo = " & DBSet(RS1!codCampo, "N")
        Sql2 = Sql2 & " and precio <> 0 "
        Sql2 = Sql2 & " order by codcalid "
    
        ' prorrateamos el importe de la variedad campo segun los kilos de la calidad
        Set Rs2 = New ADODB.Recordset
        Rs2.Open Sql2, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        
        TotCalidad = 0
        
        While Not Rs2.EOF
            ' imporvar - kilosvar
            '    x     - kiloscal
            ImpCalidad = 0
            If DBLet(RS1!KilosNet, "N") <> 0 Then
                ImpCalidad = Round2(DBLet(RS1!imporvar, "N") * DBLet(Rs2!KilosNet, "N") / DBLet(RS1!KilosNet, "N"), 2)
            End If
            TotCalidad = TotCalidad + ImpCalidad
            Precio = 0
            If DBLet(Rs2!KilosNet, "N") <> 0 Then
                Precio = Round2(ImpCalidad / DBLet(Rs2!KilosNet, "N"), 4)
            End If
            
            Sql3 = "update tmpFact_calidad set imporcal = " & DBSet(ImpCalidad, "N") & ","
            Sql3 = Sql3 & "precio = " & DBSet(Precio, "N")
            Sql3 = Sql3 & " where codtipom = " & DBSet(TMov, "T")
            Sql3 = Sql3 & " and numfactu = " & DBSet(Factu, "N")
            Sql3 = Sql3 & " and fecfactu = " & DBSet(FecFac, "F")
            Sql3 = Sql3 & " and codvarie = " & DBSet(RS1!Codvarie, "N")
            Sql3 = Sql3 & " and codcampo = " & DBSet(RS1!codCampo, "N")
            Sql3 = Sql3 & " and codcalid = " & DBSet(Rs2!codcalid, "N")
            
            conn.Execute Sql3
            
            UltCalid = Rs2!codcalid
            UltKilos = Rs2!KilosNet
            
            Rs2.MoveNext
        Wend
        
        ' en el ultimo registro aplicamos el redondeo
        If TotCalidad <> DBLet(RS1!imporvar, "N") Then
            ImpCalidad = DBLet(RS1!imporvar, "N") - TotCalidad
            Precio = Round2(ImpCalidad / UltKilos, 4)
            
            Sql3 = "update tmpFact_calidad set imporcal = imporcal + " & DBSet(ImpCalidad, "N") & ","
'            Sql3 = Sql3 & "precio = " & DBSet(Precio, "N")
            Sql3 = Sql3 & " precio = round(imporcal / " & DBSet(UltKilos, "N") & ", 4) "
            Sql3 = Sql3 & " where codtipom = " & DBSet(TMov, "T")
            Sql3 = Sql3 & " and numfactu = " & DBSet(Factu, "N")
            Sql3 = Sql3 & " and fecfactu = " & DBSet(FecFac, "F")
            Sql3 = Sql3 & " and codvarie = " & DBSet(RS1!Codvarie, "N")
            Sql3 = Sql3 & " and codcampo = " & DBSet(RS1!codCampo, "N")
            Sql3 = Sql3 & " and codcalid = " & DBSet(UltCalid, "N")
            
            conn.Execute Sql3
        
        End If
        
        Set Rs2 = Nothing
        RS1.MoveNext
    Wend
    
    Set RS1 = Nothing

    RecalcularCalidades = True
    
    Exit Function
    
eRecalcularCalidades:
    MuestraError Err.Number, "Recalculo de Calidades", Err.Description
End Function

'========================
Private Function CuadrarBasesFactura(TMov As String, Factu As String, FecFac As String, Base As Currency) As Boolean
Dim SQL As String
Dim Sql2 As String
Dim RS1 As ADODB.Recordset
Dim Rs2 As ADODB.Recordset
Dim ImpCalidad As Currency
Dim TotalVar As Currency
Dim TotCalidad As Currency
Dim UltCalid As Integer
Dim UltKilos As Currency
Dim Precio As Currency
Dim Diferencia As Currency
Dim Calidad As Currency

    On Error GoTo eCuadrarBasesFactura

    CuadrarBasesFactura = False
    
    SQL = "select sum(imporvar) from tmpFact_variedad  "
    SQL = SQL & " where codtipom = " & DBSet(TMov, "T")
    SQL = SQL & " and numfactu = " & DBSet(Factu, "N")
    SQL = SQL & " and fecfactu = " & DBSet(FecFac, "F")
    
    ' si la factura no cuadra
    If DevuelveValor(SQL) <> Round2(Base, 2) Then
        Diferencia = Round2(Base, 2) - DevuelveValor(SQL)
    
        SQL = "select codcampo, codvarie from tmpFact_variedad "
        SQL = SQL & " where codtipom = " & DBSet(TMov, "T")
        SQL = SQL & " and numfactu = " & DBSet(Factu, "N")
        SQL = SQL & " and fecfactu = " & DBSet(FecFac, "F")
            
        Set RS1 = New ADODB.Recordset
        RS1.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        
        If Not RS1.EOF Then
            SQL = "select imporvar from tmpFact_variedad  "
            SQL = SQL & " where codtipom = " & DBSet(TMov, "T")
            SQL = SQL & " and numfactu = " & DBSet(Factu, "N")
            SQL = SQL & " and fecfactu = " & DBSet(FecFac, "F")
            SQL = SQL & " and codvarie = " & DBSet(RS1!Codvarie, "N")
            SQL = SQL & " and codcampo = " & DBSet(RS1!codCampo, "N")
            TotCalidad = DevuelveValor(SQL) + Diferencia
        
        
            SQL = "update tmpFact_variedad  set imporvar = " & DBSet(TotCalidad, "N")
            SQL = SQL & " where codtipom = " & DBSet(TMov, "T")
            SQL = SQL & " and numfactu = " & DBSet(Factu, "N")
            SQL = SQL & " and fecfactu = " & DBSet(FecFac, "F")
            SQL = SQL & " and codvarie = " & DBSet(RS1!Codvarie, "N")
            SQL = SQL & " and codcampo = " & DBSet(RS1!codCampo, "N")
            
            conn.Execute SQL
            
            SQL = "select imporvar from tmpFact_variedad  "
            SQL = SQL & " where codtipom = " & DBSet(TMov, "T")
            SQL = SQL & " and numfactu = " & DBSet(Factu, "N")
            SQL = SQL & " and fecfactu = " & DBSet(FecFac, "F")
            SQL = SQL & " and codvarie = " & DBSet(RS1!Codvarie, "N")
            SQL = SQL & " and codcampo = " & DBSet(RS1!codCampo, "N")
            TotCalidad = DevuelveValor(SQL)
            
            SQL = "update tmpFact_variedad  set preciomed = round(imporvar / kilosnet,4) "
            SQL = SQL & " where codtipom = " & DBSet(TMov, "T")
            SQL = SQL & " and numfactu = " & DBSet(Factu, "N")
            SQL = SQL & " and fecfactu = " & DBSet(FecFac, "F")
            SQL = SQL & " and codvarie = " & DBSet(RS1!Codvarie, "N")
            SQL = SQL & " and codcampo = " & DBSet(RS1!codCampo, "N")
            
            conn.Execute SQL
            
' en las calidades de momento no hacemos nada
'            ' en las calidades metemos en la primera calidad la diferencia
'            SQL = "select min(codcalid) from tmpFact_calidad "
'            SQL = SQL & " where codtipom = " & DBSet(TMov, "T")
'            SQL = SQL & " and numfactu = " & DBSet(Factu, "N")
'            SQL = SQL & " and fecfactu = " & DBSet(FecFac, "F")
'            SQL = SQL & " and codvarie = " & DBSet(RS1!codvarie, "N")
'            SQL = SQL & " and codcampo = " & DBSet(RS1!CodCampo, "N")
'
'            Calidad = DevuelveValor(SQL)
'
'
'            SQL = "select imporcalidad from tmpFact_calidad "
'            SQL = SQL & " where codtipom = " & DBSet(TMov, "T")
'            SQL = SQL & " and numfactu = " & DBSet(Factu, "N")
'            SQL = SQL & " and fecfactu = " & DBSet(FecFac, "F")
'            SQL = SQL & " and codvarie = " & DBSet(RS1!codvarie, "N")
'            SQL = SQL & " and codcampo = " & DBSet(RS1!CodCampo, "N")
'            SQL = SQL & " and codcalid = " & DBSet(Calidad, "N")
'
'            TotCalidad = DevuelveValor(SQL) + Diferencia
'
'
'            SQL = "update tmpFact_calidad set imporcalidad = " & DBSet(TotCalidad, "N")
'            SQL = SQL & " where codtipom = " & DBSet(TMov, "T")
'            SQL = SQL & " and numfactu = " & DBSet(Factu, "N")
'            SQL = SQL & " and fecfactu = " & DBSet(FecFac, "F")
'            SQL = SQL & " and codvarie = " & DBSet(RS1!codvarie, "N")
'            SQL = SQL & " and codcampo = " & DBSet(RS1!CodCampo, "N")
'            SQL = SQL & " and codcalid = " & DBSet(Calidad, "N")
'
'            conn.Execute SQL
'
'            SQL = "update tmpFact_calidad set preciocalidad = round(imporcalidad / kilosnet,4) "
'            SQL = SQL & " where codtipom = " & DBSet(TMov, "T")
'            SQL = SQL & " and numfactu = " & DBSet(Factu, "N")
'            SQL = SQL & " and fecfactu = " & DBSet(FecFac, "F")
'            SQL = SQL & " and codvarie = " & DBSet(RS1!codvarie, "N")
'            SQL = SQL & " and codcampo = " & DBSet(RS1!CodCampo, "N")
'            SQL = SQL & " and codcalid = " & DBSet(Calidad, "N")
'
'            conn.Execute SQL
        End If
    
        Set RS1 = Nothing
    
    End If

    CuadrarBasesFactura = True
    
    Exit Function
    
eCuadrarBasesFactura:
    MuestraError Err.Number, "Cuadrar Bases de Factura", Err.Description
End Function

'=======================================================

Public Function ComprobarFechaVenci(FechaVenci As Date, Dia1 As Byte, Dia2 As Byte, Dia3 As Byte) As Date
Dim newFecha As Date
Dim B As Boolean

'=== Modificada Laura: 23/01/2007
    On Error GoTo ErrObtFec
    B = False
    
    '--- comprobar que tiene dias de pago para obtener nueva fecha
    If Not (Dia1 > 0 Or Dia2 > 0 Or Dia3 > 0) Then
        'si no tiene dias de pago la fecha es OK y fin
        ComprobarFechaVenci = FechaVenci
        Exit Function
    End If
        
    
    '--- Obtener nueva fecha del vencimiento
    newFecha = FechaVenci
    
    Do
        'si dia de la fecha vencimiento es uno de los 3 dias de pagos fecha es OK
        If Day(newFecha) = Dia1 Or Day(newFecha) = Dia2 Or Day(newFecha) = Dia3 Then
'            newFecha = CStr(newFecha)
            B = True
        Else
            'mientras esta en el mismo mes vamos aumentando dias hasta encontrar un dia de pago
            newFecha = DateAdd("d", 1, CDate(newFecha))
        End If
    Loop Until B = True Or Year(newFecha) = Year(FechaVenci) + 3
    
    ComprobarFechaVenci = newFecha
    Exit Function
    
ErrObtFec:
    MuestraError Err.Number, "Obtener Fecha vencimiento según dias de pago.", Err.Description
End Function


Public Function ComprobarMesNoGira(FecVenci As Date, MesNG As Byte, DiaVtoAt As Byte, Dia1 As Byte, Dia2 As Byte, Dia3 As Byte) As Date
Dim F As String

    If Month(FecVenci) = MesNG Then
        If DiaVtoAt > 0 Then
            F = DiaVtoAt & "/"
        Else
            F = Day(FecVenci) & "/"
        End If
        
        If Month(FecVenci) + 1 < 13 Then
            F = F & Month(FecVenci) + 1 & "/" & Year(FecVenci)
        Else
            F = F & "01/" & Year(FecVenci) + 1
        End If
        FecVenci = Format(F, "dd/mm/yyyy")
    End If
    ComprobarMesNoGira = FecVenci
End Function

Public Function FacturacionAnticiposGastos(cTabla As String, cWhere As String, FecFac As String, Pb1 As ProgressBar, cad As String) As Boolean
Dim SQL As String
Dim Sql3 As String

Dim Rs As ADODB.Recordset

Dim AntSocio As String
Dim AntVarie As String
Dim ActSocio As String
Dim ActVarie As String
Dim actCampo As String
Dim AntCampo As String
Dim ActCalid As String
Dim AntCalid As String

Dim HayReg As Boolean

Dim NumError As Long
Dim vImporte As Currency
Dim vPorcIva As String

Dim PrimFac As String
Dim UltFac As String

Dim tipoMov As String
Dim B As Boolean
Dim vSeccion As CSeccion
Dim Kilos As Currency
Dim KilosCal As Currency
Dim Importe As Currency

Dim devuelve As String
Dim Existe As Boolean

'Dim baseimpo As Currency
'Dim BaseReten As Currency
Dim Neto As Currency
'Dim ImpoIva As Currency
'Dim ImpoReten As Currency
'Dim TotalFac As Currency
Dim Recolect As Byte
Dim vPrecio As Currency
Dim ConGastos As Byte

    On Error GoTo eFacturacion

    FacturacionAnticiposGastos = False
    
    tipoMov = "FAA"
    
    BorrarTMPs
    B = CrearTMPs()
    If Not B Then
         Exit Function
    End If
    
    conn.BeginTrans
    
    SQL = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute SQL

    
    SQL = "SELECT rhisfruta.codsocio, rhisfruta.codvarie, rhisfruta.codcampo,  "
    SQL = SQL & "rhisfruta_clasif.codcalid, "
    SQL = SQL & "sum(rhisfruta_clasif.kilosnet) as kilosnet"
    SQL = SQL & " FROM  " & cTabla

    
    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        SQL = SQL & " WHERE " & cWhere
    End If
    SQL = SQL & " group by 1, 2, 3, 4 "
    SQL = SQL & " order by 1, 2, 3, 4 "
    
    
    Set vSeccion = New CSeccion
    
    If vSeccion.LeerDatos(vParamAplic.Seccionhorto) Then
        If Not vSeccion.AbrirConta Then
            Exit Function
        End If
    End If
    
    HayReg = False
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    If Not Rs.EOF Then
        AntSocio = CStr(DBLet(Rs!Codsocio, "N"))
        AntVarie = CStr(DBLet(Rs!Codvarie, "N"))
        AntCampo = CStr(DBLet(Rs!codCampo, "N"))
        AntCalid = CStr(DBLet(Rs!codcalid, "N"))
        
        ActSocio = CStr(DBLet(Rs!Codsocio, "N"))
        ActVarie = CStr(DBLet(Rs!Codvarie, "N"))
        actCampo = CStr(DBLet(Rs!codCampo, "N"))
        ActCalid = CStr(DBLet(Rs!codcalid, "N"))
    
        Set vSocio = New cSocio
        If vSocio.LeerDatos(ActSocio) Then
            If vSocio.LeerDatosSeccion(ActSocio, vParamAplic.Seccionhorto) Then
                baseimpo = 0
                BaseReten = 0
                ImpoIva = 0
                ImpoReten = 0
                TotalFac = 0
                BaseAFO = 0
                ImpoAFO = 0
                PorcAFO = 0
                
                Kilos = 0
                Importe = 0
                
                KilosCal = 0
                
                vPorcIva = ""
                vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vSocio.CodIva, "N")
                PorcIva = CCur(ImporteSinFormato(vPorcIva))
                
                tipoMov = vSocio.CodTipomAnt
                
                Set vTipoMov = New CTiposMov
                
                numfactu = vTipoMov.ConseguirContador(tipoMov)
                Do
                    numfactu = vTipoMov.ConseguirContador(tipoMov)
                    devuelve = DevuelveDesdeBDNew(cAgro, "rfactsoc", "numfactu", "codtipom", tipoMov, "T", , "numfactu", CStr(numfactu), "N", "fecfactu", FecFac, "F")
                    If devuelve <> "" Then
                        'Ya existe el contador incrementarlo
                        Existe = True
                        vTipoMov.IncrementarContador (tipoMov)
                        numfactu = vTipoMov.ConseguirContador(tipoMov)
                    Else
                        Existe = False
                    End If
                Loop Until Not Existe
                
                vParamAplic.PrimFactAnt = numfactu
                
            End If
        End If
    End If
    
    While Not Rs.EOF And B
        ActVarie = DBLet(Rs!Codvarie, "N")
        actCampo = DBSet(Rs!codCampo, "N")
        ActSocio = DBSet(Rs!Codsocio, "N")
        
        
        If (ActVarie <> AntVarie Or actCampo <> AntCampo Or ActSocio <> AntSocio) Then
            ' insertar linea de variedad, campo
            Sql3 = "select sum(imprecol) from rhisfruta where "
            If cad <> "" Then Sql3 = Sql3 & cad & " and "
            Sql3 = Sql3 & " rhisfruta.codvarie = " & DBSet(AntVarie, "N")
            Sql3 = Sql3 & " and codcampo = " & DBSet(AntCampo, "N") & " and codsocio = " & DBSet(AntSocio, "N")
            
            Importe = DevuelveValor(Sql3)
            
            
            B = InsertLinea(tipoMov, CStr(numfactu), FecFac, AntVarie, AntCampo, CStr(Kilos), CStr(Importe), "0")
            
            baseimpo = baseimpo + Importe
            
            If B Then
                AntVarie = ActVarie
                AntCampo = actCampo
                Kilos = 0
                Importe = 0
            End If
        End If
        
        If ActSocio <> AntSocio Then
            ImpoIva = Round2(baseimpo * PorcIva / 100, 2)
        
            Select Case DBLet(vSocio.TipoIRPF, "N")
                Case 0
                    ImpoReten = Round2((baseimpo + ImpoIva) * vParamAplic.PorcreteFacSoc / 100, 2)
                    BaseReten = (baseimpo + ImpoIva)
                    PorcReten = vParamAplic.PorcreteFacSoc
                Case 1
                    ImpoReten = Round2(baseimpo * vParamAplic.PorcreteFacSoc / 100, 2)
                    BaseReten = baseimpo
                    PorcReten = vParamAplic.PorcreteFacSoc
                Case 2
                    ImpoReten = 0
                    BaseReten = 0
                    PorcReten = 0
            End Select
        
            TotalFac = baseimpo + ImpoIva - ImpoReten - ImpoAFO
            
            IncrementarProgresNew Pb1, 1
            
            'insertar cabecera de factura
            B = InsertCabecera(tipoMov, CStr(numfactu), FecFac, True)
            
            If B Then B = InsertResumen(tipoMov, CStr(numfactu))
            
            If B Then B = vTipoMov.IncrementarContador(tipoMov)
            
            If B Then
                AntSocio = ActSocio
                
                Set vSocio = Nothing
                Set vSocio = New cSocio
                If vSocio.LeerDatos(ActSocio) Then
                    If vSocio.LeerDatosSeccion(AntSocio, vParamAplic.Seccionhorto) Then
                        vPorcIva = ""
                        vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vSocio.CodIva, "N")
                        PorcIva = CCur(ImporteSinFormato(vPorcIva))
                    End If
                    
                    tipoMov = vSocio.CodTipomAnt
                End If
                baseimpo = 0
                BaseReten = 0
                Neto = 0
                ImpoIva = 0
                ImpoReten = 0
                TotalFac = 0
                BaseAFO = 0
                ImpoAFO = 0
                
                numfactu = vTipoMov.ConseguirContador(tipoMov)
                Do
                    numfactu = vTipoMov.ConseguirContador(tipoMov)
                    devuelve = DevuelveDesdeBDNew(cAgro, "rfactsoc", "numfactu", "codtipom", tipoMov, "T", , "numfactu", CStr(numfactu), "N", "fecfactu", FecFac, "F")
                    If devuelve <> "" Then
                        'Ya existe el contador incrementarlo
                        Existe = True
                        vTipoMov.IncrementarContador (tipoMov)
                        numfactu = vTipoMov.ConseguirContador(tipoMov)
                    Else
                        Existe = False
                    End If
                Loop Until Not Existe
           End If
        End If
        
        ConGastos = DevuelveValor("select gastosrec from rcalidad where codvarie=" & DBSet(Rs!Codvarie, "N") & " and codcalid = " & DBSet(Rs!codcalid, "N"))
        
'        If DBLet(ConGastos, "N") = 1 Then
            KilosCal = DBLet(Rs!KilosNet, "N")
            Kilos = Kilos + KilosCal
'        Else
'            KilosCal = 0
'            KilosCal = DBLet(Rs!KilosNet, "N")
'        End If
        
        Importe = vImporte
        
        If KilosCal <> 0 Then
            B = InsertLineaCalidad(tipoMov, CStr(numfactu), FecFac, AntVarie, AntCampo, CStr(Rs!codcalid), CStr(DBLet(KilosCal, "N")), 0) ' CStr(vImporte))
        End If
        
        HayReg = True
        
        Rs.MoveNext
    Wend
    
    ' ultimo registro si ha entrado
    If B And HayReg Then
        ' insertar linea de variedad
        If B Then
            Sql3 = "select sum(imprecol) from rhisfruta where "
            If cad <> "" Then Sql3 = Sql3 & cad & " and "
            Sql3 = Sql3 & " rhisfruta.codvarie = " & DBSet(ActVarie, "N")
            Sql3 = Sql3 & " and codcampo = " & DBSet(actCampo, "N") & " and codsocio = " & DBSet(ActSocio, "N")
            
            Importe = DevuelveValor(Sql3)
            
            
            B = InsertLinea(tipoMov, CStr(numfactu), FecFac, ActVarie, actCampo, CStr(Kilos), CStr(Importe), "0")
            
            baseimpo = baseimpo + Importe
        End If
        
        ImpoIva = Round2(baseimpo * PorcIva / 100, 2)

        Select Case DBLet(vSocio.TipoIRPF, "N")
            Case 0
                ImpoReten = Round2((baseimpo + ImpoIva) * vParamAplic.PorcreteFacSoc / 100, 2)
                BaseReten = (baseimpo + ImpoIva)
                PorcReten = vParamAplic.PorcreteFacSoc
            Case 1
                ImpoReten = Round2(baseimpo * vParamAplic.PorcreteFacSoc / 100, 2)
                BaseReten = baseimpo
                PorcReten = vParamAplic.PorcreteFacSoc
            Case 2
                ImpoReten = 0
                BaseReten = 0
                PorcReten = 0
        End Select
        
'        BaseAFO = baseimpo
'        PorcAFO = vParamAplic.PorcenAFO
'        ImpoAFO = Round2(BaseAFO * PorcAFO / 100, 2)

        TotalFac = baseimpo + ImpoIva - ImpoReten - ImpoAFO
        
        IncrementarProgresNew Pb1, 1
        
        vParamAplic.UltFactAnt = numfactu
        
        'insertar cabecera de factura
        B = InsertCabecera(tipoMov, CStr(numfactu), FecFac, True)
        
        If B Then B = InsertResumen(tipoMov, CStr(numfactu))
        
        If B Then B = vTipoMov.IncrementarContador(tipoMov)
        
        
        If B Then B = ModificarCalidadesFacturasGastos()
        
        'pasamos las temporales a las tablas
        If B Then B = PasarTemporales()
        
        If B Then B = (vParamAplic.Modificar = 1)
    End If
    
'    BorrarTMPs
    
    vSeccion.CerrarConta
    Set vSeccion = Nothing
    Set vSocio = Nothing
    
eFacturacion:
    If Err.Number <> 0 Or Not B Then
        conn.RollbackTrans
        FacturacionAnticiposGastos = False
    Else
        conn.CommitTrans
        FacturacionAnticiposGastos = True
    End If
End Function



Private Function ModificarCalidadesFacturasGastos() As Boolean
Dim SQL As String
Dim Sql2 As String
Dim Sql3 As String
Dim Rs As ADODB.Recordset
Dim Rs2 As ADODB.Recordset
Dim TotalKilos As Long
Dim ImporteTotal As Currency
Dim Importe As Currency
Dim Precio As Currency
Dim Diferencia As Currency
Dim AntCodcalid As Currency
Dim AntKilosNet As Currency


    On Error GoTo eModificarCalidadesFacturasGastos


    ModificarCalidadesFacturasGastos = False
    
    
    SQL = "select * from tmpfact_variedad order by 1,2,3"
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    While Not Rs.EOF
        Sql2 = "select sum(kilosnet) from tmpfact_calidad where codtipom = " & DBSet(Rs!CodTipom, "T")
        Sql2 = Sql2 & " and numfactu = " & DBSet(Rs!numfactu, "N")
        Sql2 = Sql2 & " and fecfactu = " & DBSet(Rs!fecfactu, "F")
        Sql2 = Sql2 & " and codvarie = " & DBSet(Rs!Codvarie, "N")
        Sql2 = Sql2 & " and codcampo = " & DBSet(Rs!codCampo, "N")
        ' solo esto
        Sql2 = Sql2 & " and codcalid in (select codcalid from rcalidad where codvarie = " & DBSet(Rs!Codvarie, "N") & " and gastosrec = 1)"
        
        TotalKilos = DevuelveValor(Sql2)
    
        Sql2 = "select * from tmpfact_calidad where codtipom = " & DBSet(Rs!CodTipom, "T")
        Sql2 = Sql2 & " and numfactu = " & DBSet(Rs!numfactu, "N")
        Sql2 = Sql2 & " and fecfactu = " & DBSet(Rs!fecfactu, "F")
        Sql2 = Sql2 & " and codvarie = " & DBSet(Rs!Codvarie, "N")
        Sql2 = Sql2 & " and codcampo = " & DBSet(Rs!codCampo, "N")
        ' solo esto
        Sql2 = Sql2 & " and codcalid in (select codcalid from rcalidad where codvarie = " & DBSet(Rs!Codvarie, "N") & " and gastosrec = 1)"
        
        Set Rs2 = New ADODB.Recordset
        Rs2.Open Sql2, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
        ImporteTotal = 0
    
        While Not Rs2.EOF
            Importe = Round2(DBLet(Rs2!KilosNet, "N") * DBLet(Rs!imporvar, "N") / TotalKilos, 2)
            
            Precio = 0
            If DBLet(Rs2!KilosNet, "N") <> 0 Then
                Precio = Round2(Importe / Rs2!KilosNet, 4)
            End If
        
            ImporteTotal = ImporteTotal + Importe
            
            Sql3 = "update tmpfact_calidad set imporcal = " & DBSet(Importe, "N")
            Sql3 = Sql3 & ", precio = " & DBSet(Precio, "N")
            Sql3 = Sql3 & " where codtipom = " & DBSet(Rs!CodTipom, "T")
            Sql3 = Sql3 & " and numfactu = " & DBSet(Rs!numfactu, "N")
            Sql3 = Sql3 & " and fecfactu = " & DBSet(Rs!fecfactu, "F")
            Sql3 = Sql3 & " and codvarie = " & DBSet(Rs!Codvarie, "N")
            Sql3 = Sql3 & " and codcampo = " & DBSet(Rs!codCampo, "N")
            Sql3 = Sql3 & " and codcalid = " & DBSet(Rs2!codcalid, "N")
            
            conn.Execute Sql3
            
            AntCodcalid = Rs2!codcalid
            AntKilosNet = Rs2!KilosNet
            
            Rs2.MoveNext
        Wend
        
        Diferencia = DBLet(Rs!imporvar, "N") - ImporteTotal
        
        If Diferencia <> 0 Then
            'actualizamos el ultimo registro
            ' importe
            Sql3 = "update tmpfact_calidad set imporcal = imporcal + " & DBSet(Diferencia, "N")
            Sql3 = Sql3 & " where codtipom = " & DBSet(Rs!CodTipom, "T")
            Sql3 = Sql3 & " and numfactu = " & DBSet(Rs!numfactu, "N")
            Sql3 = Sql3 & " and fecfactu = " & DBSet(Rs!fecfactu, "F")
            Sql3 = Sql3 & " and codvarie = " & DBSet(Rs!Codvarie, "N")
            Sql3 = Sql3 & " and codcampo = " & DBSet(Rs!codCampo, "N")
            Sql3 = Sql3 & " and codcalid = " & DBSet(AntCodcalid, "N")
            
            conn.Execute Sql3
        
            ' precio
            Sql3 = "update tmpfact_calidad set precio = round(imporcal / kilosnet,4) "
            Sql3 = Sql3 & " where codtipom = " & DBSet(Rs!CodTipom, "T")
            Sql3 = Sql3 & " and numfactu = " & DBSet(Rs!numfactu, "N")
            Sql3 = Sql3 & " and fecfactu = " & DBSet(Rs!fecfactu, "F")
            Sql3 = Sql3 & " and codvarie = " & DBSet(Rs!Codvarie, "N")
            Sql3 = Sql3 & " and codcampo = " & DBSet(Rs!codCampo, "N")
            Sql3 = Sql3 & " and codcalid = " & DBSet(AntCodcalid, "N")
            
            conn.Execute Sql3
        
        
        End If
        
        Set Rs2 = Nothing
        
        Rs.MoveNext
    Wend
    
    Set Rs = Nothing
    
    ModificarCalidadesFacturasGastos = True
    Exit Function
    
eModificarCalidadesFacturasGastos:
    MuestraError Err.Number, "Modificar Calidades Facturas Gastos", Err.Description
End Function




Public Function TraspasoPartesFacturas(cadSQL As String, cadWHERE As String, FechaFact As String, Banpr As String, ByRef PBar1 As ProgressBar, ByRef LblBar As Label, ImprimeLasFacturasGeneradas As Boolean, ByRef vTipoM As String, TextosCSB As String, Forpa As String) As Boolean
'IN -> cadSQL: cadena para seleccion de los Partes que vamos a Facturar
'      FechaFact: Fecha de la Factura
'      BanPr: Cod. de Banco Propio
'      Pbar1:  Una progressbar. Se puede mandar un NOTHING, y no pasa nada. Si no se manda
'              es que estamos en un proceso corto o que no necesitabaos un pb1, con lo cual NO muestro el PB1
'      Imprime: Si despues de generarlo los imprime
'
'       vTipom:  Que tipo de albaran es, para luego la impresion saber que factura imprime
'      TextosCSB:  Si lleva llevara 3 lineas para meter ent tesoreria

'Desde Albaranes Genera las Facturas correspondientes
Dim RsAlb As ADODB.Recordset 'Ordenados por: tipofac,clien,dpto,forma pago, dtoppago, dtognral
Dim B As Boolean
Dim SQL As String

'Aqui Guardamos los datos del Albaran Anterior para comparar con el actual
Dim AntSocio As Long
Dim antDirec As Long
Dim antForpa As Byte
Dim antDtoPP As Single, antDtoGn As Single

'direc/dpto actual para controlar el valor nulo
Dim actDirec As Long

'Concatenamos todas las facturas generadas para listarlas en el informe
Dim ListFactu As String
Dim TipFactu As String

Dim vFactuADV As CFacturaADV
Dim INC As Integer
Dim Condicion As Boolean 'condicion que comprueba para romper la agrupacion de albaranes a 1 factura

'Por si no mando una progressbar, que no de errores
Dim PgbVisible As Boolean

    On Error GoTo ETraspasoAlbFac

    TraspasoPartesFacturas = False

    ListFactu = ""
    TipFactu = ""
        
    'comprobamos que no haya nadie facturando
    DesBloqueoManual ("ADVFAC") 'facturas de adv
    If Not BloqueoManual("ADVFAC", "1") Then
        MsgBox "No se puede facturar. Hay otro usuario facturando.", vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If
    
    'Bloqueamos todos los albaranes que vamos a facturar (cabeceras y lineas)
    'Nota: esta bloqueando tambien los registros de la tabla clientes: sclien correspondientes
    SQL = " (advpartes INNER JOIN rsocios ON advpartes.codsocio=rsocios.codsocio ) INNER JOIN advpartes_lineas ON advpartes.numparte=advpartes_lineas.numparte "
    If Not BloqueaRegistro(SQL, cadWHERE) Then
        Screen.MousePointer = vbDefault
        'comprobamos que no haya nadie facturando
        DesBloqueoManual ("ADVFAC")
        Exit Function
    End If
    
   
    'Inicializar la Progress Bar
    PgbVisible = False
    If Not (PBar1 Is Nothing) Then
        If PBar1.visible Then PgbVisible = True
    End If
    If PgbVisible Then
        If InStr(1, cadSQL, "rsocios") Then
'            Sql = Replace(cadSQL, "scaalb.*, clientes.periodof", "count(*)") 'si hay INNER JOIN con clientes
            SQL = Replace(cadSQL, "*", "count(*)") 'si hay INNER JOIN con sclien
        Else
            SQL = Replace(cadSQL, "*", "count(*)") 'si NO hay INNER JOIN con sclien
        End If
        
        
        Set RsAlb = New ADODB.Recordset
        RsAlb.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not RsAlb.EOF Then
            CargarProgresNew PBar1, CInt(RsAlb.Fields(0))
            LblBar.Caption = "Inicializando el proceso..."
        End If
        RsAlb.Close
        Set RsAlb = Nothing
    End If
    
        
    'EMPEZAMOS LA FACTURA
    Set vFactuADV = New CFacturaADV
    vFactuADV.fecfactu = FechaFact 'Fecha para las Facturas

    'Marcar Partes que se van a Facturar
    '----------------------------------------
    SQL = cadSQL & " ORDER BY advpartes.codsocio "
    Set RsAlb = New ADODB.Recordset
    RsAlb.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        
    'Agrupar los Partes posibles en una misma Factura
    'Calcular y Grabar Factura en la Tabla de Facturas
    'Partes(advpartes, advpartes_lineas) -> Factura (facturas, facturas_envases)
    '----------------------------------------------------
    'Agrupar albaranes en 1 factura por : codclien,codforpa
    B = True
    
    AntSocio = 0 'socio
    
    cadW = ""
    Errores = ""
    INC = 0
    
    While Not RsAlb.EOF
        TipoAlb = "PAR"
        INC = INC + 1
             
        '[Monica]18/05/2012
        If vParamAplic.Cooperativa = 3 Then
            LblBar.Caption = "Facturando: Albaranes Venta"
        Else
            LblBar.Caption = "Facturando: Partes ADV"
        End If
         Condicion = (AntSocio <> RsAlb!Codsocio)
         
'             If (antClien <> RSalb!CodClien) Or (antDirec <> actDirec) Or (antForpa <> RSalb!codforpa) Or (antDtoPP <> RSalb!DtoPPago) Or (antDtoGn <> RSalb!DtoGnral) Then
         If Condicion Then
         '-----
            If cadW <> "" Then 'Facturacion PEndiente
                cadW = cadW & ") "
                
                If Not vFactuADV.PasarPartesAFactura2(TipoAlb, cadW, TextosCSB, Forpa, ErroresAux, False) Then
                    If B Then B = False
                    AnyadirAvisos ErroresAux
                Else 'añadirlo a la lista de facturas a imprimir
                    If ListFactu = "" Then
                        ListFactu = vFactuADV.numfactu
                    Else
                        ListFactu = ListFactu & "," & vFactuADV.numfactu
                    End If
                    If TipFactu = "" Then
                        TipFactu = "'" & vFactuADV.CodTipom & "'"
                    Else
                        TipFactu = TipFactu & ",'" & vFactuADV.CodTipom & "'"
                    End If
                End If
                If PgbVisible Then
                    LblBar.Caption = "Socio: " & Format(vFactuADV.Socio, "000000") & " " & vFactuADV.NombreSocio
                    IncrementarProgresNew PBar1, INC
                    INC = 0
                End If
                espera 0.2
                
                'Empezamos una nueva Factura
                cadW = ""
            End If
            'Generar una Factura nueva
            vFactuADV.Socio = RsAlb!Codsocio
            vFactuADV.NombreSocio = RsAlb!nomsocio
            vFactuADV.DomicilioSocio = DBLet(RsAlb!dirsocio, "T")
            vFactuADV.CPostal = DBLet(RsAlb!codpostal, "T")
            vFactuADV.Poblacion = DBLet(RsAlb!pobsocio, "T")
            vFactuADV.Provincia = DBLet(RsAlb!prosocio, "T")
            vFactuADV.nif = DBLet(RsAlb!nifSocio, "T")
            vFactuADV.Telefono = DBLet(RsAlb!telsoci1, "T")
            vFactuADV.ForPago = Forpa
'[Monica] 09/02/2010 la forma de pago está en la contabilidad de adv
'            vFactuADV.TipForPago = DevuelveDesdeBDNew(cAgro, "forpago", "tipoforp", "codforpa", Forpa, "N")
            If vParamAplic.ContabilidadNueva Then
                vFactuADV.TipForPago = DevuelveDesdeBDNew(cConta, "formapago", "tipforpa", "codforpa", Forpa, "N")
            Else
                vFactuADV.TipForPago = DevuelveDesdeBDNew(cConta, "sforpa", "tipforpa", "codforpa", Forpa, "N")
            End If
            cadW = "  advpartes.numparte IN (" & RsAlb!Numparte
        Else
            cadW = cadW & ", " & RsAlb!Numparte
        End If
    
        'Guardamos datos del registro anterior
        AntSocio = RsAlb!Codsocio
        RsAlb.MoveNext
    Wend
    RsAlb.Close
    Set RsAlb = Nothing
        
    'Facturar la ultima Factura generada del blucle
    If cadW <> "" Then
        cadW = cadW & ")"
        If PgbVisible Then LblBar.Caption = "Socio: " & Format(vFactuADV.Socio, "000000") & " - " & vFactuADV.NombreSocio
        
        If Not vFactuADV.PasarPartesAFactura2(TipoAlb, cadW, TextosCSB, Forpa, ErroresAux, False) Then
            If B Then B = False
            AnyadirAvisos "Error Facturando el Socio: " & Format(vFactuADV.Socio, "000000") & " " & vFactuADV.NombreSocio & vbCrLf & ErroresAux
        Else 'añadirlo a la lista de facturas a imprimir
            If ListFactu = "" Then
                ListFactu = vFactuADV.numfactu
            Else
                ListFactu = ListFactu & "," & vFactuADV.numfactu
            End If
        
            If TipFactu = "" Then
                TipFactu = "'" & vFactuADV.CodTipom & "'"
            Else
                TipFactu = TipFactu & ",'" & vFactuADV.CodTipom & "'"
            End If
        
        End If
        If PgbVisible Then
'            LblBar.Caption = "Cliente: " & Format(vFactu.Cliente, "000000") & " - " & vFactu.NombreClien
            IncrementarProgresNew PBar1, INC
        End If
        espera 0.2
    End If
    
    TipoFac = vFactuADV.CodTipom
    Set vFactuADV = Nothing
    TraspasoPartesFacturas = True
    
    If B Then
        LblBar.Caption = "Proceso finalizado correctamente."
        '[Monica]18/05/2012
        If vParamAplic.Cooperativa = 3 Then
            MsgBox "Las Facturas de los Albaranes seleccionados se generaron correctamente.", vbInformation
        Else
            MsgBox "Las Facturas de los Partes seleccionados se generaron correctamente.", vbInformation
        End If
    Else
        LblBar.Caption = "Proceso finalizado con errores."
        SQL = "ATENCIÓN:" & vbCrLf
        MsgBox SQL & "No todas las Facturas se generaron correctamente!!!.", vbExclamation
        If Errores <> "" Then MostrarAvisos
    End If
    
    espera 0.2
    
    'Desbloqueamos ya no estamos facturando
    DesBloqueoManual ("ADVFAC")
    TerminaBloquear
    
    
    If ImprimeLasFacturasGeneradas Then
        If ListFactu <> "" Then
            ImprimirFacturas ListFactu, FechaFact, , False, TipFactu
        End If
    End If
    
ETraspasoAlbFac:
    If Err.Number <> 0 Then
        If vParamAplic.Cooperativa = 3 Then
            MuestraError Err.Number, "Facturando Albaranes", Err.Description
        Else
            MuestraError Err.Number, "Facturando Partes", Err.Description
        End If
        'comprobamos que no haya nadie facturando
        DesBloqueoManual ("ADVFAC")
        TerminaBloquear
    End If
End Function




Private Sub AnyadirAvisos(Donde As String)
    Errores = Errores & vbCrLf & vbCrLf & Donde & vbCrLf
End Sub



Private Sub MostrarAvisos()
    frmMensajes.vCampos = Errores
    frmMensajes.OpcionMensaje = 13
    frmMensajes.Show vbModal
End Sub



Public Sub ImprimirFacturas(listaF As String, fechaF As String, Optional SQL As String, Optional FormatoFacturaTPV As Boolean, Optional TFactu As String)
Dim cadFormula As String
Dim CadParam As String
Dim numParam As Byte
Dim cadSelect As String 'select para insertar en tabla temporal
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String
Dim NombreTabla As String

    cadFormula = ""
    CadParam = ""
    cadSelect = ""
    numParam = 0
    NombreTabla = "advfacturas"

    '===================================================
    '============ PARAMETROS ===========================
    indRPT = 32 'Facturas ADV
    
    If Not PonerParamRPT(indRPT, CadParam, numParam, nomDocu) Then
        Exit Sub
    End If

    'Nombre fichero .rpt a Imprimir
    frmImprimir.NombreRPT = nomDocu


    If SQL <> "" Then
        'Llamo desde el menu de Reimprimir facturas y tengo construida la
        'cadena de seleccion D/H tipoMov, D/H NumFactu, D/H fecfactu
        cadSelect = SQL
        cadFormula = listaF
        CadParam = CadParam & fechaF
        numParam = numParam + 1
    Else
        'Llama desde PasarAlbaranes a  Facturas y al terminar las imprime
        '===================================================
        '================= FORMULA =========================
        'Cadena para seleccion Nº de Factura
        '---------------------------------------------------
        'Cod Tipo Movimiento
        '[Monica]21/03/2011: puede que haya mas de un tipo de movimiento 'FAP FIN en facturas de adv
        If TFactu = "" Then
            devuelve = "({" & NombreTabla & ".codtipom}='" & TipoFac & "') "
            If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
        Else
            devuelve = "({" & NombreTabla & ".codtipom} IN [" & TFactu & "])"
            If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
        End If
        
        'Nº Factura
        devuelve = "({" & NombreTabla & ".numfactu} IN [" & listaF & "])"
        If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
    
        'fecha factu
        devuelve = "(year({" & NombreTabla & ".fecfactu}) = " & Year(fechaF) & ")"
        If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub

        cadSelect = cadFormula

        cadSelect = Replace(cadSelect, "[", "(")
        cadSelect = Replace(cadSelect, "]", ")")
    End If
    
    If Not HayRegParaInforme(NombreTabla, cadSelect) Then Exit Sub

     With frmImprimir
            .FormulaSeleccion = cadFormula
            .OtrosParametros = CadParam
            .NumeroParametros = numParam
            .SoloImprimir = False
            .EnvioEMail = False
            .Opcion = 0
            .Titulo = "Factura ADV Socios"
            .ConSubInforme = False
            .Show vbModal
    End With
    If frmVisReport.EstaImpreso Then
         ActualizarRegistrosFac "advfacturas", cadSelect
    End If

End Sub


Public Sub ImprimirFacturasBOD(listaF As String, fechaF As String, Optional SQL As String, Optional FormatoFacturaTPV As Boolean)
Dim cadFormula As String
Dim CadParam As String
Dim numParam As Byte
Dim cadSelect As String 'select para insertar en tabla temporal
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String
Dim NombreTabla As String

    cadFormula = ""
    CadParam = ""
    cadSelect = ""
    numParam = 0
    NombreTabla = "rbodfacturas"

    '===================================================
    '============ PARAMETROS ===========================
    indRPT = 41 'Facturas BOD
    
    If Not PonerParamRPT(indRPT, CadParam, numParam, nomDocu) Then
        Exit Sub
    End If

    'Nombre fichero .rpt a Imprimir
    frmImprimir.NombreRPT = nomDocu


    If SQL <> "" Then
        'Llamo desde el menu de Reimprimir facturas y tengo construida la
        'cadena de seleccion D/H tipoMov, D/H NumFactu, D/H fecfactu
        cadSelect = SQL
        cadFormula = listaF
        CadParam = CadParam & fechaF
        numParam = numParam + 1
    Else
        'Llama desde PasarAlbaranes a  Facturas y al terminar las imprime
        '===================================================
        '================= FORMULA =========================
        'Cadena para seleccion Nº de Factura
        '---------------------------------------------------
        'Cod Tipo Movimiento
        devuelve = "({" & NombreTabla & ".codtipom}='" & TipoFac & "') "
        If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
    
        'Nº Factura
        devuelve = "({" & NombreTabla & ".numfactu} IN [" & listaF & "])"
        If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
    
        'fecha factu
        devuelve = "(year({" & NombreTabla & ".fecfactu}) = " & Year(fechaF) & ")"
        If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub

        cadSelect = cadFormula

        cadSelect = Replace(cadSelect, "[", "(")
        cadSelect = Replace(cadSelect, "]", ")")
    End If
    
    If Not HayRegParaInforme(NombreTabla, cadSelect) Then Exit Sub

     With frmImprimir
            .FormulaSeleccion = cadFormula
            .OtrosParametros = CadParam
            .NumeroParametros = numParam
            .SoloImprimir = False
            .EnvioEMail = False
            .Opcion = 0
            .Titulo = "Factura Retirada Socios"
            .ConSubInforme = False
            .Show vbModal
    End With
    If frmVisReport.EstaImpreso Then
         ActualizarRegistrosFac "rbodfacturas", cadSelect
    End If

End Sub





Public Function FacturacionLiquidacionesCatadau(cTabla As String, cWhere As String, FecFac As String, Pb1 As ProgressBar, EsComplemen As Boolean, FecDesde As String, FecHasta As String, vFechas As String, NoPermitirFactNegativas As Boolean, EsTercero As Boolean) As Boolean
Dim SQL As String
Dim Rs As ADODB.Recordset
Dim RS1 As ADODB.Recordset

Dim AntSocio As String
Dim AntVarie As String
Dim ActSocio As String
Dim ActVarie As String
Dim actCampo As String
Dim AntCampo As String
Dim ActCalid As String
Dim AntCalid As String
Dim AntFecIni As String
Dim ActFecIni As String
Dim AntFecFin As String
Dim ActFecFin As String

Dim HayReg As Boolean

Dim NumError As Long
Dim vImporte As Currency
Dim vPorcIva As String

Dim PrimFac As String
Dim UltFac As String

Dim tipoMov As String
Dim B As Boolean
Dim vSeccion As CSeccion
Dim Kilos As Currency
Dim KilosCal As Currency
Dim Importe As Currency

Dim devuelve As String
Dim Existe As Boolean

'Dim baseimpo As Currency
'Dim BaseReten As Currency
'Dim Neto As Currency
'Dim ImpoIva As Currency
'Dim ImpoReten As Currency
'Dim TotalFac As Currency
Dim Recolect As Byte
Dim vPrecio As Currency

Dim Sql2 As String
Dim Sql3 As String

Dim Gastos As Currency
Dim vPorcGasto As String


    On Error GoTo eFacturacion

    FacturacionLiquidacionesCatadau = False
    
'08052009 antes dentro de transacciones
    BorrarTMPs
    B = CrearTMPs()
    If Not B Then
         Exit Function
    End If
'08052009
    
    conn.BeginTrans
    
    '10/05/2013: antes de borrar cargo la tabla auxiliar de albaranes, voy a utilizar la tmpexecel
    SQL = "delete from tmpexcel where codusu = " & vUsu.Codigo
    conn.Execute SQL
    
    SQL = "insert into tmpexcel (codusu, numalbar, fecalbar, codvarie, codsocio, kilosnet) select distinct " & vUsu.Codigo & ","
    SQL = SQL & " importe1, fecha1, importe2, rhisfruta.codsocio, sum(importe4)  from tmpinformes inner join rhisfruta on tmpinformes.importe1 = rhisfruta.numalbar "
    SQL = SQL & " where codusu = " & vUsu.Codigo
    SQL = SQL & " group by 1,2,3,4,5 "
    conn.Execute SQL
    'hasta aqui 10/05/2013
    
    '[Monica]01/10/2018: solo en el caso de castelduc quiere distinto contador
    If EsTercero Then
        tipoMov = "FLT"
    Else
        tipoMov = "FAL"
    End If
    
    SQL = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute SQL

    SQL = "SELECT  tmpliquidacion.codsocio, tmpliquidacion.codvarie,"
    SQL = SQL & "tmpliquidacion.codcampo, tmpliquidacion.codcalid, "
    SQL = SQL & "sum(tmpliquidacion.kilosnet) as kilosnet, sum(tmpliquidacion.importe) as importe "
    SQL = SQL & " FROM  tmpliquidacion "
    SQL = SQL & " where codusu = " & DBSet(vUsu.Codigo, "N")

    ' ordenado por socio, variedad, campo, calidad
    SQL = SQL & " group by tmpliquidacion.codsocio, tmpliquidacion.codvarie, tmpliquidacion.codcampo, tmpliquidacion.codcalid "
    SQL = SQL & " order by tmpliquidacion.codsocio, tmpliquidacion.codvarie, tmpliquidacion.codcampo, tmpliquidacion.codcalid "
    
    Set vSeccion = New CSeccion
    
    If vSeccion.LeerDatos(vParamAplic.Seccionhorto) Then
        If Not vSeccion.AbrirConta Then
            Exit Function
        End If
    End If
    
    HayReg = False
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    If Not Rs.EOF Then
        AntSocio = CStr(DBLet(Rs!Codsocio, "N"))
        AntVarie = CStr(DBLet(Rs!Codvarie, "N"))
        AntCampo = CStr(DBLet(Rs!codCampo, "N"))
        AntCalid = CStr(DBLet(Rs!codcalid, "N"))
        
        ActSocio = CStr(DBLet(Rs!Codsocio, "N"))
        ActVarie = CStr(DBLet(Rs!Codvarie, "N"))
        actCampo = CStr(DBLet(Rs!codCampo, "N"))
        ActCalid = CStr(DBLet(Rs!codcalid, "N"))
    
        Set vSocio = New cSocio
        If vSocio.LeerDatos(ActSocio) Then
            If vSocio.LeerDatosSeccion(ActSocio, vParamAplic.Seccionhorto) Then
                baseimpo = 0
                BaseReten = 0
                ImpoIva = 0
                ImpoReten = 0
                TotalFac = 0
                ImpoAFO = 0
                BaseAFO = 0
                
                Anticipos = 0
                
                Kilos = 0
                Importe = 0
                
                KilosCal = 0
                
                '[Monica]05/03/2014: alzira entra a tramos
                If vParamAplic.Cooperativa = 4 Then
                    '[Monica]29/04/2011: INTERNAS
                    If vSocio.EsFactADVInt Then
                        vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vParamAplic.CodIvaExeADV, "N")
                    Else
                        vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vSocio.CodIva, "N")
                    End If
                    PorcIva = CCur(ImporteSinFormato(vPorcIva))
                Else
                    vPorcIva = ""
                    vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vSocio.CodIva, "N")
                    PorcIva = CCur(ImporteSinFormato(vPorcIva))
                End If
                
                tipoMov = vSocio.CodTipomLiq
                
                Set vTipoMov = New CTiposMov
                
                numfactu = vTipoMov.ConseguirContador(tipoMov)
                Do
                    numfactu = vTipoMov.ConseguirContador(tipoMov)
                    devuelve = DevuelveDesdeBDNew(cAgro, "rfactsoc", "numfactu", "codtipom", tipoMov, "T", , "numfactu", CStr(numfactu), "N", "fecfactu", FecFac, "F")
                    If devuelve <> "" Then
                        'Ya existe el contador incrementarlo
                        Existe = True
                        vTipoMov.IncrementarContador (tipoMov)
                        numfactu = vTipoMov.ConseguirContador(tipoMov)
                    Else
                        Existe = False
                    End If
                Loop Until Not Existe
                
                vParamAplic.PrimFactLiq = numfactu
                
            End If
        End If
    End If
    
    While Not Rs.EOF And B
        ActCalid = DBLet(Rs!codcalid, "N")
        ActVarie = DBLet(Rs!Codvarie, "N")
        actCampo = DBSet(Rs!codCampo, "N")
        ActSocio = DBSet(Rs!Codsocio, "N")
        
        If (ActCalid <> AntCalid Or AntCampo <> actCampo Or AntVarie <> ActVarie Or AntSocio <> ActSocio) Then
            ' kilos e importe por variedad campo
            Kilos = Kilos + KilosCal
            Importe = Importe + vImporte
            
            baseimpo = baseimpo + vImporte
            
            B = InsertLineaCalidad(tipoMov, CStr(numfactu), FecFac, AntVarie, AntCampo, CStr(AntCalid), CStr(DBLet(KilosCal, "N")), CStr(vImporte))
            KilosCal = 0
            vImporte = 0
            
            AntCalid = ActCalid
        End If
        
        If (ActVarie <> AntVarie Or actCampo <> AntCampo Or ActSocio <> AntSocio) Then
            If B Then ' descontamos el porcentaje de gastos de cooperativa
                Gastos = 0
                
                vPorcGasto = ""
                vPorcGasto = DevuelveDesdeBDNew(cAgro, "rcoope", "porcgast", "codcoope", vSocio.Cooperativa, "N")
                If vPorcGasto = "" Then vPorcGasto = "0"
                
                Gastos = Round2(Importe * CCur(ImporteSinFormato(vPorcGasto)) / 100, 2)
                Importe = Importe - Gastos
                baseimpo = baseimpo - Gastos
                
            End If
            
            If B Then ' descontamos los gastos de los albaranes
                Gastos = ObtenerGastosAlbaranesNew(AntSocio, AntVarie, AntCampo, cTabla, cWhere)
                Importe = Importe - Gastos
                baseimpo = baseimpo - Gastos
            End If

                        
            ' insertar linea de variedad, campo
            If B Then
                '[Monica]05/03/2014: a tramos Alzira
                If vParamAplic.Cooperativa = 4 Then
                    B = InsertLinea(tipoMov, CStr(numfactu), FecFac, AntVarie, AntCampo, CStr(Kilos), CStr(Importe), CStr(Gastos))
                Else
                    B = InsertLinea(tipoMov, CStr(numfactu), FecFac, AntVarie, AntCampo, CStr(Kilos), CStr(Importe), "0")
                End If
            End If
            
            If B Then
                ' tenemos que descontar los anticipos que tengamos para ello ( que no sean de gastos )
                Sql2 = "select rfactsoc_variedad.imporvar, rfactsoc_variedad.numfactu, rfactsoc_variedad.fecfactu "
                Sql2 = Sql2 & " from rfactsoc_variedad INNER JOIN rfactsoc ON rfactsoc_variedad.codtipom = rfactsoc.codtipom and "
                Sql2 = Sql2 & " rfactsoc_variedad.numfactu = rfactsoc.numfactu and rfactsoc_variedad.fecfactu = rfactsoc.fecfactu "
                Sql2 = Sql2 & " where rfactsoc_variedad.codtipom = " & DBSet(vSocio.CodTipomAnt, "T") ' antes era 'FAA' "
                Sql2 = Sql2 & " and rfactsoc.codsocio = " & DBSet(AntSocio, "N")
                Sql2 = Sql2 & " and rfactsoc.esanticipogasto = 0 "
                Sql2 = Sql2 & " and codvarie = " & DBSet(AntVarie, "N")
                Sql2 = Sql2 & " and codcampo = " & DBSet(AntCampo, "N")
                Sql2 = Sql2 & " and rfactsoc_variedad.descontado = 0"
                '[Monica]16/06/2016: la factura de anticipo no tiene que estar rectificada
                Sql2 = Sql2 & " and not (rfactsoc.codtipom, rfactsoc.numfactu, rfactsoc.fecfactu) in (select rectif_codtipom,rectif_numfactu,rectif_fecfactu from rfactsoc where not rectif_codtipom is null) "
                
                
                
                '[Monica]23/11/2012: en el caso de Natural solo tenemos que quitar los anticipos entre las fechas que me pongan
                If vParamAplic.Cooperativa = 9 Then
'[Monica]11/12/2013: sustituimos por los anticipos que queremos descontar
'                    If FecDesde <> "" Then Sql2 = Sql2 & " and rfactsoc.fecfactu >= " & DBSet(FecDesde, "F")
'                    If FecHasta <> "" Then Sql2 = Sql2 & " and rfactsoc.fecfactu <= " & DBSet(FecHasta, "F")
' si no seleccionamos ninguna no descontaremos ningun anticipo
                    If vFechas <> "" Then
                        Sql2 = Sql2 & " and rfactsoc.fecfactu in (" & vFechas & ")"
                    Else
                        Sql2 = Sql2 & " and rfactsoc.fecfactu = '1900-01-01' "
                    End If

                End If
                
                Set RS1 = New ADODB.Recordset
                RS1.Open Sql2, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
                
                While Not RS1.EOF
                    '[Monica]10/03/2014: si no permitimos facturas negativas
                    If baseimpo < DBLet(RS1.Fields(0).Value, "N") And NoPermitirFactNegativas Then
                
                
                    Else
                        baseimpo = baseimpo - DBLet(RS1.Fields(0).Value, "N")
                        Anticipos = Anticipos + DBLet(RS1.Fields(0).Value, "N")
                        
                        ' indicamos que los anticipos ya han sido descontados
                        Sql3 = "update rfactsoc_variedad set descontado = 1 where codtipom = " & DBSet(vSocio.CodTipomAnt, "T") ' antes era 'FAC'
                        Sql3 = Sql3 & " and numfactu = " & DBSet(RS1!numfactu, "N")
                        Sql3 = Sql3 & " and fecfactu = " & DBSet(RS1!fecfactu, "F") & " and codvarie = " & DBSet(AntVarie, "N")
                        Sql3 = Sql3 & " and codcampo = " & DBSet(AntCampo, "N")
                        
                        conn.Execute Sql3
                        
                        ' insertamos en la tabla de anticipos de liquidacion venta campo
                        Sql3 = "insert into tmpFact_anticipos (codtipom, numfactu, fecfactu, codtipomanti, numfactuanti, fecfactuanti, codvarieanti, codcampoanti, baseimpo) values ("
                        Sql3 = Sql3 & DBSet(vSocio.CodTipomLiq, "T") & "," ' antes era 'FAL'
                        Sql3 = Sql3 & DBSet(numfactu, "N") & "," & DBSet(FecFac, "F") & ","
                        Sql3 = Sql3 & DBSet(vSocio.CodTipomAnt, "T") & "," ' antes era 'FAA'
                        Sql3 = Sql3 & DBSet(RS1!numfactu, "N") & "," & DBSet(RS1!fecfactu, "F") & ","
                        Sql3 = Sql3 & DBSet(AntVarie, "N") & "," & DBSet(AntCampo, "N") & "," & DBSet(RS1!imporvar, "N") & ")"
                        
                        conn.Execute Sql3
                    End If
                    
                    RS1.MoveNext
                Wend
                
                Set RS1 = Nothing
                ' fin descontar anticipos
            
            End If
            
            '[Monica]10/05/2013: insertamos los albaranes
            If B Then
                '[Monica]05/03/2014: alzira entra a tramos
                If vParamAplic.Cooperativa = 4 Then
                    B = InsertarAlbaranesFactura(tipoMov, CStr(numfactu), FecFac, AntSocio, AntVarie, AntCampo, cTabla, cWhere, 0)
                Else
                    B = InsertarAlbaranesFactura(tipoMov, CStr(numfactu), FecFac, AntSocio, AntVarie, AntCampo, "", "", 5)
                End If
            End If
            
            If B Then
                AntVarie = ActVarie
                AntCampo = actCampo
                Kilos = 0
                Importe = 0
            End If
        End If
        
        If ActSocio <> AntSocio Then
            
            ImpoIva = Round2(baseimpo * PorcIva / 100, 2)
        
            Select Case DBLet(vSocio.TipoIRPF, "N")
                Case 0
                    ImpoReten = Round2((baseimpo + ImpoIva) * vParamAplic.PorcreteFacSoc / 100, 2)
                    BaseReten = (baseimpo + ImpoIva)
                    PorcReten = vParamAplic.PorcreteFacSoc
                Case 1
                    ImpoReten = Round2(baseimpo * vParamAplic.PorcreteFacSoc / 100, 2)
                    BaseReten = baseimpo
                    PorcReten = vParamAplic.PorcreteFacSoc
                Case 2
                    ImpoReten = 0
                    BaseReten = 0
                    PorcReten = 0
            End Select
            
            ImpoAFO = Round2((baseimpo + Anticipos) * vParamAplic.PorcenAFO / 100, 2)
            BaseAFO = baseimpo + Anticipos
            PorcAFO = vParamAplic.PorcenAFO
        
            TotalFac = baseimpo + ImpoIva - ImpoReten - ImpoAFO
            
            IncrementarProgresNew Pb1, 1
            
            'insertar cabecera de factura
            B = InsertCabecera(tipoMov, CStr(numfactu), FecFac, , , EsComplemen)
            
            If vParamAplic.Cooperativa = 4 Then
                '[Monica]04/01/2012: marcamos la factura como contabilizada y como pdte de recibir el nro de factura
                If B And vSocio.EmiteFact Then B = MarcarFactura(tipoMov, CStr(numfactu), FecFac, True)
            End If
            
            '[Monica]22/06/2017: en el caso de natural de montaña se marca como contabilizada
            If vParamAplic.Cooperativa = 9 Then
                If B Then B = MarcarFactura(tipoMov, CStr(numfactu), FecFac, False)
            End If
            
            
            If B Then B = InsertResumen(tipoMov, CStr(numfactu))
            
            If B Then B = RecalcularCalidades(tipoMov, CStr(numfactu), FecFac)
            
            If B Then B = vTipoMov.IncrementarContador(tipoMov)
            
            If B Then
                AntSocio = ActSocio
                
                Set vSocio = Nothing
                Set vSocio = New cSocio
                If vSocio.LeerDatos(ActSocio) Then
                    If vSocio.LeerDatosSeccion(AntSocio, vParamAplic.Seccionhorto) Then
                        '[Monica]05/03/2014: alzira entra a tramos
                        If vParamAplic.Cooperativa = 4 Then
                            '[Monica]29/04/2011: INTERNAS
                            If vSocio.EsFactADVInt Then
                                vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vParamAplic.CodIvaExeADV, "N")
                            Else
                                vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vSocio.CodIva, "N")
                            End If
                            PorcIva = CCur(ImporteSinFormato(vPorcIva))
                        Else
                            vPorcIva = ""
                            vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vSocio.CodIva, "N")
                            PorcIva = CCur(ImporteSinFormato(vPorcIva))
                        End If
                    End If
                    
                    tipoMov = vSocio.CodTipomLiq
                End If
                baseimpo = 0
                BaseReten = 0
                ImpoIva = 0
                ImpoReten = 0
                TotalFac = 0
                BaseAFO = 0
                ImpoAFO = 0
                
                Anticipos = 0
                
                numfactu = vTipoMov.ConseguirContador(tipoMov)
                Do
                    numfactu = vTipoMov.ConseguirContador(tipoMov)
                    devuelve = DevuelveDesdeBDNew(cAgro, "rfactsoc", "numfactu", "codtipom", tipoMov, "T", , "numfactu", CStr(numfactu), "N", "fecfactu", FecFac, "F")
                    If devuelve <> "" Then
                        'Ya existe el contador incrementarlo
                        Existe = True
                        vTipoMov.IncrementarContador (tipoMov)
                        numfactu = vTipoMov.ConseguirContador(tipoMov)
                    Else
                        Existe = False
                    End If
                Loop Until Not Existe
           End If
        End If
        
'        Recolect = DBLet(RS!Recolect, "N")
'
'        Select Case Recolect
'            Case 0
'                vPrecio = DBLet(RS!precoop, "N")
'                vImporte = vImporte + Round2(DBLet(RS!KilosNet, "N") * RS!precoop, 2)
'            Case 1
'                vPrecio = DBLet(RS!presocio, "N")
'                vImporte = vImporte + Round2(DBLet(RS!KilosNet, "N") * RS!presocio, 2)
'        End Select
        
        vImporte = DBLet(Rs!Importe, "N")
        KilosCal = DBLet(Rs!KilosNet, "N")
        vPrecio = Round2(vImporte / KilosCal, 2)
        
        HayReg = True
        
        Rs.MoveNext
    Wend
    
    ' ultimo registro si ha entrado
    If B And HayReg Then
        ' insertar linea de calidad
        Kilos = Kilos + KilosCal
        Importe = Importe + vImporte
        
        baseimpo = baseimpo + vImporte
        
        
        If B Then B = InsertLineaCalidad(tipoMov, CStr(numfactu), FecFac, ActVarie, actCampo, CStr(ActCalid), CStr(DBLet(KilosCal, "N")), CStr(vImporte))
        
        
        If B Then ' descontamos el porcentaje de gastos de cooperativa
            Gastos = 0
            
            vPorcGasto = ""
            vPorcGasto = DevuelveDesdeBDNew(cAgro, "rcoope", "porcgast", "codcoope", vSocio.Cooperativa, "N")
            If vPorcGasto = "" Then vPorcGasto = "0"
            
            Gastos = Round2(Importe * CCur(ImporteSinFormato(vPorcGasto)) / 100, 2)
            Importe = Importe - Gastos
            baseimpo = baseimpo - Gastos
        End If
        
        If B Then ' descontamos los gastos de los albaranes
            Gastos = ObtenerGastosAlbaranesNew(AntSocio, AntVarie, AntCampo, cTabla, cWhere)
            Importe = Importe - Gastos
            baseimpo = baseimpo - Gastos
        End If
                    
        ' insertar linea de variedad
        If B Then
            '[Monica]05/03/2014: para el caso de alzira
            If vParamAplic.Cooperativa = 4 Then
                B = InsertLinea(tipoMov, CStr(numfactu), FecFac, CStr(ActVarie), CStr(actCampo), CStr(Kilos), CStr(Importe), CStr(Gastos))
            Else
                B = InsertLinea(tipoMov, CStr(numfactu), FecFac, CStr(ActVarie), CStr(actCampo), CStr(Kilos), CStr(Importe), "0")
            End If
        End If
        
        ' tenemos que descontar los anticipos que tengamos para ello
        If B Then
            Sql2 = "select rfactsoc_variedad.imporvar, rfactsoc_variedad.numfactu, rfactsoc_variedad.fecfactu "
            Sql2 = Sql2 & " from rfactsoc_variedad INNER JOIN rfactsoc ON rfactsoc_variedad.codtipom = rfactsoc.codtipom and "
            Sql2 = Sql2 & " rfactsoc_variedad.numfactu = rfactsoc.numfactu and rfactsoc_variedad.fecfactu = rfactsoc.fecfactu "
            Sql2 = Sql2 & " where rfactsoc_variedad.codtipom = " & DBSet(vSocio.CodTipomAnt, "T") ' antes era 'FAA' "
            Sql2 = Sql2 & " and rfactsoc.codsocio = " & DBSet(ActSocio, "N")
            Sql2 = Sql2 & " and rfactsoc.esanticipogasto = 0 "
            Sql2 = Sql2 & " and codvarie = " & DBSet(ActVarie, "N")
            Sql2 = Sql2 & " and codcampo = " & DBSet(actCampo, "N")
            Sql2 = Sql2 & " and rfactsoc_variedad.descontado = 0"
            '[Monica]16/06/2016: la factura de anticipo no tiene que estar rectificada
            Sql2 = Sql2 & " and not (rfactsoc.codtipom, rfactsoc.numfactu, rfactsoc.fecfactu) in (select rectif_codtipom,rectif_numfactu,rectif_fecfactu from rfactsoc where not rectif_codtipom is null) "
            
            
            
            
            '[Monica]23/11/2012: en el caso de Natural solo tenemos que quitar los anticipos entre las fechas que me pongan
            If vParamAplic.Cooperativa = 9 Then
'[Monica]11/112/2013: sustituido por las fechas que ellos seleccionen
'                If FecDesde <> "" Then Sql2 = Sql2 & " and rfactsoc.fecfactu >= " & DBSet(FecDesde, "F")
'                If FecHasta <> "" Then Sql2 = Sql2 & " and rfactsoc.fecfactu <= " & DBSet(FecHasta, "F")
' si no seleccionamos ninguna no descontaremos ningun anticipo
                If vFechas <> "" Then
                    Sql2 = Sql2 & " and rfactsoc.fecfactu in (" & vFechas & ")"
                Else
                    Sql2 = Sql2 & " and rfactsoc.fecfactu = '1900-01-01' "
                End If

            End If
            
            Set RS1 = New ADODB.Recordset
            RS1.Open Sql2, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
            
            While Not RS1.EOF
                '[Monica]10/03/2014: si no permitimos facturas negativas
                If baseimpo < DBLet(RS1.Fields(0).Value, "N") And NoPermitirFactNegativas Then
            
            
                Else
                    baseimpo = baseimpo - DBLet(RS1.Fields(0).Value, "N")
                    Anticipos = Anticipos + DBLet(RS1.Fields(0).Value, "N")
                    
                    ' indicamos que los anticipos ya han sido descontados
                    Sql3 = "update rfactsoc_variedad set descontado = 1 where codtipom = " & DBSet(vSocio.CodTipomAnt, "T") ' antes era 'FAC'
                    Sql3 = Sql3 & " and numfactu = " & DBSet(RS1!numfactu, "N")
                    Sql3 = Sql3 & " and fecfactu = " & DBSet(RS1!fecfactu, "F") & " and codvarie = " & DBSet(ActVarie, "N")
                    Sql3 = Sql3 & " and codcampo = " & DBSet(actCampo, "N")
                    
                    conn.Execute Sql3
                    
                    ' insertamos en la tabla de anticipos de liquidacion venta campo
                    Sql3 = "insert into tmpFact_anticipos (codtipom, numfactu, fecfactu, codtipomanti, numfactuanti, fecfactuanti, codvarieanti, codcampoanti, baseimpo) values ("
                    Sql3 = Sql3 & DBSet(vSocio.CodTipomLiq, "T") & "," ' antes era 'FAL'
                    Sql3 = Sql3 & DBSet(numfactu, "N") & "," & DBSet(FecFac, "F") & ","
                    Sql3 = Sql3 & DBSet(vSocio.CodTipomAnt, "T") & "," ' antes era 'FAA'
                    Sql3 = Sql3 & DBSet(RS1!numfactu, "N") & "," & DBSet(RS1!fecfactu, "F") & ","
                    Sql3 = Sql3 & DBSet(ActVarie, "N") & "," & DBSet(actCampo, "N") & "," & DBSet(RS1!imporvar, "N") & ")"
                    
                    conn.Execute Sql3
                End If
                
                RS1.MoveNext
            Wend
            
            Set RS1 = Nothing
            ' fin descontar anticipos
        
        End If
        
        '[Monica]10/05/2013: insertamos los albaranes
        If B Then
            '[Monica]05/03/2014: alzira entra a tramos
            If vParamAplic.Cooperativa = 4 Then
                B = InsertarAlbaranesFactura(tipoMov, CStr(numfactu), FecFac, AntSocio, AntVarie, AntCampo, cTabla, cWhere, 0)
            Else
                B = InsertarAlbaranesFactura(tipoMov, CStr(numfactu), FecFac, ActSocio, ActVarie, actCampo, "", "", 5)
            End If
        End If
        
        
        ImpoIva = Round2(baseimpo * PorcIva / 100, 2)

        Select Case DBLet(vSocio.TipoIRPF, "N")
            Case 0
                ImpoReten = Round2((baseimpo + ImpoIva) * vParamAplic.PorcreteFacSoc / 100, 2)
                BaseReten = (baseimpo + ImpoIva)
                PorcReten = vParamAplic.PorcreteFacSoc
            Case 1
                ImpoReten = Round2(baseimpo * vParamAplic.PorcreteFacSoc / 100, 2)
                BaseReten = baseimpo
                PorcReten = vParamAplic.PorcreteFacSoc
            Case 2
                ImpoReten = 0
                BaseReten = 0
                PorcReten = 0
        End Select
        
        ImpoAFO = Round2((baseimpo + Anticipos) * vParamAplic.PorcenAFO / 100, 2)
        BaseAFO = baseimpo + Anticipos
        PorcAFO = vParamAplic.PorcenAFO

        TotalFac = baseimpo + ImpoIva - ImpoReten - ImpoAFO
        
        IncrementarProgresNew Pb1, 1
        
        vParamAplic.UltFactLiq = numfactu
        
        'insertar cabecera de factura
        B = InsertCabecera(tipoMov, CStr(numfactu), FecFac, , , EsComplemen)
        
        If vParamAplic.Cooperativa = 4 Then
            '[Monica]04/01/2012: marcamos la factura como contabilizada y como pdte de recibir el nro de factura
            If B And vSocio.EmiteFact Then B = MarcarFactura(tipoMov, CStr(numfactu), FecFac, True)
        End If
        
        '[Monica]22/06/2017: en el caso de natural de montaña se marca como contabilizada
        If vParamAplic.Cooperativa = 9 Then
            If B Then B = MarcarFactura(tipoMov, CStr(numfactu), FecFac, False)
        End If
        
        If B Then B = InsertResumen(tipoMov, CStr(numfactu))
        
        If B Then B = RecalcularCalidades(tipoMov, CStr(numfactu), FecFac)
        
        If B Then B = vTipoMov.IncrementarContador(tipoMov)
        
        'pasamos las temporales a las tablas
        If B Then B = PasarTemporales()
        
        If B Then B = (vParamAplic.Modificar = 1)
    End If
    
'    BorrarTMPs
    
    vSeccion.CerrarConta
    Set vSeccion = Nothing
    Set vSocio = Nothing
    
eFacturacion:
    If Err.Number <> 0 Or Not B Then
        conn.RollbackTrans
        FacturacionLiquidacionesCatadau = False
    Else
        conn.CommitTrans
        FacturacionLiquidacionesCatadau = True
    End If
End Function



Private Function ObtenerGastosAlbaranesNew(Socio As String, Varie As String, campo As String, cTabla As String, cWhere As String) As Currency
Dim SQL As String
Dim RS1 As ADODB.Recordset

    On Error Resume Next
    
    ObtenerGastosAlbaranesNew = 0
    
    SQL = "select sum(gastos) as total "
    SQL = SQL & " from tmpliquidacion1  "
    SQL = SQL & " where tmpliquidacion1.codsocio = " & DBSet(Socio, "N")
    SQL = SQL & " and  tmpliquidacion1.codvarie = " & DBSet(Varie, "N")
    SQL = SQL & " and tmpliquidacion1.CodCampo = " & DBSet(campo, "N")
    SQL = SQL & " and tmpliquidacion1.codusu = " & vUsu.Codigo
    
    
    Set RS1 = New ADODB.Recordset
    RS1.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    If Not RS1.EOF Then ObtenerGastosAlbaranesNew = DBLet(RS1.Fields(0).Value, "N")

    Set RS1 = Nothing
    

End Function




Public Function FacturacionAnticiposCatadau(cTabla As String, cWhere As String, FecFac As String, Pb1 As ProgressBar, Optional EsTercero As Boolean) As Boolean
Dim SQL As String
Dim Rs As ADODB.Recordset

Dim AntSocio As String
Dim AntVarie As String
Dim ActSocio As String
Dim ActVarie As String
Dim actCampo As String
Dim AntCampo As String
Dim ActCalid As String
Dim AntCalid As String

Dim HayReg As Boolean

Dim NumError As Long
Dim vImporte As Currency
Dim vPorcIva As String

Dim PrimFac As String
Dim UltFac As String

Dim tipoMov As String
Dim B As Boolean
Dim vSeccion As CSeccion
Dim Kilos As Currency
Dim KilosCal As Currency
Dim Importe As Currency

Dim devuelve As String
Dim Existe As Boolean

'Dim baseimpo As Currency
'Dim BaseReten As Currency
Dim Neto As Currency
'Dim ImpoIva As Currency
'Dim ImpoReten As Currency
'Dim TotalFac As Currency
Dim Recolect As Byte
Dim vPrecio As Currency

    On Error GoTo eFacturacion

    FacturacionAnticiposCatadau = False
    
    
    If EsTercero Then
        tipoMov = "FAT"
    Else
        tipoMov = "FAA"
    End If
    
    BorrarTMPs
    B = CrearTMPs()
    If Not B Then
         Exit Function
    End If
    
    conn.BeginTrans
    
'    '10/05/2013: antes de borrar cargo la tabla auxiliar de albaranes, voy a utilizar la tmpexecel
'    SQL = "delete from tmpinformes2 where codusu = " & vUsu.Codigo
'    conn.Execute SQL
'
'    SQL = "insert into tmpexcel (codusu, numalbar, fecalbar, codvarie, codsocio, kilosnet) select distinct " & vUsu.Codigo & ","
'    SQL = SQL & " importe1, fecha1, importe2, rhisfruta.codsocio, sum(importe4) from tmpinformes inner join rhisfruta on tmpinformes.importe1 = rhisfruta.numalbar "
'    SQL = SQL & " where codusu = " & vUsu.Codigo
'    SQL = SQL & " group by 1,2,3,4,5 "
'    conn.Execute SQL
    'hasta aqui 10/05/2013
    
    SQL = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute SQL

    SQL = "SELECT  tmpliquidacion.codsocio, tmpliquidacion.codvarie,"
    SQL = SQL & "tmpliquidacion.codcampo, tmpliquidacion.codcalid, "
    SQL = SQL & "sum(tmpliquidacion.kilosnet) as kilosnet, sum(tmpliquidacion.importe) as importe "
    SQL = SQL & " FROM  tmpliquidacion "
    SQL = SQL & " where codusu = " & DBSet(vUsu.Codigo, "N")
    
    ' ordenado por socio, variedad, campo, calidad
    SQL = SQL & " group by tmpliquidacion.codsocio, tmpliquidacion.codvarie, tmpliquidacion.codcampo, tmpliquidacion.codcalid "
    SQL = SQL & " order by tmpliquidacion.codsocio, tmpliquidacion.codvarie, tmpliquidacion.codcampo, tmpliquidacion.codcalid "
    
    
    Set vSeccion = New CSeccion
    
    If vSeccion.LeerDatos(vParamAplic.Seccionhorto) Then
        If Not vSeccion.AbrirConta Then
            Exit Function
        End If
    End If
    
    HayReg = False
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    If Not Rs.EOF Then
        AntSocio = CStr(DBLet(Rs!Codsocio, "N"))
        AntVarie = CStr(DBLet(Rs!Codvarie, "N"))
        AntCampo = CStr(DBLet(Rs!codCampo, "N"))
        AntCalid = CStr(DBLet(Rs!codcalid, "N"))
        
        ActSocio = CStr(DBLet(Rs!Codsocio, "N"))
        ActVarie = CStr(DBLet(Rs!Codvarie, "N"))
        actCampo = CStr(DBLet(Rs!codCampo, "N"))
        ActCalid = CStr(DBLet(Rs!codcalid, "N"))
    
        Set vSocio = New cSocio
        If vSocio.LeerDatos(ActSocio) Then
            If vSocio.LeerDatosSeccion(ActSocio, vParamAplic.Seccionhorto) Then
                baseimpo = 0
                BaseReten = 0
                ImpoIva = 0
                ImpoReten = 0
                TotalFac = 0
                BaseAFO = 0
                ImpoAFO = 0
                PorcAFO = 0
                
                Kilos = 0
                Importe = 0
                
                KilosCal = 0
                
                '[Monica]05/03/2014: entra alzira a tramos
                If vParamAplic.Cooperativa = 4 Then
                    '[Monica]29/04/2011: INTERNAS
                    If vSocio.EsFactADVInt Then
                        vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vParamAplic.CodIvaExeADV, "N")
                    Else
                        vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vSocio.CodIva, "N")
                    End If
                    PorcIva = CCur(ImporteSinFormato(vPorcIva))
                Else
                    vPorcIva = ""
                    vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vSocio.CodIva, "N")
                    PorcIva = CCur(ImporteSinFormato(vPorcIva))
                End If
                tipoMov = vSocio.CodTipomAnt
                
                Set vTipoMov = New CTiposMov
                
                numfactu = vTipoMov.ConseguirContador(tipoMov)
                Do
                    numfactu = vTipoMov.ConseguirContador(tipoMov)
                    devuelve = DevuelveDesdeBDNew(cAgro, "rfactsoc", "numfactu", "codtipom", tipoMov, "T", , "numfactu", CStr(numfactu), "N", "fecfactu", FecFac, "F")
                    If devuelve <> "" Then
                        'Ya existe el contador incrementarlo
                        Existe = True
                        vTipoMov.IncrementarContador (tipoMov)
                        numfactu = vTipoMov.ConseguirContador(tipoMov)
                    Else
                        Existe = False
                    End If
                Loop Until Not Existe
                
                vParamAplic.PrimFactAnt = numfactu
                
            End If
        End If
    End If
    
    While Not Rs.EOF And B
        ActCalid = DBLet(Rs!codcalid, "N")
        ActVarie = DBLet(Rs!Codvarie, "N")
        actCampo = DBSet(Rs!codCampo, "N")
        ActSocio = DBSet(Rs!Codsocio, "N")
        
        If (ActCalid <> AntCalid Or AntCampo <> actCampo Or AntVarie <> ActVarie Or AntSocio <> ActSocio) Then
            ' kilos e importe por variedad campo
            Kilos = Kilos + KilosCal
            Importe = Importe + vImporte
            
            baseimpo = baseimpo + vImporte
            
            B = InsertLineaCalidad(tipoMov, CStr(numfactu), FecFac, AntVarie, AntCampo, CStr(AntCalid), CStr(DBLet(KilosCal, "N")), CStr(vImporte))
            KilosCal = 0
            vImporte = 0
            
            AntCalid = ActCalid
        End If
        
        If (ActVarie <> AntVarie Or actCampo <> AntCampo Or ActSocio <> AntSocio) Then
            ' insertar linea de variedad, campo
            B = InsertLinea(tipoMov, CStr(numfactu), FecFac, AntVarie, AntCampo, CStr(Kilos), CStr(Importe), "0")
            
            '[Monica]10/05/2013: insertamos los albaranes
            If B Then
                '[Monica]05/03/2014: entra a tramos Alzira
                If vParamAplic.Cooperativa = 4 Then
                    B = InsertarAlbaranesFactura(tipoMov, CStr(numfactu), FecFac, AntSocio, AntVarie, AntCampo, cTabla, cWhere, 0)
                Else
                    B = InsertarAlbaranesFactura(tipoMov, CStr(numfactu), FecFac, AntSocio, AntVarie, AntCampo, "", "", 5)
                End If
            End If

            If B Then
                AntVarie = ActVarie
                AntCampo = actCampo
                Kilos = 0
                Importe = 0
            End If
        End If
        
        If ActSocio <> AntSocio Then
            ImpoIva = Round2(baseimpo * PorcIva / 100, 2)
        
            Select Case DBLet(vSocio.TipoIRPF, "N")
                Case 0
                    ImpoReten = Round2((baseimpo + ImpoIva) * vParamAplic.PorcreteFacSoc / 100, 2)
                    BaseReten = (baseimpo + ImpoIva)
                    PorcReten = vParamAplic.PorcreteFacSoc
                Case 1
                    ImpoReten = Round2(baseimpo * vParamAplic.PorcreteFacSoc / 100, 2)
                    BaseReten = baseimpo
                    PorcReten = vParamAplic.PorcreteFacSoc
                Case 2
                    ImpoReten = 0
                    BaseReten = 0
                    PorcReten = 0
            End Select
        
            TotalFac = baseimpo + ImpoIva - ImpoReten - ImpoAFO
            
            IncrementarProgresNew Pb1, 1
            
            'insertar cabecera de factura
            B = InsertCabecera(tipoMov, CStr(numfactu), FecFac)
            
            If vParamAplic.Cooperativa = 4 Then
                '[Monica]04/01/2012: marcamos la factura como contabilizada y como pdte de recibir el nro de factura
                If B And vSocio.EmiteFact Then B = MarcarFactura(tipoMov, CStr(numfactu), FecFac, True)
            End If
            
            '[Monica]22/06/2017: en el caso de natural de montaña se marca como contabilizada
            If vParamAplic.Cooperativa = 9 Then
                If B Then B = MarcarFactura(tipoMov, CStr(numfactu), FecFac, False)
            End If
            
            If B Then B = InsertResumen(tipoMov, CStr(numfactu))
            
            If B Then B = vTipoMov.IncrementarContador(tipoMov)
            
            If B Then
                AntSocio = ActSocio
                
                Set vSocio = Nothing
                Set vSocio = New cSocio
                If vSocio.LeerDatos(ActSocio) Then
                    If vSocio.LeerDatosSeccion(AntSocio, vParamAplic.Seccionhorto) Then
                        '[Monica]05/03/2014: entra alzira a tramos
                        If vParamAplic.Cooperativa = 4 Then
                            '[Monica]29/04/2011: INTERNAS
                            If vSocio.EsFactADVInt Then
                                vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vParamAplic.CodIvaExeADV, "N")
                            Else
                                vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vSocio.CodIva, "N")
                            End If
                            PorcIva = CCur(ImporteSinFormato(vPorcIva))
                        Else
                            
                            vPorcIva = ""
                            vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vSocio.CodIva, "N")
                            PorcIva = CCur(ImporteSinFormato(vPorcIva))
                        End If
                    End If
                    
                    tipoMov = vSocio.CodTipomAnt
                End If
                baseimpo = 0
                BaseReten = 0
                Neto = 0
                ImpoIva = 0
                ImpoReten = 0
                TotalFac = 0
                BaseAFO = 0
                ImpoAFO = 0
                
                numfactu = vTipoMov.ConseguirContador(tipoMov)
                Do
                    numfactu = vTipoMov.ConseguirContador(tipoMov)
                    devuelve = DevuelveDesdeBDNew(cAgro, "rfactsoc", "numfactu", "codtipom", tipoMov, "T", , "numfactu", CStr(numfactu), "N", "fecfactu", FecFac, "F")
                    If devuelve <> "" Then
                        'Ya existe el contador incrementarlo
                        Existe = True
                        vTipoMov.IncrementarContador (tipoMov)
                        numfactu = vTipoMov.ConseguirContador(tipoMov)
                    Else
                        Existe = False
                    End If
                Loop Until Not Existe
           End If
        End If
        
'        Recolect = DBLet(RS!Recolect, "N")
'
'        Select Case Recolect
'            Case 0
'                vPrecio = DBLet(RS!precoop, "N")
'                vImporte = vImporte + Round2(DBLet(RS!KilosNet, "N") * RS!precoop, 2)
'            Case 1
'                vPrecio = DBLet(RS!presocio, "N")
'                vImporte = vImporte + Round2(DBLet(RS!KilosNet, "N") * RS!presocio, 2)
'        End Select
'
'        KilosCal = KilosCal + DBLet(RS!KilosNet, "N")
        
        vImporte = DBLet(Rs!Importe, "N")
        KilosCal = DBLet(Rs!KilosNet, "N")
        vPrecio = Round2(vImporte / KilosCal, 2)
        
        HayReg = True
        
        Rs.MoveNext
    Wend
    
    ' ultimo registro si ha entrado
    If B And HayReg Then
        ' insertar linea de calidad
        Kilos = Kilos + KilosCal
        Importe = Importe + vImporte
        
        baseimpo = baseimpo + vImporte
        
        If B Then B = InsertLineaCalidad(tipoMov, CStr(numfactu), FecFac, ActVarie, actCampo, CStr(ActCalid), CStr(DBLet(KilosCal, "N")), CStr(vImporte))
        
        ' insertar linea de variedad
        If B Then B = InsertLinea(tipoMov, CStr(numfactu), FecFac, CStr(ActVarie), CStr(actCampo), CStr(Kilos), CStr(Importe), "0")
        
        
        '[Monica]10/05/2013: insertamos los albaranes
        If B Then
            '[Monica]05/03/2014: entra a tramos alzira
            If vParamAplic.Cooperativa = 4 Then
                B = InsertarAlbaranesFactura(tipoMov, CStr(numfactu), FecFac, ActSocio, ActVarie, actCampo, cTabla, cWhere, 0)
            Else
                B = InsertarAlbaranesFactura(tipoMov, CStr(numfactu), FecFac, ActSocio, ActVarie, actCampo, "", "", 5)
            End If
        End If
        
        ImpoIva = Round2(baseimpo * PorcIva / 100, 2)

        Select Case DBLet(vSocio.TipoIRPF, "N")
            Case 0
                ImpoReten = Round2((baseimpo + ImpoIva) * vParamAplic.PorcreteFacSoc / 100, 2)
                BaseReten = (baseimpo + ImpoIva)
                PorcReten = vParamAplic.PorcreteFacSoc
            Case 1
                ImpoReten = Round2(baseimpo * vParamAplic.PorcreteFacSoc / 100, 2)
                BaseReten = baseimpo
                PorcReten = vParamAplic.PorcreteFacSoc
            Case 2
                ImpoReten = 0
                BaseReten = 0
                PorcReten = 0
        End Select
        
'        BaseAFO = baseimpo
'        PorcAFO = vParamAplic.PorcenAFO
'        ImpoAFO = Round2(BaseAFO * PorcAFO / 100, 2)

        TotalFac = baseimpo + ImpoIva - ImpoReten - ImpoAFO
        
        IncrementarProgresNew Pb1, 1
        
        vParamAplic.UltFactAnt = numfactu
        
        'insertar cabecera de factura
        B = InsertCabecera(tipoMov, CStr(numfactu), FecFac)
        
        If vParamAplic.Cooperativa = 4 Then
            '[Monica]04/01/2012: marcamos la factura como contabilizada y como pdte de recibir el nro de factura
            If B And vSocio.EmiteFact Then B = MarcarFactura(tipoMov, CStr(numfactu), FecFac, True)
        End If
        
        
        '[Monica]22/06/2017: en el caso de natural de montaña se marca como contabilizada
        If vParamAplic.Cooperativa = 9 Then
            If B Then B = MarcarFactura(tipoMov, CStr(numfactu), FecFac, False)
        End If
        
        
        If B Then B = InsertResumen(tipoMov, CStr(numfactu))
        
        If B Then B = vTipoMov.IncrementarContador(tipoMov)
        
        'pasamos las temporales a las tablas
        If B Then B = PasarTemporales()
        
        If B Then B = (vParamAplic.Modificar = 1)
    End If
    
'    BorrarTMPs
    
    vSeccion.CerrarConta
    Set vSeccion = Nothing
    Set vSocio = Nothing
    
eFacturacion:
    If Err.Number <> 0 Or Not B Then
        conn.RollbackTrans
        FacturacionAnticiposCatadau = False
    
    Else
        conn.CommitTrans
        FacturacionAnticiposCatadau = True
    End If
End Function


Private Function ActualizarRegistrosFac(cTabla As String, cWhere As String) As Boolean
'Actualizar la marca de impreso
Dim SQL As String

    On Error GoTo eActualizarRegistros

    ActualizarRegistrosFac = False
    SQL = "update " & cTabla & ", usuarios.stipom set impreso = 1 "
    SQL = SQL & " where usuarios.stipom.codtipom = " & cTabla & ".codtipom "
    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        SQL = SQL & " and " & cWhere
    End If
    
    conn.Execute SQL
    
    ActualizarRegistrosFac = True
    Exit Function
    
eActualizarRegistros:
    MuestraError Err.Number, "Actualizando registros", Err.Description
End Function
'
' La diferencia con la FacturacionLiquidacionCatadau esta en que a diferencia de Catadau, aqui la factura es
' por campo: cada campo estará en una factura aunque sea del mismo socio
'
Public Function FacturacionLiquidacionIndustria(cTabla As String, cWhere As String, FecFac As String, Pb1 As ProgressBar, Optional CadenaAlbaranes As String) As Boolean
Dim SQL As String
Dim Sql3 As String

Dim Rs As ADODB.Recordset

Dim AntSocio As String
Dim AntVarie As String
Dim ActSocio As String
Dim ActVarie As String
Dim actCampo As String
Dim AntCampo As String
Dim ActCalid As String
Dim AntCalid As String

Dim HayReg As Boolean

Dim NumError As Long
Dim vImporte As Currency
Dim vPorcIva As String

Dim PrimFac As String
Dim UltFac As String

Dim tipoMov As String
Dim B As Boolean
Dim vSeccion As CSeccion
Dim Kilos As Currency
Dim KilosCal As Currency
Dim Importe As Currency

Dim devuelve As String
Dim Existe As Boolean

'Dim baseimpo As Currency
'Dim BaseReten As Currency
Dim Neto As Currency
'Dim ImpoIva As Currency
'Dim ImpoReten As Currency
'Dim TotalFac As Currency
Dim Recolect As Byte
Dim vPrecio As Currency
Dim ConGastos As Byte
Dim Gastos As Currency

    On Error GoTo eFacturacion

    FacturacionLiquidacionIndustria = False
'08052009 antes dentro de transacciones
    BorrarTMPs
    B = CrearTMPs()
    If Not B Then
         Exit Function
    End If
'08052009
    
    conn.BeginTrans
    
    tipoMov = "FLI"
    
    
    SQL = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute SQL

    SQL = "SELECT  tmpliquidacion.codsocio, tmpliquidacion.codvarie,"
    SQL = SQL & "tmpliquidacion.codcampo, tmpliquidacion.codcalid, "
    SQL = SQL & "sum(tmpliquidacion.kilosnet) as kilosnet, sum(tmpliquidacion.importe) as importe "
    SQL = SQL & " FROM  tmpliquidacion "
    SQL = SQL & " where codusu = " & DBSet(vUsu.Codigo, "N")

    ' ordenado por socio, variedad, campo, calidad
    SQL = SQL & " group by tmpliquidacion.codsocio, tmpliquidacion.codcampo, tmpliquidacion.codvarie, tmpliquidacion.codcalid "
    SQL = SQL & " order by tmpliquidacion.codsocio, tmpliquidacion.codcampo, tmpliquidacion.codvarie, tmpliquidacion.codcalid "
    
    Set vSeccion = New CSeccion
    
    If vSeccion.LeerDatos(vParamAplic.Seccionhorto) Then
        If Not vSeccion.AbrirConta Then
            Exit Function
        End If
    End If
    
    HayReg = False
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    If Not Rs.EOF Then
        AntSocio = CStr(DBLet(Rs!Codsocio, "N"))
        AntVarie = CStr(DBLet(Rs!Codvarie, "N"))
        AntCampo = CStr(DBLet(Rs!codCampo, "N"))
        AntCalid = CStr(DBLet(Rs!codcalid, "N"))
        
        ActSocio = CStr(DBLet(Rs!Codsocio, "N"))
        ActVarie = CStr(DBLet(Rs!Codvarie, "N"))
        actCampo = CStr(DBLet(Rs!codCampo, "N"))
        ActCalid = CStr(DBLet(Rs!codcalid, "N"))
    
        Set vSocio = New cSocio
        If vSocio.LeerDatos(ActSocio) Then
            If vSocio.LeerDatosSeccion(ActSocio, vParamAplic.Seccionhorto) Then
                baseimpo = 0
                BaseReten = 0
                ImpoIva = 0
                ImpoReten = 0
                TotalFac = 0
                ImpoAFO = 0
                BaseAFO = 0
                
                Anticipos = 0
                
                Kilos = 0
                Importe = 0
                
                KilosCal = 0
                
                vPorcIva = ""
                vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vSocio.CodIva, "N")
                PorcIva = CCur(ImporteSinFormato(vPorcIva))
                
'                tipoMov = vSocio.CodTipomLiq
                
                Set vTipoMov = New CTiposMov
                
                numfactu = vTipoMov.ConseguirContador(tipoMov)
                Do
                    numfactu = vTipoMov.ConseguirContador(tipoMov)
                    devuelve = DevuelveDesdeBDNew(cAgro, "rfactsoc", "numfactu", "codtipom", tipoMov, "T", , "numfactu", CStr(numfactu), "N", "fecfactu", FecFac, "F")
                    If devuelve <> "" Then
                        'Ya existe el contador incrementarlo
                        Existe = True
                        vTipoMov.IncrementarContador (tipoMov)
                        numfactu = vTipoMov.ConseguirContador(tipoMov)
                    Else
                        Existe = False
                    End If
                Loop Until Not Existe
                
                vParamAplic.PrimFactLiq = numfactu
                
            End If
        End If
    End If
    
    While Not Rs.EOF And B
        ActCalid = DBLet(Rs!codcalid, "N")
        ActVarie = DBLet(Rs!Codvarie, "N")
        actCampo = DBSet(Rs!codCampo, "N")
        ActSocio = DBSet(Rs!Codsocio, "N")
        
        If (ActCalid <> AntCalid Or AntCampo <> actCampo Or AntVarie <> ActVarie Or AntSocio <> ActSocio) Then
            ' kilos e importe por variedad campo
            Kilos = Kilos + KilosCal
            Importe = Importe + vImporte
            
            baseimpo = baseimpo + vImporte
            
            B = InsertLineaCalidad(tipoMov, CStr(numfactu), FecFac, AntVarie, AntCampo, CStr(AntCalid), CStr(DBLet(KilosCal, "N")), CStr(vImporte))
            KilosCal = 0
            vImporte = 0
            
            AntCalid = ActCalid
        End If
        
        If (ActVarie <> AntVarie Or actCampo <> AntCampo Or ActSocio <> AntSocio) Then
            If B Then ' descontamos los gastos de los albaranes
                Gastos = ObtenerGastosAlbaranesNew(AntSocio, AntVarie, AntCampo, cTabla, cWhere)
                Importe = Importe - Gastos
                baseimpo = baseimpo - Gastos
            End If
            
            '[Monica]08/04/2010: grabamos los albaranes que intervienen en la linea de factura
            If B Then
                B = InsertarAlbaranesFactura(tipoMov, CStr(numfactu), FecFac, AntSocio, AntVarie, AntCampo, cTabla, cWhere, 1)
            End If
            
                        
            ' insertar linea de variedad, campo
            If B Then
                B = InsertLinea(tipoMov, CStr(numfactu), FecFac, AntVarie, AntCampo, CStr(Kilos), CStr(Importe), CStr(Gastos))
            End If
            
            
            If B Then
                AntVarie = ActVarie
                Kilos = 0
                Importe = 0
            End If
        End If
        
        If actCampo <> AntCampo Or ActSocio <> AntSocio Then
            
            ImpoIva = Round2(baseimpo * PorcIva / 100, 2)
        
            Select Case DBLet(vSocio.TipoIRPF, "N")
                Case 0
                    ImpoReten = Round2((baseimpo + ImpoIva) * vParamAplic.PorcreteFacSoc / 100, 2)
                    BaseReten = (baseimpo + ImpoIva)
                    PorcReten = vParamAplic.PorcreteFacSoc
                Case 1
                    ImpoReten = Round2(baseimpo * vParamAplic.PorcreteFacSoc / 100, 2)
                    BaseReten = baseimpo
                    PorcReten = vParamAplic.PorcreteFacSoc
                Case 2
                    ImpoReten = 0
                    BaseReten = 0
                    PorcReten = 0
            End Select
            
' No hay fondo de aportacion en las facturas de industria
'            ImpoAFO = Round2((baseimpo + Anticipos) * vParamAplic.PorcenAFO / 100, 2)
'            BaseAFO = baseimpo + Anticipos
'            PorcAFO = vParamAplic.PorcenAFO
        

            TotalFac = baseimpo + ImpoIva - ImpoReten '- ImpoAFO
            
            IncrementarProgresNew Pb1, 1
            
            'insertar cabecera de factura
            B = InsertCabecera(tipoMov, CStr(numfactu), FecFac)
            
            '[Monica]04/01/2012: marcamos la factura como contabilizada y como pdte de recibir el nro de factura
            If B And vSocio.EmiteFact Then B = MarcarFactura(tipoMov, CStr(numfactu), FecFac, True)
            
            If B Then B = InsertResumen(tipoMov, CStr(numfactu))
            
            If B Then B = RecalcularCalidades(tipoMov, CStr(numfactu), FecFac)
            
            If B Then B = vTipoMov.IncrementarContador(tipoMov)
            
            If B Then
                AntSocio = ActSocio
                AntCampo = actCampo
                
                Set vSocio = Nothing
                Set vSocio = New cSocio
                If vSocio.LeerDatos(ActSocio) Then
                    If vSocio.LeerDatosSeccion(AntSocio, vParamAplic.Seccionhorto) Then
                        vPorcIva = ""
                        vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vSocio.CodIva, "N")
                        PorcIva = CCur(ImporteSinFormato(vPorcIva))
                    End If
                    
'                    tipoMov = vSocio.CodTipomLiq
                End If
                baseimpo = 0
                BaseReten = 0
                ImpoIva = 0
                ImpoReten = 0
                TotalFac = 0
'                BaseAFO = 0
'                ImpoAFO = 0
                
                
                numfactu = vTipoMov.ConseguirContador(tipoMov)
                Do
                    numfactu = vTipoMov.ConseguirContador(tipoMov)
                    devuelve = DevuelveDesdeBDNew(cAgro, "rfactsoc", "numfactu", "codtipom", tipoMov, "T", , "numfactu", CStr(numfactu), "N", "fecfactu", FecFac, "F")
                    If devuelve <> "" Then
                        'Ya existe el contador incrementarlo
                        Existe = True
                        vTipoMov.IncrementarContador (tipoMov)
                        numfactu = vTipoMov.ConseguirContador(tipoMov)
                    Else
                        Existe = False
                    End If
                Loop Until Not Existe
           End If
        End If
        
'        Recolect = DBLet(RS!Recolect, "N")
'
'        Select Case Recolect
'            Case 0
'                vPrecio = DBLet(RS!precoop, "N")
'                vImporte = vImporte + Round2(DBLet(RS!KilosNet, "N") * RS!precoop, 2)
'            Case 1
'                vPrecio = DBLet(RS!presocio, "N")
'                vImporte = vImporte + Round2(DBLet(RS!KilosNet, "N") * RS!presocio, 2)
'        End Select
        
        vImporte = DBLet(Rs!Importe, "N")
        KilosCal = DBLet(Rs!KilosNet, "N")
        vPrecio = Round2(vImporte / KilosCal, 2)
        
        HayReg = True
        
        Rs.MoveNext
    Wend
    
    ' ultimo registro si ha entrado
    If B And HayReg Then
        ' insertar linea de calidad
        Kilos = Kilos + KilosCal
        Importe = Importe + vImporte
        
        baseimpo = baseimpo + vImporte
        
        
        If B Then B = InsertLineaCalidad(tipoMov, CStr(numfactu), FecFac, ActVarie, actCampo, CStr(ActCalid), CStr(DBLet(KilosCal, "N")), CStr(vImporte))
        
        If B Then ' descontamos los gastos de los albaranes
            Gastos = ObtenerGastosAlbaranesNew(AntSocio, AntVarie, AntCampo, cTabla, cWhere)
            Importe = Importe - Gastos
            baseimpo = baseimpo - Gastos
        End If
        
        '[Monica]08/04/2010: grabamos los albaranes que intervienen en la linea de factura
        If B Then
            B = InsertarAlbaranesFactura(tipoMov, CStr(numfactu), FecFac, AntSocio, AntVarie, AntCampo, cTabla, cWhere, 1, CadenaAlbaranes)
        End If
                    
        ' insertar linea de variedad
        If B Then B = InsertLinea(tipoMov, CStr(numfactu), FecFac, CStr(ActVarie), CStr(actCampo), CStr(Kilos), CStr(Importe), CStr(Gastos))
        
        ImpoIva = Round2(baseimpo * PorcIva / 100, 2)

        Select Case DBLet(vSocio.TipoIRPF, "N")
            Case 0
                ImpoReten = Round2((baseimpo + ImpoIva) * vParamAplic.PorcreteFacSoc / 100, 2)
                BaseReten = (baseimpo + ImpoIva)
                PorcReten = vParamAplic.PorcreteFacSoc
            Case 1
                ImpoReten = Round2(baseimpo * vParamAplic.PorcreteFacSoc / 100, 2)
                BaseReten = baseimpo
                PorcReten = vParamAplic.PorcreteFacSoc
            Case 2
                ImpoReten = 0
                BaseReten = 0
                PorcReten = 0
        End Select
        
' No hay fondo de aportacion
'        ImpoAFO = Round2((baseimpo + Anticipos) * vParamAplic.PorcenAFO / 100, 2)
'        BaseAFO = baseimpo + Anticipos
'        PorcAFO = vParamAplic.PorcenAFO

        TotalFac = baseimpo + ImpoIva - ImpoReten '- ImpoAFO
        
        IncrementarProgresNew Pb1, 1
        
        vParamAplic.UltFactLiq = numfactu
        
        'insertar cabecera de factura
        B = InsertCabecera(tipoMov, CStr(numfactu), FecFac)
        
        '[Monica]04/01/2012: marcamos la factura como contabilizada y como pdte de recibir el nro de factura
        If B And vSocio.EmiteFact Then B = MarcarFactura(tipoMov, CStr(numfactu), FecFac, True)
        
        If B Then B = InsertResumen(tipoMov, CStr(numfactu))
        
        If B Then B = RecalcularCalidades(tipoMov, CStr(numfactu), FecFac)
        
        If B Then B = vTipoMov.IncrementarContador(tipoMov)
        
        'pasamos las temporales a las tablas
        If B Then B = PasarTemporales()
        
        If B Then B = (vParamAplic.Modificar = 1)
    End If
    
'    BorrarTMPs
    
    vSeccion.CerrarConta
    Set vSeccion = Nothing
    Set vSocio = Nothing
    
eFacturacion:
    If Err.Number <> 0 Or Not B Then
        conn.RollbackTrans
        FacturacionLiquidacionIndustria = False
    Else
        conn.CommitTrans
        FacturacionLiquidacionIndustria = True
    End If
End Function





Public Function TraspasoAlbaranesFacturas(cadSQL As String, cadWHERE As String, FechaFact As String, Banpr As String, ByRef PBar1 As ProgressBar, ByRef LblBar As Label, ImprimeLasFacturasGeneradas As Boolean, ByRef vTipoM As String, TextosCSB As String, Forpa As String, TAlmzBod As Byte) As Boolean
'IN -> cadSQL: cadena para seleccion de los Partes que vamos a Facturar
'      FechaFact: Fecha de la Factura
'      BanPr: Cod. de Banco Propio
'      Pbar1:  Una progressbar. Se puede mandar un NOTHING, y no pasa nada. Si no se manda
'              es que estamos en un proceso corto o que no necesitabaos un pb1, con lo cual NO muestro el PB1
'      Imprime: Si despues de generarlo los imprime
'
'       vTipom:  Que tipo de albaran es, para luego la impresion saber que factura imprime
'      TextosCSB:  Si lleva llevara 3 lineas para meter ent tesoreria
'      TAlmzBod: 0 = almazara    1 = bodega


'Desde Albaranes Genera las Facturas correspondientes
Dim RsAlb As ADODB.Recordset 'Ordenados por: tipofac,clien,dpto,forma pago, dtoppago, dtognral
Dim B As Boolean
Dim SQL As String

'Aqui Guardamos los datos del Albaran Anterior para comparar con el actual
Dim AntSocio As Long
Dim antDirec As Long
Dim antForpa As Byte
Dim antDtoPP As Single, antDtoGn As Single

'direc/dpto actual para controlar el valor nulo
Dim actDirec As Long

'Concatenamos todas las facturas generadas para listarlas en el informe
Dim ListFactu As String
Dim vFactuBOD As CFacturaBOD
Dim INC As Integer
Dim Condicion As Boolean 'condicion que comprueba para romper la agrupacion de albaranes a 1 factura

'Por si no mando una progressbar, que no de errores
Dim PgbVisible As Boolean

    On Error GoTo ETraspasoAlbFac

    TraspasoAlbaranesFacturas = False

    ListFactu = ""
        
    'comprobamos que no haya nadie facturando
    DesBloqueoManual ("BODFAC") 'facturas de bodega
    If Not BloqueoManual("BODFAC", "1") Then
        MsgBox "No se puede facturar. Hay otro usuario facturando.", vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If
    
    'Bloqueamos todos los albaranes que vamos a facturar (cabeceras y lineas)
    'Nota: esta bloqueando tambien los registros de la tabla clientes: sclien correspondientes
    SQL = " (rbodalbaran INNER JOIN rsocios ON rbodalbaran.codsocio=rsocios.codsocio ) INNER JOIN rbodalbaran_variedad ON rbodalbaran.numalbar=rbodalbaran_variedad.numalbar "
    SQL = "(" & SQL & ") INNER JOIN variedades ON rbodalbaran_variedad.codvarie = variedades.codvarie "
    SQL = "(" & SQL & ") INNER JOIN productos ON variedades.codprodu = productos.codprodu "
    If Not BloqueaRegistro(SQL, cadWHERE) Then
        Screen.MousePointer = vbDefault
        'comprobamos que no haya nadie facturando
        DesBloqueoManual ("BODFAC")
        Exit Function
    End If
    
   
    'Inicializar la Progress Bar
    PgbVisible = False
    If Not (PBar1 Is Nothing) Then
        If PBar1.visible Then PgbVisible = True
    End If
    If PgbVisible Then
        SQL = Replace(cadSQL, "*", "count(*)")
        
        Set RsAlb = New ADODB.Recordset
        RsAlb.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not RsAlb.EOF Then
            CargarProgresNew PBar1, CInt(RsAlb.Fields(0))
            LblBar.Caption = "Inicializando el proceso..."
        End If
        RsAlb.Close
        Set RsAlb = Nothing
    End If
    
        
    'EMPEZAMOS LA FACTURA
    Set vFactuBOD = New CFacturaBOD
    vFactuBOD.fecfactu = FechaFact 'Fecha para las Facturas

    'Marcar Partes que se van a Facturar
    '----------------------------------------
    SQL = cadSQL & " ORDER BY rbodalbaran.codsocio "
    Set RsAlb = New ADODB.Recordset
    RsAlb.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        
    'Agrupar los Albaranes posibles en una misma Factura
    'Calcular y Grabar Factura en la Tabla de Facturas
    'Albaranes de retirada(rbodalbaran, rbodalbaran_variedad) -> Factura (rbodfacturas, rbodfacturas_alb)
    '----------------------------------------------------
    B = True
    
    AntSocio = 0 'socio
    
    cadW = ""
    Errores = ""
    INC = 0
    
    While Not RsAlb.EOF
         INC = INC + 1
             
         LblBar.Caption = "Facturando: Albaranes Retirada"
         
         Condicion = (AntSocio <> RsAlb!Codsocio)
         
'             If (antClien <> RSalb!CodClien) Or (antDirec <> actDirec) Or (antForpa <> RSalb!codforpa) Or (antDtoPP <> RSalb!DtoPPago) Or (antDtoGn <> RSalb!DtoGnral) Then
         If Condicion Then
         '-----
            If cadW <> "" Then 'Facturacion PEndiente
                cadW = cadW & ") "
                
                If Not vFactuBOD.PasarAlbaranesAFactura2(TAlmzBod, cadW, TextosCSB, Forpa, ErroresAux, False) Then
                    If B Then B = False
                    AnyadirAvisos ErroresAux
                Else 'añadirlo a la lista de facturas a imprimir
                    If ListFactu = "" Then
                        ListFactu = vFactuBOD.numfactu
                    Else
                        ListFactu = ListFactu & "," & vFactuBOD.numfactu
                    End If
                End If
                If PgbVisible Then
                    LblBar.Caption = "Socio: " & Format(vFactuBOD.Socio, "000000") & " " & vFactuBOD.NombreSocio
                    IncrementarProgresNew PBar1, INC
                    INC = 0
                    DoEvents
                End If
'                espera 0.1
                
                'Empezamos una nueva Factura
                cadW = ""
            End If
            'Generar una Factura nueva
            vFactuBOD.Socio = RsAlb!Codsocio
            vFactuBOD.NombreSocio = RsAlb!nomsocio
            vFactuBOD.DomicilioSocio = DBLet(RsAlb!dirsocio, "T")
            vFactuBOD.CPostal = DBLet(RsAlb!codpostal, "T")
            vFactuBOD.Poblacion = DBLet(RsAlb!pobsocio, "T")
            vFactuBOD.Provincia = DBLet(RsAlb!prosocio, "T")
            vFactuBOD.nif = DBLet(RsAlb!nifSocio, "T")
            vFactuBOD.Telefono = DBLet(RsAlb!telsoci1, "T")
            vFactuBOD.ForPago = Forpa
            vFactuBOD.TipForPago = DBSet(DevuelveDesdeBDNew(cAgro, "forpago", "tipoforp", "codforpa", Forpa, "N"), "N")
            cadW = "  rbodalbaran.numalbar IN (" & RsAlb!numalbar
        Else
            cadW = cadW & ", " & RsAlb!numalbar
        End If
    
        'Guardamos datos del registro anterior
        AntSocio = RsAlb!Codsocio
        RsAlb.MoveNext
    Wend
    RsAlb.Close
    Set RsAlb = Nothing
        
    'Facturar la ultima Factura generada del blucle
    If cadW <> "" Then
        cadW = cadW & ")"
        If PgbVisible Then LblBar.Caption = "Socio: " & Format(vFactuBOD.Socio, "000000") & " - " & vFactuBOD.NombreSocio
        
        If Not vFactuBOD.PasarAlbaranesAFactura2(TAlmzBod, cadW, TextosCSB, Forpa, ErroresAux, False) Then
            If B Then B = False
            AnyadirAvisos "Error Facturando el Socio: " & Format(vFactuBOD.Socio, "000000") & " " & vFactuBOD.NombreSocio & vbCrLf & ErroresAux
        Else 'añadirlo a la lista de facturas a imprimir
            If ListFactu = "" Then
                ListFactu = vFactuBOD.numfactu
            Else
                ListFactu = ListFactu & "," & vFactuBOD.numfactu
            End If
        End If
        If PgbVisible Then
'            LblBar.Caption = "Cliente: " & Format(vFactu.Cliente, "000000") & " - " & vFactu.NombreClien
            IncrementarProgresNew PBar1, INC
        End If
        espera 0.2
    End If
    
    TipoFac = vFactuBOD.CodTipom
    Set vFactuBOD = Nothing
    TraspasoAlbaranesFacturas = True
    
    If B Then
        LblBar.Caption = "Proceso finalizado correctamente."
        MsgBox "Las Facturas de los Albaranes de Retirada seleccionados se generaron correctamente.", vbInformation
    Else
        LblBar.Caption = "Proceso finalizado con errores."
        SQL = "ATENCIÓN:" & vbCrLf
        MsgBox SQL & "No todas las Facturas se generaron correctamente!!!.", vbExclamation
        If Errores <> "" Then MostrarAvisos
    End If
    
    espera 0.2
    
    'Desbloqueamos ya no estamos facturando
    DesBloqueoManual ("BODFAC")
    TerminaBloquear
    
    
    If ImprimeLasFacturasGeneradas Then
        If ListFactu <> "" Then
            ImprimirFacturasBOD ListFactu, FechaFact, , False
        End If
    End If
    
ETraspasoAlbFac:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Facturando Albaranes de Retirada", Err.Description
        'comprobamos que no haya nadie facturando
        DesBloqueoManual ("BODFAC")
        TerminaBloquear
    End If
End Function




Public Function FacturacionAnticiposBodega(cTabla As String, cWhere As String, FecFac As String, Pb1 As ProgressBar) As Boolean
Dim SQL As String
Dim Rs As ADODB.Recordset

Dim AntSocio As String
Dim AntVarie As String
Dim ActSocio As String
Dim ActVarie As String
Dim actCampo As String
Dim AntCampo As String
Dim ActCalid As String
Dim AntCalid As String

Dim HayReg As Boolean

Dim NumError As Long
Dim vImporte As Currency
Dim vPorcIva As String

Dim PrimFac As String
Dim UltFac As String

Dim tipoMov As String
Dim B As Boolean
Dim vSeccion As CSeccion
Dim Kilos As Currency
Dim KilosCal As Currency
Dim Importe As Currency


Dim Kilos2 As Currency

Dim devuelve As String
Dim Existe As Boolean

'Dim baseimpo As Currency
'Dim BaseReten As Currency
Dim Neto As Currency
'Dim ImpoIva As Currency
'Dim ImpoReten As Currency
'Dim TotalFac As Currency
Dim Recolect As Byte
Dim vPrecio As Currency

    On Error GoTo eFacturacion

    FacturacionAnticiposBodega = False
    
    tipoMov = "FNB"
    
    BorrarTMPs
    B = CrearTMPs()
    If Not B Then
         Exit Function
    End If
    
    conn.BeginTrans
    
    
    SQL = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute SQL

    SQL = "SELECT  rhisfruta.codsocio, rhisfruta.codvarie, rhisfruta.codcampo, "
    SQL = SQL & " rprecios.precioindustria, "
    SQL = SQL & "rprecios.tipofact, sum(kilosnet) as kilosnet2 , sum(rhisfruta.kilosnet * kgradobonif) as kilosnet "
    SQL = SQL & " FROM  " & cTabla

    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        SQL = SQL & " WHERE " & cWhere
    End If
    ' ordenado por socio, variedad, campo, calidad
    SQL = SQL & " group by rhisfruta.codsocio, rhisfruta.codvarie, rhisfruta.codcampo, rprecios.precioindustria,rprecios.tipofact"
    SQL = SQL & " order by rhisfruta.codsocio, rhisfruta.codvarie, rhisfruta.codcampo, rprecios.precioindustria,rprecios.tipofact"
    
    Set vSeccion = New CSeccion
    
    If vSeccion.LeerDatos(vParamAplic.SeccionBodega) Then
        If Not vSeccion.AbrirConta Then
            Exit Function
        End If
    End If
    
    HayReg = False
    
    
    '[Monica]10/11/2010: calculamos el grado bonificado
'    CalcularGradoBonificado ctabla, cwhere
    If Not CalcularGradoBonificadoRealizado(cTabla, cWhere) Then
        MsgBox "No se ha realizado el cálculo del grado bonificado. Revise.", vbExclamation
        Exit Function
    End If
    
    
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    If Not Rs.EOF Then
        AntSocio = CStr(DBLet(Rs!Codsocio, "N"))
        AntVarie = CStr(DBLet(Rs!Codvarie, "N"))
        AntCampo = CStr(DBLet(Rs!codCampo, "N"))
        
        ActSocio = CStr(DBLet(Rs!Codsocio, "N"))
        ActVarie = CStr(DBLet(Rs!Codvarie, "N"))
        actCampo = CStr(DBLet(Rs!codCampo, "N"))
   
        Set vSocio = New cSocio
        If vSocio.LeerDatos(ActSocio) Then
            If vSocio.LeerDatosSeccion(ActSocio, vParamAplic.SeccionBodega) Then
                baseimpo = 0
                BaseReten = 0
                ImpoIva = 0
                ImpoReten = 0
                TotalFac = 0
                BaseAFO = 0
                ImpoAFO = 0
                PorcAFO = 0
                
                Kilos = 0  ' suma los kilogrados
                Importe = 0
                Kilos2 = 0 ' me suma los kilos netos
                 
                KilosCal = 0
                
                '[Monica]03/02/2016: Metemos las facturas internas en Quatretonda
                If vParamAplic.Cooperativa = 7 Then
                    vPorcIva = ""
                    If vSocio.EsFactADVInt Then
                                                                                                 '[Monica]16/06/2016: antes vParamAplic.CodIvaExeADV
                        vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vSeccion.TipIvaExento, "N")
                    Else
                        vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vSocio.CodIva, "N")
                    End If
                Else
                    vPorcIva = ""
                    vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vSocio.CodIva, "N")
                End If
                
                PorcIva = CCur(ImporteSinFormato(vPorcIva))
                
                tipoMov = vSocio.CodTipomAntBod
                
                Set vTipoMov = New CTiposMov
                If vTipoMov.Leer(tipoMov) Then
                    numfactu = vTipoMov.ConseguirContador(tipoMov)
                    Do
                        numfactu = vTipoMov.ConseguirContador(tipoMov)
                        devuelve = DevuelveDesdeBDNew(cAgro, "rfactsoc", "numfactu", "codtipom", tipoMov, "T", , "numfactu", CStr(numfactu), "N", "fecfactu", FecFac, "F")
                        If devuelve <> "" Then
                            'Ya existe el contador incrementarlo
                            Existe = True
                            vTipoMov.IncrementarContador (tipoMov)
                            numfactu = vTipoMov.ConseguirContador(tipoMov)
                        Else
                            Existe = False
                        End If
                    Loop Until Not Existe
                    
                    vParamAplic.PrimFactAntBOD = numfactu
                Else
                    B = False
                End If
            End If
        End If
    End If
    
    While Not Rs.EOF And B
        ActVarie = DBLet(Rs!Codvarie, "N")
        actCampo = DBSet(Rs!codCampo, "N")
        ActSocio = DBSet(Rs!Codsocio, "N")
        
'29/10/2010
'        If (AntCampo <> actCampo Or AntVarie <> ActVarie Or AntSocio <> ActSocio) Then
         If (AntVarie <> ActVarie Or AntSocio <> ActSocio) Then
            ' kilos e importe por variedad campo
            Kilos = Kilos + KilosCal
            Importe = Importe + vImporte
            
            baseimpo = baseimpo + vImporte
            
'            b = InsertLineaCalidad(tipoMov, CStr(numfactu), FecFac, AntVarie, AntCampo, CStr(AntCalid), CStr(DBLet(KilosCal, "N")), CStr(vImporte))
            
            ' insertar en las lineas de albaran
            If B Then B = InsertLineaAlbaranBodega(tipoMov, CStr(numfactu), FecFac, AntSocio, AntVarie, cTabla, cWhere)
            
            KilosCal = 0
            vImporte = 0
            
            AntVarie = ActVarie
            
        End If
        
'29/10/2010
'        If (ActVarie <> AntVarie Or actCampo <> AntCampo Or ActSocio <> AntSocio) Then
'        If (ActVarie <> AntVarie Or ActSocio <> AntSocio) Then
'            ' insertar linea de variedad, campo
'            b = InsertLinea(tipoMov, CStr(numfactu), FecFac, AntVarie, AntCampo, CStr(Kilos2), CStr(Importe), "0", CStr(KilosCal))
'
'            If b Then
'                AntVarie = ActVarie
'                AntCampo = actCampo
'                Kilos = 0
'                Importe = 0
'                Kilos2 = 0
'
'            End If
'        End If
        
        If ActSocio <> AntSocio Then
            ImpoIva = Round2(baseimpo * PorcIva / 100, 2)
        
            Select Case DBLet(vSocio.TipoIRPF, "N")
                Case 0
                    ImpoReten = Round2((baseimpo + ImpoIva) * vParamAplic.PorcreteFacSoc / 100, 2)
                    BaseReten = (baseimpo + ImpoIva)
                    PorcReten = vParamAplic.PorcreteFacSoc
                Case 1
                    ImpoReten = Round2(baseimpo * vParamAplic.PorcreteFacSoc / 100, 2)
                    BaseReten = baseimpo
                    PorcReten = vParamAplic.PorcreteFacSoc
                Case 2
                    ImpoReten = 0
                    BaseReten = 0
                    PorcReten = 0
            End Select
        
            TotalFac = baseimpo + ImpoIva - ImpoReten - ImpoAFO
            
            IncrementarProgresNew Pb1, 1
            
            'insertar cabecera de factura
            B = InsertCabecera(tipoMov, CStr(numfactu), FecFac)
            
            If B Then B = InsertResumen(tipoMov, CStr(numfactu))
            
            If B Then B = vTipoMov.IncrementarContador(tipoMov)
            
            If B Then
                AntSocio = ActSocio
                
                Set vSocio = Nothing
                Set vSocio = New cSocio
                If vSocio.LeerDatos(ActSocio) Then
                    If vSocio.LeerDatosSeccion(AntSocio, vParamAplic.SeccionBodega) Then
                        '[Monica]03/02/2016: Metemos las facturas internas en Quatretonda
                        If vParamAplic.Cooperativa = 7 Then
                            vPorcIva = ""
                            If vSocio.EsFactADVInt Then                                                   '[Monica]16/06/2016: antes vParamAplic.CodIvaExeADV
                                vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vSeccion.TipIvaExento, "N")
                            Else
                                vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vSocio.CodIva, "N")
                            End If
                        Else
                            vPorcIva = ""
                            vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vSocio.CodIva, "N")
                        End If
                        
                        PorcIva = CCur(ImporteSinFormato(vPorcIva))
                    End If
                    
                    tipoMov = vSocio.CodTipomAntBod
                End If
                baseimpo = 0
                BaseReten = 0
                Neto = 0
                ImpoIva = 0
                ImpoReten = 0
                TotalFac = 0
                BaseAFO = 0
                ImpoAFO = 0
                
                If vTipoMov.Leer(tipoMov) Then
                    numfactu = vTipoMov.ConseguirContador(tipoMov)
                    Do
                        numfactu = vTipoMov.ConseguirContador(tipoMov)
                        devuelve = DevuelveDesdeBDNew(cAgro, "rfactsoc", "numfactu", "codtipom", tipoMov, "T", , "numfactu", CStr(numfactu), "N", "fecfactu", FecFac, "F")
                        If devuelve <> "" Then
                            'Ya existe el contador incrementarlo
                            Existe = True
                            vTipoMov.IncrementarContador (tipoMov)
                            numfactu = vTipoMov.ConseguirContador(tipoMov)
                        Else
                            Existe = False
                        End If
                    Loop Until Not Existe
                Else
                    B = False
                End If
           End If
        End If
        
'        vPrecio = DBLet(Rs!precioindustria, "N")
'        vImporte = vImporte + Round2(DBLet(Rs!KilosNet, "N") * Rs!precioindustria, 2)
'
'        KilosCal = KilosCal + DBLet(Rs!KilosNet, "N") ' kilogrados
'        Kilos2 = Kilos2 + DBLet(Rs!Kilosnet2, "N") ' kilos netos
        
        vPrecio = DBLet(Rs!precioindustria, "N")
        vImporte = vImporte + Round2(DBLet(Rs!KilosNet, "N") * Rs!precioindustria, 2)

        KilosCal = KilosCal + DBLet(Rs!KilosNet, "N") ' kilogrados
        Kilos2 = Kilos2 + DBLet(Rs!Kilosnet2, "N") ' kilos netos
        
        B = InsertLinea(tipoMov, CStr(numfactu), FecFac, AntVarie, DBLet(Rs!codCampo, "N"), CStr(DBLet(Rs!Kilosnet2, "N")), CStr(Round2(DBLet(Rs!KilosNet, "N") * Rs!precioindustria, 2)), "0", CStr(DBLet(Rs!KilosNet, "N")))
'        AntVarie = ActVarie
'        AntCampo = actCampo
'
        
        HayReg = True
        
        Rs.MoveNext
    Wend
    
    ' ultimo registro si ha entrado
    If B And HayReg Then
        ' insertar linea de calidad
        Kilos = Kilos + KilosCal
        Importe = Importe + vImporte
        
        baseimpo = baseimpo + vImporte
        
'        If b Then b = InsertLineaCalidad(tipoMov, CStr(numfactu), FecFac, ActVarie, actCampo, CStr(ActCalid), CStr(DBLet(KilosCal, "N")), CStr(vImporte))
        
        ' insertar en las lineas de albaran
        If B Then B = InsertLineaAlbaranBodega(tipoMov, CStr(numfactu), FecFac, ActSocio, ActVarie, cTabla, cWhere)
        
'        ' insertar linea de variedad
'        If b Then b = InsertLinea(tipoMov, CStr(numfactu), FecFac, CStr(ActVarie), actCampo, CStr(Kilos2), CStr(vImporte), "0", CStr(Kilos))
        
        ImpoIva = Round2(baseimpo * PorcIva / 100, 2)

        Select Case DBLet(vSocio.TipoIRPF, "N")
            Case 0
                ImpoReten = Round2((baseimpo + ImpoIva) * vParamAplic.PorcreteFacSoc / 100, 2)
                BaseReten = (baseimpo + ImpoIva)
                PorcReten = vParamAplic.PorcreteFacSoc
            Case 1
                ImpoReten = Round2(baseimpo * vParamAplic.PorcreteFacSoc / 100, 2)
                BaseReten = baseimpo
                PorcReten = vParamAplic.PorcreteFacSoc
            Case 2
                ImpoReten = 0
                BaseReten = 0
                PorcReten = 0
        End Select
        
'        BaseAFO = baseimpo
'        PorcAFO = vParamAplic.PorcenAFO
'        ImpoAFO = Round2(BaseAFO * PorcAFO / 100, 2)

        TotalFac = baseimpo + ImpoIva - ImpoReten - ImpoAFO
        
        IncrementarProgresNew Pb1, 1
        
        vParamAplic.UltFactAntBOD = numfactu
        
        'insertar cabecera de factura
        B = InsertCabecera(tipoMov, CStr(numfactu), FecFac)
        
        If B Then B = InsertResumen(tipoMov, CStr(numfactu))
        
        If B Then B = vTipoMov.IncrementarContador(tipoMov)
        
        'pasamos las temporales a las tablas
        If B Then B = PasarTemporales()
        
        If B Then B = (vParamAplic.Modificar = 1)
    End If
    
'    BorrarTMPs
    
    vSeccion.CerrarConta
    Set vSeccion = Nothing
    Set vSocio = Nothing
    
eFacturacion:
    If Err.Number <> 0 Or Not B Then
        conn.RollbackTrans
        FacturacionAnticiposBodega = False
    Else
        conn.CommitTrans
        FacturacionAnticiposBodega = True
    End If
End Function





Public Function FacturacionLiquidacionesBodega(cTabla As String, cWhere As String, FecFac As String, Pb1 As ProgressBar, EsComplementaria As Boolean) As Boolean
Dim SQL As String
Dim Rs As ADODB.Recordset
Dim RS1 As ADODB.Recordset

Dim AntSocio As String
Dim AntVarie As String
Dim ActSocio As String
Dim ActVarie As String
Dim actCampo As String
Dim AntCampo As String
Dim ActAlbar As String
Dim AntAlbar As String

Dim HayReg As Boolean

Dim NumError As Long
Dim vImporte As Currency
Dim vPorcIva As String

Dim PrimFac As String
Dim UltFac As String

Dim tipoMov As String
Dim B As Boolean
Dim vSeccion As CSeccion
Dim Kilos As Currency
Dim KilosCal As Currency
Dim Importe As Currency

Dim devuelve As String
Dim Existe As Boolean

'Dim baseimpo As Currency
'Dim BaseReten As Currency
'Dim Neto As Currency
'Dim ImpoIva As Currency
'Dim ImpoReten As Currency
'Dim TotalFac As Currency
Dim Recolect As Byte
Dim vPrecio As Currency

Dim Sql2 As String
Dim Sql3 As String
Dim Sql5 As String


Dim GastosCoop As Currency
Dim GastosAlb As Currency
Dim vPorcGasto As String


    On Error GoTo eFacturacion

    FacturacionLiquidacionesBodega = False
    
'08052009 antes dentro de transacciones
    BorrarTMPs
    B = CrearTMPs()
    If Not B Then
         Exit Function
    End If
'08052009
    
    conn.BeginTrans
    
    tipoMov = "FLB"
    
    
    SQL = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute SQL

    SQL = "SELECT  rhisfruta.codsocio, rhisfruta.codvarie, rhisfruta.codcampo, rhisfruta.numalbar, "
    SQL = SQL & " rhisfruta.fecalbar,  rhisfruta.kilosbru, rhisfruta.kgradobonif as prestimado,  "
    SQL = SQL & "rprecios.precioindustria, rhisfruta.kilosnet "
    SQL = SQL & " FROM  " & cTabla

    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        SQL = SQL & " WHERE " & cWhere
    End If
    ' ordenado por socio, variedad, campo, numlabar, fecalbar
    SQL = SQL & " order by rhisfruta.codsocio, rhisfruta.codvarie, rhisfruta.codcampo, rhisfruta.numalbar, rhisfruta.fecalbar "
    
    Set vSeccion = New CSeccion
    
    If vSeccion.LeerDatos(vParamAplic.SeccionBodega) Then
        If Not vSeccion.AbrirConta Then
            Exit Function
        End If
    End If
    
    '[Monica]23/11/2012: añadida lo de si es complementaria
    If Not EsComplementaria Then
        '[Monica]10/11/2010: calculamos el grado bonificado
        If Not CalcularGradoBonificadoRealizado(cTabla, cWhere) Then
            MsgBox "No se ha realizado el cálculo del grado bonificado. Revise.", vbExclamation
            Exit Function
        End If
    End If
    
    HayReg = False
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    If Not Rs.EOF Then
        AntSocio = CStr(DBLet(Rs!Codsocio, "N"))
        AntVarie = CStr(DBLet(Rs!Codvarie, "N"))
        AntCampo = CStr(DBLet(Rs!codCampo, "N"))
        AntAlbar = CStr(DBLet(Rs!numalbar, "N"))
        
        ActSocio = CStr(DBLet(Rs!Codsocio, "N"))
        ActVarie = CStr(DBLet(Rs!Codvarie, "N"))
        actCampo = CStr(DBLet(Rs!codCampo, "N"))
        ActAlbar = CStr(DBLet(Rs!numalbar, "N"))
        Set vSocio = New cSocio
        If vSocio.LeerDatos(ActSocio) Then
            If vSocio.LeerDatosSeccion(ActSocio, vParamAplic.SeccionBodega) Then
                baseimpo = 0
                BaseReten = 0
                ImpoIva = 0
                ImpoReten = 0
                TotalFac = 0
                
                Anticipos = 0
                
                Kilos = 0
                Importe = 0
                
                '[Monica]03/02/2016: Metemos las facturas internas en Quatretonda
                If vParamAplic.Cooperativa = 7 Then
                    vPorcIva = ""
                    If vSocio.EsFactADVInt Then                                                  '[Monica]16/06/2016: antes vParamAplic.CodIvaExeADV
                        vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vSeccion.TipIvaExento, "N")
                    Else
                        vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vSocio.CodIva, "N")
                    End If
                Else
                    vPorcIva = ""
                    vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vSocio.CodIva, "N")
                End If
                
                PorcIva = CCur(ImporteSinFormato(vPorcIva))
                
                tipoMov = vSocio.CodTipomLiqBod
                
                Set vTipoMov = New CTiposMov
                
                numfactu = vTipoMov.ConseguirContador(tipoMov)
                Do
                    numfactu = vTipoMov.ConseguirContador(tipoMov)
                    devuelve = DevuelveDesdeBDNew(cAgro, "rfactsoc", "numfactu", "codtipom", tipoMov, "T", , "numfactu", CStr(numfactu), "N", "fecfactu", FecFac, "F")
                    If devuelve <> "" Then
                        'Ya existe el contador incrementarlo
                        Existe = True
                        vTipoMov.IncrementarContador (tipoMov)
                        numfactu = vTipoMov.ConseguirContador(tipoMov)
                    Else
                        Existe = False
                    End If
                Loop Until Not Existe
                
                vParamAplic.PrimFactLiqBOD = numfactu
                
            End If
        End If
    End If
    
    While Not Rs.EOF And B
        ActVarie = DBLet(Rs!Codvarie, "N")
        actCampo = DBSet(Rs!codCampo, "N")
        ActSocio = DBSet(Rs!Codsocio, "N")
        
        
        If (ActVarie <> AntVarie Or actCampo <> AntCampo Or ActSocio <> AntSocio) Then
            If B Then ' descontamos el porcentaje de gastos de cooperativa
                GastosCoop = 0
                
                vPorcGasto = ""
                vPorcGasto = vParamAplic.PorcGtoMantBOD
                If vPorcGasto = "" Then vPorcGasto = "0"
                
                '[Monica]23/11/2012: añadida lo de si es complementaria
                If Not EsComplementaria Then
                    GastosCoop = Round2(Importe * CCur(ImporteSinFormato(vPorcGasto)) / 100, 2)
                End If
                
                Importe = Importe - GastosCoop
                baseimpo = baseimpo - GastosCoop
                
            End If
            
            If B Then ' descontamos los gastos de los albaranes
                'Para el resto sigue como estaba
                '[Monica]23/11/2012: añadida lo de si es complementaria
                GastosAlb = 0
                If Not EsComplementaria Then
                    GastosAlb = ObtenerGastosAlbaranes(AntSocio, AntVarie, AntCampo, cTabla, cWhere, 1)
                End If
                Importe = Importe - GastosAlb
                baseimpo = baseimpo - GastosAlb
            End If
                        
            ' insertar linea de variedad, campo
            If B Then
                B = InsertLinea(tipoMov, CStr(numfactu), FecFac, AntVarie, AntCampo, CStr(Kilos), CStr(Importe), CStr(GastosAlb))
            End If
            
            If B Then
                '[Monica]23/11/2012: añadida lo de si es complementaria, solo descontamos los anticipos si no es complementaria
                If Not EsComplementaria Then
                    ' tenemos que descontar los anticipos que tengamos para ello
                    Sql2 = "select rfactsoc_variedad.imporvar, rfactsoc_variedad.numfactu, rfactsoc_variedad.fecfactu "
                    Sql2 = Sql2 & " from rfactsoc_variedad INNER JOIN rfactsoc ON rfactsoc_variedad.codtipom = rfactsoc.codtipom and "
                    Sql2 = Sql2 & " rfactsoc_variedad.numfactu = rfactsoc.numfactu and rfactsoc_variedad.fecfactu = rfactsoc.fecfactu "
                    Sql2 = Sql2 & " where rfactsoc_variedad.codtipom = " & DBSet(vSocio.CodTipomAntBod, "T") ' antes era 'FAA' "
                    Sql2 = Sql2 & " and rfactsoc.codsocio = " & DBSet(AntSocio, "N")
                    Sql2 = Sql2 & " and codvarie = " & DBSet(AntVarie, "N")
                    Sql2 = Sql2 & " and codcampo = " & DBSet(AntCampo, "N")
                    Sql2 = Sql2 & " and rfactsoc_variedad.descontado = 0"
                    
                    Set RS1 = New ADODB.Recordset
                    RS1.Open Sql2, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
                    
                    While Not RS1.EOF
                        baseimpo = baseimpo - DBLet(RS1.Fields(0).Value, "N")
                        Anticipos = Anticipos + DBLet(RS1.Fields(0).Value, "N")
                        
                        ' indicamos que los anticipos ya han sido descontados
                        Sql3 = "update rfactsoc_variedad set descontado = 1 where codtipom = " & DBSet(vSocio.CodTipomAntBod, "T") ' antes era 'FAC'
                        Sql3 = Sql3 & " and numfactu = " & DBSet(RS1!numfactu, "N")
                        Sql3 = Sql3 & " and fecfactu = " & DBSet(RS1!fecfactu, "F") & " and codvarie = " & DBSet(AntVarie, "N")
                        Sql3 = Sql3 & " and codcampo = " & DBSet(AntCampo, "N")
                        
                        conn.Execute Sql3
                        
                        ' insertamos en la tabla de anticipos de liquidacion
                        Sql3 = "insert into tmpFact_anticipos (codtipom, numfactu, fecfactu, codtipomanti, numfactuanti, fecfactuanti, codvarieanti, codcampoanti, baseimpo) values ("
                        Sql3 = Sql3 & DBSet(vSocio.CodTipomLiqBod, "T") & ","
                        Sql3 = Sql3 & DBSet(numfactu, "N") & "," & DBSet(FecFac, "F") & ","
                        Sql3 = Sql3 & DBSet(vSocio.CodTipomAntBod, "T") & ","
                        Sql3 = Sql3 & DBSet(RS1!numfactu, "N") & "," & DBSet(RS1!fecfactu, "F") & ","
                        Sql3 = Sql3 & DBSet(AntVarie, "N") & "," & DBSet(AntCampo, "N") & "," & DBSet(RS1!imporvar, "N") & ")"
                        
                        conn.Execute Sql3
                        
                        RS1.MoveNext
                    Wend
                    
                    Set RS1 = Nothing
                    ' fin descontar anticipos
                End If
            End If
            
            If B Then
                AntVarie = ActVarie
                AntCampo = actCampo
                Kilos = 0
                Importe = 0
            End If
        End If
        
        If ActSocio <> AntSocio Then
'            ' me machaco la base imponible por culpa de los redondeos
'            Sql5 = "select sum(if(importe is null,0,importe)) from tmpfact_albaran where codtipom =" & DBSet(tipoMov, "T")
'            Sql5 = Sql5 & " and numfactu = " & DBSet(numfactu, "N") & " and fecfactu = " & DBSet(FecFac, "F")
'
'            baseimpo = DevuelveValor(Sql5)
'            baseimpo = baseimpo - Round2(baseimpo * vParamAplic.PorcGtoMantBOD / 100, 2)
'
'            Sql5 = "select sum(if(imporgasto is null,0,imporgasto)) from tmpfact_albaran where codtipom =" & DBSet(tipoMov, "T")
'            Sql5 = Sql5 & " and numfactu = " & DBSet(numfactu, "N") & " and fecfactu = " & DBSet(FecFac, "F")
'
'            baseimpo = baseimpo - DevuelveValor(Sql5)
'
'            Sql5 = "select sum(if(baseimpo is null,0,baseimpo)) from tmpfact_anticipos where codtipom =" & DBSet(tipoMov, "T")
'            Sql5 = Sql5 & " and numfactu = " & DBSet(numfactu, "N") & " and fecfactu = " & DBSet(FecFac, "F")
'
'            baseimpo = baseimpo - DevuelveValor(Sql5)
            
            ImpoIva = Round2(baseimpo * PorcIva / 100, 2)
        
            Select Case DBLet(vSocio.TipoIRPF, "N")
                Case 0
                    ImpoReten = Round2((baseimpo + ImpoIva) * vParamAplic.PorcreteFacSoc / 100, 2)
                    BaseReten = (baseimpo + ImpoIva)
                    PorcReten = vParamAplic.PorcreteFacSoc
                Case 1
                    ImpoReten = Round2(baseimpo * vParamAplic.PorcreteFacSoc / 100, 2)
                    BaseReten = baseimpo
                    PorcReten = vParamAplic.PorcreteFacSoc
                Case 2
                    ImpoReten = 0
                    BaseReten = 0
                    PorcReten = 0
            End Select
            
        
            TotalFac = baseimpo + ImpoIva - ImpoReten
            
            IncrementarProgresNew Pb1, 1
            
            'insertar cabecera de factura
            B = InsertCabecera(tipoMov, CStr(numfactu), FecFac, False, False, EsComplementaria)
            
            If B Then B = InsertResumen(tipoMov, CStr(numfactu))
            
            If B Then B = vTipoMov.IncrementarContador(tipoMov)
            
            If B Then
                AntSocio = ActSocio
                
                Set vSocio = Nothing
                Set vSocio = New cSocio
                If vSocio.LeerDatos(ActSocio) Then
                    If vSocio.LeerDatosSeccion(AntSocio, vParamAplic.SeccionBodega) Then
                        '[Monica]03/02/2016: Metemos las facturas internas en Quatretonda
                        If vParamAplic.Cooperativa = 7 Then
                            vPorcIva = ""
                            If vSocio.EsFactADVInt Then                                             '[Monica]16/06/2016: antes vParamAplic.CodIvaExeADV
                                vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vSeccion.TipIvaExento, "N")
                            Else
                                vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vSocio.CodIva, "N")
                            End If
                        Else
                            vPorcIva = ""
                            vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vSocio.CodIva, "N")
                        End If
                        PorcIva = CCur(ImporteSinFormato(vPorcIva))
                    End If
                    
                    tipoMov = vSocio.CodTipomLiqBod
                End If
                baseimpo = 0
                BaseReten = 0
                ImpoIva = 0
                ImpoReten = 0
                TotalFac = 0
                
                Anticipos = 0
                
                numfactu = vTipoMov.ConseguirContador(tipoMov)
                Do
                    numfactu = vTipoMov.ConseguirContador(tipoMov)
                    devuelve = DevuelveDesdeBDNew(cAgro, "rfactsoc", "numfactu", "codtipom", tipoMov, "T", , "numfactu", CStr(numfactu), "N", "fecfactu", FecFac, "F")
                    If devuelve <> "" Then
                        'Ya existe el contador incrementarlo
                        Existe = True
                        vTipoMov.IncrementarContador (tipoMov)
                        numfactu = vTipoMov.ConseguirContador(tipoMov)
                    Else
                        Existe = False
                    End If
                Loop Until Not Existe
           End If
        End If
        
        vPrecio = DBLet(Rs!precioindustria, "N")
        '[Monica]23/11/2012: añadida lo de si es complementaria
        If Not EsComplementaria Then
            vImporte = Round2(DBLet(Rs!KilosNet, "N") * Rs!precioindustria * Rs!PrEstimado, 2)
        Else
            vImporte = Round2(DBLet(Rs!KilosNet, "N") * Rs!precioindustria, 2)
        End If
        
        B = InsertLineaAlbaran(tipoMov, CStr(numfactu), FecFac, Rs, CStr(vPrecio), CStr(vImporte))
        
        Importe = Importe + vImporte
        baseimpo = baseimpo + vImporte
        Kilos = Kilos + DBLet(Rs!KilosNet)
        
        HayReg = True
        
        Rs.MoveNext
    Wend
    
    ' ultimo registro si ha entrado
    If B And HayReg Then
        If B Then ' descontamos el porcentaje de gastos de cooperativa
            GastosCoop = 0
            
            vPorcGasto = ""
            vPorcGasto = vParamAplic.PorcGtoMantBOD
            If vPorcGasto = "" Then vPorcGasto = "0"
            '[Monica]23/11/2012: añadida lo de si es complementaria
            If Not EsComplementaria Then
                GastosCoop = Round2(Importe * CCur(ImporteSinFormato(vPorcGasto)) / 100, 2)
            End If
            Importe = Importe - GastosCoop
            baseimpo = baseimpo - GastosCoop
        End If
        
        If B Then ' descontamos los gastos de los albaranes
            GastosAlb = ObtenerGastosAlbaranes(AntSocio, AntVarie, AntCampo, cTabla, cWhere, 1)
            Importe = Importe - GastosAlb
            baseimpo = baseimpo - GastosAlb
        End If
        
                    
        ' insertar linea de variedad
        If B Then B = InsertLinea(tipoMov, CStr(numfactu), FecFac, CStr(ActVarie), CStr(actCampo), CStr(Kilos), CStr(Importe), CStr(GastosAlb))
        
        ' tenemos que descontar los anticipos que tengamos para ello
        If B Then
            '[Monica]23/11/2012: añadida lo de si es complementaria
            If Not EsComplementaria Then
                Sql2 = "select rfactsoc_variedad.imporvar, rfactsoc_variedad.numfactu, rfactsoc_variedad.fecfactu "
                Sql2 = Sql2 & " from rfactsoc_variedad INNER JOIN rfactsoc ON rfactsoc_variedad.codtipom = rfactsoc.codtipom and "
                Sql2 = Sql2 & " rfactsoc_variedad.numfactu = rfactsoc.numfactu and rfactsoc_variedad.fecfactu = rfactsoc.fecfactu "
                Sql2 = Sql2 & " where rfactsoc_variedad.codtipom = " & DBSet(vSocio.CodTipomAntBod, "T") ' antes era 'FAA' "
                Sql2 = Sql2 & " and rfactsoc.codsocio = " & DBSet(ActSocio, "N")
                Sql2 = Sql2 & " and codvarie = " & DBSet(ActVarie, "N")
                Sql2 = Sql2 & " and codcampo = " & DBSet(actCampo, "N")
                Sql2 = Sql2 & " and rfactsoc_variedad.descontado = 0"
                
                Set RS1 = New ADODB.Recordset
                RS1.Open Sql2, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
                
                While Not RS1.EOF
                    baseimpo = baseimpo - DBLet(RS1.Fields(0).Value, "N")
                    Anticipos = Anticipos + DBLet(RS1.Fields(0).Value, "N")
                    
                    ' indicamos que los anticipos ya han sido descontados
                    Sql3 = "update rfactsoc_variedad set descontado = 1 where codtipom = " & DBSet(vSocio.CodTipomAntBod, "T") ' antes era 'FAC'
                    Sql3 = Sql3 & " and numfactu = " & DBSet(RS1!numfactu, "N")
                    Sql3 = Sql3 & " and fecfactu = " & DBSet(RS1!fecfactu, "F") & " and codvarie = " & DBSet(ActVarie, "N")
                    Sql3 = Sql3 & " and codcampo = " & DBSet(actCampo, "N")
                    
                    conn.Execute Sql3
                    
                    ' insertamos en la tabla de anticipos de liquidacion venta campo
                    Sql3 = "insert into tmpFact_anticipos (codtipom, numfactu, fecfactu, codtipomanti, numfactuanti, fecfactuanti, codvarieanti, codcampoanti, baseimpo) values ("
                    Sql3 = Sql3 & DBSet(vSocio.CodTipomLiqBod, "T") & "," ' antes era 'FAL'
                    Sql3 = Sql3 & DBSet(numfactu, "N") & "," & DBSet(FecFac, "F") & ","
                    Sql3 = Sql3 & DBSet(vSocio.CodTipomAntBod, "T") & "," ' antes era 'FAA'
                    Sql3 = Sql3 & DBSet(RS1!numfactu, "N") & "," & DBSet(RS1!fecfactu, "F") & ","
                    Sql3 = Sql3 & DBSet(ActVarie, "N") & "," & DBSet(actCampo, "N") & "," & DBSet(RS1!imporvar, "N") & ")"
                    
                    conn.Execute Sql3
                    
                    RS1.MoveNext
                Wend
                
                Set RS1 = Nothing
                ' fin descontar anticipos
            End If
        End If
        
        
        ImpoIva = Round2(baseimpo * PorcIva / 100, 2)

        Select Case DBLet(vSocio.TipoIRPF, "N")
            Case 0
                ImpoReten = Round2((baseimpo + ImpoIva) * vParamAplic.PorcreteFacSoc / 100, 2)
                BaseReten = (baseimpo + ImpoIva)
                PorcReten = vParamAplic.PorcreteFacSoc
            Case 1
                ImpoReten = Round2(baseimpo * vParamAplic.PorcreteFacSoc / 100, 2)
                BaseReten = baseimpo
                PorcReten = vParamAplic.PorcreteFacSoc
            Case 2
                ImpoReten = 0
                BaseReten = 0
                PorcReten = 0
        End Select
        
        TotalFac = baseimpo + ImpoIva - ImpoReten
        
        IncrementarProgresNew Pb1, 1
        
        vParamAplic.UltFactLiqBOD = numfactu
        
        'insertar cabecera de factura
        B = InsertCabecera(tipoMov, CStr(numfactu), FecFac, False, False, EsComplementaria)
        
        If B Then B = InsertResumen(tipoMov, CStr(numfactu))
        
        If B Then B = vTipoMov.IncrementarContador(tipoMov)
        
        'pasamos las temporales a las tablas
        If B Then B = PasarTemporales()
        
        If B Then B = (vParamAplic.Modificar = 1)
    End If
    
'    BorrarTMPs
    
    vSeccion.CerrarConta
    Set vSeccion = Nothing
    Set vSocio = Nothing
    
eFacturacion:
    If Err.Number <> 0 Or Not B Then
        conn.RollbackTrans
        FacturacionLiquidacionesBodega = False
    Else
        conn.CommitTrans
        FacturacionLiquidacionesBodega = True
    End If
End Function



'########   ALMAZARA   ##########

Public Function FacturacionAnticiposAlmazara(cTabla As String, cWhere As String, FecFac As String, Pb1 As ProgressBar) As Boolean
Dim SQL As String
Dim Rs As ADODB.Recordset

Dim AntSocio As String
Dim AntVarie As String
Dim ActSocio As String
Dim ActVarie As String
Dim actCampo As String
Dim AntCampo As String
Dim ActCalid As String
Dim AntCalid As String

Dim HayReg As Boolean

Dim NumError As Long
Dim vImporte As Currency
Dim vPorcIva As String

Dim PrimFac As String
Dim UltFac As String

Dim tipoMov As String
Dim B As Boolean
Dim vSeccion As CSeccion
Dim Kilos As Currency
Dim KilosCal As Currency
Dim Importe As Currency

Dim devuelve As String
Dim Existe As Boolean

'Dim baseimpo As Currency
'Dim BaseReten As Currency
Dim Neto As Currency
'Dim ImpoIva As Currency
'Dim ImpoReten As Currency
'Dim TotalFac As Currency
Dim Recolect As Byte
Dim vPrecio As Currency
Dim campo As String

    On Error GoTo eFacturacion

    FacturacionAnticiposAlmazara = False
    
    tipoMov = "FNZ"
    
    BorrarTMPs
    B = CrearTMPs()
    If Not B Then
         Exit Function
    End If
    
    conn.BeginTrans
    
    
    SQL = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute SQL

    SQL = "SELECT  rhisfruta.codsocio, rhisfruta.codvarie, rhisfruta.numalbar, rhisfruta.fecalbar, rhisfruta.kilosbru, rhisfruta.prestimado, "
    SQL = SQL & "rprecios.precioindustria, "
    SQL = SQL & "rprecios.tipofact, sum(rhisfruta.kilosnet) as kilosnet "
    SQL = SQL & " FROM  " & cTabla

    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        SQL = SQL & " WHERE " & cWhere
    End If
    ' ordenado por socio, variedad, campo, calidad
    SQL = SQL & " group by rhisfruta.codsocio, rhisfruta.codvarie, rhisfruta.numalbar, rhisfruta.fecalbar, rhisfruta.kilosbru, rhisfruta.prestimado, rprecios.precioindustria,rprecios.tipofact"
    SQL = SQL & " order by rhisfruta.codsocio, rhisfruta.codvarie, rhisfruta.numalbar, rhisfruta.fecalbar, rhisfruta.kilosbru, rhisfruta.prestimado, rprecios.precioindustria,rprecios.tipofact"
    
    Set vSeccion = New CSeccion
    
    If vSeccion.LeerDatos(vParamAplic.SeccionAlmaz) Then
        If Not vSeccion.AbrirConta Then
            Exit Function
        End If
    End If
    
    HayReg = False
    
   ' en almazara no se insertan campos: metemos el minimo codcampo sin condiciones
    campo = "0" 'DevuelveValor("select min(codcampo) from rcampos")
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    If Not Rs.EOF Then
        AntSocio = CStr(DBLet(Rs!Codsocio, "N"))
        AntVarie = CStr(DBLet(Rs!Codvarie, "N"))
        
        ActSocio = CStr(DBLet(Rs!Codsocio, "N"))
        ActVarie = CStr(DBLet(Rs!Codvarie, "N"))
   
        Set vSocio = New cSocio
        If vSocio.LeerDatos(ActSocio) Then
            If vSocio.LeerDatosSeccion(ActSocio, vParamAplic.SeccionAlmaz) Then
                baseimpo = 0
                BaseReten = 0
                ImpoIva = 0
                ImpoReten = 0
                TotalFac = 0
                BaseAFO = 0
                ImpoAFO = 0
                PorcAFO = 0
                
                Kilos = 0
                Importe = 0
                
                KilosCal = 0
                
                vPorcIva = ""
                vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vSocio.CodIva, "N")
                PorcIva = CCur(ImporteSinFormato(vPorcIva))
                
                tipoMov = vSocio.CodTipomAntAlmz
                
                Set vTipoMov = New CTiposMov
                If vTipoMov.Leer(tipoMov) Then
                    numfactu = vTipoMov.ConseguirContador(tipoMov)
                    Do
                        numfactu = vTipoMov.ConseguirContador(tipoMov)
                        devuelve = DevuelveDesdeBDNew(cAgro, "rfactsoc", "numfactu", "codtipom", tipoMov, "T", , "numfactu", CStr(numfactu), "N", "fecfactu", FecFac, "F")
                        If devuelve <> "" Then
                            'Ya existe el contador incrementarlo
                            Existe = True
                            vTipoMov.IncrementarContador (tipoMov)
                            numfactu = vTipoMov.ConseguirContador(tipoMov)
                        Else
                            Existe = False
                        End If
                    Loop Until Not Existe
                    
                    vParamAplic.PrimFactAntAlmz = numfactu
                Else
                    B = False
                End If
            End If
        End If
    End If
    
    While Not Rs.EOF And B
        ActVarie = DBLet(Rs!Codvarie, "N")
        ActSocio = DBSet(Rs!Codsocio, "N")
        
        If (AntVarie <> ActVarie Or AntSocio <> ActSocio) Then
            ' kilos e importe por variedad campo
            
            baseimpo = baseimpo + Importe
            
            ' insertar linea de variedad, campo
            B = InsertLinea(tipoMov, CStr(numfactu), FecFac, AntVarie, campo, CStr(Kilos), CStr(Importe), "0")
            
            If B Then
                AntVarie = ActVarie
                Kilos = 0
                Importe = 0
            End If
        End If
        
        If ActSocio <> AntSocio Then
            ImpoIva = Round2(baseimpo * PorcIva / 100, 2)
        
            Select Case DBLet(vSocio.TipoIRPF, "N")
                Case 0
                    ImpoReten = Round2((baseimpo + ImpoIva) * vParamAplic.PorcreteFacSoc / 100, 2)
                    BaseReten = (baseimpo + ImpoIva)
                    PorcReten = vParamAplic.PorcreteFacSoc
                Case 1
                    ImpoReten = Round2(baseimpo * vParamAplic.PorcreteFacSoc / 100, 2)
                    BaseReten = baseimpo
                    PorcReten = vParamAplic.PorcreteFacSoc
                Case 2
                    ImpoReten = 0
                    BaseReten = 0
                    PorcReten = 0
            End Select
        
            TotalFac = baseimpo + ImpoIva - ImpoReten
            
            IncrementarProgresNew Pb1, 1
            
            'insertar cabecera de factura
            B = InsertCabecera(tipoMov, CStr(numfactu), FecFac)
            
            If B Then B = InsertResumen(tipoMov, CStr(numfactu))
            
            If B Then B = vTipoMov.IncrementarContador(tipoMov)
            
            If B Then
                AntSocio = ActSocio
                
                Set vSocio = Nothing
                Set vSocio = New cSocio
                If vSocio.LeerDatos(ActSocio) Then
                    If vSocio.LeerDatosSeccion(AntSocio, vParamAplic.SeccionAlmaz) Then
                        vPorcIva = ""
                        vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vSocio.CodIva, "N")
                        PorcIva = CCur(ImporteSinFormato(vPorcIva))
                    End If
                    
                    tipoMov = vSocio.CodTipomAntAlmz
                End If
                baseimpo = 0
                BaseReten = 0
                Neto = 0
                ImpoIva = 0
                ImpoReten = 0
                TotalFac = 0
                
                If vTipoMov.Leer(tipoMov) Then
                    numfactu = vTipoMov.ConseguirContador(tipoMov)
                    Do
                        numfactu = vTipoMov.ConseguirContador(tipoMov)
                        devuelve = DevuelveDesdeBDNew(cAgro, "rfactsoc", "numfactu", "codtipom", tipoMov, "T", , "numfactu", CStr(numfactu), "N", "fecfactu", FecFac, "F")
                        If devuelve <> "" Then
                            'Ya existe el contador incrementarlo
                            Existe = True
                            vTipoMov.IncrementarContador (tipoMov)
                            numfactu = vTipoMov.ConseguirContador(tipoMov)
                        Else
                            Existe = False
                        End If
                    Loop Until Not Existe
                Else
                    B = False
                End If
           End If
        End If
        
'[Monica]añadidas estas 3 lineas eliminada la del precio para el anticipo
        vPrecio = DBLet(Rs!precioindustria, "N")
        vImporte = Round2(DBLet(Rs!KilosNet, "N") * DBLet(Rs!PrEstimado, "N") / 100 * Rs!precioindustria, 2)
        
        B = InsertLineaAlbaran(tipoMov, CStr(numfactu), FecFac, Rs, CStr(vPrecio), CStr(vImporte), campo)
    
'        vPrecio = DBLet(Rs!precioindustria, "N")
'[Monica] hasta aqui

        Importe = Importe + Round2(DBLet(Rs!KilosNet, "N") * DBLet(Rs!PrEstimado, "N") / 100 * Rs!precioindustria, 2)
        
        Kilos = Kilos + DBLet(Rs!KilosNet, "N")
        
        HayReg = True
        
        Rs.MoveNext
    Wend
    
    ' ultimo registro si ha entrado
    If B And HayReg Then
        baseimpo = baseimpo + Importe
        
        ' insertar linea de variedad
        If B Then B = InsertLinea(tipoMov, CStr(numfactu), FecFac, CStr(ActVarie), CStr(campo), CStr(Kilos), CStr(Importe), "0")
        
        ImpoIva = Round2(baseimpo * PorcIva / 100, 2)

        Select Case DBLet(vSocio.TipoIRPF, "N")
            Case 0
                ImpoReten = Round2((baseimpo + ImpoIva) * vParamAplic.PorcreteFacSoc / 100, 2)
                BaseReten = (baseimpo + ImpoIva)
                PorcReten = vParamAplic.PorcreteFacSoc
            Case 1
                ImpoReten = Round2(baseimpo * vParamAplic.PorcreteFacSoc / 100, 2)
                BaseReten = baseimpo
                PorcReten = vParamAplic.PorcreteFacSoc
            Case 2
                ImpoReten = 0
                BaseReten = 0
                PorcReten = 0
        End Select
        
        TotalFac = baseimpo + ImpoIva - ImpoReten
        
        IncrementarProgresNew Pb1, 1
        
        vParamAplic.UltFactAntAlmz = numfactu
        
        'insertar cabecera de factura
        B = InsertCabecera(tipoMov, CStr(numfactu), FecFac)
        
        If B Then B = InsertResumen(tipoMov, CStr(numfactu))
        
        If B Then B = vTipoMov.IncrementarContador(tipoMov)
        
        'pasamos las temporales a las tablas
        If B Then B = PasarTemporales()
        
        If B Then B = (vParamAplic.Modificar = 1)
    End If
    
'    BorrarTMPs
    
    vSeccion.CerrarConta
    Set vSeccion = Nothing
    Set vSocio = Nothing
    
eFacturacion:
    If Err.Number <> 0 Or Not B Then
        conn.RollbackTrans
        FacturacionAnticiposAlmazara = False
    Else
        conn.CommitTrans
        FacturacionAnticiposAlmazara = True
    End If
End Function





Public Function FacturacionLiquidacionesAlmazara(cTabla As String, cWhere As String, FecFac As String, Pb1 As ProgressBar) As Boolean
Dim SQL As String
Dim Rs As ADODB.Recordset
Dim RS1 As ADODB.Recordset

Dim AntSocio As String
Dim AntVarie As String
Dim ActSocio As String
Dim ActVarie As String
Dim actCampo As String
Dim AntCampo As String
Dim ActAlbar As String
Dim AntAlbar As String

Dim HayReg As Boolean

Dim NumError As Long
Dim vImporte As Currency
Dim vPorcIva As String

Dim PrimFac As String
Dim UltFac As String

Dim tipoMov As String
Dim B As Boolean
Dim vSeccion As CSeccion
Dim Kilos As Currency
Dim KilosCal As Currency
Dim Importe As Currency

Dim devuelve As String
Dim Existe As Boolean

'Dim baseimpo As Currency
'Dim BaseReten As Currency
'Dim Neto As Currency
'Dim ImpoIva As Currency
'Dim ImpoReten As Currency
'Dim TotalFac As Currency
Dim Recolect As Byte
Dim vPrecio As Currency

Dim Sql2 As String
Dim Sql3 As String
Dim Sql5 As String


Dim GastosCoop As Currency
Dim GastosAlb As Currency
Dim vPorcGasto As String

Dim campo As String

    On Error GoTo eFacturacion

    FacturacionLiquidacionesAlmazara = False
    
'08052009 antes dentro de transacciones
    BorrarTMPs
    B = CrearTMPs()
    If Not B Then
         Exit Function
    End If
'08052009
    
    conn.BeginTrans
    
    tipoMov = "FLZ"
    
    
    SQL = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute SQL

    SQL = "SELECT  rhisfruta.codsocio, rhisfruta.codvarie, rhisfruta.numalbar, "
    SQL = SQL & " rhisfruta.fecalbar,  rhisfruta.kilosbru, rhisfruta.prestimado,  "
    SQL = SQL & "rprecios.precioindustria, rhisfruta.kilosnet "
    SQL = SQL & " FROM  " & cTabla

    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        SQL = SQL & " WHERE " & cWhere
    End If
    ' ordenado por socio, variedad, campo, numlabar, fecalbar
    SQL = SQL & " order by rhisfruta.codsocio, rhisfruta.codvarie, rhisfruta.numalbar, rhisfruta.fecalbar "
    
    Set vSeccion = New CSeccion
    
    If vSeccion.LeerDatos(vParamAplic.SeccionAlmaz) Then
        If Not vSeccion.AbrirConta Then
            Exit Function
        End If
    End If
    
    HayReg = False
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    If Not Rs.EOF Then
        AntSocio = CStr(DBLet(Rs!Codsocio, "N"))
        AntVarie = CStr(DBLet(Rs!Codvarie, "N"))
        AntAlbar = CStr(DBLet(Rs!numalbar, "N"))
        
        ActSocio = CStr(DBLet(Rs!Codsocio, "N"))
        ActVarie = CStr(DBLet(Rs!Codvarie, "N"))
        ActAlbar = CStr(DBLet(Rs!numalbar, "N"))
        Set vSocio = New cSocio
        If vSocio.LeerDatos(ActSocio) Then
            If vSocio.LeerDatosSeccion(ActSocio, vParamAplic.SeccionAlmaz) Then
                baseimpo = 0
                BaseReten = 0
                ImpoIva = 0
                ImpoReten = 0
                TotalFac = 0
                
                Anticipos = 0
                
                Kilos = 0
                Importe = 0
                
                vPorcIva = ""
                vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vSocio.CodIva, "N")
                PorcIva = CCur(ImporteSinFormato(vPorcIva))
                
                tipoMov = vSocio.CodTipomLiqAlmz
                
                Set vTipoMov = New CTiposMov
                
                numfactu = vTipoMov.ConseguirContador(tipoMov)
                Do
                    numfactu = vTipoMov.ConseguirContador(tipoMov)
                    devuelve = DevuelveDesdeBDNew(cAgro, "rfactsoc", "numfactu", "codtipom", tipoMov, "T", , "numfactu", CStr(numfactu), "N", "fecfactu", FecFac, "F")
                    If devuelve <> "" Then
                        'Ya existe el contador incrementarlo
                        Existe = True
                        vTipoMov.IncrementarContador (tipoMov)
                        numfactu = vTipoMov.ConseguirContador(tipoMov)
                    Else
                        Existe = False
                    End If
                Loop Until Not Existe
                
                vParamAplic.PrimFactLiqAlmz = numfactu
                
            End If
        End If
    End If
   
   ' en almazara no se insertan campos: metemos el minimo codcampo sin condiciones
    campo = "0" ' DevuelveValor("select min(codcampo) from rcampos")
    
    While Not Rs.EOF And B
        ActVarie = DBLet(Rs!Codvarie, "N")
        ActSocio = DBSet(Rs!Codsocio, "N")
        
        If (ActVarie <> AntVarie Or ActSocio <> AntSocio) Then
            If B Then ' descontamos los gastos de los albaranes
                'Para el resto sigue como estaba
                GastosAlb = ObtenerGastosAlbaranes(AntSocio, AntVarie, campo, cTabla, cWhere, 1, 1)
                Importe = Importe - GastosAlb
                baseimpo = baseimpo - GastosAlb
            End If
                        
            ' insertar linea de variedad, campo
            If B Then
                B = InsertLinea(tipoMov, CStr(numfactu), FecFac, AntVarie, campo, CStr(Kilos), CStr(Importe), CStr(GastosAlb))
            End If
            
            If B Then
                ' tenemos que descontar los anticipos que tengamos para ello
                Sql2 = "select rfactsoc_variedad.imporvar, rfactsoc_variedad.numfactu, rfactsoc_variedad.fecfactu "
                Sql2 = Sql2 & " from rfactsoc_variedad INNER JOIN rfactsoc ON rfactsoc_variedad.codtipom = rfactsoc.codtipom and "
                Sql2 = Sql2 & " rfactsoc_variedad.numfactu = rfactsoc.numfactu and rfactsoc_variedad.fecfactu = rfactsoc.fecfactu "
                Sql2 = Sql2 & " where rfactsoc_variedad.codtipom = " & DBSet(vSocio.CodTipomAntAlmz, "T") ' antes era 'FAA' "
                Sql2 = Sql2 & " and rfactsoc.codsocio = " & DBSet(AntSocio, "N")
'                Sql2 = Sql2 & " and codcampo = " & DBSet(Campo, "N")
                Sql2 = Sql2 & " and codvarie = " & DBSet(AntVarie, "N")
                Sql2 = Sql2 & " and rfactsoc_variedad.descontado = 0"
                
                Set RS1 = New ADODB.Recordset
                RS1.Open Sql2, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
                
                While Not RS1.EOF
                    baseimpo = baseimpo - DBLet(RS1.Fields(0).Value, "N")
                    Anticipos = Anticipos + DBLet(RS1.Fields(0).Value, "N")
                    
                    ' indicamos que los anticipos ya han sido descontados
                    Sql3 = "update rfactsoc_variedad set descontado = 1 where codtipom = " & DBSet(vSocio.CodTipomAntAlmz, "T") ' antes era 'FAC'
                    Sql3 = Sql3 & " and numfactu = " & DBSet(RS1!numfactu, "N")
                    Sql3 = Sql3 & " and fecfactu = " & DBSet(RS1!fecfactu, "F") & " and codvarie = " & DBSet(AntVarie, "N")
'                    Sql3 = Sql3 & " and codcampo = " & DBSet(Campo, "N")
                    
                    conn.Execute Sql3
                    
                    ' insertamos en la tabla de anticipos de liquidacion
                    Sql3 = "insert into tmpFact_anticipos (codtipom, numfactu, fecfactu, codtipomanti, numfactuanti, fecfactuanti, codvarieanti, codcampoanti, baseimpo) values ("
                    Sql3 = Sql3 & DBSet(vSocio.CodTipomLiqAlmz, "T") & ","
                    Sql3 = Sql3 & DBSet(numfactu, "N") & "," & DBSet(FecFac, "F") & ","
                    Sql3 = Sql3 & DBSet(vSocio.CodTipomAntAlmz, "T") & ","
                    Sql3 = Sql3 & DBSet(RS1!numfactu, "N") & "," & DBSet(RS1!fecfactu, "F") & ","
                    Sql3 = Sql3 & DBSet(AntVarie, "N") & "," & DBSet(campo, "N") & "," & DBSet(RS1!imporvar, "N") & ")"
                    
                    conn.Execute Sql3
                    
                    RS1.MoveNext
                Wend
                
                Set RS1 = Nothing
                ' fin descontar anticipos
            
            End If
            
            
            If B Then
                AntVarie = ActVarie
                Kilos = 0
                Importe = 0
            End If
        End If
        
        If ActSocio <> AntSocio Then
'            ' me machaco la base imponible por culpa de los redondeos
'            Sql5 = "select sum(if(importe is null,0,importe)) from tmpfact_albaran where codtipom =" & DBSet(tipoMov, "T")
'            Sql5 = Sql5 & " and numfactu = " & DBSet(numfactu, "N") & " and fecfactu = " & DBSet(FecFac, "F")
'
'            baseimpo = DevuelveValor(Sql5)
'
'            Sql5 = "select sum(if(imporgasto is null,0,imporgasto)) from tmpfact_albaran where codtipom =" & DBSet(tipoMov, "T")
'            Sql5 = Sql5 & " and numfactu = " & DBSet(numfactu, "N") & " and fecfactu = " & DBSet(FecFac, "F")
'
'            GastosAlb = DevuelveValor(Sql5)
'
'            Sql5 = "select sum(if(baseimpo is null,0,baseimpo)) from tmpfact_anticipos where codtipom =" & DBSet(tipoMov, "T")
'            Sql5 = Sql5 & " and numfactu = " & DBSet(numfactu, "N") & " and fecfactu = " & DBSet(FecFac, "F")
'
'            Anticipos = DevuelveValor(Sql5)
'
'            baseimpo = baseimpo - GastosAlb - Anticipos
            
            ImpoIva = Round2(baseimpo * PorcIva / 100, 2)
        
            Select Case DBLet(vSocio.TipoIRPF, "N")
                Case 0
                    ImpoReten = Round2((baseimpo + ImpoIva) * vParamAplic.PorcreteFacSoc / 100, 2)
                    BaseReten = (baseimpo + ImpoIva)
                    PorcReten = vParamAplic.PorcreteFacSoc
                Case 1
                    ImpoReten = Round2(baseimpo * vParamAplic.PorcreteFacSoc / 100, 2)
                    BaseReten = baseimpo
                    PorcReten = vParamAplic.PorcreteFacSoc
                Case 2
                    ImpoReten = 0
                    BaseReten = 0
                    PorcReten = 0
            End Select
            
        
            TotalFac = baseimpo + ImpoIva - ImpoReten
            
            IncrementarProgresNew Pb1, 1
            
            'insertar cabecera de factura
            B = InsertCabecera(tipoMov, CStr(numfactu), FecFac)
            
            If B Then B = InsertResumen(tipoMov, CStr(numfactu))
            
            If B Then B = vTipoMov.IncrementarContador(tipoMov)
            
            If B Then
                AntSocio = ActSocio
                
                Set vSocio = Nothing
                Set vSocio = New cSocio
                If vSocio.LeerDatos(ActSocio) Then
                    If vSocio.LeerDatosSeccion(AntSocio, vParamAplic.SeccionAlmaz) Then
                        vPorcIva = ""
                        vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vSocio.CodIva, "N")
                        PorcIva = CCur(ImporteSinFormato(vPorcIva))
                    End If
                    
                    tipoMov = vSocio.CodTipomLiqAlmz
                End If
                baseimpo = 0
                BaseReten = 0
                ImpoIva = 0
                ImpoReten = 0
                TotalFac = 0
                
                Anticipos = 0
                
                numfactu = vTipoMov.ConseguirContador(tipoMov)
                Do
                    numfactu = vTipoMov.ConseguirContador(tipoMov)
                    devuelve = DevuelveDesdeBDNew(cAgro, "rfactsoc", "numfactu", "codtipom", tipoMov, "T", , "numfactu", CStr(numfactu), "N", "fecfactu", FecFac, "F")
                    If devuelve <> "" Then
                        'Ya existe el contador incrementarlo
                        Existe = True
                        vTipoMov.IncrementarContador (tipoMov)
                        numfactu = vTipoMov.ConseguirContador(tipoMov)
                    Else
                        Existe = False
                    End If
                Loop Until Not Existe
           End If
        End If
        
        vPrecio = DBLet(Rs!precioindustria, "N")
        vImporte = Round2(DBLet(Rs!KilosNet, "N") * DBLet(Rs!PrEstimado, "N") / 100 * Rs!precioindustria, 2)
        
        B = InsertLineaAlbaran(tipoMov, CStr(numfactu), FecFac, Rs, CStr(vPrecio), CStr(vImporte), campo)
        
        Importe = Importe + vImporte
        baseimpo = baseimpo + vImporte
        Kilos = Kilos + DBLet(Rs!KilosNet)
        
        HayReg = True
        
        Rs.MoveNext
    Wend
    
    ' ultimo registro si ha entrado
    If B And HayReg Then
        If B Then ' descontamos los gastos de los albaranes
            GastosAlb = ObtenerGastosAlbaranes(AntSocio, AntVarie, AntCampo, cTabla, cWhere, 1, 1)
            Importe = Importe - GastosAlb
            baseimpo = baseimpo - GastosAlb
        End If
        
                    
        ' insertar linea de variedad
        If B Then B = InsertLinea(tipoMov, CStr(numfactu), FecFac, CStr(ActVarie), CStr(campo), CStr(Kilos), CStr(Importe), CStr(GastosAlb))
        
        ' tenemos que descontar los anticipos que tengamos para ello
        If B Then
            Sql2 = "select rfactsoc_variedad.imporvar, rfactsoc_variedad.numfactu, rfactsoc_variedad.fecfactu "
            Sql2 = Sql2 & " from rfactsoc_variedad INNER JOIN rfactsoc ON rfactsoc_variedad.codtipom = rfactsoc.codtipom and "
            Sql2 = Sql2 & " rfactsoc_variedad.numfactu = rfactsoc.numfactu and rfactsoc_variedad.fecfactu = rfactsoc.fecfactu "
            Sql2 = Sql2 & " where rfactsoc_variedad.codtipom = " & DBSet(vSocio.CodTipomAntAlmz, "T") ' antes era 'FAA' "
            Sql2 = Sql2 & " and rfactsoc.codsocio = " & DBSet(ActSocio, "N")
            Sql2 = Sql2 & " and codvarie = " & DBSet(ActVarie, "N")
'            Sql2 = Sql2 & " and codcampo = " & DBSet(Campo, "N")
            Sql2 = Sql2 & " and rfactsoc_variedad.descontado = 0"
            
            Set RS1 = New ADODB.Recordset
            RS1.Open Sql2, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
            
            While Not RS1.EOF
                baseimpo = baseimpo - DBLet(RS1.Fields(0).Value, "N")
                Anticipos = Anticipos + DBLet(RS1.Fields(0).Value, "N")
                
                ' indicamos que los anticipos ya han sido descontados
                Sql3 = "update rfactsoc_variedad set descontado = 1 where codtipom = " & DBSet(vSocio.CodTipomAntAlmz, "T") ' antes era 'FAC'
                Sql3 = Sql3 & " and numfactu = " & DBSet(RS1!numfactu, "N")
                Sql3 = Sql3 & " and fecfactu = " & DBSet(RS1!fecfactu, "F") & " and codvarie = " & DBSet(ActVarie, "N")
'                Sql3 = Sql3 & " and codcampo = " & DBSet(Campo, "N")
                
                conn.Execute Sql3
                
                ' insertamos en la tabla de anticipos de liquidacion venta campo
                Sql3 = "insert into tmpFact_anticipos (codtipom, numfactu, fecfactu, codtipomanti, numfactuanti, fecfactuanti, codvarieanti, codcampoanti, baseimpo) values ("
                Sql3 = Sql3 & DBSet(vSocio.CodTipomLiqAlmz, "T") & "," ' antes era 'FAL'
                Sql3 = Sql3 & DBSet(numfactu, "N") & "," & DBSet(FecFac, "F") & ","
                Sql3 = Sql3 & DBSet(vSocio.CodTipomAntAlmz, "T") & "," ' antes era 'FAA'
                Sql3 = Sql3 & DBSet(RS1!numfactu, "N") & "," & DBSet(RS1!fecfactu, "F") & ","
                Sql3 = Sql3 & DBSet(ActVarie, "N") & "," & DBSet(campo, "N") & "," & DBSet(RS1!imporvar, "N") & ")"
                
                conn.Execute Sql3
                
                RS1.MoveNext
            Wend
            
            Set RS1 = Nothing
            ' fin descontar anticipos
        
        End If
        
'        ' me machaco la base imponible por culpa de los redondeos
'        Sql5 = "select sum(if(importe is null,0,importe)) from tmpfact_albaran where codtipom =" & DBSet(tipoMov, "T")
'        Sql5 = Sql5 & " and numfactu = " & DBSet(numfactu, "N") & " and fecfactu = " & DBSet(FecFac, "F")
'
'        baseimpo = DevuelveValor(Sql5)
'
'        Sql5 = "select sum(if(imporgasto is null,0,imporgasto)) from tmpfact_albaran where codtipom =" & DBSet(tipoMov, "T")
'        Sql5 = Sql5 & " and numfactu = " & DBSet(numfactu, "N") & " and fecfactu = " & DBSet(FecFac, "F")
'
'        GastosAlb = DevuelveValor(Sql5)
'
'        Sql5 = "select sum(if(baseimpo is null,0,baseimpo)) from tmpfact_anticipos where codtipom =" & DBSet(tipoMov, "T")
'        Sql5 = Sql5 & " and numfactu = " & DBSet(numfactu, "N") & " and fecfactu = " & DBSet(FecFac, "F")
'
'        Anticipos = DevuelveValor(Sql5)
        
'        baseimpo = baseimpo - GastosAlb - Anticipos
        
        ImpoIva = Round2(baseimpo * PorcIva / 100, 2)

        Select Case DBLet(vSocio.TipoIRPF, "N")
            Case 0
                ImpoReten = Round2((baseimpo + ImpoIva) * vParamAplic.PorcreteFacSoc / 100, 2)
                BaseReten = (baseimpo + ImpoIva)
                PorcReten = vParamAplic.PorcreteFacSoc
            Case 1
                ImpoReten = Round2(baseimpo * vParamAplic.PorcreteFacSoc / 100, 2)
                BaseReten = baseimpo
                PorcReten = vParamAplic.PorcreteFacSoc
            Case 2
                ImpoReten = 0
                BaseReten = 0
                PorcReten = 0
        End Select
        
        TotalFac = baseimpo + ImpoIva - ImpoReten
        
        IncrementarProgresNew Pb1, 1
        
        vParamAplic.UltFactLiqAlmz = numfactu
        
        'insertar cabecera de factura
        B = InsertCabecera(tipoMov, CStr(numfactu), FecFac)
        
        If B Then B = InsertResumen(tipoMov, CStr(numfactu))
        
        If B Then B = vTipoMov.IncrementarContador(tipoMov)
        
        'pasamos las temporales a las tablas
        If B Then B = PasarTemporales()
        
        If B Then B = (vParamAplic.Modificar = 1)
    End If
    
'    BorrarTMPs
    
    vSeccion.CerrarConta
    Set vSeccion = Nothing
    Set vSocio = Nothing
    
eFacturacion:
    If Err.Number <> 0 Or Not B Then
        conn.RollbackTrans
        FacturacionLiquidacionesAlmazara = False
    Else
        conn.CommitTrans
        FacturacionLiquidacionesAlmazara = True
    End If
End Function


Private Function InsertarAlbaranesFactura(cCodTipom As String, cNumfactu As String, cFecfactu As String, Socio As String, Varie As String, campo As String, cTabla As String, cWhere As String, Tipo As Byte, Optional CadenaAlbaranes As String) As Boolean
' Tipo = 0 --> para facturas de liquidacion de Alzira
' Tipo = 1 --> para facturas de liquidacion de industria de Alzira

Dim SQL As String
    
    On Error GoTo eInsertarAlbaranesFactura
    
    
    InsertarAlbaranesFactura = False
    
    Select Case Tipo
        Case 0 ' liquidaciones normales de alzira
            SQL = "insert into tmpFact_albaran (codtipom, numfactu, fecfactu, numalbar, fecalbar, codvarie, codcampo,"
            SQL = SQL & "kilosbru, kilosnet, grado, precio, importe, imporgasto) "
            SQL = SQL & " SELECT DISTINCT " & DBSet(cCodTipom, "T") & "," & DBSet(cNumfactu, "N") & "," & DBSet(cFecfactu, "F") & ","
            SQL = SQL & " rhisfruta.numalbar, rhisfruta.fecalbar, rhisfruta.codvarie, rhisfruta.codcampo, rhisfruta.kilosbru, "
            SQL = SQL & "rhisfruta.kilosnet,0,0,0,0"
            SQL = SQL & " from rhisfruta "
            SQL = SQL & " where numalbar in (select rhisfruta.numalbar from " & cTabla & " where " & cWhere
            SQL = SQL & " and rhisfruta.codsocio = " & DBSet(Socio, "N")
            SQL = SQL & " and rhisfruta.codvarie = " & DBSet(Varie, "N")
            SQL = SQL & " and rhisfruta.codcampo = " & DBSet(campo, "N") & ")"
            
            conn.Execute SQL
    
        Case 1 ' liquidaciones de industria de alzira
            SQL = "insert into tmpFact_albaran (codtipom, numfactu, fecfactu, numalbar, fecalbar, codvarie, codcampo,"
            SQL = SQL & "kilosbru, kilosnet, grado, precio, importe, imporgasto) "
            SQL = SQL & " SELECT DISTINCT " & DBSet(cCodTipom, "T") & "," & DBSet(cNumfactu, "N") & "," & DBSet(cFecfactu, "F") & ","
            SQL = SQL & " rhisfruta.numalbar, rhisfruta.fecalbar, rhisfruta.codvarie, rhisfruta.codcampo, rhisfruta.kilosbru, "
            SQL = SQL & "rhisfruta.kilosnet,0,0,0,0"
            SQL = SQL & " from rhisfruta, tmpliquidacion1 "
            SQL = SQL & " where tmpliquidacion1.codusu = " & vUsu.Codigo
            SQL = SQL & " and tmpliquidacion1.codsocio = " & DBSet(Socio, "N")
            SQL = SQL & " and tmpliquidacion1.codvarie = " & DBSet(Varie, "N")
            SQL = SQL & " and tmpliquidacion1.codcampo = " & DBSet(campo, "N")
            SQL = SQL & " and tmpliquidacion1.codsocio = rhisfruta.codsocio "
            SQL = SQL & " and tmpliquidacion1.codvarie = rhisfruta.codvarie "
            SQL = SQL & " and tmpliquidacion1.codcampo = rhisfruta.codcampo "
            SQL = SQL & " and rhisfruta.fecalbar >= tmpliquidacion1.fechaini "
            SQL = SQL & " and rhisfruta.fecalbar <= tmpliquidacion1.fechafin "
            SQL = SQL & " and rhisfruta.tipoentr = 3 " ' industria directo
            If CadenaAlbaranes <> "" Then
                SQL = SQL & " and " & CadenaAlbaranes
            End If
            
            conn.Execute SQL
    
        Case 2 ' venta campo
            SQL = "insert into tmpFact_albaran (codtipom, numfactu, fecfactu, numalbar, fecalbar, codvarie, codcampo,"
            SQL = SQL & "kilosbru, kilosnet, grado, precio, importe, imporgasto) "
            SQL = SQL & " SELECT DISTINCT " & DBSet(cCodTipom, "T") & "," & DBSet(cNumfactu, "N") & "," & DBSet(cFecfactu, "F") & ","
            SQL = SQL & " rhisfruta.numalbar, rhisfruta.fecalbar, rhisfruta.codvarie, rhisfruta.codcampo, rhisfruta.kilosbru, "
            SQL = SQL & "rhisfruta.kilosnet,0,0,rhisfruta.impentrada,0"
            SQL = SQL & " from rhisfruta "
            SQL = SQL & " where numalbar in (select rhisfruta.numalbar from " & cTabla & " where " & cWhere
            SQL = SQL & " and rhisfruta.codsocio = " & DBSet(Socio, "N")
            SQL = SQL & " and rhisfruta.codvarie = " & DBSet(Varie, "N")
            SQL = SQL & " and rhisfruta.codcampo = " & DBSet(campo, "N") & ")"
            
            conn.Execute SQL
    
        Case 3 ' anticipo genericos
            SQL = "insert into tmpFact_albaran (codtipom, numfactu, fecfactu, numalbar, fecalbar, codvarie, codcampo,"
            SQL = SQL & "kilosbru, kilosnet, grado, precio, importe, imporgasto) "
            SQL = SQL & " SELECT DISTINCT " & DBSet(cCodTipom, "T") & "," & DBSet(cNumfactu, "N") & "," & DBSet(cFecfactu, "F") & ","
            SQL = SQL & " rclasifica.numnotac, rclasifica.fechaent, rclasifica.codvarie, rclasifica.codcampo, rclasifica.kilosbru, "
            SQL = SQL & "rclasifica.kilosnet,0,0,0,0"
            SQL = SQL & " from rclasifica "
            SQL = SQL & " where numnotac in (select rclasifica.numnotac from rclasifica where " & cWhere
            SQL = SQL & " and rclasifica.codsocio = " & DBSet(Socio, "N")
            SQL = SQL & " and rclasifica.codvarie = " & DBSet(Varie, "N") & ")"
            
            conn.Execute SQL
            
            SQL = "insert into tmpFact_albaran (codtipom, numfactu, fecfactu, numalbar, fecalbar, codvarie, codcampo,"
            SQL = SQL & "kilosbru, kilosnet, grado, precio, importe, imporgasto) "
            SQL = SQL & " SELECT DISTINCT " & DBSet(cCodTipom, "T") & "," & DBSet(cNumfactu, "N") & "," & DBSet(cFecfactu, "F") & ","
            SQL = SQL & " rhisfruta.numalbar, rhisfruta.fecalbar, rhisfruta.codvarie, rhisfruta.codcampo, rhisfruta.kilosbru, "
            SQL = SQL & "rhisfruta.kilosnet,0,0,0,0"
            SQL = SQL & " from rhisfruta "
            SQL = SQL & " where numalbar in (select rhisfruta.numalbar from rhisfruta where " & Replace(cWhere, "fechaent", "fecalbar")
            SQL = SQL & " and rhisfruta.codsocio = " & DBSet(Socio, "N")
            SQL = SQL & " and rhisfruta.codvarie = " & DBSet(Varie, "N") & ")"
            
            conn.Execute SQL
            
        Case 4 ' liquidaciones quatretonda
            SQL = "insert into tmpFact_albaran (codtipom, numfactu, fecfactu, numalbar, fecalbar, codvarie, codcampo,"
            SQL = SQL & "kilosbru, kilosnet, grado, precio, importe, imporgasto) "
            SQL = SQL & " SELECT DISTINCT " & DBSet(cCodTipom, "T") & "," & DBSet(cNumfactu, "N") & "," & DBSet(cFecfactu, "F") & ","
            SQL = SQL & " rhisfruta.numalbar, rhisfruta.fecalbar, rhisfruta.codvarie, rhisfruta.codcampo, rhisfruta.kilosbru, "
            SQL = SQL & "rhisfruta.kilosnet,0,0,0,0"
            SQL = SQL & " from rhisfruta "
            SQL = SQL & " where numalbar in (select rhisfruta.numalbar from " & cTabla & " where " & Replace(cWhere, "fechaent", "fecalbar")
            SQL = SQL & " and rhisfruta.codsocio = " & DBSet(Socio, "N")
'            Sql = Sql & " and rhisfruta.codcampo = " & DBSet(campo, "N")
            SQL = SQL & " and rhisfruta.codvarie = " & DBSet(Varie, "N") & ")"
            
            conn.Execute SQL
            
            
         Case 5 ' insertar albaranes de Proceso Catadau
            SQL = "insert into tmpFact_albaran (codtipom, numfactu, fecfactu, numalbar, fecalbar, codvarie, codcampo,"
            SQL = SQL & "kilosnet, grado, precio, importe, imporgasto) "
            SQL = SQL & " SELECT " & DBSet(cCodTipom, "T") & "," & DBSet(cNumfactu, "N") & "," & DBSet(cFecfactu, "F") & ","
            'SQL = SQL & " rhisfruta.numalbar, rhisfruta.fecalbar, rhisfruta.codvarie, rhisfruta.codcampo, tmpexcel.kilosnet, "
            SQL = SQL & " tmpinformes2.importe1, tmpinformes2.fecha1, tmpinformes2.importe2, rhisfruta.codcampo, "
            SQL = SQL & " sum(importe4),0,0,sum(importe5),importeb1 "
            SQL = SQL & " from rhisfruta inner join tmpinformes2 on rhisfruta.numalbar = tmpinformes2.importe1  "
            SQL = SQL & " where codusu = " & vUsu.Codigo
            SQL = SQL & " and rhisfruta.codsocio = " & DBSet(Socio, "N")
            SQL = SQL & " and rhisfruta.codvarie = " & DBSet(Varie, "N")
            SQL = SQL & " and rhisfruta.codcampo = " & DBSet(campo, "N")
            SQL = SQL & " group by 1,2,3,4,5,6,7,9,10,12  "
            conn.Execute SQL
         
         
            
    End Select
    
    InsertarAlbaranesFactura = True
    Exit Function
    
eInsertarAlbaranesFactura:
    MensError = "Error en la inserción en rfactsoc_albaranes de la factura " & cNumfactu & " del socio " & Socio
    MuestraError Err.Number, MensError
End Function


Public Function FacturacionTransporte(cTabla As String, cWhere As String, ctabla1 As String, cwhere1 As String, FecFac As String, Pb1 As ProgressBar, Fdesde As String, Fhasta As String) As Boolean
Dim vTrans As CTransportista
Dim tipoMov As String

Dim AntTrans As String
Dim ActTrans As String

Dim AntAlbar As String
Dim ActAlbar As String

Dim AntVarie As String
Dim ActVarie As String

Dim AntCampo As String
Dim actCampo As String
Dim B As Boolean
Dim SQL As String
Dim Sql2 As String

Dim Rs As ADODB.Recordset
Dim HayReg As Boolean
Dim vImporte As Currency
Dim vPorcIva As String
Dim devuelve As String
Dim Existe As Boolean

Dim Nregs As Long

Dim CodTraba As String
Dim ImpPagado As Currency
Dim DiasTrab As Long

On Error GoTo EFacturacionTransporte

    FacturacionTransporte = False

'08052009 antes dentro de transacciones
    BorrarTMPs
    B = CrearTMPs()
    If Not B Then
         Exit Function
    End If
'08052009
    
    conn.BeginTrans

    tipoMov = "FTR"
    
    
    Sql2 = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute Sql2
    
    
    'numalbar,fecalbar,codvarie,codcampo,kilosbru,kilosnet,precio,importe,
    SQL = "select rclasifica.codtrans, 0 numalbar, rclasifica.fechaent, rclasifica.codvarie, "
    SQL = SQL & "rclasifica.codcampo, rclasifica.numnotac, sum(if(isnull(rclasifica.kilosnet),0,rclasifica.kilosnet)) kilosnet, sum(if(isnull(rclasifica.kilosbru),0,rclasifica.kilosbru)) kilosbru, sum(if(isnull(rclasifica.impacarr),0,rclasifica.impacarr)) importe, sum(if(isnull(rclasifica.kilostra),0,rclasifica.kilostra)) kilostra from " & cTabla
    If cWhere <> "" Then SQL = SQL & " where " & cWhere
    
    SQL = SQL & " group by 1, 2, 3, 4, 5, 6"
    SQL = SQL & " having sum(rclasifica.impacarr) <> 0 "
    SQL = SQL & " union "
    SQL = SQL & "select rhisfruta_entradas.codtrans, rhisfruta_entradas.numalbar, rhisfruta_entradas.fechaent, rhisfruta.codvarie, "
    SQL = SQL & "rhisfruta.codcampo, rhisfruta_entradas.numnotac, sum(if(isnull(rhisfruta_entradas.kilosnet),0,rhisfruta_entradas.kilosnet)) kilosnet, sum(if(isnull(rhisfruta_entradas.kilosbru),0,rhisfruta_entradas.kilosbru)) kilosbru, sum(if(isnull(rhisfruta_entradas.impacarr),0,rhisfruta_entradas.impacarr)) importe, sum(if(isnull(rhisfruta_entradas.kilostra),0,rhisfruta_entradas.kilostra)) kilostra from " & ctabla1
    If cwhere1 <> "" Then SQL = SQL & " where " & cwhere1
    
    SQL = SQL & " group by 1, 2, 3, 4, 5, 6"
    SQL = SQL & " having sum(rhisfruta_entradas.impacarr) <> 0 "
    SQL = SQL & " order by 1, 2, 3, 4, 5, 6"
    
    
    Nregs = TotalRegistrosConsulta(SQL)
    Pb1.visible = True
    Pb1.Max = Nregs
    Pb1.Value = 0
    DoEvents
    
    HayReg = False
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    If Not Rs.EOF Then
        AntTrans = CStr(DBLet(Rs!codTrans, "T"))
        AntAlbar = CStr(DBLet(Rs!numalbar, "N"))
        AntVarie = CStr(DBLet(Rs!Codvarie, "N"))
        AntCampo = CStr(DBLet(Rs!codCampo, "N"))
        
        ActTrans = CStr(DBLet(Rs!codTrans, "T"))
        ActAlbar = CStr(DBLet(Rs!numalbar, "N"))
        ActVarie = CStr(DBLet(Rs!Codvarie, "N"))
        actCampo = CStr(DBLet(Rs!codCampo, "N"))
    
        Set vTrans = New CTransportista
        If vTrans.LeerDatos(ActTrans) Then
                baseimpo = 0
                BaseReten = 0
                ImpoIva = 0
                ImpoReten = 0
                TotalFac = 0
                ImpoAFO = 0
                BaseAFO = 0
                
                vPorcIva = ""
                '[Monica]17/10/2011: FACTURAS INTERNAS
                If vTrans.EsFactTraInterna Then
                    vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vParamAplic.CodIvaExeADV, "N")
                Else
                    vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vTrans.CodIva, "N")
                End If
                
                If vPorcIva = "" Then
                    MsgBox "El transportista " & vTrans.Codigo & " no tiene iva. Revise.", vbExclamation
                    B = False
                Else
                    PorcIva = CCur(ImporteSinFormato(vPorcIva))
                End If
                
'                tipoMov = vSocio.CodTipomLiq
                
                If B Then
                    '[Monica] 27/07/2010 dependiendo del parametro hemos de coger el contador global o el del transportista
                    '[Monica]05/11/2012 si es una factura interna en Alzira cogemos el contador global, no el del transportista
                    If vParamAplic.TipoContadorTRA = 0 Or (vParamAplic.Cooperativa = 4 And vTrans.EsFactTraInterna) Then ' contador global
                        Set vTipoMov = New CTiposMov
                        numfactu = vTipoMov.ConseguirContador(tipoMov)
                        Do
                            numfactu = vTipoMov.ConseguirContador(tipoMov)
                            devuelve = DevuelveDesdeBDNew(cAgro, "rfacttra", "numfactu", "codtipom", tipoMov, "T", , "numfactu", CStr(numfactu), "N", "fecfactu", FecFac, "F")
                            If devuelve <> "" Then
                                'Ya existe el contador incrementarlo
                                Existe = True
                                vTipoMov.IncrementarContador (tipoMov)
                                numfactu = vTipoMov.ConseguirContador(tipoMov)
                            Else
                                Existe = False
                            End If
                        Loop Until Not Existe
                    Else
                        numfactu = vTrans.ConseguirContador()
                        Do
                            numfactu = vTrans.ConseguirContador()
                            SQL = "select numfactu from rfacttra where codtipom = " & DBSet(tipoMov, "T") & " and numfactu = " & DBSet(numfactu, "N") & " and fecfactu = " & DBSet(FecFac, "F") & " and codtrans = " & DBSet(vTrans.Codigo, "T")
                            devuelve = DevuelveValor(SQL) 'DevuelveDesdeBDNew(cAgro, "rfacttra", "numfactu", "codtipom", tipoMov, "T", , "numfactu", CStr(numfactu), "N", "fecfactu", FecFac, "F")
                            If devuelve <> 0 Then
                                'Ya existe el contador incrementarlo
                                Existe = True
                                vTrans.IncrementarContador
                                numfactu = vTrans.ConseguirContador()
                            Else
                                Existe = False
                            End If
                        Loop Until Not Existe
                    End If
                End If
        Else
            B = False
        End If
    End If
    
    While Not Rs.EOF And B
        ActTrans = DBLet(Rs!codTrans, "T")
        ActAlbar = DBSet(Rs!numalbar, "N")
        ActVarie = DBSet(Rs!Codvarie, "N")
        actCampo = DBSet(Rs!codCampo, "N")
        
'        If (ActVarie <> AntVarie Or ActCampo <> AntCampo Or ActAlbar <> AntAlbar Or ActTrans <> AntTrans) Then
'
'            ' insertar linea de variedad, campo
'            If b Then
'                b = InsertLineaTrans(tipomov, CStr(numfactu), FecFac, AntVarie, AntCampo, CStr(Kilos), CStr(Importe), CStr(Gastos))
'            End If
'
'            If b Then
'                AntVarie = ActVarie
'                AntCampo = ActCampo
'                AntAlbar = ActAlbar
'                Kilos = 0
'                Importe = 0
'            End If
'        End If
        
        If ActTrans <> AntTrans Then
            '[Monica]15/10/2010: Añadido que se descuente el importe bruto pagado como trabajador solo para Picassent
            If vParamAplic.Cooperativa = 2 Or vParamAplic.Cooperativa = 16 Then
                SQL = "select codtraba from rtransporte where codtrans = " & DBSet(AntTrans, "T")
                CodTraba = DevuelveValor(SQL)
                
                SQL = "select if(isnull(sum(importe)),0,sum(importe)) from horas where codtraba = " & DBSet(CodTraba, "N")
                SQL = SQL & " and fechahora >= " & DBSet(Fdesde, "F")
                SQL = SQL & " and fechahora <= " & DBSet(Fhasta, "F")
                '[Monica]01/02/2013
                SQL = SQL & " and codvarie = 861 "
                ImpPagado = DevuelveValor(SQL)
        
                baseimpo = baseimpo - ImpPagado
                
                SQL = "select count(distinct fechahora) from horas where codtraba = " & DBSet(CodTraba, "N")
                SQL = SQL & " and fechahora >= " & DBSet(Fdesde, "F")
                SQL = SQL & " and fechahora <= " & DBSet(Fhasta, "F")
                '[Monica]01/02/2013
                SQL = SQL & " and codvarie = 861 "
                DiasTrab = DevuelveValor(SQL)
                
                baseimpo = baseimpo - Round2(DiasTrab * vParamAplic.EurosTrabdiaNOMI, 2)
            End If
        
            ImpoIva = Round2(baseimpo * PorcIva / 100, 2)
        
            Select Case DBLet(vTrans.TipoIRPF, "N")
                Case 0
                    ImpoReten = Round2((baseimpo + ImpoIva) * vParamAplic.PorcreteFacSoc / 100, 2)
                    BaseReten = (baseimpo + ImpoIva)
                    PorcReten = vParamAplic.PorcreteFacSoc
                Case 1
                    ImpoReten = Round2(baseimpo * vParamAplic.PorcreteFacSoc / 100, 2)
                    BaseReten = baseimpo
                    PorcReten = vParamAplic.PorcreteFacSoc
                Case 2
                    ImpoReten = 0
                    BaseReten = 0
                    PorcReten = 0
                Case 3 ' modulos en el regimen de transportista
                    ImpoReten = Round2(baseimpo * vParamAplic.PorcreteFacTra / 100, 2)
                    BaseReten = (baseimpo)
                    PorcReten = vParamAplic.PorcreteFacTra
            End Select
            
            TotalFac = baseimpo + ImpoIva - ImpoReten
            
            'insertar cabecera de factura
            B = InsertCabeceraTrans(tipoMov, CStr(numfactu), FecFac, vTrans)
            
            If B Then B = InsertResumen(tipoMov, CStr(numfactu), vTrans.Codigo)
            
            If B Then
                '[Monica]05/11/2012 si es una factura interna en Alzira cogemos el contador global, no el del transportista
                If vParamAplic.TipoContadorTRA = 0 Or (vParamAplic.Cooperativa = 4 And vTrans.EsFactTraInterna) Then
                    B = vTipoMov.IncrementarContador(tipoMov)
                Else
                    B = vTrans.IncrementarContador()
                End If
            End If
            
            If B Then
                AntTrans = ActTrans
                
                Set vTrans = Nothing
                Set vTrans = New CTransportista
                If vTrans.LeerDatos(ActTrans) Then
                    vPorcIva = ""
                    '[Monica]17/10/2011: FACTURAS INTERNAS
                    If vTrans.EsFactTraInterna Then
                        vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vParamAplic.CodIvaExeADV, "N")
                    Else
                        vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vTrans.CodIva, "N")
                    End If
                    If vPorcIva = "" Then
                        MsgBox "El transportista " & vTrans.Codigo & " no tiene iva. Revise.", vbExclamation
                        B = False
                        Exit Function
                    Else
                        PorcIva = CCur(ImporteSinFormato(vPorcIva))
                    End If
                Else
                    B = False
                End If
                
                baseimpo = 0
                BaseReten = 0
                ImpoIva = 0
                ImpoReten = 0
                TotalFac = 0
                
                
                If B Then
                    '[Monica]05/11/2012 si es una factura interna en Alzira cogemos el contador global, no el del transportista
                    If vParamAplic.TipoContadorTRA = 0 Or (vParamAplic.Cooperativa = 4 And vTrans.EsFactTraInterna) Then ' contador global
                        If vTipoMov Is Nothing Then Set vTipoMov = New CTiposMov
                        numfactu = vTipoMov.ConseguirContador(tipoMov)
                        Do
                            numfactu = vTipoMov.ConseguirContador(tipoMov)
                            devuelve = DevuelveDesdeBDNew(cAgro, "rfacttra", "numfactu", "codtipom", tipoMov, "T", , "numfactu", CStr(numfactu), "N", "fecfactu", FecFac, "F")
                            If devuelve <> "" Then
                                'Ya existe el contador incrementarlo
                                Existe = True
                                vTipoMov.IncrementarContador (tipoMov)
                                numfactu = vTipoMov.ConseguirContador(tipoMov)
                            Else
                                Existe = False
                            End If
                        Loop Until Not Existe
                    Else
                        numfactu = vTrans.ConseguirContador()
                        Do
                            numfactu = vTrans.ConseguirContador()
'                            devuelve = DevuelveDesdeBDNew(cAgro, "rfacttra", "numfactu", "codtipom", tipoMov, "T", , "numfactu", CStr(numfactu), "N", "fecfactu", FecFac, "F")
                            SQL = "select numfactu from rfacttra where codtipom = " & DBSet(tipoMov, "T") & " and numfactu = " & DBSet(numfactu, "N") & " and fecfactu = " & DBSet(FecFac, "F") & " and codtrans = " & DBSet(vTrans.Codigo, "T")
                            devuelve = DevuelveValor(SQL) 'DevuelveDesdeBDNew(cAgro, "rfacttra", "numfactu", "codtipom", tipoMov, "T", , "numfactu", CStr(numfactu), "N", "fecfactu", FecFac, "F")
                            If devuelve <> 0 Then
                                'Ya existe el contador incrementarlo
                                Existe = True
                                vTrans.IncrementarContador
                                numfactu = vTrans.ConseguirContador()
                            Else
                                Existe = False
                            End If
                        Loop Until Not Existe
                    End If
                End If
           
           End If
        End If
        
        If B Then
            B = InsertLineaTrans(tipoMov, CStr(numfactu), FecFac, Rs)
        End If
        
        IncrementarProgresNew Pb1, 1
        
        baseimpo = baseimpo + DBLet(Rs!Importe, "N")
        
        HayReg = True
        
        Rs.MoveNext
    Wend
    
    ' ultimo registro si ha entrado
    If B And HayReg Then
        '[Monica]15/10/2010: Añadido que se descuente el importe bruto pagado como trabajador solo para Picassent
        If vParamAplic.Cooperativa = 2 Or vParamAplic.Cooperativa = 16 Then
            SQL = "select codtraba from rtransporte where codtrans = " & DBSet(ActTrans, "T")
            CodTraba = DevuelveValor(SQL)
            
            SQL = "select if(isnull(sum(importe)),0,sum(importe)) from horas where codtraba = " & DBSet(CodTraba, "N")
            SQL = SQL & " and fechahora >= " & DBSet(Fdesde, "F")
            SQL = SQL & " and fechahora <= " & DBSet(Fhasta, "F")
            '[Monica]01/02/2013
            SQL = SQL & " and codvarie = 861 "
            ImpPagado = DevuelveValor(SQL)
    
            baseimpo = baseimpo - ImpPagado
            
            SQL = "select count(distinct fechahora) from horas where codtraba = " & DBSet(CodTraba, "N")
            SQL = SQL & " and fechahora >= " & DBSet(Fdesde, "F")
            SQL = SQL & " and fechahora <= " & DBSet(Fhasta, "F")
            '[Monica]01/02/2013
            SQL = SQL & " and codvarie = 861 "
            DiasTrab = DevuelveValor(SQL)
            
            baseimpo = baseimpo - Round2(DiasTrab * vParamAplic.EurosTrabdiaNOMI, 2)
        End If
        
        ' insertar linea de calidad
        ImpoIva = Round2(baseimpo * PorcIva / 100, 2)

        Select Case DBLet(vTrans.TipoIRPF, "N")
            Case 0
                ImpoReten = Round2((baseimpo + ImpoIva) * vParamAplic.PorcreteFacSoc / 100, 2)
                BaseReten = (baseimpo + ImpoIva)
                PorcReten = vParamAplic.PorcreteFacSoc
            Case 1
                ImpoReten = Round2(baseimpo * vParamAplic.PorcreteFacSoc / 100, 2)
                BaseReten = baseimpo
                PorcReten = vParamAplic.PorcreteFacSoc
            Case 2
                ImpoReten = 0
                BaseReten = 0
                PorcReten = 0
            Case 3 ' modulos en el regimen de transportista
                ImpoReten = Round2(baseimpo * vParamAplic.PorcreteFacTra / 100, 2)
                BaseReten = (baseimpo)
                PorcReten = vParamAplic.PorcreteFacTra
        End Select
        
        TotalFac = baseimpo + ImpoIva - ImpoReten
        
        IncrementarProgresNew Pb1, 1
        
        'insertar cabecera de factura
        B = InsertCabeceraTrans(tipoMov, CStr(numfactu), FecFac, vTrans)
        
        If B Then B = InsertResumen(tipoMov, CStr(numfactu), vTrans.Codigo)
        
        If B Then
            '[Monica]05/11/2012 si es una factura interna en Alzira cogemos el contador global, no el del transportista
            If vParamAplic.TipoContadorTRA = 0 Or (vParamAplic.Cooperativa = 4 And vTrans.EsFactTraInterna) Then
                B = vTipoMov.IncrementarContador(tipoMov)
            Else
                B = vTrans.IncrementarContador()
            End If
        End If
        
        'pasamos las temporales a las tablas
        If B Then B = PasarTemporalesTrans()
        
    End If
    
    Set vTrans = Nothing
    If Not vTipoMov Is Nothing Then Set vTipoMov = Nothing
    
EFacturacionTransporte:
    If Err.Number <> 0 Or Not B Then
        If Err.Number <> 0 Then MuestraError Err.Number, "Facturación Transporte:", Err.Description
        conn.RollbackTrans
        FacturacionTransporte = False
    Else
        conn.CommitTrans
        FacturacionTransporte = True
    End If
                
    Pb1.visible = False
    
End Function



'Insertar Cabecera de factura
Public Function InsertCabeceraTrans(tipoMov As String, numfactu As String, FecFac As String, vTrans As CTransportista) As Boolean

    Dim SQL As String
    
    On Error GoTo eInsertCabe
    
    MensError = ""
    InsertCabeceraTrans = False
    
    SQL = "insert into rfacttra (codtipom, numfactu, fecfactu, codtrans, baseimpo, tipoiva, porc_iva, "
    SQL = SQL & "imporiva, tipoirpf, basereten, porc_ret, impreten, baseaport, porc_apo, impapor, totalfac,"
    SQL = SQL & " impreso, contabilizado, pasaridoc, rectif_codtipom, rectif_numfactu, rectif_fecfactu, rectif_motivo) "
    SQL = SQL & " values ('" & tipoMov & "'," & DBSet(numfactu, "N") & "," & DBSet(FecFac, "F") & "," & DBSet(vTrans.Codigo, "T") & ","
    SQL = SQL & DBSet(baseimpo, "N") & "," & vTrans.CodIva & "," & DBSet(PorcIva, "N") & ","
    SQL = SQL & DBSet(ImpoIva, "N") & "," & DBSet(vTrans.TipoIRPF, "N") & "," & DBSet(BaseReten, "N") & ","
    SQL = SQL & DBSet(PorcReten, "N") & "," & DBSet(ImpoReten, "N") & "," & DBSet(BaseAFO, "N", "S") & "," & DBSet(PorcAFO, "N", "S") & "," & DBSet(ImpoAFO, "N", "S") & "," & DBSet(TotalFac, "N") & ","
    SQL = SQL & "0,0,0,"
    SQL = SQL & ValorNulo & ","
    SQL = SQL & ValorNulo & ","
    SQL = SQL & ValorNulo & ","
    SQL = SQL & ValorNulo
    SQL = SQL & ")"
    
    conn.Execute SQL
    
    InsertCabeceraTrans = True
    
    Exit Function

eInsertCabe:
    MensError = "Error en la inserción en rfacttra de la factura " & numfactu & " del transportista " & vTrans.Codigo
    MuestraError Err.Number, MensError
End Function




'Insertar Linea de factura (variedad)
Public Function InsertLineaTrans(tipoMov As String, numfactu As String, FecFac As String, ByRef Rs As ADODB.Recordset) As Boolean
Dim Precio As Currency

    Dim SQL As String
    Dim ImpLinea As Currency
    
    On Error GoTo eInsertLinea
    
    InsertLineaTrans = False
    
    MensError = ""
    Precio = 0
    If vParamAplic.Cooperativa = 2 Or vParamAplic.Cooperativa = 16 Then ' si es picassent los kilos son los transportados
        If CCur(ImporteSinFormato(Rs!KilosTra)) <> 0 Then
            Precio = Round2(CCur(ImporteSinFormato(DBLet(Rs!Importe, "N"))) / CCur(ImporteSinFormato(DBLet(Rs!KilosTra, "N"))), 4)
        End If
    Else
        If CCur(ImporteSinFormato(Rs!KilosNet)) <> 0 Then
            Precio = Round2(CCur(ImporteSinFormato(DBLet(Rs!Importe, "N"))) / CCur(ImporteSinFormato(DBLet(Rs!KilosNet, "N"))), 4)
        End If
    End If
    
    'numalbar,fecalbar,codvarie,codcampo,kilosbru,kilosnet,precio,importe
    
    SQL = "insert into tmpFact_albarantra (codtipom, numfactu, fecfactu, numalbar, fecalbar, codvarie, codcampo, "
    SQL = SQL & "kilosbru, kilosnet, precio, importe, codtrans, fechaent) values ("
    SQL = SQL & "'" & tipoMov & "'," & DBSet(numfactu, "N") & "," & DBSet(FecFac, "F") & ","
    SQL = SQL & DBSet(Rs!numalbar, "N") & "," & DBSet(Rs!FechaEnt, "F") & "," & DBSet(Rs!Codvarie, "N") & "," & DBSet(Rs!codCampo, "N") & ","
    If vParamAplic.Cooperativa = 2 Or vParamAplic.Cooperativa = 16 Then
        SQL = SQL & DBSet(DBLet(Rs!NumNotac, "N"), "N") & "," & DBSet(DBLet(Rs!KilosTra, "N"), "N") & ","
    Else
        SQL = SQL & DBSet(DBLet(Rs!NumNotac, "N"), "N") & "," & DBSet(DBLet(Rs!KilosNet, "N"), "N") & ","
    End If
    SQL = SQL & DBSet(Precio, "N") & ","
    SQL = SQL & DBSet(DBLet(Rs!Importe, "N"), "N") & ","
    SQL = SQL & DBSet(Rs!codTrans, "T") & ","
    '[Monica]21/05/2013: añadida la fecha de entrada
    SQL = SQL & DBSet(Rs!FechaEnt, "F") & ")"
    
    conn.Execute SQL

    InsertLineaTrans = True
    Exit Function
    
eInsertLinea:
    If Err.Number <> 0 Then
        MensError = "Se ha producido un error en la inserción de la linea de factura"
        MuestraError Err.Number, MensError, Err.Descripc
    End If
End Function



Public Function FacturacionLiquidacionesAlmazaraValsur(cTabla As String, cWhere As String, FecFac As String, Pb1 As ProgressBar, FIni As String, FFin As String) As Boolean
Dim SQL As String
Dim Rs As ADODB.Recordset
Dim RS1 As ADODB.Recordset

Dim AntSocio As String
Dim AntVarie As String
Dim ActSocio As String
Dim ActVarie As String
Dim actCampo As String
Dim AntCampo As String
Dim ActAlbar As String
Dim AntAlbar As String

Dim HayReg As Boolean

Dim NumError As Long
Dim vImporte As Currency
Dim vPorcIva As String

Dim PrimFac As String
Dim UltFac As String

Dim tipoMov As String
Dim B As Boolean
Dim vSeccion As CSeccion
Dim Kilos As Currency
Dim KilosCal As Currency
Dim Importe As Currency

Dim devuelve As String
Dim Existe As Boolean

'Dim baseimpo As Currency
'Dim BaseReten As Currency
'Dim Neto As Currency
'Dim ImpoIva As Currency
'Dim ImpoReten As Currency
'Dim TotalFac As Currency
Dim Recolect As Byte
Dim vPrecio As Currency

Dim Sql2 As String
Dim Sql3 As String
Dim Sql4 As String
Dim Sql5 As String


Dim GastosCoop As Currency
Dim GastosCoop2 As Currency
Dim GastosAlb As Currency
Dim vPorcGasto As String

Dim campo As String

Dim LitrosConsumidos As Currency
Dim LitrosProducidos As Currency
Dim PrecioConsumido As Currency
Dim PrecioProducido As Currency

Dim jj As Integer

' añadido
Dim ImporteRetirado As Currency
Dim ImporteMoltura As Currency
Dim ImporteMoltura1 As Currency  'Gastos de molturacion litros comercializados
Dim ImporteEnvasado As Currency
Dim PrecioRetirado As Currency
Dim KilosComer As Long
Dim KilosConsu As Long

Dim Rdto As Currency
Dim SqlGastos As String

Dim PrecioMoltura As Currency
Dim PrecioEnvasado As Currency

Dim Rs4 As ADODB.Recordset

Dim Litros As Currency
Dim cantidad As Currency



    On Error GoTo eFacturacion

    FacturacionLiquidacionesAlmazaraValsur = False
    
'08052009 antes dentro de transacciones
    BorrarTMPs
    B = CrearTMPs()
    If Not B Then
         Exit Function
    End If
'08052009
    
    conn.BeginTrans
    
    tipoMov = "FLZ"
    
    SQL = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute SQL
                                        '%%%%rhisfruta.codvarie
    SQL = "SELECT  rhisfruta.codsocio, variedades.codclase codvarie, rhisfruta.numalbar, "
    SQL = SQL & " rhisfruta.fecalbar,  rhisfruta.kilosbru, rhisfruta.prestimado,  "
    SQL = SQL & "rprecios_calidad.precoop, rprecios_calidad.presocio, rhisfruta.kilosnet"
    SQL = SQL & " FROM  " & cTabla

    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        SQL = SQL & " where " & cWhere
    End If
    ' ordenado por socio, variedad, campo, numlabar, fecalbar
    'Sql = Sql & " order by rhisfruta.codsocio, rhisfruta.codvarie, rhisfruta.numalbar, rhisfruta.fecalbar "
    '%%%%%%%%%
    SQL = SQL & " order by 1,2,3,4 "
    
    Set vSeccion = New CSeccion
    
    If vSeccion.LeerDatos(vParamAplic.SeccionAlmaz) Then
        If Not vSeccion.AbrirConta Then
            Exit Function
        End If
    End If
    
    HayReg = False
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    If Not Rs.EOF Then
        AntSocio = CStr(DBLet(Rs!Codsocio, "N"))
        AntVarie = CStr(DBLet(Rs!Codvarie, "N"))
        AntAlbar = CStr(DBLet(Rs!numalbar, "N"))
        
        ActSocio = CStr(DBLet(Rs!Codsocio, "N"))
        ActVarie = CStr(DBLet(Rs!Codvarie, "N"))
        ActAlbar = CStr(DBLet(Rs!numalbar, "N"))
'        Rdto = CStr(DBLet(Rs!PrEstimado, "N"))
        
        Set vSocio = New cSocio
        If vSocio.LeerDatos(ActSocio) Then
            If vSocio.LeerDatosSeccion(ActSocio, vParamAplic.SeccionAlmaz) Then
                baseimpo = 0
                BaseReten = 0
                ImpoIva = 0
                ImpoReten = 0
                TotalFac = 0
                
                Anticipos = 0
                
                Kilos = 0
                Importe = 0
                
                KilosComer = 0
                
                LitrosProducidos = 0
                
                vPorcIva = ""
                vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vSocio.CodIva, "N")
                PorcIva = CCur(ImporteSinFormato(vPorcIva))

'[Monica]14/04/2011: ahora lo vuelven a querer
'[Monica]07/04/2011: ahora no lo quieren en linea va en el precio
'                ' [Monica] 05/07/2010  añadido el gasto de cooperativa
                vPorcGasto = ""
                vPorcGasto = DevuelveDesdeBDNew(cAgro, "rcoope", "porcgast", "codcoope", vSocio.Cooperativa, "N")
                If vPorcGasto = "" Then vPorcGasto = "0"
                
                tipoMov = vSocio.CodTipomLiqAlmz
                
                Set vTipoMov = New CTiposMov
                
                numfactu = vTipoMov.ConseguirContador(tipoMov)
                Do
                    numfactu = vTipoMov.ConseguirContador(tipoMov)
                    devuelve = DevuelveDesdeBDNew(cAgro, "rfactsoc", "numfactu", "codtipom", tipoMov, "T", , "numfactu", CStr(numfactu), "N", "fecfactu", FecFac, "F")
                    If devuelve <> "" Then
                        'Ya existe el contador incrementarlo
                        Existe = True
                        vTipoMov.IncrementarContador (tipoMov)
                        numfactu = vTipoMov.ConseguirContador(tipoMov)
                    Else
                        Existe = False
                    End If
                Loop Until Not Existe
                
                vParamAplic.PrimFactLiqAlmz = numfactu
                
            End If
        End If
    End If
   
   ' en almazara no se insertan campos: metemos el minimo codcampo sin condiciones
    campo = "0" ' DevuelveValor("select min(codcampo) from rcampos")
    jj = 0
    
    While Not Rs.EOF And B
        ActVarie = DBLet(Rs!Codvarie, "N")
        ActSocio = DBSet(Rs!Codsocio, "N")
        
        
        If (ActVarie <> AntVarie Or ActSocio <> AntSocio) Then
            ' litros consumidos a otro precio

            Sql3 = "select rbodalbaran_variedad.*, variedades.* "
            Sql3 = Sql3 & " from rbodalbaran_variedad, rbodalbaran, variedades where rbodalbaran.codsocio = " & DBSet(AntSocio, "N")
            Sql3 = Sql3 & " and rbodalbaran.fechaalb >= " & DBSet(FIni, "F") & " and rbodalbaran.fechaalb <= " & DBSet(FFin, "F")
            Sql3 = Sql3 & " and variedades.codclase = " & DBSet(AntVarie, "N")
            Sql3 = Sql3 & " and rbodalbaran.numalbar = rbodalbaran_variedad.numalbar "
            Sql3 = Sql3 & " and rbodalbaran_variedad.codvarie = variedades.codvarie "
            Sql3 = Sql3 & " order by rbodalbaran_variedad.numalbar desc, rbodalbaran_variedad.numlinea desc"

            Sql4 = "select rbodalbaran_variedad.codvarie, variedades.eurdesta, variedades.eursegsoc, sum(cantidad) cantidad, round(variedades.eurdesta * sum(cantidad), 2) importevta, round(variedades.eursegsoc * sum(cantidad), 2) importeenv "
            Sql4 = Sql4 & " from rbodalbaran_variedad, rbodalbaran, variedades where rbodalbaran.codsocio = " & DBSet(AntSocio, "N")
            Sql4 = Sql4 & " and rbodalbaran.fechaalb >= " & DBSet(FIni, "F") & " and rbodalbaran.fechaalb <= " & DBSet(FFin, "F")
            Sql4 = Sql4 & " and variedades.codclase = " & DBSet(AntVarie, "N")
            Sql4 = Sql4 & " and rbodalbaran.numalbar = rbodalbaran_variedad.numalbar "
            Sql4 = Sql4 & " and rbodalbaran_variedad.codvarie = variedades.codvarie "
            Sql4 = Sql4 & " group by 1,2,3  order by 1,2, 3 "

            Sql5 = "select sum(cantidad) from (" & Sql4 & ") aaaaa"

            LitrosConsumidos = DevuelveValor(Sql5)

            
            If LitrosProducidos > LitrosConsumidos Then
                ' añadido
'                Sql4 = "select min(precioar) from rbodalbaran_variedad, rbodalbaran where rbodalbaran.codsocio = " & DBSet(AntSocio, "N")
'                Sql4 = Sql4 & " and rbodalbaran.fechaalb >= " & DBSet(FIni, "F") & " and rbodalbaran.fechaalb <= " & DBSet(FFin, "F")
'                Sql4 = Sql4 & " and rbodalbaran_variedad.codvarie = " & DBSet(AntVarie, "N")
'                Sql4 = Sql4 & " and rbodalbaran.numalbar = rbodalbaran_variedad.numalbar "
            
'                PrecioRetirado = DevuelveValor(Sql4)
                
                Rdto = Round2(LitrosProducidos * 100 / Kilos, 4)
                
                KilosComer = Round2((LitrosProducidos - LitrosConsumidos) * 100 / Rdto, 0)
                KilosConsu = Kilos - KilosComer
                
'                ImporteRetirado = Round2(LitrosConsumidos * PrecioRetirado, 2)
'                Sql5 = "select sum(importevta) from (" & Sql4 & ") aaaaa"
'                ImporteRetirado = DevuelveValor(Sql5)
                ImporteRetirado = CalculoImporteRetirado(Sql3, CStr(LitrosConsumidos), False)
                
                PrecioMoltura = DevuelveValor("select eurmanob from variedades where codvarie = " & DBSet(AntVarie, "N"))
                
                ImporteMoltura = Round2(KilosConsu * PrecioMoltura, 2)
                ImporteMoltura1 = Round2(KilosComer * PrecioMoltura, 2)
'                Sql5 = "select sum(importeenv) from (" & Sql4 & ") aaaaa"
'                ImporteEnvasado = DevuelveValor(Sql5) 'Round2(LitrosConsumidos * vParamAplic.GtoEnvasado, 2)
                ImporteEnvasado = CalculoImporteRetirado(Sql3, LitrosConsumidos, True)
                
                ' fañadido
            
                baseimpo = baseimpo + Round2((ImporteRetirado - ImporteMoltura - ImporteEnvasado) + ((LitrosProducidos - LitrosConsumidos) * PrecioProducido), 2) - ImporteMoltura1
                Importe = Round2((ImporteRetirado - ImporteMoltura - ImporteEnvasado) + ((LitrosProducidos - LitrosConsumidos) * PrecioProducido) - ImporteMoltura1, 2)
             
                jj = jj + 1
            
                
                GastosCoop = 0
                GastosCoop2 = 0
                
               ' [Monica] 05/07/2010 descontamos los gastos de la cooperativa en la linea
                If B Then ' descontamos el porcentaje de gastos de cooperativa
                    GastosCoop = 0
'                    GastosCoop = Round2((LitrosProducidos - LitrosConsumidos) * PrecioProducido * CCur(ImporteSinFormato(vPorcGasto)) / 100, 2)
'                    GastosCoop = Round2((LitrosConsumidos * PrecioConsumido) * CCur(ImporteSinFormato(vPorcGasto)) / 100, 2)
                    GastosCoop = Round2((ImporteRetirado) * CCur(ImporteSinFormato(vPorcGasto)) / 100, 2)
                    GastosCoop2 = 0
                    GastosCoop2 = Round2((ImporteRetirado + ((LitrosProducidos - LitrosConsumidos) * PrecioProducido)) * CCur(ImporteSinFormato(vPorcGasto)) / 100, 2)

                    baseimpo = baseimpo - GastosCoop2
                    Importe = Importe - GastosCoop2

                End If
            
                ' insertamos en las lineas de albaranes las lineas de litros consumidos a precioconsumido
                ' y la lina de litros producidos a precio producido
                SQL = "insert into tmpfact_albaran (codtipom, numfactu, fecfactu, numalbar, fecalbar, "
                SQL = SQL & "codvarie, codcampo, kilosbru, kilosnet, grado, precio, importe, imporgasto, "
                SQL = SQL & "prretirada, prmoltura, prenvasado, litrosconsu) values ("
                SQL = SQL & "'" & tipoMov & "'," & DBSet(numfactu, "N") & "," & DBSet(FecFac, "F") & ","
                SQL = SQL & DBSet(jj, "N") & "," & DBSet(FecFac, "F") & "," & DBSet(AntVarie, "N") & ","
'                SQL = SQL & "0,0," & DBSet(LitrosProducidos - LitrosConsumidos, "N") & ",0," & DBSet(PrecioProducido, "N") & ","
                SQL = SQL & "0," ' campo
                SQL = SQL & DBSet(KilosComer, "N") & "," ' en kilos bruto pongo los kilos
                SQL = SQL & DBSet(LitrosProducidos - LitrosConsumidos, "N") & "," & DBSet(GastosCoop2 - GastosCoop, "N") & "," & DBSet(PrecioProducido, "N") & ","
                SQL = SQL & DBSet(Round2(((LitrosProducidos - LitrosConsumidos) * PrecioProducido) - ImporteMoltura1, 2), "N") & ",1,0,"
                SQL = SQL & DBSet(PrecioMoltura, "N") & ",0," & DBSet(LitrosConsumidos, "N") & ")"
                
                conn.Execute SQL
                
                Litros = LitrosConsumidos
                
                Set Rs4 = New ADODB.Recordset
                Rs4.Open Sql4, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                
                
                While Not Rs4.EOF And Litros <> 0
                    jj = jj + 1
                    
                    If DBLet(Rs4!cantidad, "N") < Litros Then
                        Litros = Litros - DBLet(Rs4!cantidad, "N")
                        cantidad = DBLet(Rs4!cantidad, "N")
                    Else
                        cantidad = Litros
                        Litros = 0
                    End If
                
                    
                    SQL = "insert into tmpfact_albaran (codtipom, numfactu, fecfactu, numalbar, fecalbar, "
                    SQL = SQL & "codvarie, codcampo, kilosbru, kilosnet, grado, precio, importe, imporgasto, "
                    SQL = SQL & "prretirada, prmoltura, prenvasado, litrosconsu) values ("
                    SQL = SQL & "'" & tipoMov & "'," & DBSet(numfactu, "N") & "," & DBSet(FecFac, "F") & ","
                    SQL = SQL & DBSet(jj, "N") & "," & DBSet(FecFac, "F") & "," & DBSet(AntVarie, "N") & ","
    '                SQL = SQL & "0,0," & DBSet(LitrosConsumidos, "N") & ",0," & DBSet(PrecioConsumido, "N") & ","
                    SQL = SQL & "0," 'campo
                    SQL = SQL & DBSet(KilosConsu, "N") & "," ' kilos consumidos
                    SQL = SQL & DBSet(LitrosConsumidos, "N") & "," & DBSet(GastosCoop, "N") & "," & DBSet(PrecioConsumido, "N") & ","
                    SQL = SQL & DBSet((ImporteRetirado - ImporteMoltura - ImporteEnvasado), "N") & ",2,"
                    SQL = SQL & DBSet(Rs4!EurDesta, "N") & "," & DBSet(PrecioMoltura, "N") & "," & DBSet(Rs4!eursegsoc, "N") & ","
                    SQL = SQL & DBSet(cantidad, "N") & ")"
                    conn.Execute SQL
                    
                    Rs4.MoveNext
                Wend
                Set Rs4 = Nothing
            Else
                
                ' añadido
'                Sql4 = "select min(precioar) from rbodalbaran_variedad, rbodalbaran where rbodalbaran.codsocio = " & DBSet(AntSocio, "N")
'                Sql4 = Sql4 & " and rbodalbaran.fechaalb >= " & DBSet(FIni, "F") & " and rbodalbaran.fechaalb <= " & DBSet(FFin, "F")
'                Sql4 = Sql4 & " and rbodalbaran_variedad.codvarie = " & DBSet(AntVarie, "N")
'                Sql4 = Sql4 & " and rbodalbaran.numalbar = rbodalbaran_variedad.numalbar "
'
'                PrecioRetirado = DevuelveValor(Sql4)
'
'                Sql5 = "select eurdesta from (" & Sql4 & ") aaaaa"
'                PrecioRetirado = DevuelveValor(Sql5)
                
                
                Rdto = Round2(LitrosProducidos * 100 / Kilos, 4)
                
                KilosConsu = Round2(LitrosProducidos * 100 / Rdto, 0)
                
                
'               ImporteRetirado = Round2(LitrosProducidos * PrecioRetirado, 2)
                ImporteRetirado = CalculoImporteRetirado(Sql3, LitrosProducidos, False)
                PrecioRetirado = Round2(ImporteRetirado / LitrosProducidos, 4)
                
                PrecioMoltura = DevuelveValor("select eurmanob from variedades where codvarie = " & DBSet(AntVarie, "N"))
                
                ImporteMoltura = Round2(KilosConsu * PrecioMoltura, 2)
'                PrecioEnvasado = DevuelveValor("select eursegsoc from variedades where codvarie = " & DBSet(AntVarie, "N"))
'                ImporteEnvasado = Round2(LitrosProducidos * PrecioEnvasado, 2)
                ImporteEnvasado = CalculoImporteRetirado(Sql3, LitrosProducidos, True)
                PrecioEnvasado = Round2(ImporteEnvasado / LitrosProducidos, 4)
                ' fañadido

                baseimpo = baseimpo + Round2((ImporteRetirado - ImporteMoltura - ImporteEnvasado), 2)
                Importe = Round2((ImporteRetirado - ImporteMoltura - ImporteEnvasado), 2)
             
            
                jj = jj + 1
                
                GastosCoop = 0
                If B Then ' descontamos el porcentaje de gastos de cooperativa
                    GastosCoop = 0
                    GastosCoop = Round2(ImporteRetirado * CCur(ImporteSinFormato(vPorcGasto)) / 100, 2)
                    
                    baseimpo = baseimpo - GastosCoop
                    Importe = Importe - GastosCoop2
                End If
            
                Litros = LitrosProducidos
            
            
            
                Set Rs4 = New ADODB.Recordset
                Rs4.Open Sql4, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                
                While Not Rs4.EOF And Litros <> 0
                    jj = jj + 1
            
                    If DBLet(Rs4!cantidad, "N") < Litros Then
                        Litros = Litros - DBLet(Rs4!cantidad, "N")
                        cantidad = DBLet(Rs4!cantidad, "N")
                    Else
                        cantidad = Litros
                        Litros = 0
                    End If
            
            
                    ' insertamos en las lineas de albaranes las lineas de litros consumidos a precioconsumido
                    ' y la lina de litros producidos a precio producido
                    SQL = "insert into tmpfact_albaran (codtipom, numfactu, fecfactu, numalbar, fecalbar, "
                    SQL = SQL & "codvarie, codcampo, kilosbru, kilosnet, grado, precio, importe, imporgasto, "
                    SQL = SQL & "prretirada, prmoltura, prenvasado, litrosconsu) values ("
                    SQL = SQL & "'" & tipoMov & "'," & DBSet(numfactu, "N") & "," & DBSet(FecFac, "F") & ","
                    SQL = SQL & DBSet(jj, "N") & "," & DBSet(FecFac, "F") & "," & DBSet(AntVarie, "N") & ","
                    SQL = SQL & "0," 'campo
                    SQL = SQL & DBSet(KilosConsu, "N") & ","
                    SQL = SQL & DBSet(LitrosProducidos, "N") & "," & DBSet(GastosCoop, "N") & "," & DBSet(PrecioConsumido, "N") & ","
                    SQL = SQL & DBSet((ImporteRetirado - ImporteMoltura - ImporteEnvasado), "N") & ",2," '[Monica]28/03/2014: antes ",1,"
                    SQL = SQL & DBSet(Rs4!EurDesta, "N") & "," & DBSet(PrecioMoltura, "N") & "," & DBSet(Rs4!eursegsoc, "N") & ","
                    SQL = SQL & DBSet(cantidad, "N") & ")"
                    
                    conn.Execute SQL
                    
                    Rs4.MoveNext
               Wend
               Set Rs4 = Nothing
            End If
            
            
            If B Then ' descontamos los gastos de los albaranes
                'Para el resto sigue como estaba
                GastosAlb = ObtenerGastosAlbaranes(AntSocio, AntVarie, campo, cTabla, cWhere, 1, 1)
                Importe = Importe - GastosAlb
                baseimpo = baseimpo - GastosAlb
            End If
                        

                        
            SqlGastos = "select sum(grado) from tmpfact_albaran where codtipom = " & DBSet(tipoMov, "T") & " and numfactu = " & DBSet(numfactu, "N")
            SqlGastos = SqlGastos & " and fecfactu = " & DBSet(FecFac, "F") & " and codvarie = " & DBSet(AntVarie, "N")
            SqlGastos = SqlGastos & " and codcampo = " & DBSet(campo, "N")
            
            GastosAlb = GastosAlb + DevuelveValor(SqlGastos)
            
            ' insertar linea de variedad, campo
            If B Then
                B = InsertLinea(tipoMov, CStr(numfactu), FecFac, AntVarie, campo, CStr(Kilos), CStr(Importe), CStr(GastosAlb))
            End If
            
            If B Then
                ' tenemos que descontar los anticipos que tengamos para ello
                Sql2 = "select rfactsoc_variedad.imporvar, rfactsoc_variedad.numfactu, rfactsoc_variedad.fecfactu "
                Sql2 = Sql2 & " from rfactsoc_variedad INNER JOIN rfactsoc ON rfactsoc_variedad.codtipom = rfactsoc.codtipom and "
                Sql2 = Sql2 & " rfactsoc_variedad.numfactu = rfactsoc.numfactu and rfactsoc_variedad.fecfactu = rfactsoc.fecfactu "
                Sql2 = Sql2 & " where rfactsoc_variedad.codtipom = " & DBSet(vSocio.CodTipomAntAlmz, "T") ' antes era 'FAA' "
                Sql2 = Sql2 & " and rfactsoc.codsocio = " & DBSet(AntSocio, "N")
'                Sql2 = Sql2 & " and codcampo = " & DBSet(Campo, "N")
                Sql2 = Sql2 & " and codvarie = " & DBSet(AntVarie, "N")
                Sql2 = Sql2 & " and rfactsoc_variedad.descontado = 0"
                
                Set RS1 = New ADODB.Recordset
                RS1.Open Sql2, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
                
                While Not RS1.EOF
                    baseimpo = baseimpo - DBLet(RS1.Fields(0).Value, "N")
                    Anticipos = Anticipos + DBLet(RS1.Fields(0).Value, "N")
                    
                    ' indicamos que los anticipos ya han sido descontados
                    Sql3 = "update rfactsoc_variedad set descontado = 1 where codtipom = " & DBSet(vSocio.CodTipomAntAlmz, "T") ' antes era 'FAC'
                    Sql3 = Sql3 & " and numfactu = " & DBSet(RS1!numfactu, "N")
                    Sql3 = Sql3 & " and fecfactu = " & DBSet(RS1!fecfactu, "F") & " and codvarie = " & DBSet(AntVarie, "N")
'                    Sql3 = Sql3 & " and codcampo = " & DBSet(Campo, "N")
                    
                    conn.Execute Sql3
                    
                    ' insertamos en la tabla de anticipos de liquidacion
                    Sql3 = "insert into tmpFact_anticipos (codtipom, numfactu, fecfactu, codtipomanti, numfactuanti, fecfactuanti, codvarieanti, codcampoanti, baseimpo) values ("
                    Sql3 = Sql3 & DBSet(vSocio.CodTipomLiqAlmz, "T") & ","
                    Sql3 = Sql3 & DBSet(numfactu, "N") & "," & DBSet(FecFac, "F") & ","
                    Sql3 = Sql3 & DBSet(vSocio.CodTipomAntAlmz, "T") & ","
                    Sql3 = Sql3 & DBSet(RS1!numfactu, "N") & "," & DBSet(RS1!fecfactu, "F") & ","
                    Sql3 = Sql3 & DBSet(AntVarie, "N") & "," & DBSet(campo, "N") & "," & DBSet(RS1!imporvar, "N") & ")"
                    
                    conn.Execute Sql3
                    
                    RS1.MoveNext
                Wend
                
                Set RS1 = Nothing
                ' fin descontar anticipos
            
            End If
            
            
            If B Then
                AntVarie = ActVarie
                Kilos = 0
                Importe = 0
                LitrosProducidos = 0
'                Rdto = CStr(DBLet(Rs!PrEstimado, "N"))
                KilosComer = 0
            End If
        End If
        
        If ActSocio <> AntSocio Then
            
            ImpoIva = Round2(baseimpo * PorcIva / 100, 2)
        
            Select Case DBLet(vSocio.TipoIRPF, "N")
                Case 0
                    ImpoReten = Round2((baseimpo + ImpoIva) * vParamAplic.PorcreteFacSoc / 100, 2)
                    BaseReten = (baseimpo + ImpoIva)
                    PorcReten = vParamAplic.PorcreteFacSoc
                Case 1
                    ImpoReten = Round2(baseimpo * vParamAplic.PorcreteFacSoc / 100, 2)
                    BaseReten = baseimpo
                    PorcReten = vParamAplic.PorcreteFacSoc
                Case 2
                    ImpoReten = 0
                    BaseReten = 0
                    PorcReten = 0
            End Select
            
        
            TotalFac = baseimpo + ImpoIva - ImpoReten
            
            IncrementarProgresNew Pb1, 1
            
            'insertar cabecera de factura
            B = InsertCabecera(tipoMov, CStr(numfactu), FecFac)
            
            If B Then B = InsertResumen(tipoMov, CStr(numfactu))
            
            If B Then B = vTipoMov.IncrementarContador(tipoMov)
            
            If B Then
                AntSocio = ActSocio
                
                Set vSocio = Nothing
                Set vSocio = New cSocio
                If vSocio.LeerDatos(ActSocio) Then
                    If vSocio.LeerDatosSeccion(AntSocio, vParamAplic.SeccionAlmaz) Then
                        vPorcIva = ""
                        vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vSocio.CodIva, "N")
                        PorcIva = CCur(ImporteSinFormato(vPorcIva))
'[Monica]14/04/2011: ahora lo vuelven a querer
'[Monica]07/04/2011: ahora no lo quieren en linea va en el precio
                        ' [Monica] 05/07/2010  añadido el gasto de cooperativa
                        vPorcGasto = ""
                        vPorcGasto = DevuelveDesdeBDNew(cAgro, "rcoope", "porcgast", "codcoope", vSocio.Cooperativa, "N")
                        If vPorcGasto = "" Then vPorcGasto = "0"
                    
                    End If
                    
                    tipoMov = vSocio.CodTipomLiqAlmz
                End If
                baseimpo = 0
                BaseReten = 0
                ImpoIva = 0
                ImpoReten = 0
                TotalFac = 0
                
                Anticipos = 0
                
                jj = 0
                
                numfactu = vTipoMov.ConseguirContador(tipoMov)
                Do
                    numfactu = vTipoMov.ConseguirContador(tipoMov)
                    devuelve = DevuelveDesdeBDNew(cAgro, "rfactsoc", "numfactu", "codtipom", tipoMov, "T", , "numfactu", CStr(numfactu), "N", "fecfactu", FecFac, "F")
                    If devuelve <> "" Then
                        'Ya existe el contador incrementarlo
                        Existe = True
                        vTipoMov.IncrementarContador (tipoMov)
                        numfactu = vTipoMov.ConseguirContador(tipoMov)
                    Else
                        Existe = False
                    End If
                Loop Until Not Existe
           End If
        End If
        
        'vPrecio = DBLet(Rs!precioindustria, "N")
        'vImporte = Round2(DBLet(Rs!KilosNet, "N") * DBLet(Rs!PrEstimado, "N") / 100 * Rs!precioindustria, 2)
        
        LitrosProducidos = LitrosProducidos + Round2(DBLet(Rs!KilosNet, "N") * DBLet(Rs!PrEstimado, "N") / 100, 0)
        PrecioProducido = DBLet(Rs!PreSocio, "N")
        PrecioConsumido = DBLet(Rs!PreCoop, "N")
        
        vPrecio = PrecioProducido
        vImporte = Round2(LitrosProducidos * PrecioProducido, 2)
        
        
        KilosComer = KilosComer + DBLet(Rs!KilosNet, "N")
        
        
'[Monica]de momento no grabo los albaranes que intervienen
'        b = InsertLineaAlbaran(tipomov, CStr(numfactu), FecFac, Rs, CStr(vPrecio), CStr(vImporte), campo)
        
'        Importe = Importe + vImporte
'        BaseImpo = BaseImpo + vImporte
        Kilos = Kilos + DBLet(Rs!KilosNet)
        
        HayReg = True
        
        Rs.MoveNext
    Wend
    
    ' ultimo registro si ha entrado
    If B And HayReg Then
        Sql3 = "select rbodalbaran_variedad.*, variedades.* "
        Sql3 = Sql3 & " from rbodalbaran_variedad, rbodalbaran, variedades where rbodalbaran.codsocio = " & DBSet(AntSocio, "N")
        Sql3 = Sql3 & " and rbodalbaran.fechaalb >= " & DBSet(FIni, "F") & " and rbodalbaran.fechaalb <= " & DBSet(FFin, "F")
        Sql3 = Sql3 & " and variedades.codclase = " & DBSet(AntVarie, "N")
        Sql3 = Sql3 & " and rbodalbaran.numalbar = rbodalbaran_variedad.numalbar "
        Sql3 = Sql3 & " and rbodalbaran_variedad.codvarie = variedades.codvarie "
        Sql3 = Sql3 & " order by rbodalbaran_variedad.numalbar desc, rbodalbaran_variedad.numlinea desc"
        
        ' litros consumidos a otro precio
        Sql4 = "select rbodalbaran_variedad.codvarie, variedades.eurdesta, variedades.eursegsoc, sum(cantidad) cantidad, round(variedades.eurdesta * sum(cantidad), 2) importevta, round(variedades.eursegsoc * sum(cantidad), 2) importeenv "
        Sql4 = Sql4 & " from rbodalbaran_variedad, rbodalbaran, variedades where rbodalbaran.codsocio = " & DBSet(AntSocio, "N")
        Sql4 = Sql4 & " and rbodalbaran.fechaalb >= " & DBSet(FIni, "F") & " and rbodalbaran.fechaalb <= " & DBSet(FFin, "F")
        Sql4 = Sql4 & " and variedades.codclase = " & DBSet(AntVarie, "N")
        Sql4 = Sql4 & " and rbodalbaran.numalbar = rbodalbaran_variedad.numalbar "
        Sql4 = Sql4 & " and rbodalbaran_variedad.codvarie = variedades.codvarie "
        Sql4 = Sql4 & " group by 1,2,3  order by 1,2, 3 "
        
        
        Sql5 = "select sum(cantidad) from (" & Sql4 & ") aaaaa"
        
        LitrosConsumidos = DevuelveValor(Sql5)
        
        
        If LitrosProducidos > LitrosConsumidos Then
        
        
            ' añadido
'                Sql4 = "select min(precioar) from rbodalbaran_variedad, rbodalbaran where rbodalbaran.codsocio = " & DBSet(AntSocio, "N")
'                Sql4 = Sql4 & " and rbodalbaran.fechaalb >= " & DBSet(FIni, "F") & " and rbodalbaran.fechaalb <= " & DBSet(FFin, "F")
'                Sql4 = Sql4 & " and rbodalbaran_variedad.codvarie = " & DBSet(AntVarie, "N")
'                Sql4 = Sql4 & " and rbodalbaran.numalbar = rbodalbaran_variedad.numalbar "
        
'                PrecioRetirado = DevuelveValor(Sql4)
            
            Rdto = Round2(LitrosProducidos * 100 / Kilos, 4)
            
            KilosComer = Round2((LitrosProducidos - LitrosConsumidos) * 100 / Rdto, 0)
            KilosConsu = Kilos - KilosComer


'                ImporteRetirado = Round2(LitrosConsumidos * PrecioRetirado, 2)
'            Sql5 = "select sum(importevta) from (" & Sql4 & ") aaaaa"
'            ImporteRetirado = DevuelveValor(Sql5)
            ImporteRetirado = CalculoImporteRetirado(Sql3, LitrosConsumidos, False)
            
            PrecioMoltura = DevuelveValor("select eurmanob from variedades where codvarie = " & DBSet(AntVarie, "N"))
            
            ImporteMoltura = Round2(KilosConsu * PrecioMoltura, 2)
            ImporteMoltura1 = Round2(KilosComer * PrecioMoltura, 2)
'            Sql5 = "select sum(importeenv) from (" & Sql4 & ") aaaaa"
'            ImporteEnvasado = DevuelveValor(Sql5) 'Round2(LitrosConsumidos * vParamAplic.GtoEnvasado, 2)
            ImporteEnvasado = CalculoImporteRetirado(Sql3, LitrosConsumidos, True)
            
            ' fañadido
        
            baseimpo = baseimpo + Round2((ImporteRetirado - ImporteMoltura - ImporteEnvasado) + ((LitrosProducidos - LitrosConsumidos) * PrecioProducido), 2) - ImporteMoltura1
            Importe = Round2((ImporteRetirado - ImporteMoltura - ImporteEnvasado) + ((LitrosProducidos - LitrosConsumidos) * PrecioProducido) - ImporteMoltura1, 2)
         
            jj = jj + 1
        
            
            GastosCoop = 0
            GastosCoop2 = 0
            
           ' [Monica] 05/07/2010 descontamos los gastos de la cooperativa en la linea
            If B Then ' descontamos el porcentaje de gastos de cooperativa
                GastosCoop = 0
'                    GastosCoop = Round2((LitrosProducidos - LitrosConsumidos) * PrecioProducido * CCur(ImporteSinFormato(vPorcGasto)) / 100, 2)
'                    GastosCoop = Round2((LitrosConsumidos * PrecioConsumido) * CCur(ImporteSinFormato(vPorcGasto)) / 100, 2)
                GastosCoop = Round2((ImporteRetirado) * CCur(ImporteSinFormato(vPorcGasto)) / 100, 2)
                GastosCoop2 = 0
                GastosCoop2 = Round2((ImporteRetirado + ((LitrosProducidos - LitrosConsumidos) * PrecioProducido)) * CCur(ImporteSinFormato(vPorcGasto)) / 100, 2)

                baseimpo = baseimpo - GastosCoop2
                Importe = Importe - GastosCoop2

            End If
        
            ' insertamos en las lineas de albaranes las lineas de litros consumidos a precioconsumido
            ' y la lina de litros producidos a precio producido
            SQL = "insert into tmpfact_albaran (codtipom, numfactu, fecfactu, numalbar, fecalbar, "
            SQL = SQL & "codvarie, codcampo, kilosbru, kilosnet, grado, precio, importe, imporgasto, "
            SQL = SQL & "prretirada, prmoltura, prenvasado, litrosconsu) values ("
            SQL = SQL & "'" & tipoMov & "'," & DBSet(numfactu, "N") & "," & DBSet(FecFac, "F") & ","
            SQL = SQL & DBSet(jj, "N") & "," & DBSet(FecFac, "F") & "," & DBSet(AntVarie, "N") & ","
'                SQL = SQL & "0,0," & DBSet(LitrosProducidos - LitrosConsumidos, "N") & ",0," & DBSet(PrecioProducido, "N") & ","
            SQL = SQL & "0," ' campo
            SQL = SQL & DBSet(KilosComer, "N") & "," ' en kilos bruto pongo los kilos
            SQL = SQL & DBSet(LitrosProducidos - LitrosConsumidos, "N") & "," & DBSet(GastosCoop2 - GastosCoop, "N") & "," & DBSet(PrecioProducido, "N") & ","
            SQL = SQL & DBSet(Round2(((LitrosProducidos - LitrosConsumidos) * PrecioProducido) - ImporteMoltura1, 2), "N") & ",1,0,"
            SQL = SQL & DBSet(PrecioMoltura, "N") & ",0," & DBSet(LitrosConsumidos, "N") & ")"
            
            conn.Execute SQL
            
            
            Set Rs4 = New ADODB.Recordset
            Rs4.Open Sql4, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            
            
            Litros = LitrosConsumidos
            
            While Not Rs4.EOF And Litros <> 0
                jj = jj + 1
                
                If DBLet(Rs4!cantidad, "N") < Litros Then
                    Litros = Litros - DBLet(Rs4!cantidad, "N")
                    cantidad = DBLet(Rs4!cantidad, "N")
                Else
                    cantidad = Litros
                    Litros = 0
                End If
            
            
                SQL = "insert into tmpfact_albaran (codtipom, numfactu, fecfactu, numalbar, fecalbar, "
                SQL = SQL & "codvarie, codcampo, kilosbru, kilosnet, grado, precio, importe, imporgasto, "
                SQL = SQL & "prretirada, prmoltura, prenvasado, litrosconsu) values ("
                SQL = SQL & "'" & tipoMov & "'," & DBSet(numfactu, "N") & "," & DBSet(FecFac, "F") & ","
                SQL = SQL & DBSet(jj, "N") & "," & DBSet(FecFac, "F") & "," & DBSet(AntVarie, "N") & ","
'                SQL = SQL & "0,0," & DBSet(LitrosConsumidos, "N") & ",0," & DBSet(PrecioConsumido, "N") & ","
                SQL = SQL & "0," 'campo
                SQL = SQL & DBSet(KilosConsu, "N") & "," ' kilos consumidos
                SQL = SQL & DBSet(LitrosConsumidos, "N") & "," & DBSet(GastosCoop, "N") & "," & DBSet(PrecioConsumido, "N") & ","
                SQL = SQL & DBSet((ImporteRetirado - ImporteMoltura - ImporteEnvasado), "N") & ",2,"
                SQL = SQL & DBSet(Rs4!EurDesta, "N") & "," & DBSet(PrecioMoltura, "N") & "," & DBSet(Rs4!eursegsoc, "N") & ","
                SQL = SQL & DBSet(cantidad, "N") & ")"
                conn.Execute SQL
                
                Rs4.MoveNext
            Wend
            Set Rs4 = Nothing
        Else
            
            ' añadido
'                Sql4 = "select min(precioar) from rbodalbaran_variedad, rbodalbaran where rbodalbaran.codsocio = " & DBSet(AntSocio, "N")
'                Sql4 = Sql4 & " and rbodalbaran.fechaalb >= " & DBSet(FIni, "F") & " and rbodalbaran.fechaalb <= " & DBSet(FFin, "F")
'                Sql4 = Sql4 & " and rbodalbaran_variedad.codvarie = " & DBSet(AntVarie, "N")
'                Sql4 = Sql4 & " and rbodalbaran.numalbar = rbodalbaran_variedad.numalbar "
'
'                PrecioRetirado = DevuelveValor(Sql4)
'            Sql5 = "select eurdesta from (" & Sql4 & ") aaaaa"
'            PrecioRetirado = DevuelveValor(Sql5)
            
            
            Rdto = Round2(LitrosProducidos * 100 / Kilos, 4)
            
            KilosConsu = Round2(LitrosProducidos * 100 / Rdto, 0)
'            ImporteRetirado = Round2(LitrosProducidos * PrecioRetirado, 2)
            ImporteRetirado = CalculoImporteRetirado(Sql3, LitrosProducidos, False)
            PrecioRetirado = Round2(ImporteRetirado / LitrosProducidos, 4)

            
            PrecioMoltura = DevuelveValor("select eurmanob from variedades where codvarie = " & DBSet(AntVarie, "N"))
            ImporteMoltura = Round2(KilosConsu * PrecioMoltura, 2)
            
'            ImporteEnvasado = Round2(LitrosProducidos * vParamAplic.GtoEnvasado, 2)
'            PrecioEnvasado = DevuelveValor("select eursegsoc from variedades where codvarie = " & DBSet(AntVarie, "N"))
'            ImporteEnvasado = Round2(LitrosProducidos * PrecioEnvasado, 2)
            ImporteEnvasado = CalculoImporteRetirado(Sql3, LitrosProducidos, True)
            PrecioEnvasado = Round2(ImporteEnvasado / LitrosProducidos, 4)
           
            ' fañadido

            baseimpo = baseimpo + Round2((ImporteRetirado - ImporteMoltura - ImporteEnvasado), 2)
            Importe = Round2((ImporteRetirado - ImporteMoltura - ImporteEnvasado), 2)
         
        
            jj = jj + 1
            
            GastosCoop = 0
            If B Then ' descontamos el porcentaje de gastos de cooperativa
                GastosCoop = 0
                GastosCoop = Round2(ImporteRetirado * CCur(ImporteSinFormato(vPorcGasto)) / 100, 2)
                
                baseimpo = baseimpo - GastosCoop
                Importe = Importe - GastosCoop2
            End If
        
        
        
            Set Rs4 = New ADODB.Recordset
            Rs4.Open Sql4, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            
            
            Litros = LitrosProducidos
            
            While Not Rs4.EOF And Litros <> 0
                jj = jj + 1
                
                If DBLet(Rs4!cantidad, "N") < Litros Then
                    Litros = Litros - DBLet(Rs4!cantidad, "N")
                    cantidad = DBLet(Rs4!cantidad, "N")
                Else
                    cantidad = Litros
                    Litros = 0
                End If
        
                ' insertamos en las lineas de albaranes las lineas de litros consumidos a precioconsumido
                ' y la lina de litros producidos a precio producido
                SQL = "insert into tmpfact_albaran (codtipom, numfactu, fecfactu, numalbar, fecalbar, "
                SQL = SQL & "codvarie, codcampo, kilosbru, kilosnet, grado, precio, importe, imporgasto, "
                SQL = SQL & "prretirada, prmoltura, prenvasado, litrosconsu) values ("
                SQL = SQL & "'" & tipoMov & "'," & DBSet(numfactu, "N") & "," & DBSet(FecFac, "F") & ","
                SQL = SQL & DBSet(jj, "N") & "," & DBSet(FecFac, "F") & "," & DBSet(AntVarie, "N") & ","
                SQL = SQL & "0," 'campo
                SQL = SQL & DBSet(KilosConsu, "N") & ","
                SQL = SQL & DBSet(LitrosProducidos, "N") & "," & DBSet(GastosCoop, "N") & "," & DBSet(PrecioConsumido, "N") & ","
                SQL = SQL & DBSet((ImporteRetirado - ImporteMoltura - ImporteEnvasado), "N") & ",2," '[Monica]28/03/2014: antes ",1,"
                SQL = SQL & DBSet(Rs4!EurDesta, "N") & "," & DBSet(PrecioMoltura, "N") & "," & DBSet(Rs4!eursegsoc, "N") & ","
                SQL = SQL & DBSet(cantidad, "N") & ")"
                
                conn.Execute SQL
                
                Rs4.MoveNext
           Wend
           Set Rs4 = Nothing
           
        End If
        
        
        If B Then ' descontamos los gastos de los albaranes
            'Para el resto sigue como estaba
            GastosAlb = ObtenerGastosAlbaranes(AntSocio, AntVarie, campo, cTabla, cWhere, 1, 1)
            Importe = Importe - GastosAlb
            baseimpo = baseimpo - GastosAlb
        End If
    
    
        
        GastosCoop = 0
                    
        SqlGastos = "select sum(grado) from tmpfact_albaran where codtipom = " & DBSet(tipoMov, "T") & " and numfactu = " & DBSet(numfactu, "N")
        SqlGastos = SqlGastos & " and fecfactu = " & DBSet(FecFac, "F") & " and codvarie = " & DBSet(ActVarie, "N")
        SqlGastos = SqlGastos & " and codcampo = " & DBSet(campo, "N")
        
        GastosAlb = GastosAlb + DevuelveValor(SqlGastos)
                    
                    
        ' insertar linea de variedad
        If B Then B = InsertLinea(tipoMov, CStr(numfactu), FecFac, CStr(ActVarie), CStr(campo), CStr(Kilos), CStr(Importe), CStr(GastosAlb))
        
        ' tenemos que descontar los anticipos que tengamos para ello
        If B Then
            Sql2 = "select rfactsoc_variedad.imporvar, rfactsoc_variedad.numfactu, rfactsoc_variedad.fecfactu "
            Sql2 = Sql2 & " from rfactsoc_variedad INNER JOIN rfactsoc ON rfactsoc_variedad.codtipom = rfactsoc.codtipom and "
            Sql2 = Sql2 & " rfactsoc_variedad.numfactu = rfactsoc.numfactu and rfactsoc_variedad.fecfactu = rfactsoc.fecfactu "
            Sql2 = Sql2 & " where rfactsoc_variedad.codtipom = " & DBSet(vSocio.CodTipomAntAlmz, "T") ' antes era 'FAA' "
            Sql2 = Sql2 & " and rfactsoc.codsocio = " & DBSet(ActSocio, "N")
            Sql2 = Sql2 & " and codvarie = " & DBSet(ActVarie, "N")
'            Sql2 = Sql2 & " and codcampo = " & DBSet(Campo, "N")
            Sql2 = Sql2 & " and rfactsoc_variedad.descontado = 0"
            
            Set RS1 = New ADODB.Recordset
            RS1.Open Sql2, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
            
            While Not RS1.EOF
                baseimpo = baseimpo - DBLet(RS1.Fields(0).Value, "N")
                Anticipos = Anticipos + DBLet(RS1.Fields(0).Value, "N")
                
                ' indicamos que los anticipos ya han sido descontados
                Sql3 = "update rfactsoc_variedad set descontado = 1 where codtipom = " & DBSet(vSocio.CodTipomAntAlmz, "T") ' antes era 'FAC'
                Sql3 = Sql3 & " and numfactu = " & DBSet(RS1!numfactu, "N")
                Sql3 = Sql3 & " and fecfactu = " & DBSet(RS1!fecfactu, "F") & " and codvarie = " & DBSet(ActVarie, "N")
'                Sql3 = Sql3 & " and codcampo = " & DBSet(Campo, "N")
                
                conn.Execute Sql3
                
                ' insertamos en la tabla de anticipos de liquidacion venta campo
                Sql3 = "insert into tmpFact_anticipos (codtipom, numfactu, fecfactu, codtipomanti, numfactuanti, fecfactuanti, codvarieanti, codcampoanti, baseimpo) values ("
                Sql3 = Sql3 & DBSet(vSocio.CodTipomLiqAlmz, "T") & "," ' antes era 'FAL'
                Sql3 = Sql3 & DBSet(numfactu, "N") & "," & DBSet(FecFac, "F") & ","
                Sql3 = Sql3 & DBSet(vSocio.CodTipomAntAlmz, "T") & "," ' antes era 'FAA'
                Sql3 = Sql3 & DBSet(RS1!numfactu, "N") & "," & DBSet(RS1!fecfactu, "F") & ","
                Sql3 = Sql3 & DBSet(ActVarie, "N") & "," & DBSet(campo, "N") & "," & DBSet(RS1!imporvar, "N") & ")"
                
                conn.Execute Sql3
                
                RS1.MoveNext
            Wend
            
            Set RS1 = Nothing
            ' fin descontar anticipos
        
        End If
        
        ImpoIva = Round2(baseimpo * PorcIva / 100, 2)

        Select Case DBLet(vSocio.TipoIRPF, "N")
            Case 0
                ImpoReten = Round2((baseimpo + ImpoIva) * vParamAplic.PorcreteFacSoc / 100, 2)
                BaseReten = (baseimpo + ImpoIva)
                PorcReten = vParamAplic.PorcreteFacSoc
            Case 1
                ImpoReten = Round2(baseimpo * vParamAplic.PorcreteFacSoc / 100, 2)
                BaseReten = baseimpo
                PorcReten = vParamAplic.PorcreteFacSoc
            Case 2
                ImpoReten = 0
                BaseReten = 0
                PorcReten = 0
        End Select
        
        TotalFac = baseimpo + ImpoIva - ImpoReten
        
        IncrementarProgresNew Pb1, 1
        
        vParamAplic.UltFactLiqAlmz = numfactu
        
        'insertar cabecera de factura
        B = InsertCabecera(tipoMov, CStr(numfactu), FecFac)
        
        If B Then B = InsertResumen(tipoMov, CStr(numfactu))
        
        If B Then B = vTipoMov.IncrementarContador(tipoMov)
        
        'pasamos las temporales a las tablas
        If B Then B = PasarTemporales()
        
        If B Then B = (vParamAplic.Modificar = 1)
    End If
    
'    BorrarTMPs
    
    vSeccion.CerrarConta
    Set vSeccion = Nothing
    Set vSocio = Nothing
    
eFacturacion:
    If Err.Number <> 0 Or Not B Then
        conn.RollbackTrans
        FacturacionLiquidacionesAlmazaraValsur = False
    Else
        conn.CommitTrans
        FacturacionLiquidacionesAlmazaraValsur = True
    End If
End Function


Public Function CalculoImporteRetirado(SQL As String, LitrosProducidos As Currency, EsImpEnvasado As Boolean) As Currency

Dim Rs As ADODB.Recordset
Dim Litros As Currency
Dim Importe As Currency
Dim Precio As Currency
'                Sql3 = "select rbodalbaran_variedad.*, variedades.* "
'                Sql3 = Sql3 & " from rbodalbaran_variedad, rbodalbaran, variedades where rbodalbaran.codsocio = " & DBSet(AntSocio, "N")
'                Sql3 = Sql3 & " and rbodalbaran.fechaalb >= " & DBSet(FIni, "F") & " and rbodalbaran.fechaalb <= " & DBSet(FFin, "F")
'                Sql3 = Sql3 & " and variedades.codclase = " & DBSet(AntVarie, "N")
'                Sql3 = Sql3 & " and rbodalbaran.numalbar = rbodalbaran_variedad.numalbar "
'                Sql3 = Sql3 & " and rbodalbaran_variedad.codvarie = variedades.codvarie "
'                Sql3 = Sql3 & " order by rbodalbaran_variedad.numalbar desc, rbodalbaran_variedad.numlinea desc"

    Litros = LitrosProducidos
    Importe = 0
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not Rs.EOF And Litros <> 0
        If DBLet(Rs!ampliaci) <> "Regularización de Precios" Then
            If DBLet(Rs!cantidad, "N") <= Litros Then
                Litros = Litros - DBLet(Rs!cantidad, "N")
                If EsImpEnvasado Then
                    Precio = DBLet(Rs!eursegsoc, "N")
                    Importe = Importe + CalcularImporte(CStr(DBLet(Rs!cantidad, "N")), CStr(Precio), CStr(Rs!dtolinea), 0, 0, 0)
                Else
                    Importe = Importe + DBLet(Rs!ImporteL, "N")
                End If
            Else
                If EsImpEnvasado Then
                    Precio = DBLet(Rs!eursegsoc, "N")
                Else
                    Precio = DBLet(Rs!EurDesta, "N")
                End If
                Importe = Importe + CalcularImporte(CStr(DBLet(Litros, "N")), CStr(Precio), CStr(Rs!dtolinea), 0, 0, 0)
                Litros = 0
            End If
        End If
        Rs.MoveNext
    Wend
    Set Rs = Nothing
    CalculoImporteRetirado = Importe

End Function



'*****
'   proceso en donde se crea unicamente una factura de anticipo de vemta campo que posteriormente
'   se descontará en la factura de liquidacion de venta campo
'
Public Function FacturaAnticipoVentaCampo(Socio As String, campo As String, Importe As String, FecFac As String, EsTerceros As Boolean) As Long
Dim SQL As String
Dim Sql2 As String
Dim Rs As ADODB.Recordset
Dim RS1 As ADODB.Recordset

Dim AntSocio As String
Dim ActSocio As String

Dim HayReg As Boolean
Dim Variedad As String


Dim NumError As Long
Dim vImporte As Currency
Dim vPorcIva As String

Dim PrimFac As String
Dim UltFac As String
Dim tipoMov As String

Dim Sql3 As String
Dim devuelve As String
Dim B As Boolean
Dim vSeccion As CSeccion
Dim Existe As Boolean

    On Error GoTo eFacturacion
    
'08052009 antes dentro de transaccion
    BorrarTMPs
    B = CrearTMPs()
    If Not B Then
         Exit Function
    End If
'08052009

    conn.BeginTrans


    tipoMov = "FAC" ' Factura Anticipo Campo

    '[Monica]29/05/2017: si es tercero se le generará un informe
    If vParamAplic.Cooperativa = 2 Then
        If EsTerceros Then
            tipoMov = "CAT" ' Campo Anticipo Tercero
        Else
            tipoMov = "FAC" ' Factura Anticipo Campo
        End If
    End If
    
    SQL = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute SQL

    SQL = "SELECT rcampos.codvarie from rcampos where codcampo = " & DBSet(campo, "N")
    Variedad = DevuelveValor(SQL)
    
    Set vSeccion = New CSeccion
    
    If vSeccion.LeerDatos(vParamAplic.Seccionhorto) Then
        If Not vSeccion.AbrirConta Then
            Exit Function
        End If
    End If
    
    Set vSocio = New cSocio
    If vSocio.LeerDatos(Socio) Then
        If vSocio.LeerDatosSeccion(Socio, vParamAplic.Seccionhorto) Then
            baseimpo = CCur(Importe)
            BaseReten = 0
            ImpoIva = 0
            ImpoReten = 0
            TotalFac = 0
            
            Anticipos = 0
            
            vPorcIva = ""
            vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vSocio.CodIva, "N")
            PorcIva = CCur(ImporteSinFormato(vPorcIva))
            
            
            tipoMov = vSocio.CodTipomAntVC
            
            '[Monica]29/05/2017: si es picassent y terceros
            If vParamAplic.Cooperativa = 2 Then
                If EsTerceros Then tipoMov = "CAT"
            End If
            
            
            Set vTipoMov = New CTiposMov
            numfactu = vTipoMov.ConseguirContador(tipoMov)
            Do
                numfactu = vTipoMov.ConseguirContador(tipoMov)
                devuelve = DevuelveDesdeBDNew(cAgro, "rfactsoc", "numfactu", "codtipom", tipoMov, "T", , "numfactu", CStr(numfactu), "N", "fecfactu", FecFac, "F")
                If devuelve <> "" Then
                    'Ya existe el contador incrementarlo
                    Existe = True
                    vTipoMov.IncrementarContador (tipoMov)
                    numfactu = vTipoMov.ConseguirContador(tipoMov)
                Else
                    Existe = False
                End If
            Loop Until Not Existe
            
            vParamAplic.PrimFactAntVC = numfactu
            
            ImpoIva = Round2(baseimpo * PorcIva / 100, 2)
        
            Select Case DBLet(vSocio.TipoIRPF, "N")
                Case 0
                    ImpoReten = Round2((baseimpo + ImpoIva) * vParamAplic.PorcreteFacSoc / 100, 2)
                    BaseReten = (baseimpo + ImpoIva)
                    PorcReten = vParamAplic.PorcreteFacSoc
                Case 1
                    ImpoReten = Round2(baseimpo * vParamAplic.PorcreteFacSoc / 100, 2)
                    BaseReten = baseimpo
                    PorcReten = vParamAplic.PorcreteFacSoc
                Case 2
                    ImpoReten = 0
                    BaseReten = 0
                    PorcReten = 0
            End Select
        
            TotalFac = baseimpo + ImpoIva - ImpoReten
            
            ' insertar linea de variedad, campo
            B = InsertLinea(tipoMov, CStr(numfactu), FecFac, CStr(DBLet(Variedad, "N")), CStr(DBLet(campo, "N")), CStr(DBLet(0, "N")), CStr(DBLet(Importe, "N")), 0)
            
            If B Then
                ' insertamos los totales en la calidad venta campo de la variedad (rfactsoc_calidad)
                Sql2 = "select codcalid from rcalidad where codvarie = " & DBSet(Variedad, "N")
                Sql2 = Sql2 & " and tipcalid = 2 " ' calidad de venta campo
                
                Set RS1 = New ADODB.Recordset
                RS1.Open Sql2, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
                
                If Not RS1.EOF Then
                    B = InsertLineaCalidad(tipoMov, CStr(numfactu), FecFac, CStr(DBLet(Variedad, "N")), CStr(DBLet(campo, "N")), CStr(DBLet(RS1!codcalid, "N")), CStr(DBLet(0, "N")), CStr(DBLet(Importe, "N")))
                End If
                Set RS1 = Nothing
            End If
            
            'insertar cabecera de factura
            If B Then B = InsertCabecera(tipoMov, CStr(numfactu), FecFac)
            
            '[Monica]29/05/2017: en el caso de ser tercero se marca como contabililizada
            If vParamAplic.Cooperativa = 2 Then
                If EsTerceros Then
                   If B Then B = MarcarFactura(tipoMov, CStr(numfactu), FecFac, True)
                End If
            End If
            
            
            If B Then B = InsertResumen(tipoMov, CStr(numfactu))
            
            If B Then B = vTipoMov.IncrementarContador(tipoMov)
            
            vParamAplic.UltFactAntVC = numfactu
            
            'pasamos las temporales a las tablas
            If B Then B = PasarTemporales()
            
            If B Then B = (vParamAplic.Modificar = 1)
            
        End If
    
        BorrarTMPs
    End If
    
    vSeccion.CerrarConta
    Set vSeccion = Nothing
    Set vSocio = Nothing
    
eFacturacion:
    If Err.Number <> 0 Or Not B Then
        conn.RollbackTrans
        FacturaAnticipoVentaCampo = False
    Else
        conn.CommitTrans
        FacturaAnticipoVentaCampo = True
    End If
End Function



Public Function FacturacionTransporteSocio(cTabla As String, cWhere As String, ctabla1 As String, cwhere1 As String, FecFac As String, Pb1 As ProgressBar, Fdesde As String, Fhasta As String, Optional EsTercero As Boolean) As Boolean
Dim tipoMov As String

Dim AntSocio As String
Dim ActSocio As String

Dim AntAlbar As String
Dim ActAlbar As String

Dim AntVarie As String
Dim ActVarie As String

Dim AntCampo As String
Dim actCampo As String
Dim B As Boolean
Dim SQL As String
Dim Sql2 As String

Dim Rs As ADODB.Recordset
Dim HayReg As Boolean
Dim vImporte As Currency
Dim vPorcIva As String
Dim devuelve As String
Dim Existe As Boolean

Dim Nregs As Long

Dim CodTraba As String

Dim Importe As Currency
Dim Precio As Currency
Dim Kilos As Long

Dim KilosLin As Long
Dim ImporteLin As Currency

Dim GasAcarreo As Currency
Dim PrecAcarreo As Currency

Dim ImpPenal As Currency

Dim ImporteNota As Currency


On Error GoTo EFacturacionTransporteSocio

    FacturacionTransporteSocio = False

'08052009 antes dentro de transacciones
    BorrarTMPs
    B = CrearTMPs()
    If Not B Then
         Exit Function
    End If
'08052009
    
    conn.BeginTrans

    '[Monica]10/10/2013: distinguimos si es tercero o no solo para Picassent
    If EsTercero Then
        tipoMov = "FTT"
    Else
        tipoMov = "FTS"
    End If
    
    
    Sql2 = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute Sql2
    
    '[Monica]19/12/2013: se utiliza para en la factura de transporte socio en el caso de Alzira que el gasto de recolectar se
    '                    calcula con el precio de cada calidad
    Sql2 = "delete from tmpinfventas where codusu = " & vUsu.Codigo
    conn.Execute Sql2
    
    SQL = "select rclasifica.codsocio, rclasifica.codvarie, "
    SQL = SQL & "rclasifica.codcampo, rclasifica.numnotac, rclasifica.fechaent, rclasifica.transportadopor, rclasifica.recolect, rclasifica.codtarif, sum(if(isnull(rclasifica.kilosnet),0,rclasifica.kilosnet)) kilosnet, sum(if(isnull(rclasifica.impacarr),0,rclasifica.impacarr)) impacarr, sum(if(isnull(rclasifica.imprecol),0,rclasifica.imprecol)) imprecol, sum(if(isnull(rclasifica.kilostra),0,rclasifica.kilostra)) kilostra, sum(if(isnull(rclasifica.imppenal),0,rclasifica.imppenal)) imppenal, 0 tipo from " & cTabla
    If cWhere <> "" Then
        SQL = SQL & " where " & cWhere
    Else
        SQL = SQL & " where (1=1) "
    End If
    '[Monica]11/09/2017: añadida esta condicion de que la variedad no esté en las relacionadas
    SQL = SQL & " and not rclasifica.codvarie in (select codvarie1 from variedades_rel)"
    
    SQL = SQL & " group by 1, 2, 3, 4, 5, 6, 7, 8"
    SQL = SQL & " union "
    
    '[Monica]11/09/2017: añadida toda esta parte de la union
    SQL = SQL & "select rclasifica.codsocio, variedades_rel.codvarie, "
    SQL = SQL & "rclasifica.codcampo, rclasifica.numnotac, rclasifica.fechaent, rclasifica.transportadopor, rclasifica.recolect, rclasifica.codtarif, sum(if(isnull(rclasifica.kilosnet),0,rclasifica.kilosnet)) kilosnet, sum(if(isnull(rclasifica.impacarr),0,rclasifica.impacarr)) impacarr, sum(if(isnull(rclasifica.imprecol),0,rclasifica.imprecol)) imprecol, sum(if(isnull(rclasifica.kilostra),0,rclasifica.kilostra)) kilostra, sum(if(isnull(rclasifica.imppenal),0,rclasifica.imppenal)) imppenal, 0 tipo from (" & cTabla & ") "
    SQL = SQL & " INNER JOIN variedades_rel ON rclasifica.codvarie = variedades_rel.codvarie1 "
    If cWhere <> "" Then
        SQL = SQL & " where " & cWhere
    Else
        SQL = SQL & " where (1=1) "
    End If
    SQL = SQL & " and rclasifica.codvarie in (select codvarie1 from variedades_rel)"
    
    SQL = SQL & " group by 1, 2, 3, 4, 5, 6, 7, 8"
    SQL = SQL & " union "
    ' hasta aqui
    
    
    SQL = SQL & "select rhisfruta.codsocio, rhisfruta.codvarie, "
    SQL = SQL & "rhisfruta.codcampo, rhisfruta_entradas.numnotac, rhisfruta_entradas.fechaent, rhisfruta.transportadopor, rhisfruta.recolect, rhisfruta_entradas.codtarif, sum(if(isnull(rhisfruta_entradas.kilosnet),0,rhisfruta_entradas.kilosnet)) kilosnet, sum(if(isnull(rhisfruta_entradas.impacarr),0,rhisfruta_entradas.impacarr)) impacarr, sum(if(isnull(rhisfruta_entradas.imprecol),0,rhisfruta_entradas.imprecol)) imprecol, sum(if(isnull(rhisfruta_entradas.kilostra),0,rhisfruta_entradas.kilostra)) kilostra, sum(if(isnull(rhisfruta_entradas.imppenal),0,rhisfruta_entradas.imppenal)) imppenal, 1 tipo from " & ctabla1
    If cwhere1 <> "" Then
        SQL = SQL & " where " & cwhere1
    Else
        SQL = SQL & " where (1=1) "
    End If
    '[Monica]11/09/2017: añadida esta condicion de que la variedad no esté en las relacionadas
    SQL = SQL & " and not rhisfruta.codvarie in (select codvarie1 from variedades_rel)"
    
    SQL = SQL & " group by 1, 2, 3, 4, 5, 6, 7, 8"
    
    '[Monica]11/09/2017: añadida toda esta parte de la union
    SQL = SQL & " union "
    SQL = SQL & " select rhisfruta.codsocio, rhisfruta.codvarie, "
    SQL = SQL & "rhisfruta.codcampo, rhisfruta_entradas.numnotac, rhisfruta_entradas.fechaent, rhisfruta.transportadopor, rhisfruta.recolect, rhisfruta_entradas.codtarif, sum(if(isnull(rhisfruta_entradas.kilosnet),0,rhisfruta_entradas.kilosnet)) kilosnet, sum(if(isnull(rhisfruta_entradas.impacarr),0,rhisfruta_entradas.impacarr)) impacarr, sum(if(isnull(rhisfruta_entradas.imprecol),0,rhisfruta_entradas.imprecol)) imprecol, sum(if(isnull(rhisfruta_entradas.kilostra),0,rhisfruta_entradas.kilostra)) kilostra, sum(if(isnull(rhisfruta_entradas.imppenal),0,rhisfruta_entradas.imppenal)) imppenal, 1 tipo from (" & ctabla1 & ") "
    SQL = SQL & " INNER JOIN variedades_rel ON rhisfruta.codvarie = variedades_rel.codvarie1 "
    
    If cwhere1 <> "" Then
        SQL = SQL & " where " & cwhere1
    Else
        SQL = SQL & " where (1=1) "
    End If
    SQL = SQL & " and rhisfruta.codvarie in (select codvarie1 from variedades_rel)"
    ' hasta aqui
    
    
    SQL = SQL & " group by 1, 2, 3, 4, 5, 6, 7, 8"
    SQL = SQL & " order by 1, 2, 3, 4, 5, 6, 7, 8"
    
    
    Nregs = TotalRegistrosConsulta(SQL)
    Pb1.visible = True
    Pb1.Max = Nregs
    Pb1.Value = 0
    DoEvents
    
    HayReg = False
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    
    If Not Rs.EOF Then
        AntSocio = CStr(DBLet(Rs!Codsocio, "N"))
        AntAlbar = CStr(DBLet(Rs!NumNotac, "N"))
        AntVarie = CStr(DBLet(Rs!Codvarie, "N"))
        AntCampo = CStr(DBLet(Rs!codCampo, "N"))
        
        ActSocio = CStr(DBLet(Rs!Codsocio, "N"))
        ActAlbar = CStr(DBLet(Rs!NumNotac, "N"))
        ActVarie = CStr(DBLet(Rs!Codvarie, "N"))
        actCampo = CStr(DBLet(Rs!codCampo, "N"))
    
        Set vSocio = New cSocio
        If vSocio.LeerDatos(ActSocio) Then
            If vSocio.LeerDatosSeccion(ActSocio, vParamAplic.Seccionhorto) Then
                baseimpo = 0
                BaseReten = 0
                ImpoIva = 0
                ImpoReten = 0
                TotalFac = 0
                ImpoAFO = 0
                BaseAFO = 0
                
                KilosLin = 0
                ImporteLin = 0
                
                vPorcIva = ""
                vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vSocio.CodIva, "N")
                
                If vPorcIva = "" Then
                    MsgBox "El socio " & vSocio.Codigo & " no tiene iva. Revise.", vbExclamation
                    B = False
                Else
                    PorcIva = CCur(ImporteSinFormato(vPorcIva))
                End If
                
'                tipoMov = vSocio.CodTipomLiq
                
                If B Then
                    Set vTipoMov = New CTiposMov
                    
                    numfactu = vTipoMov.ConseguirContador(tipoMov)
                    Do
                        numfactu = vTipoMov.ConseguirContador(tipoMov)
                        devuelve = DevuelveDesdeBDNew(cAgro, "rfactsoc", "numfactu", "codtipom", tipoMov, "T", , "numfactu", CStr(numfactu), "N", "fecfactu", FecFac, "F")
                        If devuelve <> "" Then
                            'Ya existe el contador incrementarlo
                            Existe = True
                            vTipoMov.IncrementarContador (tipoMov)
                            numfactu = vTipoMov.ConseguirContador(tipoMov)
                        Else
                            Existe = False
                        End If
                    Loop Until Not Existe
                End If
            Else
                MsgBox "El socio " & ActSocio & " no se encuentra en la sección de Horto. Revise.", vbExclamation
                B = False
            End If
        End If
    End If
    
    While Not Rs.EOF And B
        ActSocio = DBLet(Rs!Codsocio, "N")
        ActAlbar = DBSet(Rs!NumNotac, "N")
        ActVarie = DBSet(Rs!Codvarie, "N")
        actCampo = DBSet(Rs!codCampo, "N")
        
        If ActSocio <> AntSocio Or ActVarie <> AntVarie Or actCampo <> AntCampo Then
            If B Then B = InsertLinea(tipoMov, CStr(numfactu), FecFac, CStr(DBLet(AntVarie, "N")), CStr(DBLet(AntCampo, "N")), CStr(KilosLin), CStr(ImporteLin), 0)
            
            AntAlbar = ActAlbar
            AntVarie = ActVarie
            AntCampo = actCampo
            
            KilosLin = 0
            ImporteLin = 0
        End If
        
        
        If ActSocio <> AntSocio Then
        
            ImpoIva = Round2(baseimpo * PorcIva / 100, 2)
        
            Select Case DBLet(vSocio.TipoIRPF, "N")
                Case 0
                    ImpoReten = Round2((baseimpo + ImpoIva) * vParamAplic.PorcreteFacSoc / 100, 2)
                    BaseReten = (baseimpo + ImpoIva)
                    PorcReten = vParamAplic.PorcreteFacSoc
                Case 1
                    ImpoReten = Round2(baseimpo * vParamAplic.PorcreteFacSoc / 100, 2)
                    BaseReten = baseimpo
                    PorcReten = vParamAplic.PorcreteFacSoc
                Case 2
                    ImpoReten = 0
                    BaseReten = 0
                    PorcReten = 0
            End Select
            
            TotalFac = baseimpo + ImpoIva - ImpoReten
            
            'insertar cabecera de factura
            B = InsertCabecera(tipoMov, CStr(numfactu), FecFac)
            
            '[Monica]10/10/2013: marcamos la factura como contabilizada y como pdte de recibir el nro de factura
            '                    solo si es Picassent y estamos facturando a socios terceros
            If B And ((vParamAplic.Cooperativa = 2 Or vParamAplic.Cooperativa = 16) And EsTercero) Then B = MarcarFactura(tipoMov, CStr(numfactu), FecFac, True)
            
            '[Monica]07/11/2013: marcamos la factura como contabilizada y como pdte de recibir el nro de factura
            If B And vSocio.EmiteFact And vParamAplic.Cooperativa = 4 Then B = MarcarFactura(tipoMov, CStr(numfactu), FecFac, True)
            
            If B Then B = InsertResumen(tipoMov, CStr(numfactu))
            
            If B Then
                B = vTipoMov.IncrementarContador(tipoMov)
            End If
            
            If B Then
                AntSocio = ActSocio
                
                Set vSocio = Nothing
                Set vSocio = New cSocio
                If vSocio.LeerDatos(ActSocio) Then
                    If vSocio.LeerDatosSeccion(ActSocio, vParamAplic.Seccionhorto) Then
                        vPorcIva = ""
                        vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vSocio.CodIva, "N")
                        If vPorcIva = "" Then
                            MsgBox "El socio " & vSocio.Codigo & " no tiene iva. Revise.", vbExclamation
                            B = False
                            Exit Function
                        Else
                            PorcIva = CCur(ImporteSinFormato(vPorcIva))
                        End If
                    Else
                        MsgBox "El socio " & ActSocio & " no se encuentra en la sección de Horto. Revise.", vbExclamation
                        B = False
                    End If
                End If
                
                baseimpo = 0
                BaseReten = 0
                ImpoIva = 0
                ImpoReten = 0
                TotalFac = 0
                
                
                If B Then
                    numfactu = vTipoMov.ConseguirContador(tipoMov)
                    Do
                        numfactu = vTipoMov.ConseguirContador(tipoMov)
                        devuelve = DevuelveDesdeBDNew(cAgro, "rfactsoc", "numfactu", "codtipom", tipoMov, "T", , "numfactu", CStr(numfactu), "N", "fecfactu", FecFac, "F")
                        If devuelve <> "" Then
                            'Ya existe el contador incrementarlo
                            Existe = True
                            vTipoMov.IncrementarContador (tipoMov)
                            numfactu = vTipoMov.ConseguirContador(tipoMov)
                        Else
                            Existe = False
                        End If
                    Loop Until Not Existe
                End If
           
           End If
        End If
        GasAcarreo = 0
        ImpPenal = 0
        If vParamAplic.Cooperativa = 2 Or vParamAplic.Cooperativa = 16 Then
            Importe = 0
            If DBLet(Rs!transportadopor, "N") = 1 Then Importe = Importe + DBLet(Rs!impacarr, "N")
            If DBLet(Rs!Recolect, "N") = 1 Then
                Importe = Importe + DBLet(Rs!imprecol, "N")
                If DBLet(Rs!ImpPenal, "N") <> 0 Then Importe = Importe - DBLet(Rs!ImpPenal, "N")
            End If
            
            Kilos = DBLet(Rs!KilosTra, "N")
            Precio = 0
            If Kilos <> 0 Then
                Precio = Round2(Importe / Kilos, 4)
            End If
            GasAcarreo = DBLet(Rs!impacarr, "N")
            ImpPenal = DBLet(Rs!ImpPenal, "N")
        Else
            If vParamAplic.Cooperativa = 4 Then
                Precio = DevuelveValor("select eurecole from variedades where codvarie = " & DBSet(Rs!Codvarie, "N"))
                Importe = 0
                If DBLet(Rs!transportadopor, "N") = 1 Then
                    PrecAcarreo = 0
                    SQL = ""
'                    If IsNull(Rs!codtarif) Then
'                        MsgBox "nota" & Rs!numnotac, vbExclamation
'                    End If
                    SQL = DevuelveDesdeBDNew(cAgro, "rtarifatra", "preciokg", "codtarif", DBLet(Rs!Codtarif, "N"), "N")
                    If SQL <> "" Then
                         PrecAcarreo = CCur(SQL)
                    End If
                    
                    GasAcarreo = Round2(DBLet(Rs!KilosTra, "N") * PrecAcarreo, 2)
                    
                    Importe = Importe + GasAcarreo
                End If
                
                If DBLet(Rs!Recolect, "N") = 1 Then
                    '[Monica]18/12/2013: antes para Alzira el gasto de recoleccion era kilos por un precio de la variedad
                    '                    ahora se va a calcular pro el precio de la calidad
                    'Importe = Importe + Round2(DBLet(Rs!KilosTra, "N") * Precio, 2)
                    ImporteNota = CalculoPorCalidad(CStr(Rs!NumNotac), Rs!Tipo)
                    
                    Importe = Importe + ImporteNota
                    
                End If
                
                Kilos = DBLet(Rs!KilosTra, "N")
            Else
                Precio = DevuelveValor("select eurecole from variedades where codvarie = " & DBSet(Rs!Codvarie, "N"))
                Importe = 0
                If DBLet(Rs!transportadopor, "N") = 1 Then Importe = Importe + DBLet(Rs!impacarr, "N")
                If DBLet(Rs!Recolect, "N") = 1 Then Importe = Importe + Round2(DBLet(Rs!KilosNet, "N") * Precio, 2)
                Kilos = DBLet(Rs!KilosNet, "N")
                GasAcarreo = DBLet(Rs!impacarr, "N")
            End If
        End If
        
        If B Then
            B = InsertLineaNota(tipoMov, CStr(numfactu), FecFac, Rs, CStr(Kilos), CStr(Precio), CStr(Importe), CStr(GasAcarreo), CStr(ImpPenal))
        End If
            
        ImporteLin = ImporteLin + Importe
        KilosLin = KilosLin + Kilos
        
        IncrementarProgresNew Pb1, 1
        
        baseimpo = baseimpo + DBLet(Importe, "N")
        
        HayReg = True
        
        Rs.MoveNext
    Wend
    
    ' ultimo registro si ha entrado
    If B And HayReg Then
        If B Then B = InsertLinea(tipoMov, CStr(numfactu), FecFac, CStr(DBLet(ActVarie, "N")), CStr(DBLet(actCampo, "N")), CStr(KilosLin), CStr(ImporteLin), 0)
        
        ImpoIva = Round2(baseimpo * PorcIva / 100, 2)

        Select Case DBLet(vSocio.TipoIRPF, "N")
            Case 0
                ImpoReten = Round2((baseimpo + ImpoIva) * vParamAplic.PorcreteFacSoc / 100, 2)
                BaseReten = (baseimpo + ImpoIva)
                PorcReten = vParamAplic.PorcreteFacSoc
            Case 1
                ImpoReten = Round2(baseimpo * vParamAplic.PorcreteFacSoc / 100, 2)
                BaseReten = baseimpo
                PorcReten = vParamAplic.PorcreteFacSoc
            Case 2
                ImpoReten = 0
                BaseReten = 0
                PorcReten = 0
        End Select
        
        TotalFac = baseimpo + ImpoIva - ImpoReten
        
        IncrementarProgresNew Pb1, 1
        
        'insertar cabecera de factura
        B = InsertCabecera(tipoMov, CStr(numfactu), FecFac)
        
        '[Monica]10/10/2013: marcamos la factura como contabilizada y como pdte de recibir el nro de factura
        '                    solo si es Picassent y estamos facturando a socios terceros
        If B And ((vParamAplic.Cooperativa = 2 Or vParamAplic.Cooperativa = 16) And EsTercero) Then B = MarcarFactura(tipoMov, CStr(numfactu), FecFac, True)
        
        '[Monica]07/11/2013: marcamos la factura como contabilizada y como pdte de recibir el nro de factura
        If B And vSocio.EmiteFact And vParamAplic.Cooperativa = 4 Then B = MarcarFactura(tipoMov, CStr(numfactu), FecFac, True)
        
        If B Then B = InsertResumen(tipoMov, CStr(numfactu))
        
        If B Then
             B = vTipoMov.IncrementarContador(tipoMov)
        End If
        
        'pasamos las temporales a las tablas
        If B Then B = PasarTemporales()
        
    End If
    
    Set vSocio = Nothing
    
EFacturacionTransporteSocio:
    If Err.Number <> 0 Or Not B Then
        If Err.Number <> 0 Then MuestraError Err.Number, "Facturación Transporte/Recoleccion a Socio:", Err.Description
        conn.RollbackTrans
        FacturacionTransporteSocio = False
    Else
        conn.CommitTrans
        FacturacionTransporteSocio = True
    End If
                
    Pb1.visible = False
    
End Function

'[Monica]18/12/2013: Calculo por calidad
Private Function CalculoPorCalidad(Nota As String, Tipo As Byte) As Currency
Dim SQL As String
Dim Importe As Currency
Dim Albaran As Long
Dim Precio As Currency
Dim GastosTot As Currency
Dim KilosTot As Long
Dim KilosNota As Long

    On Error GoTo eCalculoPorCalidad
    
    CalculoPorCalidad = 0
    
    Importe = 0
    If Tipo = 0 Then ' viene de rclasifica
        SQL = "select * from rclasifica_clasif where numnotac = " & DBSet(Nota, "N")
        
        Set Rs = New ADODB.Recordset
        Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not Rs.EOF
            Precio = ObtenerPrecioRecoldeCalidad(CStr(Rs!Codvarie), CStr(Rs!codcalid), 0)
            Importe = Importe + Round2(Precio * Rs!KilosNet, 2)
            
            Rs.MoveNext
        Wend
        Set Rs = Nothing
    Else ' viene de rhisfruta
        ' si el albaran entero ya esta procesado no hacemos nada
        Albaran = DevuelveValor("select numalbar from rhisfruta_entradas where numnotac = " & DBSet(Nota, "N"))
        SQL = "select * from tmpinfventas where codusu = " & vUsu.Codigo & " and numalbar = " & DBSet(Albaran, "N")
        If TotalRegistrosConsulta(SQL) = 0 Then
            SQL = "select * from rhisfruta_clasif where numalbar = " & DBSet(Albaran, "N")
        
            Set Rs = New ADODB.Recordset
            Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            While Not Rs.EOF
                Precio = ObtenerPrecioRecoldeCalidad(CStr(Rs!Codvarie), CStr(Rs!codcalid), 0)
                Importe = Importe + Round2(Precio * DBLet(Rs!KilosNet, "N"), 2)
            
                Rs.MoveNext
            Wend
            Set Rs = Nothing
            
            SQL = "insert into tmpinfventas (codusu,numalbar,gastos1) values (" & vUsu.Codigo & "," & DBSet(Albaran, "N") & "," & DBSet(Importe, "N") & ")"
            conn.Execute SQL
                    
            GastosTot = Importe
        Else
            SQL = "select gastos1 from tmpinfventas where codusu = " & vUsu.Codigo & " and numalbar = " & DBSet(Albaran, "N")
            GastosTot = DevuelveValor(SQL)
        
        End If
        ' prorrateo pq me pueden llegar notas de entrada que esten en la clasificacion con lo que no tendré el nro de albaran
        
        SQL = "select kilosnet from rhisfruta_entradas where numnotac = " & DBSet(Nota, "N")
        KilosNota = DevuelveValor(SQL)
        
        SQL = "select kilosnet from rhisfruta where numalbar = " & DBSet(Albaran, "N")
        KilosTot = DevuelveValor(SQL)
            
        Importe = Round2(GastosTot * KilosNota / KilosTot, 2)
            
    End If
    
    CalculoPorCalidad = Importe
    
    Exit Function
eCalculoPorCalidad:
    MuestraError Err.Number, "Calculo por Calidad", Err.Description
End Function

' Funcion que almacena en rfactsoc_albaran las notas con el importe de acarreo + importe de recoleccion
' si lo tiene en la factura de gastos de transporte y acarreo socio FTS

'Insertar Linea de factura (albaran)
Public Function InsertLineaNota(tipoMov As String, numfactu As String, FecFac As String, ByRef Rs As ADODB.Recordset, Kilos As String, Precio As String, Importe As String, GasAcarreo As String, vImppenal As String) As Boolean
'(rfactsoc_albaran)
'codcampo tiene valor cuando venimos de almazara que hemos tenido que buscarlo porque en el cursor Rs no lo tenemos
Dim GastosAlb As Currency
Dim Tipo As String

    Dim SQL As String
    Dim ImpLinea As Currency
    
    On Error GoTo eInsertLinea
    
    MensError = ""
    
    InsertLineaNota = False
    
    
    'insertamos el albaran
    SQL = "insert into tmpfact_albaran (codtipom, numfactu, fecfactu, numalbar, fecalbar, "
    SQL = SQL & "codvarie, codcampo, kilosbru, kilosnet, grado, precio, importe, imporgasto, imppenal) values ("
    SQL = SQL & "'" & tipoMov & "'," & DBSet(numfactu, "N") & "," & DBSet(FecFac, "F") & ","
    SQL = SQL & DBSet(Rs!NumNotac, "N") & "," & DBSet(Rs!FechaEnt, "F") & "," & DBSet(Rs!Codvarie, "N") & ","
    SQL = SQL & DBSet(Rs!codCampo, "N") & ","
    SQL = SQL & DBSet(Kilos, "N") & "," & DBSet(Kilos, "N") & ","
    SQL = SQL & DBSet(0, "N") & "," & DBSet(Precio, "N") & "," & DBSet(Importe, "N") & ","
    '[Monica]19/11/2010: En la columna de gastos metemos los gastos de acarreo, en importe tenemos acarreo + recoleccion
    SQL = SQL & DBSet(GasAcarreo, "N") & ","
    '[Monica]10/10/2013: metemos los gastos de imppenal
    SQL = SQL & DBSet(vImppenal, "N") & ")"
    
    conn.Execute SQL
    InsertLineaNota = True
    Exit Function
    
eInsertLinea:
    If Err.Number <> 0 Then
        MensError = "Se ha producido un error en la inserción de la linea de albaran de factura "
        MuestraError Err.Number, MensError, Err.Description
    End If
End Function



'Public Function InsertLineaAlbaranBodega(tipoMov As String, numfactu As String, FecFac As String, Variedad As String, Campo As String, Albaran As String, Kilosnet As String, KiloGrado As String, Importe As String) As Boolean
''(rfactsoc_albaran)
''codcampo tiene valor cuando venimos de almazara que hemos tenido que buscarlo porque en el cursor Rs no lo tenemos
'Dim GastosAlb As Currency
'Dim Tipo As String
'Dim Precio As Currency
'Dim Grado As Currency
'
'    Dim SQL As String
'    Dim ImpLinea As Currency
'
'    On Error GoTo eInsertLinea
'
'    MensError = ""
'
'    InsertLineaAlbaranBodega = False
'
'    Precio = 0
'    If CLng(KiloGrado) <> 0 Then Precio = Round2(Importe / KiloGrado, 4)
'
'    Grado = 0
'    If CLng(Kilosnet) <> 0 Then Grado = Round2(KiloGrado / Kilosnet, 2)
'
'    'insertamos el albaran
'    SQL = "insert into tmpfact_albaran (codtipom, numfactu, fecfactu, numalbar, fecalbar, "
'    SQL = SQL & "codvarie, codcampo, kilosbru, kilosnet, grado, precio, importe, imporgasto) values ("
'    SQL = SQL & "'" & tipoMov & "'," & DBSet(numfactu, "N") & "," & DBSet(FecFac, "F") & ","
'    SQL = SQL & DBSet(Albaran, "N") & "," & DBSet(FecAlbar, "F") & "," & DBSet(Variedad, "N") & ","
'    SQL = SQL & DBSet(Campo, "N") & ",0,"
'    SQL = SQL & DBSet(Kilosnet, "N") & "," & DBSet(Grado, "N") & ","
'    SQL = SQL & DBSet(Precio, "N") & "," & DBSet(Importe, "N") & ","
'    SQL = SQL & DBSet(0, "N") & ")"
'
'    conn.Execute SQL
'    InsertLineaAlbaranBodega = True
'    Exit Function
'
'eInsertLinea:
'    If Err.Number <> 0 Then
'        MensError = "Se ha producido un error en la inserción de la linea de albaran de factura "
'        MuestraError Err.Number, MensError, Err.Description
'    End If
'End Function



Public Function InsertLineaAlbaranBodega(tipoMov As String, numfactu As String, FecFac As String, Socio As String, Variedad As String, Tabla1 As String, Where1 As String) As Boolean
'(rfactsoc_albaran)
'codcampo tiene valor cuando venimos de almazara que hemos tenido que buscarlo porque en el cursor Rs no lo tenemos
Dim GastosAlb As Currency
Dim Tipo As String
Dim Sql2 As String
Dim Rs2 As ADODB.Recordset

Dim Importe As Currency

Dim SQL As String
Dim ImpLinea As Currency
Dim SqlValues As String
    
    On Error GoTo eInsertLinea
    
    MensError = ""
    
    InsertLineaAlbaranBodega = False
    
    GastosAlb = 0
    
    Sql2 = "SELECT  rhisfruta.codsocio, rhisfruta.codvarie, rhisfruta.numalbar, rhisfruta.fecalbar, rhisfruta.codcampo, "
    Sql2 = Sql2 & " rprecios.precioindustria, "
    Sql2 = Sql2 & "rprecios.tipofact, kilosnet , kgradobonif as prestimado  "
    Sql2 = Sql2 & " FROM  " & Tabla1

    If Where1 <> "" Then
        Where1 = QuitarCaracterACadena(Where1, "{")
        Where1 = QuitarCaracterACadena(Where1, "}")
        Where1 = QuitarCaracterACadena(Where1, "_1")
        Sql2 = Sql2 & " WHERE " & Where1
    End If
    
    Sql2 = Sql2 & " and rhisfruta.codsocio = " & DBSet(Socio, "N")
    Sql2 = Sql2 & " and rhisfruta.codvarie = " & DBSet(Variedad, "N")
    
    ' ordenado por socio, variedad, campo, calidad
    Sql2 = Sql2 & " group by rhisfruta.codsocio, rhisfruta.codvarie, rhisfruta.numalbar, rhisfruta.fecalbar, rhisfruta.codcampo, rprecios.precioindustria, rprecios.tipofact"
    Sql2 = Sql2 & " order by rhisfruta.codsocio, rhisfruta.codvarie, rhisfruta.numalbar, rhisfruta.fecalbar, rhisfruta.codcampo, rprecios.precioindustria, rprecios.tipofact"
    
    Set Rs2 = New ADODB.Recordset
    Rs2.Open Sql2, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    
    'insertamos el albaran
    SQL = "insert into tmpfact_albaran (codtipom, numfactu, fecfactu, numalbar, fecalbar, "
    SQL = SQL & "codvarie, codcampo, kilosbru, kilosnet, grado, precio, importe, imporgasto) values "
    
    SqlValues = ""
    
    While Not Rs2.EOF
    
        Importe = Round2(DBLet(Rs2!KilosNet, "N") * DBLet(Rs2!PrEstimado, "N") * DBLet(Rs2!precioindustria, "N"), 2)
    
        SqlValues = SqlValues & "('" & tipoMov & "'," & DBSet(numfactu, "N") & "," & DBSet(FecFac, "F") & ","
        SqlValues = SqlValues & DBSet(Rs2!numalbar, "N") & "," & DBSet(Rs2!Fecalbar, "F") & "," & DBSet(Variedad, "N") & ","
        SqlValues = SqlValues & DBSet(Rs2!codCampo, "N") & ",0," & DBSet(Rs2!KilosNet, "N") & "," & DBSet(Rs2!PrEstimado, "N") & ","
        SqlValues = SqlValues & DBSet(Rs2!precioindustria, "N") & "," & DBSet(Importe, "N") & ",0),"
    
        Rs2.MoveNext
        
    Wend
    
    Set Rs2 = Nothing
    
    If SqlValues <> "" Then
        conn.Execute SQL & Mid(SqlValues, 1, Len(SqlValues) - 1)
    End If
    
    InsertLineaAlbaranBodega = True
    Exit Function
    
eInsertLinea:
    If Err.Number <> 0 Then
        MensError = "Se ha producido un error en la inserción de la linea de albaran de factura bodega"
        MuestraError Err.Number, MensError, Err.Description
    End If
End Function


' Procedimiento que carga el campo de rhisfruta.kgradobonif
Public Function CalcularGradoBonificado(Tabla1 As String, Where1 As String, ByRef Pb1 As ProgressBar)
Dim SQL As String
Dim Rs As ADODB.Recordset
Dim RS1 As ADODB.Recordset
Dim Rs2 As ADODB.Recordset
Dim SQL1 As String
Dim Sql2 As String
Dim Porcen As Currency
Dim Grado As Currency

    On Error GoTo eCalcularGradoBonificado

    conn.BeginTrans

    CalcularGradoBonificado = False
    
    
    SQL = "Select rhisfruta.* FROM " & QuitarCaracterACadena(Tabla1, "_1")
    If Where1 <> "" Then
        Where1 = QuitarCaracterACadena(Where1, "_1")
        SQL = SQL & " WHERE " & Where1
    End If
    
    Pb1.Max = TotalRegistrosConsulta(SQL)
    Pb1.visible = True
    Pb1.Value = 0
    
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
    While Not Rs.EOF
        IncrementarProgresNew Pb1, 1
        DoEvents
        
        SQL1 = "select porcentaje from rbonifica_lineas where codvarie = " & DBSet(Rs!Codvarie, "N")
        SQL1 = SQL1 & " and desdegrado <= " & DBSet(Rs!PrEstimado, "N")
        SQL1 = SQL1 & " and " & DBSet(Rs!PrEstimado, "N") & " <= hastagrado "
        
        Set RS1 = New ADODB.Recordset
        RS1.Open SQL1, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        If Not RS1.EOF Then
            Porcen = DBLet(RS1.Fields(0).Value, "N")
            Grado = DBLet(Rs!PrEstimado, "N")
        Else
            'cogemos el registro con el hasta mayor para coger el porcentaje
            Porcen = 0
            Grado = DBLet(Rs!PrEstimado, "N")
            
            Sql2 = "select * from rbonifica_lineas "
            Sql2 = Sql2 & " where codvarie = " & DBSet(Rs!Codvarie, "N")
            Sql2 = Sql2 & " and hastagrado = (select max(hastagrado) from rbonifica_lineas"
            Sql2 = Sql2 & " where codvarie = " & DBSet(Rs!Codvarie, "N") & ")"
            
            Set Rs2 = New ADODB.Recordset
            Rs2.Open Sql2, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            
            If Not Rs2.EOF Then
                Porcen = DBLet(Rs2!Porcentaje, "N")
                Grado = DBLet(Rs2!hastagrado, "N")
            End If
            Set Rs2 = Nothing
            
        End If
        
        SQL1 = "update rhisfruta set kgradobonif = " & DBSet(Grado + Round2(Grado * Porcen / 100, 2), "N")
        SQL1 = SQL1 & " where numalbar = " & DBSet(Rs!numalbar, "N")
        
        conn.Execute SQL1
    
        Rs.MoveNext
    Wend
    Set Rs = Nothing

    CalcularGradoBonificado = True
    conn.CommitTrans
    Pb1.visible = False
    Exit Function

eCalcularGradoBonificado:
    Pb1.visible = False
    conn.RollbackTrans
    MuestraError Err.Number, "Calcular Grado Bonificado", Err.Description
End Function


' Funcion que indica si se ha cargado el campo rhisfruta.kgradobonif
Public Function CalcularGradoBonificadoRealizado(Tabla1 As String, Where1 As String)
Dim SQL As String
Dim Rs As ADODB.Recordset
Dim RS1 As ADODB.Recordset
Dim Rs2 As ADODB.Recordset
Dim SQL1 As String
Dim Sql2 As String
Dim Porcen As Currency
Dim Grado As Currency

    On Error Resume Next

    CalcularGradoBonificadoRealizado = False
    
    SQL = "Select count(*) FROM " & QuitarCaracterACadena(Tabla1, "_1")
    If Where1 <> "" Then
        Where1 = QuitarCaracterACadena(Where1, "_1")
        SQL = SQL & " WHERE " & Where1
    End If
    SQL = SQL & " and (rhisfruta.kgradobonif is null or rhisfruta.kgradobonif = 0) "
    
    CalcularGradoBonificadoRealizado = (TotalRegistros(SQL) = 0)

End Function

Public Function FacturacionLiquidacionesAlmazaraCastelduc(cTabla As String, cWhere As String, FecFac As String, Pb1 As ProgressBar) As Boolean
Dim SQL As String
Dim Rs As ADODB.Recordset
Dim RS1 As ADODB.Recordset

Dim AntSocio As String
Dim AntVarie As String
Dim ActSocio As String
Dim ActVarie As String
Dim actCampo As String
Dim AntCampo As String
Dim ActAlbar As String
Dim AntAlbar As String

Dim HayReg As Boolean

Dim NumError As Long
Dim vImporte As Currency
Dim vPorcIva As String

Dim PrimFac As String
Dim UltFac As String

Dim tipoMov As String
Dim B As Boolean
Dim vSeccion As CSeccion
Dim Kilos As Currency
Dim KilosCal As Currency
Dim Importe As Currency

Dim devuelve As String
Dim Existe As Boolean

'Dim baseimpo As Currency
'Dim BaseReten As Currency
'Dim Neto As Currency
'Dim ImpoIva As Currency
'Dim ImpoReten As Currency
'Dim TotalFac As Currency
Dim Recolect As Byte
Dim vPrecio As Currency

Dim Sql2 As String
Dim Sql3 As String
Dim Sql5 As String


Dim GastosCoop As Currency
Dim GastosAlb As Currency
Dim vPorcGasto As String

Dim campo As String

    On Error GoTo eFacturacion

    FacturacionLiquidacionesAlmazaraCastelduc = False
    
'08052009 antes dentro de transacciones
    BorrarTMPs
    B = CrearTMPs()
    If Not B Then
         Exit Function
    End If
'08052009
    
    conn.BeginTrans
    
    tipoMov = "FLZ"
    
    
    SQL = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute SQL

    SQL = "SELECT  rhisfruta.codsocio, rhisfruta.codvarie, rhisfruta.numalbar, "
    SQL = SQL & " rhisfruta.fecalbar,  rhisfruta.kilosbru, rhisfruta.prestimado, rhisfruta.prliquidalmz, "
    SQL = SQL & " rhisfruta.kilosnet "
    SQL = SQL & " FROM  " & cTabla

    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        SQL = SQL & " WHERE " & cWhere
    End If
    ' ordenado por socio, variedad, campo, numlabar, fecalbar
    SQL = SQL & " order by rhisfruta.codsocio, rhisfruta.codvarie, rhisfruta.numalbar, rhisfruta.fecalbar "
    
    Set vSeccion = New CSeccion
    
    If vSeccion.LeerDatos(vParamAplic.SeccionAlmaz) Then
        If Not vSeccion.AbrirConta Then
            Exit Function
        End If
    End If
    
    HayReg = False
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    If Not Rs.EOF Then
        AntSocio = CStr(DBLet(Rs!Codsocio, "N"))
        AntVarie = CStr(DBLet(Rs!Codvarie, "N"))
        AntAlbar = CStr(DBLet(Rs!numalbar, "N"))
        
        ActSocio = CStr(DBLet(Rs!Codsocio, "N"))
        ActVarie = CStr(DBLet(Rs!Codvarie, "N"))
        ActAlbar = CStr(DBLet(Rs!numalbar, "N"))
        Set vSocio = New cSocio
        If vSocio.LeerDatos(ActSocio) Then
            If vSocio.LeerDatosSeccion(ActSocio, vParamAplic.SeccionAlmaz) Then
                baseimpo = 0
                BaseReten = 0
                ImpoIva = 0
                ImpoReten = 0
                TotalFac = 0
                
                Anticipos = 0
                
                Kilos = 0
                Importe = 0
                
                vPorcIva = ""
                vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vSocio.CodIva, "N")
                PorcIva = CCur(ImporteSinFormato(vPorcIva))
                
                tipoMov = vSocio.CodTipomLiqAlmz
                
                Set vTipoMov = New CTiposMov
                
                numfactu = vTipoMov.ConseguirContador(tipoMov)
                Do
                    numfactu = vTipoMov.ConseguirContador(tipoMov)
                    devuelve = DevuelveDesdeBDNew(cAgro, "rfactsoc", "numfactu", "codtipom", tipoMov, "T", , "numfactu", CStr(numfactu), "N", "fecfactu", FecFac, "F")
                    If devuelve <> "" Then
                        'Ya existe el contador incrementarlo
                        Existe = True
                        vTipoMov.IncrementarContador (tipoMov)
                        numfactu = vTipoMov.ConseguirContador(tipoMov)
                    Else
                        Existe = False
                    End If
                Loop Until Not Existe
                
                vParamAplic.PrimFactLiqAlmz = numfactu
                
            End If
        End If
    End If
   
   ' en almazara no se insertan campos: metemos el minimo codcampo sin condiciones
    campo = "0" ' DevuelveValor("select min(codcampo) from rcampos")
    
    While Not Rs.EOF And B
        ActVarie = DBLet(Rs!Codvarie, "N")
        ActSocio = DBSet(Rs!Codsocio, "N")
        
        If (ActVarie <> AntVarie Or ActSocio <> AntSocio) Then
            If B Then ' descontamos los gastos de los albaranes
                'Para el resto sigue como estaba
                GastosAlb = ObtenerGastosAlbaranes(AntSocio, AntVarie, campo, cTabla, cWhere, 1, 1)
                Importe = Importe - GastosAlb
                baseimpo = baseimpo - GastosAlb
            End If
                        
            ' insertar linea de variedad, campo
            If B Then
                B = InsertLinea(tipoMov, CStr(numfactu), FecFac, AntVarie, campo, CStr(Kilos), CStr(Importe), CStr(GastosAlb))
            End If
            
            If B Then
                ' tenemos que descontar los anticipos que tengamos para ello
                Sql2 = "select rfactsoc_variedad.imporvar, rfactsoc_variedad.numfactu, rfactsoc_variedad.fecfactu "
                Sql2 = Sql2 & " from rfactsoc_variedad INNER JOIN rfactsoc ON rfactsoc_variedad.codtipom = rfactsoc.codtipom and "
                Sql2 = Sql2 & " rfactsoc_variedad.numfactu = rfactsoc.numfactu and rfactsoc_variedad.fecfactu = rfactsoc.fecfactu "
                Sql2 = Sql2 & " where rfactsoc_variedad.codtipom = " & DBSet(vSocio.CodTipomAntAlmz, "T") ' antes era 'FAA' "
                Sql2 = Sql2 & " and rfactsoc.codsocio = " & DBSet(AntSocio, "N")
'                Sql2 = Sql2 & " and codcampo = " & DBSet(Campo, "N")
                Sql2 = Sql2 & " and codvarie = " & DBSet(AntVarie, "N")
                Sql2 = Sql2 & " and rfactsoc_variedad.descontado = 0"
                
                Set RS1 = New ADODB.Recordset
                RS1.Open Sql2, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
                
                While Not RS1.EOF
                    baseimpo = baseimpo - DBLet(RS1.Fields(0).Value, "N")
                    Anticipos = Anticipos + DBLet(RS1.Fields(0).Value, "N")
                    
                    ' indicamos que los anticipos ya han sido descontados
                    Sql3 = "update rfactsoc_variedad set descontado = 1 where codtipom = " & DBSet(vSocio.CodTipomAntAlmz, "T") ' antes era 'FAC'
                    Sql3 = Sql3 & " and numfactu = " & DBSet(RS1!numfactu, "N")
                    Sql3 = Sql3 & " and fecfactu = " & DBSet(RS1!fecfactu, "F") & " and codvarie = " & DBSet(AntVarie, "N")
'                    Sql3 = Sql3 & " and codcampo = " & DBSet(Campo, "N")
                    
                    conn.Execute Sql3
                    
                    ' insertamos en la tabla de anticipos de liquidacion
                    Sql3 = "insert into tmpFact_anticipos (codtipom, numfactu, fecfactu, codtipomanti, numfactuanti, fecfactuanti, codvarieanti, codcampoanti, baseimpo) values ("
                    Sql3 = Sql3 & DBSet(vSocio.CodTipomLiqAlmz, "T") & ","
                    Sql3 = Sql3 & DBSet(numfactu, "N") & "," & DBSet(FecFac, "F") & ","
                    Sql3 = Sql3 & DBSet(vSocio.CodTipomAntAlmz, "T") & ","
                    Sql3 = Sql3 & DBSet(RS1!numfactu, "N") & "," & DBSet(RS1!fecfactu, "F") & ","
                    Sql3 = Sql3 & DBSet(AntVarie, "N") & "," & DBSet(campo, "N") & "," & DBSet(RS1!imporvar, "N") & ")"
                    
                    conn.Execute Sql3
                    
                    RS1.MoveNext
                Wend
                
                Set RS1 = Nothing
                ' fin descontar anticipos
            
            End If
            
            
            If B Then
                AntVarie = ActVarie
                Kilos = 0
                Importe = 0
            End If
        End If
        
        If ActSocio <> AntSocio Then
'            ' me machaco la base imponible por culpa de los redondeos
'            Sql5 = "select sum(if(importe is null,0,importe)) from tmpfact_albaran where codtipom =" & DBSet(tipoMov, "T")
'            Sql5 = Sql5 & " and numfactu = " & DBSet(numfactu, "N") & " and fecfactu = " & DBSet(FecFac, "F")
'
'            baseimpo = DevuelveValor(Sql5)
'
'            Sql5 = "select sum(if(imporgasto is null,0,imporgasto)) from tmpfact_albaran where codtipom =" & DBSet(tipoMov, "T")
'            Sql5 = Sql5 & " and numfactu = " & DBSet(numfactu, "N") & " and fecfactu = " & DBSet(FecFac, "F")
'
'            GastosAlb = DevuelveValor(Sql5)
'
'            Sql5 = "select sum(if(baseimpo is null,0,baseimpo)) from tmpfact_anticipos where codtipom =" & DBSet(tipoMov, "T")
'            Sql5 = Sql5 & " and numfactu = " & DBSet(numfactu, "N") & " and fecfactu = " & DBSet(FecFac, "F")
'
'            Anticipos = DevuelveValor(Sql5)
'
'            baseimpo = baseimpo - GastosAlb - Anticipos
            
            ImpoIva = Round2(baseimpo * PorcIva / 100, 2)
        
            Select Case DBLet(vSocio.TipoIRPF, "N")
                Case 0
                    ImpoReten = Round2((baseimpo + ImpoIva) * vParamAplic.PorcreteFacSoc / 100, 2)
                    BaseReten = (baseimpo + ImpoIva)
                    PorcReten = vParamAplic.PorcreteFacSoc
                Case 1
                    ImpoReten = Round2(baseimpo * vParamAplic.PorcreteFacSoc / 100, 2)
                    BaseReten = baseimpo
                    PorcReten = vParamAplic.PorcreteFacSoc
                Case 2
                    ImpoReten = 0
                    BaseReten = 0
                    PorcReten = 0
            End Select
            
        
            TotalFac = baseimpo + ImpoIva - ImpoReten
            
            IncrementarProgresNew Pb1, 1
            
            'insertar cabecera de factura
            B = InsertCabecera(tipoMov, CStr(numfactu), FecFac)
            
            If B Then B = InsertResumen(tipoMov, CStr(numfactu))
            
            If B Then B = vTipoMov.IncrementarContador(tipoMov)
            
            If B Then
                AntSocio = ActSocio
                
                Set vSocio = Nothing
                Set vSocio = New cSocio
                If vSocio.LeerDatos(ActSocio) Then
                    If vSocio.LeerDatosSeccion(AntSocio, vParamAplic.SeccionAlmaz) Then
                        vPorcIva = ""
                        vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vSocio.CodIva, "N")
                        PorcIva = CCur(ImporteSinFormato(vPorcIva))
                    End If
                    
                    tipoMov = vSocio.CodTipomLiqAlmz
                End If
                baseimpo = 0
                BaseReten = 0
                ImpoIva = 0
                ImpoReten = 0
                TotalFac = 0
                
                Anticipos = 0
                
                numfactu = vTipoMov.ConseguirContador(tipoMov)
                Do
                    numfactu = vTipoMov.ConseguirContador(tipoMov)
                    devuelve = DevuelveDesdeBDNew(cAgro, "rfactsoc", "numfactu", "codtipom", tipoMov, "T", , "numfactu", CStr(numfactu), "N", "fecfactu", FecFac, "F")
                    If devuelve <> "" Then
                        'Ya existe el contador incrementarlo
                        Existe = True
                        vTipoMov.IncrementarContador (tipoMov)
                        numfactu = vTipoMov.ConseguirContador(tipoMov)
                    Else
                        Existe = False
                    End If
                Loop Until Not Existe
           End If
        End If
        
        vPrecio = DBLet(Rs!Prliquidalmz, "N")
        vImporte = Round2(DBLet(Rs!KilosNet, "N") * DBLet(Rs!Prliquidalmz, "N"), 2)
        
        B = InsertLineaAlbaran(tipoMov, CStr(numfactu), FecFac, Rs, CStr(vPrecio), CStr(vImporte), campo)
        
        Importe = Importe + vImporte
        baseimpo = baseimpo + vImporte
        Kilos = Kilos + DBLet(Rs!KilosNet)
        
        HayReg = True
        
        Rs.MoveNext
    Wend
    
    ' ultimo registro si ha entrado
    If B And HayReg Then
        If B Then ' descontamos los gastos de los albaranes
            GastosAlb = ObtenerGastosAlbaranes(AntSocio, AntVarie, AntCampo, cTabla, cWhere, 1, 1)
            Importe = Importe - GastosAlb
            baseimpo = baseimpo - GastosAlb
        End If
        
                    
        ' insertar linea de variedad
        If B Then B = InsertLinea(tipoMov, CStr(numfactu), FecFac, CStr(ActVarie), CStr(campo), CStr(Kilos), CStr(Importe), CStr(GastosAlb))
        
        ' tenemos que descontar los anticipos que tengamos para ello
        If B Then
            Sql2 = "select rfactsoc_variedad.imporvar, rfactsoc_variedad.numfactu, rfactsoc_variedad.fecfactu "
            Sql2 = Sql2 & " from rfactsoc_variedad INNER JOIN rfactsoc ON rfactsoc_variedad.codtipom = rfactsoc.codtipom and "
            Sql2 = Sql2 & " rfactsoc_variedad.numfactu = rfactsoc.numfactu and rfactsoc_variedad.fecfactu = rfactsoc.fecfactu "
            Sql2 = Sql2 & " where rfactsoc_variedad.codtipom = " & DBSet(vSocio.CodTipomAntAlmz, "T") ' antes era 'FAA' "
            Sql2 = Sql2 & " and rfactsoc.codsocio = " & DBSet(ActSocio, "N")
            Sql2 = Sql2 & " and codvarie = " & DBSet(ActVarie, "N")
'            Sql2 = Sql2 & " and codcampo = " & DBSet(Campo, "N")
            Sql2 = Sql2 & " and rfactsoc_variedad.descontado = 0"
            
            Set RS1 = New ADODB.Recordset
            RS1.Open Sql2, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
            
            While Not RS1.EOF
                baseimpo = baseimpo - DBLet(RS1.Fields(0).Value, "N")
                Anticipos = Anticipos + DBLet(RS1.Fields(0).Value, "N")
                
                ' indicamos que los anticipos ya han sido descontados
                Sql3 = "update rfactsoc_variedad set descontado = 1 where codtipom = " & DBSet(vSocio.CodTipomAntAlmz, "T") ' antes era 'FAC'
                Sql3 = Sql3 & " and numfactu = " & DBSet(RS1!numfactu, "N")
                Sql3 = Sql3 & " and fecfactu = " & DBSet(RS1!fecfactu, "F") & " and codvarie = " & DBSet(ActVarie, "N")
'                Sql3 = Sql3 & " and codcampo = " & DBSet(Campo, "N")
                
                conn.Execute Sql3
                
                ' insertamos en la tabla de anticipos de liquidacion venta campo
                Sql3 = "insert into tmpFact_anticipos (codtipom, numfactu, fecfactu, codtipomanti, numfactuanti, fecfactuanti, codvarieanti, codcampoanti, baseimpo) values ("
                Sql3 = Sql3 & DBSet(vSocio.CodTipomLiqAlmz, "T") & "," ' antes era 'FAL'
                Sql3 = Sql3 & DBSet(numfactu, "N") & "," & DBSet(FecFac, "F") & ","
                Sql3 = Sql3 & DBSet(vSocio.CodTipomAntAlmz, "T") & "," ' antes era 'FAA'
                Sql3 = Sql3 & DBSet(RS1!numfactu, "N") & "," & DBSet(RS1!fecfactu, "F") & ","
                Sql3 = Sql3 & DBSet(ActVarie, "N") & "," & DBSet(campo, "N") & "," & DBSet(RS1!imporvar, "N") & ")"
                
                conn.Execute Sql3
                
                RS1.MoveNext
            Wend
            
            Set RS1 = Nothing
            ' fin descontar anticipos
        
        End If
        
'        ' me machaco la base imponible por culpa de los redondeos
'        Sql5 = "select sum(if(importe is null,0,importe)) from tmpfact_albaran where codtipom =" & DBSet(tipoMov, "T")
'        Sql5 = Sql5 & " and numfactu = " & DBSet(numfactu, "N") & " and fecfactu = " & DBSet(FecFac, "F")
'
'        baseimpo = DevuelveValor(Sql5)
'
'        Sql5 = "select sum(if(imporgasto is null,0,imporgasto)) from tmpfact_albaran where codtipom =" & DBSet(tipoMov, "T")
'        Sql5 = Sql5 & " and numfactu = " & DBSet(numfactu, "N") & " and fecfactu = " & DBSet(FecFac, "F")
'
'        GastosAlb = DevuelveValor(Sql5)
'
'        Sql5 = "select sum(if(baseimpo is null,0,baseimpo)) from tmpfact_anticipos where codtipom =" & DBSet(tipoMov, "T")
'        Sql5 = Sql5 & " and numfactu = " & DBSet(numfactu, "N") & " and fecfactu = " & DBSet(FecFac, "F")
'
'        Anticipos = DevuelveValor(Sql5)
        
'        baseimpo = baseimpo - GastosAlb - Anticipos
        
        ImpoIva = Round2(baseimpo * PorcIva / 100, 2)

        Select Case DBLet(vSocio.TipoIRPF, "N")
            Case 0
                ImpoReten = Round2((baseimpo + ImpoIva) * vParamAplic.PorcreteFacSoc / 100, 2)
                BaseReten = (baseimpo + ImpoIva)
                PorcReten = vParamAplic.PorcreteFacSoc
            Case 1
                ImpoReten = Round2(baseimpo * vParamAplic.PorcreteFacSoc / 100, 2)
                BaseReten = baseimpo
                PorcReten = vParamAplic.PorcreteFacSoc
            Case 2
                ImpoReten = 0
                BaseReten = 0
                PorcReten = 0
        End Select
        
        TotalFac = baseimpo + ImpoIva - ImpoReten
        
        IncrementarProgresNew Pb1, 1
        
        vParamAplic.UltFactLiqAlmz = numfactu
        
        'insertar cabecera de factura
        B = InsertCabecera(tipoMov, CStr(numfactu), FecFac)
        
        If B Then B = InsertResumen(tipoMov, CStr(numfactu))
        
        If B Then B = vTipoMov.IncrementarContador(tipoMov)
        
        'pasamos las temporales a las tablas
        If B Then B = PasarTemporales()
        
        If B Then B = (vParamAplic.Modificar = 1)
    End If
    
'    BorrarTMPs
    
    vSeccion.CerrarConta
    Set vSeccion = Nothing
    Set vSocio = Nothing
    
eFacturacion:
    If Err.Number <> 0 Or Not B Then
        conn.RollbackTrans
        FacturacionLiquidacionesAlmazaraCastelduc = False
    Else
        conn.CommitTrans
        FacturacionLiquidacionesAlmazaraCastelduc = True
    End If
End Function



Public Function FacturacionLiquidacionesPicassent(cTabla As String, cWhere As String, FecFac As String, Pb1 As ProgressBar, TipoPrec As Byte, DescontarFVarias As Boolean) As Boolean
Dim SQL As String
Dim Rs As ADODB.Recordset
Dim RS1 As ADODB.Recordset

Dim AntSocio As String
Dim AntVarie As String
Dim ActSocio As String
Dim ActVarie As String
Dim actCampo As String
Dim AntCampo As String
Dim ActCalid As String
Dim AntCalid As String

Dim HayReg As Boolean

Dim NumError As Long
Dim vImporte As Currency
Dim vPorcIva As String

Dim PrimFac As String
Dim UltFac As String

Dim tipoMov As String
Dim B As Boolean
Dim vSeccion As CSeccion
Dim Kilos As Currency
Dim KilosCal As Currency
Dim Importe As Currency

Dim devuelve As String
Dim Existe As Boolean

'Dim baseimpo As Currency
'Dim BaseReten As Currency
'Dim Neto As Currency
'Dim ImpoIva As Currency
'Dim ImpoReten As Currency
'Dim TotalFac As Currency
Dim Recolect As Byte
Dim vPrecio As Currency

Dim Sql2 As String
Dim Sql3 As String
Dim SqlAlbaranes As String

Dim GastosCoop As Currency
Dim GastosAlb As Currency
Dim vPorcGasto As String

Dim SqlAFO As String

Dim vBonifica As Currency
Dim Bonifica As Currency
Dim PorcBoni As Currency
Dim PorcComi As Currency

    On Error GoTo eFacturacion

    FacturacionLiquidacionesPicassent = False
    
'08052009 antes dentro de transacciones
    BorrarTMPs
    B = CrearTMPs()
    If Not B Then
         Exit Function
    End If
'08052009
    
    conn.BeginTrans
    
    tipoMov = "FAL"
    
    
    SQL = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute SQL

    SQL = "SELECT  rhisfruta.codsocio, rhisfruta.codvarie, rhisfruta.codcampo,"
    SQL = SQL & "rhisfruta.recolect, rhisfruta_clasif.codcalid, rhisfruta.fecalbar, rhisfruta.numalbar, "
    SQL = SQL & "rprecios.fechaini, rprecios.fechafin, rprecios_calidad.tipofact,max(rprecios.contador) contador, sum(rhisfruta_clasif.kilosnet) as kilosnet "
    SQL = SQL & " FROM  " & cTabla


    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        SQL = SQL & " WHERE " & cWhere
    End If
     
    
    ' ordenado por socio, variedad, campo, calidad
    SQL = SQL & " group by rhisfruta.codsocio, rhisfruta.codvarie, rhisfruta.codcampo, rhisfruta_clasif.codcalid, rhisfruta.fecalbar, rhisfruta.numalbar, rhisfruta.recolect "
    SQL = SQL & " order by rhisfruta.codsocio, rhisfruta.codvarie, rhisfruta.codcampo, rhisfruta_clasif.codcalid, rhisfruta.fecalbar, rhisfruta.numalbar, rhisfruta.recolect "
    
    Set vSeccion = New CSeccion
    
    If vSeccion.LeerDatos(vParamAplic.Seccionhorto) Then
        If Not vSeccion.AbrirConta Then
            Exit Function
        End If
    End If
    
    HayReg = False
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    If Not Rs.EOF Then
        AntSocio = CStr(DBLet(Rs!Codsocio, "N"))
        AntVarie = CStr(DBLet(Rs!Codvarie, "N"))
        AntCampo = CStr(DBLet(Rs!codCampo, "N"))
        AntCalid = CStr(DBLet(Rs!codcalid, "N"))
        
        ActSocio = CStr(DBLet(Rs!Codsocio, "N"))
        ActVarie = CStr(DBLet(Rs!Codvarie, "N"))
        actCampo = CStr(DBLet(Rs!codCampo, "N"))
        ActCalid = CStr(DBLet(Rs!codcalid, "N"))
    
        Set vSocio = New cSocio
        If vSocio.LeerDatos(ActSocio) Then
            If vSocio.LeerDatosSeccion(ActSocio, vParamAplic.Seccionhorto) Then
                baseimpo = 0
                BaseReten = 0
                ImpoIva = 0
                ImpoReten = 0
                TotalFac = 0
                ImpoAFO = 0
                BaseAFO = 0
                
                Anticipos = 0
                
                Kilos = 0
                Importe = 0
                
                KilosCal = 0
                
                vPorcIva = ""
                '[Monica]29/04/2011: INTERNAS
                If vSocio.EsFactADVInt Then
                    vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vParamAplic.CodIvaExeADV, "N")
                Else
                    vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vSocio.CodIva, "N")
                End If
                PorcIva = CCur(ImporteSinFormato(vPorcIva))
                
                tipoMov = vSocio.CodTipomLiq
                
                Set vTipoMov = New CTiposMov
                
                numfactu = vTipoMov.ConseguirContador(tipoMov)
                Do
                    numfactu = vTipoMov.ConseguirContador(tipoMov)
                    devuelve = DevuelveDesdeBDNew(cAgro, "rfactsoc", "numfactu", "codtipom", tipoMov, "T", , "numfactu", CStr(numfactu), "N", "fecfactu", FecFac, "F")
                    If devuelve <> "" Then
                        'Ya existe el contador incrementarlo
                        Existe = True
                        vTipoMov.IncrementarContador (tipoMov)
                        numfactu = vTipoMov.ConseguirContador(tipoMov)
                    Else
                        Existe = False
                    End If
                Loop Until Not Existe
                
                vParamAplic.PrimFactLiq = numfactu
                
            End If
        End If
    End If
    
    While Not Rs.EOF And B
        ActCalid = DBLet(Rs!codcalid, "N")
        ActVarie = DBLet(Rs!Codvarie, "N")
        actCampo = DBSet(Rs!codCampo, "N")
        ActSocio = DBSet(Rs!Codsocio, "N")
        
        If (ActCalid <> AntCalid Or AntCampo <> actCampo Or AntVarie <> ActVarie Or AntSocio <> ActSocio) Then
            ' kilos e importe por variedad campo
            Kilos = Kilos + KilosCal
            Importe = Importe + vImporte
            Bonifica = Bonifica + vBonifica
            
            baseimpo = baseimpo + vImporte + vBonifica
            
            B = InsertLineaCalidad(tipoMov, CStr(numfactu), FecFac, AntVarie, AntCampo, CStr(AntCalid), CStr(DBLet(KilosCal, "N")), CStr(vImporte + vBonifica))
            KilosCal = 0
            vImporte = 0
            vBonifica = 0
            
            AntCalid = ActCalid
        End If
        
        If (ActVarie <> AntVarie Or actCampo <> AntCampo Or ActSocio <> AntSocio) Then
            If B Then ' descontamos el porcentaje de gastos de cooperativa
                GastosCoop = 0
                
                vPorcGasto = ""
                vPorcGasto = DevuelveDesdeBDNew(cAgro, "rcoope", "porcgast", "codcoope", vSocio.Cooperativa, "N")
                If vPorcGasto = "" Then vPorcGasto = "0"
                
                '[Monica]25/02/2011: Sólo hay gastos si no es complementaria ( Añadido el if )
                If TipoPrec <> 3 Then
                    GastosCoop = Round2(Importe * CCur(ImporteSinFormato(vPorcGasto)) / 100, 2)
                    Importe = Importe - GastosCoop
                    baseimpo = baseimpo - GastosCoop
                End If
            End If
            
'            If b Then ' descontamos los gastos de los albaranes
''[MONICA] 08/09/2009 : los gastos de transporte se suman en ObtenerGastosAlbaranes, quito lo de David
''                '17 AGOSTO 2009
''                ' David###   Para VALSUR los gastos se suman
''                If vParamAplic.Cooperativa = 1 Then
''                    GastosAlb = ObtenerGastosAlbaranes(AntSocio, AntVarie, AntCampo, cTabla, cWhere)
''                    Importe = Importe + GastosAlb
''                    baseimpo = baseimpo + GastosAlb
''
''                Else
''                    'Para el resto sigue como estaba
'                '[Monica]25/02/2011: Sólo hay gastos si no es complementaria ( Añadido el if )
'                If TipoPrec <> 3 Then
'                    GastosAlb = ObtenerGastosAlbaranes(AntSocio, AntVarie, AntCampo, cTabla, cWhere, 1)
'                    Importe = Importe - GastosAlb
'                    BaseImpo = BaseImpo - GastosAlb
'                End If
'            End If
            
            '[Monica]08/04/2010: grabamos los albaranes que intervienen en la linea de factura
            If B Then
                B = InsertarAlbaranesFactura(tipoMov, CStr(numfactu), FecFac, AntSocio, AntVarie, AntCampo, cTabla, cWhere, 0)
            End If
                        
            ' insertar linea de variedad, campo
            If B Then
                B = InsertLinea(tipoMov, CStr(numfactu), FecFac, AntVarie, AntCampo, CStr(Kilos), CStr(Importe + Bonifica), CStr(GastosAlb))
            End If
            
            If B Then
                ' tenemos que descontar los anticipos que tengamos para ello
                Sql2 = "select rfactsoc_variedad.imporvar, rfactsoc_variedad.numfactu, rfactsoc_variedad.fecfactu "
                Sql2 = Sql2 & " from rfactsoc_variedad INNER JOIN rfactsoc ON rfactsoc_variedad.codtipom = rfactsoc.codtipom and "
                Sql2 = Sql2 & " rfactsoc_variedad.numfactu = rfactsoc.numfactu and rfactsoc_variedad.fecfactu = rfactsoc.fecfactu "
                Sql2 = Sql2 & " where rfactsoc_variedad.codtipom = " & DBSet(vSocio.CodTipomAnt, "T") ' antes era 'FAA' "
                Sql2 = Sql2 & " and rfactsoc.codsocio = " & DBSet(AntSocio, "N")
                Sql2 = Sql2 & " and codvarie = " & DBSet(AntVarie, "N")
                Sql2 = Sql2 & " and codcampo = " & DBSet(AntCampo, "N")
                Sql2 = Sql2 & " and rfactsoc_variedad.descontado = 0"
                
                Set RS1 = New ADODB.Recordset
                RS1.Open Sql2, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
                
                While Not RS1.EOF
                    baseimpo = baseimpo - DBLet(RS1.Fields(0).Value, "N")
                    Anticipos = Anticipos + DBLet(RS1.Fields(0).Value, "N")
                    
                    ' indicamos que los anticipos ya han sido descontados
                    Sql3 = "update rfactsoc_variedad set descontado = 1 where codtipom = " & DBSet(vSocio.CodTipomAnt, "T") ' antes era 'FAC'
                    Sql3 = Sql3 & " and numfactu = " & DBSet(RS1!numfactu, "N")
                    Sql3 = Sql3 & " and fecfactu = " & DBSet(RS1!fecfactu, "F") & " and codvarie = " & DBSet(AntVarie, "N")
                    Sql3 = Sql3 & " and codcampo = " & DBSet(AntCampo, "N")
                    
                    conn.Execute Sql3
                    
                    ' insertamos en la tabla de anticipos de liquidacion venta campo
                    Sql3 = "insert into tmpFact_anticipos (codtipom, numfactu, fecfactu, codtipomanti, numfactuanti, fecfactuanti, codvarieanti, codcampoanti, baseimpo) values ("
                    Sql3 = Sql3 & DBSet(vSocio.CodTipomLiq, "T") & "," ' antes era 'FAL'
                    Sql3 = Sql3 & DBSet(numfactu, "N") & "," & DBSet(FecFac, "F") & ","
                    Sql3 = Sql3 & DBSet(vSocio.CodTipomAnt, "T") & "," ' antes era 'FAA'
                    Sql3 = Sql3 & DBSet(RS1!numfactu, "N") & "," & DBSet(RS1!fecfactu, "F") & ","
                    Sql3 = Sql3 & DBSet(AntVarie, "N") & "," & DBSet(AntCampo, "N") & "," & DBSet(RS1!imporvar, "N") & ")"
                    
                    conn.Execute Sql3
                    
                    RS1.MoveNext
                Wend
                
                Set RS1 = Nothing
                ' fin descontar anticipos
            
            End If
            
            
            If B Then
                AntVarie = ActVarie
                AntCampo = actCampo
                Kilos = 0
                Importe = 0
                Bonifica = 0
            End If
        End If
        
        If ActSocio <> AntSocio Then
            
            ImpoIva = Round2(baseimpo * PorcIva / 100, 2)
        
            Select Case DBLet(vSocio.TipoIRPF, "N")
                Case 0
                    ImpoReten = Round2((baseimpo + ImpoIva) * vParamAplic.PorcreteFacSoc / 100, 2)
                    BaseReten = (baseimpo + ImpoIva)
                    PorcReten = vParamAplic.PorcreteFacSoc
                Case 1
                    ImpoReten = Round2(baseimpo * vParamAplic.PorcreteFacSoc / 100, 2)
                    BaseReten = baseimpo
                    PorcReten = vParamAplic.PorcreteFacSoc
                Case 2
                    ImpoReten = 0
                    BaseReten = 0
                    PorcReten = 0
            End Select
            
            If TipoPrec <> 3 Then
                ' El importe AFO lo tiene que tener guardado en la tabla intermedia
                ImpoAFO = DevuelveValor("select sum(importe) from raporreparto where codsocio = " & DBSet(vSocio.Codigo, "N") & " and tipoentr = 0")
            Else
                ImpoAFO = 0
            End If
            BaseAFO = 0
            PorcAFO = 0

            TotalFac = baseimpo + ImpoIva - ImpoReten - ImpoAFO
            
            IncrementarProgresNew Pb1, 1
            
            '[Monica]07/02/2012: indicamos si es una factura de liquidacion complementaria
            'insertar cabecera de factura
            B = InsertCabecera(tipoMov, CStr(numfactu), FecFac, , , (TipoPrec = 3))
            
            If B Then B = InsertResumen(tipoMov, CStr(numfactu))
            
'Mirar si quito lo de reclacular calidades
'            If b Then b = RecalcularCalidades(TipoMov, CStr(numfactu), FecFac)
            
'Recalculo de todos los importes de tmpfact_calidades y tmpfact_variedades para que cuadre con la base de cabecera
'            If b Then b = CuadrarBasesFactura(TipoMov, CStr(numfactu), FecFac, BaseImpo)

            '[Monica]15/04/2013: Descontamos facturas varias
            If DescontarFVarias Then
                If B Then B = InsertFacturasVarias(tipoMov, CStr(numfactu), FecFac, 1, 0)
            End If
            
            If B Then B = vTipoMov.IncrementarContador(tipoMov)
            
            If B Then
                AntSocio = ActSocio
                
                Set vSocio = Nothing
                Set vSocio = New cSocio
                If vSocio.LeerDatos(ActSocio) Then
                    If vSocio.LeerDatosSeccion(AntSocio, vParamAplic.Seccionhorto) Then
                        vPorcIva = ""
                        '[Monica]29/04/2011: INTERNAS
                        If vSocio.EsFactADVInt Then
                            vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vParamAplic.CodIvaExeADV, "N")
                        Else
                            vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vSocio.CodIva, "N")
                        End If
                        
                        PorcIva = CCur(ImporteSinFormato(vPorcIva))
                    End If
                    tipoMov = vSocio.CodTipomLiq
                                        
                End If
                baseimpo = 0
                BaseReten = 0
                ImpoIva = 0
                ImpoReten = 0
                TotalFac = 0
                BaseAFO = 0
                ImpoAFO = 0
                
                Anticipos = 0
                
                numfactu = vTipoMov.ConseguirContador(tipoMov)
                Do
                    numfactu = vTipoMov.ConseguirContador(tipoMov)
                    devuelve = DevuelveDesdeBDNew(cAgro, "rfactsoc", "numfactu", "codtipom", tipoMov, "T", , "numfactu", CStr(numfactu), "N", "fecfactu", FecFac, "F")
                    If devuelve <> "" Then
                        'Ya existe el contador incrementarlo
                        Existe = True
                        vTipoMov.IncrementarContador (tipoMov)
                        numfactu = vTipoMov.ConseguirContador(tipoMov)
                    Else
                        Existe = False
                    End If
                Loop Until Not Existe
           End If
        End If
        
        Recolect = DBLet(Rs!Recolect, "N")
        
        '[Monica]01/09/2010: añadido ésto, antes los precios los sacabamos en el propio select
        Dim Sql9 As String
        Dim Rs9 As ADODB.Recordset
        Dim PreCoop As Currency
        Dim PreSocio As Currency
        
        Sql9 = "select precoop, presocio from rprecios_calidad where codvarie = " & DBSet(Rs!Codvarie, "N")
        Sql9 = Sql9 & " and tipofact = " & DBSet(Rs!TipoFact, "N")
        Sql9 = Sql9 & " and contador = " & DBSet(Rs!Contador, "N")
        Sql9 = Sql9 & " and codcalid = " & DBSet(Rs!codcalid, "N")
        
        Set Rs9 = New ADODB.Recordset
        Rs9.Open Sql9, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
        
        If Not Rs9.EOF Then
            PreCoop = DBLet(Rs9.Fields(0).Value, "N")
            PreSocio = DBLet(Rs9.Fields(1).Value, "N")
            PorcBoni = 0
        
            Select Case Recolect
                Case 0
                    ' si el precio es positivo miramos si hay porcentaje de bonificacion para esa fecha
                    If PreCoop > 0 Then
                        PorcBoni = DevuelveValor("select porcbonif from rbonifentradas where codvarie = " & DBSet(Rs!Codvarie, "N") & " and fechaent = " & DBSet(Rs!Fecalbar, "F"))
                        
                        '[Monica]03/02/2012: Si el precio es positivo vemos si tiene comision el campo y se lo descontamos si es positivo
                        PorcComi = DevuelveValor("select dtoprecio from rcampos where codcampo = " & DBSet(Rs!codCampo, "N"))
                        If CCur(PorcComi) <> 0 Then
                            PreCoop = PreCoop - Round2(PreCoop * PorcComi / 100, 4)
                        End If
                    End If
                    vPrecio = DBLet(PreCoop, "N")
                    vImporte = vImporte + Round2(DBLet(Rs!KilosNet, "N") * PreCoop, 2)
                    
                    vBonifica = vBonifica + Round2(DBLet(Rs!KilosNet, "N") * PreCoop * (PorcBoni / 100), 2)
                Case 1
                    ' si el precio es positivo miramos si hay porcentaje de bonificacion para esa fecha
                    If PreSocio > 0 Then
                        PorcBoni = DevuelveValor("select porcbonif from rbonifentradas where codvarie = " & DBSet(Rs!Codvarie, "N") & " and fechaent = " & DBSet(Rs!Fecalbar, "F"))
                        
                        '[Monica]03/02/2012: Si el precio es positivo vemos si tiene comision el campo y se lo descontamos si es positivo
                        PorcComi = DevuelveValor("select dtoprecio from rcampos where codcampo = " & DBSet(Rs!codCampo, "N"))
                        If CCur(PorcComi) <> 0 Then
                            PreSocio = PreSocio - Round2(PreSocio * PorcComi / 100, 4)
                        End If
                    End If
                    vPrecio = DBLet(PreSocio, "N")
                    vImporte = vImporte + Round2(DBLet(Rs!KilosNet, "N") * PreSocio, 2)
                    
                    vBonifica = vBonifica + Round2(DBLet(Rs!KilosNet, "N") * PreSocio * (PorcBoni / 100), 2)
            End Select
            
            KilosCal = KilosCal + DBLet(Rs!KilosNet, "N")
        
            
        End If
        
        Set Rs9 = Nothing
        
        'hasta aqui
        HayReg = True
        
        Rs.MoveNext
    Wend
    
    ' ultimo registro si ha entrado
    If B And HayReg Then
        ' insertar linea de calidad
        Kilos = Kilos + KilosCal
        Importe = Importe + vImporte
        Bonifica = Bonifica + vBonifica
        
        baseimpo = baseimpo + vImporte + vBonifica
        
        
        If B Then B = InsertLineaCalidad(tipoMov, CStr(numfactu), FecFac, ActVarie, actCampo, CStr(ActCalid), CStr(DBLet(KilosCal, "N")), CStr(vImporte + vBonifica))
        
        
        If B Then ' descontamos el porcentaje de gastos de cooperativa
            GastosCoop = 0
            
            vPorcGasto = ""
            vPorcGasto = DevuelveDesdeBDNew(cAgro, "rcoope", "porcgast", "codcoope", vSocio.Cooperativa, "N")
            If vPorcGasto = "" Then vPorcGasto = "0"
            
            '[Monica]25/02/2011: Sólo hay gastos si no es complementaria ( Añadido el if )
            If TipoPrec <> 3 Then
                GastosCoop = Round2(Importe * CCur(ImporteSinFormato(vPorcGasto)) / 100, 2)
                Importe = Importe - GastosCoop
                baseimpo = baseimpo - GastosCoop
            End If
        End If
        
'        If b Then ' descontamos los gastos de los albaranes
''[MONICA] 08/09/2009 : los gastos de transporte se suman en ObtenerGastosAlbaranes, quito lo de David
''            '17 AGOSTO 2009
''            ' David###   Para VALSUR los gastos se suman
''            GastosAlb = ObtenerGastosAlbaranes(AntSocio, AntVarie, AntCampo, cTabla, cWhere)
''            If vParamAplic.Cooperativa = 1 Then
''                Importe = Importe + GastosAlb
''                baseimpo = baseimpo + GastosAlb
''            Else
''                Importe = Importe - GastosAlb
''                baseimpo = baseimpo - GastosAlb
''            End If
'
'            '[Monica]25/02/2011: Sólo hay gastos si no es complementaria ( Añadido el if )
'            If TipoPrec <> 3 Then
'                GastosAlb = ObtenerGastosAlbaranes(AntSocio, AntVarie, AntCampo, cTabla, cWhere, 1)
'                Importe = Importe - GastosAlb
'                BaseImpo = BaseImpo - GastosAlb
'            End If
'
'        End If
        
        '[Monica]08/04/2010: grabamos los albaranes que intervienen en la linea de factura
        If B Then
            B = InsertarAlbaranesFactura(tipoMov, CStr(numfactu), FecFac, AntSocio, AntVarie, AntCampo, cTabla, cWhere, 0)
        End If
                    
                    
        ' insertar linea de variedad
        If B Then B = InsertLinea(tipoMov, CStr(numfactu), FecFac, CStr(ActVarie), CStr(actCampo), CStr(Kilos), CStr(Importe + Bonifica), CStr(GastosAlb))
        
        ' tenemos que descontar los anticipos que tengamos para ello
        If B Then
            Sql2 = "select rfactsoc_variedad.imporvar, rfactsoc_variedad.numfactu, rfactsoc_variedad.fecfactu "
            Sql2 = Sql2 & " from rfactsoc_variedad INNER JOIN rfactsoc ON rfactsoc_variedad.codtipom = rfactsoc.codtipom and "
            Sql2 = Sql2 & " rfactsoc_variedad.numfactu = rfactsoc.numfactu and rfactsoc_variedad.fecfactu = rfactsoc.fecfactu "
            Sql2 = Sql2 & " where rfactsoc_variedad.codtipom = " & DBSet(vSocio.CodTipomAnt, "T") ' antes era 'FAA' "
            Sql2 = Sql2 & " and rfactsoc.codsocio = " & DBSet(ActSocio, "N")
            Sql2 = Sql2 & " and codvarie = " & DBSet(ActVarie, "N")
            Sql2 = Sql2 & " and codcampo = " & DBSet(actCampo, "N")
            Sql2 = Sql2 & " and rfactsoc_variedad.descontado = 0"
            
            Set RS1 = New ADODB.Recordset
            RS1.Open Sql2, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
            
            While Not RS1.EOF
                baseimpo = baseimpo - DBLet(RS1.Fields(0).Value, "N")
                Anticipos = Anticipos + DBLet(RS1.Fields(0).Value, "N")
                
                ' indicamos que los anticipos ya han sido descontados
                Sql3 = "update rfactsoc_variedad set descontado = 1 where codtipom = " & DBSet(vSocio.CodTipomAnt, "T") ' antes era 'FAC'
                Sql3 = Sql3 & " and numfactu = " & DBSet(RS1!numfactu, "N")
                Sql3 = Sql3 & " and fecfactu = " & DBSet(RS1!fecfactu, "F") & " and codvarie = " & DBSet(ActVarie, "N")
                Sql3 = Sql3 & " and codcampo = " & DBSet(actCampo, "N")
                
                conn.Execute Sql3
                
                ' insertamos en la tabla de anticipos de liquidacion venta campo
                Sql3 = "insert into tmpFact_anticipos (codtipom, numfactu, fecfactu, codtipomanti, numfactuanti, fecfactuanti, codvarieanti, codcampoanti, baseimpo) values ("
                Sql3 = Sql3 & DBSet(vSocio.CodTipomLiq, "T") & "," ' antes era 'FAL'
                Sql3 = Sql3 & DBSet(numfactu, "N") & "," & DBSet(FecFac, "F") & ","
                Sql3 = Sql3 & DBSet(vSocio.CodTipomAnt, "T") & "," ' antes era 'FAA'
                Sql3 = Sql3 & DBSet(RS1!numfactu, "N") & "," & DBSet(RS1!fecfactu, "F") & ","
                Sql3 = Sql3 & DBSet(ActVarie, "N") & "," & DBSet(actCampo, "N") & "," & DBSet(RS1!imporvar, "N") & ")"
                
                conn.Execute Sql3
                
                RS1.MoveNext
            Wend
            
            Set RS1 = Nothing
            ' fin descontar anticipos
        
        End If
        
        
        ImpoIva = Round2(baseimpo * PorcIva / 100, 2)

        Select Case DBLet(vSocio.TipoIRPF, "N")
            Case 0
                ImpoReten = Round2((baseimpo + ImpoIva) * vParamAplic.PorcreteFacSoc / 100, 2)
                BaseReten = (baseimpo + ImpoIva)
                PorcReten = vParamAplic.PorcreteFacSoc
            Case 1
                ImpoReten = Round2(baseimpo * vParamAplic.PorcreteFacSoc / 100, 2)
                BaseReten = baseimpo
                PorcReten = vParamAplic.PorcreteFacSoc
            Case 2
                ImpoReten = 0
                BaseReten = 0
                PorcReten = 0
        End Select
        
        If TipoPrec <> 3 Then ' si no es complementaria se calcula el impafo
            ImpoAFO = DevuelveValor("select sum(importe) from raporreparto where codsocio = " & DBSet(vSocio.Codigo, "N") & " and tipoentr = 0")
        Else
            ImpoAFO = 0
        End If
        BaseAFO = 0
        PorcAFO = 0

        TotalFac = baseimpo + ImpoIva - ImpoReten - ImpoAFO
        
        IncrementarProgresNew Pb1, 1
        
        vParamAplic.UltFactLiq = numfactu
        
        '[Monica]07/02/2012: indicamos si es una factura de liquidacion complementaria
        'insertar cabecera de factura
        B = InsertCabecera(tipoMov, CStr(numfactu), FecFac, , , (TipoPrec = 3))
        
        '[Monica]15/04/2013: Descontamos facturas varias
        If DescontarFVarias Then
            If B Then B = InsertFacturasVarias(tipoMov, CStr(numfactu), FecFac, 1, 0)
        End If
        
        If B Then B = InsertResumen(tipoMov, CStr(numfactu))

'Mirar si quito lo de reclacular calidades
'        If b Then b = RecalcularCalidades(TipoMov, CStr(numfactu), FecFac)
        
'Recalculo de todos los importes de rfactsoc_calidades y rfactsoc_variedades para que cuadre con la base de cabecera
'        If b Then b = CuadrarBasesFactura(TipoMov, CStr(numfactu), FecFac, BaseImpo)
        
        If B Then B = vTipoMov.IncrementarContador(tipoMov)
        
        'pasamos las temporales a las tablas
        If B Then B = PasarTemporales()
        
        '[Monica]23/07/2012: si no es complementaria se calculan los gastos
        If TipoPrec <> 3 Then
            ' solo para Picassent: he de insertar las lineas de gastos al pie de factura que estan como gastos de albaranes
            If B Then B = DescontarGastosAPie()
        End If
        
        If B Then B = (vParamAplic.Modificar = 1)
    End If
    
'    BorrarTMPs
    
    vSeccion.CerrarConta
    Set vSeccion = Nothing
    Set vSocio = Nothing
    
eFacturacion:
    If Err.Number <> 0 Or Not B Then
        conn.RollbackTrans
        FacturacionLiquidacionesPicassent = False
    Else
        conn.CommitTrans
        FacturacionLiquidacionesPicassent = True
    End If
End Function


Private Function DescontarGastosAPie() As Boolean
Dim SQL As String
Dim Rs As ADODB.Recordset
Dim Sql2 As String
Dim SqlLin As String
Dim CadValues As String
Dim NumLin As Long


    
    On Error GoTo eDescontarGastosAPie
    
    DescontarGastosAPie = False
    
    Sql2 = "insert into rfactsoc_gastos (codtipom, numfactu, fecfactu, numlinea, codgasto, importe) values "
    
    
    CadValues = ""
    
    SQL = "select distinct tmpfact_albaran.codtipom, tmpfact_albaran.numfactu, tmpfact_albaran.fecfactu, "
    SQL = SQL & " rhisfruta_gastos.codgasto, sum(rhisfruta_gastos.importe) impgasto "
    SQL = SQL & " from tmpfact_albaran inner join rhisfruta_gastos on tmpfact_albaran.numalbar = rhisfruta_gastos.numalbar "
    
    '[Monica]07/04/2017: para el caso de coopic no descontamos los gastos que sean de transporte
    If vParamAplic.Cooperativa = 16 Then
        SQL = SQL & " where rhisfruta_gastos.codgasto <> " & vParamAplic.CodGastoTRA
    End If
    
    SQL = SQL & " group by 1, 2, 3, 4"
    
    '[Monica]07/04/2017: para el caso de coopic no descontamos los gastos que sean de transporte
    If vParamAplic.Cooperativa = 16 Then
        SQL = SQL & " having impgasto <> 0"
    End If
    
    SQL = SQL & " order by 1, 2, 3, 4"
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    NumLin = 0
    If Not Rs.EOF Then
        SqlLin = "select max(numlinea) from  rfactsoc_gastos where codtipom = " & DBSet(Rs.Fields(0).Value, "T")
        SqlLin = SqlLin & " and numfactu = " & DBSet(Rs.Fields(1).Value, "N")
        SqlLin = SqlLin & " and fecfactu = " & DBSet(Rs.Fields(2).Value, "F")
        
        NumLin = DevuelveValor(SqlLin)
    End If
    
    While Not Rs.EOF
    
        NumLin = NumLin + 1
    
        CadValues = CadValues & "(" & DBSet(Rs.Fields(0).Value, "T") & "," & DBSet(Rs.Fields(1).Value, "N") & "," & DBSet(Rs.Fields(2).Value, "F") & ","
        CadValues = CadValues & DBSet(NumLin, "N") & "," & DBSet(Rs.Fields(3).Value, "N") & "," & DBSet(Rs.Fields(4).Value, "N") & "),"
        
        Rs.MoveNext
    Wend
    
    Set Rs = Nothing
    
    If CadValues <> "" Then
         'quitamos la ultima coma
        CadValues = Mid(CadValues, 1, Len(CadValues) - 1)
        
        conn.Execute Sql2 & CadValues
    End If
    
    DescontarGastosAPie = True
    Exit Function
    
eDescontarGastosAPie:
    MuestraError Err.Number, "Descontar Gastos a Pie", Err.Description
End Function


'funcion que indica si hay albaranes en ese rango que ya hayan sido liquidados
' dentro de la funcion de liquidacion de alzira proceso valsur

Private Function AlbaranesFacturados(cTabla As String, cWhere As String) As Boolean
Dim SQL As String
Dim cadena As String
Dim Cadena2 As String
Dim Rs As ADODB.Recordset
    
    On Error GoTo eAlbaranesFacturados
    
    AlbaranesFacturados = True
    
    SQL = "select rfactsoc_albaran.numalbar, rfactsoc_albaran.fecalbar "
    SQL = SQL & " from rfactsoc_albaran "
    SQL = SQL & " where numalbar in (select rhisfruta.numalbar from " & cTabla & " where " & cWhere & ")"
    SQL = SQL & " order by 1"
            
    If TotalRegistros(SQL) > 0 Then
        Set Rs = New ADODB.Recordset
        Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
        cadena = ""
    
        While Not Rs.EOF
            cadena = cadena & Format(DBLet(Rs!numalbar, "N"), "0000000") & ", "
        
            Rs.MoveNext
        Wend
        Set Rs = Nothing
        
        
        Cadena2 = "Los siguientes albaranes ya están facturados. ¿Qué desea hacer?" & vbCrLf & vbCrLf & "Liquidarlos todos(Sí), Sólo pendientes(No) o Cancelar proceso" & vbCrLf & vbCrLf & Mid(cadena, 1, Len(cadena) - 2)

        Select Case MsgBox(Cadena2, vbQuestion + vbYesNoCancel + vbDefaultButton1)
            Case vbYes
                ' indicamos como si no hubieran albaranes facturados para poder continuar con el proceso
                ' de liquidacion o de anticipos
                AlbaranesFacturados = True

            Case vbNo
                ' se liquidan todos los albaranes no facturados
                AlbaranesFacturados = True

                cWhere = cWhere & " and rhisfruta.numalbar not in (" & Mid(cadena, 1, Len(cadena) - 2) & ")"

            Case vbCancel
                ' abortamos el proceso
                AlbaranesFacturados = False
        End Select
    End If
    Exit Function
    
eAlbaranesFacturados:
    AlbaranesFacturados = False
    MensError = "Albaranes Facturados"
    MuestraError Err.Number, MensError
End Function



Public Function FacturacionAnticiposGenerico(cTabla As String, cWhere As String, FecFac As String, Pb1 As ProgressBar, FecIni As String, FecFin As String, DeRetirada As Boolean, EsVetoRuso As Boolean) As Boolean
Dim SQL As String
Dim Sql3 As String

Dim Rs As ADODB.Recordset

Dim AntSocio As String
Dim AntVarie As String
Dim ActSocio As String
Dim ActVarie As String
Dim actCampo As String
Dim AntCampo As String
Dim ActCalid As String
Dim AntCalid As String

Dim HayReg As Boolean

Dim NumError As Long
Dim vImporte As Currency
Dim vPorcIva As String

Dim PrimFac As String
Dim UltFac As String

Dim tipoMov As String
Dim B As Boolean
Dim vSeccion As CSeccion
Dim Kilos As Currency
Dim KilosCal As Currency
Dim Importe As Currency

Dim devuelve As String
Dim Existe As Boolean

'Dim baseimpo As Currency
'Dim BaseReten As Currency
Dim Neto As Currency
'Dim ImpoIva As Currency
'Dim ImpoReten As Currency
'Dim TotalFac As Currency
Dim Recolect As Byte
Dim vPrecio As Currency
Dim ConGastos As Byte

Dim Sql8 As String
Dim Precio As Currency


    On Error GoTo eFacturacion

    FacturacionAnticiposGenerico = False
    
    tipoMov = "FAA"
    '[Monica]23/12/2014: si es veto ruso cogemos otro tipo de movimiento
    If EsVetoRuso Then tipoMov = "VAA"
    
    BorrarTMPs
    B = CrearTMPs()
    If Not B Then
         Exit Function
    End If
    
    conn.BeginTrans
    
    SQL = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute SQL

    
    SQL = "SELECT tmpliquidacion.codsocio, tmpliquidacion.codvarie,   "
    SQL = SQL & "sum(tmpliquidacion.kilosnet) as kilosnet"
    SQL = SQL & " FROM  tmpliquidacion "
    SQL = SQL & " WHERE codusu = " & vUsu.Codigo
    SQL = SQL & " group by 1, 2 "
    SQL = SQL & " order by 1, 2 "
    
    
    Set vSeccion = New CSeccion
    
    If vSeccion.LeerDatos(vParamAplic.Seccionhorto) Then
        If Not vSeccion.AbrirConta Then
            Exit Function
        End If
    End If
    
    HayReg = False
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    If Not Rs.EOF Then
        AntSocio = CStr(DBLet(Rs!Codsocio, "N"))
        AntVarie = CStr(DBLet(Rs!Codvarie, "N"))
'        AntCampo = CStr(DBLet(Rs!codcampo, "N"))
        
        ActSocio = CStr(DBLet(Rs!Codsocio, "N"))
        ActVarie = CStr(DBLet(Rs!Codvarie, "N"))
'        actCampo = CStr(DBLet(Rs!codcampo, "N"))
    
        Set vSocio = New cSocio
        If vSocio.LeerDatos(ActSocio) Then
            If vSocio.LeerDatosSeccion(ActSocio, vParamAplic.Seccionhorto) Then
                baseimpo = 0
                BaseReten = 0
                ImpoIva = 0
                ImpoReten = 0
                TotalFac = 0
                BaseAFO = 0
                ImpoAFO = 0
                PorcAFO = 0
                
                Kilos = 0
                Importe = 0
                
                KilosCal = 0
                
                vPorcIva = ""
                vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vSocio.CodIva, "N")
                PorcIva = CCur(ImporteSinFormato(vPorcIva))
                
                tipoMov = vSocio.CodTipomAnt
                
                '[Monica]23/12/2014: si es veto ruso cogemos otro tipo de movimiento
                If EsVetoRuso Then tipoMov = "VAA"
                
                Set vTipoMov = New CTiposMov
                
                numfactu = vTipoMov.ConseguirContador(tipoMov)
                Do
                    numfactu = vTipoMov.ConseguirContador(tipoMov)
                    devuelve = DevuelveDesdeBDNew(cAgro, "rfactsoc", "numfactu", "codtipom", tipoMov, "T", , "numfactu", CStr(numfactu), "N", "fecfactu", FecFac, "F")
                    If devuelve <> "" Then
                        'Ya existe el contador incrementarlo
                        Existe = True
                        vTipoMov.IncrementarContador (tipoMov)
                        numfactu = vTipoMov.ConseguirContador(tipoMov)
                    Else
                        Existe = False
                    End If
                Loop Until Not Existe
                
                vParamAplic.PrimFactAnt = numfactu
                
            End If
        End If
    End If
    
    While Not Rs.EOF And B
        ActVarie = DBLet(Rs!Codvarie, "N")
'        actCampo = DBSet(Rs!codcampo, "N")
        ActSocio = DBSet(Rs!Codsocio, "N")
        
        If ActSocio <> AntSocio Then
            ImpoIva = Round2(baseimpo * PorcIva / 100, 2)
        
            Select Case DBLet(vSocio.TipoIRPF, "N")
                Case 0
                    ImpoReten = Round2((baseimpo + ImpoIva) * vParamAplic.PorcreteFacSoc / 100, 2)
                    BaseReten = (baseimpo + ImpoIva)
                    PorcReten = vParamAplic.PorcreteFacSoc
                Case 1
                    ImpoReten = Round2(baseimpo * vParamAplic.PorcreteFacSoc / 100, 2)
                    BaseReten = baseimpo
                    PorcReten = vParamAplic.PorcreteFacSoc
                Case 2
                    ImpoReten = 0
                    BaseReten = 0
                    PorcReten = 0
            End Select
        
            TotalFac = baseimpo + ImpoIva - ImpoReten - ImpoAFO
            
            IncrementarProgresNew Pb1, 1
            
            'insertar cabecera de factura
            B = InsertCabecera(tipoMov, CStr(numfactu), FecFac, False, DeRetirada)
            
            If B Then B = InsertResumen(tipoMov, CStr(numfactu))
            
            If B Then B = vTipoMov.IncrementarContador(tipoMov)
            
            If B Then
                AntSocio = ActSocio
                
                Set vSocio = Nothing
                Set vSocio = New cSocio
                If vSocio.LeerDatos(ActSocio) Then
                    If vSocio.LeerDatosSeccion(AntSocio, vParamAplic.Seccionhorto) Then
                        vPorcIva = ""
                        vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vSocio.CodIva, "N")
                        PorcIva = CCur(ImporteSinFormato(vPorcIva))
                    End If
                    
                    tipoMov = vSocio.CodTipomAnt
                    
                    '[Monica]23/12/2014: si es veto ruso cogemos otro tipo de movimiento
                    If EsVetoRuso Then tipoMov = "VAA"
                
                End If
                baseimpo = 0
                BaseReten = 0
                Neto = 0
                ImpoIva = 0
                ImpoReten = 0
                TotalFac = 0
                BaseAFO = 0
                ImpoAFO = 0
                
                numfactu = vTipoMov.ConseguirContador(tipoMov)
                Do
                    numfactu = vTipoMov.ConseguirContador(tipoMov)
                    devuelve = DevuelveDesdeBDNew(cAgro, "rfactsoc", "numfactu", "codtipom", tipoMov, "T", , "numfactu", CStr(numfactu), "N", "fecfactu", FecFac, "F")
                    If devuelve <> "" Then
                        'Ya existe el contador incrementarlo
                        Existe = True
                        vTipoMov.IncrementarContador (tipoMov)
                        numfactu = vTipoMov.ConseguirContador(tipoMov)
                    Else
                        Existe = False
                    End If
                Loop Until Not Existe
           End If
        End If
        
        KilosCal = DBLet(Rs!KilosNet, "N")
        Kilos = Kilos + KilosCal
        
        ' insertar linea de variedad, campo
        Sql8 = "select precioindustria from rprecios where (codvarie, tipofact, contador) = ("
        Sql8 = Sql8 & "SELECT codvarie, tipofact, max(contador) FROM rprecios WHERE codvarie=" & DBSet(Rs!Codvarie, "N") & " and "
        If DeRetirada Then
            Sql8 = Sql8 & " tipofact = 5 and fechaini = " & DBSet(FecIni, "F")
        Else
            Sql8 = Sql8 & " tipofact = 4 and fechaini = " & DBSet(FecIni, "F")
        End If
        Sql8 = Sql8 & " and fechafin = " & DBSet(FecFin, "F") & " and precioindustria <> 0 and precioindustria is not null "
        Sql8 = Sql8 & " group by 1, 2) "
        
        Precio = DevuelveValor(Sql8)
        Importe = Round2(Kilos * Precio, 2)
        B = InsertLinea(tipoMov, CStr(numfactu), FecFac, ActVarie, 0, CStr(Kilos), CStr(Importe), "0")
        
        baseimpo = baseimpo + Importe
        
        If B Then
            B = InsertarAlbaranesFactura(tipoMov, CStr(numfactu), FecFac, AntSocio, CStr(Rs!Codvarie), 0, "", "fechaent between " & DBSet(FecIni, "F") & " and " & DBSet(FecFin, "F"), 3)
        End If
        
        
        If B Then
            AntVarie = ActVarie
'            AntCampo = actCampo
            Kilos = 0
            Importe = 0
        End If
        
'
'        If KilosCal <> 0 Then
'            b = InsertLineaCalidad(tipoMov, CStr(numfactu), FecFac, AntVarie, AntCampo, CStr(Rs!codcalid), CStr(DBLet(KilosCal, "N")), 0) ' CStr(vImporte))
'        End If
        
        HayReg = True
        
        Rs.MoveNext
    Wend
    
    ' ultimo registro si ha entrado
    If B And HayReg Then
        ' insertar linea de variedad
'        If b Then
'            ' insertar linea de variedad, campo
'            Sql8 = "select precioindustria from rprecios where (codvarie, tipofact, contador) = ("
'            Sql8 = Sql8 & "SELECT codvarie, tipofact, max(contador) FROM rprecios WHERE codvarie=" & DBSet(AntVarie, "N") & " and "
'            Sql8 = Sql8 & " tipofact = 4 and fechaini = " & DBSet(FecIni, "F")
'            Sql8 = Sql8 & " and fechafin = " & DBSet(FecFin, "F") & " and precioindustria <> 0 and precioindustria is not null "
'            Sql8 = Sql8 & " group by 1, 2) "
'
'            Precio = DevuelveValor(Sql8)
'            Importe = Round2(Kilos * Precio, 2)
'
'            b = InsertLinea(tipoMov, CStr(numfactu), FecFac, ActVarie, actCampo, CStr(Kilos), CStr(Importe), "0")
'
'            baseimpo = baseimpo + Importe
'        End If
'
        ImpoIva = Round2(baseimpo * PorcIva / 100, 2)

        Select Case DBLet(vSocio.TipoIRPF, "N")
            Case 0
                ImpoReten = Round2((baseimpo + ImpoIva) * vParamAplic.PorcreteFacSoc / 100, 2)
                BaseReten = (baseimpo + ImpoIva)
                PorcReten = vParamAplic.PorcreteFacSoc
            Case 1
                ImpoReten = Round2(baseimpo * vParamAplic.PorcreteFacSoc / 100, 2)
                BaseReten = baseimpo
                PorcReten = vParamAplic.PorcreteFacSoc
            Case 2
                ImpoReten = 0
                BaseReten = 0
                PorcReten = 0
        End Select
        
'        BaseAFO = baseimpo
'        PorcAFO = vParamAplic.PorcenAFO
'        ImpoAFO = Round2(BaseAFO * PorcAFO / 100, 2)

        TotalFac = baseimpo + ImpoIva - ImpoReten - ImpoAFO
        
        IncrementarProgresNew Pb1, 1
        
        vParamAplic.UltFactAnt = numfactu
        
        'insertar cabecera de factura
        B = InsertCabecera(tipoMov, CStr(numfactu), FecFac, False, DeRetirada)
        
        If B Then B = InsertResumen(tipoMov, CStr(numfactu))
        
        If B Then B = vTipoMov.IncrementarContador(tipoMov)
        
        
'        If b Then b = ModificarCalidadesFacturasGastos()
        
        'pasamos las temporales a las tablas
        If B Then B = PasarTemporales()
        
        If B Then B = (vParamAplic.Modificar = 1)
    End If
    
'    BorrarTMPs
    
    vSeccion.CerrarConta
    Set vSeccion = Nothing
    Set vSocio = Nothing
    
eFacturacion:
    If Err.Number <> 0 Or Not B Then
        conn.RollbackTrans
        FacturacionAnticiposGenerico = False
    Else
        conn.CommitTrans
        FacturacionAnticiposGenerico = True
    End If
End Function

Public Function FacturacionLiquidacionesQuatretonda(cTabla As String, cWhere As String, FecFac As String, Pb1 As ProgressBar, EsComplemen As Boolean, Seccion As String) As Boolean
Dim SQL As String
Dim Rs As ADODB.Recordset
Dim RS1 As ADODB.Recordset

Dim AntSocio As String
Dim AntVarie As String
Dim ActSocio As String
Dim ActVarie As String
Dim actCampo As String
Dim AntCampo As String
Dim ActCalid As String
Dim AntCalid As String
Dim AntFecIni As String
Dim ActFecIni As String
Dim AntFecFin As String
Dim ActFecFin As String

Dim HayReg As Boolean

Dim NumError As Long
Dim vImporte As Currency
Dim vPorcIva As String

Dim PrimFac As String
Dim UltFac As String

Dim tipoMov As String
Dim B As Boolean
Dim vSeccion As CSeccion
Dim Kilos As Currency
Dim KilosCal As Currency
Dim Importe As Currency

Dim devuelve As String
Dim Existe As Boolean

'Dim baseimpo As Currency
'Dim BaseReten As Currency
'Dim Neto As Currency
'Dim ImpoIva As Currency
'Dim ImpoReten As Currency
'Dim TotalFac As Currency
Dim Recolect As Byte
Dim vPrecio As Currency

Dim Sql2 As String
Dim Sql3 As String

Dim Gastos As Currency
Dim vPorcGasto As String

Dim Sql4 As String
Dim ImpoGastos As Currency

Dim KilosRet As Currency
Dim ImporRet As Currency


    On Error GoTo eFacturacion

    FacturacionLiquidacionesQuatretonda = False
    
'08052009 antes dentro de transacciones
    BorrarTMPs
    B = CrearTMPs()
    If Not B Then
         Exit Function
    End If
'08052009
    
    conn.BeginTrans
    
    tipoMov = "FAL"
    
    SQL = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute SQL

    SQL = "SELECT  tmpliquidacion.codsocio, tmpliquidacion.codvarie,"
    SQL = SQL & "tmpliquidacion.codcalid, "
    SQL = SQL & "sum(tmpliquidacion.kilosnet) as kilosnet, sum(tmpliquidacion.importe) as importe "
    SQL = SQL & " FROM  tmpliquidacion "
    SQL = SQL & " where codusu = " & DBSet(vUsu.Codigo, "N")

    ' ordenado por socio, variedad, campo, calidad
    SQL = SQL & " group by tmpliquidacion.codsocio, tmpliquidacion.codvarie, tmpliquidacion.codcalid "
    SQL = SQL & " order by tmpliquidacion.codsocio, tmpliquidacion.codvarie, tmpliquidacion.codcalid "
    
    Set vSeccion = New CSeccion
'[Monica]25/06/2012: seccion
'    If vSeccion.LeerDatos(vParamAplic.Seccionhorto) Then
    If vSeccion.LeerDatos(Seccion) Then
        If Not vSeccion.AbrirConta Then
            Exit Function
        End If
    End If
    
    HayReg = False
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    If Not Rs.EOF Then
        AntSocio = CStr(DBLet(Rs!Codsocio, "N"))
        AntVarie = CStr(DBLet(Rs!Codvarie, "N"))
'        AntCampo = CStr(DBLet(Rs!codcampo, "N"))
        AntCalid = CStr(DBLet(Rs!codcalid, "N"))
        
        ActSocio = CStr(DBLet(Rs!Codsocio, "N"))
        ActVarie = CStr(DBLet(Rs!Codvarie, "N"))
'        actCampo = CStr(DBLet(Rs!codcampo, "N"))
        ActCalid = CStr(DBLet(Rs!codcalid, "N"))
    
        Set vSocio = New cSocio
        If vSocio.LeerDatos(ActSocio) Then
'[Monica]25/06/2012: seccion
'            If vSocio.LeerDatosSeccion(ActSocio, vParamAplic.Seccionhorto) Then
            If vSocio.LeerDatosSeccion(ActSocio, Seccion) Then
                baseimpo = 0
                BaseReten = 0
                ImpoIva = 0
                ImpoReten = 0
                TotalFac = 0
                ImpoAFO = 0
                BaseAFO = 0
                
                Anticipos = 0
                
                Kilos = 0
                Importe = 0
                
                KilosCal = 0
                
                vPorcIva = ""
                vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vSocio.CodIva, "N")
                PorcIva = CCur(ImporteSinFormato(vPorcIva))
                
                tipoMov = vSocio.CodTipomLiq
                
                Set vTipoMov = New CTiposMov
                
                numfactu = vTipoMov.ConseguirContador(tipoMov)
                Do
                    numfactu = vTipoMov.ConseguirContador(tipoMov)
                    devuelve = DevuelveDesdeBDNew(cAgro, "rfactsoc", "numfactu", "codtipom", tipoMov, "T", , "numfactu", CStr(numfactu), "N", "fecfactu", FecFac, "F")
                    If devuelve <> "" Then
                        'Ya existe el contador incrementarlo
                        Existe = True
                        vTipoMov.IncrementarContador (tipoMov)
                        numfactu = vTipoMov.ConseguirContador(tipoMov)
                    Else
                        Existe = False
                    End If
                Loop Until Not Existe
                
                vParamAplic.PrimFactLiq = numfactu
                
            End If
        End If
    End If
    
    While Not Rs.EOF And B
        ActCalid = DBLet(Rs!codcalid, "N")
        ActVarie = DBLet(Rs!Codvarie, "N")
'        actCampo = DBSet(Rs!codcampo, "N")
        ActSocio = DBSet(Rs!Codsocio, "N")
        
        If (ActCalid <> AntCalid Or AntVarie <> ActVarie Or AntSocio <> ActSocio) Then
            ' kilos e importe por variedad campo
            Kilos = Kilos + KilosCal
            Importe = Importe + vImporte
            
            baseimpo = baseimpo + vImporte
            
            B = InsertLineaCalidad(tipoMov, CStr(numfactu), FecFac, AntVarie, 0, CStr(AntCalid), CStr(DBLet(KilosCal, "N")), CStr(vImporte))
            KilosCal = 0
            vImporte = 0
            
            AntCalid = ActCalid
        End If
        
        If (ActVarie <> AntVarie Or ActSocio <> AntSocio) Then
            If B Then ' descontamos el porcentaje de gastos de cooperativa
                Gastos = 0
                
                vPorcGasto = ""
                vPorcGasto = DevuelveDesdeBDNew(cAgro, "rcoope", "porcgast", "codcoope", vSocio.Cooperativa, "N")
                If vPorcGasto = "" Then vPorcGasto = "0"
                
                Gastos = Round2(Importe * CCur(ImporteSinFormato(vPorcGasto)) / 100, 2)
                Importe = Importe - Gastos
                baseimpo = baseimpo - Gastos
                
            End If
            
            If B Then ' descontamos los gastos de los albaranes
'                Gastos = ObtenerGastosAlbaranesNew(AntSocio, AntVarie, AntCampo, cTabla, cWhere)
                
                ' kilos retirada
                Sql4 = "select sum(codcampo) from tmpliquidacion1 "
                Sql4 = Sql4 & " where codsocio = " & DBSet(AntSocio, "N")
                Sql4 = Sql4 & " and codusu = " & vUsu.Codigo
                Sql4 = Sql4 & " and codvarie = " & DBSet(AntVarie, "N")
                
                ImpoGastos = DevuelveValor(Sql4)
                    
                Kilos = Kilos - ImpoGastos
                KilosRet = ImpoGastos
                
                ' importe retirada
                Sql4 = "select sum(gastos) from tmpliquidacion1 "
                Sql4 = Sql4 & " where codsocio = " & DBSet(AntSocio, "N")
                Sql4 = Sql4 & " and codusu = " & vUsu.Codigo
                Sql4 = Sql4 & " and codvarie = " & DBSet(AntVarie, "N")
                
                ImpoGastos = DevuelveValor(Sql4)
                
                Importe = Importe - ImpoGastos
                ImporRet = ImpoGastos
                
                baseimpo = baseimpo - ImpoGastos
                    
                ImpoGastos = 0
                
                
                ' insertamos en la tabla de anticipos de retirada
                Sql3 = "insert into tmpFact_retirada (codtipom, numfactu, fecfactu, codtipomanti, numfactuanti, fecfactuanti, codvarieanti, codcampoanti, kilosret, imporret) select "
                Sql3 = Sql3 & DBSet(vSocio.CodTipomLiq, "T") & "," ' antes era 'FAL'
                Sql3 = Sql3 & DBSet(numfactu, "N") & "," & DBSet(FecFac, "F") & ","
                '[Monica]23/12/2014: el tipo de movimiento me lo da la factura de anticipo cambio lo de abajo
                'Sql3 = Sql3 & DBSet(vSocio.CodTipomAnt, "T") & "," ' antes era 'FAA'
                Sql3 = Sql3 & " rfactsoc.codtipom,"
                
                Sql3 = Sql3 & " rfactsoc.numfactu, rfactsoc.fecfactu, rfactsoc_variedad.codvarie, 0, sum(if(kilosnet is null,0,kilosnet)), sum(if(Imporvar is null,0,imporvar)) "
                Sql3 = Sql3 & " from rfactsoc INNER JOIN rfactsoc_variedad ON rfactsoc.codtipom = rfactsoc_variedad.codtipom and rfactsoc.numfactu = rfactsoc_variedad.numfactu and rfactsoc.fecfactu = rfactsoc_variedad.fecfactu "
                Sql3 = Sql3 & " where rfactsoc.codtipom in ('FAA','VAA') " ' = " & DBSet(vSocio.CodTipomAnt, "T")
                Sql3 = Sql3 & " and rfactsoc.esretirada = 1 "
                Sql3 = Sql3 & " and rfactsoc.codsocio = " & DBSet(AntSocio, "N")
                Sql3 = Sql3 & " and rfactsoc_variedad.codvarie = " & DBSet(AntVarie, "N")
                Sql3 = Sql3 & " and rfactsoc_variedad.descontado = 0 "
                Sql3 = Sql3 & " group by 1,2,3,4,5,6,7 "
                
                conn.Execute Sql3
                
                
                '[Monica]05/12/2011: marcamos los anticipos de retirada que han intervenido
                        '[Monica]23/12/2014: ahora los tipos de movimiento los pongo a piñon pq tenemos anticipos normales y de veto ruso
                Sql3 = "update rfactsoc_variedad, rfactsoc set rfactsoc_variedad.descontado = 1 where rfactsoc_variedad.codtipom in ('FAA','VAA') " ' & DBSet(vSocio.CodTipomAnt, "T") ' antes era 'FAC'
                Sql3 = Sql3 & " and rfactsoc.codsocio = " & DBSet(AntSocio, "N")
                Sql3 = Sql3 & " and rfactsoc_variedad.codvarie = " & DBSet(AntVarie, "N")
                Sql3 = Sql3 & " and rfactsoc.esretirada = 1 "
                Sql3 = Sql3 & " and rfactsoc_variedad.descontado = 0 "
                Sql3 = Sql3 & " and rfactsoc.codtipom = rfactsoc_variedad.codtipom "
                Sql3 = Sql3 & " and rfactsoc.numfactu = rfactsoc_variedad.numfactu "
                Sql3 = Sql3 & " and rfactsoc.fecfactu = rfactsoc_variedad.fecfactu "
                
                conn.Execute Sql3
                
                
'                Importe = Importe - Gastos
'                baseimpo = baseimpo - Gastos
                
            End If
'demomento
            '[Monica]08/04/2010: grabamos los albaranes que intervienen en la linea de factura
            If B Then
                B = InsertarAlbaranesFactura(tipoMov, CStr(numfactu), FecFac, AntSocio, AntVarie, 0, cTabla, cWhere, 4)
            End If
                        
            ' insertar linea de variedad, campo
            If B Then
                B = InsertLinea(tipoMov, CStr(numfactu), FecFac, AntVarie, 0, CStr(Kilos), CStr(Importe), "0")
            End If
            
            If B Then
                ' tenemos que descontar los anticipos que tengamos para ello ( que no sean de gastos, que no sean de retirada )
                Sql2 = "select rfactsoc_variedad.imporvar, rfactsoc_variedad.numfactu, rfactsoc_variedad.fecfactu "
                Sql2 = Sql2 & " from rfactsoc_variedad INNER JOIN rfactsoc ON rfactsoc_variedad.codtipom = rfactsoc.codtipom and "
                Sql2 = Sql2 & " rfactsoc_variedad.numfactu = rfactsoc.numfactu and rfactsoc_variedad.fecfactu = rfactsoc.fecfactu "
                Sql2 = Sql2 & " where rfactsoc_variedad.codtipom = " & DBSet(vSocio.CodTipomAnt, "T") ' antes era 'FAA' "
                Sql2 = Sql2 & " and rfactsoc.codsocio = " & DBSet(AntSocio, "N")
                Sql2 = Sql2 & " and rfactsoc.esanticipogasto = 0 "
                Sql2 = Sql2 & " and rfactsoc.esretirada = 0 "
                Sql2 = Sql2 & " and codvarie = " & DBSet(AntVarie, "N")
'                Sql2 = Sql2 & " and codcampo = " & DBSet(AntCampo, "N")
                Sql2 = Sql2 & " and rfactsoc_variedad.descontado = 0"
                
                Set RS1 = New ADODB.Recordset
                RS1.Open Sql2, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
                
                While Not RS1.EOF
                    baseimpo = baseimpo - DBLet(RS1.Fields(0).Value, "N")
                    Anticipos = Anticipos + DBLet(RS1.Fields(0).Value, "N")
                    
                    ' indicamos que los anticipos ya han sido descontados
                    Sql3 = "update rfactsoc_variedad set descontado = 1 where codtipom = " & DBSet(vSocio.CodTipomAnt, "T")
                    Sql3 = Sql3 & " and numfactu = " & DBSet(RS1!numfactu, "N")
                    Sql3 = Sql3 & " and fecfactu = " & DBSet(RS1!fecfactu, "F") & " and codvarie = " & DBSet(AntVarie, "N")
'                    Sql3 = Sql3 & " and codcampo = " & DBSet(AntCampo, "N")
                    
                    conn.Execute Sql3
                    
                    ' insertamos en la tabla de anticipos de liquidacion venta campo
                    Sql3 = "insert into tmpFact_anticipos (codtipom, numfactu, fecfactu, codtipomanti, numfactuanti, fecfactuanti, codvarieanti, codcampoanti, baseimpo) values ("
                    Sql3 = Sql3 & DBSet(vSocio.CodTipomLiq, "T") & "," ' antes era 'FAL'
                    Sql3 = Sql3 & DBSet(numfactu, "N") & "," & DBSet(FecFac, "F") & ","
                    Sql3 = Sql3 & DBSet(vSocio.CodTipomAnt, "T") & "," ' antes era 'FAA'
                    Sql3 = Sql3 & DBSet(RS1!numfactu, "N") & "," & DBSet(RS1!fecfactu, "F") & ","
'                    Sql3 = Sql3 & DBSet(AntVarie, "N") & "," & DBSet(AntCampo, "N") & "," & DBSet(RS1!imporvar, "N") & ")"
                    Sql3 = Sql3 & DBSet(AntVarie, "N") & "," & "0," & DBSet(RS1!imporvar, "N") & ")"
                    
                    conn.Execute Sql3
                    
                    RS1.MoveNext
                Wend
                
                Set RS1 = Nothing
                ' fin descontar anticipos
            
            End If
            
            
            If B Then
                AntVarie = ActVarie
                AntCampo = actCampo
                Kilos = 0
                Importe = 0
                KilosRet = 0
                ImporRet = 0
            End If
        End If
        
        If ActSocio <> AntSocio Then
            
            ImpoIva = Round2(baseimpo * PorcIva / 100, 2)
        
            Select Case DBLet(vSocio.TipoIRPF, "N")
                Case 0
                    ImpoReten = Round2((baseimpo + ImpoIva) * vParamAplic.PorcreteFacSoc / 100, 2)
                    BaseReten = (baseimpo + ImpoIva)
                    PorcReten = vParamAplic.PorcreteFacSoc
                Case 1
                    ImpoReten = Round2(baseimpo * vParamAplic.PorcreteFacSoc / 100, 2)
                    BaseReten = baseimpo
                    PorcReten = vParamAplic.PorcreteFacSoc
                Case 2
                    ImpoReten = 0
                    BaseReten = 0
                    PorcReten = 0
            End Select
            
            ImpoAFO = Round2((baseimpo + Anticipos) * vParamAplic.PorcenAFO / 100, 2)
            BaseAFO = baseimpo + Anticipos
            PorcAFO = vParamAplic.PorcenAFO
        
            TotalFac = baseimpo + ImpoIva - ImpoReten - ImpoAFO
            
            IncrementarProgresNew Pb1, 1
            
            '[Monica]07/02/2012: indicamos si es una factura de liquidacion complementaria
            'insertar cabecera de factura
            B = InsertCabecera(tipoMov, CStr(numfactu), FecFac, , , EsComplemen)
            
            If B Then B = InsertResumen(tipoMov, CStr(numfactu))
            
            If B Then B = RecalcularCalidades(tipoMov, CStr(numfactu), FecFac)
            
            If B Then B = vTipoMov.IncrementarContador(tipoMov)
            
            If B Then
                AntSocio = ActSocio
                
                Set vSocio = Nothing
                Set vSocio = New cSocio
                If vSocio.LeerDatos(ActSocio) Then
'[Monica]25/06/2012: Seccion
'                    If vSocio.LeerDatosSeccion(AntSocio, vParamAplic.Seccionhorto) Then
                    If vSocio.LeerDatosSeccion(AntSocio, Seccion) Then
                        vPorcIva = ""
                        vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vSocio.CodIva, "N")
                        PorcIva = CCur(ImporteSinFormato(vPorcIva))
                    End If
                    
                    tipoMov = vSocio.CodTipomLiq
                End If
                baseimpo = 0
                BaseReten = 0
                ImpoIva = 0
                ImpoReten = 0
                TotalFac = 0
                BaseAFO = 0
                ImpoAFO = 0
                
                Anticipos = 0
                
                numfactu = vTipoMov.ConseguirContador(tipoMov)
                Do
                    numfactu = vTipoMov.ConseguirContador(tipoMov)
                    devuelve = DevuelveDesdeBDNew(cAgro, "rfactsoc", "numfactu", "codtipom", tipoMov, "T", , "numfactu", CStr(numfactu), "N", "fecfactu", FecFac, "F")
                    If devuelve <> "" Then
                        'Ya existe el contador incrementarlo
                        Existe = True
                        vTipoMov.IncrementarContador (tipoMov)
                        numfactu = vTipoMov.ConseguirContador(tipoMov)
                    Else
                        Existe = False
                    End If
                Loop Until Not Existe
           End If
        End If
        
'        Recolect = DBLet(RS!Recolect, "N")
'
'        Select Case Recolect
'            Case 0
'                vPrecio = DBLet(RS!precoop, "N")
'                vImporte = vImporte + Round2(DBLet(RS!KilosNet, "N") * RS!precoop, 2)
'            Case 1
'                vPrecio = DBLet(RS!presocio, "N")
'                vImporte = vImporte + Round2(DBLet(RS!KilosNet, "N") * RS!presocio, 2)
'        End Select
        
        vImporte = DBLet(Rs!Importe, "N")
        KilosCal = DBLet(Rs!KilosNet, "N")
        vPrecio = Round2(vImporte / KilosCal, 2)
        
        HayReg = True
        
        Rs.MoveNext
    Wend
    
    ' ultimo registro si ha entrado
    If B And HayReg Then
        ' insertar linea de calidad
        Kilos = Kilos + KilosCal
        Importe = Importe + vImporte
        
        baseimpo = baseimpo + vImporte
        
        
        If B Then B = InsertLineaCalidad(tipoMov, CStr(numfactu), FecFac, ActVarie, actCampo, CStr(ActCalid), CStr(DBLet(KilosCal, "N")), CStr(vImporte))
        
        
        If B Then ' descontamos el porcentaje de gastos de cooperativa
            Gastos = 0
            
            vPorcGasto = ""
            vPorcGasto = DevuelveDesdeBDNew(cAgro, "rcoope", "porcgast", "codcoope", vSocio.Cooperativa, "N")
            If vPorcGasto = "" Then vPorcGasto = "0"
            
            Gastos = Round2(Importe * CCur(ImporteSinFormato(vPorcGasto)) / 100, 2)
            Importe = Importe - Gastos
            baseimpo = baseimpo - Gastos
        End If
        
        If B Then ' descontamos los gastos de los albaranes
'            Gastos = ObtenerGastosAlbaranesNew(AntSocio, AntVarie, AntCampo, cTabla, cWhere)
'            Importe = Importe - Gastos
'            baseimpo = baseimpo - Gastos
        
            ' kilos retirada
            Sql4 = "select sum(codcampo) from tmpliquidacion1 "
            Sql4 = Sql4 & " where codsocio = " & DBSet(AntSocio, "N")
            Sql4 = Sql4 & " and codusu = " & vUsu.Codigo
            Sql4 = Sql4 & " and codvarie = " & DBSet(AntVarie, "N")
            
            ImpoGastos = DevuelveValor(Sql4)
                
            Kilos = Kilos - ImpoGastos
            KilosRet = KilosRet + ImpoGastos
            
            ' importe retirada
            Sql4 = "select sum(gastos) from tmpliquidacion1 "
            Sql4 = Sql4 & " where codsocio = " & DBSet(AntSocio, "N")
            Sql4 = Sql4 & " and codusu = " & vUsu.Codigo
            Sql4 = Sql4 & " and codvarie = " & DBSet(AntVarie, "N")
            
            ImpoGastos = DevuelveValor(Sql4)
            
            Importe = Importe - ImpoGastos
            ImporRet = ImporRet + ImpoGastos
            
            baseimpo = baseimpo - ImpoGastos
                
            ImpoGastos = 0
    
            
            ' insertamos en la tabla de anticipos de retirada
            Sql3 = "insert into tmpFact_retirada (codtipom, numfactu, fecfactu, codtipomanti, numfactuanti, fecfactuanti, codvarieanti, codcampoanti, kilosret, imporret) select "
            Sql3 = Sql3 & DBSet(vSocio.CodTipomLiq, "T") & "," ' antes era 'FAL'
            Sql3 = Sql3 & DBSet(numfactu, "N") & "," & DBSet(FecFac, "F") & ","
            '[Monica]23/12/2014:VR
            Sql3 = Sql3 & " rfactsoc.codtipom, " 'DBSet(vSocio.CodTipomAnt, "T") & "," ' antes era 'FAA'
            Sql3 = Sql3 & " rfactsoc.numfactu, rfactsoc.fecfactu, rfactsoc_variedad.codvarie, 0, sum(if(kilosnet is null,0,kilosnet)), sum(if(Imporvar is null,0,imporvar)) "
            Sql3 = Sql3 & " from rfactsoc INNER JOIN rfactsoc_variedad ON rfactsoc.codtipom = rfactsoc_variedad.codtipom and rfactsoc.numfactu = rfactsoc_variedad.numfactu and rfactsoc.fecfactu = rfactsoc_variedad.fecfactu "
            Sql3 = Sql3 & " where rfactsoc.codtipom in ('FAA','VAA') " '= " & DBSet(vSocio.CodTipomAnt, "T")
            Sql3 = Sql3 & " and rfactsoc.esretirada = 1 "
            Sql3 = Sql3 & " and rfactsoc.codsocio = " & DBSet(AntSocio, "N")
            Sql3 = Sql3 & " and rfactsoc_variedad.codvarie = " & DBSet(AntVarie, "N")
            Sql3 = Sql3 & " and rfactsoc_variedad.descontado = 0 "
            Sql3 = Sql3 & " group by 1,2,3,4,5,6,7 "
            
            conn.Execute Sql3
            
            '[Monica]05/12/2011: marcamos los anticipos de retirada que han intervenido
                                                                                            '[Monica]23/12/2014:VR
            Sql3 = "update rfactsoc_variedad, rfactsoc set rfactsoc_variedad.descontado = 1 where rfactsoc_variedad.codtipom in ('FAA','VAA') " '= " & DBSet(vSocio.CodTipomAnt, "T") ' antes era 'FAC'
            Sql3 = Sql3 & " and rfactsoc.codsocio = " & DBSet(AntSocio, "N")
            Sql3 = Sql3 & " and rfactsoc_variedad.codvarie = " & DBSet(AntVarie, "N")
            Sql3 = Sql3 & " and rfactsoc.esretirada = 1 "
            Sql3 = Sql3 & " and rfactsoc_variedad.descontado = 0 "
            Sql3 = Sql3 & " and rfactsoc.codtipom = rfactsoc_variedad.codtipom "
            Sql3 = Sql3 & " and rfactsoc.numfactu = rfactsoc_variedad.numfactu "
            Sql3 = Sql3 & " and rfactsoc.fecfactu = rfactsoc_variedad.fecfactu "
            
            conn.Execute Sql3
        
        
        End If
'demomento
        '[Monica]08/04/2010: grabamos los albaranes que intervienen en la linea de factura
        If B Then
            B = InsertarAlbaranesFactura(tipoMov, CStr(numfactu), FecFac, AntSocio, AntVarie, 0, cTabla, cWhere, 4)
        End If
                    
        ' insertar linea de variedad
        If B Then B = InsertLinea(tipoMov, CStr(numfactu), FecFac, CStr(ActVarie), CStr(0), CStr(Kilos), CStr(Importe), "0")
        
        ' tenemos que descontar los anticipos que tengamos para ello
        If B Then
            Sql2 = "select rfactsoc_variedad.imporvar, rfactsoc_variedad.numfactu, rfactsoc_variedad.fecfactu "
            Sql2 = Sql2 & " from rfactsoc_variedad INNER JOIN rfactsoc ON rfactsoc_variedad.codtipom = rfactsoc.codtipom and "
            Sql2 = Sql2 & " rfactsoc_variedad.numfactu = rfactsoc.numfactu and rfactsoc_variedad.fecfactu = rfactsoc.fecfactu "
            Sql2 = Sql2 & " where rfactsoc_variedad.codtipom = " & DBSet(vSocio.CodTipomAnt, "T") ' antes era 'FAA' "
            Sql2 = Sql2 & " and rfactsoc.codsocio = " & DBSet(ActSocio, "N")
            Sql2 = Sql2 & " and rfactsoc.esanticipogasto = 0 "
            Sql2 = Sql2 & " and rfactsoc.esretirada = 0 "
            Sql2 = Sql2 & " and codvarie = " & DBSet(ActVarie, "N")
'            Sql2 = Sql2 & " and codcampo = " & DBSet(actCampo, "N")
            Sql2 = Sql2 & " and rfactsoc_variedad.descontado = 0"
            
            Set RS1 = New ADODB.Recordset
            RS1.Open Sql2, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
            
            While Not RS1.EOF
                baseimpo = baseimpo - DBLet(RS1.Fields(0).Value, "N")
                Anticipos = Anticipos + DBLet(RS1.Fields(0).Value, "N")
                
                ' indicamos que los anticipos ya han sido descontados
                Sql3 = "update rfactsoc_variedad set descontado = 1 where codtipom = " & DBSet(vSocio.CodTipomAnt, "T") ' antes era 'FAC'
                Sql3 = Sql3 & " and numfactu = " & DBSet(RS1!numfactu, "N")
                Sql3 = Sql3 & " and fecfactu = " & DBSet(RS1!fecfactu, "F") & " and codvarie = " & DBSet(ActVarie, "N")
'                Sql3 = Sql3 & " and codcampo = " & DBSet(actCampo, "N")
                
                conn.Execute Sql3
                
                ' insertamos en la tabla de anticipos de liquidacion venta campo
                Sql3 = "insert into tmpFact_anticipos (codtipom, numfactu, fecfactu, codtipomanti, numfactuanti, fecfactuanti, codvarieanti, codcampoanti, baseimpo) values ("
                Sql3 = Sql3 & DBSet(vSocio.CodTipomLiq, "T") & "," ' antes era 'FAL'
                Sql3 = Sql3 & DBSet(numfactu, "N") & "," & DBSet(FecFac, "F") & ","
                Sql3 = Sql3 & DBSet(vSocio.CodTipomAnt, "T") & "," ' antes era 'FAA'
                Sql3 = Sql3 & DBSet(RS1!numfactu, "N") & "," & DBSet(RS1!fecfactu, "F") & ","
                Sql3 = Sql3 & DBSet(ActVarie, "N") & "," & "0," & DBSet(RS1!imporvar, "N") & ")"
'                Sql3 = Sql3 & DBSet(ActVarie, "N") & "," & DBSet(actCampo, "N") & "," & DBSet(RS1!imporvar, "N") & ")"
                 
                conn.Execute Sql3
                
                RS1.MoveNext
            Wend
            
            Set RS1 = Nothing
            ' fin descontar anticipos
        
        End If
        
        
        ImpoIva = Round2(baseimpo * PorcIva / 100, 2)

        Select Case DBLet(vSocio.TipoIRPF, "N")
            Case 0
                ImpoReten = Round2((baseimpo + ImpoIva) * vParamAplic.PorcreteFacSoc / 100, 2)
                BaseReten = (baseimpo + ImpoIva)
                PorcReten = vParamAplic.PorcreteFacSoc
            Case 1
                ImpoReten = Round2(baseimpo * vParamAplic.PorcreteFacSoc / 100, 2)
                BaseReten = baseimpo
                PorcReten = vParamAplic.PorcreteFacSoc
            Case 2
                ImpoReten = 0
                BaseReten = 0
                PorcReten = 0
        End Select
        
        ImpoAFO = Round2((baseimpo + Anticipos) * vParamAplic.PorcenAFO / 100, 2)
        BaseAFO = baseimpo + Anticipos
        PorcAFO = vParamAplic.PorcenAFO

        TotalFac = baseimpo + ImpoIva - ImpoReten - ImpoAFO
        
        IncrementarProgresNew Pb1, 1
        
        vParamAplic.UltFactLiq = numfactu
        
        '[Monica]07/02/2012: indicamos si es una factura de liquidacion complementaria
        'insertar cabecera de factura
        B = InsertCabecera(tipoMov, CStr(numfactu), FecFac, , , EsComplemen)
        
        If B Then B = InsertResumen(tipoMov, CStr(numfactu))
        
        If B Then B = RecalcularCalidades(tipoMov, CStr(numfactu), FecFac)
        
        If B Then B = vTipoMov.IncrementarContador(tipoMov)
        
        'pasamos las temporales a las tablas
        If B Then B = PasarTemporales()
        
        If B Then B = (vParamAplic.Modificar = 1)
    End If
    
'    BorrarTMPs
    
    vSeccion.CerrarConta
    Set vSeccion = Nothing
    Set vSocio = Nothing
    
eFacturacion:
    If Err.Number <> 0 Or Not B Then
        conn.RollbackTrans
        FacturacionLiquidacionesQuatretonda = False
    Else
        conn.CommitTrans
        FacturacionLiquidacionesQuatretonda = True
    End If
End Function




'=================ME PUEDE SERVIR PARA LA FACTURA DE RETIRADA DE QUATRETONDA

'Public Function FacturacionAnticiposGenerico(cTabla As String, cWhere As String, FecFac As String, Pb1 As ProgressBar, FecIni As String, FecFin As String) As Boolean
'Dim Sql As String
'Dim Sql3 As String
'
'Dim Rs As ADODB.Recordset
'
'Dim AntSocio As String
'Dim AntVarie As String
'Dim ActSocio As String
'Dim ActVarie As String
'Dim actCampo As String
'Dim AntCampo As String
'Dim ActCalid As String
'Dim AntCalid As String
'
'Dim HayReg As Boolean
'
'Dim NumError As Long
'Dim vImporte As Currency
'Dim vPorcIva As String
'
'Dim PrimFac As String
'Dim UltFac As String
'
'Dim tipoMov As String
'Dim b As Boolean
'Dim vSeccion As CSeccion
'Dim Kilos As Currency
'Dim KilosCal As Currency
'Dim Importe As Currency
'
'Dim devuelve As String
'Dim Existe As Boolean
'
''Dim baseimpo As Currency
''Dim BaseReten As Currency
'Dim Neto As Currency
''Dim ImpoIva As Currency
''Dim ImpoReten As Currency
''Dim TotalFac As Currency
'Dim Recolect As Byte
'Dim vPrecio As Currency
'Dim ConGastos As Byte
'
'Dim Sql8 As String
'Dim Precio As Currency
'
'
'    On Error GoTo eFacturacion
'
'    FacturacionAnticiposGenerico = False
'
'    tipoMov = "FAA"
'
'    BorrarTMPs
'    b = CrearTMPs()
'    If Not b Then
'         Exit Function
'    End If
'
'    conn.BeginTrans
'
'    Sql = "delete from tmpinformes where codusu = " & vUsu.Codigo
'    conn.Execute Sql
'
'
'    Sql = "SELECT tmpliquidacion.codsocio, tmpliquidacion.codvarie, tmpliquidacion.codcampo,  "
'    Sql = Sql & "sum(tmpliquidacion.kilosnet) as kilosnet"
'    Sql = Sql & " FROM  tmpliquidacion "
'    Sql = Sql & " WHERE codusu = " & vUsu.Codigo
'    Sql = Sql & " group by 1, 2, 3 "
'    Sql = Sql & " order by 1, 2, 3 "
'
'
'    Set vSeccion = New CSeccion
'
'    If vSeccion.LeerDatos(vParamAplic.Seccionhorto) Then
'        If Not vSeccion.AbrirConta Then
'            Exit Function
'        End If
'    End If
'
'    HayReg = False
'
'    Set Rs = New ADODB.Recordset
'    Rs.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
'
'    If Not Rs.EOF Then
'        AntSocio = CStr(DBLet(Rs!Codsocio, "N"))
'        AntVarie = CStr(DBLet(Rs!codvarie, "N"))
'        AntCampo = CStr(DBLet(Rs!codcampo, "N"))
'
'        ActSocio = CStr(DBLet(Rs!Codsocio, "N"))
'        ActVarie = CStr(DBLet(Rs!codvarie, "N"))
'        actCampo = CStr(DBLet(Rs!codcampo, "N"))
'
'        Set vSocio = New CSocio
'        If vSocio.LeerDatos(ActSocio) Then
'            If vSocio.LeerDatosSeccion(ActSocio, vParamAplic.Seccionhorto) Then
'                baseimpo = 0
'                BaseReten = 0
'                ImpoIva = 0
'                ImpoReten = 0
'                TotalFac = 0
'                BaseAFO = 0
'                ImpoAFO = 0
'                PorcAFO = 0
'
'                Kilos = 0
'                Importe = 0
'
'                KilosCal = 0
'
'                vPorcIva = ""
'                vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vSocio.CodIva, "N")
'                PorcIva = CCur(ImporteSinFormato(vPorcIva))
'
'                tipoMov = vSocio.CodTipomAnt
'
'                Set vTipoMov = New CTiposMov
'
'                numfactu = vTipoMov.ConseguirContador(tipoMov)
'                Do
'                    numfactu = vTipoMov.ConseguirContador(tipoMov)
'                    devuelve = DevuelveDesdeBDNew(cAgro, "rfactsoc", "numfactu", "codtipom", tipoMov, "T", , "numfactu", CStr(numfactu), "N", "fecfactu", FecFac, "F")
'                    If devuelve <> "" Then
'                        'Ya existe el contador incrementarlo
'                        Existe = True
'                        vTipoMov.IncrementarContador (tipoMov)
'                        numfactu = vTipoMov.ConseguirContador(tipoMov)
'                    Else
'                        Existe = False
'                    End If
'                Loop Until Not Existe
'
'                vParamAplic.PrimFactAnt = numfactu
'
'            End If
'        End If
'    End If
'
'    While Not Rs.EOF And b
'        ActVarie = DBLet(Rs!codvarie, "N")
'        actCampo = DBSet(Rs!codcampo, "N")
'        ActSocio = DBSet(Rs!Codsocio, "N")
'
'        If (ActVarie <> AntVarie Or actCampo <> AntCampo Or ActSocio <> AntSocio) Then
'            ' insertar linea de variedad, campo
'            Sql8 = "select precioindustria from rprecios where (codvarie, tipofact, contador) = ("
'            Sql8 = Sql8 & "SELECT codvarie, tipofact, max(contador) FROM rprecios WHERE codvarie=" & DBSet(AntVarie, "N") & " and "
'            Sql8 = Sql8 & " tipofact = 4 and fechaini = " & DBSet(FecIni, "F")
'            Sql8 = Sql8 & " and fechafin = " & DBSet(FecFin, "F") & " and precioindustria <> 0 and precioindustria is not null "
'            Sql8 = Sql8 & " group by 1, 2) "
'
'            Precio = DevuelveValor(Sql8)
'            Importe = Round2(Kilos * Precio, 2)
'            b = InsertLinea(tipoMov, CStr(numfactu), FecFac, AntVarie, AntCampo, CStr(Kilos), CStr(Importe), "0")
'
'            If b Then
'                b = InsertarAlbaranesFactura(tipoMov, CStr(numfactu), FecFac, AntSocio, AntVarie, AntCampo, "", "fechaent between " & DBSet(FecIni, "F") & " and " & DBSet(FecFin, "F"), 3)
'            End If
'
'            baseimpo = baseimpo + Importe
'
'            If b Then
'                AntVarie = ActVarie
'                AntCampo = actCampo
'                Kilos = 0
'                Importe = 0
'            End If
'        End If
'
'        If ActSocio <> AntSocio Then
'            ImpoIva = Round2(baseimpo * PorcIva / 100, 2)
'
'            Select Case DBLet(vSocio.TipoIRPF, "N")
'                Case 0
'                    ImpoReten = Round2((baseimpo + ImpoIva) * vParamAplic.PorcreteFacSoc / 100, 2)
'                    BaseReten = (baseimpo + ImpoIva)
'                    PorcReten = vParamAplic.PorcreteFacSoc
'                Case 1
'                    ImpoReten = Round2(baseimpo * vParamAplic.PorcreteFacSoc / 100, 2)
'                    BaseReten = baseimpo
'                    PorcReten = vParamAplic.PorcreteFacSoc
'                Case 2
'                    ImpoReten = 0
'                    BaseReten = 0
'                    PorcReten = 0
'            End Select
'
'            TotalFac = baseimpo + ImpoIva - ImpoReten - ImpoAFO
'
'            IncrementarProgresNew Pb1, 1
'
'            'insertar cabecera de factura
'            b = InsertCabecera(tipoMov, CStr(numfactu), FecFac, False)
'
'            If b Then b = InsertResumen(tipoMov, CStr(numfactu))
'
'            If b Then b = vTipoMov.IncrementarContador(tipoMov)
'
'            If b Then
'                AntSocio = ActSocio
'
'                Set vSocio = Nothing
'                Set vSocio = New CSocio
'                If vSocio.LeerDatos(ActSocio) Then
'                    If vSocio.LeerDatosSeccion(AntSocio, vParamAplic.Seccionhorto) Then
'                        vPorcIva = ""
'                        vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vSocio.CodIva, "N")
'                        PorcIva = CCur(ImporteSinFormato(vPorcIva))
'                    End If
'
'                    tipoMov = vSocio.CodTipomAnt
'                End If
'                baseimpo = 0
'                BaseReten = 0
'                Neto = 0
'                ImpoIva = 0
'                ImpoReten = 0
'                TotalFac = 0
'                BaseAFO = 0
'                ImpoAFO = 0
'
'                numfactu = vTipoMov.ConseguirContador(tipoMov)
'                Do
'                    numfactu = vTipoMov.ConseguirContador(tipoMov)
'                    devuelve = DevuelveDesdeBDNew(cAgro, "rfactsoc", "numfactu", "codtipom", tipoMov, "T", , "numfactu", CStr(numfactu), "N", "fecfactu", FecFac, "F")
'                    If devuelve <> "" Then
'                        'Ya existe el contador incrementarlo
'                        Existe = True
'                        vTipoMov.IncrementarContador (tipoMov)
'                        numfactu = vTipoMov.ConseguirContador(tipoMov)
'                    Else
'                        Existe = False
'                    End If
'                Loop Until Not Existe
'           End If
'        End If
'
'        KilosCal = DBLet(Rs!KilosNet, "N")
'        Kilos = Kilos + KilosCal
'
''
''        If KilosCal <> 0 Then
''            b = InsertLineaCalidad(tipoMov, CStr(numfactu), FecFac, AntVarie, AntCampo, CStr(Rs!codcalid), CStr(DBLet(KilosCal, "N")), 0) ' CStr(vImporte))
''        End If
'
'        HayReg = True
'
'        Rs.MoveNext
'    Wend
'
'    ' ultimo registro si ha entrado
'    If b And HayReg Then
'        ' insertar linea de variedad
'        If b Then
'            ' insertar linea de variedad, campo
'            Sql8 = "select precioindustria from rprecios where (codvarie, tipofact, contador) = ("
'            Sql8 = Sql8 & "SELECT codvarie, tipofact, max(contador) FROM rprecios WHERE codvarie=" & DBSet(AntVarie, "N") & " and "
'            Sql8 = Sql8 & " tipofact = 4 and fechaini = " & DBSet(FecIni, "F")
'            Sql8 = Sql8 & " and fechafin = " & DBSet(FecFin, "F") & " and precioindustria <> 0 and precioindustria is not null "
'            Sql8 = Sql8 & " group by 1, 2) "
'
'            Precio = DevuelveValor(Sql8)
'            Importe = Round2(Kilos * Precio, 2)
'
'            b = InsertLinea(tipoMov, CStr(numfactu), FecFac, ActVarie, actCampo, CStr(Kilos), CStr(Importe), "0")
'
'            If b Then
'                b = InsertarAlbaranesFactura(tipoMov, CStr(numfactu), FecFac, ActSocio, ActVarie, ActCampo, "", "fechaent between " & DBSet(FecIni, "F") & " and " & DBSet(FecFin, "F"), 3)
'            End If
'
'            baseimpo = baseimpo + Importe
'        End If
'
'        ImpoIva = Round2(baseimpo * PorcIva / 100, 2)
'
'        Select Case DBLet(vSocio.TipoIRPF, "N")
'            Case 0
'                ImpoReten = Round2((baseimpo + ImpoIva) * vParamAplic.PorcreteFacSoc / 100, 2)
'                BaseReten = (baseimpo + ImpoIva)
'                PorcReten = vParamAplic.PorcreteFacSoc
'            Case 1
'                ImpoReten = Round2(baseimpo * vParamAplic.PorcreteFacSoc / 100, 2)
'                BaseReten = baseimpo
'                PorcReten = vParamAplic.PorcreteFacSoc
'            Case 2
'                ImpoReten = 0
'                BaseReten = 0
'                PorcReten = 0
'        End Select
'
''        BaseAFO = baseimpo
''        PorcAFO = vParamAplic.PorcenAFO
''        ImpoAFO = Round2(BaseAFO * PorcAFO / 100, 2)
'
'        TotalFac = baseimpo + ImpoIva - ImpoReten - ImpoAFO
'
'        IncrementarProgresNew Pb1, 1
'
'        vParamAplic.UltFactAnt = numfactu
'
'        'insertar cabecera de factura
'        b = InsertCabecera(tipoMov, CStr(numfactu), FecFac, False)
'
'        If b Then b = InsertResumen(tipoMov, CStr(numfactu))
'
'        If b Then b = vTipoMov.IncrementarContador(tipoMov)
'
'
''        If b Then b = ModificarCalidadesFacturasGastos()
'
'        'pasamos las temporales a las tablas
'        If b Then b = PasarTemporales()
'
'        If b Then b = (vParamAplic.Modificar = 1)
'    End If
'
''    BorrarTMPs
'
'    vSeccion.CerrarConta
'    Set vSeccion = Nothing
'    Set vSocio = Nothing
'
'eFacturacion:
'    If Err.Number <> 0 Or Not b Then
'        conn.RollbackTrans
'        FacturacionAnticiposGenerico = False
'    Else
'        conn.CommitTrans
'        FacturacionAnticiposGenerico = True
'    End If
'End Function
'




Public Function FacturacionLiquidacionDirecta(Albaran As String, FecFac As String, Precio As String) As Boolean
Dim SQL As String
Dim Rs As ADODB.Recordset
Dim RS1 As ADODB.Recordset

Dim AntSocio As String
Dim AntVarie As String
Dim ActSocio As String
Dim ActVarie As String
Dim actCampo As String
Dim AntCampo As String
Dim ActCalid As String
Dim AntCalid As String

Dim HayReg As Boolean

Dim NumError As Long
Dim vImporte As Currency
Dim vPorcIva As String

Dim PrimFac As String
Dim UltFac As String

Dim tipoMov As String
Dim B As Boolean
Dim vSeccion As CSeccion
Dim Kilos As Currency
Dim KilosCal As Currency
Dim Importe As Currency

Dim devuelve As String
Dim Existe As Boolean

'Dim baseimpo As Currency
'Dim BaseReten As Currency
'Dim Neto As Currency
'Dim ImpoIva As Currency
'Dim ImpoReten As Currency
'Dim TotalFac As Currency
Dim Recolect As Byte
Dim vPrecio As Currency

Dim Sql2 As String
Dim Sql3 As String
Dim SqlAlbaranes As String

Dim GastosCoop As Currency
Dim GastosAlb As Currency
Dim vPorcGasto As String


    On Error GoTo eFacturacion

    FacturacionLiquidacionDirecta = False
    
'08052009 antes dentro de transacciones
    BorrarTMPs
    B = CrearTMPs()
    If Not B Then
         Exit Function
    End If
'08052009
    
    conn.BeginTrans
    
    tipoMov = "FAL"
    
    SQL = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute SQL

    SQL = "SELECT  rhisfruta.codsocio, rhisfruta.codvarie,"
    SQL = SQL & "rhisfruta.codcampo, rhisfruta.recolect, rhisfruta_clasif.codcalid, " & DBSet(Precio, "N") & ", sum(rhisfruta_clasif.kilosnet) as kilosnet "
    SQL = SQL & " FROM  rhisfruta inner join rhisfruta_clasif on rhisfruta.numalbar = rhisfruta_clasif.numalbar where rhisfruta.numalbar = " & DBSet(Albaran, "N")

    ' ordenado por socio, variedad, campo, calidad
    SQL = SQL & " group by rhisfruta.codsocio, rhisfruta.codvarie, rhisfruta.codcampo, rhisfruta_clasif.codcalid, rhisfruta.recolect "
    SQL = SQL & " order by rhisfruta.codsocio, rhisfruta.codvarie, rhisfruta.codcampo, rhisfruta_clasif.codcalid, rhisfruta.recolect "
    
    Set vSeccion = New CSeccion
    
    If vSeccion.LeerDatos(vParamAplic.Seccionhorto) Then
        If Not vSeccion.AbrirConta Then
            Exit Function
        End If
    End If
    
    HayReg = False
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    If Not Rs.EOF Then
        AntSocio = CStr(DBLet(Rs!Codsocio, "N"))
        AntVarie = CStr(DBLet(Rs!Codvarie, "N"))
        AntCampo = CStr(DBLet(Rs!codCampo, "N"))
        AntCalid = CStr(DBLet(Rs!codcalid, "N"))
        
        ActSocio = CStr(DBLet(Rs!Codsocio, "N"))
        ActVarie = CStr(DBLet(Rs!Codvarie, "N"))
        actCampo = CStr(DBLet(Rs!codCampo, "N"))
        ActCalid = CStr(DBLet(Rs!codcalid, "N"))
    
        Set vSocio = New cSocio
        If vSocio.LeerDatos(ActSocio) Then
            If vSocio.LeerDatosSeccion(ActSocio, vParamAplic.Seccionhorto) Then
                baseimpo = 0
                BaseReten = 0
                ImpoIva = 0
                ImpoReten = 0
                TotalFac = 0
                ImpoAFO = 0
                BaseAFO = 0
                
                Anticipos = 0
                
                Kilos = 0
                Importe = 0
                
                KilosCal = 0
                
                vPorcIva = ""
                '[Monica]29/04/2011: INTERNAS
                If vSocio.EsFactADVInt Then
                    vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vParamAplic.CodIvaExeADV, "N")
                Else
                    vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vSocio.CodIva, "N")
                End If
                PorcIva = CCur(ImporteSinFormato(vPorcIva))
                
                tipoMov = vSocio.CodTipomLiq
                
                Set vTipoMov = New CTiposMov
                
                numfactu = vTipoMov.ConseguirContador(tipoMov)
                Do
                    numfactu = vTipoMov.ConseguirContador(tipoMov)
                    devuelve = DevuelveDesdeBDNew(cAgro, "rfactsoc", "numfactu", "codtipom", tipoMov, "T", , "numfactu", CStr(numfactu), "N", "fecfactu", FecFac, "F")
                    If devuelve <> "" Then
                        'Ya existe el contador incrementarlo
                        Existe = True
                        vTipoMov.IncrementarContador (tipoMov)
                        numfactu = vTipoMov.ConseguirContador(tipoMov)
                    Else
                        Existe = False
                    End If
                Loop Until Not Existe
                
                vParamAplic.PrimFactLiq = numfactu
                
            End If
        End If
    End If
    
    While Not Rs.EOF And B
        ActCalid = DBLet(Rs!codcalid, "N")
        ActVarie = DBLet(Rs!Codvarie, "N")
        actCampo = DBSet(Rs!codCampo, "N")
        ActSocio = DBSet(Rs!Codsocio, "N")
        
        If (ActCalid <> AntCalid Or AntCampo <> actCampo Or AntVarie <> ActVarie Or AntSocio <> ActSocio) Then
            ' kilos e importe por variedad campo
            Kilos = Kilos + KilosCal
            Importe = Importe + vImporte
            
            baseimpo = baseimpo + vImporte
            
            B = InsertLineaCalidad(tipoMov, CStr(numfactu), FecFac, AntVarie, AntCampo, CStr(AntCalid), CStr(DBLet(KilosCal, "N")), CStr(vImporte))
            KilosCal = 0
            vImporte = 0
            
            AntCalid = ActCalid
        End If
        
        If (ActVarie <> AntVarie Or actCampo <> AntCampo Or ActSocio <> AntSocio) Then
            If B Then ' descontamos el porcentaje de gastos de cooperativa
                GastosCoop = 0
                
                vPorcGasto = ""
                vPorcGasto = DevuelveDesdeBDNew(cAgro, "rcoope", "porcgast", "codcoope", vSocio.Cooperativa, "N")
                If vPorcGasto = "" Then vPorcGasto = "0"
                
                GastosCoop = Round2(Importe * CCur(ImporteSinFormato(vPorcGasto)) / 100, 2)
                Importe = Importe - GastosCoop
                baseimpo = baseimpo - GastosCoop
            End If
            
            If B Then ' descontamos los gastos de los albaranes
                GastosAlb = ObtenerGastosAlbaranes(AntSocio, AntVarie, AntCampo, "rhifruta", "rhifruta.numalbar = " & DBSet(Albaran, "N"), 1)
                Importe = Importe - GastosAlb
                baseimpo = baseimpo - GastosAlb
            End If
            
            '[Monica]08/04/2010: grabamos los albaranes que intervienen en la linea de factura
            If B Then
                B = InsertarAlbaranesFactura(tipoMov, CStr(numfactu), FecFac, AntSocio, AntVarie, AntCampo, "rhisfruta", "rhifruta.numalbar = " & DBSet(Albaran, "N"), 0)
            End If
                        
            ' insertar linea de variedad, campo
            If B Then
                B = InsertLinea(tipoMov, CStr(numfactu), FecFac, AntVarie, AntCampo, CStr(Kilos), CStr(Importe), CStr(GastosAlb))
            End If
            
            If B Then
                ' tenemos que descontar los anticipos que tengamos para ello
                Sql2 = "select rfactsoc_variedad.imporvar, rfactsoc_variedad.numfactu, rfactsoc_variedad.fecfactu "
                Sql2 = Sql2 & " from rfactsoc_variedad INNER JOIN rfactsoc ON rfactsoc_variedad.codtipom = rfactsoc.codtipom and "
                Sql2 = Sql2 & " rfactsoc_variedad.numfactu = rfactsoc.numfactu and rfactsoc_variedad.fecfactu = rfactsoc.fecfactu "
                Sql2 = Sql2 & " where rfactsoc_variedad.codtipom = " & DBSet(vSocio.CodTipomAnt, "T") ' antes era 'FAA' "
                Sql2 = Sql2 & " and rfactsoc.codsocio = " & DBSet(AntSocio, "N")
                Sql2 = Sql2 & " and codvarie = " & DBSet(AntVarie, "N")
                Sql2 = Sql2 & " and codcampo = " & DBSet(AntCampo, "N")
                Sql2 = Sql2 & " and rfactsoc_variedad.descontado = 0"
                
                Set RS1 = New ADODB.Recordset
                RS1.Open Sql2, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
                
                While Not RS1.EOF
                    baseimpo = baseimpo - DBLet(RS1.Fields(0).Value, "N")
                    Anticipos = Anticipos + DBLet(RS1.Fields(0).Value, "N")
                    
                    ' indicamos que los anticipos ya han sido descontados
                    Sql3 = "update rfactsoc_variedad set descontado = 1 where codtipom = " & DBSet(vSocio.CodTipomAnt, "T") ' antes era 'FAC'
                    Sql3 = Sql3 & " and numfactu = " & DBSet(RS1!numfactu, "N")
                    Sql3 = Sql3 & " and fecfactu = " & DBSet(RS1!fecfactu, "F") & " and codvarie = " & DBSet(AntVarie, "N")
                    Sql3 = Sql3 & " and codcampo = " & DBSet(AntCampo, "N")
                    
                    conn.Execute Sql3
                    
                    ' insertamos en la tabla de anticipos de liquidacion venta campo
                    Sql3 = "insert into tmpFact_anticipos (codtipom, numfactu, fecfactu, codtipomanti, numfactuanti, fecfactuanti, codvarieanti, codcampoanti, baseimpo) values ("
                    Sql3 = Sql3 & DBSet(vSocio.CodTipomLiq, "T") & "," ' antes era 'FAL'
                    Sql3 = Sql3 & DBSet(numfactu, "N") & "," & DBSet(FecFac, "F") & ","
                    Sql3 = Sql3 & DBSet(vSocio.CodTipomAnt, "T") & "," ' antes era 'FAA'
                    Sql3 = Sql3 & DBSet(RS1!numfactu, "N") & "," & DBSet(RS1!fecfactu, "F") & ","
                    Sql3 = Sql3 & DBSet(AntVarie, "N") & "," & DBSet(AntCampo, "N") & "," & DBSet(RS1!imporvar, "N") & ")"
                    
                    conn.Execute Sql3
                    
                    RS1.MoveNext
                Wend
                
                Set RS1 = Nothing
                ' fin descontar anticipos
            
            End If
            
            
            If B Then
                AntVarie = ActVarie
                AntCampo = actCampo
                Kilos = 0
                Importe = 0
            End If
        End If
        
        If ActSocio <> AntSocio Then
            
            ImpoIva = Round2(baseimpo * PorcIva / 100, 2)
        
            Select Case DBLet(vSocio.TipoIRPF, "N")
                Case 0
                    ImpoReten = Round2((baseimpo + ImpoIva) * vParamAplic.PorcreteFacSoc / 100, 2)
                    BaseReten = (baseimpo + ImpoIva)
                    PorcReten = vParamAplic.PorcreteFacSoc
                Case 1
                    ImpoReten = Round2(baseimpo * vParamAplic.PorcreteFacSoc / 100, 2)
                    BaseReten = baseimpo
                    PorcReten = vParamAplic.PorcreteFacSoc
                Case 2
                    ImpoReten = 0
                    BaseReten = 0
                    PorcReten = 0
            End Select
            
            ImpoAFO = Round2((baseimpo + Anticipos) * vParamAplic.PorcenAFO / 100, 2)
            BaseAFO = baseimpo + Anticipos
            PorcAFO = vParamAplic.PorcenAFO
        
            TotalFac = baseimpo + ImpoIva - ImpoReten - ImpoAFO
            
            
            '[Monica]07/02/2012: indicamos si es una factura de liquidacion complementaria
            'insertar cabecera de factura
            B = InsertCabecera(tipoMov, CStr(numfactu), FecFac, , , False)
            
            '[Monica]04/01/2012: marcamos la factura como contabilizada y como pdte de recibir el nro de factura
            If B And vSocio.EmiteFact Then B = MarcarFactura(tipoMov, CStr(numfactu), FecFac, True)
            
            If B Then B = InsertResumen(tipoMov, CStr(numfactu))
            
'Mirar si quito lo de reclacular calidades
            If B Then B = RecalcularCalidades(tipoMov, CStr(numfactu), FecFac)
            
'Recalculo de todos los importes de tmpfact_calidades y tmpfact_variedades para que cuadre con la base de cabecera
            If B Then B = CuadrarBasesFactura(tipoMov, CStr(numfactu), FecFac, baseimpo)
            
            
            If B Then B = vTipoMov.IncrementarContador(tipoMov)
            
            If B Then
                AntSocio = ActSocio
                
                Set vSocio = Nothing
                Set vSocio = New cSocio
                If vSocio.LeerDatos(ActSocio) Then
                    If vSocio.LeerDatosSeccion(AntSocio, vParamAplic.Seccionhorto) Then
                        vPorcIva = ""
                        '[Monica]29/04/2011: INTERNAS
                        If vSocio.EsFactADVInt Then
                            vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vParamAplic.CodIvaExeADV, "N")
                        Else
                            vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vSocio.CodIva, "N")
                        End If
                        
                        PorcIva = CCur(ImporteSinFormato(vPorcIva))
                    End If
                    tipoMov = vSocio.CodTipomLiq
                                        
                End If
                baseimpo = 0
                BaseReten = 0
                ImpoIva = 0
                ImpoReten = 0
                TotalFac = 0
                BaseAFO = 0
                ImpoAFO = 0
                
                Anticipos = 0
                
                numfactu = vTipoMov.ConseguirContador(tipoMov)
                Do
                    numfactu = vTipoMov.ConseguirContador(tipoMov)
                    devuelve = DevuelveDesdeBDNew(cAgro, "rfactsoc", "numfactu", "codtipom", tipoMov, "T", , "numfactu", CStr(numfactu), "N", "fecfactu", FecFac, "F")
                    If devuelve <> "" Then
                        'Ya existe el contador incrementarlo
                        Existe = True
                        vTipoMov.IncrementarContador (tipoMov)
                        numfactu = vTipoMov.ConseguirContador(tipoMov)
                    Else
                        Existe = False
                    End If
                Loop Until Not Existe
           End If
        End If
        
        vPrecio = CCur(ComprobarCero(Precio))
        vImporte = vImporte + (DBLet(Rs!KilosNet, "N") * vPrecio)
        
        KilosCal = KilosCal + DBLet(Rs!KilosNet, "N")
        
        HayReg = True
        
        Rs.MoveNext
    Wend
    
    ' ultimo registro si ha entrado
    If B And HayReg Then
        ' insertar linea de calidad
        Kilos = Kilos + KilosCal
        Importe = Importe + vImporte
        
        baseimpo = baseimpo + vImporte
        
        
        If B Then B = InsertLineaCalidad(tipoMov, CStr(numfactu), FecFac, ActVarie, actCampo, CStr(ActCalid), CStr(DBLet(KilosCal, "N")), CStr(vImporte))
        
        
        If B Then ' descontamos el porcentaje de gastos de cooperativa
            GastosCoop = 0
            
            vPorcGasto = ""
            vPorcGasto = DevuelveDesdeBDNew(cAgro, "rcoope", "porcgast", "codcoope", vSocio.Cooperativa, "N")
            If vPorcGasto = "" Then vPorcGasto = "0"
            
            '[Monica]25/02/2011: Sólo hay gastos si no es complementaria ( Añadido el if )
            GastosCoop = Round2(Importe * CCur(ImporteSinFormato(vPorcGasto)) / 100, 2)
            Importe = Importe - GastosCoop
            baseimpo = baseimpo - GastosCoop
        End If
        
        If B Then ' descontamos los gastos de los albaranes
            '[Monica]25/02/2011: Sólo hay gastos si no es complementaria ( Añadido el if )
            GastosAlb = ObtenerGastosAlbaranes(AntSocio, AntVarie, AntCampo, "rhisfruta", "rhisfruta.numalbar = " & Albaran, 1)
            Importe = Importe - GastosAlb
            baseimpo = baseimpo - GastosAlb
        End If
        
        '[Monica]08/04/2010: grabamos los albaranes que intervienen en la linea de factura
        If B Then
            B = InsertarAlbaranesFactura(tipoMov, CStr(numfactu), FecFac, AntSocio, AntVarie, AntCampo, "rhisfruta", "rhisfruta.numalbar = " & Albaran, 0)
        End If
                    
                    
        ' insertar linea de variedad
        If B Then B = InsertLinea(tipoMov, CStr(numfactu), FecFac, CStr(ActVarie), CStr(actCampo), CStr(Kilos), CStr(Importe), CStr(GastosAlb))
        
        ' tenemos que descontar los anticipos que tengamos para ello
        If B Then
            Sql2 = "select rfactsoc_variedad.imporvar, rfactsoc_variedad.numfactu, rfactsoc_variedad.fecfactu "
            Sql2 = Sql2 & " from rfactsoc_variedad INNER JOIN rfactsoc ON rfactsoc_variedad.codtipom = rfactsoc.codtipom and "
            Sql2 = Sql2 & " rfactsoc_variedad.numfactu = rfactsoc.numfactu and rfactsoc_variedad.fecfactu = rfactsoc.fecfactu "
            Sql2 = Sql2 & " where rfactsoc_variedad.codtipom = " & DBSet(vSocio.CodTipomAnt, "T") ' antes era 'FAA' "
            Sql2 = Sql2 & " and rfactsoc.codsocio = " & DBSet(ActSocio, "N")
            Sql2 = Sql2 & " and codvarie = " & DBSet(ActVarie, "N")
            Sql2 = Sql2 & " and codcampo = " & DBSet(actCampo, "N")
            Sql2 = Sql2 & " and rfactsoc_variedad.descontado = 0"
            
            Set RS1 = New ADODB.Recordset
            RS1.Open Sql2, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
            
            While Not RS1.EOF
                baseimpo = baseimpo - DBLet(RS1.Fields(0).Value, "N")
                Anticipos = Anticipos + DBLet(RS1.Fields(0).Value, "N")
                
                ' indicamos que los anticipos ya han sido descontados
                Sql3 = "update rfactsoc_variedad set descontado = 1 where codtipom = " & DBSet(vSocio.CodTipomAnt, "T") ' antes era 'FAC'
                Sql3 = Sql3 & " and numfactu = " & DBSet(RS1!numfactu, "N")
                Sql3 = Sql3 & " and fecfactu = " & DBSet(RS1!fecfactu, "F") & " and codvarie = " & DBSet(ActVarie, "N")
                Sql3 = Sql3 & " and codcampo = " & DBSet(actCampo, "N")
                
                conn.Execute Sql3
                
                ' insertamos en la tabla de anticipos de liquidacion venta campo
                Sql3 = "insert into tmpFact_anticipos (codtipom, numfactu, fecfactu, codtipomanti, numfactuanti, fecfactuanti, codvarieanti, codcampoanti, baseimpo) values ("
                Sql3 = Sql3 & DBSet(vSocio.CodTipomLiq, "T") & "," ' antes era 'FAL'
                Sql3 = Sql3 & DBSet(numfactu, "N") & "," & DBSet(FecFac, "F") & ","
                Sql3 = Sql3 & DBSet(vSocio.CodTipomAnt, "T") & "," ' antes era 'FAA'
                Sql3 = Sql3 & DBSet(RS1!numfactu, "N") & "," & DBSet(RS1!fecfactu, "F") & ","
                Sql3 = Sql3 & DBSet(ActVarie, "N") & "," & DBSet(actCampo, "N") & "," & DBSet(RS1!imporvar, "N") & ")"
                
                conn.Execute Sql3
                
                RS1.MoveNext
            Wend
            
            Set RS1 = Nothing
            ' fin descontar anticipos
        
        End If
        
        
        ImpoIva = Round2(baseimpo * PorcIva / 100, 2)

        Select Case DBLet(vSocio.TipoIRPF, "N")
            Case 0
                ImpoReten = Round2((baseimpo + ImpoIva) * vParamAplic.PorcreteFacSoc / 100, 2)
                BaseReten = (baseimpo + ImpoIva)
                PorcReten = vParamAplic.PorcreteFacSoc
            Case 1
                ImpoReten = Round2(baseimpo * vParamAplic.PorcreteFacSoc / 100, 2)
                BaseReten = baseimpo
                PorcReten = vParamAplic.PorcreteFacSoc
            Case 2
                ImpoReten = 0
                BaseReten = 0
                PorcReten = 0
        End Select
        
        ImpoAFO = Round2((baseimpo + Anticipos) * vParamAplic.PorcenAFO / 100, 2)
        BaseAFO = baseimpo + Anticipos
        PorcAFO = vParamAplic.PorcenAFO

        TotalFac = baseimpo + ImpoIva - ImpoReten - ImpoAFO
        
        
        vParamAplic.UltFactLiq = numfactu
        
        '[Monica]07/02/2012: indicamos si es una factura de liquidacion complementaria
        'insertar cabecera de factura
        B = InsertCabecera(tipoMov, CStr(numfactu), FecFac, , , False)
        
        '[Monica]04/01/2012: marcamos la factura como contabilizada y como pdte de recibir el nro de factura
        If B And vSocio.EmiteFact Then B = MarcarFactura(tipoMov, CStr(numfactu), FecFac, True)
        
        If B Then B = InsertResumen(tipoMov, CStr(numfactu))

'Mirar si quito lo de reclacular calidades
        If B Then B = RecalcularCalidades(tipoMov, CStr(numfactu), FecFac)
        
'Recalculo de todos los importes de rfactsoc_calidades y rfactsoc_variedades para que cuadre con la base de cabecera
        If B Then B = CuadrarBasesFactura(tipoMov, CStr(numfactu), FecFac, baseimpo)
        
        If B Then B = vTipoMov.IncrementarContador(tipoMov)
        
        'pasamos las temporales a las tablas
        If B Then B = PasarTemporales()
        
        If B Then B = (vParamAplic.Modificar = 1)
    End If
    
'    BorrarTMPs
    
    vSeccion.CerrarConta
    Set vSeccion = Nothing
    Set vSocio = Nothing
    
eFacturacion:
    If Err.Number <> 0 Or Not B Then
        conn.RollbackTrans
        FacturacionLiquidacionDirecta = False
    Else
        conn.CommitTrans
        FacturacionLiquidacionDirecta = True
    End If
End Function



Public Sub RecalculoBasesIvaFactura(ByRef Rs As ADODB.Recordset, ByRef ImpTot As Variant, ByRef Tipiva As Variant, ByRef Impbas As Variant, ByRef impiva As Variant, ByRef PorIva As Variant, ByRef TotFac As Currency, ByRef ImpREC As Variant, ByRef PorRec As Variant, ByRef PorRet As Variant, ByRef ImpRet As Variant, TipoRetencion As Integer, Optional Socio As String, Optional Tipo As String)

    Dim i As Integer
    Dim SQL As String
    Dim baseimpo As Dictionary
    Dim CodIva As Integer
    Dim totimp As Currency
    Dim Base As Currency
    
    
    
    Set baseimpo = New Dictionary

    ' inicializamos los importes de los totales de la cabecera
    
    TotFac = 0
    totimp = 0
    Base = 0
    ImpRet = 0
    For i = 0 To 2
         Tipiva(i) = 0
         ImpTot(i) = 0
         Impbas(i) = 0
         impiva(i) = 0
         PorIva(i) = 0
         PorRec(i) = 0
         ImpREC(i) = 0
    Next i

    ' recorremos todas las lineas de la factura
    If Not Rs.EOF Then Rs.MoveFirst
    While Not Rs.EOF
        CodIva = DBLet(Rs!TipoIVA, "N") ' DevuelveDesdeBDNewFac("tiposiva", "codigiva", "sartic", "codartic", DBLet(RS!codartic), "N")
        baseimpo(Val(CodIva)) = DBLet(baseimpo(Val(CodIva)), "N") + DBLet(Rs!Importe, "N")

        Rs.MoveNext
    Wend

    For i = 0 To baseimpo.Count - 1
        If i <= 2 Then
            Tipiva(i) = baseimpo.Keys(i)
            Impbas(i) = baseimpo.Items(i)
 
            PorIva(i) = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", CStr(Tipiva(i)), "N")
            PorRec(i) = DevuelveDesdeBDNew(cConta, "tiposiva", "porcerec", "codigiva", CStr(Tipiva(i)), "N")
            impiva(i) = DBLet(Round2(Impbas(i) * PorIva(i) / 100, 2), "N")
            ImpREC(i) = DBLet(Round2(Impbas(i) * PorRec(i) / 100, 2), "N")
            ImpTot(i) = Impbas(i) + impiva(i) + ImpREC(i)
            TotFac = TotFac + ImpTot(i)
 
'antes el iva estaba incluido
'            PorIva(i) = DevuelveDesdeBDNewFac(cConta, "tiposiva", "porceiva", "codigiva", CStr(Tipiva(i)), "N")
'            Impbas(i) = Round2(Imptot(i) / (1 + (PorIva(i) / 100)), 2)
'            impiva(i) = Imptot(i) - Impbas(i)
'            TotFac = TotFac + Imptot(i)
        
        
        End If
    Next i
    'si hay retencion la calculamos
    If PorRet <> 0 Then
        Base = 0
        
        If Tipo = "FVP" Then  ' facturas varias de proveedor
            ' la base de retencion va a depender del tipo de socio (modulos, estimacion directa o entidad)
            If Socio <> "" Then SQL = DevuelveValor("select tipoirpf from rsocios where codsocio = " & DBSet(Socio, "N"))
            Select Case SQL
                Case 0
                    For i = 0 To baseimpo.Count - 1
                        Base = Base + Impbas(i) + impiva(i)
                    Next i
                Case 1
                    For i = 0 To baseimpo.Count - 1
                        Base = Base + Impbas(i)
                    Next i
                Case 2
                
            End Select
        
        Else
            Dim TipoRet As Integer
            TipoRet = DevuelveValor("select tipo from usuarios.wtiporeten where codigo = " & DBSet(TipoRetencion, "N"))
        
            For i = 0 To baseimpo.Count - 1
                Base = Base + Impbas(i)
                
                '[Monica]20/08/2018: añadido lo del tipo de retencion
                If TipoRet = 1 Then
                    Base = Base + impiva(i)
                End If
            Next i
        End If
        ImpRet = Round2(Base * PorRet / 100, 2)
        TotFac = TotFac - ImpRet
    Else
        ImpRet = 0
    End If
End Sub


Public Function FacturacionAnticiposAlmazaraCastelduc(cTabla As String, cWhere As String, FecFac As String, Pb1 As ProgressBar) As Boolean
Dim SQL As String
Dim Rs As ADODB.Recordset

Dim AntSocio As String
Dim AntVarie As String
Dim ActSocio As String
Dim ActVarie As String
Dim actCampo As String
Dim AntCampo As String
Dim ActCalid As String
Dim AntCalid As String

Dim HayReg As Boolean

Dim NumError As Long
Dim vImporte As Currency
Dim vPorcIva As String

Dim PrimFac As String
Dim UltFac As String

Dim tipoMov As String
Dim B As Boolean
Dim vSeccion As CSeccion
Dim Kilos As Currency
Dim KilosCal As Currency
Dim Importe As Currency

Dim devuelve As String
Dim Existe As Boolean

'Dim baseimpo As Currency
'Dim BaseReten As Currency
Dim Neto As Currency
'Dim ImpoIva As Currency
'Dim ImpoReten As Currency
'Dim TotalFac As Currency
Dim Recolect As Byte
Dim vPrecio As Currency
Dim campo As String

    On Error GoTo eFacturacion

    FacturacionAnticiposAlmazaraCastelduc = False
    
    tipoMov = "FNZ"
    
    BorrarTMPs
    B = CrearTMPs()
    If Not B Then
         Exit Function
    End If
    
    conn.BeginTrans
    
    
    SQL = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute SQL

    SQL = "SELECT  rhisfruta.codsocio, rhisfruta.codvarie, "
    SQL = SQL & "rprecios.precioindustria, "
    SQL = SQL & "rprecios.tipofact, sum(rhisfruta.kilosnet) as kilosnet "
    SQL = SQL & " FROM  " & cTabla

    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        SQL = SQL & " WHERE " & cWhere
    End If
    ' ordenado por socio, variedad, campo, calidad
    SQL = SQL & " group by rhisfruta.codsocio, rhisfruta.codvarie, rprecios.precioindustria,rprecios.tipofact"
    SQL = SQL & " order by rhisfruta.codsocio, rhisfruta.codvarie, rprecios.precioindustria,rprecios.tipofact"
    
    Set vSeccion = New CSeccion
    
    If vSeccion.LeerDatos(vParamAplic.SeccionAlmaz) Then
        If Not vSeccion.AbrirConta Then
            Exit Function
        End If
    End If
    
    HayReg = False
    
   ' en almazara no se insertan campos: metemos el minimo codcampo sin condiciones
    campo = "0" 'DevuelveValor("select min(codcampo) from rcampos")
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    If Not Rs.EOF Then
        AntSocio = CStr(DBLet(Rs!Codsocio, "N"))
        AntVarie = CStr(DBLet(Rs!Codvarie, "N"))
        
        ActSocio = CStr(DBLet(Rs!Codsocio, "N"))
        ActVarie = CStr(DBLet(Rs!Codvarie, "N"))
   
        Set vSocio = New cSocio
        If vSocio.LeerDatos(ActSocio) Then
            If vSocio.LeerDatosSeccion(ActSocio, vParamAplic.SeccionAlmaz) Then
                baseimpo = 0
                BaseReten = 0
                ImpoIva = 0
                ImpoReten = 0
                TotalFac = 0
                BaseAFO = 0
                ImpoAFO = 0
                PorcAFO = 0
                
                Kilos = 0
                Importe = 0
                
                KilosCal = 0
                
                vPorcIva = ""
                vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vSocio.CodIva, "N")
                PorcIva = CCur(ImporteSinFormato(vPorcIva))
                
                tipoMov = vSocio.CodTipomAntAlmz
                
                Set vTipoMov = New CTiposMov
                If vTipoMov.Leer(tipoMov) Then
                    numfactu = vTipoMov.ConseguirContador(tipoMov)
                    Do
                        numfactu = vTipoMov.ConseguirContador(tipoMov)
                        devuelve = DevuelveDesdeBDNew(cAgro, "rfactsoc", "numfactu", "codtipom", tipoMov, "T", , "numfactu", CStr(numfactu), "N", "fecfactu", FecFac, "F")
                        If devuelve <> "" Then
                            'Ya existe el contador incrementarlo
                            Existe = True
                            vTipoMov.IncrementarContador (tipoMov)
                            numfactu = vTipoMov.ConseguirContador(tipoMov)
                        Else
                            Existe = False
                        End If
                    Loop Until Not Existe
                    
                    vParamAplic.PrimFactAntAlmz = numfactu
                Else
                    B = False
                End If
            End If
        End If
    End If
    
    While Not Rs.EOF And B
        ActVarie = DBLet(Rs!Codvarie, "N")
        ActSocio = DBSet(Rs!Codsocio, "N")
        
        If (AntVarie <> ActVarie Or AntSocio <> ActSocio) Then
            ' kilos e importe por variedad campo
            
            baseimpo = baseimpo + Importe
            
            ' insertar linea de variedad, campo
            B = InsertLinea(tipoMov, CStr(numfactu), FecFac, AntVarie, campo, CStr(Kilos), CStr(Importe), "0")
            
            If B Then
                AntVarie = ActVarie
                Kilos = 0
                Importe = 0
            End If
        End If
        
        If ActSocio <> AntSocio Then
            ImpoIva = Round2(baseimpo * PorcIva / 100, 2)
        
            Select Case DBLet(vSocio.TipoIRPF, "N")
                Case 0
                    ImpoReten = Round2((baseimpo + ImpoIva) * vParamAplic.PorcreteFacSoc / 100, 2)
                    BaseReten = (baseimpo + ImpoIva)
                    PorcReten = vParamAplic.PorcreteFacSoc
                Case 1
                    ImpoReten = Round2(baseimpo * vParamAplic.PorcreteFacSoc / 100, 2)
                    BaseReten = baseimpo
                    PorcReten = vParamAplic.PorcreteFacSoc
                Case 2
                    ImpoReten = 0
                    BaseReten = 0
                    PorcReten = 0
            End Select
        
            TotalFac = baseimpo + ImpoIva - ImpoReten
            
            IncrementarProgresNew Pb1, 1
            
            'insertar cabecera de factura
            B = InsertCabecera(tipoMov, CStr(numfactu), FecFac)
            
            If B Then B = InsertResumen(tipoMov, CStr(numfactu))
            
            If B Then B = vTipoMov.IncrementarContador(tipoMov)
            
            If B Then
                AntSocio = ActSocio
                
                Set vSocio = Nothing
                Set vSocio = New cSocio
                If vSocio.LeerDatos(ActSocio) Then
                    If vSocio.LeerDatosSeccion(AntSocio, vParamAplic.SeccionAlmaz) Then
                        vPorcIva = ""
                        vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vSocio.CodIva, "N")
                        PorcIva = CCur(ImporteSinFormato(vPorcIva))
                    End If
                    
                    tipoMov = vSocio.CodTipomAntAlmz
                End If
                baseimpo = 0
                BaseReten = 0
                Neto = 0
                ImpoIva = 0
                ImpoReten = 0
                TotalFac = 0
                
                If vTipoMov.Leer(tipoMov) Then
                    numfactu = vTipoMov.ConseguirContador(tipoMov)
                    Do
                        numfactu = vTipoMov.ConseguirContador(tipoMov)
                        devuelve = DevuelveDesdeBDNew(cAgro, "rfactsoc", "numfactu", "codtipom", tipoMov, "T", , "numfactu", CStr(numfactu), "N", "fecfactu", FecFac, "F")
                        If devuelve <> "" Then
                            'Ya existe el contador incrementarlo
                            Existe = True
                            vTipoMov.IncrementarContador (tipoMov)
                            numfactu = vTipoMov.ConseguirContador(tipoMov)
                        Else
                            Existe = False
                        End If
                    Loop Until Not Existe
                Else
                    B = False
                End If
           End If
        End If
        
'[Monica]añadidas estas 3 lineas eliminada la del precio para el anticipo
        vPrecio = DBLet(Rs!precioindustria, "N")
        vImporte = Round2(DBLet(Rs!KilosNet, "N") * vPrecio, 2)
        
        B = InsertLineaAlbaranNew(tipoMov, CStr(numfactu), FecFac, Rs, CStr(vPrecio), CStr(vImporte), cTabla, cWhere)
    
'        vPrecio = DBLet(Rs!precioindustria, "N")
'[Monica] hasta aqui

        Importe = Importe + Round2(DBLet(Rs!KilosNet, "N") * Rs!precioindustria, 2)
        
        Kilos = Kilos + DBLet(Rs!KilosNet, "N")
        
        HayReg = True
        
        Rs.MoveNext
    Wend
    
    ' ultimo registro si ha entrado
    If B And HayReg Then
        baseimpo = baseimpo + Importe
        
        ' insertar linea de variedad
        If B Then B = InsertLinea(tipoMov, CStr(numfactu), FecFac, CStr(ActVarie), CStr(campo), CStr(Kilos), CStr(Importe), "0")
        
        ImpoIva = Round2(baseimpo * PorcIva / 100, 2)

        Select Case DBLet(vSocio.TipoIRPF, "N")
            Case 0
                ImpoReten = Round2((baseimpo + ImpoIva) * vParamAplic.PorcreteFacSoc / 100, 2)
                BaseReten = (baseimpo + ImpoIva)
                PorcReten = vParamAplic.PorcreteFacSoc
            Case 1
                ImpoReten = Round2(baseimpo * vParamAplic.PorcreteFacSoc / 100, 2)
                BaseReten = baseimpo
                PorcReten = vParamAplic.PorcreteFacSoc
            Case 2
                ImpoReten = 0
                BaseReten = 0
                PorcReten = 0
        End Select
        
        TotalFac = baseimpo + ImpoIva - ImpoReten
        
        IncrementarProgresNew Pb1, 1
        
        vParamAplic.UltFactAntAlmz = numfactu
        
        'insertar cabecera de factura
        B = InsertCabecera(tipoMov, CStr(numfactu), FecFac)
        
        If B Then B = InsertResumen(tipoMov, CStr(numfactu))
        
        If B Then B = vTipoMov.IncrementarContador(tipoMov)
        
        'pasamos las temporales a las tablas
        If B Then B = PasarTemporales()
        
        If B Then B = (vParamAplic.Modificar = 1)
    End If
    
'    BorrarTMPs
    
    vSeccion.CerrarConta
    Set vSeccion = Nothing
    Set vSocio = Nothing
    
eFacturacion:
    If Err.Number <> 0 Or Not B Then
        conn.RollbackTrans
        FacturacionAnticiposAlmazaraCastelduc = False
    Else
        conn.CommitTrans
        FacturacionAnticiposAlmazaraCastelduc = True
    End If
End Function


'Insertar Linea de factura (albaran)
Public Function InsertLineaAlbaranNew(tipoMov As String, numfactu As String, FecFac As String, ByRef Rs As ADODB.Recordset, Precio As String, Importe As String, cTabla As String, cWhere As String) As Boolean
'(rfactsoc_albaran)
'codcampo tiene valor cuando venimos de almazara que hemos tenido que buscarlo porque en el cursor Rs no lo tenemos
Dim GastosAlb As Currency
Dim Tipo As String

    Dim SQL As String
    Dim ImpLinea As Currency
    
    On Error GoTo eInsertLinea
    
    MensError = ""
    
    InsertLineaAlbaranNew = False
    
    Tipo = DevuelveDesdeBDNew(cAgro, "usuarios.stipom", "tipodocu", "codtipom", tipoMov, "T")
    If CInt(Tipo) = 7 Then ' si se trata de un anticipo de almazara no descontamos gastos
        GastosAlb = 0
    Else
        GastosAlb = DevuelveValor("select sum(importe) from rhisfruta_gastos where numalbar = " & DBSet(Rs!numalbar, "N"))
    End If
    
    'insertamos el albaran
    SQL = "insert into tmpfact_albaran (codtipom, numfactu, fecfactu, numalbar, fecalbar, "
    SQL = SQL & "codvarie, codcampo, kilosbru, kilosnet, grado, precio, importe, imporgasto) select "
    SQL = SQL & "'" & tipoMov & "'," & DBSet(numfactu, "N") & "," & DBSet(FecFac, "F") & ", numalbar, fecalbar, rhisfruta.codvarie, 0, rhisfruta.kilosbru, rhisfruta.kilosnet, "
    SQL = SQL & " prestimado," & DBSet(Precio, "N") & ",round(" & DBSet(Precio, "N") & " * kilosnet,2),0"
    SQL = SQL & " from  " & cTabla
    SQL = SQL & " where rhisfruta.codsocio = " & DBSet(Rs!Codsocio, "N")
    SQL = SQL & " and rhisfruta.codvarie = " & DBSet(Rs!Codvarie, "N") & " and " & cWhere
    
    conn.Execute SQL
    InsertLineaAlbaranNew = True
    Exit Function
    
eInsertLinea:
    If Err.Number <> 0 Then
        MensError = "Se ha producido un error en la inserción de la linea de albaran de factura "
        MuestraError Err.Number, MensError, Err.Description
    End If
End Function




Public Function EstaFacturado(Albaran As Long) As Boolean
Dim SQL As String
Dim Facturado As Boolean
    
    EstaFacturado = False
    
    SQL = "select count(*) from rfactsoc_albaran where numalbar = " & DBSet(Albaran, "N")
    Facturado = (TotalRegistros(SQL) <> 0)
    
    SQL = "select count(*) from rlifter where numalbar = " & DBSet(Albaran, "N")
    
    EstaFacturado = Facturado Or (TotalRegistros(SQL) <> 0)

End Function



'*****
'   proceso en donde se crea unicamente una factura de anticipo que posteriormente
'   se descontará en la factura de liquidacion de venta campo
'

'[Monica]07/11/2013: añadido el parametro de si es tercero solo para Picassent

Public Function FacturaAnticipoSinEntrada(Socio As String, campo As String, Importe As String, FecFac As String, Optional EsTercero As Boolean, Optional Varie As String) As Long
Dim SQL As String
Dim Sql2 As String
Dim Rs As ADODB.Recordset
Dim RS1 As ADODB.Recordset

Dim AntSocio As String
Dim ActSocio As String

Dim HayReg As Boolean
Dim Variedad As String


Dim NumError As Long
Dim vImporte As Currency
Dim vPorcIva As String

Dim PrimFac As String
Dim UltFac As String
Dim tipoMov As String

Dim Sql3 As String
Dim devuelve As String
Dim B As Boolean
Dim vSeccion As CSeccion
Dim Existe As Boolean

    On Error GoTo eFacturacion
    
'08052009 antes dentro de transaccion
    BorrarTMPs
    B = CrearTMPs()
    If Not B Then
         Exit Function
    End If
'08052009

    conn.BeginTrans

    '[Monica]07/11/2013: añadida la opcion de si es tercero
    If EsTercero Then
        tipoMov = "FAT"
    Else
        tipoMov = "FAA"
    End If

    SQL = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute SQL

    '[Monica]02/11/2017: en el caso de que me pongan la variedad puede ser relacionada, no hace falta buscarla
    If Varie = "" Then
        SQL = "SELECT rcampos.codvarie from rcampos where codcampo = " & DBSet(campo, "N")
        Variedad = DevuelveValor(SQL)
    Else
        Variedad = Varie
    End If
    
    Set vSeccion = New CSeccion
    
    If vSeccion.LeerDatos(vParamAplic.Seccionhorto) Then
        If Not vSeccion.AbrirConta Then
            Exit Function
        End If
    End If
    
    Set vSocio = New cSocio
    If vSocio.LeerDatos(Socio) Then
        If vSocio.LeerDatosSeccion(Socio, vParamAplic.Seccionhorto) Then
            baseimpo = CCur(Importe)
            BaseReten = 0
            ImpoIva = 0
            ImpoReten = 0
            TotalFac = 0
            
            Anticipos = 0
            
            vPorcIva = ""
            vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vSocio.CodIva, "N")
            PorcIva = CCur(ImporteSinFormato(vPorcIva))
            
            '[Monica]07/11/2013: depende de si es tercero
            If EsTercero Then
                tipoMov = "FAT"
            Else
                tipoMov = vSocio.CodTipomAnt
            End If
            
            Set vTipoMov = New CTiposMov
            numfactu = vTipoMov.ConseguirContador(tipoMov)
            Do
                numfactu = vTipoMov.ConseguirContador(tipoMov)
                devuelve = DevuelveDesdeBDNew(cAgro, "rfactsoc", "numfactu", "codtipom", tipoMov, "T", , "numfactu", CStr(numfactu), "N", "fecfactu", FecFac, "F")
                If devuelve <> "" Then
                    'Ya existe el contador incrementarlo
                    Existe = True
                    vTipoMov.IncrementarContador (tipoMov)
                    numfactu = vTipoMov.ConseguirContador(tipoMov)
                Else
                    Existe = False
                End If
            Loop Until Not Existe
            
            vParamAplic.PrimFactAnt = numfactu
            
            ImpoIva = Round2(baseimpo * PorcIva / 100, 2)
        
            Select Case DBLet(vSocio.TipoIRPF, "N")
                Case 0
                    ImpoReten = Round2((baseimpo + ImpoIva) * vParamAplic.PorcreteFacSoc / 100, 2)
                    BaseReten = (baseimpo + ImpoIva)
                    PorcReten = vParamAplic.PorcreteFacSoc
                Case 1
                    ImpoReten = Round2(baseimpo * vParamAplic.PorcreteFacSoc / 100, 2)
                    BaseReten = baseimpo
                    PorcReten = vParamAplic.PorcreteFacSoc
                Case 2
                    ImpoReten = 0
                    BaseReten = 0
                    PorcReten = 0
            End Select
        
            TotalFac = baseimpo + ImpoIva - ImpoReten
            
            ' insertar linea de variedad, campo
            B = InsertLinea(tipoMov, CStr(numfactu), FecFac, CStr(DBLet(Variedad, "N")), CStr(DBLet(campo, "N")), CStr(DBLet(0, "N")), CStr(DBLet(Importe, "N")), 0)
            
            If B Then
                ' insertamos los totales en la calidad venta campo de la variedad (rfactsoc_calidad)
                Sql2 = "select codcalid from rcalidad where codvarie = " & DBSet(Variedad, "N")
'                Sql2 = Sql2 & " and tipcalid = 2 " ' calidad de venta campo
                
                Set RS1 = New ADODB.Recordset
                RS1.Open Sql2, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
                
                If Not RS1.EOF Then
                    B = InsertLineaCalidad(tipoMov, CStr(numfactu), FecFac, CStr(DBLet(Variedad, "N")), CStr(DBLet(campo, "N")), CStr(DBLet(RS1!codcalid, "N")), CStr(DBLet(0, "N")), CStr(DBLet(Importe, "N")))
                End If
                Set RS1 = Nothing
            End If
            
            'insertar cabecera de factura
            If B Then B = InsertCabecera(tipoMov, CStr(numfactu), FecFac)
            
            '[Monica]07/11/2013: si es tercero he de marcarla como contabilizada
            '                    en ppio solo es para Picassent
            If EsTercero Then
                If B Then B = MarcarFactura(tipoMov, CStr(numfactu), FecFac, True)
            End If
            
            If B Then B = InsertResumen(tipoMov, CStr(numfactu))
            
            If B Then B = vTipoMov.IncrementarContador(tipoMov)
            
            vParamAplic.UltFactAnt = numfactu
            
            'pasamos las temporales a las tablas
            If B Then B = PasarTemporales()
            
            If B Then B = (vParamAplic.Modificar = 1)
            
        End If
    
        BorrarTMPs
    End If
    
    vSeccion.CerrarConta
    Set vSeccion = Nothing
    Set vSocio = Nothing
    
eFacturacion:
    If Err.Number <> 0 Or Not B Then
        conn.RollbackTrans
        FacturaAnticipoSinEntrada = False
    Else
        conn.CommitTrans
        FacturaAnticipoSinEntrada = True
    End If
End Function

'################################################################################################
'########## NUEVA FACTURACION DE ANTICIPOS Y LIQUIDACIONES PARA PICASSENT ( ahora es por tramos )
'################################################################################################
' Igual que lo tenia antes, pero añadiendo un paso previo de precio segun fecha de albaran

Public Function FacturacionAnticiposPicassentNew(cTabla As String, cWhere As String, FecFac As String, Pb1 As ProgressBar, DescontarFVarias As Boolean, EsTerceros As Boolean) As Boolean
Dim SQL As String
Dim Rs As ADODB.Recordset

Dim AntSocio As String
Dim AntVarie As String
Dim ActSocio As String
Dim ActVarie As String
Dim actCampo As String
Dim AntCampo As String
Dim ActCalid As String
Dim AntCalid As String

Dim HayReg As Boolean

Dim NumError As Long
Dim vImporte As Currency
Dim vPorcIva As String

Dim PrimFac As String
Dim UltFac As String

Dim tipoMov As String
Dim B As Boolean
Dim vSeccion As CSeccion
Dim Kilos As Currency
Dim KilosCal As Currency
Dim Importe As Currency
Dim Bonifica As Currency

Dim devuelve As String
Dim Existe As Boolean

'Dim baseimpo As Currency
'Dim BaseReten As Currency
Dim Neto As Currency
'Dim ImpoIva As Currency
'Dim ImpoReten As Currency
'Dim TotalFac As Currency
Dim Recolect As Byte
Dim vPrecio As Currency
Dim vBonifica As Currency
Dim PorcBoni As Currency
Dim PorcComi As Currency

Dim HayPrecio As Boolean


    On Error GoTo eFacturacion

    FacturacionAnticiposPicassentNew = False
    
    If EsTerceros Then
        tipoMov = "FAT" ' facturas de anticipos de terceros
    Else
        tipoMov = "FAA"
    End If
    
    BorrarTMPs
    B = CrearTMPs()
    If Not B Then
         Exit Function
    End If
    
    conn.BeginTrans
    
    SQL = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute SQL

    SQL = "SELECT  rhisfruta.codsocio, rhisfruta.codvarie,"
    SQL = SQL & "rhisfruta.codcampo, rhisfruta_clasif.codcalid, rhisfruta.numalbar, rhisfruta.fecalbar, sum(coalesce(rhisfruta_clasif.kilosnet,0)) as kilosnet "
    SQL = SQL & " FROM  " & cTabla

    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        SQL = SQL & " WHERE " & cWhere
    End If
    ' ordenado por socio, variedad, campo, calidad
    SQL = SQL & " group by rhisfruta.codsocio, rhisfruta.codvarie, rhisfruta.codcampo, rhisfruta_clasif.codcalid, rhisfruta.numalbar, rhisfruta.fecalbar "
    SQL = SQL & " order by rhisfruta.codsocio, rhisfruta.codvarie, rhisfruta.codcampo, rhisfruta_clasif.codcalid, rhisfruta.numalbar, rhisfruta.fecalbar "
    
    Set vSeccion = New CSeccion
    
    If vSeccion.LeerDatos(vParamAplic.Seccionhorto) Then
        If Not vSeccion.AbrirConta Then
            Exit Function
        End If
    End If
    
    HayReg = False
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    If Not Rs.EOF Then
        AntSocio = CStr(DBLet(Rs!Codsocio, "N"))
        AntVarie = CStr(DBLet(Rs!Codvarie, "N"))
        AntCampo = CStr(DBLet(Rs!codCampo, "N"))
        AntCalid = CStr(DBLet(Rs!codcalid, "N"))
        
        ActSocio = CStr(DBLet(Rs!Codsocio, "N"))
        ActVarie = CStr(DBLet(Rs!Codvarie, "N"))
        actCampo = CStr(DBLet(Rs!codCampo, "N"))
        ActCalid = CStr(DBLet(Rs!codcalid, "N"))
    
        Set vSocio = New cSocio
        If vSocio.LeerDatos(ActSocio) Then
            If vSocio.LeerDatosSeccion(ActSocio, vParamAplic.Seccionhorto) Then
                baseimpo = 0
                BaseReten = 0
                ImpoIva = 0
                ImpoReten = 0
                TotalFac = 0
                BaseAFO = 0
                ImpoAFO = 0
                PorcAFO = 0
                
                Kilos = 0
                Importe = 0
                Bonifica = 0
                
                KilosCal = 0
                
                vPorcIva = ""
                vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vSocio.CodIva, "N")
                PorcIva = CCur(ImporteSinFormato(vPorcIva))
                
                If EsTerceros Then
                    tipoMov = "FAT"
                Else
                    tipoMov = vSocio.CodTipomAnt
                End If
                
                Set vTipoMov = New CTiposMov
                
                numfactu = vTipoMov.ConseguirContador(tipoMov)
                Do
                    numfactu = vTipoMov.ConseguirContador(tipoMov)
                    devuelve = DevuelveDesdeBDNew(cAgro, "rfactsoc", "numfactu", "codtipom", tipoMov, "T", , "numfactu", CStr(numfactu), "N", "fecfactu", FecFac, "F")
                    If devuelve <> "" Then
                        'Ya existe el contador incrementarlo
                        Existe = True
                        vTipoMov.IncrementarContador (tipoMov)
                        numfactu = vTipoMov.ConseguirContador(tipoMov)
                    Else
                        Existe = False
                    End If
                Loop Until Not Existe
                
                vParamAplic.PrimFactAnt = numfactu
                
            End If
        End If
    End If
    
    While Not Rs.EOF And B
        ActCalid = DBLet(Rs!codcalid, "N")
        ActVarie = DBLet(Rs!Codvarie, "N")
        actCampo = DBSet(Rs!codCampo, "N")
        ActSocio = DBSet(Rs!Codsocio, "N")
        
        If (ActCalid <> AntCalid Or AntCampo <> actCampo Or AntVarie <> ActVarie Or AntSocio <> ActSocio) Then
            ' kilos e importe por variedad campo
            '[Monica]24/02/2014: añadida condicion
            If HayPrecio Then
                Kilos = Kilos + KilosCal
                Importe = Importe + vImporte
                Bonifica = Bonifica + vBonifica
                
                baseimpo = baseimpo + vImporte
                
                B = InsertLineaCalidad(tipoMov, CStr(numfactu), FecFac, AntVarie, AntCampo, CStr(AntCalid), CStr(DBLet(KilosCal, "N")), CStr(vImporte), CStr(vBonifica))
            End If
            KilosCal = 0
            vImporte = 0
            vBonifica = 0
            
            AntCalid = ActCalid
        End If
        
        If (ActVarie <> AntVarie Or actCampo <> AntCampo Or ActSocio <> AntSocio) Then
            ' insertar linea de variedad, campo
            '[Monica]24/02/2014: añadida condicion
            If Kilos <> 0 Then
                B = InsertLinea(tipoMov, CStr(numfactu), FecFac, AntVarie, AntCampo, CStr(Kilos), CStr(Importe), "0", CStr(Bonifica))
            End If
            
            If B Then
                AntVarie = ActVarie
                AntCampo = actCampo
                Kilos = 0
                Importe = 0
                Bonifica = 0
            End If
        End If
        
        If ActSocio <> AntSocio Then
            
           '[Monica]24/02/2014: añadida condicion
            If baseimpo <> 0 Then
            
                ImpoIva = Round2(baseimpo * PorcIva / 100, 2)
            
                Select Case DBLet(vSocio.TipoIRPF, "N")
                    Case 0
                        ImpoReten = Round2((baseimpo + ImpoIva) * vParamAplic.PorcreteFacSoc / 100, 2)
                        BaseReten = (baseimpo + ImpoIva)
                        PorcReten = vParamAplic.PorcreteFacSoc
                    Case 1
                        ImpoReten = Round2(baseimpo * vParamAplic.PorcreteFacSoc / 100, 2)
                        BaseReten = baseimpo
                        PorcReten = vParamAplic.PorcreteFacSoc
                    Case 2
                        ImpoReten = 0
                        BaseReten = 0
                        PorcReten = 0
                End Select
            
                TotalFac = baseimpo + ImpoIva - ImpoReten - ImpoAFO
                
                
                'insertar cabecera de factura
                B = InsertCabecera(tipoMov, CStr(numfactu), FecFac)
                
                '[Monica]24/12/2013: si es tercero he de marcarla como contabilizada
                If EsTerceros Then
                    If B Then B = MarcarFactura(tipoMov, CStr(numfactu), FecFac, True)
                End If
                
                
                If B Then B = InsertResumen(tipoMov, CStr(numfactu))
                
                '[Monica]15/04/2013: Introducimos las facturas varias a descontar
                If DescontarFVarias Then
                    If B Then B = InsertFacturasVarias(tipoMov, CStr(numfactu), FecFac, 0, 0)
                End If
                
                If B Then B = vTipoMov.IncrementarContador(tipoMov)
            Else
                B = True
                
            End If
                
            IncrementarProgresNew Pb1, 1
            
            
            If B Then
                AntSocio = ActSocio
                
                Set vSocio = Nothing
                Set vSocio = New cSocio
                If vSocio.LeerDatos(ActSocio) Then
                    If vSocio.LeerDatosSeccion(AntSocio, vParamAplic.Seccionhorto) Then
                        vPorcIva = ""
                        vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vSocio.CodIva, "N")
                        PorcIva = CCur(ImporteSinFormato(vPorcIva))
                    End If
                    If EsTerceros Then
                        tipoMov = "FAT"
                    Else
                        tipoMov = vSocio.CodTipomAnt
                    End If
                End If
                baseimpo = 0
                BaseReten = 0
                Neto = 0
                ImpoIva = 0
                ImpoReten = 0
                TotalFac = 0
                BaseAFO = 0
                ImpoAFO = 0
                
                numfactu = vTipoMov.ConseguirContador(tipoMov)
                Do
                    numfactu = vTipoMov.ConseguirContador(tipoMov)
                    devuelve = DevuelveDesdeBDNew(cAgro, "rfactsoc", "numfactu", "codtipom", tipoMov, "T", , "numfactu", CStr(numfactu), "N", "fecfactu", FecFac, "F")
                    If devuelve <> "" Then
                        'Ya existe el contador incrementarlo
                        Existe = True
                        vTipoMov.IncrementarContador (tipoMov)
                        numfactu = vTipoMov.ConseguirContador(tipoMov)
                    Else
                        Existe = False
                    End If
                Loop Until Not Existe
           End If
        End If
        
        
        Dim Sql9 As String
        Dim Rs9 As ADODB.Recordset
        Dim Precio As Currency
        
        Sql9 = "select precio1 from tmpinformes2 where importe1 = " & DBSet(Rs!numalbar, "N") & " and importe2 = " & DBSet(Rs!Codvarie, "N")
        Sql9 = Sql9 & " and importe3  = " & DBSet(Rs!codcalid, "N") & " and codusu = " & vUsu.Codigo
        
        Set Rs9 = New ADODB.Recordset
        Rs9.Open Sql9, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
        
        If Not Rs9.EOF Then
            '[Monica]24/02/2014: añadida variable
            HayPrecio = True
            
            Precio = DBLet(Rs9.Fields(0).Value, "N")
            PorcBoni = 0
            PorcComi = 0
            ' si el precio es positivo miramos si hay porcentaje de bonificacion para esa fecha
            If Precio > 0 Then
                PorcBoni = DevuelveValor("select porcbonif from rbonifentradas where codvarie = " & DBSet(Rs!Codvarie, "N") & " and fechaent = " & DBSet(Rs!Fecalbar, "F"))
                
                '[Monica]03/02/2012: Si el precio es positivo vemos si tiene comision el campo y se lo descontamos si es positivo
                PorcComi = DevuelveValor("select dtoprecio from rcampos where codcampo = " & DBSet(Rs!codCampo, "N"))
                If CCur(PorcComi) <> 0 Then
                    Precio = Precio - Round2(Precio * PorcComi / 100, 4)
                End If
            End If
            
            '[Monica]25/01/2016: para el caso de Picassent si la calidad no tiene bonificacion PorcBoni = 0
            If Not EsCalidadConBonificacion(CStr(Rs!Codvarie), CStr(Rs!codcalid)) Then PorcBoni = 0
            
        
            vPrecio = DBLet(Precio, "N")
            vImporte = vImporte + Round2(DBLet(Rs!KilosNet, "N") * Precio * (1 + (PorcBoni / 100)), 2)
            vBonifica = vBonifica + Round2(DBLet(Rs!KilosNet, "N") * Precio * (1 + (PorcBoni / 100)), 2) - Round2(DBLet(Rs!KilosNet, "N") * Precio, 2)
            
            KilosCal = KilosCal + DBLet(Rs!KilosNet, "N")
            
        Else
            HayPrecio = False
        End If
        
        Set Rs9 = Nothing
        
        '[Monica]20/03/2014: miramos si hay precio para la calidad
        Sql9 = "select count(*) from tmpinformes2 where importe5 = " & DBSet(Rs!codCampo, "N") & " and importe2 = " & DBSet(Rs!Codvarie, "N") & " and importeb1 = " & DBSet(Rs!Codsocio, "N")
        Sql9 = Sql9 & " and importe3  = " & DBSet(Rs!codcalid, "N") & " and codusu = " & vUsu.Codigo
        HayPrecio = (TotalRegistros(Sql9) <> 0)
        
        
        'hasta aqui
        HayReg = True
        
        Rs.MoveNext
    Wend
    
    ' ultimo registro si ha entrado
    If B And HayReg Then
        ' insertar linea de calidad
        '[Monica]24/02/2014: añadida condicion
        If HayPrecio Then
            Kilos = Kilos + KilosCal
            Importe = Importe + vImporte
            Bonifica = Bonifica + vBonifica
            
            baseimpo = baseimpo + vImporte
            
            If B Then B = InsertLineaCalidad(tipoMov, CStr(numfactu), FecFac, ActVarie, actCampo, CStr(ActCalid), CStr(DBLet(KilosCal, "N")), CStr(vImporte), CStr(vBonifica))
        End If
        
        '[Monica]24/02/2014: añadida condicion
        If Kilos <> 0 Then
            ' insertar linea de variedad
            If B Then B = InsertLinea(tipoMov, CStr(numfactu), FecFac, CStr(ActVarie), CStr(actCampo), CStr(Kilos), CStr(Importe), "0", CStr(Bonifica))
        End If
        
        '[Monica]24/02/2014: añadida condicion
        If baseimpo <> 0 Then
            ImpoIva = Round2(baseimpo * PorcIva / 100, 2)
    
            Select Case DBLet(vSocio.TipoIRPF, "N")
                Case 0
                    ImpoReten = Round2((baseimpo + ImpoIva) * vParamAplic.PorcreteFacSoc / 100, 2)
                    BaseReten = (baseimpo + ImpoIva)
                    PorcReten = vParamAplic.PorcreteFacSoc
                Case 1
                    ImpoReten = Round2(baseimpo * vParamAplic.PorcreteFacSoc / 100, 2)
                    BaseReten = baseimpo
                    PorcReten = vParamAplic.PorcreteFacSoc
                Case 2
                    ImpoReten = 0
                    BaseReten = 0
                    PorcReten = 0
            End Select
            
    '        BaseAFO = baseimpo
    '        PorcAFO = vParamAplic.PorcenAFO
    '        ImpoAFO = Round2(BaseAFO * PorcAFO / 100, 2)
    
            TotalFac = baseimpo + ImpoIva - ImpoReten - ImpoAFO
            
            
            'insertar cabecera de factura
            B = InsertCabecera(tipoMov, CStr(numfactu), FecFac)
        
            '[Monica]24/12/2013: si es tercero he de marcarla como contabilizada
            If EsTerceros Then
                '[Monica]27/07/2017: antes no se ponia que se tuviera que recibir factura
                If B Then B = MarcarFactura(tipoMov, CStr(numfactu), FecFac, True)
            End If
            
            
            '[Monica]15/04/2013: Introducimos las facturas varias a descontar
            If DescontarFVarias Then
                If B Then B = InsertFacturasVarias(tipoMov, CStr(numfactu), FecFac, 0, 0)
            End If
            
            If B Then B = InsertResumen(tipoMov, CStr(numfactu))
            
            If B Then B = vTipoMov.IncrementarContador(tipoMov)
        
        Else
            B = True
        End If
        
        IncrementarProgresNew Pb1, 1
        
        vParamAplic.UltFactAnt = numfactu
        
        'pasamos las temporales a las tablas
        If B Then B = PasarTemporales()
        
        If B Then B = (vParamAplic.Modificar = 1)
    End If
    
'    BorrarTMPs
    
    vSeccion.CerrarConta
    Set vSeccion = Nothing
    Set vSocio = Nothing
    
eFacturacion:
    If Err.Number <> 0 Or Not B Then
        conn.RollbackTrans
        FacturacionAnticiposPicassentNew = False
    Else
        conn.CommitTrans
        FacturacionAnticiposPicassentNew = True
    End If
End Function


Public Function FacturacionLiquidacionesPicassentNew(cTabla As String, cWhere As String, FecFac As String, Pb1 As ProgressBar, TipoPrec As Byte, DescontarFVarias As Boolean, EsTerceros As Boolean) As Boolean
Dim SQL As String
Dim Rs As ADODB.Recordset
Dim RS1 As ADODB.Recordset
Dim Sql4 As String

Dim AntSocio As String
Dim AntVarie As String
Dim ActSocio As String
Dim ActVarie As String
Dim actCampo As String
Dim AntCampo As String
Dim ActCalid As String
Dim AntCalid As String

Dim HayReg As Boolean

Dim NumError As Long
Dim vImporte As Currency
Dim vPorcIva As String

Dim PrimFac As String
Dim UltFac As String

Dim tipoMov As String
Dim B As Boolean
Dim vSeccion As CSeccion
Dim Kilos As Currency
Dim KilosCal As Currency
Dim Importe As Currency

Dim devuelve As String
Dim Existe As Boolean

'Dim baseimpo As Currency
'Dim BaseReten As Currency
'Dim Neto As Currency
'Dim ImpoIva As Currency
'Dim ImpoReten As Currency
'Dim TotalFac As Currency
Dim Recolect As Byte
Dim vPrecio As Currency

Dim Sql2 As String
Dim Sql3 As String
Dim SqlAlbaranes As String

Dim GastosCoop As Currency
Dim GastosAlb As Currency
Dim vPorcGasto As String

Dim SqlAFO As String

Dim vBonifica As Currency
Dim Bonifica As Currency
Dim PorcBoni As Currency
Dim PorcComi As Currency

Dim Incremento As Currency

Dim HayPrecio As Boolean

Dim Concepto As String

    On Error GoTo eFacturacion

    FacturacionLiquidacionesPicassentNew = False
    
'08052009 antes dentro de transacciones
    BorrarTMPs
    B = CrearTMPs()
    If Not B Then
         Exit Function
    End If
'08052009
    
    conn.BeginTrans
    
    If EsTerceros Then
        tipoMov = "FLT" ' facturas de liquidacion de terceros
    Else
        tipoMov = "FAL"
    End If
    
    SQL = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute SQL

    SQL = "SELECT  rhisfruta.codsocio, rhisfruta.codvarie, rhisfruta.codcampo,"
    SQL = SQL & "rhisfruta.recolect, rhisfruta_clasif.codcalid, rhisfruta.fecalbar, rhisfruta.numalbar, "
    SQL = SQL & "sum(coalesce(rhisfruta_clasif.kilosnet,0)) as kilosnet "
    SQL = SQL & " FROM  " & cTabla


    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        SQL = SQL & " WHERE " & cWhere
    End If

    ' ordenado por socio, variedad, campo, calidad
    SQL = SQL & " group by rhisfruta.codsocio, rhisfruta.codvarie, rhisfruta.codcampo, rhisfruta_clasif.codcalid, rhisfruta.numalbar, rhisfruta.fecalbar,  rhisfruta.recolect "
    SQL = SQL & " having sum(coalesce(rhisfruta_clasif.kilosnet,0)) <> 0"
    SQL = SQL & " order by rhisfruta.codsocio, rhisfruta.codvarie, rhisfruta.codcampo, rhisfruta_clasif.codcalid, rhisfruta.numalbar, rhisfruta.fecalbar, rhisfruta.recolect "
    

'    Sql = "SELECT  tmpinformes2.importeb1 codsocio, tmpinformes2.importe2 codvarie,"
'    Sql = Sql & "tmpinformes2.importe5 codcampo, tmpinformes2.importe3 codcalid, "
'    Sql = Sql & "sum(tmpinformes2.importe4) as kilosnet, sum(round(tmpinformes2.importe4 * tmpinformes2.precio1,2)) as importe "
'    Sql = Sql & " FROM  tmpinformes2 "
'    Sql = Sql & " where codusu = " & DBSet(vUsu.Codigo, "N")
'
'    ' ordenado por socio, variedad, campo, calidad
'    Sql = Sql & " group by tmpinformes2.codsocio, tmpinformes2.codvarie, tmpinformes2.codcampo, tmpinformes2.codcalid "
'    Sql = Sql & " order by tmpinformes2.codsocio, tmpinformes2.codvarie, tmpinformes2.codcampo, tmpinformes2.codcalid "
'
    
    
    Set vSeccion = New CSeccion
    
    If vSeccion.LeerDatos(vParamAplic.Seccionhorto) Then
        If Not vSeccion.AbrirConta Then
            Exit Function
        End If
    End If
    
    HayReg = False
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    If Not Rs.EOF Then
        AntSocio = CStr(DBLet(Rs!Codsocio, "N"))
        AntVarie = CStr(DBLet(Rs!Codvarie, "N"))
        AntCampo = CStr(DBLet(Rs!codCampo, "N"))
        AntCalid = CStr(DBLet(Rs!codcalid, "N"))
        
        ActSocio = CStr(DBLet(Rs!Codsocio, "N"))
        ActVarie = CStr(DBLet(Rs!Codvarie, "N"))
        actCampo = CStr(DBLet(Rs!codCampo, "N"))
        ActCalid = CStr(DBLet(Rs!codcalid, "N"))
    
        Set vSocio = New cSocio
        If vSocio.LeerDatos(ActSocio) Then
            If vSocio.LeerDatosSeccion(ActSocio, vParamAplic.Seccionhorto) Then
                baseimpo = 0
                BaseReten = 0
                ImpoIva = 0
                ImpoReten = 0
                TotalFac = 0
                ImpoAFO = 0
                BaseAFO = 0
                
                Anticipos = 0
                
                Kilos = 0
                Importe = 0
                
                KilosCal = 0
                
                vPorcIva = ""
                '[Monica]29/04/2011: INTERNAS
                If vSocio.EsFactADVInt Then
                    vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vParamAplic.CodIvaExeADV, "N")
                Else
                    vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vSocio.CodIva, "N")
                End If
                PorcIva = CCur(ImporteSinFormato(vPorcIva))
                
                If EsTerceros Then
                    tipoMov = "FLT"
                Else
                    tipoMov = vSocio.CodTipomLiq
                End If
                
                Set vTipoMov = New CTiposMov
                
                numfactu = vTipoMov.ConseguirContador(tipoMov)
                Do
                    numfactu = vTipoMov.ConseguirContador(tipoMov)
                    devuelve = DevuelveDesdeBDNew(cAgro, "rfactsoc", "numfactu", "codtipom", tipoMov, "T", , "numfactu", CStr(numfactu), "N", "fecfactu", FecFac, "F")
                    If devuelve <> "" Then
                        'Ya existe el contador incrementarlo
                        Existe = True
                        vTipoMov.IncrementarContador (tipoMov)
                        numfactu = vTipoMov.ConseguirContador(tipoMov)
                    Else
                        Existe = False
                    End If
                Loop Until Not Existe
                
                vParamAplic.PrimFactLiq = numfactu
                
            End If
        End If
    End If
    
    While Not Rs.EOF And B
        ActCalid = DBLet(Rs!codcalid, "N")
        ActVarie = DBLet(Rs!Codvarie, "N")
        actCampo = DBSet(Rs!codCampo, "N")
        ActSocio = DBSet(Rs!Codsocio, "N")
        
        If (ActCalid <> AntCalid Or AntCampo <> actCampo Or AntVarie <> ActVarie Or AntSocio <> ActSocio) Then
            '[Monica]24/02/2014: añadida condicion
            If HayPrecio Then
        
                ' kilos e importe por variedad campo
                Kilos = Kilos + KilosCal
                Importe = Importe + vImporte
                Bonifica = Bonifica + vBonifica
                
                baseimpo = baseimpo + vImporte + vBonifica
                
                B = InsertLineaCalidad(tipoMov, CStr(numfactu), FecFac, AntVarie, AntCampo, CStr(AntCalid), CStr(DBLet(KilosCal, "N")), CStr(vImporte + vBonifica))
            
            End If
            
            KilosCal = 0
            vImporte = 0
            vBonifica = 0
            
            AntCalid = ActCalid
        End If
        
        If (ActVarie <> AntVarie Or actCampo <> AntCampo Or ActSocio <> AntSocio) Then
        
            '[Monica]24/02/2014: añadida condicion
            If Kilos <> 0 Then
        
                If B Then ' descontamos el porcentaje de gastos de cooperativa
                    GastosCoop = 0
                    
                    vPorcGasto = ""
                    vPorcGasto = DevuelveDesdeBDNew(cAgro, "rcoope", "porcgast", "codcoope", vSocio.Cooperativa, "N")
                    If vPorcGasto = "" Then vPorcGasto = "0"
                    
                    '[Monica]25/02/2011: Sólo hay gastos si no es complementaria ( Añadido el if )
                    If TipoPrec <> 3 Then
                        GastosCoop = Round2(Importe * CCur(ImporteSinFormato(vPorcGasto)) / 100, 2)
                        Importe = Importe - GastosCoop
                        baseimpo = baseimpo - GastosCoop
                    End If
                End If
                
                
                '[Monica]08/04/2010: grabamos los albaranes que intervienen en la linea de factura
                If B Then
                    B = InsertarAlbaranesFactura(tipoMov, CStr(numfactu), FecFac, AntSocio, AntVarie, AntCampo, cTabla, cWhere, 0)
                End If
                            
                ' insertar linea de variedad, campo
                If B Then
                    B = InsertLinea(tipoMov, CStr(numfactu), FecFac, AntVarie, AntCampo, CStr(Kilos), CStr(Importe + Bonifica), CStr(GastosAlb))
                End If
            
            
                '[Monica]10/01/2014: en el caso de que haya incremento hemos de insertarlo y aumentar la base
                If B Then
                    If ActVarie <> AntVarie Or ActSocio <> AntSocio Then
                        Sql4 = "select sum(ringresos.importe) from ringresos where codsocio = " & DBSet(AntSocio, "N")
                        Sql4 = Sql4 & " and codvarie = " & DBSet(AntVarie, "N")
                        
                        Incremento = DevuelveValor(Sql4)
        
                        If Incremento <> 0 Then
                            B = InsertLinea(tipoMov, CStr(numfactu), FecFac, AntVarie, 0, 0, CStr(Incremento), 0, 0)
                            baseimpo = baseimpo + Incremento
                        
                            '[Monica]10/04/2017: insertamos las lineas de ingresos que tiene que imprimir Coopic en la factura
                            If B Then
                                B = InsertLineasIngresos(tipoMov, CStr(numfactu), FecFac, AntVarie, AntSocio)
                            End If
                        
                        End If
                        
                        'borramos la linea de ingresos para el socio variedad
                        Sql4 = "delete from ringresos where codsocio = " & DBSet(AntSocio, "N")
                        Sql4 = Sql4 & " and codvarie = " & DBSet(AntVarie, "N")
                        
                        conn.Execute Sql4
                    End If
                End If
                
                If B Then
                    ' tenemos que descontar los anticipos que tengamos para ello
                    Sql2 = "select rfactsoc_variedad.imporvar, rfactsoc_variedad.numfactu, rfactsoc_variedad.fecfactu "
                    Sql2 = Sql2 & " from rfactsoc_variedad INNER JOIN rfactsoc ON rfactsoc_variedad.codtipom = rfactsoc.codtipom and "
                    Sql2 = Sql2 & " rfactsoc_variedad.numfactu = rfactsoc.numfactu and rfactsoc_variedad.fecfactu = rfactsoc.fecfactu "
                    '[Monica]21/01/2014: no contemplabamos los anticipos de terceros
                    If EsTerceros Then
                        Sql2 = Sql2 & " where rfactsoc_variedad.codtipom = 'FAT'"
                    Else
                        Sql2 = Sql2 & " where rfactsoc_variedad.codtipom = " & DBSet(vSocio.CodTipomAnt, "T") ' antes era 'FAA' "
                    End If
                    Sql2 = Sql2 & " and rfactsoc.codsocio = " & DBSet(AntSocio, "N")
                    Sql2 = Sql2 & " and codvarie = " & DBSet(AntVarie, "N")
                    Sql2 = Sql2 & " and codcampo = " & DBSet(AntCampo, "N")
                    Sql2 = Sql2 & " and rfactsoc_variedad.descontado = 0"
                    
                    Set RS1 = New ADODB.Recordset
                    RS1.Open Sql2, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
                    
                    While Not RS1.EOF
                        baseimpo = baseimpo - DBLet(RS1.Fields(0).Value, "N")
                        Anticipos = Anticipos + DBLet(RS1.Fields(0).Value, "N")
                        
                        ' indicamos que los anticipos ya han sido descontados
                        Sql3 = "update rfactsoc_variedad set descontado = 1 where "
                        '[Monica]21/01/2014: no contemplabamos los anticipos de terceros
                        If EsTerceros Then
                            Sql3 = Sql3 & " codtipom = 'FAT'"  ' antes era 'FAC'
                        Else
                            Sql3 = Sql3 & " codtipom = " & DBSet(vSocio.CodTipomAnt, "T") ' antes era 'FAC'
                        End If
                        Sql3 = Sql3 & " and numfactu = " & DBSet(RS1!numfactu, "N")
                        Sql3 = Sql3 & " and fecfactu = " & DBSet(RS1!fecfactu, "F") & " and codvarie = " & DBSet(AntVarie, "N")
                        Sql3 = Sql3 & " and codcampo = " & DBSet(AntCampo, "N")
                        
                        conn.Execute Sql3
                        
                        ' insertamos en la tabla de anticipos de liquidacion venta campo
                        Sql3 = "insert into tmpFact_anticipos (codtipom, numfactu, fecfactu, codtipomanti, numfactuanti, fecfactuanti, codvarieanti, codcampoanti, baseimpo) values ("
                        
                        '[Monica]21/01/2014: no contemplabamos los anticipos de terceros
                        If EsTerceros Then
                            Sql3 = Sql3 & "'FLT',"
                        Else
                            Sql3 = Sql3 & DBSet(vSocio.CodTipomLiq, "T") & "," ' antes era 'FAL'
                        End If
                        
                        Sql3 = Sql3 & DBSet(numfactu, "N") & "," & DBSet(FecFac, "F") & ","
                        '[Monica]21/01/2014: consideramos terceros
                        If EsTerceros Then
                            Sql3 = Sql3 & "'FAT'," ' antes era 'FAA'
                        Else
                            Sql3 = Sql3 & DBSet(vSocio.CodTipomAnt, "T") & "," ' antes era 'FAA'
                        End If
                        Sql3 = Sql3 & DBSet(RS1!numfactu, "N") & "," & DBSet(RS1!fecfactu, "F") & ","
                        Sql3 = Sql3 & DBSet(AntVarie, "N") & "," & DBSet(AntCampo, "N") & "," & DBSet(RS1!imporvar, "N") & ")"
                        
                        conn.Execute Sql3
                        
                        RS1.MoveNext
                    Wend
                    
                    Set RS1 = Nothing
                    ' fin descontar anticipos
                
                End If
            Else

                B = True

            End If
                
            
            If B Then
                AntVarie = ActVarie
                AntCampo = actCampo
                Kilos = 0
                Importe = 0
                Bonifica = 0
                Incremento = 0
            End If
        End If
        
        If ActSocio <> AntSocio Then
            
            '[Monica]24/02/2014: añadida condicion
            If baseimpo <> 0 Then
            
                ImpoIva = Round2(baseimpo * PorcIva / 100, 2)
            
                Select Case DBLet(vSocio.TipoIRPF, "N")
                    Case 0
                        ImpoReten = Round2((baseimpo + ImpoIva) * vParamAplic.PorcreteFacSoc / 100, 2)
                        BaseReten = (baseimpo + ImpoIva)
                        PorcReten = vParamAplic.PorcreteFacSoc
                    Case 1
                        ImpoReten = Round2(baseimpo * vParamAplic.PorcreteFacSoc / 100, 2)
                        BaseReten = baseimpo
                        PorcReten = vParamAplic.PorcreteFacSoc
                    Case 2
                        ImpoReten = 0
                        BaseReten = 0
                        PorcReten = 0
                End Select
                
                If TipoPrec <> 3 Then
                    ' El importe AFO lo tiene que tener guardado en la tabla intermedia
                    ImpoAFO = DevuelveValor("select sum(importe) from raporreparto where codsocio = " & DBSet(vSocio.Codigo, "N") & " and tipoentr = 0")
                Else
                    ImpoAFO = 0
                End If
                BaseAFO = 0
                PorcAFO = 0
    
                TotalFac = baseimpo + ImpoIva - ImpoReten - ImpoAFO
            
            
                '[Monica]07/02/2012: indicamos si es una factura de liquidacion complementaria
                'insertar cabecera de factura
                B = InsertCabecera(tipoMov, CStr(numfactu), FecFac, , , (TipoPrec = 3))
                
                '[Monica]24/12/2013: si es tercero he de marcarla como contabilizada
                If EsTerceros Then
                    If B Then B = MarcarFactura(tipoMov, CStr(numfactu), FecFac, True)
                End If
                
                If B Then B = InsertResumen(tipoMov, CStr(numfactu))
                
                vParamAplic.UltFactLiq = numfactu
    
    
    'Mirar si quito lo de reclacular calidades
    '            If b Then b = RecalcularCalidades(TipoMov, CStr(numfactu), FecFac)
                
    'Recalculo de todos los importes de tmpfact_calidades y tmpfact_variedades para que cuadre con la base de cabecera
    '            If b Then b = CuadrarBasesFactura(TipoMov, CStr(numfactu), FecFac, BaseImpo)
    
                '[Monica]15/04/2013: Descontamos facturas varias
                If DescontarFVarias Then
                    If B Then B = InsertFacturasVarias(tipoMov, CStr(numfactu), FecFac, 1, 0)
                End If
                
                
                If B Then B = vTipoMov.IncrementarContador(tipoMov)
            Else
            
                B = True
                
            End If
            
            IncrementarProgresNew Pb1, 1
            
            
            If B Then
                AntSocio = ActSocio
                
                Set vSocio = Nothing
                Set vSocio = New cSocio
                If vSocio.LeerDatos(ActSocio) Then
                    If vSocio.LeerDatosSeccion(AntSocio, vParamAplic.Seccionhorto) Then
                        vPorcIva = ""
                        '[Monica]29/04/2011: INTERNAS
                        If vSocio.EsFactADVInt Then
                            vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vParamAplic.CodIvaExeADV, "N")
                        Else
                            vPorcIva = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vSocio.CodIva, "N")
                        End If
                        
                        PorcIva = CCur(ImporteSinFormato(vPorcIva))
                    End If
                    If EsTerceros Then
                        tipoMov = "FLT"
                    Else
                        tipoMov = vSocio.CodTipomLiq
                    End If
                                        
                End If
                baseimpo = 0
                BaseReten = 0
                ImpoIva = 0
                ImpoReten = 0
                TotalFac = 0
                BaseAFO = 0
                ImpoAFO = 0
                
                Anticipos = 0
                
                numfactu = vTipoMov.ConseguirContador(tipoMov)
                Do
                    numfactu = vTipoMov.ConseguirContador(tipoMov)
                    devuelve = DevuelveDesdeBDNew(cAgro, "rfactsoc", "numfactu", "codtipom", tipoMov, "T", , "numfactu", CStr(numfactu), "N", "fecfactu", FecFac, "F")
                    If devuelve <> "" Then
                        'Ya existe el contador incrementarlo
                        Existe = True
                        vTipoMov.IncrementarContador (tipoMov)
                        numfactu = vTipoMov.ConseguirContador(tipoMov)
                    Else
                        Existe = False
                    End If
                Loop Until Not Existe
           End If
        End If
        
        Recolect = DBLet(Rs!Recolect, "N")
        
        
        Dim Sql9 As String
        Dim Rs9 As ADODB.Recordset
        Dim Precio As Currency
        
        Sql9 = "select precio1 from tmpinformes2 where fecha1 = " & DBSet(Rs!Fecalbar, "F") & " and importe2 = " & DBSet(Rs!Codvarie, "N")
        Sql9 = Sql9 & " and importe3  = " & DBSet(Rs!codcalid, "N") & " and codusu = " & vUsu.Codigo
        Sql9 = Sql9 & " and importe1 = " & DBSet(Rs!numalbar, "N") & " and importeb1 = " & DBSet(Rs!Codsocio, "N")
        
        Set Rs9 = New ADODB.Recordset
        Rs9.Open Sql9, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
        
        
        If Not Rs9.EOF Then
            '[Monica]24/02/2014: añadida variable
            HayPrecio = True
        
            Precio = DBLet(Rs9.Fields(0).Value, "N")
            PorcBoni = 0
            PorcComi = 0
            ' si el precio es positivo miramos si hay porcentaje de bonificacion para esa fecha
            If Precio > 0 Then
                PorcBoni = DevuelveValor("select porcbonif from rbonifentradas where codvarie = " & DBSet(Rs!Codvarie, "N") & " and fechaent = " & DBSet(Rs!Fecalbar, "F"))
                
                '[Monica]03/02/2012: Si el precio es positivo vemos si tiene comision el campo y se lo descontamos si es positivo
                PorcComi = DevuelveValor("select dtoprecio from rcampos where codcampo = " & DBSet(Rs!codCampo, "N"))
                If CCur(PorcComi) <> 0 Then
                    Precio = Precio - Round2(Precio * PorcComi / 100, 4)
                End If
            End If
            
            
            '[Monica]25/01/2016: para el caso de Picassent si la calidad no tiene bonificacion PorcBoni = 0
            If Not EsCalidadConBonificacion(CStr(Rs!Codvarie), CStr(Rs!codcalid)) Then PorcBoni = 0
            
            
            vPrecio = DBLet(Precio, "N")
            vImporte = vImporte + Round2(DBLet(Rs!KilosNet, "N") * Precio, 2)
            
            vBonifica = vBonifica + Round2(DBLet(Rs!KilosNet, "N") * Precio * (PorcBoni / 100), 2)
            
            KilosCal = KilosCal + DBLet(Rs!KilosNet, "N")
        
        Else
            '[Monica]24/02/2014: añadida condicion
            HayPrecio = False
            
        End If
        
        Set Rs9 = Nothing
        
        '[Monica]20/03/2014: miramos si hay precio para la calidad
        Sql9 = "select count(*) from tmpinformes2 where importe5 = " & DBSet(Rs!codCampo, "N") & " and importe2 = " & DBSet(Rs!Codvarie, "N") & " and importeb1 = " & DBSet(Rs!Codsocio, "N")
        Sql9 = Sql9 & " and importe3  = " & DBSet(Rs!codcalid, "N") & " and codusu = " & vUsu.Codigo
        HayPrecio = (TotalRegistros(Sql9) <> 0)
        
        
        
        'hasta aqui
        HayReg = True
        
        Rs.MoveNext
    Wend
    
    ' ultimo registro si ha entrado
    If B And HayReg Then
        ' insertar linea de calidad
        
       '[Monica]24/02/2014: añadida condicion
        If HayPrecio Then
            Kilos = Kilos + KilosCal
            Importe = Importe + vImporte
            Bonifica = Bonifica + vBonifica
            
            baseimpo = baseimpo + vImporte + vBonifica
            
            
            If B Then B = InsertLineaCalidad(tipoMov, CStr(numfactu), FecFac, ActVarie, actCampo, CStr(ActCalid), CStr(DBLet(KilosCal, "N")), CStr(vImporte + vBonifica))
        Else
            
            B = True
        
        End If
        
        
        '[Monica]24/02/2014: añadida condicion
        If Kilos <> 0 Then
        
            If B Then ' descontamos el porcentaje de gastos de cooperativa
                GastosCoop = 0
                
                vPorcGasto = ""
                vPorcGasto = DevuelveDesdeBDNew(cAgro, "rcoope", "porcgast", "codcoope", vSocio.Cooperativa, "N")
                If vPorcGasto = "" Then vPorcGasto = "0"
                
                '[Monica]25/02/2011: Sólo hay gastos si no es complementaria ( Añadido el if )
                If TipoPrec <> 3 Then
                    GastosCoop = Round2(Importe * CCur(ImporteSinFormato(vPorcGasto)) / 100, 2)
                    Importe = Importe - GastosCoop
                    baseimpo = baseimpo - GastosCoop
                End If
            End If
            
            
            '[Monica]08/04/2010: grabamos los albaranes que intervienen en la linea de factura
            If B Then
                B = InsertarAlbaranesFactura(tipoMov, CStr(numfactu), FecFac, AntSocio, AntVarie, AntCampo, cTabla, cWhere, 0)
            End If
                        
                        
            ' insertar linea de variedad
            If B Then B = InsertLinea(tipoMov, CStr(numfactu), FecFac, CStr(ActVarie), CStr(actCampo), CStr(Kilos), CStr(Importe + Bonifica), CStr(GastosAlb))
            
            '[Monica]10/01/2014: en el caso de que haya incremento hemos de insertarlo y aumentar la base
            If B Then
                Sql4 = "select sum(ringresos.importe) from ringresos where codsocio = " & DBSet(ActSocio, "N")
                Sql4 = Sql4 & " and codvarie = " & DBSet(ActVarie, "N")
                
                Incremento = DevuelveValor(Sql4)
    
                If Incremento <> 0 Then
                    B = InsertLinea(tipoMov, CStr(numfactu), FecFac, AntVarie, 0, 0, CStr(Incremento), 0, 0)
                    baseimpo = baseimpo + Incremento
                
                    '[Monica]10/04/2017: insertamos las lineas de ingresos que tiene que imprimir Coopic en la factura
                    If B Then
                        B = InsertLineasIngresos(tipoMov, CStr(numfactu), FecFac, ActVarie, ActSocio)
                    End If
                End If
                
                'borramos la linea de ingresos para el socio variedad
                Sql4 = "delete from ringresos where codsocio = " & DBSet(ActSocio, "N")
                Sql4 = Sql4 & " and codvarie = " & DBSet(ActVarie, "N")
                
                conn.Execute Sql4
            End If
            
            ' tenemos que descontar los anticipos que tengamos para ello
            If B Then
                Sql2 = "select rfactsoc_variedad.imporvar, rfactsoc_variedad.numfactu, rfactsoc_variedad.fecfactu "
                Sql2 = Sql2 & " from rfactsoc_variedad INNER JOIN rfactsoc ON rfactsoc_variedad.codtipom = rfactsoc.codtipom and "
                Sql2 = Sql2 & " rfactsoc_variedad.numfactu = rfactsoc.numfactu and rfactsoc_variedad.fecfactu = rfactsoc.fecfactu "
                '[Monica]21/01/2014: sobre terceros
                If EsTerceros Then
                    Sql2 = Sql2 & " where rfactsoc_variedad.codtipom = 'FAT'" '& DBSet(vSocio.CodTipomAnt, "T") ' antes era 'FAA' "
                Else
                    Sql2 = Sql2 & " where rfactsoc_variedad.codtipom = " & DBSet(vSocio.CodTipomAnt, "T") ' antes era 'FAA' "
                End If
                Sql2 = Sql2 & " and rfactsoc.codsocio = " & DBSet(ActSocio, "N")
                Sql2 = Sql2 & " and codvarie = " & DBSet(ActVarie, "N")
                Sql2 = Sql2 & " and codcampo = " & DBSet(actCampo, "N")
                Sql2 = Sql2 & " and rfactsoc_variedad.descontado = 0"
                
                Set RS1 = New ADODB.Recordset
                RS1.Open Sql2, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
                
                While Not RS1.EOF
                    baseimpo = baseimpo - DBLet(RS1.Fields(0).Value, "N")
                    Anticipos = Anticipos + DBLet(RS1.Fields(0).Value, "N")
                    
                    ' indicamos que los anticipos ya han sido descontados
                    Sql3 = "update rfactsoc_variedad set descontado = 1 where codtipom = "
                    '[Monica]21/01/2014: sobre terceros
                    If EsTerceros Then
                        Sql3 = Sql3 & "'FAT'" ' antes era 'FAC'
                    Else
                        Sql3 = Sql3 & DBSet(vSocio.CodTipomAnt, "T") ' antes era 'FAC'
                    End If
                    Sql3 = Sql3 & " and numfactu = " & DBSet(RS1!numfactu, "N")
                    Sql3 = Sql3 & " and fecfactu = " & DBSet(RS1!fecfactu, "F") & " and codvarie = " & DBSet(ActVarie, "N")
                    Sql3 = Sql3 & " and codcampo = " & DBSet(actCampo, "N")
                    
                    conn.Execute Sql3
                    
                    ' insertamos en la tabla de anticipos de liquidacion venta campo
                    Sql3 = "insert into tmpFact_anticipos (codtipom, numfactu, fecfactu, codtipomanti, numfactuanti, fecfactuanti, codvarieanti, codcampoanti, baseimpo) values ("
                    '[Monica]21/01/2014: sobre terceros
                    If EsTerceros Then
                        Sql3 = Sql3 & "'FLT'" & "," ' antes era 'FAL'
                    Else
                        Sql3 = Sql3 & DBSet(vSocio.CodTipomLiq, "T") & "," ' antes era 'FAL'
                    End If
                    Sql3 = Sql3 & DBSet(numfactu, "N") & "," & DBSet(FecFac, "F") & ","
                    '[Monica]21/01/2014: consideramos terceros
                    If EsTerceros Then
                        Sql3 = Sql3 & "'FAT'," ' antes era 'FAA'
                    Else
                        Sql3 = Sql3 & DBSet(vSocio.CodTipomAnt, "T") & "," ' antes era 'FAA'
                    End If
                    Sql3 = Sql3 & DBSet(RS1!numfactu, "N") & "," & DBSet(RS1!fecfactu, "F") & ","
                    Sql3 = Sql3 & DBSet(ActVarie, "N") & "," & DBSet(actCampo, "N") & "," & DBSet(RS1!imporvar, "N") & ")"
                    
                    conn.Execute Sql3
                    
                    RS1.MoveNext
                Wend
                
                Set RS1 = Nothing
                ' fin descontar anticipos
            
            End If
            
        Else
            
            B = True
            
        End If
        
        
        '[Monica]24/02/2014: añadida condicion
        If baseimpo <> 0 Then
        
            ImpoIva = Round2(baseimpo * PorcIva / 100, 2)
    
            Select Case DBLet(vSocio.TipoIRPF, "N")
                Case 0
                    ImpoReten = Round2((baseimpo + ImpoIva) * vParamAplic.PorcreteFacSoc / 100, 2)
                    BaseReten = (baseimpo + ImpoIva)
                    PorcReten = vParamAplic.PorcreteFacSoc
                Case 1
                    ImpoReten = Round2(baseimpo * vParamAplic.PorcreteFacSoc / 100, 2)
                    BaseReten = baseimpo
                    PorcReten = vParamAplic.PorcreteFacSoc
                Case 2
                    ImpoReten = 0
                    BaseReten = 0
                    PorcReten = 0
            End Select
            
            If TipoPrec <> 3 Then ' si no es complementaria se calcula el impafo
                ImpoAFO = DevuelveValor("select sum(importe) from raporreparto where codsocio = " & DBSet(vSocio.Codigo, "N") & " and tipoentr = 0")
            Else
                ImpoAFO = 0
            End If
            BaseAFO = 0
            PorcAFO = 0
    
            TotalFac = baseimpo + ImpoIva - ImpoReten - ImpoAFO
            
        
            vParamAplic.UltFactLiq = numfactu
        
            '[Monica]07/02/2012: indicamos si es una factura de liquidacion complementaria
            'insertar cabecera de factura
            B = InsertCabecera(tipoMov, CStr(numfactu), FecFac, , , (TipoPrec = 3))
            
            '[Monica]24/12/2013: si es tercero he de marcarla como contabilizada
            If EsTerceros Then
                If B Then B = MarcarFactura(tipoMov, CStr(numfactu), FecFac, True)
            End If
            
            '[Monica]15/04/2013: Descontamos facturas varias
            If DescontarFVarias Then
                If B Then B = InsertFacturasVarias(tipoMov, CStr(numfactu), FecFac, 1, 0)
            End If
            
            If B Then B = InsertResumen(tipoMov, CStr(numfactu))

'Mirar si quito lo de reclacular calidades
'        If b Then b = RecalcularCalidades(TipoMov, CStr(numfactu), FecFac)
        
'Recalculo de todos los importes de rfactsoc_calidades y rfactsoc_variedades para que cuadre con la base de cabecera
'        If b Then b = CuadrarBasesFactura(TipoMov, CStr(numfactu), FecFac, BaseImpo)
        
            If B Then B = vTipoMov.IncrementarContador(tipoMov)
        
        End If
        
        IncrementarProgresNew Pb1, 1
        
        
        'pasamos las temporales a las tablas
        If B Then B = PasarTemporales()
        
        '[Monica]23/07/2012: si no es complementaria se calculan los gastos
        If TipoPrec <> 3 Then
            ' solo para Picassent: he de insertar las lineas de gastos al pie de factura que estan como gastos de albaranes
            If B Then B = DescontarGastosAPie()
        End If
        
        If B Then B = (vParamAplic.Modificar = 1)
    End If
    
'    BorrarTMPs
    
    vSeccion.CerrarConta
    Set vSeccion = Nothing
    Set vSocio = Nothing
    
eFacturacion:
    If Err.Number <> 0 Or Not B Then
        conn.RollbackTrans
        FacturacionLiquidacionesPicassentNew = False
    Else
        conn.CommitTrans
        FacturacionLiquidacionesPicassentNew = True
    End If
End Function



