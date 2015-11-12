VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFacturaBOD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'VENTAS
'FACTURA: se corresponde con la tabla ariges.scafac (cabecera de facturas)
'-------------------------------------------------------------------------

'ATRIBUTOS
'Variables locales que contienen valores de propiedad
Private mCodTipoM As String 'Tipo de Movimiento="FAV", "FAR" (Factura venta o reparacion)
Private mNumFactu As Long 'Nº Factura
Private mFecFactu As String 'FEcha Factura
Private mCodsocio As String 'Cod. socio

Private mNomsocio As String 'nombre del socio
Private mDomSocio As String 'domicilio
Private mCodPobla As String 'cpostal
Private mPobsocio As String 'poblacion
Private mProsocio As String 'provincia
Private mNIFsocio As String 'NIF proveedor
Private mTelSocio As String 'telefono

Private mCoddirec As String 'codigo direc./dpto
Private mNomdirec As String  'nombre direc./dpto
Private mCodAgent As Integer 'codigo de agente

Private mIban As String ' iban
Private mCodbanco As String 'cod. banco
Private mCodsucur As String 'cod. sucur
Private mDigContr As String 'digito control
Private mCuentaba As String 'cuenta bancaria

Private mCodForpa As String 'Cod. Forma de pago
Private mTipForpa As Byte 'tipo de forma de pago: efectivo,...

Private mDtoPpago As Currency 'descuento pronto pago
Private mDtoGnral As Currency 'descuento general
Private mBrutofac As Currency 'importe bruto de la factura
Private mImpPPago As Currency 'Importe al aplicar al bruto el dto pronto pago
Private mImpGnral As Currency  'Importe al aplicar al bruto el dto general
Private mBaseImp As Currency  'Base imponible de la factura (bruto - dtopp - dtogn)

Private mBaseiva1 As Currency  'base imponible tipo IVA 1
Private mBaseiva2 As Currency  'base imponible tipo IVA 2
Private mBaseiva3 As Currency  'base imponible tipo IVA 3
Private mTipoiva1 As Byte   'Cod. tipo de iva 1
Private mTipoiva2 As Byte   'Cod. tipo de iva 2
Private mTipoiva3 As Byte   'Cod. tipo de iva 3
Private mPorciva1 As Currency    '% de iva 1
Private mPorciva2 As Currency    '% de iva 2
Private mPorciva3 As Currency    '% de iva 3

Private mImpoiva1 As Currency   'Importe de iva 1
Private mImpoiva2 As Currency   'Importe de iva 2
Private mImpoiva3 As Currency   'Importe de iva 3

Private mTotalfac As String 'total factura

Private mLetraSer As String 'Letra de serie para los cobros
Private mBanPr As String 'Banco propio que ponemos como cuenta prevista de cobro
Private mCtaPrev As String 'Cuenta prevista de cobro

Private mNumtermi As Integer 'terminal de venta del TPV
Private mNumventa As Long 'Nº de venta del TPV


'David.
'Para las facturas de mantenimeinto, hay un campo que es "concepto factura" varchar(60)
'Si tiene valor entonces ese valor es el que grabaremos en la linea y como obervacion de la factura
'meteremos lo de antes ("texto" de: mes de año.  Ejemplo: manteimiento de Junio de 2007
'Si no tiene valor seguira como estaba
Private mObservacion As String

' VBLES para los Mantenimientos -------------------------------------------
Dim TipCoMan As String 'tipo de contrato para Facturas de Manteniminetos
Dim OpeFactu As String 'Operador conectado o el que pasamos como parametro
Dim AlmFactu As String 'Cod Almacen del trabajador conectado, para insertar lines mantenim, desplaz, bonificaciones,....
Dim MesFactu As String
Dim numMante As String
Dim ArticFac As String 'cod articulo que se va a utilizar para facturar mantenimientos
'--------------------------------------------------------------------------

' VBLES para los Desplazamientos ----------------------------------
Dim TotalKm As Integer
Dim PrecioKm As Currency
Dim ImporteL As Currency
Dim Ampliacion As String  'ampliacion de la linea: "Dia,Dia,Dia..." de los albaranes con desplazamiento
Dim codArtic As String
Dim despIVA As Integer 'tipo de IVA del articulo para desplazamientos
'------------------------------------------------------------------

' VBLES para las Bonificaciones -----------------------------------
Dim TotBonif1 As Currency 'Total de bonificacion a factura con articulo de tipo de IVA1
Dim IVABonif1 As Byte 'tipo de IVA
Dim TotBonif2 As Currency
Dim IVABonif2 As Byte
Dim TotBonif3 As Currency
Dim IVABonif3 As Byte
Dim SQLBonif As String
'------------------------------------------------------------------

' VBLES para los Cheques Regalo -----------------------------------
Dim ImpCheque As Currency
'------------------------------------------------------------------




'vbles para las facturas que vienen de una venta de ticket ---------
Private NumTicket As String
Private NumAlbTicket As String '01/09/06 laura
'-------------------------------------------------------------------


'# Laura 14/11/2006 recuperar facturas Alzira
Private EstaRecuFac As Boolean


'------------------------------------------------
'Propiedades del modulo CFactura
'------------------------------------------------

'**** Tipo de Movimiento de la Factura

Public Property Let CodTipom(ByVal vData As String)
     mCodTipoM = vData
End Property

Public Property Get CodTipom() As String
     CodTipom = mCodTipoM
End Property


'**** Nº de la Factura

Public Property Let numfactu(ByVal vData As Long)
     mNumFactu = vData
End Property

Public Property Get numfactu() As Long
     numfactu = mNumFactu
End Property


'**** Fecha de la Factura

Public Property Let fecfactu(ByVal vData As String)
     mFecFactu = vData
End Property

Public Property Get fecfactu() As String
     fecfactu = mFecFactu
End Property


'**** Socio de la Factura

Public Property Let Socio(ByVal vData As String)
     mCodsocio = vData
End Property

Public Property Get Socio() As String
     Socio = mCodsocio
End Property


'**** Nombre del Socio de la Factura

Public Property Let NombreSocio(ByVal vData As String)
     mNomsocio = vData
End Property

Public Property Get NombreSocio() As String
     NombreSocio = mNomsocio
End Property


'**** Domicilio del Socio de la Factura

Public Property Let DomicilioSocio(ByVal vData As String)
     mDomSocio = vData
End Property

Public Property Get DomicilioSocio() As String
     DomicilioSocio = mDomSocio
End Property


'**** CPostal del Socio de la Factura

Public Property Let CPostal(ByVal vData As String)
     mCodPobla = vData
End Property

Public Property Get CPostal() As String
     CPostal = mCodPobla
End Property


'**** Poblacion del Socio de la Factura

Public Property Let Poblacion(ByVal vData As String)
     mPobsocio = vData
End Property

Public Property Get Poblacion() As String
     Poblacion = mPobsocio
End Property


'**** Provincia del Socio de la Factura

Public Property Let Provincia(ByVal vData As String)
     mProsocio = vData
End Property

Public Property Get Provincia() As String
     Provincia = mProsocio
End Property


'**** NIF del Socio de la Factura

Public Property Let NIF(ByVal vData As String)
     mNIFsocio = vData
End Property

Public Property Get NIF() As String
     NIF = mNIFsocio
End Property


'**** Telefono del Socio de la Factura

Public Property Let Telefono(ByVal vData As String)
     mTelSocio = vData
End Property

Public Property Get Telefono() As String
     Telefono = mTelSocio
End Property


'**** Direc./Dpto de Socio de la Factura

Public Property Let DirDpto(ByVal vData As String)
     mCoddirec = vData
End Property

Public Property Get DirDpto() As String
     DirDpto = mCoddirec
End Property


'**** Nombre Direc./Dpto de Socio de la Factura

Public Property Let NombreDirDpto(ByVal vData As String)
     mNomdirec = vData
End Property

Public Property Get NombreDirDpto() As String
     NombreDirDpto = mNomdirec
End Property


'**** Agente de Socio de la Factura

Public Property Let Agente(ByVal vData As Integer)
     mCodAgent = vData
End Property

Public Property Get Agente() As Integer
     Agente = mCodAgent
End Property


'**** Forma Pago de la Factura

Public Property Let ForPago(ByVal vData As String)
     mCodForpa = vData
End Property

Public Property Get ForPago() As String
     ForPago = mCodForpa
End Property


'**** Tipo de Forma Pago de la Factura
'(para tickets saber si insertar en tesoreria o no: si es efectivo no se inserta en scobro)

Public Property Let TipForPago(ByVal vData As Byte)
     mTipForpa = vData
End Property

Public Property Get TipForPago() As Byte
     TipForPago = mTipForpa
End Property


'**** Cuenta Prevista de cobro de la Factura

Public Property Let LetraSerie(ByVal vData As String)
     mLetraSer = vData
End Property

Public Property Get LetraSerie() As String
     LetraSerie = mLetraSer
End Property


'**** Banco propio de pago de la Factura

Public Property Let BancoPr(ByVal vData As String)
     mBanPr = vData
End Property

Public Property Get BancoPr() As String
     BancoPr = mBanPr
End Property


'**** Cuenta Prevista de cobro de la Factura

Public Property Let CuentaPrev(ByVal vData As String)
     mCtaPrev = vData
End Property

Public Property Get CuentaPrev() As String
     CuentaPrev = mCtaPrev
End Property


'**** IBAN de la Factura

Public Property Let Iban(ByVal vData As String)
     mIban = vData
End Property

Public Property Get Iban() As String
     Iban = mIban
End Property




'**** Banco de la Factura

Public Property Let Banco(ByVal vData As String)
     mCodbanco = vData
End Property

Public Property Get Banco() As String
     Banco = mCodbanco
End Property


'**** Sucursal de la Factura

Public Property Let Sucursal(ByVal vData As String)
     mCodsucur = vData
End Property

Public Property Get Sucursal() As String
     Sucursal = mCodsucur
End Property


'**** Digito Control de la Factura

Public Property Let Digcontrol(ByVal vData As String)
     mDigContr = vData
End Property

Public Property Get Digcontrol() As String
     Digcontrol = mDigContr
End Property


'**** Cuenta Bancaria de la Factura

Public Property Let CuentaBan(ByVal vData As String)
     mCuentaba = vData
End Property

Public Property Get CuentaBan() As String
     CuentaBan = mCuentaba
End Property


'**** Descuento pronto pago de la Factura

Public Property Let DtoPPago(ByVal vData As Currency)
     mDtoPpago = vData
End Property

Public Property Get DtoPPago() As Currency
     DtoPPago = mDtoPpago
End Property


'**** Descuento general de la Factura

Public Property Let DtoGnral(ByVal vData As Currency)
     mDtoGnral = vData
End Property

Public Property Get DtoGnral() As Currency
     DtoGnral = mDtoGnral
End Property


'**** Importe Bruto Factura

Public Property Let BrutoFac(ByVal vData As Currency)
     mBrutofac = vData
End Property

Public Property Get BrutoFac() As Currency
     BrutoFac = mBrutofac
End Property


'**** Importe pronto pago de la Factura

Public Property Let ImpPPago(ByVal vData As Currency)
     vData = vData
     mImpPPago = vData
End Property

Public Property Get ImpPPago() As Currency
     ImpPPago = mImpPPago
End Property


'**** Importe general de la Factura

Public Property Let ImpGnral(ByVal vData As Currency)
     mImpGnral = vData
End Property

Public Property Get ImpGnral() As Currency
     ImpGnral = mImpGnral
End Property


'**** Importe Base Imponible Factura

Public Property Let BaseImp(ByVal vData As Currency)
     mBaseImp = vData
End Property

Public Property Get BaseImp() As Currency
     BaseImp = mBaseImp
End Property


'**** Base IVA1 de la Factura

Public Property Let BaseIVA1(ByVal vData As Currency)
     mBaseiva1 = vData
End Property

Public Property Get BaseIVA1() As Currency
     BaseIVA1 = mBaseiva1
End Property


'**** Base IVA2 de la Factura

Public Property Let BaseIVA2(ByVal vData As Currency)
     mBaseiva2 = vData
End Property

Public Property Get BaseIVA2() As Currency
     BaseIVA2 = mBaseiva2
End Property


'**** Base IVA3 de la Factura

Public Property Let BaseIVA3(ByVal vData As Currency)
     mBaseiva3 = vData
End Property

Public Property Get BaseIVA3() As Currency
     BaseIVA3 = mBaseiva3
End Property


'**** Tipo iva 1 de la Factura

Public Property Let TipoIVA1(ByVal vData As Byte)
     mTipoiva1 = vData
End Property

Public Property Get TipoIVA1() As Byte
     TipoIVA1 = mTipoiva1
End Property


'**** Tipo iva 2 de la Factura

Public Property Let TipoIVA2(ByVal vData As Byte)
     mTipoiva2 = vData
End Property

Public Property Get TipoIVA2() As Byte
     TipoIVA2 = mTipoiva2
End Property


'**** Tipo iva 3 de la Factura

Public Property Let TipoIVA3(ByVal vData As Byte)
     mTipoiva3 = vData
End Property

Public Property Get TipoIVA3() As Byte
     TipoIVA3 = mTipoiva3
End Property


'**** % de iva 1 de la Factura

Public Property Let PorceIVA1(ByVal vData As Currency)
     mPorciva1 = vData
End Property

Public Property Get PorceIVA1() As Currency
     PorceIVA1 = mPorciva1
End Property


'**** % de iva 2 de la Factura

Public Property Let PorceIVA2(ByVal vData As Currency)
     mPorciva2 = vData
End Property

Public Property Get PorceIVA2() As Currency
     PorceIVA2 = mPorciva2
End Property


'**** % de iva 3 de la Factura

Public Property Let PorceIVA3(ByVal vData As Currency)
     mPorciva3 = vData
End Property

Public Property Get PorceIVA3() As Currency
     PorceIVA3 = mPorciva3
End Property


'**** Importe de IVA 1 de la Factura

Public Property Let ImpIVA1(ByVal vData As Currency)
     mImpoiva1 = vData
End Property

Public Property Get ImpIVA1() As Currency
     ImpIVA1 = mImpoiva1
End Property


'**** Importe de IVA 2 de la Factura

Public Property Let ImpIVA2(ByVal vData As Currency)
     mImpoiva2 = vData
End Property

Public Property Get ImpIVA2() As Currency
     ImpIVA2 = mImpoiva2
End Property


'**** Importe de IVA 3 de la Factura

Public Property Let ImpIVA3(ByVal vData As Currency)
     mImpoiva3 = vData
End Property

Public Property Get ImpIVA3() As Currency
     ImpIVA3 = mImpoiva3
End Property


'**** Total Factura

Public Property Let TotalFac(ByVal vData As String)
     mTotalfac = vData
End Property

Public Property Get TotalFac() As String
     TotalFac = mTotalfac
End Property


'**** Nº terminal de la venta del TPV

Public Property Let NumTerminal(ByVal vData As Integer)
     mNumtermi = vData
End Property

Public Property Get NumTerminal() As Integer
     NumTerminal = mNumtermi
End Property



'**** Nº venta del TPV

Public Property Let NumVenta(ByVal vData As Long)
     mNumventa = vData
End Property

Public Property Get NumVenta() As Long
     NumVenta = mNumventa
End Property


'**** Para los MANTENIMIENTOS.  sera 1 la linea. Si esta vacio lo cargaremos con la amplicaicon normal

Public Property Let Observacion(ByVal vData As String)
     mObservacion = vData
End Property

Public Property Get Observacion() As String
     Observacion = mObservacion
End Property




'------------------------------------------------
'Procedimientos del modulo CFactura
'------------------------------------------------

Public Function LeerDatos(vCodtipom As String, vNumfactu As String, vFecfactu As String) As Boolean
'Leer los datos de una factura almacenada en la tabla scafac
'Lee de la BD: Ariges, Tabla: scafac
'OUT: True si lee los datos correctamente
Dim Rs As ADODB.Recordset
Dim Mens As String
Dim SQL As String

    On Error GoTo ELeer
    
    
    LeerDatos = False
    SQL = "SELECT codtipom,numfactu,fecfactu,codsocio,codforpa,iban,codbanco,codsucur,digcontr,cuentaba,brutofac, totalfac FROM rbodfacturas "
    SQL = SQL & " WHERE codtipom='" & vCodtipom & "'"
    SQL = SQL & " AND numfactu=" & vNumfactu
    SQL = SQL & " AND fecfactu='" & Format(vFecfactu, FormatoFecha) & "'"
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    Mens = ""
    
    If Rs.EOF Then
        LeerDatos = False
    Else
        Mens = "Datos Factura"
        mCodTipoM = Rs!CodTipom
        mNumFactu = CStr(Rs!numfactu)
        mFecFactu = CStr(Rs!fecfactu)
        mCodsocio = CStr(Rs!Codsocio)
        ForPago = CStr(Rs!Codforpa)
        TipForPago = DevuelveDesdeBDNew(cAgro, "sforpa", "tipforpa", "codforpa", ForPago, "N")
        mIban = CStr(DBLet(Rs!Iban, "T"))
        mCodbanco = CStr(DBLet(Rs!CodBanco, "N"))
        mCodsucur = CStr(DBLet(Rs!CodSucur, "N"))
        mDigContr = DBLet(Rs!digcontr, "T")
        mCuentaba = DBLet(Rs!CuentaBa, "T")
        mBrutofac = CStr(Rs!BrutoFac)
        mTotalfac = CStr(Rs!TotalFac)

        mLetraSer = ObtenerLetraSerie(mCodTipoM)
        
        'si venimos de proceso de insertar tenemos la cuenta prevista de cobro que
        'se introdujo en el formulario de pasar albaranes a factura
        If CuentaPrev = "" Then
            'leer la cuenta prevista de cobros de la tabla de contabilidad: conta.scobro
            'cuando vamos a borrar una factura porque modificarmos
            'ya que este dato se pidio al facturar (SI EXISTE LA FACTURA EN TESORERIA)
            SQL = "SELECT COUNT(*) FROM scobro WHERE numserie='" & LetraSerie & "' and codfaccl=" & numfactu
            SQL = SQL & " AND fecfaccl=" & DBSet(fecfactu, "F")
            If RegistrosAListar(SQL, cConta) > 0 Then
                SQL = DevuelveDesdeBDNew(cConta, "scobro", "ctabanc1", "numserie", mLetraSer, "T", , "codfaccl", vNumfactu, "N", "fecfaccl", vFecfactu, "F")
                If SQL <> "" Then
                    mCtaPrev = SQL
                    LeerDatos = True
                Else
                    LeerDatos = False
                    Mens = "La cuenta prevista de cobro para la factura no puede ser nula."
                End If
            Else
                LeerDatos = True
            End If
        Else
            LeerDatos = True
        End If
    End If

    Rs.Close
    Set Rs = Nothing
    Exit Function

ELeer:
    Mens = "Se ha producido un error. " & Mens & vbCrLf
    Mens = Mens & "Número: " & Err.Number & vbCrLf
    Mens = Mens & "Descripción: " & Err.Description
    MsgBox Mens, vbExclamation
    Set Rs = Nothing
    LeerDatos = False
End Function


'======================================================================
'GRABAR EN TESORERIA
'======================================================================

Public Function InsertarEnTesoreria(MenError As String) As Boolean
'Guarda datos de Tesoreria en tablas: ariges.svenci y en conta.scobros
Dim b As Boolean
Dim Rs As ADODB.Recordset
Dim rsVenci As ADODB.Recordset
Dim SQL As String, Codmacta As String, textcsb33 As String
Dim CadValues As String, cadValuesAux As String 'para insertar en svenci
Dim CadValues2 As String, CadValuesAux2 As String 'para insertar en conta.scobro
Dim CadValues3 As String
Dim FecVenci As Date, FecVenci1 As Date
Dim ImpVenci As Single 'importe para insertar en la svenci
Dim ImpVenci2 As Single 'importe para insertar en conta.scobro
Dim i As Byte


    On Error GoTo EInsertarTesoreria

    b = False
    InsertarEnTesoreria = False
    CadValues = ""
    CadValues2 = ""

    'campo para insertar en conta.scobro de Tesoreria
    textcsb33 = "'FACTURA: " & LetraSerie & "-" & Format(numfactu, "0000000") & " de Fecha " & Format(fecfactu, "dd,mm,yyyy") & "'"

    'Obtener el Nº de Vencimientos de la forma de pago
    SQL = "SELECT numerove, primerve, restoven FROM sforpa WHERE codforpa=" & ForPago
    Set rsVenci = New ADODB.Recordset
    rsVenci.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not rsVenci.EOF Then
        If rsVenci!numerove > 0 Then
            'Obtener los dias de pago del cliente
            SQL = " SELECT  diapago1, diapago2, diapago3, mesnogir, diavtoat, codmacta "
            SQL = SQL & " FROM sclien "
            SQL = SQL & " WHERE codclien=" & Socio
            Set Rs = New ADODB.Recordset
            Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            
            '[Monica]03/07/2013: añado trim(codmacta)
            Codmacta = DBSet(Trim(Rs!Codmacta), "T")
'            textcsb33 = "'FACTURA: " & LetraSerie & "-" & Format(NumFactu, "0000000") & " de Fecha " & Format(FecFactu, "dd,mm,yyyy") & "'"
            
            If Not Rs.EOF Then
                cadValuesAux = "('" & CodTipom & "', " & numfactu & ", '" & Format(fecfactu, FormatoFecha) & "', "
                CadValuesAux2 = "('" & LetraSerie & "', " & numfactu & ", '" & Format(fecfactu, FormatoFecha) & "', "
                '-------- Primer Vencimiento
                i = 1
                'FECHA VTO
                FecVenci = CDate(fecfactu)
                '=== Laura 23/01/2007
                'FecVenci = FecVenci + CByte(DBLet(rsVenci!primerve, "N"))
                FecVenci = DateAdd("d", DBLet(rsVenci!primerve, "N"), FecVenci)
                '===
                'comprobar si tiene dias de pago y obtener la fecha del vencimiento correcta
                If TipForPago <> 0 Then
                    FecVenci = ComprobarFechaVenci(FecVenci, DBLet(Rs!DiaPago1, "N"), DBLet(Rs!DiaPago2, "N"), DBLet(Rs!DiaPago3, "N"))
                Else
                    FecVenci = ComprobarFechaVenci(FecVenci, 0, 0, 0)
                End If
                'Comprobar si cliente tiene mes a no girar
                FecVenci1 = FecVenci
                If CInt(DBLet(Rs!mesnogir, "N")) <> 0 Then
                    FecVenci1 = ComprobarMesNoGira(FecVenci1, DBLet(Rs!mesnogir, "N"), DBLet(Rs!DiaVtoAt, "N"), DBLet(Rs!DiaPago1, "N"), DBLet(Rs!DiaPago2, "N"), DBLet(Rs!DiaPago3, "N"))
                End If
                
                'Comprobar si cliente tiene dia de vencimiento atrasado
                CadValues = cadValuesAux & i & ", '" & Format(FecVenci1, FormatoFecha) & "', "
                CadValues2 = CadValuesAux2 & i & ", "
                CadValues2 = CadValues2 & Codmacta & ", " & ForPago & ", '" & Format(FecVenci1, FormatoFecha) & "', "
                
                'IMPORTE del Vencimiento
                If rsVenci!numerove = 1 Then
                    ImpVenci = TotalFac
                    ImpVenci2 = TotalFac - ImpCheque
                Else
                    ImpVenci = Round2(TotalFac / rsVenci!numerove, 2)
                    ImpVenci2 = Round2((TotalFac - ImpCheque) / rsVenci!numerove, 2)
                    'Comprobar que la suma de los vencimientos cuadra con el total de la factura
                    If ImpVenci * rsVenci!numerove <> TotalFac Then
                        ImpVenci = Round(ImpVenci + (TotalFac - ImpVenci * rsVenci!numerove), 2)
                    End If
                    'Comprobar que la suma de los vencimientos cuadra con el total de la factura
                    If (ImpVenci2 * rsVenci!numerove) + ImpCheque <> TotalFac Then
                        ImpVenci2 = Round(ImpVenci2 + (TotalFac - ImpCheque - (ImpVenci2 * rsVenci!numerove)), 2)
                    End If
                End If
                CadValues = CadValues & DBSet(ImpVenci, "N") & ")"
                CadValues2 = CadValues2 & DBSet(ImpVenci2, "N") & ", '" & CuentaPrev & "', " & DBSet(Banco, "N", "S") & ", " & DBSet(Sucursal, "N", "S") & ", " & DBSet(Digcontrol, "T", "S") & ", " & DBSet(CuentaBan, "T", "S") & ", " & textcsb33 & ", " & DBSet(Agente, "N")
                'departamento
                CadValues2 = CadValues2 & "," & DBSet(Me.DirDpto, "N", "S") '& ")"
                '[Monica]22/11/2013: Tema iban
                If vEmpresa.HayNorma19_34Nueva = 1 Then
                    CadValues2 = CadValues2 & ", " & DBSet(Iban, "T", "S") & ") "
                Else
                    CadValues2 = CadValues2 & ") "
                End If
                
                
                
                'Resto Vencimientos
                '--------------------------------------------------------------------
                For i = 2 To rsVenci!numerove
                   'FECHA Resto Vencimientos
                    '=== Laura 23/01/2007
                    'FecVenci = FecVenci + DBSet(rsVenci!restoven, "N")
                    FecVenci = DateAdd("d", DBLet(rsVenci!restoven, "N"), FecVenci)
                    '===
                    'comprobar si tiene dias de pago y obtener la fecha del vencimiento correcta
                    If TipForPago <> 0 Then
                        FecVenci = ComprobarFechaVenci(FecVenci, DBLet(Rs!DiaPago1, "N"), DBLet(Rs!DiaPago2, "N"), DBLet(Rs!DiaPago3, "N"))
                    Else
                        FecVenci = ComprobarFechaVenci(FecVenci, 0, 0, 0)
                    End If
                    'Comprobar si cliente tiene mes a no girar
                    FecVenci1 = FecVenci
                    If DBLet(Rs!mesnogir, "N") <> "0" Then
                        FecVenci1 = ComprobarMesNoGira(FecVenci1, DBLet(Rs!mesnogir, "N"), DBLet(Rs!DiaVtoAt, "N"), DBLet(Rs!DiaPago1, "N"), DBLet(Rs!DiaPago2, "N"), DBLet(Rs!DiaPago3, "N"))
                    End If
                        
                    CadValues = CadValues & ", " & cadValuesAux & i & ", '" & Format(FecVenci1, FormatoFecha) & "', "
                    CadValues2 = CadValues2 & ", " & CadValuesAux2 & i & ", " & Codmacta & ", " & ForPago & ", '" & Format(FecVenci1, FormatoFecha) & "', "
                    
                    'IMPORTE Resto de Vendimientos
                    ImpVenci = Round2(TotalFac / rsVenci!numerove, 2)
                    ImpVenci2 = Round2((TotalFac - ImpCheque) / rsVenci!numerove, 2)
                    CadValues = CadValues & DBSet(ImpVenci, "N") & ")"
                    CadValues2 = CadValues2 & DBSet(ImpVenci2, "N") & ", " & DBSet(CuentaPrev, "T") & ", " & DBSet(Banco, "N", "S") & ", " & DBSet(Sucursal, "N", "S") & ", " & DBSet(Digcontrol, "T", "S") & ", " & DBSet(CuentaBan, "T", "S") & ", " & textcsb33 & ", " & DBSet(Agente, "N") & ", "
                    CadValues2 = CadValues2 & DBSet(Me.DirDpto, "N", "S") '& ")"
                    '[Monica]22/11/2013: Tema iban
                    If vEmpresa.HayNorma19_34Nueva = 1 Then
                        CadValues2 = CadValues2 & ", " & DBSet(Iban, "T", "S") & ") "
                    Else
                        CadValues2 = CadValues2 & ") "
                    End If
                    
                Next i
                
'--monica
'                '--- Cheque regalo: laura 1/12/2006
'                'si hay cheque regalo insertar una linea más para la forma de pago correspondiente y el importe del cheque
'                If ImpCheque > 0 Then
'                    CadValues2 = CadValues2 & ", " & CadValuesAux2 & i & "," & Codmacta & ", " & vParamAplic.ForPagoChequeRegalo & ", "
'
'                    'FECHA VTO
'                    CadValuesAux2 = "primerve"
'                    TipForPago = DevuelveDesdeBDNew(conAri, "sforpa", "tipforpa", "codforpa", vParamAplic.ForPagoChequeRegalo, "N", CadValuesAux2)
'                    FecVenci = CDate(FecFactu)
'                    FecVenci = FecVenci + CByte(CadValuesAux2)
'                    'comprobar si tiene dias de pago y obtener la fecha del vencimiento correcta
'                    If TipForPago <> 0 Then
'            '            FecVenci = ComprobarFechaVenci(FecVenci, DBLet(RS!DiaPago1, "N"), DBLet(RS!DiaPago2, "N"), DBLet(RS!DiaPago3, "N"))
'                        MsgBox "FALTA cheque regalo con forma de pago no en EFECTIVO", vbInformation
'                    Else
'                        FecVenci = ComprobarFechaVenci(FecVenci, 0, 0, 0)
'                    End If
'
'                    CadValues2 = CadValues2 & DBSet(FecVenci, "F") & ", "
'                    CadValues2 = CadValues2 & DBSet(ImpCheque, "N") & ", '" & CuentaPrev & "', " & DBSet(Banco, "N") & ", " & DBSet(Sucursal, "N") & ", " & DBSet(DigControl, "T") & ", " & DBSet(CuentaBan, "T") & ", " & textcsb33 & ", " & DBSet(Agente, "N")
'                    'departamento
'                    CadValues2 = CadValues2 & "," & DBSet(Me.DirDpto, "N", "S") & ")"
'                End If
            End If
        End If
        Rs.Close
        Set Rs = Nothing
    End If
    rsVenci.Close
    Set rsVenci = Nothing
    
    If CadValues <> "" Then
        SQL = "INSERT INTO svenci (codtipom, numfactu, fecfactu, ordefect, fecefect, impefect)"
        SQL = SQL & " VALUES " & CadValues
        conn.Execute SQL
    End If
    
    
    'Grabar tabla scobro de la CONTABILIDAD
    '-------------------------------------------------
    If CadValues2 <> "" Then
        '01/09/06
'        If (NumTicket = "") Or (NumTicket <> "" And TipForPago <> 0) Then
            If CuentaPrev <> "" Then
                'antes de grabar en la scobro comprobar que existe en conta.sforpa la
                'forma de pago de la factura. Sino existe insertarla
                'vemos si existe en la conta
                b = InsertarFormaPagoEnConta(ForPago, MenError)
'--monica
'                'si hay cheque regalo comprobar q existe su forma de pago en conta
'                'sino insertarla
'                If b And ImpCheque > 0 Then b = InsertarFormaPagoEnConta(vParamAplic.ForPagoChequeRegalo, MenError)
                
                
'                CadValuesAux2 = DevuelveDesdeBDNew(conConta, "sforpa", "codforpa", "codforpa", ForPago, "N")
'                'si no existe la forma de pago en conta, insertamos la de ariges
'                If CadValuesAux2 = "" Then
'                    cadValuesAux = "tipforpa"
'                    CadValuesAux2 = DevuelveDesdeBDNew(conAri, "sforpa", "nomforpa", "codforpa", ForPago, "N", cadValuesAux)
'                    'insertamos e sforpa de la CONTA
'                    SQL = "INSERT INTO sforpa(codforpa,nomforpa,tipforpa)"
'                    SQL = SQL & " VALUES(" & ForPago & ", " & DBSet(CadValuesAux2, "T") & ", " & cadValuesAux & ")"
'                    ConnConta.Execute SQL
'                End If
                
                If b Then
                    'Insertamos en la tabla scobro de la CONTA
                    SQL = "INSERT INTO scobro (numserie, codfaccl, fecfaccl, numorden, codmacta, codforpa, fecvenci, impvenci,ctabanc1, codbanco, codsucur, digcontr, cuentaba,text33csb,agente,departamento" ') "
                    '[Monica]22/11/2013: Tema iban
                    If vEmpresa.HayNorma19_34Nueva = 1 Then
                        SQL = SQL & ", iban) "
                    Else
                        SQL = SQL & ") "
                    End If
                    SQL = SQL & " VALUES " & CadValues2
                    ConnConta.Execute SQL
                End If
            End If
'        End If

    End If
    
'    If b Then b = True
    
EInsertarTesoreria:
    If Err.Number <> 0 Then
        b = False
        MenError = "Insertar en Tesoreria: " & vbCrLf & Err.Description
    End If
    InsertarEnTesoreria = b
End Function



Private Sub ReiniciarImportesFactura()
    Me.BaseImp = 0
    Me.BaseIVA1 = 0
    Me.BaseIVA2 = 0
    Me.BaseIVA3 = 0
    
    Me.BrutoFac = 0
    
    Me.ImpIVA1 = 0
    Me.ImpIVA2 = 0
    Me.ImpIVA3 = 0
    
    Me.PorceIVA1 = 0
    Me.PorceIVA2 = 0
    Me.PorceIVA3 = 0
    
    Me.TipoIVA1 = 0
    Me.TipoIVA2 = 0
    Me.TipoIVA3 = 0
    
    Me.ImpGnral = 0
    Me.ImpPPago = 0
    Me.TotalFac = 0
End Sub


Public Function CalcularDatosFacturaBOD(vSocio As CSocio) As Boolean
'cadWhere: cad para la where de la SQL que selecciona las lineas del parte o la factura
'nomTabla: nombre de la tabla de advfactura_parte
Dim Rs As ADODB.Recordset
Dim i As Integer

Dim SQL As String
Dim cadAux As String
Dim cadAux1 As String

'Aqui vamos acumulando los totales
Dim TotBruto As Currency
Dim TotNeto As Currency
Dim TotImpIVA As Currency

Dim ImpAux As Currency
Dim ImpIVA As Currency
Dim ImpREC As Currency
Dim ImpBImIVA As Currency 'Importe Base imponible a la que hay q aplicar el IVA

Dim vBruto As Currency
Dim vNeto As Currency

Dim exentoIVA As Boolean
Dim conDesplaz As Boolean
    
Dim BaseImp As Currency
Dim BaseIVA1 As Currency
Dim BaseIVA2 As Currency
Dim BaseIVA3 As Currency
    
Dim BrutoFac As Currency
    
Dim ImpIVA1 As Currency
Dim ImpIVA2 As Currency
Dim ImpIVA3 As Currency
    
Dim PorceIVA1 As Currency
Dim PorceIVA2 As Currency
Dim PorceIVA3 As Currency
    
Dim ImpREC1 As Currency
Dim ImpREC2 As Currency
Dim ImpREC3 As Currency
    
Dim PorceREC1 As Currency
Dim PorceREC2 As Currency
Dim PorceREC3 As Currency
    
Dim TipoIVA1 As Currency
Dim TipoIVA2 As Currency
Dim TipoIVA3 As Currency
    
Dim ImpDto1 As Currency
Dim ImpDto2 As Currency
Dim TotalFac As Currency

Dim IvaAnt As Integer

Dim cadwhere As String
Dim cadWhere1 As String
    
Dim Nulo2 As String
Dim Nulo3 As String
Dim vSeccion As CSeccion
    
    CalcularDatosFacturaBOD = False
    On Error GoTo ECalcular

    BaseImp = 0
    BaseIVA1 = 0
    BaseIVA2 = 0
    BaseIVA3 = 0
    
    BrutoFac = 0
    
    ImpIVA1 = 0
    ImpIVA2 = 0
    ImpIVA3 = 0
    
    PorceIVA1 = 0
    PorceIVA2 = 0
    PorceIVA3 = 0
    
    ImpREC1 = 0
    ImpREC2 = 0
    ImpREC3 = 0
    
    PorceREC1 = 0
    PorceREC2 = 0
    PorceREC3 = 0
    
    TipoIVA1 = 0
    TipoIVA2 = 0
    TipoIVA3 = 0
    
    ImpDto1 = 0
    ImpDto2 = 0
    TotalFac = 0


    'Agrupar el importe bruto por tipos de iva
    cadwhere = "rbodfacturas.codtipom = " & DBSet(CodTipom, "T") & " and rbodfacturas.numfactu = " & DBSet(numfactu, "N") & " and fecfactu = " & DBSet(fecfactu, "F")
    cadWhere1 = Replace(cadwhere, "rbodfacturas", "rbodfacturas_lineas")
    
    SQL = SQL & "SELECT rbodfacturas_lineas.codigiva, sum(importel) as bruto, sum(importel) as neto"
    SQL = SQL & " FROM rbodfacturas_lineas "
    SQL = SQL & " WHERE " & cadWhere1
    SQL = SQL & " GROUP BY 1 "
    SQL = SQL & " ORDER BY 1 "

    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText

    TotBruto = 0
    TotNeto = 0
    TotImpIVA = 0
    vBruto = 0
    vNeto = 0
    i = 1

    If Not Rs.EOF Then Rs.MoveFirst
    IvaAnt = Rs.Fields(0).Value
    While Not Rs.EOF
        
        If IvaAnt <> Rs.Fields(0).Value Then
            TotBruto = TotBruto + vBruto
            TotNeto = TotNeto + vNeto
            ImpBImIVA = vNeto
        

            'Obtener el % de IVA
            cadAux = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", CStr(IvaAnt), "N")

            'aplicar el IVA a la base imponible de ese tipo
            ImpIVA = CalcularPorcentaje(ImpBImIVA, CCur(cadAux), 2)
            
            'sumamos todos los IVAS para sumarselo a la base imponible total de la factura
            'los vamos acumulando
            TotImpIVA = TotImpIVA + ImpIVA

'            If vSocio.CodIVA = 2 Then
'                'Obtener el % de RECARGO
'                cadAux1 = DevuelveDesdeBDNew(cConta, "tiposiva", "porcerec", "codigiva", CStr(IvaAnt), "N")
'
'                'aplicar el RECARGO a la base imponible de ese tipo
'                ImpREC = CalcularPorcentaje(ImpBImIVA, CCur(cadAux1), 2)
'
'                'sumamos todos los RECARGOS para sumarselo a la base imponible total de la factura
'                'los vamos acumulando
'                TotImpIVA = TotImpIVA + ImpREC
'            Else
                cadAux1 = "0"
                ImpREC = 0
'            End If


            Select Case i
                Case 1  'IVA 1
                    TipoIVA1 = IvaAnt 'RS!codigiva

                    BaseIVA1 = ImpBImIVA 'BASE IMPONIBLE

                    PorceIVA1 = cadAux '% de IVA

                    'Importe total con IVA
                    ImpIVA1 = ImpIVA
                    
                    PorceREC1 = cadAux1 '% de REC

                    'Importe total con RECARGO
                    ImpREC1 = ImpREC

                Case 2  'IVA 2
                    TipoIVA2 = IvaAnt 'RS!codigiva

                    BaseIVA2 = ImpBImIVA 'BASE IMPONIBLE

                    PorceIVA2 = cadAux '% de IVA

                    'Importe total con IVA
                    ImpIVA2 = ImpIVA

                    PorceREC2 = cadAux1 '% de REC

                    'Importe total con RECARGO
                    ImpREC2 = ImpREC
                Case 3  'IVA 3
                    TipoIVA3 = IvaAnt 'RS!codigiva

                    BaseIVA3 = ImpBImIVA 'BASE IMPONIBLE

                    PorceIVA3 = cadAux '% de IVA

                    'Importe total con IVA
                    ImpIVA3 = ImpIVA
                    
                    PorceREC3 = cadAux1 '% de REC

                    'Importe total con RECARGO
                    ImpREC3 = ImpREC
            End Select
            
            
            i = i + 1
            IvaAnt = Rs.Fields(0).Value
            vBruto = DBLet(Rs.Fields(1).Value, "N")
            vNeto = DBLet(Rs.Fields(2).Value, "N")
        Else
            vBruto = vBruto + DBLet(Rs.Fields(1).Value, "N")
            vNeto = vNeto + DBLet(Rs.Fields(2).Value, "N")
        End If
        
        
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing

    ' ULTIMO REGISTRO
    TotBruto = TotBruto + vBruto
    TotNeto = TotNeto + vNeto
    ImpBImIVA = vNeto


    'Obtener el % de IVA
    cadAux = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", CStr(IvaAnt), "N")

    'aplicar el IVA a la base imponible de ese tipo
    ImpIVA = CalcularPorcentaje(ImpBImIVA, CCur(cadAux), 2)
    
    'sumamos todos los IVAS para sumarselo a la base imponible total de la factura
    'los vamos acumulando
    TotImpIVA = TotImpIVA + ImpIVA
    
'    If CInt(vSocio.CodIVA) = 2 Then
'        'Obtener el % de RECARGO
'        cadAux1 = DevuelveDesdeBDNew(cConta, "tiposiva", "porcerec", "codigiva", CStr(IvaAnt), "N")
'
'        'aplicar el RECARGO a la base imponible de ese tipo
'        ImpREC = CalcularPorcentaje(ImpBImIVA, CCur(cadAux1), 2)
'    Else
        cadAux1 = "0"
        ImpREC = 0
'    End If
    'sumamos todos los RECARGOS para sumarselo a la base imponible total de la factura
    'los vamos acumulando
    TotImpIVA = TotImpIVA + ImpREC



    Select Case i
        Case 1  'IVA 1
            TipoIVA1 = IvaAnt

            BaseIVA1 = ImpBImIVA 'BASE IMPONIBLE

            PorceIVA1 = cadAux '% de IVA

            'Importe total con IVA
            ImpIVA1 = ImpIVA
            
            PorceREC1 = cadAux1 '% de REC

            'Importe total con RECARGO
            ImpREC1 = ImpREC

        Case 2  'IVA 2
            TipoIVA2 = IvaAnt

            BaseIVA2 = ImpBImIVA 'BASE IMPONIBLE

            PorceIVA2 = cadAux '% de IVA

            'Importe total con IVA
            ImpIVA2 = ImpIVA

            PorceREC2 = cadAux1 '% de REC

            'Importe total con RECARGO
            ImpREC2 = ImpREC
        Case 3  'IVA 3
            TipoIVA3 = IvaAnt

            BaseIVA3 = ImpBImIVA 'BASE IMPONIBLE

            PorceIVA3 = cadAux '% de IVA

            'Importe total con IVA
            ImpIVA3 = ImpIVA
            
            PorceREC3 = cadAux1 '% de REC

            'Importe total con RECARGO
            ImpREC3 = ImpREC
    End Select

    'Base Imponible
    BaseImp = TotNeto

    'TOTAL de la factura
    TotalFac = BaseImp + TotImpIVA

    Nulo2 = "N"
    Nulo3 = "N"
    If DBSet(TipoIVA2, "N", "S") = ValorNulo Then Nulo2 = "S"
    If DBSet(TipoIVA3, "N", "S") = ValorNulo Then Nulo3 = "S"
    
    'ACTUALIZAMOS LA FACTURA (tabla facturas)
    SQL = "update rbodfacturas "
    SQL = SQL & "set baseimp1 = " & DBSet(BaseIVA1, "N")
    SQL = SQL & ",baseimp2 = " & DBSet(BaseIVA2, "N", Nulo2)
    SQL = SQL & ",baseimp3 = " & DBSet(BaseIVA3, "N", Nulo3)
    SQL = SQL & ",impoiva1 = " & DBSet(ImpIVA1, "N")
    SQL = SQL & ",impoiva2 = " & DBSet(ImpIVA2, "N", Nulo2)
    SQL = SQL & ",impoiva3 = " & DBSet(ImpIVA3, "N", Nulo3)
    SQL = SQL & ",imporec1 = " & DBSet(ImpREC1, "N")
    SQL = SQL & ",imporec2 = " & DBSet(ImpREC2, "N", Nulo2)
    SQL = SQL & ",imporec3 = " & DBSet(ImpREC3, "N", Nulo3)
    SQL = SQL & ",codiiva1 = " & DBSet(TipoIVA1, "N")
    SQL = SQL & ",codiiva2 = " & DBSet(TipoIVA2, "N", Nulo2)
    SQL = SQL & ",codiiva3 = " & DBSet(TipoIVA3, "N", Nulo3)
    SQL = SQL & ",porciva1 = " & DBSet(PorceIVA1, "N")
    SQL = SQL & ",porciva2 = " & DBSet(PorceIVA2, "N", Nulo2)
    SQL = SQL & ",porciva3 = " & DBSet(PorceIVA3, "N", Nulo3)
    SQL = SQL & ",porcrec1 = " & DBSet(PorceREC1, "N")
    SQL = SQL & ",porcrec2 = " & DBSet(PorceREC2, "N", Nulo2)
    SQL = SQL & ",porcrec3 = " & DBSet(PorceREC3, "N", Nulo3)
    SQL = SQL & ",brutofac = " & DBSet(TotBruto, "N")
    SQL = SQL & ",totalfac = " & DBSet(TotalFac, "N")
    SQL = SQL & " where " & cadwhere
    
    conn.Execute SQL

    CalcularDatosFacturaBOD = True

ECalcular:
    If Err.Number <> 0 Then
        CalcularDatosFacturaBOD = False
    Else
        CalcularDatosFacturaBOD = True
    End If
End Function



Private Function InsertarCabeceraFactu(ByRef vSocio As CSocio, Observac As String) As Boolean
'Inserta la cabecera de la factura en la tabla: advfacturas
Dim SQL As String
Dim Nulo2 As String
Dim Nulo3 As String
Dim ctaDpto As Boolean
Dim Cad As String, cad2 As String

    On Error GoTo EInsertar

    'Insertar en la tabla cabecera de la factura de ventas
    SQL = "INSERT INTO rbodfacturas (codtipom,numfactu,fecfactu,codsocio,nifsocio,nomsocio,telsoci1,"
    SQL = SQL & "dirsocio, pobsocio,prosocio,codpostal,codforpa,"
    SQL = SQL & "observac,baseimp1,baseimp2,baseimp3,impoiva1,impoiva2,impoiva3,"
    SQL = SQL & "imporec1,imporec2,imporec3,codiiva1,codiiva2,codiiva3,porciva1,porciva2,porciva3,"
    SQL = SQL & "porcrec1,porcrec2,porcrec3,brutofac,totalfac,intconta,pasaridoc) "
    SQL = SQL & " VALUES (" & DBSet(CodTipom, "T") & "," & DBSet(numfactu, "N") & "," & DBSet(fecfactu, "F") & "," & DBSet(vSocio.Codigo, "N") & ","
    SQL = SQL & DBSet(vSocio.NIF, "T") & "," & DBSet(vSocio.Nombre, "T") & "," & DBSet(vSocio.Tfno1, "T") & ","
    SQL = SQL & DBSet(vSocio.Direccion, "T") & "," & DBSet(vSocio.Poblacion, "T") & "," & DBSet(vSocio.Provincia, "T") & ","
    SQL = SQL & DBSet(vSocio.CPostal, "T") & ","
    SQL = SQL & DBSet(ForPago, "N") & ","
    SQL = SQL & DBSet(Observac, "T", "S") & ","
    SQL = SQL & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
    SQL = SQL & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
    SQL = SQL & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
    SQL = SQL & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
    SQL = SQL & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
    SQL = SQL & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
    SQL = SQL & ValorNulo & "," & ValorNulo & ","
    SQL = SQL & "0,0)"
    
    conn.Execute SQL
    
    InsertarCabeceraFactu = True
    
EInsertar:
    If Err.Number <> 0 Then
        InsertarCabeceraFactu = False
        MuestraError Err.Number, "Insertar cabecera factura.", Err.Description
    Else
        InsertarCabeceraFactu = True
    End If
End Function





Private Function ClienteExentoIVA() As Boolean
'Dim vClien As CCliente
'Dim b As Boolean
'
'    Set vClien = New CCliente
'    vClien.Codigo = cliente
'    b = vClien.exentoIVA
'    Set vClien = Nothing
'
'    ClienteExentoIVA = b
End Function




Private Function InsertarFormaPagoEnConta(nForPa As String, cadErr As String) As Boolean
Dim cadAux As String
Dim cadAux2 As String
Dim SQL As String

    On Error GoTo ErrInsForpa
    InsertarFormaPagoEnConta = False
    
    'antes de grabar en la scobro comprobar que existe en conta.sforpa la
    'forma de pago de la factura. Sino existe insertarla
    
    'vemos si existe en la conta
    cadAux = DevuelveDesdeBDNew(cConta, "sforpa", "codforpa", "codforpa", nForPa, "N")
    'si no existe la forma de pago en conta, insertamos la de ariges
    If cadAux = "" Then
        cadAux2 = "tipforpa"
        cadAux = DevuelveDesdeBDNew(cAgro, "sforpa", "nomforpa", "codforpa", nForPa, "N", cadAux2)
        If cadAux <> "" Then
            'insertamos e sforpa de la CONTA
            SQL = "INSERT INTO sforpa(codforpa,nomforpa,tipforpa)"
            SQL = SQL & " VALUES(" & nForPa & ", " & DBSet(cadAux, "T") & ", " & cadAux2 & ")"
            ConnConta.Execute SQL
            InsertarFormaPagoEnConta = True
        Else
            InsertarFormaPagoEnConta = False
        End If
    Else
        InsertarFormaPagoEnConta = True
    End If
    
    
    Exit Function
    
ErrInsForpa:
    InsertarFormaPagoEnConta = False
    cadErr = "Insertar forma de pago en Contablilidad: " & vbCrLf & Err.Description
End Function




Private Function CalcularImporteDto(cantidad As String, Precio As String, ImpDto As String, Insertado As Boolean) As String
'Insertado: indica si ya hemos insertado el registro o no
'Calcula el Importe de una linea de Oferta, Pedido, Albaran, ...
'Importe=cantidad * precio - (descuentos)
Dim vCant As Currency
Dim vImp As Currency
Dim vDto As Currency
Dim vPre As Currency
Dim Rs As ADODB.Recordset
Dim SQL As String
Dim SumaBruto As Currency

On Error Resume Next


    SQL = "select sum(imporbru) from facturas_variedad where codtipom = " & DBSet(CodTipom, "T")
    SQL = SQL & " and numfactu = " & DBSet(numfactu, "N") & " and fecfactu = " & DBSet(fecfactu, "F")
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    SumaBruto = 0
    If Not Rs.EOF Then
        SumaBruto = DBLet(Rs.Fields(0).Value, "N")
    End If
    
    'Como son de tipo string comprobar que si vale "" lo ponemos a 0
    vCant = ComprobarCero(cantidad)
    vPre = ComprobarCero(Precio)
    vDto = ComprobarCero(ImpDto)
    
    If Not Insertado Then
        SumaBruto = SumaBruto + (CCur(vCant) * CCur(vPre))
    End If
    
    If SumaBruto <> 0 Then
        vImp = (CCur(cantidad) * CCur(vPre) * CCur(vDto)) / SumaBruto
    Else
        vImp = CCur(vDto)
    End If
    
    vImp = Round2(vImp, 6)
    
    CalcularImporteDto = CStr(vImp)

End Function




Public Function PasarAlbaranesAFactura2(TipoAlb As Byte, cadSQL As String, textoCSB As String, Forpa As String, Optional ErroresAux As String, Optional EstaRecu As Boolean) As Boolean
'IN ->  cadSQL: cadena para seleccion de los Albaranes que vamos a Facturar
'       TextosCSB:  Para la tesoreria
'Desde Albaranes Genera las Facturas correspondientes
Dim b As Boolean
Dim MenError As String

    On Error GoTo ETraspasoAlbFac

    PasarAlbaranesAFactura2 = False
    b = False
    ErroresAux = ""
    '#Laura 14/11/2006 recuperar facturas de Alzira
    'si esta recuperando facturas
    EstaRecuFac = EstaRecu
    
    
'[Monica]24/02/2010: el tipo de movimiento lo cogeremos de la coopoerativa
'    Select Case TipoAlb
'        Case 0 'lineas de variedades de almazara
'            CodTipom = "FZA" ' Facturas de retirada de bodega y almaZAra
'        Case 1 ' lineas de variedades de bodega
'            CodTipom = "FAB" ' FActuras de retirada de bodega
'    End Select
      
     
    'Aqui empieza transaccion
    conn.BeginTrans
   
    'Insertar la Factura (advfacturas, advfacturas_lineas)
    MenError = "Error al Insertar en tablas de factura."
    b = InsertarFacturaBOD(cadSQL, TipoAlb)
    
ETraspasoAlbFac:
    If Err.Number <> 0 Then
        ErroresAux = Err.Description
        Err.Clear
        b = False
    End If
  
    If b Then
        conn.CommitTrans
        PasarAlbaranesAFactura2 = True
    Else
        conn.RollbackTrans
        PasarAlbaranesAFactura2 = False
        ErroresAux = MenError & vbCrLf & ErroresAux & vbCrLf
    End If
'    espera 0.4
End Function




Private Function InsertarFacturaBOD(cadSQL As String, TipoAlb As Byte) As Boolean
'Insertamos en las tablas de factura de socios: advfacturas, advfacturas_lineas
Dim b As Boolean
Dim vTipoMov As CTiposMov
Dim Existe As Boolean
Dim devuelve As String
Dim SQL As String
Dim Rs As ADODB.Recordset
Dim vSocio As CSocio
Dim TipoMovimiento As String

    InsertarFacturaBOD = False

    'cogemos los datos del albaran
    SQL = "select codsocio, fechaalb, observac from rbodalbaran where " & cadSQL
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText

    Set vSocio = New CSocio

    'si se ha modificado el cliente volver a cargar los datos
    If vSocio.Existe(Rs.Fields(0).Value) Then
        If vSocio.LeerDatos(Rs.Fields(0).Value) Then
            Socio = vSocio.Codigo

            Select Case TipoAlb
                Case 0 'lineas de variedades de almazara
                    CodTipom = vSocio.CodTipomFacAlmz '"FZA" ' Facturas de retirada de bodega y almaZAra
                Case 1 ' lineas de variedades de bodega
                    CodTipom = vSocio.CodTipomFacBod '"FAB"  FActuras de retirada de bodega
            End Select


            TipoMovimiento = CodTipom
            Set vTipoMov = New CTiposMov
            If vTipoMov.Leer(TipoMovimiento) Then
                numfactu = vTipoMov.ConseguirContador(TipoMovimiento)
                ' si existe la factura incrementamos el contador
                Do
                    devuelve = DevuelveDesdeBDNew(cAgro, "rbodfacturas", "numfactu", "codtipom", TipoMovimiento, "T", , "numfactu", numfactu, "N", "fecfactu", Rs.Fields(1).Value, "F")
                    If devuelve <> "" Then
                        'Ya existe el contador incrementarlo
                        Existe = True
                        vTipoMov.IncrementarContador (TipoMovimiento)
                        numfactu = vTipoMov.ConseguirContador(TipoMovimiento)
                    Else
                        Existe = False
                    End If
                Loop Until Not Existe
            Else
                Exit Function
            End If

            CodTipom = TipoMovimiento

            'Cabecera Factura
            b = InsertarCabeceraFactu(vSocio, DBLet(Rs.Fields(2).Value, "T"))
            If Not b Then Exit Function
    
            'Cabeceras Albaranes de la Factura
            b = InsertarCabAlbaranesFactu(cadSQL)
            If Not b Then Exit Function
    
            'Insertar lineas de partes de la factura
            b = InsertarLineasFactuBOD(cadSQL)
            If Not b Then Exit Function

            'calcular el total de la factura
            b = CalcularDatosFacturaBOD(vSocio)
            If Not b Then Exit Function

            'Incrementar contador del tipo de movimiento
            '-------------------------------------------------------------
'--monica: en cualquier caso se incrementa contador
'            If vCliente.TipoFactu = 0 Then
                vTipoMov.IncrementarContador (CodTipom)
                Set vTipoMov = Nothing
'            End If

            InsertarFacturaBOD = True

        End If
    End If
    Set vSocio = Nothing

End Function


Private Function EliminarAlbaranes(cadSQL As String) As Boolean
'Eliminamos de las tablas de Partes: advpartes, advpartes_lineas
Dim SQL As String

    On Error GoTo EEliminar

    EliminarAlbaranes = False
    
    'ELiminar lineas partes
    SQL = "DELETE FROM rbodalbaran_variedad "
    SQL = SQL & " WHERE " & Replace(cadSQL, "rbodalbaran", "rbodalbaran_variedad")
    conn.Execute SQL
    
    espera 0.1
    
    'Eliminar Cabeceras Partes
    SQL = "DELETE FROM rbodalbaran "
    SQL = SQL & " WHERE " & Replace(cadSQL, "rbodalbaran_variedad", "rbodalbaran")
    conn.Execute SQL
        
    EliminarAlbaranes = True

EEliminar:
    If Err.Number <> 0 Then
        EliminarAlbaranes = False
    Else
        EliminarAlbaranes = True
    End If
End Function


Private Function InsertarLineasFactuBOD(cadSQL As String) As Boolean
'Insertar las lineas de los Albaranes de envases de la factura en la tabla: slifac
Dim SQL As String
Dim Sql2 As String
Dim Cad As String
Dim cad2 As String
Dim NumF As Long
Dim Rs As ADODB.Recordset

    On Error GoTo EInsertarLin

    SQL = "INSERT INTO rbodfacturas_lineas(codtipom,numfactu,fecfactu,numalbar,numlinea,codvarie,unidades,cantidad,precioar,dtolinea,importel,ampliaci,codigiva) values "
    
     'pasamos lineas de partes a lineas factura
    Sql2 = " SELECT numalbar,numlinea,rbodalbaran_variedad.codvarie,unidades,cantidad,precioar,dtolinea,importel, ampliaci, rbodalbaran_variedad.codigiva "
    Sql2 = Sql2 & " FROM (rbodalbaran_variedad INNER JOIN variedades ON rbodalbaran_variedad.codvarie = variedades.codvarie) "
    Sql2 = Sql2 & " INNER JOIN productos ON variedades.codprodu = productos.codprodu "
    Sql2 = Sql2 & " WHERE " & Replace(cadSQL, "rbodalbaran", "rbodalbaran_variedad")
    
    Select Case CodTipom
        Case "FZA"
            Sql2 = Sql2 & " and productos.codgrupo = 5"
        Case "FAB"
            Sql2 = Sql2 & " and productos.codgrupo = 6"
    End Select

    Set Rs = New ADODB.Recordset
    Rs.Open Sql2, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    Cad = "(" & DBSet(CodTipom, "T") & "," & DBSet(numfactu, "N") & "," & DBSet(fecfactu, "F") & ","
    cad2 = ""
    
    NumF = 1
    
    While Not Rs.EOF
        cad2 = cad2 & Cad & DBSet(Rs.Fields(0).Value, "N") & "," & DBSet(Rs.Fields(1).Value, "N") & ","
        cad2 = cad2 & DBSet(Rs.Fields(2).Value, "N") & "," & DBSet(Rs.Fields(3).Value, "T") & ","
        cad2 = cad2 & DBSet(Rs.Fields(4).Value, "N") & "," 'cantidad
        cad2 = cad2 & DBSet(Rs.Fields(5).Value, "N") & "," 'precio
        cad2 = cad2 & DBSet(Rs.Fields(6).Value, "N") & "," 'dtolinea
        cad2 = cad2 & DBSet(Rs.Fields(7).Value, "N") & "," 'importe
        cad2 = cad2 & DBSet(Rs.Fields(8).Value, "T", "S") & "," 'ampliacion
        cad2 = cad2 & DBSet(Rs.Fields(9).Value, "N") & ")," 'codigiva
        
        NumF = NumF + 1
        
        Rs.MoveNext
    Wend
    
    ' quitamos la ultima coma
    If cad2 <> "" Then
        cad2 = Mid(cad2, 1, Len(cad2) - 1)
    
        conn.Execute SQL & cad2
    End If

    InsertarLineasFactuBOD = True
    
EInsertarLin:
    If Err.Number <> 0 Then
        InsertarLineasFactuBOD = False
        MuestraError Err.Number, "Insertar lineas factura albaranes retirada.", Err.Description
    Else
        InsertarLineasFactuBOD = True
    End If
End Function

Private Function InsertarCabAlbaranesFactu(cadSQL As String) As Boolean
'Insertar las cabeceras de los Partes de la factura ADV en la tabla: advfacturas_partes
Dim SQL As String
Dim devuelve As String
Dim MantenimientoObservacion As String

    On Error GoTo EInsertarAlb

    devuelve = ""
    cadSQL = Replace(cadSQL, "rbodalbaran_variedad", "rbodalbaran")
    
    SQL = "INSERT INTO rbodfacturas_alb (codtipom,numfactu,fecfactu,numalbar,fechaalb,observac)"
    
    SQL = SQL & " SELECT '" & CodTipom & "' as codtipom," & DBSet(numfactu, "N") & " as numfactu," & DBSet(fecfactu, "F") & " as fecfactu," & "numalbar,fechaalb,observac "
    SQL = SQL & " FROM rbodalbaran WHERE " & cadSQL
    
    conn.Execute SQL

    InsertarCabAlbaranesFactu = True
    
EInsertarAlb:
    If Err.Number <> 0 Then
        InsertarCabAlbaranesFactu = False
        MuestraError Err.Number, "Insertar cabecera albaranes retirada factura.", Err.Description
    Else
        InsertarCabAlbaranesFactu = True
    End If
End Function

