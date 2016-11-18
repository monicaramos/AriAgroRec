VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFacturaADV"
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

Private mIban As String     'cod. iban
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

Public Property Let nif(ByVal vData As String)
     mNIFsocio = vData
End Property

Public Property Get nif() As String
     nif = mNIFsocio
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
Dim Sql As String

    On Error GoTo ELeer
    
    
    LeerDatos = False
    Sql = "SELECT codtipom,numfactu,fecfactu,codsocio,codforpa,codbanco,codsucur,digcontr,cuentaba,brutofac, totalfac FROM advfacturas "
    Sql = Sql & " WHERE codtipom='" & vCodtipom & "'"
    Sql = Sql & " AND numfactu=" & vNumfactu
    Sql = Sql & " AND fecfactu='" & Format(vFecfactu, FormatoFecha) & "'"
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
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
        TipForPago = DevuelveDesdeBDNew(cAgro, "forpago", "tipforpa", "codforpa", ForPago, "N")
        mIban = DBLet(Rs!Iban, "T")
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
            If vParamAplic.ContabilidadNueva Then
                Sql = "SELECT COUNT(*) FROM cobros WHERE numserie='" & LetraSerie & "' and numfactu=" & numfactu
                Sql = Sql & " AND fecfactu=" & DBSet(fecfactu, "F")
            Else
                Sql = "SELECT COUNT(*) FROM scobro WHERE numserie='" & LetraSerie & "' and codfaccl=" & numfactu
                Sql = Sql & " AND fecfaccl=" & DBSet(fecfactu, "F")
            End If
            If RegistrosAListar(Sql, cConta) > 0 Then
                If vParamAplic.ContabilidadNueva Then
                    Sql = DevuelveDesdeBDNew(cConta, "cobros", "ctabanc1", "numserie", mLetraSer, "T", , "numfactu", vNumfactu, "N", "fecfactu", vFecfactu, "F")
                Else
                    Sql = DevuelveDesdeBDNew(cConta, "scobro", "ctabanc1", "numserie", mLetraSer, "T", , "codfaccl", vNumfactu, "N", "fecfaccl", vFecfactu, "F")
                End If
                If Sql <> "" Then
                    mCtaPrev = Sql
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


Public Function CalcularDatosFacturaADV(vSocio As cSocio) As Boolean
'cadWhere: cad para la where de la SQL que selecciona las lineas del parte o la factura
'nomTabla: nombre de la tabla de advfactura_parte
Dim Rs As ADODB.Recordset
Dim i As Integer

Dim Sql As String
Dim cadAux As String
Dim cadAux1 As String

'Aqui vamos acumulando los totales
Dim TotBruto As Currency
Dim TotNeto As Currency
Dim TotImpIVA As Currency

Dim ImpAux As Currency
Dim impiva As Currency
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

Dim cadWHERE As String
Dim cadWhere1 As String
    
Dim Nulo2 As String
Dim Nulo3 As String
Dim vSeccion As CSeccion
    
    CalcularDatosFacturaADV = False
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
    cadWHERE = "advfacturas.codtipom = " & DBSet(CodTipom, "T") & " and advfacturas.numfactu = " & DBSet(numfactu, "N") & " and fecfactu = " & DBSet(fecfactu, "F")
    cadWhere1 = Replace(cadWHERE, "advfacturas", "advfacturas_lineas")
    
    ' si la facturacion es interna el codigo de iva es el exento de parametros
    ' sino es el del articulo
    If CodTipom = "FIN" Then
        Sql = "SELECT " & vParamAplic.CodIvaExeADV & " as codigiva , sum(importel) as bruto, sum(importel) as neto"
    Else
        Sql = "SELECT advfacturas_lineas.codigiva, sum(importel) as bruto, sum(importel) as neto"
    End If
    
    Sql = Sql & " FROM advfacturas_lineas "
    Sql = Sql & " WHERE " & cadWhere1
    Sql = Sql & " GROUP BY 1 "
    Sql = Sql & " ORDER BY 1 "

    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText

    TotBruto = 0
    TotNeto = 0
    TotImpIVA = 0
    vBruto = 0
    vNeto = 0
    i = 1
    IvaAnt = Rs.Fields(0).Value

    If Not Rs.EOF Then Rs.MoveFirst
    
    While Not Rs.EOF
        If IvaAnt <> Rs.Fields(0).Value Then
            TotBruto = TotBruto + vBruto
            TotNeto = TotNeto + vNeto
            ImpBImIVA = vNeto
        

            'Obtener el % de IVA
            cadAux = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", CStr(IvaAnt), "N")

            'aplicar el IVA a la base imponible de ese tipo
            impiva = CalcularPorcentaje(ImpBImIVA, CCur(cadAux), 2)
            
            'sumamos todos los IVAS para sumarselo a la base imponible total de la factura
            'los vamos acumulando
            TotImpIVA = TotImpIVA + impiva

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
                    ImpIVA1 = impiva
                    
                    PorceREC1 = cadAux1 '% de REC

                    'Importe total con RECARGO
                    ImpREC1 = ImpREC

                Case 2  'IVA 2
                    TipoIVA2 = IvaAnt 'RS!codigiva

                    BaseIVA2 = ImpBImIVA 'BASE IMPONIBLE

                    PorceIVA2 = cadAux '% de IVA

                    'Importe total con IVA
                    ImpIVA2 = impiva

                    PorceREC2 = cadAux1 '% de REC

                    'Importe total con RECARGO
                    ImpREC2 = ImpREC
                Case 3  'IVA 3
                    TipoIVA3 = IvaAnt 'RS!codigiva

                    BaseIVA3 = ImpBImIVA 'BASE IMPONIBLE

                    PorceIVA3 = cadAux '% de IVA

                    'Importe total con IVA
                    ImpIVA3 = impiva
                    
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
    impiva = CalcularPorcentaje(ImpBImIVA, CCur(cadAux), 2)
    
    'sumamos todos los IVAS para sumarselo a la base imponible total de la factura
    'los vamos acumulando
    TotImpIVA = TotImpIVA + impiva
    
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
            ImpIVA1 = impiva
            
            PorceREC1 = cadAux1 '% de REC

            'Importe total con RECARGO
            ImpREC1 = ImpREC

        Case 2  'IVA 2
            TipoIVA2 = IvaAnt

            BaseIVA2 = ImpBImIVA 'BASE IMPONIBLE

            PorceIVA2 = cadAux '% de IVA

            'Importe total con IVA
            ImpIVA2 = impiva

            PorceREC2 = cadAux1 '% de REC

            'Importe total con RECARGO
            ImpREC2 = ImpREC
        Case 3  'IVA 3
            TipoIVA3 = IvaAnt

            BaseIVA3 = ImpBImIVA 'BASE IMPONIBLE

            PorceIVA3 = cadAux '% de IVA

            'Importe total con IVA
            ImpIVA3 = impiva
            
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
    Sql = "update advfacturas "
    Sql = Sql & "set baseimp1 = " & DBSet(BaseIVA1, "N")
    Sql = Sql & ",baseimp2 = " & DBSet(BaseIVA2, "N", Nulo2)
    Sql = Sql & ",baseimp3 = " & DBSet(BaseIVA3, "N", Nulo3)
    Sql = Sql & ",impoiva1 = " & DBSet(ImpIVA1, "N")
    Sql = Sql & ",impoiva2 = " & DBSet(ImpIVA2, "N", Nulo2)
    Sql = Sql & ",impoiva3 = " & DBSet(ImpIVA3, "N", Nulo3)
    Sql = Sql & ",imporec1 = " & DBSet(ImpREC1, "N")
    Sql = Sql & ",imporec2 = " & DBSet(ImpREC2, "N", Nulo2)
    Sql = Sql & ",imporec3 = " & DBSet(ImpREC3, "N", Nulo3)
    Sql = Sql & ",codiiva1 = " & DBSet(TipoIVA1, "N")
    Sql = Sql & ",codiiva2 = " & DBSet(TipoIVA2, "N", Nulo2)
    Sql = Sql & ",codiiva3 = " & DBSet(TipoIVA3, "N", Nulo3)
    Sql = Sql & ",porciva1 = " & DBSet(PorceIVA1, "N")
    Sql = Sql & ",porciva2 = " & DBSet(PorceIVA2, "N", Nulo2)
    Sql = Sql & ",porciva3 = " & DBSet(PorceIVA3, "N", Nulo3)
    Sql = Sql & ",porcrec1 = " & DBSet(PorceREC1, "N")
    Sql = Sql & ",porcrec2 = " & DBSet(PorceREC2, "N", Nulo2)
    Sql = Sql & ",porcrec3 = " & DBSet(PorceREC3, "N", Nulo3)
    Sql = Sql & ",brutofac = " & DBSet(TotBruto, "N")
    Sql = Sql & ",totalfac = " & DBSet(TotalFac, "N")
    Sql = Sql & " where " & cadWHERE
    
    conn.Execute Sql

    CalcularDatosFacturaADV = True

ECalcular:
    If Err.Number <> 0 Then
        CalcularDatosFacturaADV = False
    Else
        CalcularDatosFacturaADV = True
    End If
End Function



Public Function PasarAlbaranAFactura(cadSQL As String, FechaFact As String, Optional ErroresAux As String, Optional EstaRecu As Boolean) As Boolean
'IN -> cadSQL: cadena para seleccion de los Albaranes que vamos a Facturar
'Desde Albaranes Genera las Facturas correspondientes
Dim b As Boolean
Dim MenError As String

    On Error GoTo ETraspasoAlbFac

    PasarAlbaranAFactura = False
    b = False
    ErroresAux = ""
      
    'Aqui empieza transaccion
    conn.BeginTrans
    ConnConta.BeginTrans
   
    fecfactu = FechaFact
   
    'Insertar la Factura (scafac,scafac1,slifac)
    MenError = "Error al Insertar en tablas de factura."
    b = InsertarFactura(cadSQL)
    
    
'    'Insertar los Vencimientos de la Factura. Tabla: svenci
'    'Grabar en TESORERIA. Tabla de Contabilidad: sconta.scobros
'    '#Laura: 14/11/2006 Recuperar facturas Alzira
'    If b Then
'        If EstaRecuFac = False Then
'            MenError = "Error al actualizar en Tesoreria."
'            b = InsertarEnTesoreria(MenError)
'        Else
'            MenError = "Error al actualizar en Tesoreria."
'            b = InsertarEnTesoreria_RecupFac(MenError)
'        End If
'    End If
'
    
    
ETraspasoAlbFac:
    If Err.Number <> 0 Then
        ErroresAux = Err.Description
        Err.Clear
        b = False
    End If
  
    If b Then
        conn.CommitTrans
        ConnConta.CommitTrans
        PasarAlbaranAFactura = True
    Else
        conn.RollbackTrans
        ConnConta.RollbackTrans
        PasarAlbaranAFactura = False
        ErroresAux = MenError & vbCrLf & ErroresAux & vbCrLf
    End If
    espera 0.4
End Function


Private Function InsertarFactura(cadSQL As String) As Boolean
'Insertamos en las tablas de factura de cliente: scafac, scafac1, slifac
Dim b As Boolean
Dim vTipoMov As CTiposMov
Dim Existe As Boolean
Dim devuelve As String
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim vSocio As cSocio
Dim TipoMovimiento As String

    InsertarFactura = False
    
    'cogemos los datos del albaran
    Sql = "select codclien, fechaalb, observac from albaran where " & cadSQL
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    Set vSocio = New cSocio
    
    'si se ha modificado el cliente volver a cargar los datos
    If vSocio.Existe(Rs.Fields(0).Value) Then
        If vSocio.LeerDatos(Rs.Fields(0).Value) Then
             Socio = vSocio.Codigo
            
            TipoMovimiento = "FAP"
            Set vTipoMov = New CTiposMov
            If vTipoMov.Leer(TipoMovimiento) Then
                numfactu = vTipoMov.ConseguirContador(TipoMovimiento)
                ' si existe la factura incrementamos el contador
                Do
                    devuelve = DevuelveDesdeBDNew(cAgro, "advfacturas", "numfactu", "codtipom", TipoMovimiento, "T", , "numfactu", numfactu, "N", "fecfactu", Rs.Fields(1).Value, "F")
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
             
            'Insertar lineas de variedades de la factura
            b = InsertarLineasFactu(cadSQL, vSocio)
            If Not b Then Exit Function
                         
            
            'calcular el total de la factura
            b = CalcularDatosFacturaADV(vSocio)
            If Not b Then Exit Function
            
            'Incrementar contador del tipo de movimiento
            '-------------------------------------------------------------
'--monica: en cualquier caso se incrementa contador
'            If vCliente.TipoFactu = 0 Then
                vTipoMov.IncrementarContador (CodTipom)
                Set vTipoMov = Nothing
'            End If
            
            InsertarFactura = True
            
        End If
    End If
    Set vSocio = Nothing
    
End Function


Private Function InsertarCabeceraFactu(ByRef vSocio As cSocio, Observac As String) As Boolean
'Inserta la cabecera de la factura en la tabla: advfacturas
Dim Sql As String
Dim Nulo2 As String
Dim Nulo3 As String
Dim ctaDpto As Boolean
Dim Cad As String, Cad2 As String

    On Error GoTo EInsertar

    'Insertar en la tabla cabecera de la factura de ventas
    Sql = "INSERT INTO advfacturas (codtipom,numfactu,fecfactu,codsocio,nifsocio,nomsocio,telsoci1,"
    Sql = Sql & "dirsocio, pobsocio,prosocio,codpostal,codforpa,"
    Sql = Sql & "observac,baseimp1,baseimp2,baseimp3,impoiva1,impoiva2,impoiva3,"
    Sql = Sql & "imporec1,imporec2,imporec3,codiiva1,codiiva2,codiiva3,porciva1,porciva2,porciva3,"
    Sql = Sql & "porcrec1,porcrec2,porcrec3,brutofac,totalfac,intconta,pasaridoc) "
    Sql = Sql & " VALUES (" & DBSet(CodTipom, "T") & "," & DBSet(numfactu, "N") & "," & DBSet(fecfactu, "F") & "," & DBSet(vSocio.Codigo, "N") & ","
    Sql = Sql & DBSet(vSocio.nif, "T") & "," & DBSet(vSocio.Nombre, "T") & "," & DBSet(vSocio.Tfno1, "T") & ","
    Sql = Sql & DBSet(vSocio.Direccion, "T") & "," & DBSet(vSocio.Poblacion, "T") & "," & DBSet(vSocio.Provincia, "T") & ","
    Sql = Sql & DBSet(vSocio.CPostal, "T") & ","
    Sql = Sql & DBSet(ForPago, "N") & ","
    Sql = Sql & DBSet(Observac, "T", "S") & ","
    Sql = Sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
    Sql = Sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
    Sql = Sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
    Sql = Sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
    Sql = Sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
    Sql = Sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
    Sql = Sql & ValorNulo & "," & ValorNulo & ","
    Sql = Sql & "0,0)"
    
    conn.Execute Sql
    
    InsertarCabeceraFactu = True
    
EInsertar:
    If Err.Number <> 0 Then
        InsertarCabeceraFactu = False
        MuestraError Err.Number, "Insertar cabecera factura.", Err.Description
    Else
        InsertarCabeceraFactu = True
    End If
End Function


Private Function InsertarCabAlbaranesFactu(cadSQL As String) As Boolean
'Insertar las cabeceras de los Albaranes de la factura en la tabla: scafac1
'NumTicket: para cuando de venta generamos Factura en numalbar guardamos el ticket que la genero
Dim Sql As String
Dim devuelve As String
Dim MantenimientoObservacion As String

    On Error GoTo EInsertarAlb

    devuelve = ""
    If CodTipom = "FTI" Then
'        cadSQL = Replace(cadSQL, "sliven", "scaven")
    Else
        If InStr(cadSQL, "scaven") > 0 Then 'viene de generar factura en ticket
'            codTipoM = "FTI"
        Else
            cadSQL = Replace(cadSQL, "slialb", "scaalb")
        End If
    End If
    
    Sql = "INSERT INTO scafac1 (codtipom,numfactu,fecfactu,codtipoa,numalbar,fechaalb,numpedcl,fecpedcl,sementre,numofert,fecofert,referenc,codenvio,codtraba,codtrab1,codtrab2,observa1,observa2,observa3,observa4,observa5,numtermi,numventa)"
    
    If (CodTipom = "FAV" And NumTicket <> "") Then
        'Facturas generadas desde una venta de tickets: en numalbar guardamos numticket
'        SQL = SQL & " VALUES('" & codTipoM & "', " & NumFactu & ", '" & Format(FecFactu, FormatoFecha) & "','FTI'," & NumTicket & ","
        '01/09/06 laura
        Sql = Sql & " VALUES('" & CodTipom & "', " & numfactu & ", '" & Format(fecfactu, FormatoFecha) & "','ATI'," & NumTicket & ","
        
        Sql = Sql & "'" & Format(fecfactu, FormatoFecha) & "'," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," 'datos de pedido a nulo
        Sql = Sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," 'datos de oferta a nulo
        'obtenemos el codigo de envio de cliente
        devuelve = DevuelveDesdeBDNew(cAgro, "rsocios", "codenvio", "codsocio", Socio, "N")
        Sql = Sql & DBSet(devuelve, "N") & ","
        'trabajadores
        Sql = Sql & DBSet(OpeFactu, "N") & "," & ValorNulo & "," & DBSet(OpeFactu, "N") & ","
        'observaciones
        '---- Laura 10/10/06: poner en la observacion 1 el valor de scaven.observa1
        devuelve = DevuelveDesdeBDNew(cAgro, "scaven", "observa1", "numtermi", Me.NumTerminal, "N", , "numventa", Me.NumVenta, "N", "fecventa", Me.fecfactu, "F")
        Sql = Sql & DBSet(devuelve, "T") & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ","
        
        Sql = Sql & NumTerminal & "," & NumVenta & ")"
        
    ElseIf (CodTipom <> "FAM" And CodTipom <> "FTI") And cadSQL <> "" Then 'Facturas de VEnta
        Sql = Sql & " SELECT '" & CodTipom & "' as codtipom," & DBSet(numfactu, "N") & " as numfactu," & DBSet(fecfactu, "F") & " as fecfactu," & "codtipom as codtipoa, numalbar,fechaalb,numpedcl,fecpedcl,sementre,numofert,fecofert,referenc,codenvio,codtraba,codtrab1,codtrab2,observa01,observa02,observa03,observa04,observa05,numtermi,numventa "
        Sql = Sql & " FROM scaalb WHERE " & cadSQL
        
        
    Else 'Bonificaciones/Desplazamientos
        'Facturas de Mantenimiento
        'Facturas de Ticket del TPV
        Sql = Sql & " VALUES('" & CodTipom & "', " & numfactu & ", '" & Format(fecfactu, FormatoFecha) & "',"
'        If (codTipoM = "FAM" Or codTipoM = "FTI") Then
        '01/09/06 laura
        If (CodTipom = "FAM") Then
            Sql = Sql & "'',0,"
        ElseIf CodTipom = "FTI" Then
            Sql = Sql & "'ATI'," & NumAlbTicket & ","
        Else
            Sql = Sql & "'',9999999,"
        End If
        Sql = Sql & "'" & Format(fecfactu, FormatoFecha) & "'," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," 'datos de pedido a nulo
        Sql = Sql & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," 'datos de oferta a nulo
        'obtenemos el codigo de envio de cliente
        devuelve = DevuelveDesdeBDNew(cAgro, "sclien", "codenvio", "codclien", Socio, "N")
        Sql = Sql & DBSet(devuelve, "N") & ","
        'trabajadores
        Sql = Sql & DBSet(OpeFactu, "N") & "," & ValorNulo & "," & DBSet(OpeFactu, "N") & ","
        'observaciones 1
        If CodTipom = "FTI" Then
            '---- Laura 19/12/06: poner en la observacion 1 el valor de scaven.observa1
            devuelve = DevuelveDesdeBDNew(cAgro, "scaven", "observa1", "numtermi", Me.NumTerminal, "N", , "numventa", Me.NumVenta, "N", "fecventa", Me.fecfactu, "F")
        Else
            'La linea viene en el campo concefac que hemos guardado en observaciones
            If CodTipom = "FAM" Then
                
                MantenimientoObservacion = DevuelveDesdeBDNew(cAgro, "scaman", "tipopago", "codclien", Me.Socio, "N", , "nummante", numMante, "T")
                If MantenimientoObservacion = "1" Then 'Trimestral
                    MantenimientoObservacion = "TRIMESTRAL DE " & Year(fecfactu)
                ElseIf MantenimientoObservacion = "2" Then 'Semestral
                    MantenimientoObservacion = "SEMESTRAL DE " & Year(fecfactu)
                ElseIf MantenimientoObservacion = "3" Then 'Anual
                    MantenimientoObservacion = "ANUAL " & Year(fecfactu)
                Else
                    'si es mensual: ej: "JUNIO de 2005" (ponemos el mes seleccionado para facturar)
                    MantenimientoObservacion = MonthName(CLng(MesFactu)) & " de " & Year(fecfactu)
                    MantenimientoObservacion = UCase(MantenimientoObservacion)
                End If
                                                            
                If mObservacion <> "" Then
                    devuelve = MantenimientoObservacion
                    
                Else
                    devuelve = ""
                    mObservacion = MantenimientoObservacion
                End If
                
            Else
                devuelve = ""
            End If
        End If
        Sql = Sql & DBSet(devuelve, "T") & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo
        Sql = Sql & "," & DBSet(NumTerminal, "N", "S") & "," & DBSet(NumVenta, "N", "S") & ")"
    End If
    
    conn.Execute Sql

    InsertarCabAlbaranesFactu = True
    
EInsertarAlb:
    If Err.Number <> 0 Then
        InsertarCabAlbaranesFactu = False
        MuestraError Err.Number, "Insertar cabecera albaranes factura.", Err.Description
    Else
        InsertarCabAlbaranesFactu = True
    End If
End Function


Private Function InsertarLineasFactu(vSQL As String, ByRef vSocio As cSocio) As Boolean
''Insertar las lineas de los Albaranes de la factura en la tabla: slifac
'Dim sql As String
'Dim Sql2 As String
'Dim cad As String
'Dim cad2 As String
'Dim vWhere2 As String
'Dim cArt As CArticulo
'Dim Rs As ADODB.Recordset
'Dim NumF As Long
'Dim TipoIva As Integer
'Dim vSql1 As String
'
'    On Error GoTo EInsertarLin
'
'    sql = "INSERT INTO facturas_variedad (codtipom,numfactu,fecfactu,numlinea,numalbar,"
'    sql = sql & "numlinealbar,cantreal,cantfact,precibru,precinet,dtocom1,dtocom2,imporbru,"
'    sql = sql & "impornet,codigiva,unidades) values "
'
'    vWhere2 = "codtipom = " & DBSet(codTipoM, "T") & " and numfactu = " & DBSet(NumFactu, "N") & " and fecfactu = " & DBSet(FecFactu, "F")
'
'    Select Case vCliente.TipoIva
'        Case 0  'normal
''            TipoIva = vParamAplic.CodIvaNormal
''++monica:27/05/08 el tipo de iva lo sacaré de la variedad
'            TipoIva = 0
'        Case 1  'exento
'            TipoIva = vParamAplic.CodIvaExento
'        Case 2  'recargo de equivalencia
'            TipoIva = vParamAplic.CodIvaRecargo
'    End Select
'
'    vSql1 = Replace(vSql, "albaran", "albaran_variedad")
'
'    Sql2 = "select albaran_variedad.numalbar, albaran_variedad.numlinea, albaran_variedad.codforfait,"
'    Sql2 = Sql2 & " albaran_variedad.totpalet, albaran_variedad.numcajas, albaran_variedad.pesobrut,"
'    Sql2 = Sql2 & " albaran_variedad.pesoneto, albaran_variedad.preciopro, albaran_variedad.preciodef,"
'    Sql2 = Sql2 & " albaran_variedad.impcomis, forfaits.kiloscaj, albaran_variedad.unidades, "
'    Sql2 = Sql2 & " albaran_variedad.codvarie, forfaits.facturar "
'    Sql2 = Sql2 & " from albaran_variedad, forfaits where " & vSql1
'    Sql2 = Sql2 & " and albaran_variedad.codforfait = forfaits.codforfait "
'    Sql2 = Sql2 & " order by albaran_variedad.numalbar, albaran_variedad.numlinea"
'
'    Set Rs = New ADODB.Recordset
'    Rs.Open Sql2, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
'
'    cad = "(" & DBSet(codTipoM, "T") & "," & DBSet(NumFactu, "N") & "," & DBSet(FecFactu, "F") & ","
'    cad2 = ""
'
'    NumF = 1
'
'    While Not Rs.EOF
'        If vCliente.TipoIva = 0 Then
'            TipoIva = DevuelveDesdeBDNew(cAgro, "variedades", "codigiva", "codvarie", Rs.Fields(12).Value, "N")
'        End If
'        cad2 = cad2 & cad & DBSet(NumF, "N") & "," & DBSet(Rs.Fields(0).Value, "N") & "," & DBSet(Rs.Fields(1).Value, "N") & ","
'        cad2 = cad2 & DBSet(Rs.Fields(6).Value, "N") & "," 'cantreal <- pesoneto
'        cad2 = cad2 & DBSet(DBLet(Rs.Fields(4).Value, "N") * DBLet(Rs.Fields(10).Value, "N"), "N") & "," 'cantfact <- numcajas * kiloscaj
'        cad2 = cad2 & DBSet(Rs.Fields(7).Value, "N") & "," 'preciobru <- preciopro
'        cad2 = cad2 & "0," ' precinet
'        cad2 = cad2 & DBSet(vCliente.Dto1, "N") & "," ' dtocom1
'        cad2 = cad2 & DBSet(vCliente.Dto2, "N") & "," ' dtocom2
'
'        '++monica:27/05/08 dependiendo de si la confeccion es por unidades(0) o por kilos(1)
'        If DBLet(Rs.Fields(13).Value, "N") = 0 Then 'unidades
'            cad2 = cad2 & DBSet(Round2(Rs.Fields(11).Value * DBLet(Rs.Fields(7).Value, "N"), 2), "N") & "," ' imporbru
'        Else   'kilos
'            cad2 = cad2 & DBSet(Round2(Rs.Fields(6).Value * DBLet(Rs.Fields(7).Value, "N"), 2), "N") & "," ' imporbru
'        End If
'
'        cad2 = cad2 & "0," ' impornet
'        cad2 = cad2 & DBSet(TipoIva, "N") & "," ' codigiva
'        cad2 = cad2 & DBSet(Rs.Fields(11).Value, "N") & "),"
'
'        NumF = NumF + 1
'
'        Rs.MoveNext
'    Wend
'
'    Set Rs = Nothing
'
'    ' quitamos la ultima ,
'    cad2 = Mid(cad2, 1, Len(cad2) - 1)
'    cad2 = sql & cad2
'
'    conn.Execute cad2
'
'    InsertarLineasFactu = True
'
'EInsertarLin:
'    If Err.Number <> 0 Then
'        InsertarLineasFactu = False
'        MuestraError Err.Number, "Insertar lineas factura.", Err.Description
'    Else
'        InsertarLineasFactu = True
'    End If
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



Private Function CalcularImporteDto(cantidad As String, Precio As String, ImpDto As String, Insertado As Boolean) As String
'Insertado: indica si ya hemos insertado el registro o no
'Calcula el Importe de una linea de Oferta, Pedido, Albaran, ...
'Importe=cantidad * precio - (descuentos)
Dim vCant As Currency
Dim vImp As Currency
Dim vDto As Currency
Dim vPre As Currency
Dim Rs As ADODB.Recordset
Dim Sql As String
Dim SumaBruto As Currency

On Error Resume Next


    Sql = "select sum(imporbru) from facturas_variedad where codtipom = " & DBSet(CodTipom, "T")
    Sql = Sql & " and numfactu = " & DBSet(numfactu, "N") & " and fecfactu = " & DBSet(fecfactu, "F")
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
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




Public Function PasarPartesAFactura2(TipoAlb As String, cadSQL As String, textoCSB As String, Forpa As String, Optional ErroresAux As String, Optional EstaRecu As Boolean) As Boolean
'IN ->  cadSQL: cadena para seleccion de los Albaranes que vamos a Facturar
'       TextosCSB:  Para la tesoreria
'Desde Albaranes Genera las Facturas correspondientes
Dim b As Boolean
Dim MenError As String

    On Error GoTo ETraspasoAlbFac

    PasarPartesAFactura2 = False
    b = False
    ErroresAux = ""
    '#Laura 14/11/2006 recuperar facturas de Alzira
    'si esta recuperando facturas
    EstaRecuFac = EstaRecu
    
    Select Case TipoAlb
        Case "PAR" 'PAR: Partes de ADV de socios
            CodTipom = "FAP" ' FActuras de Partes de adv
    End Select
      
     
    'Aqui empieza transaccion
    conn.BeginTrans
   
    'Insertar la Factura (advfacturas, advfacturas_lineas)
    MenError = "Error al Insertar en tablas de factura."
    b = InsertarFacturaADV(cadSQL)
    

    'Eliminar los Partes incluidos en la Factura
    If b Then
        MenError = "Error al eliminar los Partes."
        b = EliminarPartes(cadSQL)
    End If
    
ETraspasoAlbFac:
    If Err.Number <> 0 Then
        ErroresAux = Err.Description
        Err.Clear
        b = False
    End If
  
    If b Then
        conn.CommitTrans
        PasarPartesAFactura2 = True
    Else
        conn.RollbackTrans
        PasarPartesAFactura2 = False
        ErroresAux = MenError & vbCrLf & ErroresAux & vbCrLf
    End If
    espera 0.4
End Function




Private Function InsertarFacturaADV(cadSQL As String) As Boolean
'Insertamos en las tablas de factura de socios: advfacturas, advfacturas_lineas
Dim b As Boolean
Dim vTipoMov As CTiposMov
Dim Existe As Boolean
Dim devuelve As String
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim vSocio As cSocio
Dim TipoMovimiento As String

    InsertarFacturaADV = False

    'cogemos los datos del albaran
    Sql = "select codsocio, fechapar, codcampo, codtrata, observac from advpartes where " & cadSQL
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText

    Set vSocio = New cSocio

    'si se ha modificado el cliente volver a cargar los datos
    If vSocio.Existe(Rs.Fields(0).Value) Then
        If vSocio.LeerDatos(Rs.Fields(0).Value) Then
            Socio = vSocio.Codigo

            If vSocio.EsFactADVInt = 1 Then
                TipoMovimiento = "FIN"
            Else
                TipoMovimiento = "FAP"
            End If
            Set vTipoMov = New CTiposMov
            If vTipoMov.Leer(TipoMovimiento) Then
                numfactu = vTipoMov.ConseguirContador(TipoMovimiento)
                ' si existe la factura incrementamos el contador
                Do
                    devuelve = DevuelveDesdeBDNew(cAgro, "advfacturas", "numfactu", "codtipom", TipoMovimiento, "T", , "numfactu", numfactu, "N", "fecfactu", Rs.Fields(1).Value, "F")
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
    
            'Cabeceras Partes de la Factura
            b = InsertarCabPartesFactu(cadSQL)
            If Not b Then Exit Function
    
            'Insertar lineas de partes de la factura
            b = InsertarLineasFactuADV(cadSQL)
            If Not b Then Exit Function

            'Insertar lineas de trabajador de partes de la factura
            b = InsertarLineasTrabFactuADV(cadSQL)
            If Not b Then Exit Function



            'calcular el total de la factura
            b = CalcularDatosFacturaADV(vSocio)
            If Not b Then Exit Function

            'Incrementar contador del tipo de movimiento
            '-------------------------------------------------------------
'--monica: en cualquier caso se incrementa contador
'            If vCliente.TipoFactu = 0 Then
                vTipoMov.IncrementarContador (CodTipom)
                Set vTipoMov = Nothing
'            End If

            InsertarFacturaADV = True

        End If
    End If
    Set vSocio = Nothing

End Function


Private Function EliminarPartes(cadSQL As String) As Boolean
'Eliminamos de las tablas de Partes: advpartes, advpartes_lineas
Dim Sql As String

    On Error GoTo EEliminar

    EliminarPartes = False
    
    'ELiminar lineas partes
    Sql = "DELETE FROM advpartes_lineas "
    Sql = Sql & " WHERE " & Replace(cadSQL, "advpartes", "advpartes_lineas")
    conn.Execute Sql
    
    espera 0.1
    
    'ELiminar lineas trabajadores
    Sql = "DELETE FROM advpartes_trabajador "
    Sql = Sql & " WHERE " & Replace(cadSQL, "advpartes", "advpartes_trabajador")
    conn.Execute Sql
    
    espera 0.1
    
    
    
    'Eliminar Cabeceras Partes
    Sql = "DELETE FROM advpartes "
    Sql = Sql & " WHERE " & Replace(cadSQL, "advpartes_lineas", "advpartes")
    conn.Execute Sql
        
    EliminarPartes = True

EEliminar:
    If Err.Number <> 0 Then
        EliminarPartes = False
    Else
        EliminarPartes = True
    End If
End Function


Private Function InsertarLineasFactuADV(cadSQL As String) As Boolean
'Insertar las lineas de los Albaranes de envases de la factura en la tabla: slifac
Dim Sql As String
Dim Sql2 As String
Dim Cad As String
Dim Cad2 As String
Dim NumF As Long
Dim Rs As ADODB.Recordset

    On Error GoTo EInsertarLin

    Sql = "INSERT INTO advfacturas_lineas (codtipom,numfactu,fecfactu,codtipop,numparte,numlinea, codalmac,codartic,cantidad,preciove,dosishab,importel,ampliaci,codigiva) values "
    
     'pasamos lineas de partes a lineas factura
    Sql2 = " SELECT numparte,numlinea,codalmac,codartic,cantidad,preciove,dosishab,importel, ampliaci, codigiva "
    Sql2 = Sql2 & " FROM advpartes_lineas "
    Sql2 = Sql2 & " WHERE " & Replace(cadSQL, "advpartes", "advpartes_lineas")
    

    Set Rs = New ADODB.Recordset
    Rs.Open Sql2, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    Cad = "(" & DBSet(CodTipom, "T") & "," & DBSet(numfactu, "N") & "," & DBSet(fecfactu, "F") & ","
    Cad2 = ""
    
    NumF = 1
    
    While Not Rs.EOF
        Cad2 = Cad2 & Cad & "'PAR'," & DBSet(Rs.Fields(0).Value, "N") & "," & DBSet(Rs.Fields(1).Value, "N") & ","
        Cad2 = Cad2 & DBSet(Rs.Fields(2).Value, "N") & "," & DBSet(Rs.Fields(3).Value, "T") & ","
        Cad2 = Cad2 & DBSet(Rs.Fields(4).Value, "N") & "," 'cantidad
        Cad2 = Cad2 & DBSet(Rs.Fields(5).Value, "N") & "," 'precio
        Cad2 = Cad2 & DBSet(Rs.Fields(6).Value, "N") & "," 'dosishab
        Cad2 = Cad2 & DBSet(Rs.Fields(7).Value, "N") & "," 'importe
        Cad2 = Cad2 & DBSet(Rs.Fields(8).Value, "T", "S") & "," 'ampliacion
        Cad2 = Cad2 & DBSet(Rs.Fields(9).Value, "N") & ")," 'codigiva
        
        NumF = NumF + 1
        
        Rs.MoveNext
    Wend
    
    ' quitamos la ultima coma
    If Cad2 <> "" Then
        Cad2 = Mid(Cad2, 1, Len(Cad2) - 1)
    
        conn.Execute Sql & Cad2
    End If

    InsertarLineasFactuADV = True
    
EInsertarLin:
    If Err.Number <> 0 Then
        InsertarLineasFactuADV = False
        MuestraError Err.Number, "Insertar lineas factura partes.", Err.Description
    Else
        InsertarLineasFactuADV = True
    End If
End Function

Private Function InsertarCabPartesFactu(cadSQL As String) As Boolean
'Insertar las cabeceras de los Partes de la factura ADV en la tabla: advfacturas_partes
Dim Sql As String
Dim devuelve As String
Dim MantenimientoObservacion As String

    On Error GoTo EInsertarAlb

    devuelve = ""
    cadSQL = Replace(cadSQL, "advpartes_lineas", "advpartes")
    
    Sql = "INSERT INTO advfacturas_partes (codtipom,numfactu,fecfactu,codtipop,numparte,fechapar,codcampo, codtrata,observac, codcuadrilla, nrohombres, nrohoras, litrospre, litrosrea)"
    
    Sql = Sql & " SELECT '" & CodTipom & "' as codtipom," & DBSet(numfactu, "N") & " as numfactu," & DBSet(fecfactu, "F") & " as fecfactu," & "'PAR', numparte,fechapar,codcampo,codtrata,observac,codcuadrilla,nrohombres,nrohoras, litrospre, litrosrea "
    Sql = Sql & " FROM advpartes WHERE " & cadSQL
    
    conn.Execute Sql

    InsertarCabPartesFactu = True
    
EInsertarAlb:
    If Err.Number <> 0 Then
        InsertarCabPartesFactu = False
        MuestraError Err.Number, "Insertar cabecera partes factura.", Err.Description
    Else
        InsertarCabPartesFactu = True
    End If
End Function

Private Function InsertarLineasTrabFactuADV(cadSQL As String) As Boolean
'Insertar las lineas de los Albaranes de envases de la factura en la tabla: slifac
Dim Sql As String
Dim Sql2 As String
Dim Cad As String
Dim Cad2 As String
Dim NumF As Long
Dim Rs As ADODB.Recordset

    On Error GoTo EInsertarLin

    Sql = "INSERT INTO advfacturas_trabajador (codtipom,numfactu,fecfactu,codtipop,numparte,numlinea,codtraba,horas,precio,importel) values "
    
     'pasamos lineas de partes a lineas factura
    Sql2 = " SELECT numparte,numlinea,codtraba,horas,precio,importel "
    Sql2 = Sql2 & " FROM advpartes_trabajador "
    Sql2 = Sql2 & " WHERE " & Replace(cadSQL, "advpartes", "advpartes_trabajador")
    

    Set Rs = New ADODB.Recordset
    Rs.Open Sql2, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    Cad = "(" & DBSet(CodTipom, "T") & "," & DBSet(numfactu, "N") & "," & DBSet(fecfactu, "F") & ","
    Cad2 = ""
    
    NumF = 1
    
    While Not Rs.EOF
        Cad2 = Cad2 & Cad & "'PAR'," & DBSet(Rs.Fields(0).Value, "N") & "," & DBSet(Rs.Fields(1).Value, "N") & ","
        Cad2 = Cad2 & DBSet(Rs.Fields(2).Value, "N") & "," 'trabajador
        Cad2 = Cad2 & DBSet(Rs.Fields(3).Value, "N") & "," 'horas
        Cad2 = Cad2 & DBSet(Rs.Fields(4).Value, "N") & "," 'precio
        Cad2 = Cad2 & DBSet(Rs.Fields(5).Value, "N") & ")," 'importel
        
        NumF = NumF + 1
        
        Rs.MoveNext
    Wend
    
    ' quitamos la ultima coma
    If Cad2 <> "" Then
        Cad2 = Mid(Cad2, 1, Len(Cad2) - 1)
    
        conn.Execute Sql & Cad2
    End If

    InsertarLineasTrabFactuADV = True
    
EInsertarLin:
    If Err.Number <> 0 Then
        InsertarLineasTrabFactuADV = False
        MuestraError Err.Number, "Insertar lineas trabajador factura partes.", Err.Description
    Else
        InsertarLineasTrabFactuADV = True
    End If
End Function

