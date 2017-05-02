VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFacturaTer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'COMPRAS
'FACTURA: rcafter (cabecera de facturas de socios terceros)
'-------------------------------------------------------------------------

'Variables locales que contienen valores de propiedad
Private mTercero As String 'Cod. socio tercero
Private mNumFactu As String 'Nº Factura
Private mFecFactu As String 'FEcha Factura
Private mFecRecep As String 'FEcha Recepcion

Private mCodmacta As String 'cuenta de la agencia de transporte para insertar in spagop

Private mNomTerce As String 'nombre del Socio
Private mDomTerce As String 'domicilio
Private mCodPobla As String 'cpostal
Private mPobTerce As String 'poblacion
Private mProTerce As String 'provincia
Private mNifTerce As String 'NIF agencia de transporte
Private mTelTerce As String 'telefono

Private mCodForpa As Integer 'Cod. Forma de pago
Private mCodTraba As String 'Cod. trabajador factura

Private mBrutofac As Currency 'total importe bruto de las lineas de la factura(sin dtos ni IVA)
Private mDtoPpago As Currency 'descuento pronto pago
Private mDtoGnral As Currency 'descuento general
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

Private mPorcRet As Currency    '% de retencion
Private mImpReten As Currency   'Importe al aplicar a la base imponible la retencion
Private mBaseReten As Currency  ' base de la retencion

Private mPorcrec1 As Currency   '% de recargo 1
Private mPorcrec2 As Currency   '% de recargo 2
Private mPorcrec3 As Currency   '% de recargo 3

Private mImpoiva1 As Currency   'Importe de iva 1
Private mImpoiva2 As Currency   'Importe de iva 2
Private mImpoiva3 As Currency   'Importe de iva 3
Private mImporec1 As Currency   'Importe de recargo 1
Private mImporec2 As Currency   'Importe de recargo 2
Private mImporec3 As Currency   'Importe de recargo 3

Private mTotalfac As Currency  'total factura (con dtos y con IVA)

Private mBanPr As String 'Banco propio que ponemos como cuenta prevista de pago
Private mCtaPrev As String 'Cuenta prevista de pago

Private mIntracomunitario As Byte   'Socio tercero intracomunitario
Private mEsAnticipo As Byte   'Es Anticipo
Private mAnticipos As String  'anticipos

Private mVariedad As String

'DAVID###
Private mCCC_Iban As String
Private mCCC_Entidad As String
Private mCCC_Oficina As String
Private mCCC_CC As String
Private mCCC_CTa As String

'[Monica]
Private mPorcCorredor As Currency   '% corredor


Private mConcepCargo As Integer
Private mImpCargo As Currency

Dim vvIban As String



'------------------------------------------------
'Propiedades del modulo CFacturaCom
'------------------------------------------------

'**** Socio Tercero  de la Factura

Public Property Let Tercero(ByVal vData As String)
     mTercero = vData
End Property

Public Property Get Tercero() As String
     Tercero = mTercero
End Property


'**** Nº de la Factura

Public Property Let numfactu(ByVal vData As String)
     mNumFactu = vData
End Property

Public Property Get numfactu() As String
     numfactu = mNumFactu
End Property


'**** Fecha de la Factura

Public Property Let fecfactu(ByVal vData As String)
     mFecFactu = vData
End Property

Public Property Get fecfactu() As String
     fecfactu = mFecFactu
End Property


'**** Fecha Recepcion de la Factura

Public Property Let FecRecep(ByVal vData As String)
     mFecRecep = vData
End Property

Public Property Get FecRecep() As String
     FecRecep = mFecRecep
End Property


'**** cta contable (codmacta) del socio tercero de la Factura

Public Property Let CtaTerce(ByVal vData As String)
     mCodmacta = vData
End Property

Public Property Get CtaTerce() As String
     CtaTerce = mCodmacta
End Property


'**** Nombre del Socio de la Factura

Public Property Let NombreTerce(ByVal vData As String)
     mNomTerce = vData
End Property

Public Property Get NombreTerce() As String
     NombreTerce = mNomTerce
End Property

'**** Domicilio del Socio Tercero de la Factura

Public Property Let DomicilioTerce(ByVal vData As String)
     mDomTerce = vData
End Property

Public Property Get DomicilioTerce() As String
     DomicilioTerce = mDomTerce
End Property


'**** CPostal del Socio de la Factura

Public Property Let CPostalTerce(ByVal vData As String)
     mCodPobla = vData
End Property

Public Property Get CPostalTerce() As String
     CPostalTerce = mCodPobla
End Property


'**** Poblacion del Socio de la Factura

Public Property Let PoblacionTerce(ByVal vData As String)
     mPobTerce = vData
End Property

Public Property Get PoblacionTerce() As String
     PoblacionTerce = mPobTerce
End Property

'**** Provincia del Socio de la Factura

Public Property Let ProvinciaTerce(ByVal vData As String)
     mProTerce = vData
End Property

Public Property Get ProvinciaTerce() As String
     ProvinciaTerce = mProTerce
End Property


'**** NIF del Socio de la Factura

Public Property Let NIFTerce(ByVal vData As String)
     mNifTerce = vData
End Property

Public Property Get NIFTerce() As String
     NIFTerce = mNifTerce
End Property


'**** Telefono del Socio de la Factura

Public Property Let TelefTerce(ByVal vData As String)
     mTelTerce = vData
End Property

Public Property Get TelefTerce() As String
     TelefTerce = mTelTerce
End Property


'**** Trabajador de la Factura

Public Property Let trabajador(ByVal vData As String)
     mCodTraba = vData
End Property

Public Property Get trabajador() As String
     trabajador = mCodTraba
End Property


'**** Forma Pago de la Factura

Public Property Let ForPago(ByVal vData As Integer)
     mCodForpa = vData
End Property

Public Property Get ForPago() As Integer
     ForPago = mCodForpa
End Property


'**** Importe Bruto Factura

Public Property Let BrutoFac(ByVal vData As Currency)
     mBrutofac = vData
End Property

Public Property Get BrutoFac() As Currency
     BrutoFac = mBrutofac
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

'**** Importe pronto pago de la Factura

Public Property Let ImpPPago(ByVal vData As Currency)
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

Public Property Let TotalFac(ByVal vData As Currency)
     mTotalfac = vData
End Property

Public Property Get TotalFac() As Currency
     TotalFac = mTotalfac
End Property


'**** Banco propio de pago de la Factura

Public Property Let BancoPr(ByVal vData As String)
     mBanPr = vData
End Property

Public Property Get BancoPr() As String
     BancoPr = mBanPr
End Property

'**** Cuenta Prevista de pago de la Factura
Public Property Let CuentaPrev(ByVal vData As String)
     mCtaPrev = vData
End Property
Public Property Get CuentaPrev() As String
     CuentaPrev = mCtaPrev
End Property

'**** Cuenta bancaria: IBAN
Public Property Let CCC_Iban(ByVal vData As String)
     mCCC_Iban = vData
End Property
Public Property Get CCC_Iban() As String
     CCC_Iban = mCCC_Iban
End Property



'**** Cuenta banacaria: ENTIDAD
Public Property Let CCC_Entidad(ByVal vData As String)
     mCCC_Entidad = vData
End Property
Public Property Get CCC_Entidad() As String
     CCC_Entidad = mCCC_Entidad
End Property

'**** Cuenta banacaria: OFICINA
Public Property Let CCC_Oficina(ByVal vData As String)
     mCCC_Oficina = vData
End Property
Public Property Get CCC_Oficina() As String
     CCC_Oficina = mCCC_Oficina
End Property

'**** Cuenta banacaria: Codigo control
Public Property Let CCC_CC(ByVal vData As String)
     mCCC_CC = vData
End Property
Public Property Get CCC_CC() As String
     CCC_CC = mCCC_CC
End Property

'Cuenta banacaria.  Cuenta
Public Property Let CCC_CTa(ByVal vData As String)
     mCCC_CTa = vData
End Property
Public Property Get CCC_CTa() As String
     CCC_CTa = mCCC_CTa
End Property


'**** % de retencion de la Factura

Public Property Let PorcReten(ByVal vData As Currency)
     mPorcRet = vData
End Property

Public Property Get PorcReten() As Currency
     PorcReten = mPorcRet
End Property

'**** Importe retencion de la Factura

Public Property Let ImpReten(ByVal vData As Currency)
     mImpReten = vData
End Property

Public Property Get ImpReten() As Currency
     ImpReten = mImpReten
End Property

'**** Base Retencion de la Factura

Public Property Let BaseReten(ByVal vData As Currency)
     mBaseReten = vData
End Property

Public Property Get BaseReten() As Currency
     BaseReten = mBaseReten
End Property


'**** % Intracomunitario

Public Property Let Intracomunitario(ByVal vData As Byte)
     mIntracomunitario = vData
End Property

Public Property Get Intracomunitario() As Byte
     Intracomunitario = mIntracomunitario
End Property


'**** Variedad

Public Property Let Variedad(ByVal vData As String)
     mVariedad = vData
End Property

Public Property Get Variedad() As String
     Variedad = mVariedad
End Property


'**** Porcentaje corredor

Public Property Let PorcCorredor(ByVal vData As Currency)
     mPorcCorredor = vData
End Property

Public Property Get PorcCorredor() As Currency
     PorcCorredor = mPorcCorredor
End Property


'**** % Es Anticipo

Public Property Let EsAnticipo(ByVal vData As Byte)
     mEsAnticipo = vData
End Property

Public Property Get EsAnticipo() As Byte
     EsAnticipo = mEsAnticipo
End Property


'**** Anticipos

Public Property Let Anticipos(ByVal vData As String)
     mAnticipos = vData
End Property

Public Property Get Anticipos() As String
     Anticipos = mAnticipos
End Property



'**** Concepto Cargo

Public Property Let ConcepCargo(ByVal vData As Integer)
     mConcepCargo = vData
End Property

Public Property Get ConcepCargo() As Integer
     ConcepCargo = mConcepCargo
End Property

'**** Importe Cargo

Public Property Let ImpCargo(ByVal vData As Currency)
     mImpCargo = vData
End Property

Public Property Get ImpCargo() As Currency
     ImpCargo = mImpCargo
End Property




'------------------------------------------------
'Procedimientos del modulo CFacturaTer
'------------------------------------------------

Public Function LeerDatos(vTercero As String, vNumfactu As String, vFecfactu As String) As Boolean
'Leer los datos de una factura almacenada en la tabla scafpc
'Lee de la BD: Ariges, Tabla: scafpc
'OUT: True si lee los datos correctamente
Dim Rs As ADODB.Recordset
Dim RS1 As ADODB.Recordset
Dim Mens As String
Dim Sql As String

    On Error GoTo ELeer
    
    LeerDatos = False
    Sql = "SELECT codsocio,numfactu,fecfactu,codforpa,brutofac, totalfac FROM rcafter "
    Sql = Sql & " WHERE codsocio=" & vTercero
    Sql = Sql & " AND numfactu='" & vNumfactu & "'"
    Sql = Sql & " AND fecfactu='" & Format(vFecfactu, FormatoFecha) & "'"

    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    Mens = ""
    If Rs.EOF Then
        LeerDatos = False
    Else
        Mens = "Datos Factura"
        mTercero = Rs!Codsocio
        mNumFactu = CStr(Rs!numfactu)
        mFecFactu = CStr(Rs!fecfactu)
        mCodForpa = CStr(Rs!Codforpa)
        mBrutofac = CStr(Rs!BrutoFac)
        mTotalfac = CStr(Rs!TotalFac)
        'obterner la ctaCble del proveedor
        
        '[Monica]03/07/2013: añado trim(codmacta)
        CtaTerce = Trim(DevuelveDesdeBDNew(cAgro, "rsocios_seccion", "codmacpro", "codsocio", vTercero, "N", , "codsecci", vParamAplic.Seccionhorto, "N"))
        
        
        'si venimos de proceso de insertar tenemos la cuenta prevista de pago que
        'se introdujo en el formulario de pasar albaranes a factura
        'si no venimos de moficicar y la leemos de spagos antes de actualizar
        If CuentaPrev = "" Then
            'leer la cuenta prevista de pagos de la tabla de contabilidad: conta.spagos
            'cuando vamos a borrar una factura porque modificarmos
            Sql = Trim(DevuelveDesdeBDNew(cConta, "spagop", "ctabanc1", "ctaprove", CtaTerce, "T", , "numfactu", vNumfactu, "T", "fecfactu", vFecfactu, "F"))
            If Sql <> "" Then
                CuentaPrev = Sql
                LeerDatos = True
            Else
                LeerDatos = False
                Mens = "La cuenta prevista de pago para la factura no puede ser nula."
            End If
        Else
            LeerDatos = True
        End If
        
'        LeerDatos = True
    End If
    Rs.Close
    Set Rs = Nothing
    Exit Function

ELeer:
    Mens = "Se ha producido un error. " & Mens & vbCrLf
    Mens = Mens & "Número: " & Err.Number & vbCrLf
    Mens = Mens & "Descripción: " & Err.Description
    MsgBox Mens, vbExclamation
    If Not Rs Is Nothing Then Set Rs = Nothing
    LeerDatos = False
End Function



'======================================================================
'GRABAR EN TESORERIA
'======================================================================

Public Function InsertarEnTesoreria(MenError As String) As Boolean
'Guarda datos de Tesoreria en tablas: ariges.svenci y en conta.scobros
Dim B As Boolean
Dim Rs As ADODB.Recordset
Dim rsVenci As ADODB.Recordset
Dim Sql As String
Dim cadValuesAux As String 'para insertar en svenci
Dim CadValues2 As String, CadValuesAux2 As String 'para insertar en conta.scobro
Dim FecVenci As Date, FecVenci1 As Date
Dim ImpVenci As Single
Dim I As Byte



    On Error GoTo EInsertarTesoreria

'    b = False
    InsertarEnTesoreria = False
    CadValues2 = ""

    Dim vSoc As cSocio
    Set vSoc = New cSocio
    
    vSoc.LeerDatos (mTercero)



    'Obtener el Nº de Vencimientos de la forma de pago
    Sql = "SELECT numerove, primerve, restoven FROM forpago WHERE codforpa=" & ForPago
    Set rsVenci = New ADODB.Recordset
    rsVenci.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    If Not rsVenci.EOF Then
        If rsVenci!numerove > 0 Then
            'Obtener los dias de pago de la tabla de parametros: spara1
            Sql = " SELECT  diapago1, diapago2, diapago3,mesnogir "
            Sql = Sql & " FROM sparam "
            Sql = Sql & " WHERE codparam=1"
            Set Rs = New ADODB.Recordset
            Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

            If Not Rs.EOF Then
               'vamos creando la cadena para insertar en spagosp de la CONTA
               '[Monica]03/07/2013: añado trim(codmacta)
               CadValuesAux2 = "("
               If vParamAplic.ContabilidadNueva Then CadValuesAux2 = CadValuesAux2 & DBSet(SerieFraPro, "T") & ","
               CadValuesAux2 = CadValuesAux2 & "'" & Trim(mCodmacta) & "', " & DBSet(numfactu, "T") & ", '" & Format(fecfactu, FormatoFecha) & "', "
                

              'Primer Vencimiento
              '------------------------------------------------------------
              I = 1
              'FECHA VTO
              FecVenci = CDate(fecfactu)
              '=== Modificado: Laura 23/01/2007
'              FecVenci = FecVenci + CByte(DBLet(rsVenci!primerve, "N"))
              FecVenci = DateAdd("d", DBLet(rsVenci!primerve, "N"), FecVenci)
              '==================================
              'comprobar si tiene dias de pago y obtener la fecha del vencimiento correcta
              FecVenci = ComprobarFechaVenci(FecVenci, DBLet(Rs!DiaPago1, "N"), DBLet(Rs!DiaPago2, "N"), DBLet(Rs!DiaPago3, "N"))

              'Comprobar si  tiene mes a no girar
              FecVenci1 = FecVenci
              If DBSet(Rs!mesnogir, "N") <> 0 Then
                  FecVenci1 = ComprobarMesNoGira(FecVenci1, DBSet(Rs!mesnogir, "N"), DBSet(0, "N"), Rs!DiaPago1, Rs!DiaPago2, Rs!DiaPago3)
              End If

              CadValues2 = CadValuesAux2 & I
              CadValues2 = CadValues2 & ", " & ForPago & ", '" & Format(FecVenci1, FormatoFecha) & "', "

              'IMPORTE del Vencimiento
              If rsVenci!numerove = 1 Then
                    ImpVenci = TotalFac
              Else
                    ImpVenci = Round(TotalFac / rsVenci!numerove, 2)
                    'Comprobar que la suma de los vencimientos cuadra con el total de la factura
                    If ImpVenci * rsVenci!numerove <> TotalFac Then
                        ImpVenci = Round(ImpVenci + (TotalFac - ImpVenci * rsVenci!numerove), 2)
                    End If
              End If
              CadValues2 = CadValues2 & DBSet(ImpVenci, "N") & ", " & DBSet(CuentaPrev, "T") & ","


              If Not vParamAplic.ContabilidadNueva Then
                    'David. Para que ponga la cuenta bancaria (SI LA tiene)
                    CadValues2 = CadValues2 & DBSet(mCCC_Entidad, "T", "S") & "," & DBSet(mCCC_Oficina, "T", "S") & ","
                    CadValues2 = CadValues2 & DBSet(mCCC_CC, "T", "S") & "," & DBSet(mCCC_CTa, "T", "S") & ","
              End If

              'David. JUNIO 07.   Los dos textos de grabacion de datos de csb
              Sql = "Factura num.: " & numfactu & "-" & Format(fecfactu, "dd/mm/yyyy")
              CadValues2 = CadValues2 & "'" & DevNombreSQL(Sql) & "',"
              Sql = "Vto a fecha: " & Format(FecVenci1, "dd/mm/yyyy")
              CadValues2 = CadValues2 & "'" & DevNombreSQL(Sql) & "'" ')"
              
              
              If vParamAplic.ContabilidadNueva Then
                  vvIban = MiFormat(mCCC_Iban, "") & MiFormat(mCCC_Entidad, "0000") & MiFormat(mCCC_Oficina, "0000") & MiFormat(mCCC_CC, "00") & MiFormat(mCCC_CTa, "0000000000")
                  
                  CadValues2 = CadValues2 & "," & DBSet(vvIban, "T") & ","
                  'nomprove, domprove, pobprove, cpprove, proprove, nifprove, codpais
                  CadValues2 = CadValues2 & DBSet(vSoc.Nombre, "T") & "," & DBSet(vSoc.Direccion, "T") & "," & DBSet(vSoc.Poblacion, "T") & "," & DBSet(vSoc.CPostal, "T") & ","
                  CadValues2 = CadValues2 & DBSet(vSoc.Provincia, "T") & "," & DBSet(vSoc.nif, "T") & ",'ES'),"
              
              Else
                    '[Monica]22/11/2013: Tema iban
                    If vEmpresa.HayNorma19_34Nueva = 1 Then
                       CadValues2 = CadValues2 & ", " & DBSet(mCCC_Iban, "T", "S") & ") "
                    Else
                       CadValues2 = CadValues2 & ") "
                    End If
              End If

              'Resto Vencimientos
              '--------------------------------------------------------------------
              For I = 2 To rsVenci!numerove
                 'FECHA Resto Vencimientos
                  '==== Modificado: Laura 23/01/2007
                  'FecVenci = FecVenci + DBSet(rsVenci!restoven, "N")
                  FecVenci = DateAdd("d", DBLet(rsVenci!restoven, "N"), FecVenci)
                  '==================================================
                  'comprobar si tiene dias de pago y obtener la fecha del vencimiento correcta
                  FecVenci = ComprobarFechaVenci(FecVenci, DBLet(Rs!DiaPago1, "N"), DBLet(Rs!DiaPago2, "N"), DBLet(Rs!DiaPago3, "N"))

                  'Comprobar si tiene mes a no girar
                  FecVenci1 = FecVenci
                  If DBSet(Rs!mesnogir, "N") <> 0 Then
                        FecVenci1 = ComprobarMesNoGira(FecVenci1, DBSet(Rs!mesnogir, "N"), DBSet(0, "N"), Rs!DiaPago1, Rs!DiaPago2, Rs!DiaPago3)
                  End If

                  CadValues2 = CadValues2 & ", " & CadValuesAux2 & I & ", " & ForPago & ", '" & Format(FecVenci1, FormatoFecha) & "', "

                  'IMPORTE Resto de Vendimientos
                  ImpVenci = Round(TotalFac / rsVenci!numerove, 2)

                  CadValues2 = CadValues2 & DBSet(ImpVenci, "N") & ", '" & CuentaPrev & "',"

                  If Not vParamAplic.ContabilidadNueva Then
                        'David. Para que ponga la cuenta bancaria (SI LA tiene)
                        CadValues2 = CadValues2 & DBSet(mCCC_Entidad, "T", "S") & "," & DBSet(mCCC_Oficina, "T", "S") & ","
                        CadValues2 = CadValues2 & DBSet(mCCC_CC, "T", "S") & "," & DBSet(mCCC_CTa, "T", "S") & ","
                  End If
                  
                  Sql = "Factura num.: " & numfactu & "-" & Format(fecfactu, "dd/mm/yyyy")
                  CadValues2 = CadValues2 & "'" & DevNombreSQL(Sql) & "',"
                  Sql = "Vto a fecha: " & Format(FecVenci1, "dd/mm/yyyy")
                  CadValues2 = CadValues2 & "'" & DevNombreSQL(Sql) & "'" ')"
                  If vParamAplic.ContabilidadNueva Then
                        vvIban = MiFormat(mCCC_Iban, "") & MiFormat(mCCC_Entidad, "0000") & MiFormat(mCCC_Oficina, "0000") & MiFormat(mCCC_CC, "00") & MiFormat(mCCC_CTa, "0000000000")
                        
                        CadValues2 = CadValues2 & "," & DBSet(vvIban, "T") & ","
                        'nomprove, domprove, pobprove, cpprove, proprove, nifprove, codpais
                        CadValues2 = CadValues2 & DBSet(vSoc.Nombre, "T") & "," & DBSet(vSoc.Direccion, "T") & "," & DBSet(vSoc.Poblacion, "T") & "," & DBSet(vSoc.CPostal, "T") & ","
                        CadValues2 = CadValues2 & DBSet(vSoc.Provincia, "T") & "," & DBSet(vSoc.nif, "T") & ",'ES'),"
                  
                  Else
                        '[Monica]22/11/2013: Tema iban
                        If vEmpresa.HayNorma19_34Nueva = 1 Then
                            CadValues2 = CadValues2 & ", " & DBSet(mCCC_Iban, "T", "S") & ") "
                        Else
                            CadValues2 = CadValues2 & ") "
                        End If
                  End If
              Next I
            End If
            Rs.Close
            Set Rs = Nothing
        End If
    End If
    rsVenci.Close
    Set rsVenci = Nothing

    'Grabar tabla spagop de la CONTABILIDAD
    '-------------------------------------------------
    If CadValues2 <> "" Then
        'antes de grabar en la spago comprobar que existe en conta.sforpa la
        'forma de pago de la factura. Sino existe insertarla

        'vemos si existe en la conta
        If vParamAplic.ContabilidadNueva Then
            CadValuesAux2 = DevuelveDesdeBDNew(cConta, "formapago", "codforpa", "codforpa", ForPago, "N")
        Else
            CadValuesAux2 = DevuelveDesdeBDNew(cConta, "sforpa", "codforpa", "codforpa", ForPago, "N")
        End If
        'si no existe la forma de pago en conta, insertamos la de ariges
        If CadValuesAux2 = "" Then
            cadValuesAux = "tipoforp"
            CadValuesAux2 = DevuelveDesdeBDNew(cAgro, "forpago", "nomforpa", "codforpa", ForPago, "N", cadValuesAux)
            'insertamos e sforpa de la CONTA
            If vParamAplic.ContabilidadNueva Then
                Sql = "INSERT INTO formapago(codforpa,nomforpa,tipforpa)"
                Sql = Sql & " VALUES(" & ForPago & ", " & DBSet(CadValuesAux2, "T") & ", " & cadValuesAux & ")"
            Else
                Sql = "INSERT INTO sforpa(codforpa,nomforpa,tipforpa)"
                Sql = Sql & " VALUES(" & ForPago & ", " & DBSet(CadValuesAux2, "T") & ", " & cadValuesAux & ")"
            End If
            ConnConta.Execute Sql
        End If

        'Insertamos en la tabla spagop de la CONTA
        'SQL = "INSERT INTO spagop (ctaprove, numfactu, fecfactu, numorden, codforpa, fecefect, impefect, ctabanc1) "
        'David. Cuenta bancaria y descripcion textos
        If vParamAplic.ContabilidadNueva Then
            Sql = "INSERT INTO pagos (numserie, codmacta, numfactu, fecfactu, numorden, codforpa, fecefect, impefect, ctabanc1,text1csb,text2csb, iban,"
            Sql = Sql & "nomprove, domprove, pobprove, cpprove, proprove, nifprove, codpais)"
        Else
            Sql = "INSERT INTO spagop (ctaprove, numfactu, fecfactu, numorden, codforpa, fecefect, impefect, ctabanc1,entidad,oficina,cc,cuentaba,text1csb,text2csb" ') "
            '[Monica]22/11/2013: Tema iban
            If vEmpresa.HayNorma19_34Nueva = 1 Then
                Sql = Sql & ", iban) "
            Else
                Sql = Sql & ") "
            End If
        End If
        Sql = Sql & " VALUES " & Mid(CadValues2, 1, Len(CadValues2) - 1)
        ConnConta.Execute Sql
    End If

    B = True
    Set vSoc = Nothing

EInsertarTesoreria:
    If Err.Number <> 0 Then
        B = False
        MenError = "Error al insertar en Tesoreria: " & Err.Description
    End If
    InsertarEnTesoreria = B
End Function



Public Function TraspasoAlbaranesAFactura(cadSQL As String) As Boolean
'IN -> cadSQL: cadena para seleccion de los Albaranes que vamos a Facturar
'Desde Albaranes Genera las Facturas correspondientes
Dim B As Boolean
Dim MenError As String
Dim vSeccion As CSeccion

    On Error GoTo ETraspasoAlbFac
    TraspasoAlbaranesAFactura = False
    
    conn.BeginTrans
    
    Set vSeccion = New CSeccion
    If vSeccion.LeerDatos(vParamAplic.Seccionhorto) Then
        If vSeccion.AbrirConta Then
            ConnConta.BeginTrans
        Else
            Exit Function
        End If
    End If
    
    
'    cadSQL = Replace(cadSQL, "slialp", "scaalp")
    
    'Cuenta Prevista de Cobro de las Facturas
    CuentaPrev = DevuelveDesdeBDNew(cAgro, "banpropi", "codmacta", "codbanpr", BancoPr, "N")
    
    
    'Insertar la Factura
    B = InsertarFactura(cadSQL, MenError)

    If B Then
        'Insertar en Tesoreria
        '#### Laura 15/11/2006 Recuperar facturas ALZIRA
        If frmTercRecFact.Check1(0).Value = 1 Then
        '####
            '[Monica]27/05/2013: insertarmos los vencimientos dependiendo del porcentaje de corredor
            If vParamAplic.Cooperativa = 12 Then
                B = InsertarEnTesoreriaMontifrut(MenError)
            Else
                B = InsertarEnTesoreria(MenError)
            End If
        End If
        
        If B Then
            B = MarcarAlbaran(MenError)
        End If
        
'        '#### Laura 15/11/2006 Recuperar facturas ALZIRA
'        If frmTercRecFact.Check1(1).Value = 1 Then
'            If b Then
'                b = Me.MarcarContabilizada(MenError)
'            End If
'        End If
'        '####
        
        
    End If
    If B Then
        conn.CommitTrans
        ConnConta.CommitTrans
        TraspasoAlbaranesAFactura = True
        MsgBox "La Factura de Socio Tercero de los Albaranes seleccionados se generó correctamente.", vbInformation
    Else
        conn.RollbackTrans
        ConnConta.RollbackTrans
        TraspasoAlbaranesAFactura = False
        MsgBox "ATENCIÓN:" & vbCrLf & "La Factura NO se generó correctamente!!!." & vbCrLf & MenError, vbInformation
    End If
    
    vSeccion.CerrarConta
    Set vSeccion = Nothing

    Exit Function


ETraspasoAlbFac:
    If Err.Number <> 0 Then
        conn.RollbackTrans
        ConnConta.RollbackTrans
        TraspasoAlbaranesAFactura = False
        MuestraError Err.Number, "Pasando Albaranes a Factura.(Recepción)", Err.Description
        vSeccion.CerrarConta
        Set vSeccion = Nothing
    End If
End Function







Private Function InsertarFactura(cadSQL As String, cadError As String) As Boolean
'Insertamos en las tablas de factura de tranportista: tcafpc, tcafpa, tlifpc
Dim B As Boolean
    
    InsertarFactura = False
    
    
    'Cabecera Factura
    cadError = ""
    B = InsertarCabeceraFactu(cadError)
    If Not B Then
        Exit Function
    End If
    
    '[Monica]06/05/2013: si no hay cadena de albaranes no los metemos es un anticipo
    If cadSQL <> "" Then
            'Cabeceras Albaranes de la Factura
            B = InsertarCabAlbaranesFactu(cadSQL, cadError)
            If Not B Then
                Exit Function
            
            '[Monica]23/09/2013: descontamos anticipos
            Else
                If Anticipos <> "" Then
                    B = InsertarAnticipos(cadError)
                End If
            End If
    Else
        If B Then B = InsertarAlbaran(cadError)
        
        If B Then
            If Anticipos <> "" Then
                B = InsertarAnticipos(cadError)
            
            End If
'        If b Then b = ActualizarFactura(cadError)
        End If
    End If
    
    InsertarFactura = True
End Function


Private Function InsertarCabeceraFactu(Optional cadErr As String) As Boolean
'Inserta la cabecera de la factura en la tabla: scafpc
Dim Sql As String
Dim vSocio As cSocio
Dim Nulo2 As String
Dim Nulo3 As String
Dim Nulo4 As String
Dim Nulo5 As String
Dim Nulo6 As String

    On Error GoTo EInsertar

    'Obtener los datos del Tercero: nombre, domicilio, NIF,....
    Set vSocio = New cSocio
    If vSocio.LeerDatos(Tercero) Then
        '#### Leer estos datos de la tabla scaalpr y no de sprove
        NombreTerce = vSocio.Nombre
        DomicilioTerce = vSocio.Direccion
        CPostalTerce = vSocio.CPostal
        PoblacionTerce = vSocio.Poblacion
        ProvinciaTerce = vSocio.Provincia
        NIFTerce = vSocio.nif
        TelefTerce = vSocio.Tfno1
'        CtaTerce = vSocio.CtaProv
    End If

    'Insertar en la tabla cabecera de la factura de compras
    Sql = "INSERT INTO rcafter (codsocio,numfactu,fecfactu,fecrecep,nomsocio,domsocio,codpobla,pobsocio,prosocio,nifsocio,telsocio,codforpa,brutofac,dtoppago,dtognral,impppago,impgnral,baseiva1,baseiva2,baseiva3, "
    Sql = Sql & "tipoiva1,tipoiva2,tipoiva3,porciva1,porciva2,porciva3,impoiva1,impoiva2,impoiva3,retfacpr,trefacpr,basereten, totalfac,intconta, intracom, porccorredor, esanticipo, concepcargo, impcargo) "
    Sql = Sql & " VALUES (" & Tercero & "," & DBSet(numfactu, "T") & "," & DBSet(fecfactu, "F") & "," & DBSet(FecRecep, "F") & ","
    Sql = Sql & DBSet(NombreTerce, "T", "N") & "," & DBSet(DomicilioTerce, "T", "N") & "," & DBSet(CPostalTerce, "T", "N") & "," & DBSet(PoblacionTerce, "T", "N") & "," & DBSet(ProvinciaTerce, "T", "N") & ","
    Sql = Sql & DBSet(NIFTerce, "T", "N") & "," & DBSet(TelefTerce, "T", "N") & "," & ForPago & ","
    Sql = Sql & DBSet(BrutoFac, "N", "N") & ","
    Sql = Sql & DBSet(DtoPPago, "N", "N") & "," & DBSet(DtoGnral, "N", "N") & ","
    Sql = Sql & DBSet(ImpPPago, "N", "N") & "," & DBSet(ImpGnral, "N", "N") & ","
    Sql = Sql & DBSet(BaseIVA1, "N", "N") & "," & DBSet(BaseIVA2, "N", "S") & "," & DBSet(BaseIVA3, "N", "S") & ","
    Nulo2 = "N"
    Nulo3 = "N"
    If BaseIVA2 = 0 Then Nulo2 = "S"
    If BaseIVA3 = 0 Then Nulo3 = "S"
    Sql = Sql & DBSet(TipoIVA1, "N", "N") & "," & DBSet(TipoIVA2, "N", Nulo2) & "," & DBSet(TipoIVA3, "N", Nulo3) & ","
    Sql = Sql & DBSet(PorceIVA1, "N", "N") & "," & DBSet(PorceIVA2, "N", Nulo2) & "," & DBSet(PorceIVA3, "N", Nulo3) & ","
    Sql = Sql & DBSet(ImpIVA1, "N", "N") & "," & DBSet(ImpIVA2, "N", Nulo2) & "," & DBSet(ImpIVA3, "N", Nulo3) & ","
    If PorcReten = 0 Then Nulo4 = "S"
    If ImpReten = 0 Then Nulo5 = "S"
    If PorcReten = 0 Then Nulo6 = "S"
    Sql = Sql & DBSet(PorcReten, "N", Nulo4) & ","
    Sql = Sql & DBSet(ImpReten, "N", Nulo5) & ","
    Sql = Sql & DBSet(BaseReten, "N", Nulo6) & ","
    Sql = Sql & DBSet(TotalFac, "N", "N") & ",0,"
    Sql = Sql & DBSet(Intracomunitario, "N") & ","
    '[Monica]27/05/2013: añadimos el porcentaje de corredor por Montifrut
    Sql = Sql & DBSet(PorcCorredor, "N") & ","
    Sql = Sql & DBSet(EsAnticipo, "N") & ","
    Sql = Sql & DBSet(ConcepCargo, "N", "S") & ","
    Sql = Sql & DBSet(ImpCargo, "N", "S") & ")"
    
    conn.Execute Sql
    Set vSocio = Nothing
    InsertarCabeceraFactu = True

EInsertar:
    If Err.Number <> 0 Then
        Set vSocio = Nothing
        InsertarCabeceraFactu = False
        cadErr = "Error al insertar cabecera factura: " & Err.Description
    Else
        InsertarCabeceraFactu = True
    End If
End Function


Private Function ActualizarFactura(Optional cadErr As String) As Boolean
'Inserta la cabecera de la factura en la tabla: scafpc
Dim Sql As String
Dim vSocio As cSocio
Dim Nulo2 As String
Dim Nulo3 As String
Dim Nulo4 As String
Dim Nulo5 As String
Dim Nulo6 As String

    On Error GoTo EInsertar

    'Obtener los datos del Tercero: nombre, domicilio, NIF,....
    Set vSocio = New cSocio
    If vSocio.LeerDatos(Tercero) Then
        '#### Leer estos datos de la tabla scaalpr y no de sprove
        NombreTerce = vSocio.Nombre
        DomicilioTerce = vSocio.Direccion
        CPostalTerce = vSocio.CPostal
        PoblacionTerce = vSocio.Poblacion
        ProvinciaTerce = vSocio.Provincia
        NIFTerce = vSocio.nif
        TelefTerce = vSocio.Tfno1
'        CtaTerce = vSocio.CtaProv
    End If

    'actalizar que es un anticipo
    Sql = "UPDATE rcafter set esanticipo = 1 where codsocio = " & DBSet(Tercero, "N") & " and numfactu = " & DBSet(numfactu, "T") & " and fecfactu = " & DBSet(fecfactu, "F")
    conn.Execute Sql
    
    
    Set vSocio = Nothing
    ActualizarFactura = True

EInsertar:
    If Err.Number <> 0 Then
        Set vSocio = Nothing
        ActualizarFactura = False
        cadErr = "Error al actualizar  factura: " & Err.Description
    Else
        ActualizarFactura = True
    End If
End Function






Private Function InsertarCabAlbaranesFactu(cadSQL As String, Optional cadErr As String) As Boolean
'Insertar las cabeceras de los Albaranes de la factura en la tabla: scafpa
Dim Sql As String

    On Error GoTo EInsertarAlb

    Sql = "INSERT INTO rlifter (codsocio,numfactu,fecfactu,numalbar,fechaalb,codvarie,kilosnet,importel,prestimado, observa1,observa2,observa3,observa4,observa5) "
    If vParamAplic.Cooperativa = 12 Then
        Sql = Sql & " SELECT " & Tercero & " as codsocio," & DBSet(numfactu, "T") & " as numfactu," & DBSet(fecfactu, "F") & " as fecfactu," & " numalbar,fecalbar,codvarie,kilosfactu,impentrada,prestimado," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo
    Else
        Sql = Sql & " SELECT " & Tercero & " as codsocio," & DBSet(numfactu, "T") & " as numfactu," & DBSet(fecfactu, "F") & " as fecfactu," & " numalbar,fecalbar,codvarie,kilosnet,impentrada,prestimado," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo
    End If
    Sql = Sql & " FROM rhisfruta WHERE " & cadSQL
    conn.Execute Sql

    InsertarCabAlbaranesFactu = True
    
EInsertarAlb:
    If Err.Number <> 0 Then
        InsertarCabAlbaranesFactu = False
        cadErr = "Error al insertar cabecera albaranes Factura: " & Err.Description
    Else
        InsertarCabAlbaranesFactu = True
    End If
End Function


Private Function InsertarAnticipos(Optional cadErr As String) As Boolean
'Insertar los anticipos
Dim Sql As String

    On Error GoTo EInsertarAnticipos

    Sql = "INSERT INTO rliantifter (codsocio,numfactu,fecfactu,codsocioanti, numfactuanti, fecfactuanti) "
    Sql = Sql & " SELECT " & Tercero & " as codsocio," & DBSet(numfactu, "T") & " as numfactu," & DBSet(fecfactu, "F") & " as fecfactu,"
    Sql = Sql & Tercero & ",numfactu, fecfactu "
    Sql = Sql & " FROM rcafter WHERE codsocio = " & DBSet(Tercero, "N")
    Sql = Sql & " and (numfactu, fecfactu) in " & Anticipos
    conn.Execute Sql


    '[Monica]25/09/2013: descontamos los anticipos que hemos marcado
    Sql = "update rlifter set descontado = 1 "
    Sql = Sql & " where codsocio = " & DBSet(Tercero, "N")
    Sql = Sql & " and (numfactu, fecfactu) in " & Anticipos
    
    conn.Execute Sql




    InsertarAnticipos = True
    
EInsertarAnticipos:
    If Err.Number <> 0 Then
        InsertarAnticipos = False
        cadErr = "Error al insertar anticipos de Factura: " & Err.Description
    Else
        InsertarAnticipos = True
    End If
End Function






Private Function InsertarAlbaran(cadSQL As String, Optional cadErr As String) As Boolean
'Insertar las cabeceras de los Albaranes de la factura en la tabla: slifter
Dim Sql As String

    On Error GoTo EInsertarAlb

    Sql = "INSERT INTO rlifter (codsocio,numfactu,fecfactu,numalbar,fechaalb,codvarie,kilosnet,importel,prestimado, observa1,observa2,observa3,observa4,observa5) values ("
    Sql = Sql & DBSet(Tercero, "N") & "," & DBSet(numfactu, "T") & "," & DBSet(fecfactu, "F") & ",0," & DBSet(fecfactu, "F") & "," & DBSet(Variedad, "N") & ",0," & DBSet(BaseIVA1, "N") & ",0," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & "," & ValorNulo & ")"
    conn.Execute Sql

    InsertarAlbaran = True
    
EInsertarAlb:
    If Err.Number <> 0 Then
        InsertarAlbaran = False
        cadErr = "Error al insertar cabecera albaranes Factura: " & Err.Description
    Else
        InsertarAlbaran = True
    End If
End Function




'Private Function InsertarLineasFactu(cadSQL As String, cadErr As String) As Boolean
''Insertar las lineas de los Albaranes de la factura en la tabla: slifpc
'Dim SQL As String
'
'    On Error GoTo EInsertarLin
'
'    SQL = "INSERT INTO rlifter (codsocio,numfactu,fecfactu,numalbar,importel) "
'    SQL = SQL & " SELECT " & Tercero & " as codsocio," & DBSet(NumFactu, "T") & " as numfactu," & DBSet(fecfactu, "F") & " as fecfactu," & " rhisfruta.numalbar, impentrada "
'    SQL = SQL & " FROM rhisfruta WHERE " & cadSQL
'
'    conn.Execute SQL
'
'    InsertarLineasFactu = True
'
'EInsertarLin:
'    If Err.Number <> 0 Then
'        InsertarLineasFactu = False
'        cadErr = "Error al insertar líneas de factura: " & Err.Description
'    Else
'        InsertarLineasFactu = True
'    End If
'End Function


'Private Function EliminarAlbaranes(cadSQL As String, cadError As String) As Boolean
''Eliminamos de las tablas de Albaranes: scaalp, slialp
'Dim Sql As String
'
'    On Error GoTo EEliminar
'
'    EliminarAlbaranes = False
'
'    'ELiminar lineas
'    Sql = "DELETE FROM slialp "
'    Sql = Sql & " WHERE " & Replace(cadSQL, "scaalp", "slialp")
'    Conn.Execute Sql
'
'    'Eliminar Cabeceras Albaranes
'    Sql = "DELETE FROM scaalp "
'    Sql = Sql & " WHERE " & cadSQL
'    Conn.Execute Sql
'
'    EliminarAlbaranes = True
'    Exit Function
'
'EEliminar:
'    EliminarAlbaranes = False
'    cadError = "Error al Eliminar albaranes: " & Err.Description
'End Function



Public Function CalcularDatosFactura(cadWHERE As String, Socio As String, Optional Insertar As Boolean) As Boolean
'cadWhere: cad para la where de la SQL que selecciona las lineas del albaran o la factura
'nomTabla: nombre de la tabla de albaranes(scaalp) o de AlbaranesXFactura(scafpa)
'           segun llamemos desde recepcion de facturas o desde Hco de Facturas
Dim Rs As ADODB.Recordset
Dim I As Integer

Dim Sql As String
Dim cadAux As String

'Aqui vamos acumulando los totales
Dim TotBruto As Currency
Dim TotImpIVA As Currency

Dim ImpAux As Currency
Dim impiva As Currency
Dim ImpBImIVA As Currency 'Importe Base imponible a la que hay q aplicar el IVA

Dim vSocio As cSocio
Dim NumConta As String

    CalcularDatosFactura = False
    On Error GoTo ECalcular

    'Obtener los datos del Socio: nombre, domicilio, NIF,....
    Set vSocio = New cSocio
    If vSocio.LeerDatos(Socio) Then
        NumConta = DevuelveDesdeBDNew(cAgro, "rseccion", "empresa_conta", "codsecci", vParamAplic.Seccionhorto, "N")
    
        If AbrirConexionConta2(CInt(NumConta)) Then
            If mIntracomunitario Then
                Sql = "SELECT sum(impentrada) as bruto"
                Sql = Sql & " FROM rhisfruta "
                Sql = Sql & " WHERE " & cadWHERE
            
                TotBruto = DevuelveValor(Sql)
                ImpBImIVA = TotBruto
                
                '---- Laura: 24/10/2006
                ImpAux = CalcularPorcentaje(TotBruto, DtoPPago, 2)
                ImpBImIVA = ImpBImIVA - ImpAux '(bruto - DtoPP)
                
                '---- Laura: 24/10/2006
                ImpAux = mDtoGnral 'CalcularPorcentaje(TotBruto, DtoGnral, 2)
                ImpBImIVA = ImpBImIVA - ImpAux '(bruto - Dtogn)
                
                'Obtener el % de IVA
                cadAux = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", vParamAplic.CodIvaIntra, "N")
                
                '---- Laura: 24/10/2006
                If cadAux = "" Then cadAux = "0"
                impiva = CalcularPorcentaje(ImpBImIVA, CCur(cadAux), 2)
                
                'sumamos todos los IVAS para sumarselo a la base imponible total de la factura
                'los vamos acumulando
                TotImpIVA = TotImpIVA + impiva
                
                TipoIVA1 = vParamAplic.CodIvaIntra
                        
                BaseIVA1 = ImpBImIVA 'BASE IMPONIBLE
                        
                PorceIVA1 = cadAux '% de IVA
                     
                ImpIVA1 = impiva
                
                'TOTALES
                BrutoFac = TotBruto
                
                'Aplicarle el dto ppago
                '---- Laura: 24/10/2006
                ImpPPago = CalcularPorcentaje(TotBruto, DtoPPago, 2)
                '----
                
                'Aplicarle el dto general
                '---- Laura: 24/10/2006
'                ImpGnral = CalcularPorcentaje(TotBruto, DtoGnral, 2)
                '----
                ImpGnral = mDtoGnral
                
                'Base Imponible
                BaseImp = TotBruto - ImpPPago - ImpGnral
                
                
                ' dependiendo del tipoirpf del socio se calcula la retencion
                Select Case vSocio.TipoIRPF
                    Case 0
                        BaseReten = BaseImp + TotImpIVA
                    Case 1
                        BaseReten = BaseImp
                    Case 2
                        BaseReten = 0
                End Select
                
                ' calculo de la retencion
                PorcReten = CCur(ComprobarCero(vParamAplic.PorcreteFacSoc))
                ImpReten = Round2(BaseReten * PorcReten / 100, 2)
            
                'TOTAL de la factura
                TotalFac = BaseImp + TotImpIVA - ImpReten
            
                CalcularDatosFactura = True
            
            Else
                ' no es factura de extranjero cogemos los ivas de las variedades
                
                Sql = "SELECT variedades.codigiva, sum(impentrada) as bruto"
                Sql = Sql & " FROM rhisfruta, variedades "
                Sql = Sql & " WHERE rhisfruta.codvarie = variedades.codvarie and " & cadWHERE
                Sql = Sql & " group by 1"
                Sql = Sql & " order by 1"
                Set Rs = New ADODB.Recordset
                Rs.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
                
                TotBruto = 0
                I = 1
                While Not Rs.EOF
                    'Aqui vamos acumulando la suma del importe bruto de las lineas
                    TotBruto = TotBruto + Rs!Bruto
                    ImpBImIVA = Rs!Bruto
                    
                    '---- Laura: 24/10/2006
                    ImpAux = CalcularPorcentaje(Rs!Bruto, DtoPPago, 2)
                    ImpBImIVA = ImpBImIVA - ImpAux '(bruto - DtoPP)
                    
                    '---- Laura: 24/10/2006
                    ImpAux = mDtoGnral 'CalcularPorcentaje(RS!Bruto, DtoGnral, 2)
                    ImpBImIVA = ImpBImIVA - ImpAux '(bruto - Dtogn)
                    
                    'Obtener el % de IVA
                    cadAux = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", CStr(Rs!Codigiva), "N")
                    
                    '---- Laura: 24/10/2006
                    If cadAux = "" Then cadAux = "0"
                    impiva = CalcularPorcentaje(ImpBImIVA, CCur(cadAux), 2)
                    
                    'sumamos todos los IVAS para sumarselo a la base imponible total de la factura
                    'los vamos acumulando
                    TotImpIVA = TotImpIVA + impiva
                    
                    Select Case I
                        Case 1  'IVA 1
                            TipoIVA1 = Rs!Codigiva
                            
                            BaseIVA1 = ImpBImIVA 'BASE IMPONIBLE
                            
                            PorceIVA1 = cadAux '% de IVA
                            
                            'Importe total con IVA
                            ImpIVA1 = impiva
                            
                        Case 2  'IVA 2
                            TipoIVA2 = Rs!Codigiva
                            
                            BaseIVA2 = ImpBImIVA 'BASE IMPONIBLE
                            
                            PorceIVA2 = cadAux '% de IVA
                            
                            'Importe total con IVA
                            ImpIVA2 = impiva
            
                        Case 3  'IVA 3
                            TipoIVA3 = Rs!Codigiva
                            
                            BaseIVA3 = ImpBImIVA 'BASE IMPONIBLE
                            
                            PorceIVA3 = cadAux '% de IVA
                            
                            'Importe total con IVA
                            ImpIVA3 = impiva
                    End Select
                    I = I + 1
                    Rs.MoveNext
                Wend
                
                Rs.Close
                Set Rs = Nothing
            
                'TOTALES
                BrutoFac = TotBruto
                
                'Aplicarle el dto ppago
                '---- Laura: 24/10/2006
                ImpPPago = CalcularPorcentaje(TotBruto, DtoPPago, 2)
                '----
                
                'Aplicarle el dto general
                '---- Laura: 24/10/2006
'                ImpGnral = CalcularPorcentaje(TotBruto, DtoGnral, 2)
                ImpGnral = mDtoGnral
                '----
                
                'Base Imponible
                BaseImp = TotBruto - ImpPPago - ImpGnral
                
                
                ' dependiendo del tipoirpf del socio se calcula la retencion
                Select Case vSocio.TipoIRPF
                    Case 0
                        BaseReten = BaseImp + TotImpIVA
                    Case 1
                        BaseReten = BaseImp
                    Case 2
                        BaseReten = 0
                End Select
                
                ' calculo de la retencion
                PorcReten = CCur(ComprobarCero(vParamAplic.PorcreteFacSoc))
                ImpReten = Round2(BaseReten * PorcReten / 100, 2)
            
                'TOTAL de la factura
                TotalFac = BaseImp + TotImpIVA - ImpReten
            
                CalcularDatosFactura = True
            End If
        End If
        
        CerrarConexionConta2
    End If
    
ECalcular:
    If Err.Number <> 0 Then
        CalcularDatosFactura = False
    Else
        CalcularDatosFactura = True
    End If
End Function


'#### Laura 15/11/2006 Recuperar facturas ALZIRA
Public Function MarcarContabilizada(Optional cadErr As String) As Boolean
'Poner la factura como contabilizada
Dim Sql As String

    On Error GoTo EActualizar
    MarcarContabilizada = False
    
    Sql = "UPDATE rcafter SET intconta=1 "
    Sql = Sql & " WHERE codsocio= " & Me.Tercero & " AND numfactu=" & DBSet(Me.numfactu, "T") & " AND fecfactu=" & DBSet(Me.fecfactu, "F")
    conn.Execute Sql
    
    MarcarContabilizada = True
    Exit Function
    
EActualizar:
    MarcarContabilizada = False
    cadErr = "Error al marcar como contabilizada:" & Err.Description
End Function

'#### marcar albaranes como cobradas
Public Function MarcarAlbaran(Optional cadErr As String) As Boolean
'Poner la factura como contabilizada
Dim Sql As String

    On Error GoTo EActualizar
    MarcarAlbaran = False

    Sql = "UPDATE rhisfruta SET cobradosn=1 "
    Sql = Sql & " WHERE codsocio= " & Me.Tercero & " AND numalbar in (select numalbar from rlifter where numfactu =" & DBSet(Me.numfactu, "T") & " AND fecfactu=" & DBSet(Me.fecfactu, "F") & ")"
    conn.Execute Sql

    MarcarAlbaran = True
    Exit Function

EActualizar:
    MarcarAlbaran = False
    cadErr = "Error al marcar como contabilizada:" & Err.Description
End Function


'######################################
'########    MONTIFRUT   ##############
'######################################
Public Function InsertarEnTesoreriaMontifrut(MenError As String) As Boolean
'Guarda datos de Tesoreria en tablas: ariges.svenci y en conta.scobros
'Porcen: Porcentaje de corredor
Dim B As Boolean
Dim Rs As ADODB.Recordset
Dim rsVenci As ADODB.Recordset
Dim Sql As String
Dim cadValuesAux As String 'para insertar en svenci
Dim CadValues2 As String, CadValuesAux2 As String 'para insertar en conta.scobro
Dim FecVenci As Date, FecVenci1 As Date
Dim ImpVenci As Single
Dim ImpVenci1 As Single
Dim I As Byte

    On Error GoTo EInsertarTesoreria


    Dim vSoc As cSocio
    Set vSoc = New cSocio
    
    vSoc.LeerDatos (mTercero)


'    b = False
    InsertarEnTesoreriaMontifrut = False
    CadValues2 = ""

    'Obtener el Nº de Vencimientos de la forma de pago
    Sql = "SELECT numerove, primerve, restoven FROM forpago WHERE codforpa=" & ForPago
    Set rsVenci = New ADODB.Recordset
    rsVenci.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    If Not rsVenci.EOF Then
        If rsVenci!numerove > 0 Then
            'Obtener los dias de pago de la tabla de parametros: spara1
            Sql = " SELECT  diapago1, diapago2, diapago3,mesnogir "
            Sql = Sql & " FROM sparam "
            Sql = Sql & " WHERE codparam=1"
            Set Rs = New ADODB.Recordset
            Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

            If Not Rs.EOF Then
               'vamos creando la cadena para insertar en spagosp de la CONTA
               '[Monica]03/07/2013: añado trim(codmacta)
               CadValuesAux2 = "("
               If vParamAplic.ContabilidadNueva Then CadValuesAux2 = CadValuesAux2 & DBSet(SerieFraPro, "T") & ","
               CadValuesAux2 = CadValuesAux2 & "'" & Trim(mCodmacta) & "', " & DBSet(numfactu, "T") & ", '" & Format(fecfactu, FormatoFecha) & "', "

              'Primer Vencimiento
              '------------------------------------------------------------
              I = 1
              'FECHA VTO
              FecVenci = CDate(fecfactu)
              '=== Modificado: Laura 23/01/2007
'              FecVenci = FecVenci + CByte(DBLet(rsVenci!primerve, "N"))
              FecVenci = DateAdd("d", DBLet(rsVenci!primerve, "N"), FecVenci)
              '==================================
              'comprobar si tiene dias de pago y obtener la fecha del vencimiento correcta
              FecVenci = ComprobarFechaVenci(FecVenci, DBLet(Rs!DiaPago1, "N"), DBLet(Rs!DiaPago2, "N"), DBLet(Rs!DiaPago3, "N"))

              'Comprobar si  tiene mes a no girar
              FecVenci1 = FecVenci
              If DBSet(Rs!mesnogir, "N") <> 0 Then
                  FecVenci1 = ComprobarMesNoGira(FecVenci1, DBSet(Rs!mesnogir, "N"), DBSet(0, "N"), Rs!DiaPago1, Rs!DiaPago2, Rs!DiaPago3)
              End If

              CadValues2 = CadValuesAux2 & I
              CadValues2 = CadValues2 & ", " & ForPago & ", '" & Format(FecVenci1, FormatoFecha) & "', "

'[Monica]14/10/2014: calculamos el importe de vencimientos
              'IMPORTE del Vencimiento
              If rsVenci!numerove = 1 Then
                    ImpVenci = TotalFac
              Else
                    ImpVenci = Round(TotalFac / rsVenci!numerove, 2)
                    'Comprobar que la suma de los vencimientos cuadra con el total de la factura
                    If ImpVenci * rsVenci!numerove <> TotalFac Then
                        ImpVenci = Round(ImpVenci + (TotalFac - ImpVenci * rsVenci!numerove), 2)
                    End If
              End If

'[Monica]14/10/2014: antes
'              ImpVenci1 = Round2(TotalFac * ComprobarCero(PorcCorredor) / 100, 2)
'              ImpVenci = TotalFac - ImpVenci1

              CadValues2 = CadValues2 & DBSet(ImpVenci, "N") & ", " & DBSet(CuentaPrev, "T") & ","
              
              If Not vParamAplic.ContabilidadNueva Then
                    'David. Para que ponga la cuenta bancaria (SI LA tiene)
                    CadValues2 = CadValues2 & DBSet(mCCC_Entidad, "T", "S") & "," & DBSet(mCCC_Oficina, "T", "S") & ","
                    CadValues2 = CadValues2 & DBSet(mCCC_CC, "T", "S") & "," & DBSet(mCCC_CTa, "T", "S") & ","
              End If

              'David. JUNIO 07.   Los dos textos de grabacion de datos de csb
              Sql = "Factura num.: " & numfactu & "-" & Format(fecfactu, "dd/mm/yyyy")
              CadValues2 = CadValues2 & "'" & DevNombreSQL(Sql) & "',"
              Sql = "Vto a fecha: " & Format(FecVenci1, "dd/mm/yyyy")
              CadValues2 = CadValues2 & "'" & DevNombreSQL(Sql) & "'" ')"
              
              If vParamAplic.ContabilidadNueva Then
                  vvIban = MiFormat(mCCC_Iban, "") & MiFormat(mCCC_Entidad, "0000") & MiFormat(mCCC_Oficina, "0000") & MiFormat(mCCC_CC, "00") & MiFormat(mCCC_CTa, "0000000000")
                  
                  CadValues2 = CadValues2 & "," & DBSet(vvIban, "T") & ","
                  'nomprove, domprove, pobprove, cpprove, proprove, nifprove, codpais
                  CadValues2 = CadValues2 & DBSet(vSoc.Nombre, "T") & "," & DBSet(vSoc.Direccion, "T") & "," & DBSet(vSoc.Poblacion, "T") & "," & DBSet(vSoc.CPostal, "T") & ","
                  CadValues2 = CadValues2 & DBSet(vSoc.Provincia, "T") & "," & DBSet(vSoc.nif, "T") & ",'ES'),"
              
              
              Else
                    '[Monica]22/11/2013: Tema iban
                    If vEmpresa.HayNorma19_34Nueva = 1 Then
                          CadValues2 = CadValues2 & ", " & DBSet(mCCC_Iban, "T", "S") & ") "
                    Else
                          CadValues2 = CadValues2 & ") "
                    End If
              End If

              'Resto Vencimientos
              '--------------------------------------------------------------------
              For I = 2 To rsVenci!numerove
'              If ImpVenci1 <> 0 Then
              
'[Monica]14/10/2014: antes
'                  i = 2
'[Monica]14/10/2014: ahora
                 'FECHA Resto Vencimientos
                  '==== Modificado: Laura 23/01/2007
                  'FecVenci = FecVenci + DBSet(rsVenci!restoven, "N")
                  FecVenci = DateAdd("d", DBLet(rsVenci!restoven, "N"), FecVenci)
                  '==================================================
                  'comprobar si tiene dias de pago y obtener la fecha del vencimiento correcta
                  FecVenci = ComprobarFechaVenci(FecVenci, DBLet(Rs!DiaPago1, "N"), DBLet(Rs!DiaPago2, "N"), DBLet(Rs!DiaPago3, "N"))

                  'Comprobar si tiene mes a no girar
                  FecVenci1 = FecVenci
                  If DBSet(Rs!mesnogir, "N") <> 0 Then
                        FecVenci1 = ComprobarMesNoGira(FecVenci1, DBSet(Rs!mesnogir, "N"), DBSet(0, "N"), Rs!DiaPago1, Rs!DiaPago2, Rs!DiaPago3)
                  End If

                  CadValues2 = CadValues2 & ", " & CadValuesAux2 & I & ", " & ForPago & ", '" & Format(FecVenci1, FormatoFecha) & "', "

'[Monica]14/10/2014: ahora
                  'IMPORTE Resto de Vendimientos
                  ImpVenci = Round(TotalFac / rsVenci!numerove, 2)
'[Monica]14/10/2014: antes impvenci1
                  CadValues2 = CadValues2 & DBSet(ImpVenci, "N") & ", '" & CuentaPrev & "',"

                  If Not vParamAplic.ContabilidadNueva Then
                        'David. Para que ponga la cuenta bancaria (SI LA tiene)
                        CadValues2 = CadValues2 & DBSet(mCCC_Entidad, "T", "S") & "," & DBSet(mCCC_Oficina, "T", "S") & ","
                        CadValues2 = CadValues2 & DBSet(mCCC_CC, "T", "S") & "," & DBSet(mCCC_CTa, "T", "S") & ","
                  End If
                  
                  Sql = "Factura num.: " & numfactu & "-" & Format(fecfactu, "dd/mm/yyyy")
                  CadValues2 = CadValues2 & "'" & DevNombreSQL(Sql) & "',"
                  Sql = "Vto a fecha: " & Format(FecVenci1, "dd/mm/yyyy")
                  CadValues2 = CadValues2 & "'" & DevNombreSQL(Sql) & "'" ')"
                  
                  If vParamAplic.ContabilidadNueva Then
                        vvIban = MiFormat(mCCC_Iban, "") & MiFormat(mCCC_Entidad, "0000") & MiFormat(mCCC_Oficina, "0000") & MiFormat(mCCC_CC, "00") & MiFormat(mCCC_CTa, "0000000000")
                        
                        CadValues2 = CadValues2 & "," & DBSet(vvIban, "T") & ","
                        'nomprove, domprove, pobprove, cpprove, proprove, nifprove, codpais
                        CadValues2 = CadValues2 & DBSet(vSoc.Nombre, "T") & "," & DBSet(vSoc.Direccion, "T") & "," & DBSet(vSoc.Poblacion, "T") & "," & DBSet(vSoc.CPostal, "T") & ","
                        CadValues2 = CadValues2 & DBSet(vSoc.Provincia, "T") & "," & DBSet(vSoc.nif, "T") & ",'ES'),"
                  
                  Else
                        '[Monica]22/11/2013: Tema iban
                        If vEmpresa.HayNorma19_34Nueva = 1 Then
                            CadValues2 = CadValues2 & ", " & DBSet(mCCC_Iban, "T", "S") & ") "
                        Else
                            CadValues2 = CadValues2 & ") "
                        End If
                  End If
              Next I
'                End If
            End If
            Rs.Close
            Set Rs = Nothing
        End If
    End If
    rsVenci.Close
    Set rsVenci = Nothing

    'Grabar tabla spagop de la CONTABILIDAD
    '-------------------------------------------------
    If CadValues2 <> "" Then
        'antes de grabar en la spago comprobar que existe en conta.sforpa la
        'forma de pago de la factura. Sino existe insertarla

        'vemos si existe en la conta
        If vParamAplic.ContabilidadNueva Then
            CadValuesAux2 = DevuelveDesdeBDNew(cConta, "formapago", "codforpa", "codforpa", ForPago, "N")
        Else
            CadValuesAux2 = DevuelveDesdeBDNew(cConta, "sforpa", "codforpa", "codforpa", ForPago, "N")
        End If
        'si no existe la forma de pago en conta, insertamos la de ariges
        If CadValuesAux2 = "" Then
            cadValuesAux = "tipoforp"
            CadValuesAux2 = DevuelveDesdeBDNew(cAgro, "forpago", "nomforpa", "codforpa", ForPago, "N", cadValuesAux)
            'insertamos e sforpa de la CONTA
            If vParamAplic.ContabilidadNueva Then
                Sql = "INSERT INTO formapago(codforpa,nomforpa,tipforpa)"
            Else
                Sql = "INSERT INTO sforpa(codforpa,nomforpa,tipforpa)"
            End If
            Sql = Sql & " VALUES(" & ForPago & ", " & DBSet(CadValuesAux2, "T") & ", " & cadValuesAux & ")"
            ConnConta.Execute Sql
        End If

        'Insertamos en la tabla spagop de la CONTA
        'SQL = "INSERT INTO spagop (ctaprove, numfactu, fecfactu, numorden, codforpa, fecefect, impefect, ctabanc1) "
        'David. Cuenta bancaria y descripcion textos
        If vParamAplic.ContabilidadNueva Then
            Sql = "INSERT INTO pagos (numserie, codmacta, numfactu, fecfactu, numorden, codforpa, fecefect, impefect, ctabanc1,text1csb,text2csb, iban,"
            Sql = Sql & "nomprove, domprove, pobprove, cpprove, proprove, nifprove, codpais)"
       
        Else
            Sql = "INSERT INTO spagop (ctaprove, numfactu, fecfactu, numorden, codforpa, fecefect, impefect, ctabanc1,entidad,oficina,cc,cuentaba,text1csb,text2csb" ') "
            '[Monica]22/11/2013: Tema iban
            If vEmpresa.HayNorma19_34Nueva = 1 Then
                Sql = Sql & ", iban) "
            Else
                Sql = Sql & ") "
            End If
        End If
        Sql = Sql & " VALUES " & CadValues2
        ConnConta.Execute Sql
    End If

    B = True
    Set vSoc = Nothing
    
EInsertarTesoreria:
    If Err.Number <> 0 Then
        B = False
        MenError = "Error al insertar en Tesoreria: " & Err.Description
    End If
    InsertarEnTesoreriaMontifrut = B
End Function




