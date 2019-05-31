VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSocio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ### [Monica] 04/10/2006
' clase modificada para la tabla de rsocios


Option Explicit

'Socios
'TABLA: rsocios


'ATRIBUTOS
'Variables locales que contienen valores de propiedad
Private mCodsocio As String 'Cod. socio
Private mNomsocio As String 'Nombre socio
Private mDirsocio As String 'Dierccion socio
Private mCodposta As String 'cod. postal socio
Private mPobsocio As String 'poblacion socio
Private mProsocio As String 'Provincia socio
Private mNIFsocio As String 'NIF socio
Private mTfno1 As String 'Telefono1 administracion
Private mTfno2 As String 'Telefono2 administracion
Private mTfno3 As String 'Telefono3 administracion
Private mMovsocio As String 'Movil Socio
Private mMaisocio As String 'Mail administracion

Private mObservac As String 'observaciones del socio



Private mFecAlta As Date 'fecha de alta

'los declaro string porque si no tendria '0000-00-00'
Private mFecBaja As String 'fecha de baja
Private mFecNacim As String 'fecha de nacimiento

'[Monica]22/11/2013: tema iban
Private mIban As String ' iban

Private mCodbanco As String 'cod. banco
Private mCodsucur As String 'cod. sucur
Private mDigContr As String 'digito control
Private mCuentaba As String 'cuenta bancaria

Private mCodSitua As Byte 'codigo de situacion
Private mCodCoope As String 'codigo de cooperativa
Private mTipoIRPF As Byte 'tipo irpf: 0=modulos 1=E.D 2=entidad
Private mTipoProd As Byte 'tipo prod: 0=socio 1=tercero 2=otras OPA
Private mCorreo As Byte 'recibe correo si o no


Private mCodTipomAnt As String 'tipo de movimiento de anticipos
Private mCodTipomLiq As String 'tipo de movimiento de liquidacion
Private mCodTipomAntVC As String 'tipo de movimiento de anticipos venta campo
Private mCodTipomLiqVC As String 'tipo de movimiento de liquidacion venta campo

Private mCodTipomAntAlmz As String 'tipo de movimiento de anticipos almazara
Private mCodTipomLiqAlmz As String 'tipo de movimiento de liquidacion almazara

Private mCodTipomAntBod As String 'tipo de movimiento de anticipos bodega
Private mCodTipomLiqBod As String 'tipo de movimiento de liquidacion bodega

Private mCodTipomFacAlmz As String 'tipo de movimiento de factura de retirada de almazara
Private mCodTipomFacBod As String 'tipo de movimiento de factura de retirada de bodega


Private vSeccion As CSeccion  ' seccion del socio

Private mCodSecci As String
Private mCtaCli As String
Private mCtaPro As String
Private mCodIva As String

Private mEsFactADVInt As Byte 'es factura de adv interna

Private mEmiteFact As Byte

Private mvarNoDevolverApor As Byte  'Sin devolucion de aportaciones

Private mvarHayEmbargo As Byte 'hay embargo

'------------------------------------------------
'Propiedades del modulo Csocio
'------------------------------------------------

'**** Codigo del socio

Public Property Let Codigo(ByVal vData As String)
     mCodsocio = vData
End Property

Public Property Get Codigo() As String
     Codigo = mCodsocio
End Property

'**** Nombre del socio

Public Property Let Nombre(ByVal vData As String)
     mNomsocio = vData
End Property

Public Property Get Nombre() As String
     Nombre = mNomsocio
End Property


'**** Direccion del socio

Public Property Let Direccion(ByVal vData As String)
     mDirsocio = vData
End Property

Public Property Get Direccion() As String
     Direccion = mDirsocio
End Property

'**** CPostal del socio

Public Property Let CPostal(ByVal vData As String)
     mCodposta = vData
End Property

Public Property Get CPostal() As String
     CPostal = mCodposta
End Property

'**** Poblacion del socio

Public Property Let Poblacion(ByVal vData As String)
     mPobsocio = vData
End Property

Public Property Get Poblacion() As String
     Poblacion = mPobsocio
End Property

'**** Provincia del socio

Public Property Let Provincia(ByVal vData As String)
     mProsocio = vData
End Property

Public Property Get Provincia() As String
     Provincia = mProsocio
End Property

'**** NIF del socio

Public Property Let nif(ByVal vData As String)
     mNIFsocio = vData
End Property

Public Property Get nif() As String
     nif = mNIFsocio
End Property

'**** Telefono1 del socio

Public Property Let Tfno1(ByVal vData As String)
     mTfno1 = vData
End Property

Public Property Get Tfno1() As String
     Tfno1 = mTfno1
End Property

'**** Telefono2 del socio

Public Property Let Tfno2(ByVal vData As String)
     mTfno2 = vData
End Property

Public Property Get Tfno2() As String
     Tfno2 = mTfno2
End Property

'**** Telefono3 del socio

Public Property Let Tfno3(ByVal vData As String)
     mTfno3 = vData
End Property

Public Property Get Tfno3() As String
     Tfno3 = mTfno3
End Property

'**** Movil del socio

Public Property Let MovilSocio(ByVal vData As String)
     mMovsocio = vData
End Property

Public Property Get MovilSocio() As String
     MovilSocio = mMovsocio
End Property




'**** e-MAil del socio

Public Property Let EMail(ByVal vData As String)
     mMaisocio = vData
End Property

Public Property Get EMail() As String
     EMail = mMaisocio
End Property

'****  fecha de nacimiento
Public Property Let FechaNac(ByVal vData As String)
     mFecNacim = vData
End Property

Public Property Get FechaNac() As String
     FechaNac = mFecNacim
End Property

'****  fecha de alta
Public Property Let FechaAlta(ByVal vData As Date)
     mFecAlta = vData
End Property

Public Property Get FechaAlta() As Date
     FechaAlta = mFecAlta
End Property

'****  fecha de baja
Public Property Let FechaBaja(ByVal vData As String)
     mFecBaja = vData
End Property

Public Property Get FechaBaja() As String
     FechaBaja = mFecBaja
End Property

'[Monica]22/11/2013: tema Iban
'**** IBAN
Public Property Let Iban(ByVal vData As String)
     mIban = vData
End Property

Public Property Get Iban() As String
     Iban = mIban
End Property

'**** Banco del socio

Public Property Let Banco(ByVal vData As String)
     mCodbanco = vData
End Property

Public Property Get Banco() As String
     Banco = mCodbanco
End Property


'**** Sucursal del socio

Public Property Let Sucursal(ByVal vData As String)
     mCodsucur = vData
End Property

Public Property Get Sucursal() As String
     Sucursal = mCodsucur
End Property


'**** Digito Control de la cuenta del socio

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


'**** situacion socio

Public Property Let Situacion(ByVal vData As Byte)
     mCodSitua = vData
End Property

Public Property Get Situacion() As Byte
     Situacion = mCodSitua
End Property


'**** Tipo de IRPF

Public Property Let TipoIRPF(ByVal vData As Byte)
     mTipoIRPF = vData
End Property

Public Property Get TipoIRPF() As Byte
     TipoIRPF = mTipoIRPF
End Property


'**** Tipo de Produccion

Public Property Let TipoProd(ByVal vData As Byte)
     mTipoProd = vData
End Property

Public Property Get TipoProd() As Byte
     TipoProd = mTipoProd
End Property


'**** Correo

Public Property Let Correo(ByVal vData As Byte)
     mCorreo = vData
End Property

Public Property Get Correo() As Byte
     Correo = mCorreo
End Property

'**** Cooperativa del socio

Public Property Let Cooperativa(ByVal vData As String)
     mCodCoope = vData
End Property

Public Property Get Cooperativa() As String
     Cooperativa = mCodCoope
End Property

'**** Tipo de movimiento de anticipos

Public Property Let CodTipomAnt(ByVal vData As String)
     mCodTipomAnt = vData
End Property

Public Property Get CodTipomAnt() As String
     CodTipomAnt = mCodTipomAnt
End Property

'**** Tipo de movimiento de liquidacion

Public Property Let CodTipomLiq(ByVal vData As String)
     mCodTipomLiq = vData
End Property

Public Property Get CodTipomLiq() As String
     CodTipomLiq = mCodTipomLiq
End Property

'**** Tipo de movimiento de anticipos de venta campo

Public Property Let CodTipomAntVC(ByVal vData As String)
     mCodTipomAntVC = vData
End Property

Public Property Get CodTipomAntVC() As String
     CodTipomAntVC = mCodTipomAntVC
End Property

'**** Tipo de movimiento de liquidacion de venta campo

Public Property Let CodTipomLiqVC(ByVal vData As String)
     mCodTipomLiqVC = vData
End Property

Public Property Get CodTipomLiqVC() As String
     CodTipomLiqVC = mCodTipomLiqVC
End Property

'**** Tipo de movimiento de anticipos almazara

Public Property Let CodTipomAntAlmz(ByVal vData As String)
     mCodTipomAntAlmz = vData
End Property

Public Property Get CodTipomAntAlmz() As String
     CodTipomAntAlmz = mCodTipomAntAlmz
End Property

'**** Tipo de movimiento de liquidacion de almazara

Public Property Let CodTipomLiqAlmz(ByVal vData As String)
     mCodTipomLiqAlmz = vData
End Property

Public Property Get CodTipomLiqAlmz() As String
     CodTipomLiqAlmz = mCodTipomLiqAlmz
End Property


'**** Tipo de movimiento de anticipos bodega

Public Property Let CodTipomAntBod(ByVal vData As String)
     mCodTipomAntBod = vData
End Property

Public Property Get CodTipomAntBod() As String
     CodTipomAntBod = mCodTipomAntBod
End Property

'**** Tipo de movimiento de liquidacion de bodega

Public Property Let CodTipomLiqBod(ByVal vData As String)
     mCodTipomLiqBod = vData
End Property

Public Property Get CodTipomLiqBod() As String
     CodTipomLiqBod = mCodTipomLiqBod
End Property


'**** Tipo de movimiento de factura de retirada de almazara

Public Property Let CodTipomFacAlmz(ByVal vData As String)
     mCodTipomFacAlmz = vData
End Property

Public Property Get CodTipomFacAlmz() As String
     CodTipomFacAlmz = mCodTipomFacAlmz
End Property


'**** Tipo de movimiento de factura de retirada de bodega

Public Property Let CodTipomFacBod(ByVal vData As String)
     mCodTipomFacBod = vData
End Property

Public Property Get CodTipomFacBod() As String
     CodTipomFacBod = mCodTipomFacBod
End Property


'**** Seccion del socio

Public Property Let Seccion(ByVal vData As String)
     mCodSecci = vData
End Property

Public Property Get Seccion() As String
     Seccion = mCodSecci
End Property

'**** Cuenta Contable Cliente

Public Property Let CtaClien(ByVal vData As String)
     mCtaCli = vData
End Property

Public Property Get CtaClien() As String
     CtaClien = mCtaCli
End Property

'**** Cuenta Contable Proveedor

Public Property Let CtaProv(ByVal vData As String)
     mCtaPro = vData
End Property

Public Property Get CtaProv() As String
     CtaProv = mCtaPro
End Property


'**** Código de IVA

Public Property Let CodIva(ByVal vData As String)
     mCodIva = vData
End Property

Public Property Get CodIva() As String
     CodIva = mCodIva
End Property


'**** Observaciones del socio

Public Property Let Observaciones(ByVal vData As String)
     mObservac = vData
End Property

Public Property Get Observaciones() As String
     Observaciones = mObservac
End Property


'**** Factura de ADV interna

Public Property Let EsFactADVInt(ByVal vData As Byte)
     mEsFactADVInt = vData
End Property

Public Property Get EsFactADVInt() As Byte
     EsFactADVInt = mEsFactADVInt
End Property


'**** Emite Factura

Public Property Let EmiteFact(ByVal vData As Byte)
     mEmiteFact = vData
End Property

Public Property Get EmiteFact() As Byte
     EmiteFact = mEmiteFact
End Property


'**** Hay embargo

Public Property Let HayEmbargo(ByVal vData As Byte)
     mvarHayEmbargo = vData
End Property

Public Property Get HayEmbargo() As Byte
     HayEmbargo = mvarHayEmbargo
End Property


'**** No devolver aportaciones
Public Property Let NoDevolverApor(ByVal vData As Byte)
     mvarNoDevolverApor = vData
End Property

Public Property Get NoDevolverApor() As Byte
     NoDevolverApor = mvarNoDevolverApor
End Property


'------------------------------------------------
'Procedimientos del modulo Csocio
'------------------------------------------------

Public Function Existe(vCodsocio As String) As Boolean
'Comprueba si existe el socio en la BD
Dim devuelve As String

    On Error GoTo EExiste
    
    devuelve = DevuelveDesdeBD("codsocio", "rsocios", "codsocio", vCodsocio, "N")
    If devuelve = "" Then
        Existe = False
        devuelve = "No existe el socio " & vCodsocio & ". "
        MsgBox devuelve, vbExclamation
    Else
        Existe = True
    End If
    
EExiste:
    If Err.Number <> 0 Then Existe = False
End Function



Public Function LeerDatos(vCodsocio As String) As Boolean
'Leer los datos de un socio dado
'Lee de la BD: Ariagrorec, Tabla: rsocio
'OUT: True si lee los datos correctamente
Dim Rs As ADODB.Recordset
Dim Sql As String
Dim b As Boolean

    On Error GoTo ELeer
    
    LeerDatos = False
    
    Sql = "SELECT codsocio, nifsocio, nomsocio, dirsocio, pobsocio, prosocio, codpostal, "
    Sql = Sql & "fechanac, telsoci1, telsoci2, telsoci3, movsocio, maisocio, codbanco, codsucur, "
    Sql = Sql & "digcontr , cuentaba, fechaalta, fechabaja, correo, tipoIRPF, tipoprod, codsitua, codcoope, "
    Sql = Sql & "observaciones, esfactadvinterna, emitefact, iban, nodevolverapor, hayembargo "
    Sql = Sql & " FROM rsocios "
    Sql = Sql & " WHERE codsocio=" & vCodsocio
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText

    If Rs.EOF Then
        LeerDatos = False
    Else
        
        mCodsocio = CStr(Rs!Codsocio)
        mNomsocio = DBLet(Rs!nomsocio, "T")
        
        mDirsocio = DBLet(Rs!dirsocio, "T")
        mCodposta = DBLet(Rs!codpostal, "T")
        mPobsocio = DBLet(Rs!pobsocio, "T")
        mProsocio = DBLet(Rs!prosocio, "T")
        mNIFsocio = Rs!nifsocio
        mTfno1 = DBLet(Rs!telsoci1, "T")
        mTfno2 = DBLet(Rs!telsoci2, "T")
        mTfno3 = DBLet(Rs!telsoci3, "T")
        mMovsocio = DBLet(Rs!movsocio, "T")
        mMaisocio = DBLet(Rs!maisocio, "T")
        
        mFecNacim = DBLet(Rs!FechaNac, "T")
        mFecAlta = Rs!FechaAlta
        mFecBaja = DBLet(Rs!FechaBaja, "T")
        
       '[Monica]22/11/2013: tema iban
        mIban = CStr(DBLet(Rs!Iban, "T"))
        
        mCodbanco = CStr(Format(ComprobarCero(DBLet(Rs!CodBanco, "N")), "0000"))
        mCodsucur = CStr(Format(ComprobarCero(DBLet(Rs!CodSucur, "N")), "0000"))
        mDigContr = DBLet(Rs!digcontr, "T")
        mCuentaba = DBLet(Rs!CuentaBa, "T")
        
        mCodSitua = CStr(DBLet(Rs!codsitua, "N"))
        mTipoIRPF = CStr(DBLet(Rs!TipoIRPF, "N"))
        mTipoProd = CStr(DBLet(Rs!TipoProd, "N"))
        mCorreo = CStr(DBLet(Rs!Correo, "N"))
        mCodCoope = CStr(DBLet(Rs!codcoope, "N"))
        mEsFactADVInt = CStr(DBLet(Rs!EsFactADVInterna, "N"))
        mEmiteFact = CStr(DBLet(Rs!EmiteFact, "N"))
        
        mObservac = DBLet(Rs!Observaciones, "T")
        
        '[Monica]27/10/2015
        mvarNoDevolverApor = DBLet(Rs!NoDevolverApor, "N") ' No devolver aportaciones

        '[Monica]14/05/2018: hay embargo
        mvarHayEmbargo = DBLet(Rs!HayEmbargo, "N") ' Hay embargo
       
       
        b = LeerTipoMovimientos
        
        LeerDatos = b
    End If

    Rs.Close
    Set Rs = Nothing
    Exit Function

ELeer:
    Sql = "Se ha producido un error. " & "Datos socio" & vbCrLf
    Sql = Sql & "Número: " & Err.Number & vbCrLf
    Sql = Sql & "Descripción: " & Err.Description
    MsgBox Sql, vbExclamation
    Set Rs = Nothing
    LeerDatos = False
End Function

Public Function SocioBloqueado() As Boolean
'devuelve true si el socio esta bloqueado
'si la situación del socio es distinta de NORMAL(codsitua=0) entonces
'mostrar un mensaje con la situación especial del socio
Dim Tipo As String
Dim devuelve As String

    On Error GoTo EBloqueado
    SocioBloqueado = False
    
    If Situacion <> 0 Then '0: situacion normal
        Tipo = "tipositu"
        devuelve = DevuelveDesdeBD("ssitua", "nomsitua", "codsitua", Situacion, "N", Tipo)
        
        If Tipo = "1" Then 'socio Bloqueado por Situación Especial.
            MsgBox UCase("Socio Bloqueado por: ") & vbCrLf & devuelve, vbInformation, "Situación Especial del Socio."
            SocioBloqueado = True
        Else
            MsgBox devuelve, vbInformation, "Situación Especial del Socio."
        End If
    End If
    
EBloqueado:
    If Err.Number <> 0 Then Err.Clear
End Function


Public Function LeerDatosSeccion(vCodsocio As String, vCodsecci As String) As Boolean
'Leer los datos de la seccion del socio
'Lee de la BD: Ariagrorec, Tabla: rsocio_seccion
'OUT: True si lee los datos correctamente
Dim Rs As ADODB.Recordset
Dim Sql As String

    On Error GoTo ELeer
    
    LeerDatosSeccion = False
    
    Sql = "SELECT codmaccli, codmacpro, codiva "
    Sql = Sql & " FROM rsocios_seccion "
    Sql = Sql & " WHERE codsocio=" & vCodsocio & " and codsecci = " & vCodsecci
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText

    If Rs.EOF Then
        LeerDatosSeccion = False
    Else
        mCodsocio = CStr(vCodsocio)
        mCodSecci = CStr(vCodsecci)
        '[Monica]03/07/2013: añado trim(codmacta)
        mCtaCli = DBLet(Trim(Rs!codmaccli), "T")
        mCtaPro = DBLet(Trim(Rs!codmacpro), "T")
        mCodIva = CStr(Rs!CodIva)
        
        LeerDatosSeccion = True
    End If

    Rs.Close
    Set Rs = Nothing
    Exit Function

ELeer:
    Sql = "Se ha producido un error. " & "Datos socio/sección" & vbCrLf
    Sql = Sql & "Número: " & Err.Number & vbCrLf
    Sql = Sql & "Descripción: " & Err.Description
    MsgBox Sql, vbExclamation
    Set Rs = Nothing
    LeerDatosSeccion = False
End Function

Private Function LeerTipoMovimientos() As Boolean
'Leer los tipos de movimiento de la cooperativa a la que pertenece
'Lee de la BD: Ariagrorec, Tabla: rcoope
'OUT: True si lee los datos correctamente
Dim Rs As ADODB.Recordset
Dim Sql As String

    On Error GoTo ELeer
    
    LeerTipoMovimientos = False
    
    Sql = "SELECT codtipomant, codtipomliq, codtipomantvc, codtipomliqvc, codtipomfacalmz, codtipomfacbod, "
    Sql = Sql & " codtipomantalmz, codtipomliqalmz, codtipomantbod, codtipomliqbod "
    Sql = Sql & " FROM rcoope "
    Sql = Sql & " WHERE codcoope=" & mCodCoope
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText

    If Rs.EOF Then
        LeerTipoMovimientos = False
    Else
        mCodTipomAnt = CStr(DBLet(Rs!CodTipomAnt, "T"))
        mCodTipomLiq = CStr(DBLet(Rs!CodTipomLiq, "T"))
        mCodTipomAntVC = CStr(DBLet(Rs!CodTipomAntVC, "T"))
        mCodTipomLiqVC = CStr(DBLet(Rs!CodTipomLiqVC, "T"))
        mCodTipomFacAlmz = CStr(DBLet(Rs!CodTipomFacAlmz, "T"))
        mCodTipomFacBod = CStr(DBLet(Rs!CodTipomFacBod, "T"))
        
        mCodTipomAntAlmz = CStr(DBLet(Rs!CodTipomAntAlmz, "T"))
        mCodTipomLiqAlmz = CStr(DBLet(Rs!CodTipomLiqAlmz, "T"))
        mCodTipomAntBod = CStr(DBLet(Rs!CodTipomAntBod, "T"))
        mCodTipomLiqBod = CStr(DBLet(Rs!CodTipomLiqBod, "T"))
        
        LeerTipoMovimientos = True
    End If

    Rs.Close
    Set Rs = Nothing
    Exit Function

ELeer:
    Sql = "Se ha producido un error. " & "Tipos movimiento socio/cooperativa" & vbCrLf
    Sql = Sql & "Número: " & Err.Number & vbCrLf
    Sql = Sql & "Descripción: " & Err.Description
    MsgBox Sql, vbExclamation
    Set Rs = Nothing
    LeerTipoMovimientos = False
End Function


Public Function Estercero(vCodsocio As String, Optional SinMensaje As Boolean) As Boolean
'Comprueba si existe el socio en la BD
Dim devuelve As String

    On Error GoTo EExiste
    
    devuelve = DevuelveDesdeBDNew(cAgro, "rsocios", "tipoprod", "codsocio", vCodsocio, "N")
    If devuelve = "" Or devuelve <> "1" Then
        Estercero = False
        devuelve = "El socio " & vCodsocio & " no es tercero. "
        If Not SinMensaje Then MsgBox devuelve, vbExclamation
    Else
        Estercero = True
    End If
    
EExiste:
    If Err.Number <> 0 Then Estercero = False
End Function


