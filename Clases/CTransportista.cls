VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTransportista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'TABLA: rtransporte
'   transportistas
'Variables locales que contienen valores de propiedad
Private mCodTrans As String 'Cod. transportista
Private mNomTrans As String 'Nombre transportista

Private mDirTrans As String 'Dierccion trans
Private mCodposta As String 'cod. postal trans
Private mPobTrans As String 'poblacion trans
Private mProTrans As String 'Provincia trans
Private mNIFTrans As String 'NIF trans

Private mMatricul As String 'matricula
Private mTaraVehi As String 'tara de vehiculo
Private mTipoIRPF As Byte 'tipo irpf: 0=modulos 1=E.D 2=entidad

Private mCodForpa As String 'cod. forma de pago
Private mCtaTra As String
Private mCodIva As String

Private mContador As String
Private mEsFactTraInterna As Byte '0=no interna 1=interna

'------------------------------------------------
'------------------------------------------------
'Propiedades del modulo CTransportista
'------------------------------------------------
'------------------------------------------------


'**** Codigo del transportista

Public Property Let Codigo(ByVal vData As String)
     mCodTrans = vData
End Property

Public Property Get Codigo() As String
     Codigo = mCodTrans
End Property


'**** Nombre del transportista

Public Property Let Nombre(ByVal vData As String)
     mNomTrans = vData
End Property

Public Property Get Nombre() As String
     Nombre = mNomTrans
End Property


'**** Matricula del transportista

Public Property Let Matricula(ByVal vData As String)
     mMatricul = vData
End Property

Public Property Get Matricula() As String
     Matricula = mMatricul
End Property

'**** Tara del vehiculo

Public Property Let TaraVehi(ByVal vData As String)
     mTaraVehi = vData
End Property

Public Property Get TaraVehi() As String
     TaraVehi = mTaraVehi
End Property


'**** Tipo de IRPF

Public Property Let TipoIRPF(ByVal vData As Byte)
     mTipoIRPF = vData
End Property

Public Property Get TipoIRPF() As Byte
     TipoIRPF = mTipoIRPF
End Property



'**** Forma de pago del transportista

Public Property Let ForPago(ByVal vData As String)
     mCodForpa = vData
End Property

Public Property Get ForPago() As String
     ForPago = mCodForpa
End Property


'**** Cuenta Contable transportista

Public Property Let CtaTRA(ByVal vData As String)
     mCtaTra = vData
End Property

Public Property Get CtaTRA() As String
     CtaTRA = mCtaTra
End Property


'**** Código de IVA

Public Property Let CodIva(ByVal vData As String)
     mCodIva = vData
End Property

Public Property Get CodIva() As String
     CodIva = mCodIva
End Property


'**** Contador de Factura

Public Property Let Contador(ByVal vData As String)
     mContador = vData
End Property

Public Property Get Contador() As String
     Contador = mContador
End Property

'**** Es Factura de Transporte Interna

Public Property Let EsFactTraInterna(ByVal vData As Byte)
     mEsFactTraInterna = vData
End Property

Public Property Get EsFactTraInterna() As Byte
     EsFactTraInterna = mEsFactTraInterna
End Property


'**** Direccion del trans

Public Property Let Direccion(ByVal vData As String)
     mDirTrans = vData
End Property

Public Property Get Direccion() As String
     Direccion = mDirTrans
End Property

'**** CPostal del trans

Public Property Let CPostal(ByVal vData As String)
     mCodposta = vData
End Property

Public Property Get CPostal() As String
     CPostal = mCodposta
End Property

'**** Poblacion del trans

Public Property Let Poblacion(ByVal vData As String)
     mPobTrans = vData
End Property

Public Property Get Poblacion() As String
     Poblacion = mPobTrans
End Property

'**** Provincia del trans

Public Property Let Provincia(ByVal vData As String)
     mProTrans = vData
End Property

Public Property Get Provincia() As String
     Provincia = mProTrans
End Property

'**** NIF del trans

Public Property Let nif(ByVal vData As String)
     mNIFTrans = vData
End Property

Public Property Get nif() As String
     nif = mNIFTrans
End Property




'------------------------------------------------
'------------------------------------------------
'Procedimientos del modulo CTransportista
'------------------------------------------------
'------------------------------------------------

Public Function Existe(vCodTrans As String) As Boolean
'Comprueba si existe el transportista en la BD
Dim devuelve As String

    On Error GoTo EExiste
    
    devuelve = DevuelveDesdeBDNew(cAgro, "rtransporte", "codtrans", "codtrans", vCodTrans, "T")
    If devuelve = "" Then
        Existe = False
        devuelve = "No existe el Transportista " & vCodTrans & ". "
        MsgBox devuelve, vbExclamation
    Else
        Existe = True
    End If
    
EExiste:
    If Err.Number <> 0 Then Existe = False
End Function


Public Function LeerDatos(vCodTrans As String) As Boolean
'Leer los datos de un tranportista dado
'Lee de la BD: Ariges, Tabla: agencias
'OUT: True si lee los datos correctamente
Dim Rs As ADODB.Recordset
Dim Mens As String
Dim Sql As String

    On Error GoTo ELeer
    LeerDatos = False
    
    Sql = "SELECT codtrans,nomtrans,matricula,taravehi,tipoirpf,codmacpro,codiva,codforpa,esfacttrainterna, "
    Sql = Sql & "dirtrans,pobtrans,codpostal,protrans,niftrans "
    Sql = Sql & " FROM rtransporte "
    Sql = Sql & " WHERE codtrans=" & DBSet(vCodTrans, "T")
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    Mens = ""
    If Rs.EOF Then
        LeerDatos = False
    Else
        Mens = "Datos Transportista"
        mCodTrans = CStr(Rs!codTrans)
        
        mNomTrans = Rs!nomtrans
        
        mCodForpa = CStr(Rs!Codforpa)
        mTipoIRPF = CStr(DBLet(Rs!TipoIRPF, "N"))
        mCtaTra = DBLet(Rs!codmacpro, "T")
        mCodIva = CStr(Rs!CodIva)
        mEsFactTraInterna = DBLet(Rs!EsFactTraInterna, "N")
        
        
        mDirTrans = DBLet(Rs!dirtrans, "T")
        mCodposta = DBLet(Rs!codpostal, "T")
        mPobTrans = DBLet(Rs!pobtrans, "T")
        mProTrans = DBLet(Rs!protrans, "T")
        mNIFTrans = Rs!NIFTrans
        
        
        LeerDatos = True
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


Public Function ConseguirContador() As Long
Dim OK As Boolean
Dim c1 As Long 'contador
Dim Rs As ADODB.Recordset
Dim Sql As String
Dim Mens As String

On Error GoTo Err1
    
    ConseguirContador = 1
    
    'Abrimos bloqueando
    Sql = "Select contador from rtransporte WHERE codtrans='" & Trim(mCodTrans) & "' FOR UPDATE"
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Not Rs.EOF Then
        mContador = Rs!Contador
        c1 = Rs!Contador
        OK = True
    Else
        MsgBox "No existe contador para ese transportista.", vbExclamation
        OK = False
    End If
    Rs.Close
    Set Rs = Nothing
     
    If OK Then ConseguirContador = c1 + 1
    Exit Function
Err1:
    Mens = "Error: " & Err.Number & " : " & Err.Description
    MsgBox Mens, vbExclamation
    ConseguirContador = -1
End Function


Public Function IncrementarContador() As Boolean
Dim c1 As Long 'contador1: recibo de entrada
Dim Sql As String
On Error GoTo Err1
    
      
    'Actualizamos el contador
    c1 = Contador
    Sql = "UPDATE rtransporte set "
    c1 = c1 + 1
    mContador = c1
    Sql = Sql & " contador=" & c1
    
    Sql = Sql & " WHERE codtrans = '" & Trim(mCodTrans) & "'"
    conn.Execute Sql
    IncrementarContador = True
    Exit Function
Err1:
    If Err.Number <> 0 Then
        IncrementarContador = False
    Else
        IncrementarContador = True
    End If
End Function

Public Function ActualizarContador() As Boolean
Dim c1 As Long 'contador1: recibo de entrada
Dim Sql As String
On Error GoTo Err1
    
      
    'Actualizamos el contador
    c1 = Contador
    Sql = "UPDATE rtransporte set "
    mContador = c1
    Sql = Sql & " contador=" & c1
    
    Sql = Sql & " WHERE codtrans = '" & Trim(mCodTrans) & "'"
    conn.Execute Sql
    ActualizarContador = True
    Exit Function
Err1:
    If Err.Number <> 0 Then
        ActualizarContador = False
    Else
        ActualizarContador = True
    End If
End Function

Public Function DevolverContador(vCodigo As String, Contador As Long) As Byte
Dim OK As Boolean
Dim c1 As Long
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim Mens As String

On Error GoTo Err1
    
    'Abrimos bloqueando
    Sql = "Select contador from rtransporte WHERE codtrans = '" & vCodigo & "' FOR UPDATE " ' "' FOR UPDATE"
    'Esto esba asi antes, sin comentar
    'Conn.Execute "Set autocommit = 0"
    DevolverContador = 1
    OK = False
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Not Rs.EOF Then
        c1 = Rs!Contador
        OK = True
    End If
    Rs.Close
    Set Rs = Nothing
    If Not OK Then
        conn.Execute "commit"
        Exit Function
    End If

    OK = False

    OK = (c1 = Contador)
    If OK Then
        'Actualizamos el contador
        Sql = "UPDATE rtransporte set "

        c1 = c1 - 1
        Sql = Sql & " contador=" & c1
        Sql = Sql & " WHERE codtrans = '" & Trim(vCodigo) & "'"
        conn.Execute Sql
    End If

    'Desbloqueamos
    conn.Execute "commit"
    DevolverContador = 1
    
Err1:
    If Err.Number <> 0 Then
        Mens = "Error: " & Err.Number & " : " & Err.Description
        MsgBox Mens, vbExclamation
        DevolverContador = 0
    End If
    conn.Execute "Set autocommit = 1"
End Function
