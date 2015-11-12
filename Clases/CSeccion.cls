VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSeccion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

'Secciones
'TABLA: rseccion


'ATRIBUTOS
'Variables locales que contienen valores de propiedad
Private mCodSecci As String 'Cod. seccion
Private mNomSecci As String 'Nombre seccion

Private mCodigoConta As Byte 'Numero de contabilidad
Private mRaizSocio As String 'Raiz Socio
Private mRaizAsoc As String 'Raiz Asociado
Private mRaizProv As String 'Raiz Proveedor
Private mTipIvaExento As Integer ' tipo de iva exento de la contabilidad


'------------------------------------------------
'Propiedades del modulo CSeccion
'------------------------------------------------

'**** Codigo de la seccion

Public Property Let Codigo(ByVal vData As String)
     mCodSecci = vData
End Property

Public Property Get Codigo() As String
     Codigo = mCodSecci
End Property

'**** Nombre de la seccion

Public Property Let Nombre(ByVal vData As String)
     mNomSecci = vData
End Property

Public Property Get Nombre() As String
     Nombre = mNomSecci
End Property


'**** contabilidad de la seccion

Public Property Let Conta(ByVal vData As Byte)
     mCodigoConta = vData
End Property

Public Property Get Conta() As Byte
     Conta = mCodigoConta
End Property


'**** Raiz Socio

Public Property Let RaizSocio(ByVal vData As String)
     mRaizSocio = vData
End Property

Public Property Get RaizSocio() As String
     RaizSocio = mRaizSocio
End Property

'**** Raiz Asociado

Public Property Let RaizAsoc(ByVal vData As String)
     mRaizAsoc = vData
End Property

Public Property Get RaizAsoc() As String
     RaizAsoc = mRaizAsoc
End Property

'**** Raiz Proveedor

Public Property Let RaizProv(ByVal vData As String)
     mRaizProv = vData
End Property

Public Property Get RaizProv() As String
     RaizProv = mRaizProv
End Property

'**** tipo de iva exento de la contabilidad

Public Property Let TipIvaExento(ByVal vData As Integer)
     mTipIvaExento = vData
End Property

Public Property Get TipIvaExento() As Integer
     TipIvaExento = mTipIvaExento
End Property


'------------------------------------------------
'Procedimientos del modulo Cseccion
'------------------------------------------------

Public Function Existe(vCodsecci As String) As Boolean
'Comprueba si existe la seccion en la BD
Dim devuelve As String

    On Error GoTo EExiste
    
    devuelve = DevuelveDesdeBD("codsecci", "rseccion", "codsecci", vCodsecci, "N")
    If devuelve = "" Then
        Existe = False
        devuelve = "No existe la seccion " & vCodsecci & ". "
        MsgBox devuelve, vbExclamation
    Else
        Existe = True
    End If
    
EExiste:
    If Err.Number <> 0 Then Existe = False
End Function



Public Function LeerDatos(vCodsecci As String) As Boolean
'Leer los datos de una seccion dada
'Lee de la BD: Ariagrorec, Tabla: rseccion
'OUT: True si lee los datos correctamente
Dim RS As ADODB.Recordset
Dim SQL As String

    On Error GoTo ELeer
    
    LeerDatos = False
    
    SQL = "SELECT codsecci, nomsecci, empresa_conta, raiz_cliente_socio, raiz_cliente_asociado, raiz_proveedor, codivaexe "
    SQL = SQL & " FROM rseccion "
    SQL = SQL & " WHERE codsecci=" & vCodsecci
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText

    If RS.EOF Then
        LeerDatos = False
    Else
        mCodSecci = CStr(RS!codsecci)
        
        mNomSecci = RS!nomsecci
        mCodigoConta = CStr(DBLet(RS!empresa_conta, "N"))
        mRaizSocio = DBLet(RS!raiz_cliente_socio, "T")
        mRaizAsoc = DBLet(RS!raiz_cliente_asociado, "T")
        mRaizProv = DBLet(RS!raiz_proveedor, "T")
        mTipIvaExento = DBLet(RS!codivaexe, "N")
        
        
        LeerDatos = True
    End If

    RS.Close
    Set RS = Nothing
    Exit Function

ELeer:
    SQL = "Se ha producido un error. " & "Datos sección" & vbCrLf
    SQL = SQL & "Número: " & Err.Number & vbCrLf
    SQL = SQL & "Descripción: " & Err.Description
    MsgBox SQL, vbExclamation
    Set RS = Nothing
    LeerDatos = False
End Function

Public Function CtaConRaizCorrecta(Codmacta As String, CliPro As Byte) As Boolean

    Select Case CliPro
        Case 0 'cliente
            CtaConRaizCorrecta = (Mid(Codmacta, 1, vEmpresa.DigitosNivelAnterior) = Trim(mRaizSocio))
        Case 1 'proveedor
            CtaConRaizCorrecta = (Mid(Codmacta, 1, vEmpresa.DigitosNivelAnterior) = Trim(mRaizProv))
    End Select

End Function

Public Function AbrirConta() As Boolean
' abrimos la bd de contabilidad
Dim SQL As String

    On Error GoTo ELeer
    
    AbrirConta = False
    
    If AbrirConexionConta2(CInt(mCodigoConta)) Then
        If vEmpresa.LeerNiveles Then
            AbrirConta = True
        End If
    End If
    Exit Function

ELeer:
    SQL = "Se ha producido un error. " & "Abrir Conta" & vbCrLf
    SQL = SQL & "Número: " & Err.Number & vbCrLf
    SQL = SQL & "Descripción: " & Err.Description
    MsgBox SQL, vbExclamation
    Set vEmpresa = Nothing
    AbrirConta = False
End Function


Public Function CerrarConta() As Boolean
' cerramos la conexion de contabilidad
    CerrarConexionConta2
End Function


Public Function AbrirAriges() As Boolean
' abrimos la bd de ariges (suministros)
Dim SQL As String

    On Error GoTo ELeer
    
    AbrirAriges = False
    
    If AbrirConexionAriges() Then
         AbrirAriges = True
    End If
    Exit Function

ELeer:
    SQL = "Se ha producido un error. " & "Abrir Ariges" & vbCrLf
    SQL = SQL & "Número: " & Err.Number & vbCrLf
    SQL = SQL & "Descripción: " & Err.Description
    MsgBox SQL, vbExclamation
    Set vEmpresa = Nothing
End Function


Public Function CerrarAriges() As Boolean
' cerramos la conexion de ariges (suministros)
    CerrarConexionAriges
End Function



