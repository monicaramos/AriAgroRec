VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Cparametros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'Variables locales que contienen valores de propiedad
'Variables asociadas a cada campo de la tabla de parametros: EMPRESAS
Private mvarCodigo As Integer       'Codigo tabla parametros empresa
Private mvarNomempre As String      'Nombre Empresa
Private mvarDomempre As String      'Domicilio Empresa
Private mvarCodpobla As String      'Cod. Postal empresa
Private mvarPobempre As String      'Poblacion empresa
Private mvarProempre As String      'Provincia empresa
Private mvarCifempre As String * 9  'C.I.F. empresa
Private mvarTelempre As String      'Teléfono empresa
Private mvarFaxempre As String      'Fax empresa
Private mvarWwwempre As String      'web empresa
Private mvarMaiempre As String      'e-mail empresa
Private mvarPercontacto As String   'persona de contacto

Private mvarFecIniCam As String       'fecha de inicio de campaña
Private mvarFecFinCam As String       'fecha fin de campaña

'Si el usuario ha modificado los datos de la empresa
'para saber si hay que actualizar los datos empresa visibles
'en el form frmPpal
Private mvarModificado As Boolean


Private SQL As String  'Para la sentencia
Private Mens As String    'Para los mensajes
Private RS As Recordset
Private Rc As Byte



'------------------------------------------------
'Propiedades del modulo Cparametros
'------------------------------------------------

Public Property Let Codigo(ByVal vData As Integer)
     mvarCodigo = vData
End Property

Public Property Get Codigo() As Integer
     Codigo = mvarCodigo
End Property


Public Property Let NombreEmpresa(ByVal vData As String)
     mvarNomempre = vData
End Property

Public Property Get NombreEmpresa() As String
     NombreEmpresa = mvarNomempre
End Property


Public Property Let DomicilioEmpresa(ByVal vData As String)
     mvarDomempre = vData
End Property

Public Property Get DomicilioEmpresa() As String
     DomicilioEmpresa = mvarDomempre
End Property


Public Property Let CPostal(ByVal vData As String)
     mvarCodpobla = vData
End Property

Public Property Get CPostal() As String
     CPostal = mvarCodpobla
End Property


Public Property Let Poblacion(ByVal vData As String)
     mvarPobempre = vData
End Property

Public Property Get Poblacion() As String
     Poblacion = mvarPobempre
End Property


Public Property Let Provincia(ByVal vData As String)
     mvarProempre = vData
End Property

Public Property Get Provincia() As String
     Provincia = mvarProempre
End Property


Public Property Let CifEmpresa(ByVal vData As String)
     mvarCifempre = vData
End Property

Public Property Get CifEmpresa() As String
     CifEmpresa = mvarCifempre
End Property

Public Property Let Telefono(ByVal vData As String)
     mvarTelempre = vData
End Property

Public Property Get Telefono() As String
     Telefono = mvarTelempre
End Property


Public Property Let Fax(ByVal vData As String)
     mvarFaxempre = vData
End Property

Public Property Get Fax() As String
     Fax = mvarFaxempre
End Property


Public Property Let WebEmpresa(ByVal vData As String)
     mvarWwwempre = vData
End Property

Public Property Get WebEmpresa() As String
     WebEmpresa = mvarWwwempre
End Property

Public Property Let MailEmpresa(ByVal vData As String)
     mvarMaiempre = vData
End Property

Public Property Get MailEmpresa() As String
     MailEmpresa = mvarMaiempre
End Property

Public Property Let PerContacto(ByVal vData As String)
     mvarPercontacto = vData
End Property

Public Property Get PerContacto() As String
     PerContacto = mvarPercontacto
End Property

Public Property Let Modificado(ByVal vData As Boolean)
     mvarModificado = vData
End Property

Public Property Get Modificado() As Boolean
     Modificado = mvarModificado
End Property

Public Property Let FecIniCam(ByVal vData As String)
     mvarFecIniCam = vData
End Property

Public Property Get FecIniCam() As String
     FecIniCam = mvarFecIniCam
End Property

Public Property Let FecFinCam(ByVal vData As String)
     mvarFecFinCam = vData
End Property

Public Property Get FecFinCam() As String
     FecFinCam = mvarFecFinCam
End Property


'------------------------------------------------
'FUNCIONES del modulo Cparametros
'------------------------------------------------

'****   LEER  ***
Public Function Leer() As Byte
    On Error GoTo Err3
    SQL = "SELECT * FROM empresas "
    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic

    If RS.EOF Then
        Leer = 1
    Else
        'DAVID
        ' Lo pongo porque para los informes con varais tablas
        'necesito linkarla y para ello debe tener un codigo
        mvarCodigo = RS!codempre
        mvarNomempre = RS!nomempre
        mvarDomempre = RS!domempre
        mvarCodpobla = RS!CodPobla
        mvarPobempre = RS!pobempre
        mvarProempre = RS!proempre
        mvarCifempre = RS!cifempre
        mvarTelempre = DBLet(RS!telempre, "T")
        mvarFaxempre = DBLet(RS!faxempre, "T")
        mvarWwwempre = DBLet(RS!wwwempre, "T")
        mvarMaiempre = DBLet(RS!maiempre, "T")
        mvarPercontacto = DBLet(RS!PerContacto, "T")
        mvarFecIniCam = DBLet(RS!FechaIni, "T")
        mvarFecFinCam = DBLet(RS!FechaFin, "T")
        
        mvarModificado = False
        Leer = 0
    End If

    RS.Close
    Set RS = Nothing
    Exit Function

Err3:
    Mens = "Se ha producido un error." & vbCrLf
    Mens = Mens & "Número: " & Err.Number & vbCrLf
    Mens = Mens & "Descripción: " & Err.Description
    MsgBox Mens, vbExclamation
    RS.Close
    Leer = 1
End Function


Public Function Modificar() As Byte
'Modifica los parametros de una empresa
'Modifica la Tabla: sparam, BD: Ariges
    On Error GoTo Err4

    mvarModificado = True

    SQL = "UPDATE empresas SET "
    SQL = SQL & " nomempre = " & DBSet(mvarNomempre, "T")
    SQL = SQL & ", domempre = " & DBSet(mvarDomempre, "T")
    SQL = SQL & ", codpobla = '" & mvarCodpobla
    SQL = SQL & "', pobempre = " & DBSet(mvarPobempre, "T")
    SQL = SQL & ", proempre = " & DBSet(mvarProempre, "T")
    SQL = SQL & ", cifempre = '" & mvarCifempre
    SQL = SQL & "', telempre = '" & mvarTelempre
    SQL = SQL & "', faxempre = '" & mvarFaxempre
    SQL = SQL & "', wwwempre = " & DBSet(mvarWwwempre, "T")
    SQL = SQL & ", maiempre = " & DBSet(mvarMaiempre, "T")
    SQL = SQL & ", percontacto = " & DBSet(mvarPercontacto, "T")
    SQL = SQL & ", fechaini = " & DBSet(mvarFecIniCam, "F")
    SQL = SQL & ", fechafin = " & DBSet(mvarFecFinCam, "F")
    SQL = SQL & " WHERE codempre =" & vEmpresa.codempre & ";"

    Conn.Execute SQL

    Exit Function

Err4:
    Mens = "Se ha producido un error." & vbCrLf
    Mens = Mens & "Número: " & Err.Number & vbCrLf
    Mens = Mens & "Descripción: " & Err.Description
    MsgBox Mens, vbExclamation
    RS.Close
    Modificar = 1
End Function


