VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCampAnt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'EMPRESA
'TABLA:

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

Private mvarBaseDatos As String    'base de datos (ariagro1, ariagro2...)



Private Sql As String  'Para la sentencia
Private Sql1 As String
Private Mens As String    'Para los mensajes
Private Rs As Recordset
Private RS1 As Recordset
Private RC As Byte



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

Public Property Let BaseDatos(ByVal vData As String)
     mvarBaseDatos = vData
End Property

Public Property Get BaseDatos() As String
     BaseDatos = mvarBaseDatos
End Property

'------------------------------------------------
'FUNCIONES del modulo CCampAnt
'------------------------------------------------

'****   LEER  ***
Public Function Leer(Optional OcultarMens As Boolean) As Byte
'Leer los parametros de una empresa
'Lee de la BD: Usuarios, Tabla: EmpresasAriagro
Dim Encontrado As Boolean


    On Error GoTo Err3
    
    Leer = 0
    
    
    Sql = "SELECT * FROM usuarios.empresasariagro WHERE ariagro<>'" & vUsu.CadenaConexion & "'"
    Sql = Sql & " order by codempre "
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    Mens = ""
    
    Encontrado = False
    
    While Not Rs.EOF And Not Encontrado
        If AbrirConexionCampAnterior(DBLet(Rs!Ariagro, "T")) Then
        
            Sql1 = "select * from empresas "
            Set RS1 = New ADODB.Recordset
            RS1.Open Sql1, ConnCAnt, adOpenForwardOnly, adLockOptimistic, adCmdText
            
            If DBLet(RS1!FechaFin, "F") = (CDate(vParam.FecIniCam) - 1) Then
                    Encontrado = True
                    
'codempre,nomempre,domempre,codpobla,pobempre,proempre,cifempre,telempre,faxempre,wwwempre,maiempre,fechaini,fechafin,percontacto
                    
                    mvarCodigo = ComprobarCero(RS1!codempre)
                    mvarNomempre = DBLet(RS1!nomempre, "T")
                    mvarDomempre = DBLet(RS1!domempre, "T")
                    mvarCodpobla = DBLet(RS1!CodPobla, "T")
                    mvarPobempre = DBLet(RS1!pobempre, "T")
                    mvarProempre = DBLet(RS1!proempre, "T")
                    mvarCifempre = DBLet(RS1!CifEmpre, "T")
                    mvarTelempre = DBLet(RS1!telempre, "T")
                    mvarFaxempre = DBLet(RS1!faxempre, "T")
                    mvarWwwempre = DBLet(RS1!wwwempre, "T")
                    mvarMaiempre = DBLet(RS1!maiempre, "T")
                    mvarPercontacto = DBLet(RS1!PerContacto, "T")
                    mvarFecIniCam = CStr(RS1!FechaIni)
                    mvarFecFinCam = CStr(RS1!FechaFin)
                    
                    mvarBaseDatos = DBLet(Rs!Ariagro, "T")
                    
                    Mens = "Datos de Ariagro Campaña anterior"
                    Leer = 0
                End If
        End If
        CerrarConexionCampAnterior
        Rs.MoveNext
    Wend
    
    If Not Encontrado Then
        If Not OcultarMens Then MsgBox "No hay campaña anterior.", vbExclamation
        Leer = 1
    End If

    Rs.Close
    Set Rs = Nothing
    Exit Function

Err3:
    Mens = "Se ha producido un error. " & Mens & vbCrLf
    Mens = Mens & "Número: " & Err.Number & vbCrLf
    Mens = Mens & "Descripción: " & Err.Description
    MsgBox Mens, vbExclamation
    Set Rs = Nothing
    Leer = 0
End Function


Public Function LeerAnterior(Optional OcultarMens As Boolean) As String
'Devuelve el nombre de la base de datos de la campaña anterior a la anterior campaña
Dim Encontrado As Boolean
Dim Rs As ADODB.Recordset
Dim RS1 As ADODB.Recordset

    On Error GoTo Err3
    
    LeerAnterior = ""
    
    Sql = "SELECT * FROM usuarios.empresasariagro WHERE ariagro<>'" & vUsu.CadenaConexion & "'"
    Sql = Sql & " order by codempre "
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    Mens = ""
    Encontrado = False
    While Not Rs.EOF And Not Encontrado
        If DBLet(Rs!Ariagro, "T") <> "" Then
            Sql1 = "select * from " & Rs!Ariagro & ".empresas where fechafin = " & DBSet(mvarFecIniCam, "F") & " - INTERVAL 1 DAY"
            Set RS1 = New ADODB.Recordset
            RS1.Open Sql1, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
            If Not RS1.EOF Then
                LeerAnterior = DBLet(RS1!Ariagro, "T")
                Encontrado = True
            End If
            Set RS1 = Nothing
        End If
        Rs.MoveNext
    Wend
    
    Exit Function

Err3:
    Mens = "Se ha producido un error. " & Mens & vbCrLf
    Mens = Mens & "Número: " & Err.Number & vbCrLf
    Mens = Mens & "Descripción: " & Err.Description
    MsgBox Mens, vbExclamation
    Set Rs = Nothing
End Function



