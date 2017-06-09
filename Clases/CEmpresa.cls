VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Cempresa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'EMPRESA
'TABLA:


'ATRIBUTOS
'Variables locales que contienen valores de propiedad
Private mvarCodempre As String 'copia local
Private mvarNomempre As String 'copia local
Private mvarNomresum As String 'copia local
Private mvarAriagro As String

Private mvarnumnivel As Integer 'copia local
Private mvarnumdigi1 As Integer 'copia local
Private mvarnumdigi2 As Integer 'copia local
Private mvarnumdigi3 As Integer 'copia local
Private mvarnumdigi4 As Integer 'copia local
Private mvarnumdigi5 As Integer 'copia local
Private mvarnumdigi6 As Integer 'copia local
Private mvarnumdigi7 As Integer 'copia local
Private mvarnumdigi8 As Integer 'copia local
Private mvarnumdigi9 As Integer 'copia local
Private mvarnumdigi10 As Integer 'copia local



'VARIABLES
'Campo calculado de numero de digitos en la ULTIMA NIVEL CONTABLE
Private mvarDigitosUltimoNivel As Byte
Private mvarDigitosNivelAnterior As Byte

'David
'-----------------------------------------------
'Para  no tener que leer "cienes y cienes" de veces
'las fechas de inicio y fin en la contabilidad
'las leo en esta primera lectura de datos
'Ya que si tiene esta tabla(empresa), tiene la de parametros
Private mvarFechaIni As Date
Private mvarFechaFin As Date
'Tambien añadire si lleva analitica
Private mvarTieneAnalitica As Boolean

Private mvarConceInt As Byte 'concepto par ambos de la contabilidad
Private mvarNumDiarioInt  As Byte 'numero de diario

Private mvarConceIntProv As Byte 'concepto par ambos de la contabilidad (facturas proveedor)
Private mvarNumDiarioIntProv  As Byte 'numero de diario

Private mvarFechaUltIVA As Date


Private Sql As String  'Para la sentencia
Private Mens As String    'Para los mensajes
Private Rs As ADODB.Recordset
Private RC As Byte


Private mvarNorma19_34Nueva As Byte

Private mvarAplicarNorma19_34Nueva As Byte


Private mvarTieneSII As Byte
Private mvarSIIFechaInicio As Date
Private mvarSIIDiasAviso As Byte



'------------------------------------------------
'Propiedades del modulo Cempresa
'------------------------------------------------

'*** codigo de la empresa

Public Property Let codempre(ByVal vData As String)
     mvarCodempre = vData
End Property

Public Property Get codempre() As String
     codempre = mvarCodempre
End Property


'*** nombre de la empresa

Public Property Let nomempre(ByVal vData As String)
     mvarNomempre = vData
End Property

Public Property Get nomempre() As String
     nomempre = mvarNomempre
End Property


'*** nombre resumen de la empresa
Public Property Let nomresum(ByVal vData As String)
     mvarNomresum = vData
End Property

Public Property Get nomresum() As String
     nomresum = mvarNomresum
End Property


'*** Base de Datos de la empresa

Public Property Let BDAriagro(ByVal vData As String)
     mvarAriagro = vData
End Property

Public Property Get BDAriagro() As String
     BDAriagro = mvarAriagro
End Property


'*** nivel de las cuentas de contabilidad

Public Property Let numnivel(ByVal vData As Integer)
     mvarnumnivel = vData
End Property

Public Property Get numnivel() As Integer
     numnivel = mvarnumnivel
End Property


'*** digito de 1er nivel de las cuentas de contabilidad
Public Property Let numdigi1(ByVal vData As Integer)
     mvarnumdigi1 = vData
End Property

Public Property Get numdigi1() As Integer
     numdigi1 = mvarnumdigi1
End Property

'*** digito de nivel 2

Public Property Let numdigi2(ByVal vData As Integer)
     mvarnumdigi2 = vData
End Property

Public Property Get numdigi2() As Integer
     numdigi2 = mvarnumdigi2
End Property

'*** digito de nivel 3

Public Property Let numdigi3(ByVal vData As Integer)
     mvarnumdigi3 = vData
End Property

Public Property Get numdigi3() As Integer
     numdigi3 = mvarnumdigi3
End Property

'*** digito de nivel 4
Public Property Let numdigi4(ByVal vData As Integer)
     mvarnumdigi4 = vData
End Property

Public Property Get numdigi4() As Integer
     numdigi4 = mvarnumdigi4
End Property

'*** digito de nivel 5
Public Property Let numdigi5(ByVal vData As Integer)
     mvarnumdigi5 = vData
End Property

Public Property Get numdigi5() As Integer
     numdigi5 = mvarnumdigi5
End Property


'*** digito de nivel 6
Public Property Let numdigi6(ByVal vData As Integer)
     mvarnumdigi6 = vData
End Property

Public Property Get numdigi6() As Integer
     numdigi6 = mvarnumdigi6
End Property

'*** digito de nivel 7
Public Property Let numdigi7(ByVal vData As Integer)
     mvarnumdigi7 = vData
End Property

Public Property Get numdigi7() As Integer
     numdigi7 = mvarnumdigi7
End Property


'*** digito de nivel 8
Public Property Let numdigi8(ByVal vData As Integer)
     mvarnumdigi8 = vData
End Property

Public Property Get numdigi8() As Integer
     numdigi8 = mvarnumdigi8
End Property


'*** digito de nivel *
Public Property Let numdigi9(ByVal vData As Integer)
     mvarnumdigi9 = vData
End Property

Public Property Get numdigi9() As Integer
     numdigi9 = mvarnumdigi9
End Property

'*** digito de nivel 10
Public Property Let numdigi10(ByVal vData As Integer)
     mvarnumdigi10 = vData
End Property

Public Property Get numdigi10() As Integer
     numdigi10 = mvarnumdigi10
End Property


'*** digitos de la cuenta de ultimo nivel

Public Property Get DigitosUltimoNivel() As Byte
     DigitosUltimoNivel = mvarDigitosUltimoNivel
End Property

'*** digitos de la cuenta de nivel anterior

Public Property Get DigitosNivelAnterior() As Byte
     DigitosNivelAnterior = mvarDigitosNivelAnterior
End Property


Public Property Let FechaIni(ByVal vData As Date)
     mvarFechaIni = vData
End Property

Public Property Get FechaIni() As Date
     FechaIni = mvarFechaIni
End Property

Public Property Let FechaFin(ByVal vData As Date)
     mvarFechaFin = vData
End Property

Public Property Get FechaFin() As Date
     FechaFin = mvarFechaFin
End Property


Public Property Let TieneAnalitica(ByVal vData As Boolean)
     mvarTieneAnalitica = vData
End Property

Public Property Get TieneAnalitica() As Boolean
     TieneAnalitica = mvarTieneAnalitica
End Property


Public Property Let ConceptoInt(ByVal vData As Byte)
     mvarConceInt = vData
End Property

Public Property Get ConceptoInt() As Byte
     ConceptoInt = mvarConceInt
End Property

Public Property Let NumDiarioInt(ByVal vData As Byte)
     mvarNumDiarioInt = vData
End Property

Public Property Get NumDiarioInt() As Byte
     NumDiarioInt = mvarNumDiarioInt
End Property

Public Property Let ConceptoIntProv(ByVal vData As Byte)
     mvarConceIntProv = vData
End Property

Public Property Get ConceptoIntProv() As Byte
     ConceptoIntProv = mvarConceIntProv
End Property

Public Property Let NumDiarioIntprov(ByVal vData As Byte)
     mvarNumDiarioIntProv = vData
End Property

Public Property Get NumDiarioIntprov() As Byte
     NumDiarioIntprov = mvarNumDiarioIntProv
End Property


Public Property Let FechaUltIVA(ByVal vData As Date)
     mvarFechaUltIVA = vData
End Property

Public Property Get FechaUltIVA() As Date
     FechaUltIVA = mvarFechaUltIVA
End Property

'[Monica]20/11/2013
'*** Indicamos si esta con la nueva norma o no
Public Property Get HayNorma19_34Nueva() As Byte
     HayNorma19_34Nueva = mvarNorma19_34Nueva
End Property

'[Monica]22/11/2013
'*** Indicamos si esta con la nueva norma o no
Public Property Get AplicarNorma19_34Nueva() As Byte
     AplicarNorma19_34Nueva = mvarAplicarNorma19_34Nueva
End Property


'[Monica]09/06/2017: fecha de inicio y dias de aviso del sii
Public Property Get TieneSII() As Byte
     TieneSII = mvarTieneSII
End Property


Public Property Get SIIDiasAviso() As Byte
     SIIDiasAviso = mvarSIIDiasAviso
End Property


Public Property Get SIIFechaInicio() As Date
     SIIFechaInicio = mvarSIIFechaInicio
End Property






'------------------------------------------------
'FUNCIONES del modulo Cempresa
'------------------------------------------------

'****   AGREGAR  ***
Public Function Agregar() As Byte
'Añadir una nueva empresa en la BD de usuarios tabla:empresasariges
    
    On Error GoTo Err1
    'La Base de datos "Usuarios" esta en el mismo servidor que BD "Ariges"
    Sql = "INSERT INTO usuarios.empresasariagro (codempre, nomempre, nomresum)"
    Sql = Sql & " VALUES ("
    'Datos
    Sql = Sql & mvarCodempre
    Sql = Sql & "," & DBSet(mvarNomempre, "T")
    Sql = Sql & "," & DBSet(mvarNomresum, "T")
    Sql = Sql & ")"

    conn.Execute Sql
    Agregar = 0
    Exit Function
    
Err1:
    Mens = "Se ha producido un error." & vbCrLf
    Mens = Mens & "Número: " & Err.Number & vbCrLf
    Mens = Mens & "Descripción: " & Err.Description
    MsgBox Mens, vbExclamation
    Agregar = 1
End Function



'****   LEER  ***
Public Function LeerDatos() As Byte
'Leer los parametros de una empresa
'Lee de la BD: Usuarios, Tabla: EmpresasAriges

    On Error GoTo Err3
    Sql = "SELECT * FROM usuarios.empresasariagro WHERE ariagro='" & vUsu.CadenaConexion & "'"
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    Mens = ""
    
    If Rs.EOF Then
        LeerDatos = 1
    Else
        Mens = "Datos basicos"
        mvarCodempre = Rs!codempre
        mvarNomempre = Rs!nomempre
        mvarNomresum = DBLet(Rs!nomresum)
        mvarAriagro = DBLet(Rs!Ariagro)
        Mens = "Datos de Gestión"
        LeerDatos = 0
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
    LeerDatos = 1
End Function


'****   LEER  ***
Public Function LeerNiveles() As Byte
'Leer los parametros de la empresa en BDatos "Conta" correspondiente
'a la empresa de "Ariges" que estamos conectados
Dim J As Integer
Dim jj As Integer

    On Error GoTo Err3
    
    Sql = "SELECT * FROM empresa "
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, ConnConta, adOpenForwardOnly, adLockOptimistic, adCmdText
    Mens = ""
    
    If Rs.EOF Then
        LeerNiveles = 1
    Else
        Mens = "Datos basicos empresa en contabilidad."
        mvarnumnivel = Rs!numnivel
        mvarnumdigi1 = Rs!numdigi1
        mvarnumdigi2 = DBLet(Rs!numdigi2, "N")
        mvarnumdigi3 = DBLet(Rs!numdigi3, "N")
        mvarnumdigi4 = DBLet(Rs!numdigi4, "N")
        mvarnumdigi5 = DBLet(Rs!numdigi5, "N")
        mvarnumdigi6 = DBLet(Rs!numdigi6, "N")
        mvarnumdigi7 = DBLet(Rs!numdigi7, "N")
        mvarnumdigi8 = DBLet(Rs!numdigi8, "N")
        mvarnumdigi9 = DBLet(Rs!numdigi9, "N")
        mvarnumdigi10 = DBLet(Rs!numdigi10, "N")
        'Fijamos la longitud maxima de digitos en ctas de ultimo nivel
        FijaLongitudMaxima
        
        ' leemos los datos de contabilidad
        mvarConceInt = DevuelveDesdeBDNew(cConta, "parametros", "concefcl", "1", "1", "N")
        mvarNumDiarioInt = DevuelveDesdeBDNew(cConta, "parametros", "numdiacl", "1", "1", "N")
        
        ' leemos los datos de contabilidad (proveedores)
        mvarConceIntProv = DevuelveDesdeBDNew(cConta, "parametros", "concefpr", "1", "1", "N")
        mvarNumDiarioIntProv = DevuelveDesdeBDNew(cConta, "parametros", "numdiapr", "1", "1", "N")
        
        LeerNiveles = 0
    End If
    Rs.Close
    
    
    'Asigno los valores por defecto a las variables de fechaini y fechafin
    If LeerNiveles = 0 Then
        LeerNiveles = 2
        Sql = "Select * from parametros"
        Rs.Open Sql, ConnConta, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Rs.EOF Then
            MsgBox "Fecha inicio y fin de ejercicio", vbExclamation
        Else
            mvarFechaIni = Rs!FechaIni
            mvarFechaFin = Rs!FechaFin
            
            mvarTieneAnalitica = Rs!Autocoste = 1
            
            '[Monica]02/09/2013: añadida la fecha de ultima liquidacion de iva
            'Para establecer del ultimo dia liquidado
            If DBLet(Rs!periodos, "N") = 1 Then ' es mensual
               J = DBLet(Rs!perfactu, "N")
            Else
                'Trimestral
                J = DBLet(Rs!perfactu, "N") * 3
            End If
            If J = 0 Then J = 1
            jj = DBLet(Rs!Anofactu, "N")
            If jj = 0 Then jj = 1900
            
            mvarFechaUltIVA = "01/" & J & "/" & jj
            mvarFechaUltIVA = DateAdd("m", 1, mvarFechaUltIVA)  'mas un mes
            mvarFechaUltIVA = DateAdd("d", -1, mvarFechaUltIVA)  'menos un dia
       
            
       End If
       Rs.Close
       
       '[Monica]22/11/2013: miramos lo del iban
       mvarNorma19_34Nueva = TieneNorma19_34Nueva
       If mvarNorma19_34Nueva = 1 Then
            mvarAplicarNorma19_34Nueva = CByte(ComprobarCero(DevuelveDesdeBDNew(cConta, "paramtesor", "Norma19_34Nueva", "1", "1", "N")))
       End If
       
       '[Monica]09/06/2017: miramos si tiene el sii
       mvarTieneSII = LeerSII
       
       
       
       
    End If
    
    
    Set Rs = Nothing
    Exit Function

Err3:
    Mens = "Se ha producido un error. " & Mens & vbCrLf
    Mens = Mens & "Número: " & Err.Number & vbCrLf
    Mens = Mens & "Descripción: " & Err.Description
    MsgBox Mens, vbExclamation
    Set Rs = Nothing
    LeerNiveles = 1
End Function

Private Function LeerSII() As Boolean
Dim Sql As String

    On Error Resume Next

    Sql = "select SIIFechaInicio, SIIDiasAviso from parametros "
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, ConnConta, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    LeerSII = True
    If Not Rs.EOF Then
        mvarSIIFechaInicio = DBLet(Rs!SIIFechaInicio, "F")
        mvarSIIDiasAviso = DBLet(Rs!SIIDiasAviso, "N")
    End If
    Rs.Close
    Set Rs = Nothing

    If Err.Number <> 0 Then
        LeerSII = False
        Err.Clear
    End If
End Function


Private Function TieneNorma19_34Nueva() As Byte
Dim Sql As String
Dim Rs As ADODB.Recordset

    On Error GoTo eTieneNorma19_34Nueva

    TieneNorma19_34Nueva = 0

    Set Rs = New ADODB.Recordset

    Sql = "select norma19_34Nueva from paramtesor"
    Rs.Open Sql, ConnConta, adOpenForwardOnly, adLockPessimistic, adCmdText
         
    TieneNorma19_34Nueva = 1
    
    Set Rs = Nothing
    Exit Function
    
eTieneNorma19_34Nueva:

End Function


'****   MODIFICAR  ***
Public Function ModificarDatos() As Byte
'Modifica los parametros de una empresa
'Modifica la Tabla:empresasariges en la BD: Usuarios
    
    On Error GoTo Err4

    'Modificamos en la tabla Usuarios.empresasariges
    Sql = "UPDATE usuarios.empresasariagro SET nomempre= '" & mvarNomempre
    Sql = Sql & "', nomresum= '" & mvarNomresum
    Sql = Sql & "' WHERE codempre=" & mvarCodempre
    conn.Execute Sql
    
    ModificarDatos = 0
    Exit Function

Err4:
    Mens = "Se ha producido un error." & vbCrLf
    Mens = Mens & "Número: " & Err.Number & vbCrLf
    Mens = Mens & "Descripción: " & Err.Description
    MsgBox Mens, vbExclamation
    Rs.Close
    ModificarDatos = 1
End Function



'---------------------------------------------------------
'FUNCIONES del modulo Cempresa referentes a Contabilidad
'---------------------------------------------------------

Private Sub FijaLongitudMaxima()
Dim I As Byte

    Select Case mvarnumnivel
        Case 1
            I = mvarnumdigi1
        Case 2
            I = mvarnumdigi2
        Case 3
            I = mvarnumdigi3
        Case 4
            I = mvarnumdigi4
        Case 5
            I = mvarnumdigi5
        Case 6
            I = mvarnumdigi6
        Case 7
            I = mvarnumdigi7
        Case 8
            I = mvarnumdigi8
        Case 9
            I = mvarnumdigi9
        Case 10
            I = mvarnumdigi10
    End Select
    mvarDigitosUltimoNivel = I
    
    Select Case mvarnumnivel - 1
        Case 1
            I = mvarnumdigi1
        Case 2
            I = mvarnumdigi2
        Case 3
            I = mvarnumdigi3
        Case 4
            I = mvarnumdigi4
        Case 5
            I = mvarnumdigi5
        Case 6
            I = mvarnumdigi6
        Case 7
            I = mvarnumdigi7
        Case 8
            I = mvarnumdigi8
        Case 9
            I = mvarnumdigi9
        Case 10
            I = mvarnumdigi10
    End Select
    mvarDigitosNivelAnterior = I
End Sub


'Devuelve, para un nivel determinado, cuantos digitos tienen las cuentas
' a ese nivel
Public Function DigitosNivel(numnivel As Integer) As Integer
    Select Case numnivel
    Case 1
        DigitosNivel = vEmpresa.numdigi1

    Case 2
        DigitosNivel = vEmpresa.numdigi2

    Case 3
        DigitosNivel = vEmpresa.numdigi3

    Case 4
        DigitosNivel = vEmpresa.numdigi4

    Case 5
        DigitosNivel = vEmpresa.numdigi5

    Case 6
        DigitosNivel = vEmpresa.numdigi6

    Case 7
        DigitosNivel = vEmpresa.numdigi7

    Case 8
        DigitosNivel = vEmpresa.numdigi8

    Case 9
        DigitosNivel = vEmpresa.numdigi9

    Case 10
        DigitosNivel = vEmpresa.numdigi10

    Case Else
        DigitosNivel = -1
    End Select
End Function

