VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTiposMov"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'variables locales que contienen valores de propiedad

'PPPPP


Private mvarTipoMov As String 'copia local
Private mvarNombreMov As String 'copia local
Private mvarMueveStock As Boolean
Private mvarContador As Long  'copia local
Private mvarTipoDocu As Byte  'copia local
Private mvarLetraSerie As String * 1 'Letra Serie Ariconta

Private Rs As ADODB.Recordset
Private Sql As String
Private Mens As String


Public Property Let TipoMovimiento(ByVal vData As String)
    mvarTipoMov = vData
End Property


Public Property Get TipoMovimiento() As String
    TipoMovimiento = mvarTipoMov
End Property

Public Property Let NombreMovimiento(ByVal vData As String)
    mvarNombreMov = vData
End Property


Public Property Get NombreMovimiento() As String
    NombreMovimiento = mvarNombreMov
End Property


Public Property Let Contador(ByVal vData As Long)
    mvarContador = vData
End Property


Public Property Get Contador() As Long
    Contador = mvarContador
End Property

Public Property Let MueveStock(ByVal vData As Boolean)
    mvarMueveStock = vData
End Property


Public Property Get MueveStock() As Boolean
    MueveStock = mvarMueveStock
End Property

Public Property Let LetraSerie(ByVal vData As String)
    mvarLetraSerie = vData
End Property


Public Property Get LetraSerie() As String
    LetraSerie = mvarLetraSerie
End Property

Public Property Let TipoDocu(ByVal vData As Byte)
    mvarTipoDocu = vData
End Property


Public Property Get TipoDocu() As Byte
    TipoDocu = mvarTipoDocu
End Property


'//////////////////////////////////////////////////////////////////7
'//
'//     Si es el ultimo numero de la serie devuelve el anterior
Public Function DevolverContador(vCodigo As String, Contador As Long) As Byte
Dim OK As Boolean
Dim c1 As Long
On Error GoTo Err1
    
    'Abrimos bloqueando
    Sql = "Select * from usuarios.stipom WHERE codtipom = '" & vCodigo & "' FOR UPDATE " ' "' FOR UPDATE"
    'Esto esba asi antes, sin comentar
    'Conn.Execute "Set autocommit = 0"
    DevolverContador = 1
    OK = False
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Not Rs.EOF Then
        mvarTipoMov = vCodigo
        mvarNombreMov = Rs!nomtipom
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
        Sql = "UPDATE usuarios.stipom set "

        c1 = c1 - 1
        Sql = Sql & " contador=" & c1
        Sql = Sql & " WHERE codtipom = '" & vCodigo & "'"
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


Public Function ConseguirContador(Tipo As String) As Long
Dim OK As Boolean
Dim c1 As Long 'contador
Dim Rs As ADODB.Recordset
On Error GoTo Err1
    
    ConseguirContador = 1
    
    'Abrimos bloqueando
    Sql = "Select * from usuarios.stipom WHERE codtipom='" & Tipo & "' FOR UPDATE"
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Not Rs.EOF Then
        mvarTipoMov = Tipo
        mvarNombreMov = Rs!nomtipom
        mvarMueveStock = Rs!muevesto
        mvarContador = Rs!Contador
        mvarTipoDocu = Rs!TipoDocu
        c1 = Rs!Contador
        OK = True
    Else
        MsgBox "No existe contador para ese tipo de movimiento.", vbExclamation
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


Public Function IncrementarContador(vCodigo As String) As Boolean
Dim c1 As Long 'contador1: recibo de entrada
On Error GoTo Err1
    
      
    'Actualizamos el contador
    c1 = Contador
    Sql = "UPDATE usuarios.stipom set "
    c1 = c1 + 1
    mvarContador = c1
    Sql = Sql & " contador=" & c1
    
    Sql = Sql & " WHERE codtipom='" & vCodigo & "'"
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



Public Function Leer(vCodigo As String) As Boolean
On Error GoTo Err3
    
    Sql = "SELECT * FROM usuarios.stipom WHERE codtipom='" & vCodigo & "'"
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockOptimistic

    If Rs.EOF Then
        Leer = False
        'No existe el tipo de Movimiento
        Sql = "No existe el Tipo de Movimiento: " & vCodigo & " en la tabla de parámetros stipom."
        Sql = Sql & vbCrLf & "Consulte al Administrador."
        MsgBox Sql, vbCritical
    Else
        mvarTipoMov = Rs!CodTipom
        mvarNombreMov = Rs!nomtipom
        mvarMueveStock = Rs!muevesto
        mvarContador = Rs!Contador
        mvarLetraSerie = DBLet(Rs!letraser, "T")
        mvarTipoDocu = DBLet(Rs!TipoDocu, "N")
        Leer = True
    End If

    Rs.Close
    Set Rs = Nothing
    Exit Function

Err3:
    If Err.Number <> 0 Then
        Mens = "Se ha producido un error." & vbCrLf
        Mens = Mens & "Número: " & Err.Number & vbCrLf
        Mens = Mens & "Descripción: " & Err.Description
        MsgBox Mens, vbExclamation
        Rs.Close
        Leer = False
    End If
End Function



