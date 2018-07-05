VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BaseDatos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'Option Explicit
'Private mtipo As String
'Private conn As ADODB.Connection
'Private mrs As ADODB.Recordset
'
'Public Property Let tipo(ByVal vData As String)
'    mtipo = vData
'End Property
'
'Public Property Get tipo() As String
'    tipo = mtipo
'End Property
'
'Public Property Let Con(ByVal vData As ADODB.Connection)
'    Set conn = vData
'End Property
'
'Public Property Get Con() As ADODB.Connection
'    Set Con = conn
'End Property
'
'
'Public Function abrir(vODBC As String, vUsu As String, vPass As String) As Boolean
'    '-- Abrimos la conexión en función del ODBC pasado
'    Set conn = New ADODB.Connection
'    conn.ConnectionString = "Provider=MSDASQL.1;Persist Security Info=False;" & _
'                            "Data Source=" & vODBC & _
'                            ";UID=" & vUsu & _
'                            ";PWD=" & vPass
'    conn.Open
'    abrir = True
'End Function
'
'Public Function abrir_MYSQL(vServidor As String, vBaseDatos As String, vUsu As String, vPass As String) As Boolean
'    '-- Abrimos la conexión en función de los datos pasados
'    Set conn = New ADODB.Connection
'    On Error GoTo abrir_MYSQL_error
'    conn.ConnectionString = "DRIVER={MySQL ODBC 3.51 Driver};DESC=" & _
'                        ";DATABASE=" & vBaseDatos & _
'                        ";SERVER=" & vServidor & _
'                        ";UID=" & vUsu & _
'                        ";PWD=" & vPass & _
'                        ";PORT=3306;OPTION=3;STMT="
'    conn.Open
'    mtipo = "MYSQL"
'    abrir_MYSQL = True
'    Exit Function
'abrir_MYSQL_error:
'    MsgBox Err.Description & " (" & Err.Number & ")", vbCritical, "BaseDatos.abrir_MYSQL"
'End Function
'
'Public Function ejecutar(vSQL As String) As Long
'' ### [Monica] 10/10/2006 añado la rutina de errores
'On Error GoTo eEjecutar
'
'    ejecutar = 0
'    If Not (conn Is Nothing) Then
'        conn.Execute vSQL
'    Else
'        '-- Falta la conexión y no podemos continuar
'    End If
'
'' ### [Monica] 10/10/2006 añado la rutina de errores
'eEjecutar:
'    If Err.Number <> 0 Then
'        ejecutar = Err.Number
''        MuestraError Err.Number, Err.Description
'    End If
'
'End Function
'
'Public Function ejecutar2(vSQL As String, ByRef menserr As String) As Long
'' ### [Monica] 10/10/2006 añado la rutina de errores
'On Error GoTo eEjecutar
'
'    ejecutar2 = 0
'    If Not (conn Is Nothing) Then
'        conn.Execute vSQL
'    Else
'        '-- Falta la conexión y no podemos continuar
'    End If
'
'' ### [Monica] 10/10/2006 añado la rutina de errores
'eEjecutar:
'    If Err.Number <> 0 Then
'        ejecutar2 = Err.Number
'        MuestraError Err.Number, menserr & " " & Err.Description
'    End If
'
'End Function
'
'Public Function cursor(vSQL As String) As ADODB.Recordset
'    '-- Devuelve un recordset con los valores obtenidos por SQL
'    Set mrs = New ADODB.Recordset
'    mrs.Open vSQL, conn, adOpenForwardOnly
'    Set cursor = mrs
'End Function
'
'Public Function fecha(vData) As String
'    '-- Devuelve un string con la fecha correspondiente
'    If vData = "0:00:00" Then
'        fecha = "NULL"
'    Else
'        If Not IsNull(vData) And IsDate(vData) Then
'            Select Case mtipo
'                Case "ACCESS"
'                    fecha = "#" & Format(vData, "yyyy/mm/dd") & "#"
'                Case "MYSQL"
'                    fecha = "'" & Format(vData, "yyyy-mm-dd") & "'"
'                Case "MULTIBASE"
'                    fecha = "'" & Format(vData, "dd/mm/yyyy") & "'"
'                Case Default
'                    fecha = "'" & Format(vData, "yyyy-mm-dd") & "'"
'            End Select
'        Else
'            fecha = "NULL"
'        End If
'    End If
'End Function
'
'Public Function Hora(vData) As String
'    '-- Devuelve un string con la fecha correspondiente
'    If vData = "0:00:00" Then
'        Hora = "NULL"
'    Else
'        If Not IsNull(vData) And IsDate(vData) Then
'            Select Case mtipo
'                Case "ACCESS"
'                    Hora = "#" & Format(vData, "hh:mm:ss") & "#"
'                Case "MYSQL"
'                    Hora = "'" & Format(vData, "hh:mm:ss") & "'"
'                Case "MULTIBASE"
'                    Hora = "'" & Format(vData, "hh:mm") & "'"
'                Case Default
'            End Select
'        Else
'            Hora = "NULL"
'        End If
'    End If
'End Function
'
'Public Function fechahora(vData) As String
'    '-- Devuelve un string con la fecha correspondiente
'    If vData = "00/00/0000 00:00:00" Then
'        fechahora = "NULL"
'    Else
'        If Not IsNull(vData) And IsDate(vData) Then
'            Select Case mtipo
'                Case "ACCESS"
'                    fechahora = "#" & Format(vData, "yyyy/mm/dd") & " " & Format(vData, "hh:mm:ss") & "#"
'                Case "MYSQL"
'                    fechahora = "'" & Format(vData, "yyyy-mm-dd") & " " & Format(vData, "hh:mm:ss") & "'"
'                Case "MULTIBASE"
'                    fechahora = "'" & Format(vData, "dd/mm/yyyy") & " " & Format(vData, "hh:mm:ss") & "'"
'                Case Default
'                    fechahora = "'" & Format(vData, "yyyy-mm-dd") & " " & Format(vData, "hh:mm:ss") & "'"
'            End Select
'        Else
'            fechahora = "NULL"
'        End If
'    End If
'End Function
'
'Public Function numero(vData, Optional decimales As Integer) As String
'    If Not IsNull(vData) Then
'        numero = CStr(vData)
'        If decimales Then
'            'numero = TransformaComasPuntos(CStr(vData / (10 ^ decimales)))
'            numero = TransformaComasPuntos(Format((vData / (10 ^ decimales)), "0." & String(decimales, "0")))
'        Else
'            numero = TransformaComasPuntos(CStr(vData))
'        End If
'    Else
'        'antes
'        'numero = "0"
'        'ahora
'        numero = ValorNulo
'    End If
'    If numero = "" Then numero = 0
'End Function
'
'Public Function Texto(vData) As String
'    If Not IsNull(vData) Then
'        Texto = "'" & vData & "'"
'    Else
'        Texto = "''"
'    End If
'End Function
'
'Public Function Borrar() As String
'    Select Case mtipo
'        Case "ACCESS"
'            Borrar = "delete * "
'        Case "MYSQL"
'            Borrar = "delete "
'        Case "MULTIBASE"
'            Borrar = "delete "
'        Case Default
'    End Select
'End Function
'
'' ### [Monica] 11/10/2006
'' añadida la parte de transacciones
'Public Sub AbrirTrans()
'
'    conn.BeginTrans
'
'End Sub
'
'Public Sub RollbackTrans()
'
'    conn.RollbackTrans
'
'End Sub
'
'Public Sub CommitTrans()
'
'    conn.CommitTrans
'
'End Sub
'

'++monica: he cogido la basedatos de rafa
Option Explicit
Private mtipo As String
Private conn As ADODB.Connection
Private mrs As ADODB.Recordset

Public Property Let Tipo(ByVal vData As String)
    mtipo = vData
End Property

Public Property Get Tipo() As String
    Tipo = mtipo
End Property


Public Function abrir(vODBC As String, vUsu As String, vPass As String) As Boolean
    '-- Abrimos la conexión en función del ODBC pasado
    Set conn = New ADODB.Connection
    conn.ConnectionString = "Provider=MSDASQL.1;Persist Security Info=False;" & _
                            "Data Source=" & vODBC
    conn.Open
End Function



Public Function abrir_MYSQL(vServidor As String, vBaseDatos As String, vUsu As String, vPass As String) As Boolean
    '-- Abrimos la conexión en función de los datos pasados
    Set conn = New ADODB.Connection
    On Error GoTo abrir_MYSQL_error
    conn.ConnectionString = "DRIVER={MySQL ODBC 3.51 Driver};DESC=" & _
                        ";DATA SOURCE=vAriagro" & _
                        ";DATABASE=" & vBaseDatos & _
                        ";SERVER=" & vServidor
    conn.Open
    mtipo = "MYSQL"
    abrir_MYSQL = True
    Exit Function
abrir_MYSQL_error:
    MsgBox Err.Description & " (" & Err.Number & ")", vbCritical, "BaseDatos.abrir_MYSQL"
End Function

Public Function abrirL(vUsu As String, vPass As String, vBaseDatos As String, vServer As String) As Boolean
    '-- Abrimos la conexión en función del ODBC pasado
    On Error GoTo err_abrirL
    Set conn = New ADODB.Connection
    conn.ConnectionString = "DRIVER={MySQL ODBC 3.51 Driver};DESC=;" _
                & ";DATA SOURCE=vAriagro" _
                & ";DATABASE=" & vBaseDatos & ";SERVER=" & vServer
                
                
    conn.Open
    abrirL = True
    Exit Function
err_abrirL:
    MsgBox Err.Description & "(Err:" & Err.Number & ")", vbCritical
    abrirL = False
End Function



Public Function ejecutar(vSQL As String) As Boolean
    If Not (conn Is Nothing) Then
        conn.Execute vSQL
    Else
        '-- Falta la conexión y no podemos continuar
    End If
End Function

Public Function cursor(vSQL As String) As ADODB.Recordset
    '-- Devuelve un recordset con los valores obtenidos por SQL
    Set mrs = New ADODB.Recordset
    mrs.Open vSQL, conn, adOpenForwardOnly
    Set cursor = mrs
End Function

Public Function Fecha(vData) As String
    '-- Devuelve un string con la fecha correspondiente
    If Not IsNull(vData) And IsDate(vData) Then
        Select Case mtipo
            Case "ACCESS"
                Fecha = "#" & Format(vData, "yyyy/mm/dd") & "#"
            Case "MYSQL"
                Fecha = "'" & Format(vData, "yyyy-mm-dd") & "'"
            Case "MULTIBASE"
                Fecha = "'" & Format(vData, "dd/mm/yyyy") & "'"
            Case Default
                Fecha = "'" & Format(vData, "yyyy-mm-dd") & "'"
        End Select
    Else
        Fecha = "NULL"
    End If
End Function

Public Function numero(vData, Optional decimales As Integer) As String
    If Not IsNull(vData) Then
        numero = CStr(vData)
        If decimales Then
            numero = TransformaComasPuntos(CStr(vData / (10 ^ decimales)))
        Else
            numero = TransformaComasPuntos(CStr(vData))
        End If
    Else
        numero = "0"
    End If
    If numero = "" Then numero = 0
End Function

Public Function Texto(vData) As String
    If Not IsNull(vData) Then
        Texto = "'" & vData & "'"
    Else
        Texto = "''"
    End If
End Function

