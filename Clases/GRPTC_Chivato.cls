VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GRPTC_Chivato"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Este es un objeto parta manejar las entradas en chivato"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'variables locales para almacenar los valores de las propiedades
Private mvarId As Long 'copia local
Private mvarBD_Org As String 'copia local
Private mvarTabla As String 'copia local
Private mvarOper As String 'copia local
Private mvarFecha As Date 'copia local
Private mvarSep As String 'copia local
Private mvarClv_Old As String 'copia local
Private mvarClv_New As String 'copia local
'variables locales para almacenar los valores de las propiedades
Private mvarXML As String 'copia local

Public Property Let XML(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.XML = 5
    mvarXML = vData
End Property


Public Property Get XML() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.XML
    XML = mvarXML
End Property


Public Property Let Clv_New(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.Clv_New = 5
    mvarClv_New = vData
End Property


Public Property Get Clv_New() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.Clv_New
    Clv_New = mvarClv_New
End Property



Public Property Let Clv_Old(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.Clv_Old = 5
    mvarClv_Old = vData
End Property


Public Property Get Clv_Old() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.Clv_Old
    Clv_Old = mvarClv_Old
End Property



Public Property Let Sep(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.Sep = 5
    mvarSep = vData
End Property


Public Property Get Sep() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.Sep
    Sep = mvarSep
End Property



Public Property Let fecha(ByVal vData As Date)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.Fecha = 5
    mvarFecha = vData
End Property


Public Property Get fecha() As Date
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.Fecha
    fecha = mvarFecha
End Property



Public Property Let Oper(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.Oper = 5
    mvarOper = vData
End Property


Public Property Get Oper() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.Oper
    Oper = mvarOper
End Property



Public Property Let Tabla(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.Tabla = 5
    mvarTabla = vData
End Property


Public Property Get Tabla() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.Tabla
    Tabla = mvarTabla
End Property



Public Property Let BD_Org(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.BD_Org = 5
    mvarBD_Org = vData
End Property


Public Property Get BD_Org() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.BD_Org
    BD_Org = mvarBD_Org
End Property



Public Property Let Id(ByVal vData As Long)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.Id = 5
    mvarId = vData
End Property


Public Property Get Id() As Long
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.Id
    Id = mvarId
End Property



Private Function Siguiente_NumOrden() As Long
    Dim sql As String
    Dim rs As ADODB.Recordset
    sql = "SELECT MAX(numorden) from chivato"
    Set rs = dbAriagro.cursor(sql)
    If (Not IsNull(rs.Fields(0))) And (Not rs.EOF) Then
        Siguiente_NumOrden = rs.Fields(0) + 1
    Else
        Siguiente_NumOrden = 1
    End If
    rs.Close
    Set rs = Nothing
End Function


Public Function Grabar() As Boolean
    Dim sql As String
    Dim nombmemo(4) As String
    Dim i As Integer
    Dim nc As Integer
    Dim resto As Integer
    Dim TamCampo As Integer
    Dim bloqueado As Boolean
    TamCampo = 254
    nc = Int(Len(mvarXML) / TamCampo)
    resto = Len(mvarXML) Mod TamCampo
    For i = 0 To nc - 1 '-- Los que ocupan 254
        nombmemo(i) = Mid(mvarXML, (TamCampo * i) + 1, 254)
    Next i
    If resto <> 0 Then _
        nombmemo(nc) = Mid(mvarXML, (TamCampo * nc) + 1, resto) '-- Si hay resto lo ponemos
    '-- Graba los datos en la tabla que corresponde
    On Error GoTo Err_Grabar
    '-- Bloqueamos la tabla chivato (OJO HAY QUE HACERLO PARA MYSQL)
    bloqueado = True
    '--- Buscamos un número de orden libre
    mvarId = Siguiente_NumOrden
    sql = "INSERT INTO chivato(nombmemo,nomtabla,operacio,fechadia,separado,claveant,clavenue,basedato,nombmem1,nombmem2,horaproc,numorden,nombmem3,nombmem4)"
    sql = sql & " VALUES("
    sql = sql & "'" & CApos(nombmemo(0)) & "'," 'nombmemo
    sql = sql & "'" & mvarTabla & "'," 'nomtabla
    sql = sql & "'" & mvarOper & "'," ' operacio
    sql = sql & "'" & Format(Date, "yyyy-mm-dd") & "'," ' fechadia
    sql = sql & "'" & mvarSep & "'," ' separado
    sql = sql & "'" & mvarClv_Old & "'," ' claveant
    sql = sql & "'" & mvarClv_New & "'," ' clavenue
    sql = sql & "'" & mvarBD_Org & "'," ' basedato
    sql = sql & "'" & CApos(nombmemo(1)) & "'," ' nombmem1
    sql = sql & "'" & CApos(nombmemo(2)) & "'," ' nombmem2
    sql = sql & "'" & Format(Now, "hh:mm:ss") & "'," ' horaproc
    sql = sql & CStr(mvarId) & "," ' numorden
    sql = sql & "'" & CApos(nombmemo(3)) & "'," ' nombmem3
    sql = sql & "'" & CApos(nombmemo(4)) & "')" ' nombmem4
    dbAriagro.ejecutar sql
    '-- Y aora rabamos la tabala
    Grabar = True ' todo ha ido bien (increible pero cierto)
    '-- Desbloqueamos la tabla
    If bloqueado Then
        '-- (OJO HABRIA QUE DESBLOQUEAR REALMENTE)
        bloqueado = False
    End If
    Exit Function
Err_Grabar:
    MsgBox "Err: " & Err.Number & vbCrLf & Err.Description, vbCritical
    If bloqueado Then
        '-- (OJO HABRIA QUE DESBLOQUEAR REALMENTE)
        bloqueado = False
    End If
End Function

