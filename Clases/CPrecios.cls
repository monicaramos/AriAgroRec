VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPrecios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'variables locales que contienen valores de propiedad

Private mvarTipoFact As Byte    ' tipo: 0=anticipo 1=liquidacion
Private mvarVariedad As String  ' variedad
Private mvarContador As Long    ' contador
Private mvarFechaDesde As Date  ' fecha desde
Private mvarFechaHasta As Date  ' fecha hasta
Private mvarTextoPer As String  ' texto

Private mvarPrecioSocio As Currency    'Precio socio
Private mvarPrecioCoop As Currency     'Precio cooperativa

Private RS As ADODB.Recordset
Private SQL As String
Private Mens As String


Public Property Let TipoFactura(ByVal vData As Byte)
'se usa cuando se asigna un valor a una propiedad, en el lado izquierdo de la asignación.
'Syntax: X.Nivel = 5
    mvarTipoFact = vData
End Property


Public Property Get TipoFactura() As Byte
'se usa cuando se asigna un valor a una propiedad, en el lado derecho de la asignación.
'Syntax: Debug.Print X.Nivel
    TipoFactura = mvarTipoFact
End Property

Public Property Let Variedad(ByVal vData As String)
'se usa cuando se asigna un valor a una propiedad, en el lado izquierdo de la asignación.
'Syntax: X.Nombre= 5
    mvarVariedad = vData
End Property


Public Property Get Variedad() As String
'se usa cuando se asigna un valor a una propiedad, en el lado derecho de la asignación.
'Syntax: Debug.Print X.Nombre
    Variedad = mvarVariedad
End Property

Public Property Let Contador(ByVal vData As Long)
'se usa cuando se asigna un valor a una propiedad, en el lado izquierdo de la asignación.
'Syntax: X.Nivel = 5
    mvarContador = vData
End Property

Public Property Get Contador() As Long
'se usa cuando se asigna un valor a una propiedad, en el lado derecho de la asignación.
'Syntax: Debug.Print X.Nivel
    Contador = mvarContador
End Property

Public Property Let FechaDesde(ByVal vData As Date)
     mvarFechaDesde = vData
End Property

Public Property Get FechaDesde() As Date
     FechaDesde = mvarFechaDesde
End Property

Public Property Let FechaHasta(ByVal vData As Date)
     mvarFechaHasta = vData
End Property

Public Property Get FechaHasta() As Date
     FechaHasta = mvarFechaHasta
End Property


Public Property Let PrecioSocio(ByVal vData As Currency)
'se usa cuando se asigna un valor a una propiedad, en el lado izquierdo de la asignación.
'Syntax: X.Nivel = 5
    mvarPrecioSocio = vData
End Property


Public Property Get PrecioSocio() As Currency
'se usa cuando se asigna un valor a una propiedad, en el lado derecho de la asignación.
'Syntax: Debug.Print X.Nivel
    PrecioSocio = mvarPrecioSocio
End Property


Public Property Let PrecioCoop(ByVal vData As Currency)
'se usa cuando se asigna un valor a una propiedad, en el lado izquierdo de la asignación.
'Syntax: X.Nivel = 5
    mvarPrecioCoop = vData
End Property


Public Property Get PrecioCoop() As Currency
'se usa cuando se asigna un valor a una propiedad, en el lado derecho de la asignación.
'Syntax: Debug.Print X.Nivel
    PrecioCoop = mvarPrecioCoop
End Property



Public Function Leer(tipo As String, Variedad As String, FecDesde As String, FecHasta As String) As Boolean
On Error GoTo Err3
    
    SQL = "select * from rprecios where (codvarie, tipofact, contador) = ("
    SQL = SQL & "SELECT codvarie, tipofact, max(contador) FROM rprecios WHERE codvarie=" & DBSet(Variedad, "N") & " and "
    SQL = SQL & " tipofact = " & tipo & " and fechaini = " & DBSet(FecDesde, "F")
    SQL = SQL & " and fechafin = " & DBSet(FecHasta, "F")
    SQL = SQL & " group by 1, 2) "
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockOptimistic

    If RS.EOF Then
        Leer = False
        'No existe el tipo de Movimiento
        SQL = "No existe el Precio para la Variedad " & Variedad & " entre estas Fechas en la tabla de precios."
        SQL = SQL & vbCrLf
        MsgBox SQL, vbCritical
    Else
        mvarTipoFact = RS!TipoFact    ' tipo: 0=anticipo 1=liquidacion
        mvarVariedad = RS!CodVarie    ' variedad
        mvarContador = RS!Contador    ' contador
        mvarFechaDesde = RS!FechaIni  ' fecha desde
        mvarFechaHasta = RS!FechaFin  ' fecha hasta
        mvarTextoPer = RS!textoper    ' texto
        
        Leer = True
    End If

    RS.Close
    Set RS = Nothing
    Exit Function

Err3:
    If Err.Number <> 0 Then
        Mens = "Se ha producido un error." & vbCrLf
        Mens = Mens & "Número: " & Err.Number & vbCrLf
        Mens = Mens & "Descripción: " & Err.Description
        MsgBox Mens, vbExclamation
        RS.Close
        Leer = False
    End If
End Function



Public Function Modificar() As Byte
 
    Exit Function
Err1:
    Mens = "Error: " & Err.Number & " : " & Err.Description
    MsgBox Mens, vbExclamation
    RS.Close
    Modificar = 1
End Function


Public Function ExistenPreciosCalidades() As Boolean
Dim SQL As String
Dim Sql1 As String
Dim b As Boolean

    ExistenPreciosCalidades = False
 
    SQL = "select codcalid, nomcalid from rcalidad where codvarie = " & mvarVariedad
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    b = True
    While Not RS.EOF And b
        Sql1 = "select count(*) from rprecios_calidad where codvarie =" & mvarVariedad
        Sql1 = Sql1 & " and tipofact = " & mvarTipoFact & " and contador = " & mvarContador
        Sql1 = Sql1 & " and codcalid = " & DBLet(RS!codcalid, "N")
        
        If TotalRegistros(Sql1) = 0 Then
            b = False
            MsgBox "La calidad " & RS.Fields(1).Value & " de la variedad " & mvarVariedad & " no tiene asignado registro de precios. Revise.", vbExclamation
        End If
    
        RS.MoveNext
    Wend
    
    Set RS = Nothing
    ExistenPreciosCalidades = b
    
    Exit Function
Err1:
    Mens = "Error: " & Err.Number & " : " & Err.Description
    MsgBox Mens, vbExclamation
    RS.Close
End Function


Public Function PreciosCalidad(Calidad As String) As Boolean
Dim SQL As String

    PreciosCalidad = False
 
    SQL = "select precoop, presocio from rprecios_calidad where codvarie = " & mvarVariedad
    SQL = SQL & " and tipofact = " & mvarTipoFact & " and contador = " & mvarContador
    SQL = SQL & " and codcalid = " & DBSet(Calidad, "N")
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    If RS.EOF Then
        PreciosCalidad = False
        'No existe el precio para la calidad
        SQL = "No existe el Registros de Precios para la Calidad."
        SQL = SQL & vbCrLf
        MsgBox SQL, vbCritical
    Else
        mvarPrecioSocio = DBLet(RS!presocio, "N")   'Precio socio
        mvarPrecioCoop = DBLet(RS!precoop, "N")   'Precio cooperativa
        
        PreciosCalidad = True
    End If

    Exit Function
Err1:
    Mens = "Error: " & Err.Number & " : " & Err.Description
    MsgBox Mens, vbExclamation
    RS.Close
End Function



Public Function ExisteAlgunPrecioCalidad(SQL As String, TipoPrecio As Byte) As Boolean
' TipoPrecio: 0 = precio recolectado cooperativa
'             1 = precio recolectado socio
'             2 = precio recolectado cooperativa o socio

Dim Sql1 As String
Dim b As Boolean

    ExisteAlgunPrecioCalidad = False
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
 
    b = False
    While Not RS.EOF And Not b
        Sql1 = "select count(*) from rprecios_calidad where codvarie =" & mvarVariedad
        Sql1 = Sql1 & " and tipofact = " & mvarTipoFact & " and contador = " & mvarContador
        Sql1 = Sql1 & " and codcalid = " & DBLet(RS!codcalid, "N")
'++monica:28/07/2009
        Select Case TipoPrecio
            Case 0
                Sql1 = Sql1 & " and (rprecios_calidad.precoop <> 0 and not rprecios_calidad.precoop is null)"
            Case 1
                Sql1 = Sql1 & " and (rprecios_calidad.presocio <> 0 and not rprecios_calidad.presocio is null)"
            Case 2
                Sql1 = Sql1 & " and ((rprecios_calidad.precoop <> 0 and not rprecios_calidad.precoop is null) or (rprecios_calidad.presocio <> 0 and not rprecios_calidad.precoop is null)) "
        End Select
'++
        If TotalRegistros(Sql1) > 0 Then
            b = True
        End If
    
        RS.MoveNext
    Wend
    
    Set RS = Nothing
    ExisteAlgunPrecioCalidad = b
    
    Exit Function
Err1:
    Mens = "Error: " & Err.Number & " : " & Err.Description
    MsgBox Mens, vbExclamation
    RS.Close
End Function

