VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ArdImagen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'-- ArdImagen:
'   Esta clase maneja objetos que representan registros de la tabla
'   timagen en AriDoc (en realidad documentos guardados en esa base de datos)
'   Esta clase presupone que existe un objeto de nombre ardDB que soporta el acceso a la base
'   de datos

'variables locales para almacenar los valores de las propiedades
Private mvarcodigo As Long 'copia local
Private mvarcodext As Integer 'copia local
Private mvarcodcarpeta As Integer 'copia local
Private mvarcampo1 As String 'copia local
Private mvarcampo2 As String 'copia local
Private mvarcampo3 As String 'copia local
Private mvarcampo4 As String 'copia local
Private mvarfecha1 As Date 'copia local
Private mvarfecha2 As Date 'copia local
Private mvarfecha3 As Date 'copia local
Private mvarimporte1 As Currency 'copia local
Private mvarimporte2 As Currency 'copia local
Private mvarobserva As String 'copia local
Private mvartamnyo As Single 'copia local
Private mvaruserprop As Long 'copia local
Private mvargroupprop As Long 'copia local
Private mvarlecturau As Long 'copia local
Private mvarlecturag As Long 'copia local
Private mvarescriturau As Variant 'copia local
Private mvarescriturag As Long 'copia local
Private mvarbloqueo As Integer 'copia local

Public Property Let bloqueo(ByVal vData As Integer)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.bloqueo = 5
    mvarbloqueo = vData
End Property


Public Property Get bloqueo() As Integer
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.bloqueo
    bloqueo = mvarbloqueo
End Property



Public Property Let escriturag(ByVal vData As Long)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.escriturag = 5
    mvarescriturag = vData
End Property


Public Property Get escriturag() As Long
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.escriturag
    escriturag = mvarescriturag
End Property



Public Property Let escriturau(ByVal vData As Variant)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.escriturau = 5
    mvarescriturau = vData
End Property


Public Property Set escriturau(ByVal vData As Variant)
'se usa al asignar un objeto a la propiedad, en la parte izquierda de una instrucción Set.
'Syntax: Set x.escriturau = Form1
    Set mvarescriturau = vData
End Property


Public Property Get escriturau() As Variant
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.escriturau
    If IsObject(mvarescriturau) Then
        Set escriturau = mvarescriturau
    Else
        escriturau = mvarescriturau
    End If
End Property



Public Property Let lecturag(ByVal vData As Long)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.lecturag = 5
    mvarlecturag = vData
End Property


Public Property Get lecturag() As Long
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.lecturag
    lecturag = mvarlecturag
End Property



Public Property Let lecturau(ByVal vData As Long)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.lecturau = 5
    mvarlecturau = vData
End Property


Public Property Get lecturau() As Long
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.lecturau
    lecturau = mvarlecturau
End Property



Public Property Let groupprop(ByVal vData As Long)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.groupprop = 5
    mvargroupprop = vData
End Property


Public Property Get groupprop() As Long
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.groupprop
    groupprop = mvargroupprop
End Property



Public Property Let userprop(ByVal vData As Long)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.userprop = 5
    mvaruserprop = vData
End Property


Public Property Get userprop() As Long
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.userprop
    userprop = mvaruserprop
End Property



Public Property Let tamnyo(ByVal vData As Single)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.tamnyo = 5
    mvartamnyo = vData
End Property


Public Property Get tamnyo() As Single
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.tamnyo
    tamnyo = mvartamnyo
End Property



Public Property Let observa(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.observa = 5
    mvarobserva = vData
End Property


Public Property Get observa() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.observa
    observa = mvarobserva
End Property



Public Property Let importe2(ByVal vData As Currency)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.importe2 = 5
    mvarimporte2 = vData
End Property


Public Property Get importe2() As Currency
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.importe2
    importe2 = mvarimporte2
End Property



Public Property Let importe1(ByVal vData As Currency)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.importe1 = 5
    mvarimporte1 = vData
End Property


Public Property Get importe1() As Currency
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.importe1
    importe1 = mvarimporte1
End Property



Public Property Let fecha3(ByVal vData As Date)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.fecha3 = 5
    mvarfecha3 = vData
End Property


Public Property Get fecha3() As Date
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.fecha3
    fecha3 = mvarfecha3
End Property



Public Property Let fecha2(ByVal vData As Date)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.fecha2 = 5
    mvarfecha2 = vData
End Property


Public Property Get fecha2() As Date
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.fecha2
    fecha2 = mvarfecha2
End Property



Public Property Let fecha1(ByVal vData As Date)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.fecha1 = 5
    mvarfecha1 = vData
End Property


Public Property Get fecha1() As Date
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.fecha1
    fecha1 = mvarfecha1
End Property



Public Property Let campo4(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.campo4 = 5
    mvarcampo4 = vData
End Property


Public Property Get campo4() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.campo4
    campo4 = mvarcampo4
End Property



Public Property Let campo3(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.campo3 = 5
    mvarcampo3 = vData
End Property


Public Property Get campo3() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.campo3
    campo3 = mvarcampo3
End Property



Public Property Let campo2(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.campo2 = 5
    mvarcampo2 = vData
End Property


Public Property Get campo2() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.campo2
    campo2 = mvarcampo2
End Property



Public Property Let campo1(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.campo1 = 5
    mvarcampo1 = vData
End Property


Public Property Get campo1() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.campo1
    campo1 = mvarcampo1
End Property



Public Property Let codcarpeta(ByVal vData As Integer)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.codcarpeta = 5
    mvarcodcarpeta = vData
End Property


Public Property Get codcarpeta() As Integer
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.codcarpeta
    codcarpeta = mvarcodcarpeta
End Property



Public Property Let codext(ByVal vData As Integer)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.codext = 5
    mvarcodext = vData
End Property


Public Property Get codext() As Integer
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.codext
    codext = mvarcodext
End Property



Public Property Let codigo(ByVal vData As Long)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.codigo = 5
    mvarcodigo = vData
End Property


Public Property Get codigo() As Long
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.codigo
    codigo = mvarcodigo
End Property

Private Function siguienteCodigo() As Long
    '-- siguienteCodigo()
    '   Devuelve el valor del siguiete código disponible en la tabla timagen
    Dim SQL As String
    Dim RS As ADODB.Recordset
    On Error GoTo err_siguienteCodigo:
    SQL = "select max(codigo) from timagen"
    Set RS = ardDB.cursor(SQL)
    siguienteCodigo = 1
    If Not RS.EOF Then
        If Not IsNull(RS.Fields(0)) Then
            siguienteCodigo = RS.Fields(0) + 1
        End If
    End If
    Exit Function
err_siguienteCodigo:
    MsgBox Err.Description & " (" & Err.Number & ")", vbCritical, "ArdImagen.siguienteCodigo"
End Function

Public Function guardar() As Boolean
    Dim SQL As String
    Dim RS As ADODB.Recordset
On Error GoTo err_guardar
    guardar = False
    '-- 1º comprobamos si existe previamente el registro
    SQL = "select * from timagen where codigo = " & ardDB.numero(codigo)
    Set RS = ardDB.cursor(SQL)
    If RS.EOF Then
        '-- No existe y lo damos de alta
        codigo = siguienteCodigo
        SQL = "insert into timagen(" & _
            "codigo, codext, codcarpeta, campo1, campo2, campo3, campo4," & _
            "fecha1, fecha2, fecha3, importe1, importe2, observa, tamnyo," & _
            "userprop, groupprop, lecturau, lecturag, escriturau, escriturag," & _
            "bloqueo) values ("
        SQL = SQL & ardDB.numero(codigo) & "," 'toma el valor del siguiente código disponible
        SQL = SQL & ardDB.numero(codext) & ","
        SQL = SQL & ardDB.numero(codcarpeta) & ","
        SQL = SQL & ardDB.Texto(CApos(campo1)) & ","
        SQL = SQL & ardDB.Texto(CApos(campo2)) & ","
        SQL = SQL & ardDB.Texto(CApos(campo3)) & ","
        SQL = SQL & ardDB.Texto(CApos(campo4)) & ","
        SQL = SQL & ardDB.Fecha(fecha1) & ","
        If fecha2 = "0:00:00" Then
            SQL = SQL & ValorNulo & ","
        Else
            SQL = SQL & ardDB.Fecha(fecha2) & ","
        End If
        If fecha3 = "0:00:00" Then
            SQL = SQL & ValorNulo & ","
        Else
            SQL = SQL & ardDB.Fecha(fecha3) & ","
        End If
        
        If importe1 = 0 Then
            SQL = SQL & ValorNulo & ","
        Else
            SQL = SQL & ardDB.numero(importe1) & ","
        End If
        If importe2 = 0 Then
            SQL = SQL & ValorNulo & ","
        Else
            SQL = SQL & ardDB.numero(importe2) & ","
        End If
        
        SQL = SQL & ardDB.Texto(observa) & ","
        SQL = SQL & ardDB.numero(tamnyo) & ","
        SQL = SQL & ardDB.numero(userprop) & ","
        SQL = SQL & ardDB.numero(groupprop) & ","
        SQL = SQL & ardDB.numero(lecturau) & ","
        SQL = SQL & ardDB.numero(lecturag) & ","
        SQL = SQL & ardDB.numero(escriturau) & ","
        SQL = SQL & ardDB.numero(escriturag) & ","
        SQL = SQL & ardDB.numero(bloqueo) & ")"

'--monica: no tiene marca de pasaridoc
'    Else
'        '-- Ya existe y lo modificamos
'        SQL = "update timagen set "
'        SQL = SQL & "codext = " & ardDB.numero(codext) & ","
'        SQL = SQL & "codcarpeta = " & ardDB.numero(codcarpeta) & ","
'        SQL = SQL & "campo1 = " & ardDB.Texto(CApos(campo1)) & ","
'        SQL = SQL & "campo2 = " & ardDB.Texto(CApos(campo2)) & ","
'        SQL = SQL & "campo3 = " & ardDB.Texto(CApos(campo3)) & ","
'        SQL = SQL & "campo4 = " & ardDB.Texto(CApos(campo4)) & ","
'        SQL = SQL & "fecha1 = " & ardDB.Fecha(fecha1) & ","
'        If fecha2 = "0:00:00" Then
'            SQL = SQL & "fecha2 = " & ValorNulo & ","
'        Else
'            SQL = SQL & "fecha2 = " & ardDB.Fecha(fecha2) & ","
'        End If
'        If fecha3 = "0:00:00" Then
'            SQL = SQL & "fecha3 = " & ValorNulo & ","
'        Else
'            SQL = SQL & "fecha3 = " & ardDB.Fecha(fecha3) & ","
'        End If
'        If importe1 = "0" Then
'            SQL = SQL & "importe1 = " & ValorNulo & ","
'        Else
'            SQL = SQL & "importe1 = " & ardDB.numero(importe1) & ","
'        End If
'        If importe2 = "0" Then
'            SQL = SQL & "importe2 = " & ValorNulo & ","
'        Else
'            SQL = SQL & "importe2 = " & ardDB.numero(importe2) & ","
'        End If
'        SQL = SQL & "observa = " & ardDB.Texto(observa) & ","
'        SQL = SQL & "tamnyo = " & ardDB.numero(tamnyo) & ","
'        SQL = SQL & "userprop = " & ardDB.numero(userprop) & ","
'        SQL = SQL & "groupprop = " & ardDB.numero(groupprop) & ","
'        SQL = SQL & "lecturau = " & ardDB.numero(lecturau) & ","
'        SQL = SQL & "lecturag = " & ardDB.numero(lecturag) & ","
'        SQL = SQL & "escriturau = " & ardDB.numero(escriturau) & ","
'        SQL = SQL & "escriturag = " & ardDB.numero(escriturag) & ","
'        SQL = SQL & "bloqueo = " & ardDB.numero(bloqueo)
'        SQL = SQL & " where codigo = " & ardDB.numero(codigo)
    End If
    ardDB.ejecutar SQL
    guardar = True
    Exit Function
err_guardar:
    MsgBox Err.Description & " (" & Err.Number & ")", vbCritical, "ArdImagen.guardar"
End Function

Public Function leer(vcodigo As Long) As Boolean
    Dim SQL As String
    Dim RS As ADODB.Recordset
On Error GoTo err_leer
    leer = False
    SQL = "select * from timagen where codigo = " & ardDB.numero(vcodigo)
    Set RS = ardDB.cursor(SQL)
    If Not RS.EOF Then
        codigo = DBLet(RS!codigo, "N")
        codext = DBLet(RS!codext, "N")
        codcarpeta = DBLet(RS!codcarpeta, "N")
        campo1 = DBLet(RS!campo1)
        campo2 = DBLet(RS!campo2)
        campo3 = DBLet(RS!campo3)
        campo4 = DBLet(RS!campo4)
        
        '++monica : añadida la condicion y el else
        If Not IsNull(RS!fecha1) Then
            fecha1 = DBLet(RS!fecha1, "F")
        End If
        '++monica : añadida la condicion y el else
        If Not IsNull(RS!fecha2) Then
            fecha2 = DBLet(RS!fecha2, "F")
        End If
        '++monica : añadida la condicion y el else
        If Not IsNull(RS!fecha3) Then
            fecha3 = DBLet(RS!fecha3, "F")
        End If
        
        importe1 = DBLet(importe1, "N")
        importe2 = DBLet(importe2, "N")
        observa = DBLet(RS!observa)
        tamnyo = DBLet(RS!tamnyo, "N")
        userprop = DBLet(RS!userprop, "N")
        groupprop = DBLet(RS!groupprop, "N")
        lecturau = DBLet(RS!lecturau, "N")
        lecturag = DBLet(RS!lecturag, "N")
        escriturau = DBLet(RS!escriturau, "N")
        escriturag = DBLet(RS!escriturag, "N")
        bloqueo = DBLet(RS!bloqueo, "N")
        leer = True
    Else
        leer = False
    End If
    Exit Function
err_leer:
    MsgBox Err.Description & " (" & Err.Number & ")", vbCritical, "ArdImagen.leer"
End Function

Public Function eliminar() As Boolean
    Dim SQL As String
    Dim RS As ADODB.Recordset
On Error GoTo err_eliminar
    eliminar = False
    
    SQL = "delete from timagen where codigo = " & ardDB.numero(codigo)
    ardDB.ejecutar SQL
    
    eliminar = True
    Exit Function
err_eliminar:
    MsgBox Err.Description & " (" & Err.Number & ")", vbCritical, "ArdImagen.eliminar"
End Function

Private Sub Class_Initialize()
    '-- Cargamos una serie de valores por defecto para cualquier registro
    '   esto debería revisarse en un futuro para cambiar por lo correcto
    userprop = 0
    groupprop = 1
    lecturau = 2147483647
    lecturag = 2147483647
    escriturau = 2147483647
    escriturag = 2147483647
    bloqueo = 0
End Sub
