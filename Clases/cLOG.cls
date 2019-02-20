VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cLOG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'ACCIONES DISPONIBLES
'
'       1 .- Eliminar facturas
'       2 .- Facturacion
'       3 .- Contabilizar facturas
'       4 .- Correccion precios
'       5 .- Facturacion desde hasta COMPRAS
'       6 .- Abocamiento manual de trazabilidad de una entrada sin abocar
'       7 .- Lectura de contadores de pozos de Escalona
'       8 .- Facturacion Talla de pozos de Escalona
'       9 .- Modificacion Campos
'       10.- Insertar contador pozos
'       11.- Modificar contador pozos
'       12.- Modificar factura de socios
'       13.- Modificar factura varias
'       14.- Modificacion registros BD
'       15.- Insertamos los albaranes en las facturas de terceros
'       16.- Eliminar volcado de traza
'       17.- Insertar facturas
'       18.- Modificar facturas

Public Function Insertar(Accion As Byte, ByRef ElUsuario As Usuario, Descripcion As String) As Boolean
Dim Sql As String

    On Error GoTo EI
    Insertar = False
    
    Sql = "insert into `slog` (`fecha`,`accion`,`usuario`,`pc`,`descripcion`,`procedencia`) values ( "
    Sql = Sql & " now()," & Accion & " ,'" & DevNombreSQL(ElUsuario.Login) & "','"
    Sql = Sql & DevNombreSQL(ElUsuario.PC) & "',"
    If Descripcion = "" Then
        Sql = Sql & "NULL,1)"
    Else
        Sql = Sql & "'" & DevNombreSQL(Descripcion) & "',1)"
    End If
    conn.Execute Sql
    
    
    Insertar = True
    Exit Function
EI:
    MuestraError Err.Number, Err.Description

End Function


Public Function DevuelveAcciones(ByRef ListaAcciones As Collection) As Boolean
 On Error GoTo EDevuelveAcciones
    ListaAcciones.Add "1|Eliminar facturas|"
    ListaAcciones.Add "2|Facturación|"
    ListaAcciones.Add "3|Contabilizar facturas|"
    ListaAcciones.Add "4|Corrección precios|"
    ListaAcciones.Add "5|Facturacion COMPRAS|"
    ListaAcciones.Add "6|Abocamiento Manual TRAZA|"
    ListaAcciones.Add "7|Lectura Contadores|"
    ListaAcciones.Add "8|Facturación Talla|"
    ListaAcciones.Add "9|Modificación Campos|"
    ListaAcciones.Add "10|Insertar Contador|"
    ListaAcciones.Add "11|Modificar Contador|"
    ListaAcciones.Add "12|Modificar Factura|"
    ListaAcciones.Add "13|Modificar Fra.Varias|"
    ListaAcciones.Add "14|Modificar Registros BD|"
    ListaAcciones.Add "15|Insertar Albaranes|"
    ListaAcciones.Add "16|Eliminar Volcado|"
    ListaAcciones.Add "17|Insertar Facturas|"
    ListaAcciones.Add "18|Modificar Facturas|"
    DevuelveAcciones = True
    Exit Function
EDevuelveAcciones:
    MuestraError Err.Number
End Function

Public Sub InsertarCambiosRegistros(Accion As Byte, ByRef ElUsuario As Usuario, Descripcion As String, tabla As String, CP As String)
Dim Sql As String
Dim Sql2 As String

    '[Monica]15/11/2013: meto el replace
    Sql = Replace(CadenaCambio, "\'", " ")

    Sql2 = "insert into slog (fecha, accion, usuario, pc, descripcion, procedencia, tabla, cp, cadena, valorant) values ("
    Sql2 = Sql2 & DBSet(Now, "FH") & "," & Accion & " ,'" & DevNombreSQL(ElUsuario.Login) & "','"
    Sql2 = Sql2 & DevNombreSQL(ElUsuario.PC) & "',"
    If Descripcion = "" Then
        Sql2 = Sql2 & "NULL,1,"
    Else
        Sql2 = Sql2 & "'" & DevNombreSQL(Descripcion) & "',1,"
    End If
    
    Sql2 = Sql2 & DBSet(tabla, "T") & ","
    Sql2 = Sql2 & DBSet(CP, "T") & ","
    Sql2 = Sql2 & DBSet(Sql, "T") & ","
    If ValorAnterior = ValorNulo Then
        Sql2 = Sql2 & ValorNulo & ")"
    Else
        '[Monica]15/11/2013: meto el replace
        Sql2 = Sql2 & DBSet(Replace(ValorAnterior, "\'", " "), "T", "S") & ")"
    End If

    conn.Execute Sql2

End Sub
