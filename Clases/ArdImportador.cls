VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ArdImportador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'-- ArdImportador:
'   Incluye lo necesario para importar documententos a Aridoc
'   presuòne la existencia de un objeto ardDB de la clase BaseDatos

'variables locales para almacenar los valores de las propiedades
Private mArdImagen As ArdImagen 'copia local


Public Function nombreCarpeta(codcarpeta As Long) As String
    '-- nombreCarpeta:
    '   Devuelve el nombre de la carpeta relacionada por su código
    '   Si la carpeta no existe devuelve ""
    Dim SQL As String
    Dim RS As ADODB.Recordset
On Error GoTo err_nombreCarpeta
    nombreCarpeta = ""
    SQL = "select * from carpetas where codcarpeta = " & ardDB.numero(codcarpeta)
    Set RS = ardDB.cursor(SQL)
    If Not RS.EOF Then
        nombreCarpeta = DBLet(RS!Nombre)
    End If
    Exit Function
err_nombreCarpeta:
    MsgBox Err.Description & " (" & Err.Number & ")", vbCritical, "ArdImportador.nombreCarpeta"
End Function

Public Function rutaAGrabar(codcarpeta As Long) As String
    '-- rutaAGrabar:
    '   A partir de un código de carpeta suministrado devuelve la
    '   ruta del almacén en el que debe residir el documento
    Dim SQL As String
    Dim RS As ADODB.Recordset
On Error GoTo err_rutaAGrabar
    rutaAGrabar = ""
    SQL = "select * from carpetas where codcarpeta = " & ardDB.numero(codcarpeta)
    Set RS = ardDB.cursor(SQL)
    If Not RS.EOF Then
        SQL = "select * from almacen where codalma = " & ardDB.numero(RS!Almacen)
        Set RS = ardDB.cursor(SQL)
        If Not RS.EOF Then
            rutaAGrabar = RS!pathreal
        End If
    End If
    Exit Function
err_rutaAGrabar:
    MsgBox Err.Description & " (" & Err.Number & ")", vbCritical, "ArdImportador.rutaAGrabar"
End Function

Public Function importaFicheroPDF(fichero As String, codcarpeta As Long, c1 As String, c2 As String, c3 As String, c4 As String, f1 As Date, f3 As Date, i1 As Currency) As Boolean
    '-- importaFicheroPDF
    '   Importa el fichero pasado a la carpeta que se le diga de Aridoc, con los datos de
    '   registro adicionales de los que se trate.
    Dim SQL As String
    Dim RS As ADODB.Recordset
    Dim ruta As String
On Error GoTo err_importaFicheroPDF
    importaFicheroPDF = False
    '-- 1º obtenemos la ruta a grabar, si no hay ruta no hay nada que hacer
    ruta = rutaAGrabar(codcarpeta)
    If ruta <> "" Then
        '-- Hay ruta, hay esperanza
        Set mArdImagen = New ArdImagen
        '-- 2º comprobamos si esa factura en concreto ha sido importada con anterioridad
'--monica
'        SQL = "select * from timagen where campo2 = " & ardDB.Texto(c2)
'        Set RS = ardDB.cursor(SQL)
'        If Not RS.EOF Then mArdImagen.leer RS!Codigo ' si ya existe traemos sus datos
'--
        With mArdImagen
            .codext = vParamAplic.Extension  ' indica que es un PDF
            .codcarpeta = codcarpeta
            .campo1 = c1
            .campo2 = c2
            .campo3 = c3
            .campo4 = c4
            .fecha1 = f1
            .fecha3 = f3
            .Importe1 = i1
            .tamnyo = Round2(FileLen(fichero) / 1024, 2)
            .observa = "Importación automática de " & c2 '& " Tel:" & c3
            '++monica: solo si lo hemos podido guardar lo copiamos
'            If .guardar Then
            .guardar
                '-- Ya tenemos el registro y ahora copiamos el fichero al destino
             FileCopy fichero, ruta & "\" & CStr(.Codigo)
'            Else
'                Exit Function
'            End If
        End With
        
    End If
    importaFicheroPDF = True
    Exit Function
err_importaFicheroPDF:
    MsgBox Err.Description & " (" & Err.Number & ")", vbCritical, "ArdImportador.importaFicheroPDF"

End Function
