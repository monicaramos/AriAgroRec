VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CStockADV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'=======================================================================
'MODULO para el control de Stock en Almacenes
'Comprueba si hay stock de un Articulo en un Almacen
'Comprueba si hay control de stock para permitir movimiento o no
'Inserta el movimiento de Entrada/salida de stock en la tabla "advsmoval"
'Actualiza el valor del stock tras el movimiento en la tabla "salmac"
'=======================================================================


'------------------------------------------------
'Variables del modulo CStock
'------------------------------------------------
Private mvarCodArtic As String
Private mvarCodAlmac As Integer

Private mvarFechaMov As String
Private mvarHoraMov As String
Private mvarTipoMov As String '0=Salida, 1=Entrada
Private mvarDetaMov As String 'referencia a codtipom de la tabla stipom
Private mvarCantidad As Single 'Cantidad de Stock a mover
Private mvarImporte As Currency 'Importe de la linea de Movimiento de stock
Private mvarCodOpe As Long 'Codigo Trabajador/Cliente/Proveedor segun el tipo de Movimiento
Private mvarDocum As String 'Codigo de la tabla de Movimientos, Traspaso, Albaran,... al que va asociado el movimiento
Private mvarLinea As Integer 'Numero de linea del documento
Private mvarCampo As Long ' codigo de campo de la entrada


'variable local
Dim MuestraMenCStock  As Boolean 'Indica si se muestra el mensaje en el control de Stock
                             'Cuando se pasa de Pedido a Albaran si alguna linea no tiene
                             'suficiente stock no se muestra mensaje, al final muestra detalle de articulos
                             'sin stock


'------------------------------------------------
'Propiedades del modulo CStock
'------------------------------------------------

Public Property Let codArtic(ByVal vData As String)
     mvarCodArtic = vData
End Property

Public Property Get codArtic() As String
     codArtic = mvarCodArtic
End Property


Public Property Let codAlmac(ByVal vData As Integer)
     mvarCodAlmac = vData
End Property

Public Property Get codAlmac() As Integer
     codAlmac = mvarCodAlmac
End Property


Public Property Let Fechamov(ByVal vData As String)
     mvarFechaMov = vData
End Property

Public Property Get Fechamov() As String
     Fechamov = mvarFechaMov
End Property



Public Property Let HoraMov(ByVal vData As String)
     mvarHoraMov = vData
End Property

Public Property Get HoraMov() As String
     HoraMov = mvarHoraMov
End Property


Public Property Let tipoMov(ByVal vData As String)
     mvarTipoMov = vData
End Property

Public Property Get tipoMov() As String
     tipoMov = mvarTipoMov
End Property


Public Property Let DetaMov(ByVal vData As String)
     mvarDetaMov = vData
End Property

Public Property Get DetaMov() As String
     DetaMov = mvarDetaMov
End Property


Public Property Let cantidad(ByVal vData As Single)
     mvarCantidad = vData
End Property

Public Property Get cantidad() As Single
     cantidad = mvarCantidad
End Property


Public Property Let Importe(ByVal vData As Currency)
     mvarImporte = vData
End Property

Public Property Get Importe() As Currency
     Importe = mvarImporte
End Property


Public Property Let Trabajador(ByVal vData As Long)
     mvarCodOpe = vData
End Property

Public Property Get Trabajador() As Long
     Trabajador = mvarCodOpe
End Property


Public Property Let Documento(ByVal vData As String)
     mvarDocum = vData
End Property

Public Property Get Documento() As String
     Documento = mvarDocum
End Property


Public Property Let LineaDocu(ByVal vData As Integer)
     mvarLinea = vData
End Property

Public Property Get LineaDocu() As Integer
     LineaDocu = mvarLinea
End Property


Public Property Let Campo(ByVal vData As Long)
     mvarCampo = vData
End Property

Public Property Get Campo() As Long
     Campo = mvarCampo
End Property


'------------------------------------------------
'Metodos del modulo CStock
'------------------------------------------------


Public Function HayStock(stockOrig As Single, stockTras As Single) As Boolean
'IN: stockOrig: stock existente en almacen Origen
'    stockTras: stock a traspasar del origen a otro almacen
'OUT: True=Hay stock disponible en almacen para realizar el movimiento
'     False=No hay suficiento stock en el almacen
    
    Dim b As Boolean
    
    If stockTras > 0 Then
        If stockOrig >= stockTras Then
        'Si cantidad en stock > cantidad a traspasar entonces
            b = True
        Else
            b = False
        End If
    Else
        b = True
    End If
    HayStock = b
End Function


Public Function ControlStock(vStock As Long) As Boolean
'Si hay control de Stock no se permite realizar el movimiento
'Sino hay control se pregunta si se permite realizar el movimiento
Dim devuelve As String
Dim b As Boolean
Dim NomArtic As String
        
        NomArtic = DevuelveDesdeBDNew(cAgro, "sartic", "nomartic", "codartic", codArtic, "T")
        
        devuelve = "Control de Stock : " & vbCrLf
        devuelve = devuelve & "----------------------" & vbCrLf & vbCrLf
        devuelve = devuelve & " No hay suficiente Stock en el Almacen del Artículo:     "
        devuelve = devuelve & vbCrLf & " Código:   " & codArtic & vbCrLf
        devuelve = devuelve & " Desc.: " & NomArtic
        devuelve = devuelve & vbCrLf & vbCrLf
        devuelve = devuelve & "(Stock = " & vStock & ")"
        
        If DetaMov = "OFE" Or DetaMov = "PEV" Then 'No hay control de stock para bloquear solo informa
            MsgBox devuelve, vbInformation
            b = False
        Else
'--monica: lo he quitado por el parametro
'            If vParamAplic.ControlStock Then
'                'Si hay control Stock no permitir traspaso
'                b = True
'                If Not MuestraMenCStock Then
'                    Select Case DetaMov
'                        Case "REG"
'                            devuelve = devuelve & vbCrLf & vbCrLf & " No se puede realizar el Movimiento de Almacen. "
'                        Case "TRA"
'                            devuelve = devuelve & vbCrLf & vbCrLf & " No se puede realizar el Traspaso de Almacen. "
'                        Case "ALV", "ALR", "ALC"
'                            devuelve = devuelve & vbCrLf & vbCrLf & " No se puede realizar el Albaran. "
'                    End Select
'                    MsgBox devuelve, vbExclamation
'                End If
'            Else
'                Select Case DetaMov
'                Case "REG"
'                    devuelve = devuelve & vbCrLf & vbCrLf & " ¿Desea realizar el Movimiento de Almacen? "
'                Case "TRA"
'                    devuelve = devuelve & vbCrLf & vbCrLf & " ¿Desea realizar el Traspaso de Almacen? "
'                Case "ALV", "ALR", "ALC"
'                    devuelve = devuelve & vbCrLf & vbCrLf & " ¿Desea realizar el Albaran? "
'                End Select
'                If MsgBox(devuelve, vbQuestion + vbYesNo) = vbYes Then
'                    b = False
'                Else
'                    b = True
'                End If
'            End If
        End If
        ControlStock = b
End Function


Private Function ObtenerStock() As String
'OUT: cantidad de Stock que hay en el almacen vColAlm del articulo vCodArtic
Dim cant As String
Dim devuelve As String

    cant = DevuelveDesdeBDNew(cAgro, "advartic_salmac", "canstock", "codartic", codArtic, "T", , "codalmac", CStr(codAlmac), "N")
    If cant = "" Then 'No hay de ese artículo en Almacen
        devuelve = "No existe en el Almacen: " & codAlmac & vbCrLf
        devuelve = devuelve & "El Artículo: " & codArtic
        MsgBox devuelve, vbExclamation
        cant = "Error"
    ElseIf CSng(cant) = 0 Then
    End If
    ObtenerStock = cant
End Function


Public Function MueveStock() As Boolean
Dim vArtic As CArticuloADV
Dim b As Boolean

    Set vArtic = New CArticuloADV
    If vArtic.LeerDatos(codArtic) Then
'--monica
'        If vArtic.EsInstalacion Then
'            'Si el articulo es una instalacion no mueve stock
'            b = False
'        Else
            ' ver si hay control de stock para el articulo
            If vArtic.CtrStock Then
'                    vStock = ObtenerStock
                b = True
            Else
                b = False
            End If
'--monica
'        End If
   End If
   Set vArtic = Nothing
   MueveStock = b
End Function



Public Function MoverStock(Optional NoMenStock As Boolean) As Boolean
Dim vStock As String
Dim b As Boolean
Dim devuelve As String
Dim vArtic As CArticuloADV

    MuestraMenCStock = NoMenStock
       
    MoverStock = False
    
    'Comprobar si hay suficiente stock para realizar el Movimiento de stock si
    'el articulo NO es INSTALACION
    'Coprobar si el articulo es instalacion
'    devuelve = DevuelveDesdeBDNew(conAri, "sartic", "codfamia", "codartic", codArtic, "T")
'    If devuelve <> "" Then
'        devuelve = DevuelveDesdeBDNew(conAri, "sfamia", "instalac", "codfamia", devuelve, "N")
    Set vArtic = New CArticuloADV
    If vArtic.LeerDatos(codArtic) Then
'--monica, no hay instalaciones
'        If vArtic.EsInstalacion Then
'            b = True
'        Else
    '        If devuelve = "0" Then 'No es instalacion y comprobar stock, si hay control de stock
                'si hay control de stock para el articulo
'                devuelve = DevuelveDesdeBDNew(conAri, "sartic", "ctrstock", "codartic", codArtic, "T")
'                If Val(devuelve) = 1 Then 'hay ctrol stock
                If vArtic.CtrStock Then
                    vStock = ObtenerStock
                    If (vStock = "Error") Then Exit Function
                    If HayStock(CSng(vStock), cantidad) Then
                        b = True
                    Else 'No hay stock para el movimiento. Ver control de Stock
                        b = Not (ControlStock(Val(vStock)))
                    End If
                Else
                    b = False
                End If
    '        ElseIf devuelve = "1" Then 'Es instalacion
    '            b = True
    '        End If
'--monica
'        End If
    End If
    Set vArtic = Nothing
    MoverStock = b
End Function


Public Function ActualizarStock(Optional EsTicket As Boolean) As Boolean
'Actualiza el stock en la tabla advsmoval y en la tabla salmac
'EsTicket: si estamos generando un ticket no se comprueba el stock auque se quede negativo
Dim b As Boolean
Dim ADonde As String, cadMen As String
Dim CadValues As String 'cadena para la SQL de insertar en la tabla salmac
Dim vTipoMov As String

    ActualizarStock = False
    b = False
    On Error GoTo EActualizarStock
    
'    'comprobar que el articulo tiene control de stock
'    'si el articulo tiene control de stock comprobamos que haya suficiente stock
'    'para realizar le movimiento
'    cadMen = DevuelveDesdeBDNew(conAri, "sartic", "ctrstock", "codartic", codartic, "T")
'    If Val(cadMen) = 1 Then 'hay ctrol de stock para le articulo
'
    If tipoMov = "S" Then
        vTipoMov = "0" 'Movimiento de Salida de Almacen
    Else
        vTipoMov = "1" 'Movimiento de Entrada de Almacen
    End If
        
    '---- Insertamos en Movimientos Artículos
    ADonde = "Insertando datos en Movimientos de Articulos (tabla: advsmoval)."
    If Fechamov = "" Then Fechamov = Now
    'Construir la cadena VALUES para insertar en la tabla "advsmoval"
    CadValues = "VALUES (" & DBSet(codArtic, "T") & ", " & codAlmac & ", " & DBSet(Fechamov, "F") & ", '"
    If HoraMov = "" Then HoraMov = Format(Fechamov, "dd/mm/yyyy") & " " & Format(Now, "hh:mm:ss")
    CadValues = CadValues & Format(HoraMov, "yyyy-mm-dd hh:mm:ss") & "', " & vTipoMov & " , '" & DetaMov & "', " & DBSet(cantidad, "N") & ", "
    CadValues = CadValues & DBSet(Importe, "N") & ", " & Trabajador & ", " & ValorNulo & ", " & DBSet(Format(Documento, "0000000"), "T") & ", " & LineaDocu & ","
    CadValues = CadValues & DBSet(Campo, "N") & ")"
    
    b = InsertarMovimArticulos(CadValues)


    '---- Modificar stock del articulo
    If b Then
        ADonde = "Actualizando Stocks Almacenes (tabla: salmac)."
        If MueveStock Then
            If vTipoMov = "1" Then 'Entrada: no se comprueba si hay stock
'                If ActualizarStockAlmacen() Then b = True
                b = ActualizarStockAlmacen()
            Else
                'If DetaMov = "FTI" Then
                If EsTicket Then
                    'En los tickets se actualiza el stock aunque no haya
                    'no se paraliza la venta
'                    If ActualizarStockAlmacen() Then b = True
                    b = ActualizarStockAlmacen()
                Else
                    'comprobar si hay stock, y si no hay si permito que
                    'se mueva o no.
                    If MoverStock(False) Then
'                        If ActualizarStockAlmacen() Then b = True
                        b = ActualizarStockAlmacen()
                    Else
                        b = False
                    End If
                End If
            End If
        End If
    End If
    
    
    '---- si es articulo de conjunto actualizar stock e insertar movimientos
    'para cada una de las líneas de conjunto
    'if b then b=ActualizarStockLinConjunto
    
    
    ActualizarStock = b
    Exit Function
    
EActualizarStock:
    If Err.Number <> 0 Then
        cadMen = "Actualizar Stock." & vbCrLf & "----------------------------" & vbCrLf
        cadMen = cadMen & ADonde
        ActualizarStock = False
        MuestraError Err.Number, cadMen, Err.Description
    End If
End Function


Private Function InsertarMovimArticulos(CadValues As String) As Boolean
Dim SQL As String
On Error Resume Next

    SQL = "INSERT INTO advsmoval (codartic, codalmac, fechamov, horamovi, tipomovi, detamovi, cantidad, impormov, codigope, letraser, document, numlinea, codcampo ) "
    SQL = SQL & CadValues
    conn.Execute SQL
    
    If Err.Number <> 0 Then
        'Hay error , almacenamos y salimos
        InsertarMovimArticulos = False
    Else
        InsertarMovimArticulos = True
    End If
End Function


Private Function ActualizarStockAlmacen() As Boolean
Dim SQL As String
On Error Resume Next

    ActualizarStockAlmacen = False
    If cantidad <> 0 Then
        If tipoMov = "E" Then 'Mov. de Entrada
            '==== Aumentar el stock en el Almacen
            SQL = "UPDATE advartic_salmac Set canstock = canstock + " & DBSet(cantidad, "N")
            SQL = SQL & " WHERE codartic =" & DBSet(codArtic, "T") & " AND "
            SQL = SQL & " codalmac =" & codAlmac
        Else 'Mov. de Salida
            '==== Disminuir Stock en Almacen Origen
            SQL = "UPDATE advartic_salmac Set canstock = canstock - " & DBSet(cantidad, "N")
            SQL = SQL & " WHERE codartic =" & DBSet(codArtic, "T") & " AND "
            SQL = SQL & " codalmac =" & codAlmac
        End If
        conn.Execute SQL
    End If
    If Err.Number <> 0 Then
        'Hay error , almacenamos y salimos
        ActualizarStockAlmacen = False
    Else
        ActualizarStockAlmacen = True
    End If
End Function


Public Function DevolverStock() As Boolean
'Cuando se elimina una linea de movimiento de stock:
'Actualiza el stock en la tabla "salmac" aumentando de nuevo la cantidad de la linea
'Eliminar la linea que se introdujo en la "advsmoval"
Dim b As Boolean
Dim ADonde As String, cadMen As String

    DevolverStock = False
    b = False
    
    On Error GoTo EDevolverStock
    
    'Modificar stock
    ADonde = "Actualizando Stocks Almacenes (tabla: advartic_salmac)."
    If MueveStock Then
        If ActualizarStockAlmacen() Then b = True
    Else
        b = True
    End If
    
    If b Then
        'Eliminar en Movimientos Artículos
        ADonde = "Eliminando datos en Movimientos de Articulos (tabla: advsmoval)."
        If EliminarMovimArticulos Then b = True
    End If
    
    DevolverStock = b
    
EDevolverStock:
    If Err.Number <> 0 Then
        cadMen = "Actualizar Stock." & vbCrLf & "----------------------------" & vbCrLf
        cadMen = cadMen & ADonde
        MuestraError Err.Number, cadMen, Err.Description
        DevolverStock = False
    End If
End Function


Private Function EliminarMovimArticulos() As Boolean
Dim SQL As String
On Error Resume Next

    SQL = "DELETE FROM advsmoval WHERE codartic=" & DBSet(codArtic, "T") & " AND codalmac=" & codAlmac
    SQL = SQL & " AND detamovi='" & DetaMov & "' AND document= " & DBSet(Documento, "T") & " AND numlinea= " & LineaDocu
    SQL = SQL & " AND fechamov=" & DBSet(Fechamov, "F")
    'en albaranes de compra añadir el proveedor
    If DetaMov = "ALC" Then
        SQL = SQL & " AND codigope=" & Me.Trabajador 'en compras aqui guardamos el proveedor
    End If
    
    conn.Execute SQL
    
    If Err.Number <> 0 Then
        'Hay error , almacenamos y salimos
        EliminarMovimArticulos = False
    Else
        EliminarMovimArticulos = True
    End If

End Function


Public Function ModificarStock(oldCant As Single) As Boolean
'Cuando se modifica una linea de movimiento de stock:
'Modificar la linea que se introdujo en la "advsmoval"
'Actualiza el stock en la tabla "salmac" aumentando/disminuyendo la diferencia de cantidad de la linea
Dim b As Boolean
Dim ADonde As String, cadMen As String
Dim difCant As Single, difTipoM As String

    ModificarStock = False
    b = False
    
    On Error GoTo EModificarStock
    
    'Modificar en Movimientos Artículos (tabla: advsmoval)
    ADonde = "Modificando datos en Movimientos de Articulos (tabla: advsmoval)."
    If ModificarMovimArticulos Then b = True
    
    If b Then
        'Modificar stock
        ADonde = "Modificando Stocks Almacenes (tabla: salmac)."
        'Obtener si se suma o resta la diferencia de cantidad
        difCant = ObtenerDiferenciaStock(oldCant, difTipoM)
        mvarCantidad = difCant
        mvarTipoMov = difTipoM
        b = ActualizarStockAlmacen()
    End If
    ModificarStock = b
    
EModificarStock:
    If Err.Number <> 0 Then
        cadMen = "Actualizar Stock." & vbCrLf & "----------------------------" & vbCrLf
        cadMen = cadMen & ADonde
        MuestraError Err.Number, cadMen, Err.Description
        ModificarStock = False
    End If
End Function


Private Function ModificarMovimArticulos() As Boolean
'Modifica la cantidad y el importe en la tabla "advsmoval"
Dim SQL As String
On Error Resume Next
    
    SQL = "UPDATE advsmoval SET cantidad=" & DBSet(cantidad, "N") & ", impormov=" & DBSet(Importe, "N") 'Importe
    SQL = SQL & " WHERE codartic=" & DBSet(codArtic, "T") & " AND codalmac=" & codAlmac
    SQL = SQL & " AND detamovi='" & DetaMov & "' AND document= " & DBSet(Documento, "T") & " AND numlinea= " & LineaDocu
    conn.Execute SQL
    
    If Err.Number <> 0 Then
        'Hay error , almacenamos y salimos
        ModificarMovimArticulos = False
    Else
        ModificarMovimArticulos = True
    End If
End Function


Private Function ObtenerDiferenciaStock(oldCant As Single, difTipoM As String) As Single
'Cuando se modifica la cantidad de una linea hay que reestablecer el stock en la
'tabla "salmac" aumentando o restando la diferencia
'IN -> oldCant: cantidad que habia antes de modificar la cantidad
'OUT -> difTipoMov: Indicara si se devera aumentar o disminuir el stock (Entrada o Salida)
'       RETURN: cantidad a aumentar o disminuir en almacen: salmac
Dim TipoM As Byte
Dim difCant As Single
Dim SQL As String
Dim Rs As ADODB.Recordset

    On Error GoTo EDifStock

    'Obtener el tipo de movimiento en la advsmoval
    SQL = "SELECT tipomovi FROM advsmoval "
    SQL = SQL & " WHERE codartic=" & DBSet(codArtic, "T") & " AND codalmac=" & codAlmac
    SQL = SQL & " AND detamovi='" & DetaMov & "' AND document= " & DBSet(Documento, "T") & " AND numlinea= " & LineaDocu
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs.EOF Then
        TipoM = Rs.Fields(0).Value
    End If
    Rs.Close
    Set Rs = Nothing
    
    'En funcion del tipo de Movimiento y la diferencia de stock se sumara o se resta la diferencia de stock
    difCant = oldCant - cantidad
    If TipoM = 1 Then 'Entrada
        If difCant >= 0 Then 'restar la diferencia
            difTipoM = "S"
        Else 'sumar la diferencia
            difTipoM = "E"
        End If
    ElseIf TipoM = 0 Then 'Salida
        If difCant >= 0 Then
            'sumar la diferencia
            difTipoM = "E"
        Else 'restar la diferencia
            difTipoM = "S"
        End If
    End If
    ObtenerDiferenciaStock = Abs(difCant)
    
EDifStock:
    If Err.Number <> 0 Then
        ObtenerDiferenciaStock = -1 'error
    End If
End Function



'## Laura 22/11/2006
'Private Function ActualizarStockArticConjunto() As Boolean
'Dim cArt As CArticulo
'Dim b As Boolean
'Dim menError As String
'
'    On Error GoTo ErrActConj
'
'    b = False
'    Set cArt = New CArticulo
'    b = cArt.LeerDatos(Me.codArtic) 'articulo q estamos actualizando stock
'    If b Then
'        If cArt.EsConjunto Then
'            'actualizar stock de las lineas de conjunto
'            b = cArt.ActualizarStockLinConjunto(Me, menError)
'        Else
'            b = True 'no es conjunto y salimos con ok
'        End If
'    End If
'    Set cArt = Nothing
'
'    ActualizarStockArticConjunto = b
'    Exit Function
'
'ErrActConj:
'    ActualizarStockArticConjunto = False
'    MuestraError Err.Number, "Actualizar stock artículo de conjunto", Err.Description
'End Function
