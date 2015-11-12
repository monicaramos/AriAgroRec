VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CArticuloADV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'ARTICULOS
'TABLA: advartic

'ATRIBUTOS
'Variables locales que contienen valores de propiedad
Private mCodartic As String 'Cod. articulo
Private mNomArtic As String 'Nombre articulo

Private mCodFamia As String 'Familia del Articulo

Private mCtrStock As Byte 'hay control de stock (si/No)

Private mCodigIVA As Byte
Private mTipoProd As Byte

Private mPrecioVe As String 'Precio de venta (P.V.P.)
Private mPreciomp As Currency 'Precio medio ponderado
Private mPreciouc As Currency 'Precio ultima compra

Private mTextoVen As String 'texto a mostrar para ventas
Private mTextoCom As String 'texto a mostrar para compras



'------------------------------------------------
'Propiedades del modulo CArticulo
'------------------------------------------------

'**** Codigo del articulo
Public Property Let Codigo(ByVal vData As String)
     mCodartic = vData
End Property

Public Property Get Codigo() As String
     Codigo = mCodartic
End Property


'**** Nombre del Articulo
Public Property Let Nombre(ByVal vData As String)
     mNomArtic = vData
End Property

Public Property Get Nombre() As String
     Nombre = mNomArtic
End Property



'**** Familia del Articulo
Public Property Let Familia(ByVal vData As String)
     mCodFamia = vData
End Property

Public Property Get Familia() As String
     Familia = mCodFamia
End Property




'**** Articulo tiene control de stock (si/no)
Public Property Let CtrStock(ByVal vData As Byte)
     mCtrStock = vData
End Property

Public Property Get CtrStock() As Byte
     CtrStock = mCtrStock
End Property


'**** Precio de Venta
Public Property Let PrecioVenta(ByVal vData As String)
     mPrecioVe = vData
End Property

Public Property Get PrecioVenta() As String
     PrecioVenta = mPrecioVe
End Property


'**** Precio medio ponderado
Public Property Let PrecioMedPon(ByVal vData As Currency)
     mPreciomp = vData
End Property

Public Property Get PrecioMedPon() As Currency
     PrecioMedPon = mPreciomp
End Property


'**** Precio ultima compra
Public Property Let PrecioUltCom(ByVal vData As Currency)
     mPreciouc = vData
End Property

Public Property Get PrecioUltCom() As Currency
     PrecioUltCom = mPreciouc
End Property


'**** Tipo de IVA del articulo
Public Property Let TipoIVA(ByVal vData As Byte)
     mCodigIVA = vData
End Property

Public Property Get TipoIVA() As Byte
     TipoIVA = mCodigIVA
End Property

'**** Tipo de Articulo (0=producto 1=trabajo 2=varios)
Public Property Let TipoProd(ByVal vData As Byte)
     mTipoProd = vData
End Property

Public Property Get TipoProd() As Byte
     TipoProd = mTipoProd
End Property



'**** Texto para Ventas
Public Property Let TextoVentas(ByVal vData As String)
     mTextoVen = vData
End Property

Public Property Get TextoVentas() As String
     TextoVentas = mTextoVen
End Property


'**** Texto para Compras
Public Property Let TextoCompras(ByVal vData As String)
     mTextoCom = vData
End Property

Public Property Get TextoCompras() As String
     TextoCompras = mTextoCom
End Property



'------------------------------------------------
'Procedimientos del modulo CArticuloADV
'------------------------------------------------

Public Function Existe(vCodArtic As String) As Boolean
'Comprueba si existe el articulo en la BD
Dim devuelve As String
Dim Nom As String

    On Error GoTo EExiste
    
    Nom = "nomartic"
    devuelve = DevuelveDesdeBD("codartic", "advartic", "codartic", vCodArtic, "T", Nom)
    If devuelve = "" Then
        Existe = False
        devuelve = "No existe el Artículo " & vCodArtic & ". "
        MsgBox devuelve, vbExclamation
    Else
        Existe = True
        Codigo = vCodArtic
        Nombre = Nom
    End If
    
EExiste:
    If Err.Number <> 0 Then Existe = False
End Function



Public Function LeerDatos(vCodArtic As String) As Boolean
'Leer los datos de un Articulo dado
'Lee de la BD: Ariagrorec, Tabla: advartic
'OUT: True si lee los datos correctamente
Dim RS As ADODB.Recordset
Dim SQL As String

    On Error GoTo ELeer
    LeerDatos = False
    
    SQL = "SELECT codartic,nomartic,codfamia,ctrstock,preciove,codigiva,textoven,textocom,preciomp,preciouc, tipoprod "
    SQL = SQL & " FROM advartic "
    SQL = SQL & " WHERE codartic=" & DBSet(vCodArtic, "T")
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText

    If RS.EOF Then
        LeerDatos = False
    Else
        Codigo = CStr(RS!codArtic)
        
        Nombre = RS!NomArtic
        Familia = RS!codfamia
        CtrStock = RS!CtrStock
        PrecioVenta = RS!preciove
        TipoIVA = RS!CodigIVA
        TipoProd = RS!TipoProd
        
        PrecioMedPon = DBLet(RS!precioMP, "N") 'precio medio ponderado
        PrecioUltCom = DBLet(RS!PrecioUC, "N") 'precio ultima compra
        TextoVentas = DBLet(RS!textoven, "T")
        TextoCompras = DBLet(RS!textocom, "T")

        LeerDatos = True
    End If

    RS.Close
    Set RS = Nothing
    Exit Function

ELeer:
    SQL = "Se ha producido un error. " & "Datos Artículo" & vbCrLf
    SQL = SQL & "Número: " & Err.Number & vbCrLf
    SQL = SQL & "Descripción: " & Err.Description
    MsgBox SQL, vbExclamation
    Set RS = Nothing
    LeerDatos = False
End Function



'Public Sub MostrarStatusArtic(Bloquea As Boolean)
''Recuperamos el estatus (codstatu) del articulo, y si tiene
''valor 1 o 2 lo mostramos en un mensaje
''codstatu=0 normal
''codstatu=1 bloqueado, muestra un mensaje
''codstatu=2 caducado, muestra un mensaje y bloquea linea de articulo
'
'    Bloquea = False
'
'    Select Case Status
'        ' 0 = Normal
'        Case 1 'BLOQUEADO
'            MsgBox "El articulo " & Codigo & " esta BLOQUEADO." & vbCrLf & "No se puede insertar una linea con este artículo.", vbExclamation, "Situación artículo"
'            Bloquea = True
'        Case 2 'CADUCADO
'            MsgBox "El articulo " & Codigo & " esta CADUCADO.", vbInformation, "Situación artículo"
'    End Select
'End Sub


Public Function ExisteEnAlmacen(codAlm As String, Optional CanStock As String) As Boolean
Dim devuelve As String
Dim Cantidad As String

    On Error GoTo EExisteAlm

    'Comprobar que existe de ese articulo en el almacen seleccionado
    Cantidad = "canstock"
    devuelve = DevuelveDesdeBDNew(cAgro, "advartic_salmac", "codartic", "codartic", Codigo, "T", Cantidad, "codalmac", codAlm, "N")
    
    If devuelve = "" Then
        MsgBox "No existe unidades del Artículo: " & Codigo & "  en el Almacen: " & codAlm, vbExclamation
        ExisteEnAlmacen = False
    Else
        ExisteEnAlmacen = True
        CanStock = Cantidad
    End If
    
EExisteAlm:
    If Err.Number <> 0 Then MuestraError Err.Number, "Comprobar Artículo", Err.Description
End Function




Public Function ExistenciaTotalAlmacenes() As Currency
'-- devuelve la cantidad de stock total en todos los almacenes del articulo
Dim rst As ADODB.Recordset
Dim SQL As String
    
    On Error GoTo EExisTotal
    
    If Codigo <> "" Then
        SQL = "SELECT SUM(canstock) FROM advartic_salmac where codartic=" & DBSet(Codigo, "T")
        Set rst = New ADODB.Recordset
        rst.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not rst.EOF Then
            ExistenciaTotalAlmacenes = rst.Fields(0).Value
        End If
        rst.Close
        Set rst = Nothing
    End If
    
    Exit Function
    
EExisTotal:
    If Err.Number <> 0 Then MuestraError Err.Number, "Obtener existencia total en Almacenes del Artículo", Err.Description
End Function


Public Function EnInventario(codAlm As String) As Boolean
'SI  se esta haciendo inventario del articulo o no
Dim devuelve As String

    devuelve = DevuelveDesdeBDNew(cAgro, "advartic_salmac", "statusin", "codartic", Codigo, "T", , "codalmac", codAlm, "N")
    
    If devuelve = "1" Then
        EnInventario = True
        MsgBox "El artículo " & Codigo & " esta inventariandose.", vbInformation
    Else
        EnInventario = False
    End If
End Function


Public Sub MostrarTextoCom()
'Recuperamos el texto para compras y ventas del articulo, y si tiene
'valor lo mostramos en un mensaje
     MsgBox TextoCompras, vbInformation, "Texto para compras"
End Sub


Public Sub MostrarTextoVen()
'Recuperamos el texto para compras y ventas del articulo, y si tiene
'valor lo mostramos en un mensaje
     MsgBox TextoVentas, vbInformation, "Texto para ventas"
End Sub



Public Function ActualizarUltFechaCompra(vFecha As String, vPrecio As String) As Boolean
'Modificar fecha ult.compra y precio ult.compra en la tabla sartic
Dim SQL As String
Dim actualiza As Boolean

    On Error GoTo EActFecha

    'Obtenemos la ult. fecha de compra que tiene el articulo
    SQL = DevuelveDesdeBDNew(cAgro, "advartic", "ultfecco", "codartic", Codigo, "T")
    
    actualiza = False
    
    If SQL = "" Then
        'No tiene fecha ult. compra
        actualiza = True
    ElseIf CDate(vFecha) >= CDate(SQL) Then
        'fecha de compra posterior
        actualiza = True
    End If
    
    If actualiza Then
        SQL = "UPDATE advartic SET ultfecco=" & DBSet(vFecha, "F")
        'actualizar el ult. precio de compra si el nuevo valor no es 0
        If CCur(vPrecio) > 0 Then SQL = SQL & ",preciouc=" & DBSet(vPrecio, "N")
        SQL = SQL & " WHERE codartic=" & DBSet(Codigo, "T", "N")
        
        conn.Execute SQL
    End If
    
EActFecha:
    If Err.Number <> 0 Then
        ActualizarUltFechaCompra = False
    Else
        ActualizarUltFechaCompra = True
    End If
End Function



Public Function ActualizarPrecioMedPond(vCantidad As Currency, vPrecio As Currency, Optional vCantAnt As Currency) As Boolean
'-- Modificar precio medio ponderado en la tabla sartic, se utiliza al crear
'-- albaranes de compra
'(IN) vCantidad=cantidad comprada o modificada
'(IN) vPrecio= precio de compra
'(IN) vCantAnt= cantidad anterior a la modificacion
Dim SQL As String
Dim actualiza As Boolean
Dim existencia As Currency
Dim Importe As Currency

    On Error GoTo EActPMP

    
    If PrecioMedPon = 0 Then
        'Obtenemos el precio medio ponderado que tiene el articulo
        'leemos de la tabla por si no se leyeron todos los campos de la sartic
        SQL = DevuelveDesdeBDNew(cAgro, "advartic", "preciomp", "codartic", Codigo, "T")
        If SQL <> "" Then PrecioMedPon = CCur(SQL)
    End If
    
    actualiza = False
    If PrecioMedPon = 0 Then
        'se actualiza con el precio de compra
        PrecioMedPon = vPrecio
        actualiza = True
    Else
        'se actualiza con la formula:
        '((existencia_total_almacenes_antes_compra * precio_medio_ponderado_actual) + importe_compra) / (existencia_total_almacenes_antes_compra + cantidad_compra)
        
        existencia = ExistenciaTotalAlmacenes
        'como aqui ya le habiamos sumado la cantidad q acabamos de comprar
        'se la quitamos para obtener la existencia antes de la compra
'        existencia = existencia - vCantidad

        'si hemos modificado la linea del albaran, a la existencia habrá q
        'quitarle la cantidad q habia antes
        If vCantAnt <> 0 Then existencia = existencia - vCantAnt
        
        'obtenemos el nuevo precio_medio_ponderado a partir del valor actual
        If (existencia + vCantidad) <> 0 Then
            Importe = Round2((vCantidad * vPrecio), 2)
            PrecioMedPon = Round2(((existencia * PrecioMedPon) + Importe) / (existencia + vCantidad), 4)
'        PrecioMedPon = Round(PrecioMedPon, 4)
            actualiza = True
        End If
    End If

    
    If actualiza Then
        SQL = "UPDATE advartic SET preciomp=" & DBSet(PrecioMedPon, "N")
        SQL = SQL & " WHERE codartic=" & DBSet(Codigo, "T", "N")
        
        conn.Execute SQL
    End If
    
    ActualizarPrecioMedPond = True
    Exit Function
    
EActPMP:
'    If Err.Number <> 0 Then
        ActualizarPrecioMedPond = False
'    Else
'        ActualizarPrecioMedPond = True
'    End If
End Function






Public Function ReestablecerUltFechaCompra(vAlmac As String) As Boolean
'Modificar fecha ult.compra y precio ult.compra en la tabla sartic
'con los valores del ultimo movimiento que encontremos en la smoval
Dim SQL As String
Dim cadSel As String
Dim RS As ADODB.Recordset

    On Error GoTo ERestFecha

    'Obtenemos el ultimo movimiento de albaran de compra que tiene el articulo
    cadSel = " detamovi='ALC' and codartic=" & DBSet(Codigo, "T") & " AND codalmac=" & vAlmac & " AND and tipomovi=1 "
    
    SQL = "select * from smoval WHERE " & cadSel
    SQL = SQL & " and horamovi =(select max(horamovi) from smoval where " & cadSel & ")"

    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not RS.EOF Then
        'actualizamos la sartic con esos valores
        SQL = "UPDATE advartic SET ultfecco=" & DBSet(RS!Fechamov, "F")
        SQL = SQL & ", preciouc=" & DBSet(RS!impormov, "N")
        SQL = SQL & " WHERE codartic=" & DBSet(Codigo, "T")
        conn.Execute SQL
    End If
    RS.Close
    Set RS = Nothing

    
ERestFecha:
    If Err.Number <> 0 Then
        ReestablecerUltFechaCompra = False
    Else
        ReestablecerUltFechaCompra = True
    End If
End Function




Public Function ReestablecerPrecioMedPon(vCantidad As Currency, vPrecio As Currency) As Boolean
'reestablecer el valor del precio medio ponderado
'       (pmp_actual * stock_actual) - (vCantidad * vPrecio)
'pmp=   ---------------------------------------------------
'       (stock_actual - cantidad)

Dim existencia As Currency 'existencia de stock en almacenes
Dim SQL As String
Dim Importe As Currency

    On Error GoTo EResPMP
    
    ReestablecerPrecioMedPon = False
    
    If PrecioMedPon = 0 Then
        'Obtenemos el precio medio ponderado que tiene el articulo
        'leemos de la tabla por si no se leyeron todos los campos de la sartic
        SQL = DevuelveDesdeBDNew(cAgro, "advartic", "preciomp", "codartic", Codigo, "T")
        If SQL <> "" Then PrecioMedPon = CCur(SQL)
    End If
    
    
    existencia = ExistenciaTotalAlmacenes
    
    If (existencia - vCantidad) <> 0 Then
        Importe = Round2((vCantidad * vPrecio), 2)
        PrecioMedPon = Round2(((PrecioMedPon * existencia) - Importe) / (existencia - vCantidad), 4)
'        PrecioMedPon = Round(PrecioMedPon, 4)
    Else
        PrecioMedPon = 0
    End If
    
    SQL = "UPDATE advartic SET preciomp=" & DBSet(PrecioMedPon, "N")
    SQL = SQL & " WHERE codartic=" & DBSet(Codigo, "T", "N")
        
    conn.Execute SQL
    
    ReestablecerPrecioMedPon = True
    Exit Function
    
EResPMP:
    ReestablecerPrecioMedPon = False
    MsgBox Err.Number, "Reestablecer precio medio ponderado.", Err.Description
End Function




Public Function ObtenerPorceIVA() As Currency
'Devuelve el % de IVA del articulo correspondiente al codigiva
Dim RS As ADODB.Recordset
Dim SQL As String

    On Error GoTo EPorce
    
    If Codigo <> "" Then 'codartic
        'Obtenemos el % de IVA a aplicar
        SQL = "SELECT porceiva FROM tiposiva WHERE codigiva=" & TipoIVA
        Set RS = New ADODB.Recordset
        RS.Open SQL, ConnConta, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        SQL = ""
        If Not RS.EOF Then
            SQL = RS.Fields(0).Value
        End If
        RS.Close
        Set RS = Nothing
        ObtenerPorceIVA = CCur(SQL)
    End If
    
EPorce:
    If Err.Number <> 0 Then
        SQL = "Se ha producido un error. " & "Obteniendo porcentaje de IVA." & vbCrLf
        SQL = SQL & "Número: " & Err.Number & vbCrLf
        SQL = SQL & "Descripción: " & Err.Description
        MsgBox SQL, vbExclamation
        Set RS = Nothing
    End If
End Function



Public Function ObtenerPrecioConIVA(Optional newPrecio As String) As Currency
'Devuelve el precio del articulo tras aplicarle el IVA correspondiente
'si newPrecio tiene valor le aplicamos el IVA a ese precio ya que habremos
'introducido el precio del articulo manualmente
'si no le pasamos newPrecio se aplicará al precio_venta del articulo
Dim SQL As String
Dim Porce As Currency ' % de IVA

    On Error GoTo EPrecio
    
    If Codigo <> "" Then 'codartic
        'Obtener el % de IvA q se le aplica al articulo
        Porce = ObtenerPorceIVA
        
        If Porce <> CCur(0) Then
            If newPrecio = "" Then
                'al precio de venta del articulo
                SQL = CalcularDto(PrecioVenta, CStr(Porce))
                ObtenerPrecioConIVA = Round(CCur(ComprobarCero(PrecioVenta)) + CCur(ComprobarCero(SQL)), 4)
            Else
                'al precio especial del articulo (para un cliente tn promociones)
                SQL = CalcularDto(newPrecio, CStr(Porce))
                ObtenerPrecioConIVA = Round(CCur(ComprobarCero(newPrecio)) + CCur(ComprobarCero(SQL)), 4)
            End If
        Else
            If newPrecio = "" Then
                ObtenerPrecioConIVA = CCur(ComprobarCero(PrecioVenta))
            Else
                ObtenerPrecioConIVA = CCur(ComprobarCero(newPrecio))
            End If
        End If
    End If
    
EPrecio:
    If Err.Number <> 0 Then
        SQL = "Se ha producido un error. " & "Calculando precio con IVA." & vbCrLf
        SQL = SQL & "Número: " & Err.Number & vbCrLf
        SQL = SQL & "Descripción: " & Err.Description
        MsgBox SQL, vbExclamation
    End If
End Function


Public Function ObtenerPrecioSinIVA(Optional newPrecio As String) As Currency
'Devuelve el precio del articulo tras quitarle el IVA correspondiente
'IN -> newPrecio: precio con IVA
'si newPrecio tiene valor le quitamos el IVA a ese precio ya que habremos
'introducido el precio del articulo manualmente
'si no le pasamos newPrecio será al precio_venta del articulo
Dim SQL As String
Dim Porce As Currency ' % de IVA
Dim PreuSinIVA As Currency

    On Error GoTo EPrecio
    
    If Codigo <> "" Then 'codartic
        'Obtener el % de IvA q se le aplica al articulo
        Porce = ObtenerPorceIVA
        Porce = Porce / 100
        
        'precio al que le vamos a quitar el IVA
        If newPrecio = "" Then
            'al precio de venta del articulo
            PreuSinIVA = CCur(ComprobarCero(PrecioVenta))
        Else
            'al precio especial del articulo (para un cliente tn promociones)
            'o art. de varios y precio manual
            PreuSinIVA = CCur(ComprobarCero(newPrecio))
            If Porce <> CCur(0) Then
                'quitarle el IVA al precio
                PreuSinIVA = PreuSinIVA / (1 + Porce)
                PreuSinIVA = Round(PreuSinIVA, 4)
            End If
        End If
        ObtenerPrecioSinIVA = PreuSinIVA
    End If

'    curImporte = curImporte / (1 + curPorce) 'importe sin iva
'    curCuota = Round((curPorce * curImporte), 2)
'    curImporte = Round(curImporte, 2)
'
'    'valores que devuelve: Importe sin iva, cuota de iva
'    ImporteSinIVA = Format(curImporte, FormatoImporte)
'    sCuota = Format(curCuota, FormatoImporte)
'
    Exit Function
    
EPrecio:
    If Err.Number <> 0 Then
        SQL = "Se ha producido un error. " & "Calculando precio sin IVA." & vbCrLf
        SQL = SQL & "Número: " & Err.Number & vbCrLf
        SQL = SQL & "Descripción: " & Err.Description
        MsgBox SQL, vbExclamation
    End If
End Function


