VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmFactRectifSocio 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informes"
   ClientHeight    =   6030
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   6945
   Icon            =   "frmFactRectifSocio.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6030
   ScaleWidth      =   6945
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cd1 
      Left            =   7680
      Top             =   5160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame FrameGenFacturaRect 
      Height          =   6015
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   6945
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   0
         Left            =   1860
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   1050
         Width           =   3330
      End
      Begin VB.Frame Frame8 
         BorderStyle     =   0  'None
         Height          =   1665
         Left            =   240
         TabIndex        =   17
         Top             =   3240
         Width           =   6375
         Begin VB.TextBox txtcodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   10
            Left            =   1620
            MaxLength       =   10
            TabIndex        =   5
            Top             =   180
            Width           =   1005
         End
         Begin VB.TextBox txtcodigo 
            Height          =   825
            Index           =   9
            Left            =   1620
            MaxLength       =   500
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   6
            Top             =   750
            Width           =   4755
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Imprimir Resumen"
            Height          =   255
            Index           =   1
            Left            =   4290
            TabIndex        =   19
            Top             =   60
            Width           =   1965
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Imprimir Factura"
            Height          =   195
            Index           =   0
            Left            =   4290
            TabIndex        =   18
            Top             =   420
            Width           =   1995
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Factura"
            ForeColor       =   &H00972E0B&
            Height          =   195
            Index           =   10
            Left            =   360
            TabIndex        =   21
            Top             =   -30
            Width           =   1035
         End
         Begin VB.Image imgFecha 
            Height          =   240
            Index           =   1
            Left            =   1260
            Picture         =   "frmFactRectifSocio.frx":000C
            Top             =   180
            Width           =   240
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Concepto"
            ForeColor       =   &H00972E0B&
            Height          =   195
            Index           =   9
            Left            =   390
            TabIndex        =   20
            Top             =   570
            Width           =   690
         End
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   0
         Left            =   5610
         TabIndex        =   8
         Top             =   5280
         Width           =   975
      End
      Begin VB.CommandButton CmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Index           =   1
         Left            =   4500
         TabIndex        =   7
         Top             =   5280
         Width           =   975
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   19
         Left            =   1875
         MaxLength       =   10
         TabIndex        =   2
         Top             =   1935
         Width           =   960
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   18
         Left            =   1875
         MaxLength       =   10
         TabIndex        =   1
         Top             =   1590
         Width           =   960
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   17
         Left            =   1860
         MaxLength       =   10
         TabIndex        =   4
         Top             =   2805
         Width           =   1005
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   16
         Left            =   1860
         MaxLength       =   10
         TabIndex        =   3
         Top             =   2460
         Width           =   1005
      End
      Begin MSComctlLib.ProgressBar Pb1 
         Height          =   255
         Left            =   540
         TabIndex        =   23
         Top             =   4920
         Width           =   6075
         _ExtentX        =   10716
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo Factura"
         ForeColor       =   &H00972E0B&
         Height          =   255
         Index           =   1
         Left            =   540
         TabIndex        =   22
         Top             =   840
         Width           =   945
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   29
         Left            =   900
         TabIndex        =   16
         Top             =   1590
         Width           =   465
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   28
         Left            =   900
         TabIndex        =   15
         Top             =   1950
         Width           =   420
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Factura"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   27
         Left            =   540
         TabIndex        =   14
         Top             =   1350
         Width           =   540
      End
      Begin VB.Label Label7 
         Caption         =   "Generación Factura Rectificativas"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   345
         Left            =   480
         TabIndex        =   13
         Top             =   300
         Width           =   5925
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   26
         Left            =   930
         TabIndex        =   12
         Top             =   2490
         Width           =   465
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   25
         Left            =   930
         TabIndex        =   11
         Top             =   2820
         Width           =   420
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   24
         Left            =   570
         TabIndex        =   10
         Top             =   2280
         Width           =   450
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   3
         Left            =   1515
         Picture         =   "frmFactRectifSocio.frx":0097
         Top             =   2805
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   2
         Left            =   1530
         Picture         =   "frmFactRectifSocio.frx":0122
         Top             =   2460
         Width           =   240
      End
   End
End
Attribute VB_Name = "frmFactRectifSocio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' +-+-+-+-+-+-+-+-+-+-+-
' +-+- Autor: LAURA +-+-
' +-+-+-+-+-+-+-+-+-+-+-

Option Explicit

Public NumCod As String 'Para indicar cod. Traspaso,Movimiento, etc. que llama
                        'Para indicar nº oferta a imprimir

Public CadTag As String 'Cadena con el Tag del campo que se va a poner en D/H en los listados
                        'Se necesita si el tipo de codigo es texto

Public Event RectificarFactura(Cliente As String, Observaciones As String)

Private Conexion As Byte
'1.- Conexión a BD Ariges  2.- Conexión a BD Conta

Private HaDevueltoDatos As Boolean


Private WithEvents frmSoc  As frmManSocios 'mantenimiento de socios
Attribute frmSoc.VB_VarHelpID = -1
 
Private WithEvents frmB As frmBuscaGrid  'Busquedas
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmC As frmCal 'calendario fechas
Attribute frmC.VB_VarHelpID = -1

'GENERALES PARA PASARLE A CRYSTAL REPORT
Private cadFormula As String 'Cadena con la FormulaSelection para Crystal Report
Private CadParam As String 'Cadena con los parametros para Crystal Report
Private numParam As Byte 'Numero de parametros que se pasan a Crystal Report
Private cadSelect As String 'Cadena para comprobar si hay datos antes de abrir Informe
Private cadTitulo As String 'Titulo para la ventana frmImprimir
Private cadNombreRPT As String 'Nombre del informe


Dim indCodigo As Integer 'indice para txtCodigo
Dim indFrame As Single 'nº de frame en el que estamos
 
'Se inicializan para cada Informe (tabla de BD a la que hace referencia
Dim Tabla As String
Dim Codigo As String 'Código para FormulaSelection de Crystal Report
Dim TipCod As String
Dim Orden1 As String 'Campo de Ordenacion (por codigo) para Cristal Report
Dim Orden2 As String 'Campo de Ordenacion (por nombre) para Cristal Report
Dim Tipo As String


Dim PrimeraVez As Boolean
Dim Contabilizada As Byte
Dim ConSubInforme As Boolean


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 0, cerrar
    If cerrar Then Unload Me

End Sub


Private Sub cmdAceptar_Click(Index As Integer)
Dim cOrden As String
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim numOp As Byte
Dim NomAlmac As String
Dim cTabla As String

Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String

    
    InicializarVbles
    
    If Not DatosOk Then Exit Sub
    
    'Añadir el parametro de Empresa
    CadParam = CadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
        
    '======== FORMULA  ====================================
    ' tipo de factura
    If Not AnyadirAFormula(cadSelect, "{rfactsoc.codtipom}='" & Mid(Combo1(0).Text, 1, 3) & "'") Then Exit Sub
    
    'D/H factura
    cDesde = Trim(txtcodigo(18).Text)
    cHasta = Trim(txtcodigo(19).Text)
    nDesde = ""
    nHasta = ""
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{rfactsoc.numfactu}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHHidrante=""") Then Exit Sub
    End If
    
    'D/H fecha
    cDesde = Trim(txtcodigo(16).Text)
    cHasta = Trim(txtcodigo(17).Text)
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{rfactsoc.fecfactu}"
        TipCod = "F"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHFecha=""") Then Exit Sub
    End If

    '  ya teniamos obligatoriamente el tipo de movimiento
    cadSelect = cadSelect & " and not (codtipom, numfactu, fecfactu) in "
    cadSelect = cadSelect & "(select rectif_codtipom, rectif_numfactu, rectif_fecfactu from rfactsoc "
    cadSelect = cadSelect & " where not rectif_codtipom is null and not rectif_numfactu is null and not rectif_fecfactu is null) "

    ProcesoFacturacionRectificativa Tabla, cadSelect
    
    DesBloqueoManual ("RECFAC") 'RECtificativas FACturas
    Pb1.visible = False
    
End Sub


Private Sub cmdCancel_Click(Index As Integer)
    Unload Me
End Sub


Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        PonerFoco txtcodigo(18)
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
Dim H As Integer, W As Integer
Dim List As Collection

    PrimeraVez = True
    limpiar Me

'    CommitConexion
    CargaCombo
    
    Tabla = "rfactsoc"
    
    indFrame = 0
    txtcodigo(10).Text = Format(Now, "dd/mm/yyyy")
    Me.Pb1.visible = False
    Combo1(0).ListIndex = 0
    
    Me.FrameGenFacturaRect.Top = -90
    Me.FrameGenFacturaRect.Left = 0
    Me.FrameGenFacturaRect.Height = 6015
    Me.FrameGenFacturaRect.Width = 6945
    W = Me.FrameGenFacturaRect.Width
    H = Me.FrameGenFacturaRect.Height
    
    Me.Check1(0).Value = 1
    Me.Check1(1).Value = 1
    
    'Esto se consigue poniendo el cancel en el opcion k corresponda
    Me.cmdCancel(0).Cancel = True
    Me.Width = W + 70
    Me.Height = H + 350
End Sub


Private Sub frmC_Selec(vFecha As Date)
    ' *** repasar si el camp es txtAux o Text1 ***
    txtcodigo(CByte(imgFecha(1).Tag)).Text = Format(vFecha, "dd/mm/yyyy") '<===
    ' ********************************************
End Sub

Private Sub frmSoc_DatoSeleccionado(CadenaSeleccion As String)
'    txtcodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000")
'    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub imgBuscar_Click(Index As Integer)
   Select Case Index
        Case 0, 1 ' Socios
            AbrirFrmSocios (Index + 23)
        Case 2, 3 ' Socios
            AbrirFrmSocios (Index + 32)
        Case 6, 7  'Socios
            AbrirFrmSocios (Index)
    End Select
    PonerFoco txtcodigo(indCodigo)
End Sub

Private Sub imgFecha_Click(Index As Integer)
    Dim esq As Long
    Dim dalt As Long
    Dim menu As Long
    Dim obj As Object
    Dim indice As Integer

    Set frmC = New frmCal
    
    esq = imgFecha(Index).Left
    dalt = imgFecha(Index).Top
        
    Set obj = imgFecha(Index).Container
      
      While imgFecha(Index).Parent.Name <> obj.Name
            esq = esq + obj.Left
            dalt = dalt + obj.Top
            Set obj = obj.Container
      Wend
    
    menu = Me.Height - Me.ScaleHeight 'ací tinc el heigth del menú i de la toolbar

    frmC.Left = esq + imgFecha(Index).Parent.Left + 30
    frmC.Top = dalt + imgFecha(Index).Parent.Top + imgFecha(Index).Height + menu - 40

    Select Case Index
        Case 1
            indice = 10
        Case 2, 3
            indice = Index + 14
    End Select

    imgFecha(1).Tag = indice '<===
    ' *** repasar si el camp es txtAux o Text1 ***
    If txtcodigo(indice).Text <> "" Then frmC.NovaData = txtcodigo(indice).Text
    ' ********************************************

    frmC.Show vbModal
    Set frmC = Nothing
    ' *** repasar si el camp es txtAux o Text1 ***
    PonerFoco txtcodigo(CByte(imgFecha(1).Tag)) '<===
    ' ********************************************
End Sub

Private Sub txtCodigo_GotFocus(Index As Integer)
    ConseguirFoco txtcodigo(Index), 3
End Sub


Private Sub txtCodigo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub txtCodigo_KeyPress(Index As Integer, KeyAscii As Integer)
'15/02/2007
'    KEYpress KeyAscii
'ahora
    If KeyAscii = teclaBuscar Then
        Select Case Index
            Case 10: KEYFecha KeyAscii, 1 'fecha
            
            Case 16: KEYFecha KeyAscii, 2 'fecha desde
            Case 17: KEYFecha KeyAscii, 3 'fecha hasta
        End Select
    Else
        KEYpress KeyAscii
    End If

End Sub

Private Sub KEYBusqueda(KeyAscii As Integer, indice As Integer)
    KeyAscii = 0
    imgBuscar_Click (indice)
End Sub

Private Sub KEYFecha(KeyAscii As Integer, indice As Integer)
    KeyAscii = 0
    imgFecha_Click (indice)
End Sub


Private Sub txtCodigo_LostFocus(Index As Integer)
Dim Cad As String, cadTipo As String 'tipo cliente

    'Quitar espacios en blanco por los lados
    txtcodigo(Index).Text = Trim(txtcodigo(Index).Text)
'    If txtCodigo(Index).Text = "" Then Exit Sub
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    Select Case Index
        Case 18, 19  ' Nro. Factura
            PonerFormatoEntero txtcodigo(Index)
            
        Case 10, 16, 17 'FECHAS
            If txtcodigo(Index).Text <> "" Then
                If PonerFormatoFecha(txtcodigo(Index)) Then
                End If
            End If
            
    End Select
End Sub

Private Sub AbrirFrmSocios(indice As Integer)
    indCodigo = indice
    Set frmSoc = New frmManSocios
    frmSoc.DatosADevolverBusqueda = "0|1|"
'    frmCli.DeConsulta = True
'    frmCli.CodigoActual = txtCodigo(indCodigo)
    frmSoc.Show vbModal
    Set frmSoc = Nothing
End Sub


Private Sub InicializarVbles()
    cadFormula = ""
    cadSelect = ""
    CadParam = ""
    numParam = 0
End Sub

Private Function PonerDesdeHasta(codD As String, codH As String, nomD As String, nomH As String, param As String) As Boolean
'IN: codD,codH --> codigo Desde/Hasta
'    nomD,nomH --> Descripcion Desde/Hasta
'Añade a cadFormula y cadSelect la cadena de seleccion:
'       "(codigo>=codD AND codigo<=codH)"
' y añade a cadParam la cadena para mostrar en la cabecera informe:
'       "codigo: Desde codD-nomd Hasta: codH-nomH"
Dim devuelve As String
Dim devuelve2 As String

    PonerDesdeHasta = False
    devuelve = CadenaDesdeHasta(codD, codH, Codigo, TipCod)
    If devuelve = "Error" Then Exit Function
    If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Function
    If TipCod <> "F" Then 'Fecha
        If Not AnyadirAFormula(cadSelect, devuelve) Then Exit Function
    Else
        devuelve2 = CadenaDesdeHastaBD(codD, codH, Codigo, TipCod)
        If devuelve2 = "Error" Then Exit Function
        If Not AnyadirAFormula(cadSelect, devuelve2) Then Exit Function
    End If
    If devuelve <> "" Then
        If param <> "" Then
            'Parametro Desde/Hasta
            CadParam = CadParam & AnyadirParametroDH(param, codD, codH, nomD, nomH)
            numParam = numParam + 1
        End If
        PonerDesdeHasta = True
    End If
End Function

Private Sub LlamarImprimir()
    With frmImprimir
        .FormulaSeleccion = cadFormula
        .OtrosParametros = CadParam
        .NumeroParametros = numParam
        .SoloImprimir = False
        .Titulo = cadTitulo
        .NombreRPT = cadNombreRPT
        .EnvioEMail = False
        .ConSubInforme = ConSubInforme
        .Opcion = 0
        .Show vbModal
    End With
End Sub

Private Sub AbrirVisReport()
    Screen.MousePointer = vbHourglass
    CadenaDesdeOtroForm = ""
    With frmVisReport
        .FormulaSeleccion = cadFormula
'        .SoloImprimir = (Me.OptVisualizar(indFrame).Value = 1)
        .OtrosParametros = CadParam
        .NumeroParametros = numParam
        '##descomen
'        .MostrarTree = MostrarTree
'        .Informe = MIPATH & Nombre
'        .InfConta = InfConta
        '##
        
'        If NombreSubRptConta <> "" Then
'            .SubInformeConta = NombreSubRptConta
'        Else
'            .SubInformeConta = ""
'        End If
        '##descomen
'        .ConSubInforme = ConSubInforme
        '##
        .Opcion = 0
'        .ExportarPDF = (chkEMAIL.Value = 1)
        .Show vbModal
    End With
    
'    If Me.chkEMAIL.Value = 1 Then
'    '####Descomentar
'        If CadenaDesdeOtroForm <> "" Then frmEMail.Show vbModal
'    End If
    Unload Me
End Sub

Private Sub AbrirEMail()
    If CadenaDesdeOtroForm <> "" Then frmEMail.Show vbModal
End Sub


Private Sub ProcesoFacturacionRectificativa(nTabla As String, cadSelect As String)
Dim cOrden As String
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim numOp As Byte

Dim cadDesde As Date
Dim cadhasta As Date
Dim cadAux As String

Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String


Dim Nregs As Long
Dim FecFac As Date

Dim b As Boolean
Dim Sql2 As String



    cadNombreRPT = "rResumFacturas.rpt"
    
    cadTitulo = "Resumen de Facturas"
                    
    'Comprobar si hay registros a Mostrar antes de abrir el Informe
    If HayRegParaInforme(nTabla, cadSelect) Then
        'comprobamos que los tipos de iva existen en la contabilidad de horto
                
        Nregs = TotalRegistrosConsulta("select * from " & nTabla & " where " & cadSelect)
        If Nregs <> 0 Then
                Me.Pb1.visible = True
                Me.Pb1.Max = Nregs
                Me.Pb1.Value = 0
                Me.Refresh
                b = FacturacionRectificativa(nTabla, cadSelect, txtcodigo(10).Text, Me.Pb1)
                If b Then
                    MsgBox "Proceso realizado correctamente.", vbExclamation
                                   
                    'IMPRESION DEL RESUMEN DE LA FACTURACION DE rectificativas
                    If Me.Check1(1).Value Then
                        cadFormula = ""
                        CadParam = CadParam & "pFecFac= """ & txtcodigo(10).Text & """|"
                        numParam = numParam + 1
                        CadParam = CadParam & "pTitulo= ""Resumen Facturación Rectificativa""|"
                        numParam = numParam + 1
                        
                        FecFac = CDate(txtcodigo(10).Text)
                        cadAux = "{rfactsoc.fecfactu}= Date(" & Year(FecFac) & "," & Month(FecFac) & "," & Day(FecFac) & ")"
                        If Not AnyadirAFormula(cadFormula, cadAux) Then Exit Sub
                        If Not AnyadirAFormula(cadFormula, "{tmpinformes.codusu} = " & vUsu.Codigo) Then Exit Sub
                        ConSubInforme = False
                        
                        LlamarImprimir
                    End If
                    'IMPRESION DE LAS FACTURAS RESULTANTES DE LA FACTURACION DE contadores de pozos
                    If Me.Check1(0).Value Then
                        cadFormula = "{rfactsoc.codtipom}='FRS'"
                        cadSelect = "{rfactsoc.codtipom}='FRS'"
                        'Nº Factura
                        cadAux = "({rfactsoc.numfactu} IN [" & FacturasGeneradasRectificativas & "])"
                        If Not AnyadirAFormula(cadFormula, cadAux) Then Exit Sub
                        cadAux = Replace(Replace(cadAux, "]", ")"), "[", "(")
                        If Not AnyadirAFormula(cadSelect, cadAux) Then Exit Sub

                        'Fecha de Factura
                        FecFac = CDate(txtcodigo(10).Text)
                        cadAux = "{rfactsoc.fecfactu}= Date(" & Year(FecFac) & "," & Month(FecFac) & "," & Day(FecFac) & ")"
                        If Not AnyadirAFormula(cadFormula, cadAux) Then Exit Sub
                        cadAux = "{rfactsoc.fecfactu}='" & Format(FecFac, FormatoFecha) & "'"
                        If Not AnyadirAFormula(cadSelect, cadAux) Then Exit Sub

                        Select Case Mid(Combo1(0).Text, 1, 3)
                            Case "FLI"
                                indRPT = 38 'Impresion de Factura Socio de Industria
                            Case Else
                                Tipo = DevuelveValor("select tipodocu from usuarios.stipom where codtipom = " & DBSet(Mid(Combo1(0).Text, 1, 3), "T"))
                                If Tipo >= 7 And Tipo <= 10 Then
                                    indRPT = 42 'Imporesion de Facturas de Bodega o Almazara
                                Else
                                    indRPT = 23 'Impresion de Factura Socio
                                End If
                        End Select
                        
                        If indRPT = 23 And (vParamAplic.Cooperativa = 2 Or vParamAplic.Cooperativa = 16) Then
                            Dim PrecioApor As Double
                            PrecioApor = DevuelveValor("select min(precio) from raporreparto")
                            
                            CadParam = CadParam & "pPrecioApor=""" & Replace(Format(PrecioApor, "#0.000000"), ",", ".") & """|"
                            numParam = numParam + 1
                        End If
                        
                        
                        If Not PonerParamRPT(indRPT, CadParam, numParam, nomDocu) Then Exit Sub
                        'Nombre fichero .rpt a Imprimir
                        cadNombreRPT = nomDocu
                        'Nombre fichero .rpt a Imprimir
                        cadTitulo = "Reimpresión de Facturas "
                        ConSubInforme = True

                        LlamarImprimir

                        If frmVisReport.EstaImpreso Then
'                            ActualizarRegistrosFac "rrecibpozos", cadSelect
                        End If
                    End If
                    'SALIR DE LA FACTURACION DE RECIBOS DE CONTADORES
                    cmdCancel_Click (0)
                End If
            Else
                MsgBox "No hay contadores a facturar.", vbExclamation
            End If
    End If
End Sub


Private Function FacturacionRectificativa(cTabla As String, cWhere As String, FecFac As String, Pb1 As ProgressBar) As Boolean
Dim Sql As String
Dim Sql2 As String
Dim Rs As ADODB.Recordset
Dim Rs2 As ADODB.Recordset

Dim AntSocio As String
Dim ActSocio As String

Dim HayReg As Boolean

Dim NumError As Long
Dim vImporte As Currency
Dim vPorcIva As String

Dim PrimFac As String
Dim UltFac As String

Dim tipoMov As String
Dim b As Boolean
Dim vSeccion As CSeccion
Dim Importe As Currency

Dim devuelve As String
Dim Existe As Boolean

Dim Neto As Currency
Dim Recolect As Byte
Dim vPrecio As Currency
Dim PorcIva As Currency
Dim vTipoMov As CTiposMov
Dim numfactu As Long
Dim ImpoIva As Currency
Dim baseimpo As Currency
Dim TotalFac As Currency


Dim ConsumoHan As Currency
Dim Acciones As Currency
Dim Consumo1 As Long
Dim Consumo2 As Long

Dim ConsTra1 As Long
Dim ConsTra2 As Long

Dim Consumo As Long
Dim ConsumoHidrante As Long

Dim Tipo As Integer


    On Error GoTo eFacturacion

    FacturacionRectificativa = False
    
    'comprobamos que no haya nadie facturando
    DesBloqueoManual ("RECFAC") 'RECtificativas FACturas
    If Not BloqueoManual("RECFAC", "1") Then
        MsgBox "No se puede facturar. Hay otro usuario facturando.", vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If
    
    '------------------------------------------------------------------------------
    '  LOG de acciones
    Set LOG = New cLOG
    LOG.Insertar 2, vUsu, "Facturas Rectificativas: " & vbCrLf & cTabla & vbCrLf & cWhere
    Set LOG = Nothing
    '-----------------------------------------------------------------------------

    
    'Bloqueamos todos los albaranes que vamos a facturar (cabeceras y lineas)
    'Nota: esta bloqueando tambien los registros de la tabla clientes: sclien correspondientes
    If Not BloqueaRegistro(cTabla, cadSelect) Then
        Screen.MousePointer = vbDefault
        'comprobamos que no haya nadie facturando
        DesBloqueoManual ("RECFAC")
        Exit Function
    End If
    
    tipoMov = "FRS"
    
    conn.BeginTrans
    
    Sql = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute Sql

    Sql = "SELECT * "
    Sql = Sql & " FROM  " & cTabla

    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        Sql = Sql & " WHERE " & cWhere
    End If
    
    ' ordenado por socio, variedad, campo, calidad
    Sql = Sql & " order by rfactsoc.codtipom, rfactsoc.numfactu, rfactsoc.fecfactu "
    
    Set vTipoMov = New CTiposMov
    
    HayReg = False
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    b = True
    
    While Not Rs.EOF And b
        HayReg = True
        
        numfactu = vTipoMov.ConseguirContador(tipoMov)
        Do
            numfactu = vTipoMov.ConseguirContador(tipoMov)
            devuelve = DevuelveDesdeBDNew(cAgro, "rfactsoc", "numfactu", "codtipom", tipoMov, "T", , "numfactu", CStr(numfactu), "N", "fecfactu", FecFac, "F")
            If devuelve <> "" Then
                'Ya existe el contador incrementarlo
                Existe = True
                vTipoMov.IncrementarContador (tipoMov)
                numfactu = vTipoMov.ConseguirContador(tipoMov)
            Else
                Existe = False
            End If
        Loop Until Not Existe
            
        IncrementarProgresNew Pb1, 1
        
        'insertar en la tabla de cabecera de facturas
        Sql = "insert into rfactsoc (codtipom,numfactu,fecfactu,codsocio,baseimpo,tipoiva,porc_iva,imporiva,tipoirpf,basereten,porc_ret,"
        Sql = Sql & "impreten,baseaport,porc_apo,impapor,totalfac,impreso,contabilizado,pasaridoc,esanticipogasto,"
        Sql = Sql & " rectif_codtipom,rectif_numfactu,rectif_fecfactu,rectif_motivo "
        '[Monica]14/06/2013: Añadidos los campos que faltaban
        If vParamAplic.Cooperativa = 12 Then
            Sql = Sql & ", esretirada, codforpa, porccorredor, tipoprecio) values ("
        Else
            Sql = Sql & ") values ("
        End If
        Sql = Sql & DBSet(tipoMov, "T") & "," & DBSet(numfactu, "N") & "," & DBSet(txtcodigo(10).Text, "F") & ","
        Sql = Sql & DBSet(Rs!Codsocio, "N") & ","
        Sql = Sql & DBSet(DBLet(Rs!baseimpo, "N") * (-1), "N") & "," ' baseimponible en negativo
        Sql = Sql & DBSet(Rs!TipoIVA, "N") & ","
        Sql = Sql & DBSet(Rs!Porc_Iva, "N") & ","
        Sql = Sql & DBSet(DBLet(Rs!ImporIva, "N") * (-1), "N") & "," ' importe iva en negativo
        Sql = Sql & DBSet(Rs!TipoIRPF, "N") & ","
        Sql = Sql & DBSet(DBLet(Rs!BaseReten, "N") * (-1), "N", "S") & "," ' base retencion en negativo
        Sql = Sql & DBSet(Rs!porc_ret, "N", "S") & ","
        Sql = Sql & DBSet(DBLet(Rs!ImpReten, "N") * (-1), "N", "S") & "," ' importe de retencion en negativo
        Sql = Sql & DBSet(DBLet(Rs!baseaport, "N") * (-1), "N", "S") & "," ' base de aportacion en negativo
        Sql = Sql & DBSet(Rs!porc_apo, "N", "S") & ","
        Sql = Sql & DBSet(DBLet(Rs!impapor, "N") * (-1), "N", "S") & "," ' importe de aportacion en negativo
        Sql = Sql & DBSet(DBLet(Rs!TotalFac, "N") * (-1), "N") & "," ' total factura en negativo
        
        If vParamAplic.Cooperativa = 12 Then
            Sql = Sql & "0,0,0," & DBSet(Rs!EsAnticipoGasto, "N") & ","
        Else
            Sql = Sql & "0,0,0,0,"
        End If
        
        Sql = Sql & DBSet(Rs!CodTipom, "T") & ","
        Sql = Sql & DBSet(Rs!numfactu, "N") & ","
        Sql = Sql & DBSet(Rs!fecfactu, "F") & ","
        Sql = Sql & DBSet(txtcodigo(9).Text, "T")
        
        If vParamAplic.Cooperativa = 12 Then
            Sql = Sql & "," & DBSet(Rs!esretirada, "N") & ","
            Sql = Sql & DBSet(Rs!Codforpa, "N") & ","
            Sql = Sql & DBSet(Rs!PorcCorredor, "N") & ","
            Sql = Sql & DBSet(Rs!TipoPrecio, "N") & ")"
        Else
            Sql = Sql & ")"
        End If
        
        conn.Execute Sql
            
        ' insertamos en la tabla rfactsoc_variedad
        Sql = "insert into rfactsoc_variedad (codtipom, numfactu, fecfactu,codvarie,codcampo,kilosnet,"
        Sql = Sql & "preciomed,imporvar,descontado,imporgasto) "
        Sql = Sql & " select " & DBSet(tipoMov, "T") & "," & DBSet(numfactu, "N") & "," & DBSet(txtcodigo(10).Text, "F") & ","
        Sql = Sql & "codvarie,codcampo, kilosnet * (-1), preciomed, imporvar * (-1), descontado, imporgasto * (-1) "
        Sql = Sql & " from rfactsoc_variedad "
        Sql = Sql & " where codtipom = " & DBSet(Rs!CodTipom, "T")
        Sql = Sql & " and numfactu = " & DBSet(Rs!numfactu, "N")
        Sql = Sql & " and fecfactu = " & DBSet(Rs!fecfactu, "F")
        
        conn.Execute Sql
            
        ' insertamos en la tabla rfactsoc_albaran
        Sql = "insert into rfactsoc_albaran (codtipom,numfactu,fecfactu,numalbar,fecalbar,codvarie,codcampo,kilosbru,"
        Sql = Sql & "kilosnet,grado,precio, importe, imporgasto) "
        Sql = Sql & " select " & DBSet(tipoMov, "T") & "," & DBSet(numfactu, "N") & "," & DBSet(txtcodigo(10).Text, "F") & ","
        Sql = Sql & "numalbar, fecalbar, codvarie, codcampo,  kilosbru * (-1), kilosnet * (-1), grado, precio, "
        Sql = Sql & "importe * (-1), imporgasto * (-1) from rfactsoc_albaran "
        Sql = Sql & " where codtipom = " & DBSet(Rs!CodTipom, "T")
        Sql = Sql & " and numfactu = " & DBSet(Rs!numfactu, "N")
        Sql = Sql & " and fecfactu = " & DBSet(Rs!fecfactu, "F")
        
        conn.Execute Sql
        
        ' insertamos en la tabla rfactsoc_anticipos
        Sql = "insert into rfactsoc_anticipos (codtipom,numfactu,fecfactu,codtipomanti,numfactuanti,fecfactuanti,"
        Sql = Sql & "codvarieanti,codcampoanti,baseimpo) "
        Sql = Sql & " select " & DBSet(tipoMov, "T") & "," & DBSet(numfactu, "N") & "," & DBSet(txtcodigo(10).Text, "F") & ","
        Sql = Sql & "codtipomanti,numfactuanti,fecfactuanti,codvarieanti,codcampoanti,baseimpo * (-1) "
        Sql = Sql & " from rfactsoc_anticipos "
        Sql = Sql & " where codtipom = " & DBSet(Rs!CodTipom, "T")
        Sql = Sql & " and numfactu = " & DBSet(Rs!numfactu, "N")
        Sql = Sql & " and fecfactu = " & DBSet(Rs!fecfactu, "F")
        
        conn.Execute Sql
                
        ' insertamos en la tabla rfactsoc_calidad
        Sql = "insert into rfactsoc_calidad (codtipom,numfactu,fecfactu,codvarie,codcampo,codcalid,kilosnet,precio,"
        Sql = Sql & "imporcal,preciocalidad,imporcalidad) "
        Sql = Sql & "select " & DBSet(tipoMov, "T") & "," & DBSet(numfactu, "N") & "," & DBSet(txtcodigo(10).Text, "F") & ","
        Sql = Sql & "codvarie, codcampo, codcalid, kilosnet * (-1), precio, imporcal * (-1), preciocalidad, imporcalidad * (-1) "
        Sql = Sql & " from rfactsoc_calidad "
        Sql = Sql & " where codtipom = " & DBSet(Rs!CodTipom, "T")
        Sql = Sql & " and numfactu = " & DBSet(Rs!numfactu, "N")
        Sql = Sql & " and fecfactu = " & DBSet(Rs!fecfactu, "F")
        
        conn.Execute Sql
        
        ' insertamos en la tabla rfactsoc_gastos
        Sql = "insert into rfactsoc_gastos (codtipom,numfactu,fecfactu,numlinea,codgasto,importe) "
        Sql = Sql & "select " & DBSet(tipoMov, "T") & "," & DBSet(numfactu, "N") & "," & DBSet(txtcodigo(10).Text, "F") & ","
        Sql = Sql & "numlinea, codgasto, importe * (-1) "
        Sql = Sql & " from rfactsoc_gastos "
        Sql = Sql & " where codtipom = " & DBSet(Rs!CodTipom, "T")
        Sql = Sql & " and numfactu = " & DBSet(Rs!numfactu, "N")
        Sql = Sql & " and fecfactu = " & DBSet(Rs!fecfactu, "F")
        
        conn.Execute Sql
        
        '[Monica]04/06/2014: en el caso de Montifrut es diferente
        If vParamAplic.Cooperativa = 12 Then
            If DBLet(Rs!EsAnticipoGasto) = 1 Then
                '[Monica]04/06/2014: si la factura que rectifico es un anticipo tengo que marcarlo como que se ha descontado
                ' pq sino en la proxima liquidacion se descontará este anticipo siendo que se ha rectificado.
                Sql = "update rfactsoc_variedad set descontado = 1 where codtipom = " & DBSet(Rs!CodTipom, "T")
                Sql = Sql & " and numfactu = " & DBSet(Rs!numfactu, "N")
                Sql = Sql & " and fecfactu = " & DBSet(Rs!fecfactu, "F")
                
                conn.Execute Sql
            
            Else
                '[Monica]04/06/2014: si la factura que rectifico es una liquidacion que tiene descontados anticipos,
                '                    los he de desmarcar como descontados para que en la proxima liquidacion se descuente
                Sql = "update rfactsoc_variedad  set descontado = 0 where (codtipom,numfactu,fecfactu,codvarie,codcampo) in "
                Sql = Sql & " (select codtipomanti, numfactuanti, fecfactuanti, codvarieanti, codcampoanti from rfactsoc_anticipos where codtipom = " & DBSet(Rs!CodTipom, "T")
                Sql = Sql & " and numfactu = " & DBSet(Rs!numfactu, "N")
                Sql = Sql & " and fecfactu = " & DBSet(Rs!fecfactu, "F") & ")"
                
                conn.Execute Sql
            
            End If
        Else
            '[Monica]08/07/2011: si la factura que rectifico es un anticipo tengo que marcarlo como que se ha descontado
            ' pq sino en la proxima liquidacion se descontará este anticipo siendo que se ha rectificado.
            Tipo = DevuelveValor("select tipodocu from usuarios.stipom where codtipom = " & DBSet(Rs!CodTipom, "T"))
            If Tipo = 1 Or Tipo = 3 Then
                Sql = "update rfactsoc_variedad set descontado = 1 where codtipom = " & DBSet(Rs!CodTipom, "T")
                Sql = Sql & " and numfactu = " & DBSet(Rs!numfactu, "N")
                Sql = Sql & " and fecfactu = " & DBSet(Rs!fecfactu, "F")
                
                conn.Execute Sql
            End If
            '[Monica]08/07/2011
        
            '[Monica]04/06/2014: si la factura que rectifico es una liquidacion que tiene descontados anticipos,
            '                    los he de desmarcar como descontados para que en la proxima liquidacion se descuente
            If Tipo = 2 Or Tipo = 4 Then
                Sql = "update rfactsoc_variedad  set descontado = 0 where (codtipom,numfactu,fecfactu,codvarie,codcampo) in "
                Sql = Sql & " (select codtipomanti, numfactuanti, fecfactuanti, codvarieanti, codcampoanti from rfactsoc_anticipos where codtipom = " & DBSet(Rs!CodTipom, "T")
                Sql = Sql & " and numfactu = " & DBSet(Rs!numfactu, "N")
                Sql = Sql & " and fecfactu = " & DBSet(Rs!fecfactu, "F") & ")"
                
                conn.Execute Sql
            End If
        End If
            
            
        If b Then b = InsertResumen(tipoMov, CStr(numfactu))
        
        If b Then b = vTipoMov.IncrementarContador(tipoMov)
        
        
        Rs.MoveNext
    Wend
    
    Set Rs = Nothing
    
    
eFacturacion:
    If Err.Number <> 0 Or Not b Then
        conn.RollbackTrans
        FacturacionRectificativa = False
    Else
        conn.CommitTrans
        FacturacionRectificativa = True
    End If

End Function

Private Function TotalFacturasSocios(cTabla As String, cWhere As String) As Long
Dim Sql As String

    TotalFacturasSocios = 0
    
    Sql = "SELECT  count(distinct rpozos.codsocio) "
    Sql = Sql & " FROM  " & cTabla

    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        Sql = Sql & " WHERE " & cWhere
    End If

    TotalFacturasSocios = TotalRegistros(Sql)

End Function

Private Function DatosOk() As Boolean
'Comprobar que los datos de la cabecera son correctos antes de Insertar o Modificar
'la cabecera del Pedido
Dim b As Boolean
Dim Sql As String

    On Error GoTo EDatosOK

    DatosOk = False
    b = True
    
    If txtcodigo(10).Text = "" Then
        MsgBox "Debe introducir un valor para la Fecha de Factura Rectificativa.", vbExclamation
        PonerFoco txtcodigo(10)
        b = False
    End If
    
    DatosOk = b
    
    
EDatosOK:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function


Private Function FacturasGeneradasRectificativas() As String
Dim Sql As String
Dim RS1 As ADODB.Recordset
Dim Cad As String
    
    On Error GoTo eFacturasGeneradas

    FacturasGeneradasRectificativas = ""

    Sql = "select nombre1, importe1 from tmpinformes where codusu = " & vUsu.Codigo
    Sql = Sql & " and nombre1 = 'FRS'"
    
    Set RS1 = New ADODB.Recordset
    RS1.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    Cad = ""
    While Not RS1.EOF
        Cad = Cad & DBLet(RS1.Fields(1).Value, "N") & ","
    
        RS1.MoveNext
    Wend
    Set RS1 = Nothing
    
    'si hay facturas quitamos la ultima coma
    If Cad <> "" Then Cad = Mid(Cad, 1, Len(Cad) - 1)
    
    FacturasGeneradasRectificativas = Cad
    Exit Function
    
eFacturasGeneradas:
    MuestraError Err.Number, "Cadena de Facturas Rectificativas Generadas ", Err.Description
End Function


Private Sub CargaCombo()
Dim Rs As ADODB.Recordset
Dim Sql As String
Dim i As Byte
    
    ' *** neteje els combos, els pose valor i seleccione el valor per defecte ***
    For i = 0 To Combo1.Count - 1
        Combo1(i).Clear
    Next i
    
    'tipo de factura todas las facturas excepto las rectificativas
    Sql = "select codtipom, nomtipom from usuarios.stipom where tipodocu >= 1 and tipodocu < 11 "

    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    i = 1
    While Not Rs.EOF
'        Sql = Replace(Rs.Fields(1).Value, "Factura", "Fac.")
        Sql = Rs.Fields(1).Value
        Sql = Rs.Fields(0).Value & " - " & Sql
        Combo1(0).AddItem Sql 'campo del codigo
        Combo1(0).ItemData(Combo1(0).NewIndex) = i
        i = i + 1
        Rs.MoveNext
    Wend
    
End Sub

