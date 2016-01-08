VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmPOZFraRecargo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Facturación con recargo"
   ClientHeight    =   7110
   ClientLeft      =   45
   ClientTop       =   4335
   ClientWidth     =   7740
   Icon            =   "frmPOZFraRecargo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmPOZFraRecargo.frx":000C
   ScaleHeight     =   7110
   ScaleWidth      =   7740
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   7
      Left            =   10380
      MaxLength       =   15
      TabIndex        =   15
      Tag             =   "Base Imponible 2 |N|N|0||scafac|baseimp2|#,###,###,##0.00|N|"
      Text            =   "Text1 7"
      Top             =   3375
      Width           =   1485
   End
   Begin VB.Frame FrameIntro 
      Height          =   1455
      Left            =   120
      TabIndex        =   5
      Top             =   540
      Width           =   7425
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   315
         Index           =   3
         Left            =   1860
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   9
         Text            =   "Text2"
         Top             =   390
         Width           =   5355
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   3
         Left            =   1050
         MaxLength       =   6
         TabIndex        =   2
         Tag             =   "Cod. Transportista|N|N|0|999|tcafpc|codtrans|000|S|"
         Text            =   "Text1"
         Top             =   390
         Width           =   780
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   1
         Left            =   1830
         MaxLength       =   10
         TabIndex        =   1
         Tag             =   "Fecha Factura|F|N|||tcafpc|fecfactu|dd/mm/yyyy|S|"
         Top             =   930
         Width           =   1305
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   0
         Left            =   4290
         MaxLength       =   10
         TabIndex        =   0
         Tag             =   "Nº Factura|T|N|||tcafpc|numfactu||S|"
         Text            =   "Text1 7"
         Top             =   900
         Width           =   1245
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   0
         Left            =   1560
         Picture         =   "frmPOZFraRecargo.frx":0A0E
         ToolTipText     =   "Buscar fecha"
         Top             =   930
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   750
         ToolTipText     =   "Buscar socio"
         Top             =   420
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Socio"
         Height          =   255
         Index           =   0
         Left            =   210
         TabIndex        =   8
         Top             =   420
         Width           =   510
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha Factura"
         Height          =   255
         Index           =   29
         Left            =   240
         TabIndex        =   7
         Top             =   930
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "% Recargo"
         Height          =   255
         Index           =   28
         Left            =   3420
         TabIndex        =   6
         Top             =   930
         Width           =   1095
      End
   End
   Begin VB.Frame FrameAux0 
      Height          =   4830
      Left            =   120
      TabIndex        =   11
      Top             =   2010
      Width           =   7440
      Begin VB.CommandButton cmdCancelar 
         Cancel          =   -1  'True
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   6030
         TabIndex        =   18
         Top             =   4260
         Width           =   1035
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   4830
         TabIndex        =   17
         Top             =   4260
         Width           =   1035
      End
      Begin VB.Frame Frame1 
         Height          =   555
         Index           =   0
         Left            =   180
         TabIndex        =   12
         Top             =   4080
         Width           =   2865
         Begin VB.Label lblIndicador 
            Alignment       =   2  'Center
            Caption         =   "Label2"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   120
            TabIndex        =   13
            Top             =   180
            Width           =   2655
         End
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   3555
         Left            =   150
         TabIndex        =   16
         Top             =   540
         Width           =   7065
         _ExtentX        =   12462
         _ExtentY        =   6271
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Tipo"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Factura"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Descripcion"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Fecha"
            Object.Width           =   1942
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Importe"
            Object.Width           =   3706
         EndProperty
      End
      Begin VB.Label Label2 
         Caption         =   "Facturas Pendientes de Pago"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   180
         TabIndex        =   14
         Top             =   240
         Width           =   4665
      End
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   5730
      Top             =   5520
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   7740
      _ExtentX        =   13653
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Pedir Datos"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Generar Factura"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
         EndProperty
      EndProperty
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   195
         Left            =   6060
         TabIndex        =   4
         Top             =   60
         Visible         =   0   'False
         Width           =   1215
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Operador"
      Height          =   255
      Index           =   1
      Left            =   1845
      TabIndex        =   10
      Top             =   900
      Width           =   735
   End
   Begin VB.Image imgBuscar 
      Height          =   240
      Index           =   1
      Left            =   1845
      Picture         =   "frmPOZFraRecargo.frx":0A99
      ToolTipText     =   "Buscar trabajador"
      Top             =   1125
      Width           =   240
   End
   Begin VB.Menu mnOpciones 
      Caption         =   "&Opciones"
      Begin VB.Menu mnPedirDatos 
         Caption         =   "&Pedir Datos"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnGenerarFac 
         Caption         =   "&Generar Factura"
      End
      Begin VB.Menu mnBarra1 
         Caption         =   "-"
      End
      Begin VB.Menu mnSalir 
         Caption         =   "&Salir"
         Shortcut        =   ^S
      End
   End
End
Attribute VB_Name = "frmPOZFraRecargo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'========== VBLES PUBLICAS ====================
Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)


'========== VBLES PRIVADAS ====================
Private WithEvents frmF As frmCal 'Form Calendario Fecha
Attribute frmF.VB_VarHelpID = -1

Private WithEvents frmSoc As frmManSocios
Attribute frmSoc.VB_VarHelpID = -1
Private WithEvents frmVar As frmComVar  'Form Mto clientes
Attribute frmVar.VB_VarHelpID = -1
Private WithEvents frmBanPr As frmComBanco 'Mto de Bancos propios
Attribute frmBanPr.VB_VarHelpID = -1
Private WithEvents frmFPa As frmComFpa 'Mto de formas de pago
Attribute frmFPa.VB_VarHelpID = -1
'Private WithEvents frmCtas As frmCtasConta 'Cuentas contables

Private Modo As Byte
'-----------------------------
'Se distinguen varios modos
'   0.-  Formulario limpio sin nungun campo rellenado
'   1.-  Preparando para hacer la busqueda
'   2.-  Ya tenemos registros y los vamos a recorrer
'        y podemos editarlos Edicion del campo
'   3.-  Insercion de nuevo registro
'   4.-  Modificar
'   5.-  Mantenimiento Lineas
'-------------------------------------------------------------------------
'-------------------------------------------------------------------------

'cadena donde se almacena la WHERE para la seleccion de los albaranes
'marcados para facturar
Dim cadWHERE As String

'Cuando vuelve del formulario de ver los albaranes seleccionados hay que volver
'a cargar los datos de los albaranes
'Dim VerAlbaranes As Boolean

Dim PrimeraVez As Boolean
Dim Bloquear As Boolean
Dim indice As Integer

'SQL de la tabla principal del formulario
Private CadenaConsulta As String

Private NombreTabla As String  'Nombre de la tabla o de la
Private kCampo As Integer
'-------------------------------------------------------------------------

Private vSocio As cSocio

Dim ModoLineas As Byte
'1.- Afegir,  2.- Modificar,  3.- Borrar,  0.-Passar control a Llínies
Dim NumTabMto As Integer 'Indica quin nº de Tab està en modo Mantenimient

Dim vWhere As String


Dim ModificaDescuento As Boolean



Private Sub Check1_LostFocus(Index As Integer)
    If Index = 1 Then
        If Check1(1).Value = 1 Then
            If vParamAplic.CodIvaIntra = 0 Then
                MsgBox "No tiene asignado el código de Iva Intracomunitario en parámetros. Revise.", vbExclamation
                Check1(1).Value = 0
            End If
        End If
    End If
End Sub

Private Sub DataGridAux_RowColChange(Index As Integer, LastRow As Variant, ByVal LastCol As Integer)
    If Not AdoAux(0).Recordset.EOF Then _
        Me.lblIndicador.Caption = AdoAux(0).Recordset.AbsolutePosition & " de " & AdoAux(0).Recordset.RecordCount
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
'    If VerAlbaranes Then RefrescarAlbaranes
'    VerAlbaranes = False
End Sub

Private Sub Form_Load()
Dim I As Integer
'    'Icono del formulario
'    Me.Icon = frmPpal.Icon
    
    ' ICONITOS DE LA BARRA
    With Me.Toolbar1
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 1   'Pedir Datos
        .Buttons(2).Image = 4   'Ver albaranes
        .Buttons(3).Image = 15   'Generar FActura
        .Buttons(6).Image = 11   'Salir
    End With
    
    ' ******* si n'hi han llínies *******
    'ICONETS DE LES BARRES ALS TABS DE LLÍNIA
    For I = 0 To ToolAux.Count - 1
        With Me.ToolAux(I)
            .HotImageList = frmPpal.imgListComun_OM16
            .DisabledImageList = frmPpal.imgListComun_BN16
            .ImageList = frmPpal.imgListComun16
            .Buttons(1).Image = 3   'Insertar
            .Buttons(2).Image = 4   'Modificar
            .Buttons(3).Image = 5   'Borrar
        End With
    Next I
    ' ***********************************
    
    
    
    'cargar IMAGES de busqueda
    Me.imgBuscar(0).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Me.imgBuscar(2).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Me.imgBuscar(3).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Me.imgBuscar(4).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    
    Me.FrameFactura.Enabled = False
    
    LimpiarCampos   'Limpia los campos TextBox
'    InicializarListView
   
    '## A mano
    NombreTabla = "rhisfruta" ' albaranes de venta
    
    'Vemos como esta guardado el valor del check
'    chkVistaPrevia.Value = CheckValueLeer(Name)
    
    'ASignamos un SQL al DATA1
    CadenaConsulta = "Select * from " & NombreTabla
    CadenaConsulta = CadenaConsulta & " where numalbar=-1"
    
    Data1.ConnectionString = conn
    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
        
    'Cargar el DataGrid de lineas de Revisiones inicialmente a nada DATA2
    PrimeraVez = True
    
    If DatosADevolverBusqueda = "" Then
        PonerModo 0
    End If
    CargaGrid 0, False
    
    PrimeraVez = False
End Sub


Private Sub LimpiarCampos()
On Error Resume Next
    limpiar Me
    'Aqui va el especifico de cada form es
    '### a mano
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub Form_Unload(Cancel As Integer)
    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
    DesBloqueoManual "FACTRA"
    TerminaBloquear
'    DesBloqueoManual ("scaalp")
End Sub


Private Sub frmBanPr_DatoSeleccionado(CadenaSeleccion As String)
    'Form de Mantenimiento de Bancos Propios
    Text1(5).Text = RecuperaValor(CadenaSeleccion, 1)
    Text1(5).Text = Format(Text1(5).Text, "0000")
    Text2(5).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmF_Selec(vFecha As Date)
Dim indice As Byte
    indice = CByte(Me.imgFecha(0).Tag)
    Text1(indice).Text = Format(vFecha, "dd/mm/yyyy")
End Sub


Private Sub frmFPa_DatoSeleccionado(CadenaSeleccion As String)
    Text1(4).Text = RecuperaValor(CadenaSeleccion, 1) 'Cod forpa
    FormateaCampo Text1(4)
    Text2(4).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom forpa
End Sub

Private Sub frmSoc_DatoSeleccionado(CadenaSeleccion As String)
Dim indice As Byte
    indice = 3
    Text1(indice).Text = RecuperaValor(CadenaSeleccion, 1) 'Cod Socios
    FormateaCampo Text1(indice)
    Text2(indice).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom socio
End Sub

Private Sub imgBuscar_Click(Index As Integer)
Dim indice As Byte

    If Modo = 2 Or Modo = 0 Then Exit Sub
'    Screen.MousePointer = vbHourglass

    Select Case Index
        Case 0 'Cod. Proveedor
            Set frmSoc = New frmManSocios
            frmSoc.DatosADevolverBusqueda = "0|1|"
            frmSoc.Show vbModal
            Set frmSoc = Nothing
            indice = 3
       
       Case 2 'Bancos Propios
            indice = 5
            Set frmBanPr = New frmComBanco
            frmBanPr.DatosADevolverBusqueda = "0|1|"
            frmBanPr.Show vbModal
            Set frmBanPr = Nothing
            
       Case 3 'Variedades
            Set frmVar = New frmComVar
            frmVar.DatosADevolverBusqueda = "0|2|"
            frmVar.Show vbModal
            Set frmVar = Nothing
            indice = 26
       
       Case 4 'Forma de pago
            Set frmFPa = New frmComFpa
            frmFPa.DatosADevolverBusqueda = "0|1|"
            frmFPa.Show vbModal
            Set frmFPa = Nothing
            indice = 4
       
    End Select
    
    PonerFoco Text1(indice)
'    Screen.MousePointer = vbDefault
End Sub

Private Sub imgFecha_Click(Index As Integer)
Dim indice As Byte
Dim esq As Long
Dim dalt As Long
Dim menu As Long
Dim obj As Object

   If Modo = 2 Or Modo = 0 Then Exit Sub
   Screen.MousePointer = vbHourglass
    
   Set frmF = New frmCal
    
   esq = imgFecha(Index).Left
   dalt = imgFecha(Index).Top
    
   Set obj = imgFecha(Index).Container

   While imgFecha(Index).Parent.Name <> obj.Name
       esq = esq + obj.Left
       dalt = dalt + obj.Top
       Set obj = obj.Container
   Wend
    
   menu = Me.Height - Me.ScaleHeight 'ací tinc el heigth del menú i de la toolbar

   frmF.Left = esq + imgFecha(Index).Parent.Left + 30
   frmF.Top = dalt + imgFecha(Index).Parent.Top + imgFecha(Index).Height + menu - 40
   
   frmF.NovaData = Now
   indice = Index + 1
   Me.imgFecha(0).Tag = indice
   
   PonerFormatoFecha Text1(indice)
   If Text1(indice).Text <> "" Then frmF.NovaData = CDate(Text1(indice).Text)

   Screen.MousePointer = vbDefault
   frmF.Show vbModal
   Set frmF = Nothing
   PonerFoco Text1(indice)
End Sub

Private Sub mnModificarDto_Click()
Dim I As Integer


    If Text1(0).Text = "" Then Exit Sub

    PonerModo 4

    Me.FrameFactura.Enabled = True
    
    ' *** bloqueje la clau primaria de la capçalera ***
    BloquearTxt Text1(0), True
    ' *********************************
    
    BloquearTxt Text1(6), True
    BloquearTxt Text1(8), False
    
    For I = 9 To 22
        BloquearTxt Text1(I), True
    Next I
    
    lblIndicador.Caption = "MODIFICA DESCUENTO"
    
    Me.FrameFactura.Enabled = True
    
    PonerFoco Text1(8)
 
    
End Sub

Private Sub mnGenerarFac_Click()
    BotonFacturar
    Set vSocio = Nothing
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnPedirDatos_Click()
    BotonPedirDatos
End Sub

Private Sub mnSalir_Click()
    Screen.MousePointer = vbDefault
    Unload Me
End Sub


'Private Sub mnVerAlbaran_Click()
'    BotonVerAlbaranes
'End Sub

Private Sub Text1_GotFocus(Index As Integer)
    kCampo = Index
    ConseguirFoco Text1(Index), Modo
End Sub


Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
    KEYdown KeyCode
End Sub


Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


'----------------------------------------------------------------
'----------------------------------------------------------------
' Cunado el campo de texto pierde el enfoque
' Es especifico de cada formulario y en el podremos controlar
' lo que queramos, desde formatear un campo si asi lo deseamos
' hasta pedir que nos devuelva los datos de la empresa
'----------------------------------------------------------------
'----------------------------------------------------------------
Private Sub Text1_LostFocus(Index As Integer)

    If Index <> 8 And Not PerderFocoGnral(Text1(Index), Modo) Then Exit Sub
        
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
          
    'Si queremos hacer algo ..
    Select Case Index
        Case 1 'Fecha factura
            PonerFormatoFecha Text1(Index)
            If Text1(Index) <> "" Then
                ' No debe existir el número de factura para el proveedor en hco
                If ExisteFacturaEnHco Then
'                    InicializarListView
                End If
            End If
            
        Case 3 'Cod Socios
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), "rsocios", "nomsocio", "codsocio")
                
                If Text2(Index).Text <> "" Then CargarFacturas Text1(Index)
            Else
                Text2(Index).Text = ""
            End If
            
        Case 0
            PonerFormatoDecimal Text1(Index), 4
        
        
    End Select
End Sub


'----------------------------------------------------------------
'----------------------------------------------------------------
'   En PONERMODO se habilitan, o no, los diverso campos del
'   formulario en funcion del modo en k vayamos a trabajar
Private Sub PonerModo(Kmodo As Byte)
Dim I As Byte, Numreg As Byte
Dim b As Boolean
On Error GoTo EPonerModo

    Modo = Kmodo
    
    'Modo 2. Hay datos y estamos visualizandolos
    '=========================================
    b = (Modo = 2)
        
    cmdAceptar.visible = (ModoLineas = 2)
    cmdAceptar.Enabled = (ModoLineas = 2)
    cmdCancelar.visible = (ModoLineas = 2)
    cmdCancelar.Enabled = (ModoLineas = 2)
    
'    'Bloquea los campos Text1 sino estamos modificando/Insertando Datos
'    'Si estamos en Insertar además limpia los campos Text1
'    'si estamos en modificar bloquea las compos que son clave primaria
'    BloquearText1 Me, Modo
    
    For I = 0 To Text1.Count - 1
        BloquearTxt Text1(I), (Modo <> 3)
    Next I
    
    'Importes siempre bloqueados
    For I = 6 To 25
        BloquearTxt Text1(I), True
    Next I
    'Campo B.Imp y Imp. IVA siempre en azul
    Text1(9).BackColor = &HFFFFC0 'Base imponible
    Text1(19).BackColor = &HFFFFC0 'Total Iva 1
    Text1(20).BackColor = &HFFFFC0 'Iva 2
    Text1(21).BackColor = &HFFFFC0 'IVa 3
    Text1(22).BackColor = &HC0C0FF 'Total factura
    Text1(25).BackColor = &HFFFFC0 'Imp.Retencion
    
    
    '---------------------------------------------
    b = (Modo <> 0 And Modo <> 2 And Modo <> 5)
    
    For I = 0 To Me.imgBuscar.Count - 1
        Me.imgBuscar(I).Enabled = b
    Next I
                    
    Me.chkVistaPrevia.Enabled = (Modo <= 2)
    
    For I = 0 To txtAux.Count - 1
        BloquearTxt txtAux(I), True
        txtAux(I).visible = False
    Next I
        
    Me.FrameIntro.Enabled = (Modo = 3)
    Me.FrameAux0.Enabled = (Modo = 5)
       
    Text2(2).visible = False
       
    'Poner el tamaño de los campos. Si es modo Busqueda el MaxLength del campo
    'debe ser mayor para adminir intervalos de busqueda.
    PonerLongCampos
'    PonerModoOpcionesMenu (Modo) 'Activar opciones de menu según modo
    PonerOpcionesMenu 'Activar opciones de menu según nivel de permisos del usuario
    
EPonerModo:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub PonerLongCampos()
'Modificar el MaxLength del campo en funcion de si es modo de búsqueda o no
'para los campos que permitan introducir criterios más largos del tamaño del campo
    PonerLongCamposGnral Me, Modo, 1
End Sub


Private Function DatosOk() As Boolean
'Comprobar que los datos del frame de introduccion son correctos antes de cargar datos
Dim vtag As CTag
Dim cad As String
Dim I As Byte
Dim vSeccion As CSeccion

    On Error GoTo EDatosOK
    DatosOk = False
    
    ' deben de introducirse todos los datos del frame
    For I = 0 To 5
        If Text1(I).Text = "" Then
            If Text1(I).Tag <> "" Then
                Set vtag = New CTag
                If vtag.Cargar(Text1(I)) Then
                    cad = vtag.Nombre
                Else
                    cad = "Campo"
                End If
                Set vtag = Nothing
            Else
                cad = "Campo"
                If I = 5 Then cad = "Cta. Prev. Pago"
                If I = 4 Then cad = "Forma de Pago"
            End If
            MsgBox cad & " no puede estar vacio. Reintroduzca", vbExclamation
            PonerModo 3
            PonerFoco Text1(I)
            Exit Function
        End If
    Next I
        
    'comprobar que la fecha de la factura sea anterior a la fecha de recepcion
    If Not EsFechaIgualPosterior(Text1(1).Text, Text1(2).Text, True, "La fecha de recepción debe ser igual o posterior a la fecha de la factura.") Then
        Exit Function
    End If
    
    'Comprobar que la fecha de RECEPCION esta dentro de los ejercicios contables
    Set vSeccion = New CSeccion
    If vSeccion.LeerDatos(vParamAplic.Seccionhorto) Then
        If vSeccion.AbrirConta Then
            I = EsFechaOKConta(CDate(Text1(2).Text))
            If I > 0 Then
                MsgBox "Fecha fuera ejercicios contables", vbExclamation
                vSeccion.CerrarConta
                Set vSeccion = Nothing
                Exit Function
            End If
        End If
    End If
    vSeccion.CerrarConta
    Set vSeccion = Nothing

'    If vParamAplic.NumeroConta <> 0 Then
'        i = EsFechaOKConta(CDate(Text1(2).Text))
'        If i > 0 Then
'            'If i = 1 Then
'                MsgBox "Fecha fuera ejercicios contables", vbExclamation
'                Exit Function
'           ' Else
'           '     cad = "La fecha es superior al ejercico contable siguiente. ¿Desea continuar?"
'           '     If MsgBox(cad, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Function
'           ' End If
'        End If
'    End If
    
'--monica:03/12/2008
'    'comprobar que se han seleccionado lineas para facturar
'    If cadWHERE = "" Then
'        MsgBox "Debe seleccionar albaranes para facturar.", vbExclamation
'        Exit Function
'    End If
    
'++monica:03/12/2008
    'comprobamos que hay lineas para facturar: o albaranes o portes de vuelta
    If cadWHERE = "" Then
        If AdoAux(0).Recordset.RecordCount = 0 Then
            MsgBox "No hay albaranes para incluir en la factura. Revise.", vbExclamation
            Exit Function
        End If
    End If
    
    ' No debe existir el número de factura para el socio tercero en hco
    If ExisteFacturaEnHco Then Exit Function
    
'--monica
'    'todos los albaranes seleccionados deben tener la misma: forma pago, dto ppago, dto gnral
'    cad = "select count(distinct codforpa,dtoppago,dtognral) from scaalp "
'    cad = cad & " WHERE " & Replace(cadWHERE, "slialp", "scaalp")
'    If RegistrosAListar(cad) > 1 Then
'        MsgBox "No se puede facturar albaranes con distintas: forma de pago, dto gral, dto ppago.", vbExclamation
'        Exit Function
'    End If
    
    
    'Si la forpa es TRANSFERENCIA entonces compruebo la si tiene cta bancaria
'    cad = "select distinct (codforpa) from scaalp "
'    cad = cad & " WHERE " & Replace(cadWHERE, "slialp", "scaalp")
    Set miRsAux = New ADODB.Recordset
'    miRsAux.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'    cad = miRsAux.Fields(0)
'    miRsAux.Close
    
    
    
    'Ahora buscamos el tipforpa del codforpa
    cad = "Select tipoforp from forpago where codforpa=" & DBSet(Text1(4).Text, "N")
    miRsAux.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    I = 0
    If miRsAux.EOF Then
        MsgBox "Error en el TIPO de forma de pago", vbExclamation
    Else
        I = 1
        cad = miRsAux.Fields(0)
        If Val(cad) = vbFPTransferencia Then
            'Compruebo que la forpa es transferencia
            I = 2
        End If
    End If
    miRsAux.Close
    Set miRsAux = Nothing
    
    
    If I = 2 Then
        'La forma de pago es transferencia. Debo comprobar que existe la cuenta bancaria
        'del proveedor
        If vSocio.CuentaBan = "" Or vSocio.Digcontrol = "" Or vSocio.Sucursal = "" Or vSocio.Banco = "" Then
            cad = "Cuenta bancaria incorrecta. Forma de pago: transferencia.    ¿Continuar?"
            If MsgBox(cad, vbQuestion + vbYesNoCancel) <> vbYes Then I = 0
        End If
    End If
    
    'Si i=0 es que o esta mal la forpa o no quiere seguir pq no tiene cuenta bancaria
    If I > 0 Then DatosOk = True
    Exit Function
    
EDatosOK:
    DatosOk = False
    MuestraError Err.Number, "Comprobar datos correctos", Err.Description
End Function



Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1  'Pedir datos
             mnPedirDatos_Click
            
        Case 2
             mnModificarDto_Click
            
        Case 3 'Generar Factura
            mnGenerarFac_Click

        Case 6    'Salir
            mnSalir_Click
    End Select
End Sub


Private Sub PonerOpcionesMenu()
Dim J As Byte

    PonerOpcionesMenuGeneral Me
    
    J = Val(Me.mnPedirDatos.HelpContextID)
    If J < vUsu.Nivel Then Me.mnPedirDatos.Enabled = False
    
    J = Val(Me.mnGenerarFac.HelpContextID)
    If J < vUsu.Nivel Then Me.mnGenerarFac.Enabled = False
End Sub


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me
End Sub

 
Private Sub BotonPedirDatos()
Dim Nombre As String

    TerminaBloquear

    'Vaciamos todos los Text
    LimpiarCampos
    Check1(1).Value = 0
    'Vaciamos el ListView
'    InicializarListView
    CargaGrid 0, False
    
    'Como no habrá albaranes seleccionados vaciamos la cadwhere
    cadWHERE = ""
    
    PonerModo 3
    
    'fecha recepcion
    Text1(2).Text = Format(Now, "dd/mm/yyyy")
    
    
    'desbloquear los registros de la saalp (si hay bloquedos)
    TerminaBloquear
    
    'si vamos
    'desBloqueo Manual de las tablas
'    DesBloqueoManual ("scaalp")
    
    PonerFoco Text1(0)
End Sub


Private Sub CargarFacturas(Socio As String)
'Recupera de la BD y muestra en el Listview todos los albaranes de compra
'que tiene el proveedor introducido.
Dim SQL As String
Dim RS As ADODB.Recordset
Dim RSFact As ADODB.Recordset

On Error GoTo ECargar
    
    Set vSeccion = New CSeccion
    If vSeccion.LeerDatos(vParamAplic.Seccionhorto) Then
        If vSeccion.AbrirConta Then
    
            SQL = "select codtipom, numfactu, fecfactu from rrecibpozos where codsocio = " & DBSet(Socio, "N")
            SQL = SQL & " and contabilizado = 1 "
            SQL = SQL & " order by fecfactu "
            
            Set RSFact = New ADODB.Recordset
            RSFact.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
            While Not RSFact.EOF
                SQL = "SELECT sum(coalesce(impvenci,0) + coalesce(gastos,0) - coalesce(impcobro,0)) importe "
                SQL = SQL & " FROM scobro INNER JOIN usuarios.stipom ON scobro.numserie = stipom.letraser "
                SQL = SQL & " WHERE stipom.codtipom = " & DBSet(RSFact!CodTipom, "T")
                SQL = SQL & " and scobro.codfaccl = " & DBSet(RSFact!numfactu, "N")
                SQL = SQL & " and scobro.fecfaccl = " & DBSet(RSFact!fecfactu, "F")
                Set RS = New ADODB.Recordset
                RS.Open SQL, ConnConta, adOpenForwardOnly, adLockPessimistic, adCmdText
                If Not RS.EOF Then
                    If DBLet(RS.Fields(0).Value) <> 0 Then
                    
                        Set It = ListView1.ListItems.Add
                        
                        'It.Tag = DevNombreSQL(RS!codCampo)
                        It.Text = DBLet(RSFact!CodTipom, "T")
                        It.SubItems(1) = Format(DBLet(RSFact!numfactu, "N"), "0000000")
                        It.SubItems(2) = RSFact!fecfactu
                        It.SubItems(3) = Format(DBLet(RS!Importe, "N"), "###,###,##0.00")
                        
                        It.Checked = False
                        
                        RS.MoveNext
                        TotalArray = TotalArray + 1
                        If TotalArray > 300 Then
                            TotalArray = 0
                            DoEvents
                        End If
                    
                    End If
                End If
                Set RS = Nothing
                
                RSFact.MoveNext
            Wend
    
            Set RSFact = Nothing
        End If
    End If
    Set vSeccion = Nothing

    If AdoAux(0).Recordset.RecordCount = 0 Then
        MsgBox "No existen albaranes pendientes de facturar para este socio.", vbExclamation
        BotonPedirDatos
    End If


ECargar:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando Albaranes", Err.Description
End Sub



Private Sub CalcularDatosFactura()
Dim I As Integer
Dim SQL As String
Dim cadAux As String
Dim ImpBruto As Currency
Dim ImpIVA As Currency
Dim vFactu As CFacturaTer
Dim RS As ADODB.Recordset
Dim Dto As Currency

    Dto = 0
    If Text1(8).Text <> "" Then
        Dto = CCur(TransformaPuntosComas(ImporteSinFormato(Text1(8).Text)))
    End If
    'Limpiar en el form los datos calculados de la factura
    'y volvemos a recalcular
    For I = 6 To 25
         Text1(I).Text = ""
    Next I

    cadAux = ""
    cadWHERE = ""
    ImpBruto = 0
    
    SQL = "select variedades.codigiva, sum(impentrada) from rhisfruta, variedades where codsocio= " & DBSet(Text1(3).Text, "N")
    If Text1(26).Text <> "" Then
        SQL = SQL & " and rhisfruta.codvarie = " & DBSet(Text1(26).Text, "N")
    End If
    SQL = SQL & " and variedades.codvarie = rhisfruta.codvarie "
    If Check1(2).Value = 0 Then
        SQL = SQL & " and rhisfruta.cobradosn = 0 "
    End If
    SQL = SQL & " group by 1 "
    SQL = SQL & " order by 1 "
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText

    If Not RS.EOF Then ImpBruto = ImpBruto + DBLet(RS.Fields(0).Value, "N")
    
    cadWHERE = "rhisfruta.codsocio=" & Val(Text1(3).Text)
    If Check1(2).Value = 0 Then
        cadWHERE = cadWHERE & " and rhisfruta.cobradosn = 0 "
    End If
    cadWHERE = cadWHERE & " and rhisfruta.impentrada <> 0 "

    If Text1(26).Text <> "" Then
        cadWHERE = cadWHERE & " and rhisfruta.codvarie = " & DBSet(Text1(26).Text, "N")
    End If

    If Not SeleccionaRegistros Then Exit Sub
    
    If Not BloqueaRegistro("rhisfruta", cadWHERE) Then
        conn.Execute "update rhisfruta set impentrada = 0 where " & cadWHERE
        cadWHERE = "rhisfruta.codsocio=" & Val(Text1(3).Text)
        
        If Check1(2).Value = 0 Then
            cadWHERE = cadWHERE & "  and rhisfruta.cobradosn = 0 "
        End If
        If Text1(26).Text <> "" Then
            cadWHERE = cadWHERE & " and rhisfruta.codvarie = " & DBSet(Text1(26).Text, "N")
        End If
        CargarAlbaranes cadWHERE
    End If
    
    cadWHERE = "rhisfruta.codsocio=" & Val(Text1(3).Text)
    cadWHERE = cadWHERE & " and rhisfruta.impentrada <> 0 "
    
    If Check1(2).Value = 0 Then
        cadWHERE = cadWHERE & " and rhisfruta.cobradosn = 0 "
    End If
        
    If Text1(26).Text <> "" Then
        cadWHERE = cadWHERE & " and rhisfruta.codvarie = " & DBSet(Text1(26).Text, "N")
    End If

    Set vFactu = New CFacturaTer
    vFactu.DtoPPago = 0
    vFactu.DtoGnral = 0
    If Dto <> 0 Then
        vFactu.DtoGnral = Dto
    End If
    vFactu.Intracomunitario = Check1(1).Value
    If vFactu.CalcularDatosFactura(cadWHERE, Text1(3).Text) Then
        Text1(6).Text = vFactu.BrutoFac
        Text1(7).Text = vFactu.ImpPPago
        Text1(8).Text = vFactu.ImpGnral
        Text1(9).Text = vFactu.BaseImp
        Text1(10).Text = vFactu.TipoIVA1
        Text1(11).Text = vFactu.TipoIVA2
        Text1(12).Text = vFactu.TipoIVA3
        Text1(13).Text = vFactu.PorceIVA1
        Text1(14).Text = vFactu.PorceIVA2
        Text1(15).Text = vFactu.PorceIVA3
        Text1(16).Text = vFactu.BaseIVA1
        Text1(17).Text = vFactu.BaseIVA2
        Text1(18).Text = vFactu.BaseIVA3
        Text1(19).Text = vFactu.ImpIVA1
        Text1(20).Text = vFactu.ImpIVA2
        Text1(21).Text = vFactu.ImpIVA3
        Text1(22).Text = vFactu.TotalFac
        Text1(23).Text = vFactu.BaseReten
        Text1(25).Text = vFactu.ImpReten
        If vFactu.ImpReten = 0 Then
            Text1(24).Text = 0
        Else
            Text1(24).Text = vFactu.PorcReten
        End If
        
        Check1(1).Value = vFactu.Intracomunitario
        
        For I = 6 To 26
            FormateaCampo Text1(I)
        Next I
        'Quitar ceros de linea IVA 2
        If Val(Text1(14).Text) = 0 And Val(Text1(11).Text) = 0 Then
            For I = 11 To 20 Step 3
                Text1(I).Text = QuitarCero(CCur(Text1(I).Text))
            Next I
        End If
        'Quitar ceros de linea IVA 3
        If Val(Text1(15).Text) = 0 And Val(Text1(12).Text) = 0 Then
            For I = 12 To 21 Step 3
                Text1(I).Text = QuitarCero(CCur(Text1(I).Text))
            Next I
        End If
        
    Else
        MuestraError Err.Number, "Calculando Factura", Err.Description
    End If
    Set vFactu = Nothing
    
    
    
End Sub

Private Function SeleccionaRegistros() As Boolean
'Comprueba que se seleccionan albaranes en la base de datos
'es decir que hay albaranes marcados
'cuando se van marcando albaranes se van añadiendo el la cadena cadWhere
Dim SQL As String

    On Error GoTo ESel
    SeleccionaRegistros = False
    
    If cadWHERE = "" Then Exit Function
    
    SQL = "Select count(*) FROM rhisfruta"
    SQL = SQL & " WHERE " & cadWHERE
    If RegistrosAListar(SQL) <> 0 Then SeleccionaRegistros = True
    Exit Function
    
ESel:
    SeleccionaRegistros = False
    MuestraError Err.Number, "No hay seleccionados Albaranes", Err.Description
End Function


Private Sub BotonFacturar()
Dim vFactu As CFacturaTer
Dim cad As String

    Screen.MousePointer = vbHourglass
    On Error GoTo Error1
    
    
    cad = ""
    If Text1(3).Text = "" Then
        cad = "Falta socio"
    Else
        If Not IsNumeric(Text1(3).Text) Then cad = "Campo socio debe ser numérico"
    End If
    If cad <> "" Then
        MsgBox cad, vbExclamation
        Exit Sub
    End If
        
    Set vSocio = New cSocio
    
    'Tiene que ller los datos del transportista
    If Not vSocio.LeerDatos(Text1(3).Text) Then Exit Sub
    
    If Not DatosOk Then
        Set vSocio = Nothing
        Exit Sub
    End If

    'Pasar los Albaranes seleccionados con cadWHERE a una factura
    Set vFactu = New CFacturaTer
    vFactu.Tercero = Text1(3).Text
    vFactu.numfactu = Text1(0).Text
    vFactu.fecfactu = Text1(1).Text
    vFactu.FecRecep = Text1(2).Text
    vFactu.trabajador = Text1(4).Text
    vFactu.BancoPr = Text1(5).Text
    vFactu.BrutoFac = ImporteFormateado(Text1(6).Text)
    vFactu.ForPago = Text1(4).Text
    vFactu.DtoPPago = 0
    vFactu.DtoGnral = 0
    vFactu.ImpPPago = ImporteFormateado(Text1(7).Text)
    vFactu.ImpGnral = ImporteFormateado(Text1(8).Text)
    vFactu.BaseIVA1 = ImporteFormateado(Text1(16).Text)
    vFactu.BaseIVA2 = ImporteFormateado(Text1(17).Text)
    vFactu.BaseIVA3 = ImporteFormateado(Text1(18).Text)
    vFactu.TipoIVA1 = ComprobarCero(Text1(10).Text)
    vFactu.TipoIVA2 = ComprobarCero(Text1(11).Text)
    vFactu.TipoIVA3 = ComprobarCero(Text1(12).Text)
    vFactu.PorceIVA1 = ComprobarCero(Text1(13).Text)
    vFactu.PorceIVA2 = ComprobarCero(Text1(14).Text)
    vFactu.PorceIVA3 = ComprobarCero(Text1(15).Text)
    vFactu.ImpIVA1 = ImporteFormateado(Text1(19).Text)
    vFactu.ImpIVA2 = ImporteFormateado(Text1(20).Text)
    vFactu.ImpIVA3 = ImporteFormateado(Text1(21).Text)
    vFactu.TotalFac = ImporteFormateado(Text1(22).Text)
    vFactu.PorcReten = ImporteFormateado(Text1(24).Text)
    vFactu.ImpReten = ImporteFormateado(Text1(25).Text)
    vFactu.BaseReten = ImporteFormateado(Text1(23).Text)
    
    'Si el proveedor tiene CTA BANCARIA se la asigno
    vFactu.CCC_Entidad = vSocio.Banco
    vFactu.CCC_Oficina = vSocio.Sucursal
    vFactu.CCC_CC = vSocio.Digcontrol
    vFactu.CCC_CTa = vSocio.CuentaBan
    
    vFactu.Intracomunitario = Check1(1).Value
    
    ' sacamos la cuenta de proveedor
    If Not vSocio.LeerDatosSeccion(vSocio.Codigo, vParamAplic.Seccionhorto) Then
        MsgBox "No se han encontrado los datos del socio de la sección Hortofrutícola", vbExclamation
        Set vFactu = Nothing
        Exit Sub
    End If
    
    vFactu.CtaTerce = vSocio.CtaProv
    
    If cadWHERE <> "" Then
        If vFactu.TraspasoAlbaranesAFactura(cadWHERE) Then BotonPedirDatos
    End If
    Set vFactu = Nothing
    
    
    Screen.MousePointer = vbDefault

Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub



Private Function ExisteFacturaEnHco() As Boolean
'Comprobamos si la factura ya existe en la tabla de Facturas a Proveedor: scafpc
Dim cad As String

    ExisteFacturaEnHco = False
    'Tiene que tener valor los 3 campos de clave primaria antes de comprobar
    If Not (Text1(0).Text <> "" And Text1(1).Text <> "" And Text1(3).Text <> "") Then Exit Function
    
    ' No debe existir el número de factura para el proveedor en hco. [06/05/2013]la fecha a mirar es la de recepcion
    cad = "SELECT count(*) FROM rcafter "
    cad = cad & " WHERE codsocio=" & Text1(3).Text & " AND numfactu=" & DBSet(Text1(0).Text, "T") & " AND year(fecfactu)=" & Year(Text1(2).Text)
    If RegistrosAListar(cad) > 0 Then
        MsgBox "Factura de Tercero ya existente. Reintroduzca.", vbExclamation
        ExisteFacturaEnHco = True
        Exit Function
    End If
End Function



'****************************************

'************* LLINIES: ****************************
Private Sub ToolAux_ButtonClick(Index As Integer, ByVal Button As MSComctlLib.Button)
Dim I As Integer
'-- pon el bloqueo aqui
    'If BLOQUEADesdeFormulario2(Me, Data1, 1) Then
    Select Case Button.Index
        Case 2
            BotonModificarLinea Index
        
    End Select
    'End If
End Sub




Private Sub BotonModificarLinea(Index As Integer)
    Dim anc As Single
    Dim I As Integer
    Dim J As Integer
    
    If AdoAux(Index).Recordset.EOF Then Exit Sub
    If AdoAux(Index).Recordset.RecordCount < 1 Then Exit Sub
    
    ModoLineas = 2 'Modificar llínia
       
    If Modo = 4 Then 'Modificar Capçalera
        cmdAceptar_Click
        If ModoLineas = 0 Then Exit Sub
    End If
       
    NumTabMto = Index
    PonerModo 5
    ' *** bloqueje la clau primaria de la capçalera ***
    BloquearTxt Text1(0), True
    ' *********************************
  
    Select Case Index
        Case 0 ' *** pose els index de llínies que tenen datagrid (en o sense tab) ***
            If DataGridAux(Index).Bookmark < DataGridAux(Index).FirstRow Or DataGridAux(Index).Bookmark > (DataGridAux(Index).FirstRow + DataGridAux(Index).VisibleRows - 1) Then
                I = DataGridAux(Index).Bookmark - DataGridAux(Index).FirstRow
                DataGridAux(Index).Scroll 0, I
                DataGridAux(Index).Refresh
            End If
              
            anc = DataGridAux(Index).Top
            If DataGridAux(Index).Row < 0 Then
                anc = anc + 210
            Else
                anc = anc + DataGridAux(Index).RowTop(DataGridAux(Index).Row) + 5
            End If
            
        
    End Select
    
    Select Case Index
        ' *** valor per defecte al modificar dels camps del grid ***
        Case 0 'importes
            txtAux(0).Text = DataGridAux(Index).Columns(0).Text
            txtAux(1).Text = DataGridAux(Index).Columns(1).Text
            txtAux(6).Text = DataGridAux(Index).Columns(2).Text
            Text2(2).Text = DataGridAux(Index).Columns(3).Text
            txtAux(3).Text = DataGridAux(Index).Columns(4).Text
            txtAux(2).Text = DataGridAux(Index).Columns(5).Text
            txtAux(4).Text = DataGridAux(Index).Columns(6).Text
            txtAux(5).Text = DataGridAux(Index).Columns(7).Text
            
            For I = 0 To 3
                BloquearTxt txtAux(I), True
            Next I
            BloquearTxt txtAux(4), False
            BloquearTxt txtAux(5), True
       
    End Select
    
    LLamaLineas Index, ModoLineas, anc
   
    ' *** foco al 1r camp visible de les llinies en grids que no siga PK (en o sense tab) ***
    Select Case Index
        Case 0 'importes
            PonerFoco txtAux(4)
    End Select
    ' ***************************************************************************************
    lblIndicador.Caption = "INSERTAR IMPORTE"
End Sub

Private Sub CargaGrid(Index As Integer, enlaza As Boolean)
Dim b As Boolean
Dim I As Byte
Dim tots As String

    On Error GoTo ECarga

    tots = MontaSQLCarga(Index, enlaza)

    CargaGridGnral Me.DataGridAux(Index), Me.AdoAux(Index), tots, PrimeraVez
    
    
    Select Case Index
        Case 0 'rhisfruta
            'si es visible|control|tipo campo|nombre campo|ancho control|
            tots = "S|txtaux(0)|T|Albarán|750|;S|txtaux(1)|T|Fecha|950|;"
            tots = tots & "S|txtaux(6)|T|Código|660|;S|Text2(2)|T|Variedad|1220|;"
            tots = tots & "S|txtaux(3)|T|Kilos Neto|1000|;S|txtaux(2)|T|Pr.Estim.|850|;"
            tots = tots & "S|txtaux(4)|T|Importe|1100|;N|txtaux(5)|T|Socio|1100|;"
            arregla tots, DataGridAux(Index), Me
        
            DataGridAux(0).Columns(1).Alignment = dbgLeft
            DataGridAux(0).Columns(5).Alignment = dbgRight
        
            b = (Modo = 4) And ((ModoLineas = 1) Or (ModoLineas = 2))
            BloquearTxt txtAux(3), Not b

    End Select
    DataGridAux(Index).ScrollBars = dbgAutomatic
    
    If Not AdoAux(0).Recordset.EOF Then
        Me.lblIndicador.Caption = AdoAux(0).Recordset.AbsolutePosition & " de " & AdoAux(0).Recordset.RecordCount
    Else
        Me.lblIndicador.Caption = ""
    End If
ECarga:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid: " & DataGridAux(Index).Tag, Err.Description
End Sub



Private Function MontaSQLCarga(Index As Integer, enlaza As Boolean) As String
'--------------------------------------------------------------------
' MontaSQlCarga:
'   Basant-se en la informació proporcionada pel vector de camps
'   crea un SQl per a executar una consulta sobre la base de datos que els
'   torne.
' Si ENLAZA -> Enlaça en el data1
'           -> Si no el carreguem sense enllaçar a cap camp
'--------------------------------------------------------------------
Dim SQL As String
Dim Tabla As String
   
    ' ********* si n'hi han tabs, dona igual si en datagrid o no ***********
    Select Case Index
        Case 0 'historico de entradas
            Tabla = "rhisfruta"
            SQL = "SELECT rhisfruta.numalbar,rhisfruta.fecalbar, rhisfruta.codvarie, variedades.nomvarie, rhisfruta.kilosnet, rhisfruta.prestimado, rhisfruta.impentrada, rhisfruta.codsocio "
            SQL = SQL & " FROM " & Tabla & " inner join variedades on rhisfruta.codvarie = variedades.codvarie "
            If enlaza Then
'                SQL = SQL & ObtenerWhereCab(True)
                SQL = SQL & " where codsocio =  " & DBSet(Text1(3).Text, "N")
                
                
                '[Monica] 04/02/2010 Todos los albaranes o solo los que no han sido cobrados
                If Check1(2).Value = 0 Then
                    SQL = SQL & " and cobradosn = 0 "   ' que no esten cobradas
                End If
                    
                If Text1(26).Text <> "" Then
                    SQL = SQL & " and rhisfruta.codvarie = " & DBSet(Text1(26).Text, "N")
                End If
            Else
                SQL = SQL & " WHERE numalbar  = -1"
            End If
            
            SQL = SQL & " ORDER BY " & Tabla & ".numalbar,  " & Tabla & ".fecalbar "
            
    End Select
    ' ********************************************************************************
    
    MontaSQLCarga = SQL
End Function


Private Sub cmdAceptar_Click()
    Screen.MousePointer = vbHourglass
    
    On Error GoTo Error1
    If Text1(0).Text = "" Then Exit Sub
    Select Case Modo
        Case 5 'LLÍNIES
            Select Case ModoLineas
                Case 2 'modificar llínies
                    ModoLineas = 0
                    PonerModo 5
                    ModificarLinea
                    If Not AdoAux(0).Recordset.EOF Then _
                        Me.lblIndicador.Caption = AdoAux(0).Recordset.AbsolutePosition & " de " & AdoAux(0).Recordset.RecordCount
            End Select
            
        CalcularDatosFactura
        ' **************************
    End Select
    Screen.MousePointer = vbDefault
 
Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub

Private Sub cmdCancelar_Click()
Dim I As Integer
Dim V
    
    Select Case Modo
        Case 5 'LLÍNIES
            Select Case ModoLineas
                Case 2 'modificar llínies
                    ModoLineas = 0
                    
                    If Not AdoAux(NumTabMto).Recordset.EOF Then
                        ' *** l'Index de Fields es el que canvie de la PK de llínies ***
                        V = AdoAux(NumTabMto).Recordset.Fields(0) 'el 2 es el nº de llinia
                        AdoAux(NumTabMto).Recordset.Find (AdoAux(NumTabMto).Recordset.Fields(0).Name & " =" & V)
                        ' ***************************************************************
                    End If
                    LLamaLineas NumTabMto, ModoLineas 'ocultar txtAux
                    PonerModo 2
                    If Not AdoAux(0).Recordset.EOF Then _
                         Me.lblIndicador.Caption = AdoAux(0).Recordset.AbsolutePosition & " de " & AdoAux(0).Recordset.RecordCount
                    End Select
    End Select
End Sub

Private Function ObtenerWhereCab(conW As Boolean) As String
Dim vWhere As String
    
    vWhere = ""
    If conW Then vWhere = " WHERE "
    ' *** canviar-ho per la clau primaria de la capçalera ***
    vWhere = vWhere & " numalbar=" & Val(txtAux(0).Text)
    ' *******************************************************
    
    ObtenerWhereCab = vWhere
End Function


Private Sub PosicionarData()
Dim cad As String, Indicador As String

    ' *** canviar-ho per tota la PK de la capçalera, no llevar els () ***
    cad = "numalbar = " & AdoAux(0).Recordset!Numalbar
    ' ***************************************
    
    ' *** gastar SituarData o SituarDataMULTI depenent de si la PK es simple o composta ***
    If SituarData(AdoAux(0), cad, Indicador) Then
        lblIndicador.Caption = Indicador
    End If
    ' ***********************************************************************************
End Sub


Private Sub LLamaLineas(Index As Integer, xModo As Byte, Optional alto As Single)
Dim jj As Integer
Dim b As Boolean

    ' *** si n'hi han tabs sense datagrid posar el If ***
    If Index <> 3 Then DeseleccionaGrid DataGridAux(Index)
    ' ***************************************************
       
    b = (xModo = 1 Or xModo = 2) 'Insertar o Modificar Llínies
    Select Case Index
        Case 0 'albaranes
            txtAux(0).visible = False
            txtAux(1).visible = False
            txtAux(2).visible = False
            txtAux(3).visible = False
            For jj = 4 To 4
                txtAux(jj).visible = b
                txtAux(jj).Top = alto
            Next jj
            
            Text2(2).visible = False
            
            
    End Select
End Sub


Private Sub ModificarLinea()
'Modifica registre en les taules de Llínies
Dim nomframe As String
Dim V As Long
Dim cad As String
    On Error Resume Next

    ' *** posa els noms del frames, tant si son de grid com si no ***
    Select Case NumTabMto
        Case 0: nomframe = "FrameAux0" 'cuentas Bancarias
    End Select
    ' **************************************************************

    If DatosOkLlin(nomframe) Then
        TerminaBloquear
        If ModificaDesdeFormulario2(Me, 2, nomframe) Then
'??monica
'            If BLOQUEADesdeFormulario2(Me, Data1, 1) Then BotonModificar
            If cadWHERE <> "" Then BloqueaRegistro "rhisfruta", cadWHERE

            If NumTabMto <> 3 Then
                V = AdoAux(NumTabMto).Recordset.Fields(0) 'el 2 es el nº de llinia
                CargaGrid NumTabMto, True
            End If

            ' *** si n'hi han tabs ***
'            SituarTab (NumTabMto)

            ' *** si n'hi han tabs que no tenen datagrid, posar el if ***
            If NumTabMto <> 3 Then
                DataGridAux(NumTabMto).SetFocus
                AdoAux(NumTabMto).Recordset.Find (AdoAux(NumTabMto).Recordset.Fields(0).Name & " =" & V)
            End If
            ' ***********************************************************

            LLamaLineas NumTabMto, 0
        End If
    End If
        
End Sub

Private Sub PonerCampos()
Dim I As Integer
Dim codpobla As String, despobla As String
Dim CPostal As String, desProvi As String, desPais As String

    If Data1.Recordset.EOF Then Exit Sub
    PonerCamposForma2 Me, Data1, 1 'opcio=1: posa el format o els camps de la capçalera
    
    CargaGrid I, True
    If Not AdoAux(I).Recordset.EOF Then _
        PonerCamposForma2 Me, AdoAux(I), 2, "FrameAux" & I
    
    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
    
    PonerModoOpcionesMenu (Modo)
End Sub


Private Function DatosOkLlin(nomframe As String) As Boolean
Dim RS As ADODB.Recordset
Dim SQL As String
Dim b As Boolean
Dim cant As Integer
Dim Mens As String
Dim vFact As Byte, vDocum As Byte
    
    On Error GoTo EDatosOKLlin

    Mens = ""
    DatosOkLlin = False

    b = CompForm2(Me, 2, nomframe) 'Comprovar formato datos ok
    If Not b Then Exit Function
    
    DatosOkLlin = b

EDatosOKLlin:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function


Private Sub PonerModoOpcionesMenu(Modo)
'Actives unes Opcions de Menú i Toolbar según el modo en que estem
Dim b As Boolean, bAux As Boolean
Dim I As Byte
    
    ' *** si n'hi han llínies que tenen grids (en o sense tab) ***
    b = (Modo = 3 Or Modo = 4 Or Modo = 2)
    For I = 0 To ToolAux.Count - 1
        ToolAux(I).Buttons(1).Enabled = b
        If b Then bAux = (b And Me.AdoAux(I).Recordset.RecordCount > 0)
        ToolAux(I).Buttons(2).Enabled = bAux
        ToolAux(I).Buttons(3).Enabled = bAux
    Next I
    ' ****************************************
    
'    ' *** si n'hi han tabs que no tenen grids ***
'    i = 3
'    If AdoAux(i).Recordset.EOF Then
'        ToolAux(i).Buttons(1).Enabled = b
'        ToolAux(i).Buttons(2).Enabled = False
'        ToolAux(i).Buttons(3).Enabled = False
'    Else
'        ToolAux(i).Buttons(1).Enabled = False
'        ToolAux(i).Buttons(2).Enabled = b
'    End If
    ' *******************************************
End Sub

Private Sub txtAux_GotFocus(Index As Integer)
    ConseguirFoco txtAux(Index), Modo
End Sub

Private Sub txtAux_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Index <> 0 And KeyCode <> 38 Then KEYdown KeyCode
End Sub

Private Sub txtAux_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtAux_LostFocus(Index As Integer)
Dim TipoDto As Byte
Dim ImpDto As String
Dim Unidades As String
Dim cantidad As String

    'Quitar espacios en blanco
    If Not PerderFocoGnralLineas(txtAux(Index), ModoLineas) Then Exit Sub
    
    Select Case Index
        Case 4 ' Importe
            If txtAux(Index).Text <> "" Then PonerFormatoDecimal txtAux(Index), 1
            
    End Select
End Sub

Private Function LimpiarImportes(vWhere As String) As Boolean
On Error GoTo eLimpiarImportes

    LimpiarImportes = False

    'primero limpiamos importes
    conn.Execute "update rhisfruta set impentrada = 0 where " & vWhere

    LimpiarImportes = True
    Exit Function

eLimpiarImportes:
    MuestraError Err.Number, "Limpiar Importes", Err.Description
End Function
                                

