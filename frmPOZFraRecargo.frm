VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmPOZFraRecargo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Facturación con recargo"
   ClientHeight    =   7500
   ClientLeft      =   45
   ClientTop       =   4035
   ClientWidth     =   7965
   Icon            =   "frmPOZFraRecargo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7500
   ScaleWidth      =   7965
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameBotonGnral 
      Height          =   705
      Left            =   135
      TabIndex        =   22
      Top             =   90
      Width           =   1470
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   210
         TabIndex        =   23
         Top             =   180
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   2
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Pedir Datos"
               Object.Tag             =   "2"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Generar Factura"
               Object.Tag             =   "2"
            EndProperty
         EndProperty
      End
   End
   Begin VB.CheckBox chkVistaPrevia 
      Caption         =   "Vista previa"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   6120
      TabIndex        =   21
      Top             =   270
      Visible         =   0   'False
      Width           =   1605
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   7
      Left            =   10380
      MaxLength       =   15
      TabIndex        =   14
      Tag             =   "Base Imponible 2 |N|N|0||scafac|baseimp2|#,###,###,##0.00|N|"
      Text            =   "Text1 7"
      Top             =   3375
      Width           =   1485
   End
   Begin VB.Frame FrameIntro 
      Height          =   1455
      Left            =   120
      TabIndex        =   4
      Top             =   780
      Width           =   7695
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   6030
         MaxLength       =   10
         TabIndex        =   3
         Tag             =   "Importe Factura|T|N|||tcafpc|numfactu||S|"
         Text            =   "Text1 7"
         Top             =   930
         Visible         =   0   'False
         Width           =   1380
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   3
         Left            =   2040
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   8
         Text            =   "Text2"
         Top             =   390
         Width           =   5400
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   3
         Left            =   1125
         MaxLength       =   6
         TabIndex        =   0
         Tag             =   "Cod. Socio|N|N|0|999999|tcafpc|codtrans|000|S|"
         Text            =   "Text1"
         Top             =   390
         Width           =   915
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   2010
         MaxLength       =   10
         TabIndex        =   1
         Tag             =   "Fecha Factura|F|N|||tcafpc|fecfactu|dd/mm/yyyy|S|"
         Top             =   930
         Width           =   1350
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   4440
         MaxLength       =   10
         TabIndex        =   2
         Tag             =   "Porcentaje Recargo|T|N|||tcafpc|numfactu||S|"
         Text            =   "Text1 7"
         Top             =   930
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Importe"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   5235
         TabIndex        =   16
         Top             =   960
         Visible         =   0   'False
         Width           =   1245
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   0
         Left            =   1740
         Picture         =   "frmPOZFraRecargo.frx":000C
         ToolTipText     =   "Buscar fecha"
         Top             =   930
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   840
         ToolTipText     =   "Buscar socio"
         Top             =   420
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Socio"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   210
         TabIndex        =   7
         Top             =   420
         Width           =   600
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha Factura"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   29
         Left            =   240
         TabIndex        =   6
         Top             =   930
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "%Recargo"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   28
         Left            =   3390
         TabIndex        =   5
         Top             =   960
         Width           =   1095
      End
   End
   Begin VB.Frame FrameAux0 
      Height          =   5040
      Left            =   120
      TabIndex        =   10
      Top             =   2190
      Width           =   7710
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   5580
         MaxLength       =   10
         TabIndex        =   20
         Tag             =   "Fecha Factura|F|N|||tcafpc|fecfactu|dd/mm/yyyy|S|"
         Top             =   4590
         Width           =   1830
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   5580
         MaxLength       =   10
         TabIndex        =   17
         Top             =   4170
         Width           =   1830
      End
      Begin VB.Frame Frame1 
         Height          =   555
         Index           =   0
         Left            =   180
         TabIndex        =   11
         Top             =   4260
         Width           =   2865
         Begin VB.Label lblIndicador 
            Alignment       =   2  'Center
            Caption         =   "Label2"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   120
            TabIndex        =   12
            Top             =   180
            Width           =   2655
         End
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   3435
         Left            =   150
         TabIndex        =   15
         Top             =   540
         Width           =   7305
         _ExtentX        =   12885
         _ExtentY        =   6059
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
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
      Begin VB.Label Label1 
         Caption         =   "TOTAL CON RECARGO"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   3195
         TabIndex        =   19
         Top             =   4590
         Width           =   2370
      End
      Begin VB.Label Label1 
         Caption         =   "TOTAL SELECCIONADO"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   3195
         TabIndex        =   18
         Top             =   4230
         Width           =   2370
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   0
         Left            =   6540
         Picture         =   "frmPOZFraRecargo.frx":0097
         Top             =   210
         Width           =   240
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   1
         Left            =   6900
         Picture         =   "frmPOZFraRecargo.frx":01E1
         Top             =   210
         Width           =   240
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
         TabIndex        =   13
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
   Begin VB.Label Label1 
      Caption         =   "Operador"
      Height          =   255
      Index           =   1
      Left            =   1845
      TabIndex        =   9
      Top             =   900
      Width           =   735
   End
   Begin VB.Image imgBuscar 
      Height          =   240
      Index           =   1
      Left            =   1845
      Picture         =   "frmPOZFraRecargo.frx":032B
      ToolTipText     =   "Buscar trabajador"
      Top             =   1125
      Width           =   240
   End
   Begin VB.Menu mnOpciones 
      Caption         =   "&Opciones"
      Enabled         =   0   'False
      Visible         =   0   'False
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
Dim Indice As Integer

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

Dim vSeccion As CSeccion

'GENERALES PARA PASARLE A CRYSTAL REPORT
Private cadFormula As String 'Cadena con la FormulaSelection para Crystal Report
Private CadParam As String 'Cadena con los parametros para Crystal Report
Private numParam As Byte 'Numero de parametros que se pasan a Crystal Report
Private cadSelect As String 'Cadena para comprobar si hay datos antes de abrir Informe
Private cadTitulo As String 'Titulo para la ventana frmImprimir
Private cadNombreRPT As String 'Nombre del informe
Private Titulo As String 'Titulo informe que se pasa a frmImprimir
Private nomRPT As String 'nombre del fichero .rpt a imprimir
Private conSubRPT As Boolean 'si tiene subinformes para enlazarlos a las tablas correctas
'-------------------------------------


Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
'    If VerAlbaranes Then RefrescarAlbaranes
'    VerAlbaranes = False

    If PrimeraVez Then
        mnPedirDatos_Click
        PrimeraVez = False
    End If
    

End Sub

Private Sub Form_Load()
Dim I As Integer
    'Icono del formulario
    Me.Icon = frmPpal.Icon

    PrimeraVez = True
    
    ' ICONITOS DE LA BARRA
'    With Me.Toolbar1
'        .ImageList = frmPpal.imgListComun
'        .Buttons(1).Image = 1   'Pedir Datos
'        .Buttons(2).Image = 15   'Generar FActura
'        .Buttons(5).Image = 11   'Salir
'    End With
    
    With Me.Toolbar1
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 1   'Pedir Datos
        .Buttons(2).Image = 15  'Generar FActura
    End With
    
    ' ******* si n'hi han llínies *******
    
    ' ***********************************
    
    
    'cargar IMAGES de busqueda
    Me.imgBuscar(0).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    
    
    LimpiarCampos   'Limpia los campos TextBox
    InicializarListView
   
    '## A mano
    NombreTabla = "rrecibpozos" ' facturas de pozos
    
    'Vemos como esta guardado el valor del check
'    chkVistaPrevia.Value = CheckValueLeer(Name)
    
    'ASignamos un SQL al DATA1
    CadenaConsulta = "Select * from " & NombreTabla
    CadenaConsulta = CadenaConsulta & " where numfactu is null"
    
    Data1.ConnectionString = conn
    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
        
    'Cargar el DataGrid de lineas de Revisiones inicialmente a nada DATA2
    PrimeraVez = True
    
    If DatosADevolverBusqueda = "" Then
        PonerModo 0
    End If
    
    
End Sub


Private Sub LimpiarCampos()
On Error Resume Next
    limpiar Me
    'Aqui va el especifico de cada form es
    '### a mano
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub Form_Unload(Cancel As Integer)
'    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
'    DesBloqueoManual "FACTRA"
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
Dim Indice As Byte
    Indice = CByte(Me.imgFecha(0).Tag)
    Text1(Indice).Text = Format(vFecha, "dd/mm/yyyy")
End Sub


Private Sub frmFPa_DatoSeleccionado(CadenaSeleccion As String)
    Text1(4).Text = RecuperaValor(CadenaSeleccion, 1) 'Cod forpa
    FormateaCampo Text1(4)
    Text2(4).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom forpa
End Sub

Private Sub frmSoc_DatoSeleccionado(CadenaSeleccion As String)
Dim Indice As Byte
    Indice = 3
    Text1(Indice).Text = RecuperaValor(CadenaSeleccion, 1) 'Cod Socios
    FormateaCampo Text1(Indice)
    Text2(Indice).Text = RecuperaValor(CadenaSeleccion, 2) 'Nom socio
End Sub

Private Sub imgBuscar_Click(Index As Integer)
Dim Indice As Byte

    If Modo = 2 Or Modo = 0 Then Exit Sub
'    Screen.MousePointer = vbHourglass

    Select Case Index
        Case 0 'Cod. Proveedor
            Set frmSoc = New frmManSocios
            frmSoc.DatosADevolverBusqueda = "0|1|"
            frmSoc.Show vbModal
            Set frmSoc = Nothing
            Indice = 3
       
       Case 2 'Bancos Propios
            Indice = 5
            Set frmBanPr = New frmComBanco
            frmBanPr.DatosADevolverBusqueda = "0|1|"
            frmBanPr.Show vbModal
            Set frmBanPr = Nothing
            
       
    End Select
    
    PonerFoco Text1(Indice)
'    Screen.MousePointer = vbDefault
End Sub

Private Sub imgFecha_Click(Index As Integer)
Dim Indice As Byte
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
   Indice = Index + 1
   Me.imgFecha(0).Tag = Indice
   
   PonerFormatoFecha Text1(Indice)
   If Text1(Indice).Text <> "" Then frmF.NovaData = CDate(Text1(Indice).Text)

   Screen.MousePointer = vbDefault
   frmF.Show vbModal
   Set frmF = Nothing
   PonerFoco Text1(Indice)
End Sub


Private Sub ListView1_ItemCheck(ByVal item As MSComctlLib.ListItem)
    If item.SubItems(4) < 0 Then item.Checked = False
    
    CalcularTotales
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
        
        Case 3 'Cod Socios
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), "rsocios", "nomsocio", "codsocio")
                
                If Text2(Index).Text <> "" Then
                    CargarFacturas Text1(Index)
                    CalcularTotales
                End If
            Else
                Text2(Index).Text = ""
            End If
            PonerFoco Text1(1)
            
        Case 0 ' porcentaje
            PonerFormatoDecimal Text1(Index), 4
          
        Case 2 ' importe recargo
            PonerFormatoDecimal Text1(Index), 3
            PonerFocoListView Me.ListView1
    End Select
    CalcularTotales
End Sub


'----------------------------------------------------------------
'----------------------------------------------------------------
'   En PONERMODO se habilitan, o no, los diverso campos del
'   formulario en funcion del modo en k vayamos a trabajar
Private Sub PonerModo(Kmodo As Byte)
Dim I As Byte, NumReg As Byte
Dim B As Boolean
On Error GoTo EPonerModo

    Modo = Kmodo
    
    'Modo 2. Hay datos y estamos visualizandolos
    '=========================================
    B = (Modo = 2)
        
'    cmdAceptar.visible = (ModoLineas = 2)
'    cmdAceptar.Enabled = (ModoLineas = 2)
'    cmdCancelar.visible = (ModoLineas = 2)
'    cmdCancelar.Enabled = (ModoLineas = 2)
    
'    'Bloquea los campos Text1 sino estamos modificando/Insertando Datos
'    'Si estamos en Insertar además limpia los campos Text1
'    'si estamos en modificar bloquea las compos que son clave primaria
'    BloquearText1 Me, Modo
    
    For I = 0 To Text1.Count - 1
        BloquearTxt Text1(I), (Modo <> 3)
    Next I
    
    
    '---------------------------------------------
    B = (Modo <> 0 And Modo <> 2 And Modo <> 5)
    
    For I = 0 To Me.imgBuscar.Count - 1
        Me.imgBuscar(I).Enabled = B
    Next I
                    
    Me.chkVistaPrevia.Enabled = (Modo <= 2)
    
        
    Me.FrameIntro.Enabled = (Modo = 3)
'    Me.FrameAux0.Enabled = (Modo = 5)
       
 
    If Modo = 3 Then
        lblIndicador.Caption = "Datos factura"
    Else
        lblIndicador.Caption = "Generar factura"
    End If
 
       
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


Private Function DatosOK() As Boolean
'Comprobar que los datos del frame de introduccion son correctos antes de cargar datos
Dim vtag As CTag
Dim cad As String
Dim I As Byte
Dim vSeccion As CSeccion

    On Error GoTo EDatosOK
    DatosOK = False
    
    ' deben de introducirse todos los datos del frame
    For I = 0 To 3
        If Text1(I).Text = "" Then
            If Text1(I).Tag <> "" Then
                Set vtag = New CTag
                If vtag.Cargar(Text1(I)) Then
                    cad = vtag.Nombre
                Else
                    cad = "Campo"
                End If
                Set vtag = Nothing
            End If
            If (I = 0 And ComprobarCero(Text1(2).Text) = "0") Or (I = 2 And ComprobarCero(Text1(0).Text) = "0") Then
                MsgBox cad & " no puede estar vacio. Reintroduzca", vbExclamation
                PonerModo 3
                PonerFoco Text1(I)
                Exit Function
            End If
        End If
    Next I
        
    'comprobamos que no me hayan puesto ambos valores
    If ComprobarCero(Text1(0).Text) <> "0" And ComprobarCero(Text1(2).Text) <> "0" Then
        MsgBox "Es incorrecto introducir ambos valores, porcentaje e importe. Revise.", vbExclamation
        PonerFoco Text1(0)
        Exit Function
    End If
        
'++monica:03/12/2008
    'comprobamos que hay lineas para facturar: facturas
    If Not HayFacturas Then
        MsgBox "No hay facturas para realizar una factura con recargo. Revise.", vbExclamation
        Exit Function
    End If
    
    
    
    'comprobar que la fecha de la factura sea anterior a la fecha de la rectificativa y nueva factura con recargo
    If Not DeFechaIgualoPosterior() Then
        MsgBox "Hay facturas seleccionadas de fecha superior a las que vamos a generar. Revise.", vbExclamation
        Exit Function
    End If
    
    'Comprobar que la fecha de nueva factura esta dentro de los ejercicios contables
    Set vSeccion = New CSeccion
    If vSeccion.LeerDatos(vParamAplic.Seccionhorto) Then
        If vSeccion.AbrirConta Then
            I = EsFechaOKConta(CDate(Text1(1).Text))
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

    
    
    
    
    
'    'Si la forpa es TRANSFERENCIA entonces compruebo la si tiene cta bancaria
''    cad = "select distinct (codforpa) from scaalp "
''    cad = cad & " WHERE " & Replace(cadWHERE, "slialp", "scaalp")
'    Set miRsAux = New ADODB.Recordset
''    miRsAux.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
''    cad = miRsAux.Fields(0)
''    miRsAux.Close
'
'
'
'    'Ahora buscamos el tipforpa del codforpa
'    Cad = "Select tipoforp from forpago where codforpa=" & DBSet(Text1(4).Text, "N")
'    miRsAux.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'    I = 0
'    If miRsAux.EOF Then
'        MsgBox "Error en el TIPO de forma de pago", vbExclamation
'    Else
'        I = 1
'        Cad = miRsAux.Fields(0)
'        If Val(Cad) = vbFPTransferencia Then
'            'Compruebo que la forpa es transferencia
'            I = 2
'        End If
'    End If
'    miRsAux.Close
'    Set miRsAux = Nothing
'
'
'    If I = 2 Then
'        'La forma de pago es transferencia. Debo comprobar que existe la cuenta bancaria
'        'del proveedor
'        If vSocio.CuentaBan = "" Or vSocio.Digcontrol = "" Or vSocio.Sucursal = "" Or vSocio.Banco = "" Then
'            Cad = "Cuenta bancaria incorrecta. Forma de pago: transferencia.    ¿Continuar?"
'            If MsgBox(Cad, vbQuestion + vbYesNoCancel) <> vbYes Then I = 0
'        End If
'    End If
'
'    'Si i=0 es que o esta mal la forpa o no quiere seguir pq no tiene cuenta bancaria
'    If I > 0 Then DatosOk = True


    DatosOK = True
    Exit Function
    
EDatosOK:
    DatosOK = False
    MuestraError Err.Number, "Comprobar datos correctos", Err.Description
End Function

Private Function HayFacturas() As Boolean
Dim Sql As String
Dim I As Integer

    HayFacturas = False
    For I = 1 To Me.ListView1.ListItems.Count
        If Me.ListView1.ListItems(I).Checked Then
            HayFacturas = True
            Exit For
        End If
    Next I


End Function

Private Function DeFechaIgualoPosterior() As Boolean
Dim Sql As String
Dim I As Integer

    DeFechaIgualoPosterior = True

    For I = 1 To Me.ListView1.ListItems.Count
        If Me.ListView1.ListItems(I).Checked Then
            If CDate(ListView1.ListItems(I).SubItems(3)) > CDate(Text1(1).Text) Then
                DeFechaIgualoPosterior = False
                Exit For
            End If
        End If
    Next I
    
End Function

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1  'Pedir datos
             mnPedirDatos_Click
            
        Case 2 'Generar Factura
            mnGenerarFac_Click

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
    'Vaciamos el ListView

    InicializarListView

    'Como no habrá albaranes seleccionados vaciamos la cadwhere
    cadWHERE = ""
    
    PonerModo 3
    
    'fecha recepcion
    Text1(1).Text = Format(Now, "dd/mm/yyyy")
    
    
    'desbloquear los registros de la saalp (si hay bloquedos)
    TerminaBloquear
    
    'si vamos
    'desBloqueo Manual de las tablas
'    DesBloqueoManual ("scaalp")
    
    PonerFoco Text1(3)
    
End Sub


Private Sub CargarFacturas(Socio As String)
'Recupera de la BD y muestra en el Listview todos los albaranes de compra
'que tiene el proveedor introducido.
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim RSFact As ADODB.Recordset
Dim It As ListItem
Dim TotalArray As Integer
Dim HayReg As Boolean
Dim Codmacta As String
On Error GoTo ECargar

    On Error GoTo ECargar

    
    Set vSeccion = New CSeccion
    If vSeccion.LeerDatos(vParamAplic.Seccionhorto) Then
        If vSeccion.AbrirConta Then
    
            Sql = "select uuu.codtipom, uuu.nomtipom, numfactu, fecfactu from rrecibpozos inner join usuarios.stipom uuu on rrecibpozos.codtipom = uuu.codtipom where codsocio = " & DBSet(Socio, "N")
            Sql = Sql & " and fecfactu "
            Sql = Sql & " and contabilizado = 1  "
            '[Monica]15/01/2016: no se muestran para poder rectificar ni las facturas rectifcativas ni
            '                   LAS DE RIEGO A MANTA --> SOBRARIA EL TIPO RRT(RECT.CONSUMO MANTA)
            Sql = Sql & " and not rrecibpozos.codtipom in ('RRC','RRM','RRT','RRV','RTA','RMT') "
            
            
            Sql = Sql & " order by 4 "
            
            Set RSFact = New ADODB.Recordset
            RSFact.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
            HayReg = False
    
            'DevuelveValor("select codmaccli from rsocios_seccion where codsocio = " & DBSet(Socio, "N") & " and codsecci = " & DBSet(vParamAplic.Seccionhorto, "N"))
    
            Codmacta = DevuelveDesdeBDNew(cAgro, "rsocios_seccion", "codmaccli", "codsocio", Socio, "N", , "codsecci", vParamAplic.Seccionhorto, "N")
                
                
            ListView1.ListItems.Clear
                
            While Not RSFact.EOF
                '[Monica]09/05/2018: fallaba sobre la nueva contabilidad
                If vParamAplic.ContabilidadNueva Then
                    Sql = "SELECT sum(coalesce(impvenci,0) + coalesce(gastos,0) - coalesce(impcobro,0)) importe "
                    Sql = Sql & " FROM cobros INNER JOIN usuarios.stipom ON cobros.numserie = stipom.letraser "
                    Sql = Sql & " WHERE stipom.codtipom = " & DBSet(RSFact!CodTipom, "T")
                    Sql = Sql & " and cobros.numfactu = " & DBSet(RSFact!numfactu, "N")
                    Sql = Sql & " and cobros.fecfactu = " & DBSet(RSFact!fecfactu, "F")
                    Sql = Sql & " and cobros.codmacta = " & DBSet(Codmacta, "T")
                
                Else
                    Sql = "SELECT sum(coalesce(impvenci,0) + coalesce(gastos,0) - coalesce(impcobro,0)) importe "
                    Sql = Sql & " FROM scobro INNER JOIN usuarios.stipom ON scobro.numserie = stipom.letraser "
                    Sql = Sql & " WHERE stipom.codtipom = " & DBSet(RSFact!CodTipom, "T")
                    Sql = Sql & " and scobro.codfaccl = " & DBSet(RSFact!numfactu, "N")
                    Sql = Sql & " and scobro.fecfaccl = " & DBSet(RSFact!fecfactu, "F")
                    Sql = Sql & " and scobro.codmacta = " & DBSet(Codmacta, "T")
                End If
                Set Rs = New ADODB.Recordset
                Rs.Open Sql, ConnConta, adOpenForwardOnly, adLockPessimistic, adCmdText
                
                If Not Rs.EOF Then
                    If DBLet(Rs.Fields(0).Value, "N") <> 0 Then
                    
                        HayReg = True
                    
                        Set It = ListView1.ListItems.Add
                        
                        'It.Tag = DevNombreSQL(RS!codCampo)
                        It.Text = DBLet(RSFact!CodTipom, "T")
                        It.SubItems(1) = RSFact!nomtipom
                        It.SubItems(2) = Format(DBLet(RSFact!numfactu, "N"), "0000000")
                        It.SubItems(3) = RSFact!fecfactu
                        'It.SubItems(4) = Format(DBLet(Rs!Importe, "N"), "###,###,##0.00")
                        
                        ' de momento saco el importe del totalfact
                        Sql = "SELECT sum(coalesce(totalfact,0))  from rrecibpozos "
                        Sql = Sql & " where codtipom = " & DBSet(RSFact!CodTipom, "T")
                        Sql = Sql & " and numfactu = " & DBSet(RSFact!numfactu, "N")
                        Sql = Sql & " and fecfactu = " & DBSet(RSFact!fecfactu, "F")
                        
                        Dim Importe As Currency
                        Importe = DevuelveValor(Sql)
                        It.SubItems(4) = Format(Importe, "###,###,##0.00")
                        
                        If Importe > 0 Then
                            It.Checked = True
                        End If
                        
                        Rs.MoveNext
                        TotalArray = TotalArray + 1
                        If TotalArray > 300 Then
                            TotalArray = 0
                            DoEvents
                        End If
                    
                    End If
                End If
                Set Rs = Nothing
                
                RSFact.MoveNext
            Wend
    
            Set RSFact = Nothing
        End If
    End If
    Set vSeccion = Nothing

    If Not HayReg Then
        MsgBox "No hay facturas pendientes de cobro para este socio.", vbExclamation
        BotonPedirDatos
    End If


ECargar:
    Set vSeccion = Nothing
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando Facturas", Err.Description
End Sub




Private Function SeleccionaRegistros() As Boolean
'Comprueba que se seleccionan albaranes en la base de datos
'es decir que hay albaranes marcados
'cuando se van marcando albaranes se van añadiendo el la cadena cadWhere
Dim Sql As String

    On Error GoTo ESel
    SeleccionaRegistros = False
    
    If cadWHERE = "" Then Exit Function
    
    Sql = "Select count(*) FROM rhisfruta"
    Sql = Sql & " WHERE " & cadWHERE
    If RegistrosAListar(Sql) <> 0 Then SeleccionaRegistros = True
    Exit Function
    
ESel:
    SeleccionaRegistros = False
    MuestraError Err.Number, "No hay seleccionados Albaranes", Err.Description
End Function


Private Sub BotonFacturar()
Dim vFactu As CFacturaTer
Dim cad As String
Dim I As Integer
Dim B As Boolean
Dim Sql As String

    Screen.MousePointer = vbHourglass
    On Error GoTo Error1
    
    If Not DatosOK Then Exit Sub
    
    
    Set vSocio = New cSocio
    
    'Tiene que llevar los datos del socio
    If Not vSocio.LeerDatos(Text1(3).Text) Then Exit Sub
    
    If Not DatosOK Then
        Set vSocio = Nothing
        Exit Sub
    End If

    conn.BeginTrans
    
    ' desmarcamos todas las facturas de ese socio
    Sql = "update rrecibpozos set imprimir = '' where codsocio = " & DBSet(Text1(3).Text, "N")
    conn.Execute Sql
    
    
    B = True
    I = 1
    While I <= Me.ListView1.ListItems.Count And B
        If Me.ListView1.ListItems(I).Checked Then
            'creamos la factura rectificativa
            B = CrearFactura(Me.ListView1.ListItems(I).Text, Me.ListView1.ListItems(I).SubItems(2), Me.ListView1.ListItems(I).SubItems(3), True)

            'creamos la nueva factura con recargo
            If B Then B = CrearFactura(Me.ListView1.ListItems(I), Me.ListView1.ListItems(I).SubItems(2), Me.ListView1.ListItems(I).SubItems(3), False)
        End If
        I = I + 1
    Wend

    If B Then
        conn.CommitTrans
    Else
        conn.RollbackTrans
    End If
        
    MsgBox "Proceso realizado correctamente.", vbExclamation
    Screen.MousePointer = vbDefault
    
    'impresion de facturas
    BotonImprimir
    
    Unload Me
    
    Exit Sub
    
Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Function CrearFactura(vTipoM As String, NumFact As String, FecFact As String, EsRectificativa As Boolean) As Boolean
Dim Sql As String
Dim vTipoMov As CTiposMov
Dim vSeccion As CSeccion
Dim CodTipom As String
Dim numfactu As Long
Dim devuelve As String
Dim Existe As Boolean
Dim Mens As String
Dim vImpRecargo As Currency
Dim SqlInsert As String
Dim SqlValues As String

Dim vGastoDevol As Currency
Dim RsGastos As ADODB.Recordset
Dim SqlGastos As String

Dim B As Boolean

    On Error GoTo eCrearFactura

    CrearFactura = False
    
    If EsRectificativa Then
        Select Case vTipoM
            Case "RCP"
                CodTipom = "RRC"
            Case "RMP"
                CodTipom = "RRM"
            Case "RMT"
                CodTipom = "RRT"
            Case "RVP"
                CodTipom = "RRV"
            Case "TAL"
                CodTipom = "RTA"
        End Select
    Else
        CodTipom = vTipoM
    End If
    
    Set vTipoMov = New CTiposMov
    If vTipoMov.Leer(CodTipom) Then
        'Comprobar si mientras tanto se incremento el contador de albaranes
        Do
            numfactu = vTipoMov.ConseguirContador(CodTipom)
            '[Monica]26/05/2016: añadida la condicion del año que no se miraba para la comprobacion de la existencia del contador
            devuelve = DevuelveDesdeBDNew(cAgro, "rrecibpozos", "numfactu", "numfactu", CStr(numfactu), "N", , "codtipom", CodTipom, "T", "year(fecfactu)", Mid(Text1(1).Text, 7, 4), "N")
            If devuelve <> "" Then
                'Ya existe el contador incrementarlo
                Existe = True
                vTipoMov.IncrementarContador (CodTipom)
                numfactu = vTipoMov.ConseguirContador(CodTipom)
            Else
                Existe = False
            End If
        Loop Until Not Existe
            
    Else 'No existe el tipo de Movimiento
        Set vTipoMov = Nothing
        Exit Function
    End If
    
    Mens = "Insertando rrecibpozos"
    
    ' cabecera de factura
    SqlInsert = "insert into rrecibpozos (codtipom,numfactu,fecfactu,numlinea,codsocio,hidrante,baseimpo,tipoiva,porc_iva,imporiva,totalfact,"
    SqlInsert = SqlInsert & "consumo,impcuota,lect_ant,fech_ant,lect_act,fech_act,consumo1,precio1,consumo2,precio2,concepto,"
    SqlInsert = SqlInsert & "contabilizado,impreso,conceptomo,importemo,conceptoar1,importear1,conceptoar2,importear2,"
    SqlInsert = SqlInsert & "conceptoar3,importear3,conceptoar4,importear4,difdias,codparti,calibre,codpozo,porcdto,impdto,"
    SqlInsert = SqlInsert & "precio,pasaridoc,parcelas,poligono,nroorden,numalbar,fecalbar,escontado,"
    SqlInsert = SqlInsert & "lect_ant2 , fech_ant2, lect_act2, fech_act2, imprimir "
    
    If EsRectificativa Then
        vImpRecargo = 0
        SqlInsert = SqlInsert & ",codtipomrec, numfacturec, fecfacturec, imprecargo, gastodedevol, porcrecargo) "
    Else
        vImpRecargo = CalculoImpRecargo(vTipoM, NumFact, FecFact)
        SqlInsert = SqlInsert & ",imprecargo, porcrecargo) "
    End If

    SqlValues = " select " & DBSet(CodTipom, "T") & "," & DBSet(numfactu, "N") & "," & DBSet(Text1(1).Text, "F") & ", numlinea, codsocio,hidrante, "
    If EsRectificativa Then
        SqlValues = SqlValues & "baseimpo * (-1),tipoiva,porc_iva,imporiva*(-1), totalfact * (-1),"
        SqlValues = SqlValues & "consumo,impcuota * (-1),lect_ant,fech_ant,lect_act,fech_act,consumo1,precio1 * (-1),consumo2,precio2 * (-1),concat(concepto,' Nro." & NumFact & " de " & FecFact & "'),"
        SqlValues = SqlValues & "0,0,conceptomo,importemo *(-1),conceptoar1,importear1 * (-1),conceptoar2,importear2 * (-1),"
        SqlValues = SqlValues & "conceptoar3,importear3 * (-1),conceptoar4,importear4 * (-1),difdias,codparti,calibre,codpozo,porcdto,impdto * (-1),"
        SqlValues = SqlValues & "precio * (-1),pasaridoc,parcelas,poligono,nroorden,numalbar,fecalbar,escontado,"
        SqlValues = SqlValues & "lect_ant2 , fech_ant2, lect_act2, fech_act2, " & DBSet(vUsu.PC, "T")
    Else
        SqlValues = SqlValues & "baseimpo + " & DBSet(vImpRecargo, "N") & ",tipoiva,porc_iva,round((baseimpo + " & DBSet(vImpRecargo, "N") & ") * porc_iva / 100,2), totalfact + " & DBSet(vImpRecargo, "N") & ","
        SqlValues = SqlValues & "consumo,impcuota,lect_ant,fech_ant,lect_act,fech_act,consumo1,precio1,consumo2,precio2,concat(concepto,' Nro." & NumFact & " de " & FecFact & "'),"
        SqlValues = SqlValues & "0,0,conceptomo,importemo,conceptoar1,importear1,conceptoar2,importear2,"
        SqlValues = SqlValues & "conceptoar3,importear3,conceptoar4,importear4,difdias,codparti,calibre,codpozo,porcdto,impdto,"
        SqlValues = SqlValues & "precio,pasaridoc,parcelas,poligono,nroorden,numalbar,fecalbar,escontado,"
        SqlValues = SqlValues & "lect_ant2 , fech_ant2, lect_act2, fech_act2, " & DBSet(vUsu.PC, "T")
    End If
    
    ' grabamos la factura a la que rectifica
    If EsRectificativa Then
        vGastoDevol = 0
'???? si tenemos que imprimir en la factura rectificativa los gastos de la original
        Set vSeccion = New CSeccion
        If vSeccion.LeerDatos(vParamAplic.Seccionhorto) Then
            If vSeccion.AbrirConta Then
                
                If vParamAplic.ContabilidadNueva Then
                    SqlGastos = "SELECT sum(coalesce(gastos,0)) gastos "
                    SqlGastos = SqlGastos & " FROM cobros INNER JOIN usuarios.stipom ON cobros.numserie = stipom.letraser "
                    SqlGastos = SqlGastos & " WHERE stipom.codtipom = " & DBSet(vTipoM, "T")
                    SqlGastos = SqlGastos & " and cobros.numfactu = " & DBSet(NumFact, "N")
                    SqlGastos = SqlGastos & " and cobros.fecfactu = " & DBSet(FecFact, "F")
                
                Else
                
                    SqlGastos = "SELECT sum(coalesce(gastos,0)) gastos "
                    SqlGastos = SqlGastos & " FROM scobro INNER JOIN usuarios.stipom ON scobro.numserie = stipom.letraser "
                    SqlGastos = SqlGastos & " WHERE stipom.codtipom = " & DBSet(vTipoM, "T")
                    SqlGastos = SqlGastos & " and scobro.codfaccl = " & DBSet(NumFact, "N")
                    SqlGastos = SqlGastos & " and scobro.fecfaccl = " & DBSet(FecFact, "F")
                End If
                Set RsGastos = New ADODB.Recordset
                RsGastos.Open SqlGastos, ConnConta, adOpenForwardOnly, adLockPessimistic, adCmdText
                If Not RsGastos.EOF Then
                    vGastoDevol = DBLet(RsGastos!Gastos, "N")
                End If
                Set RsGastos = Nothing
            End If
        End If
        Set vSeccion = Nothing
'????
        SqlValues = SqlValues & "," & DBSet(vTipoM, "T") & "," & DBSet(NumFact, "N") & "," & DBSet(FecFact, "F") & ",0, " & DBSet(vGastoDevol, "N") & ",0 "
    Else
        SqlValues = SqlValues & "," & DBSet(vImpRecargo, "N") & "," & DBSet(Text1(0).Text, "N")
    End If
    SqlValues = SqlValues & " from rrecibpozos where codtipom = " & DBSet(vTipoM, "T") & " and numfactu = " & DBSet(NumFact, "N") & " and fecfactu = " & DBSet(FecFact, "F")
    
    conn.Execute SqlInsert & SqlValues
    
    
    ' lineas de factura rrecibpozos_acc
    Mens = "Insertando rrecibpozos_acc"
    
    SqlInsert = "insert into rrecibpozos_acc (CodTipom , numfactu, fecfactu, numlinea, numfases, Acciones, observac) "
    
    SqlValues = " select " & DBSet(CodTipom, "T") & "," & DBSet(numfactu, "N") & "," & DBSet(Text1(1).Text, "F") & ",numlinea, numfases, Acciones, observac "
    SqlValues = SqlValues & " from rrecibpozos_acc where codtipom = " & DBSet(vTipoM, "T") & " and numfactu = " & DBSet(NumFact, "N") & " and fecfactu = " & DBSet(FecFact, "F")
    
    conn.Execute SqlInsert & SqlValues
    
    ' lineas de factura rrecibpozos_cam
    Mens = "Insertando rrecibpozos_cam"
    
    SqlInsert = "insert into rrecibpozos_cam (codTipom , numfactu, fecfactu, numlinea, codcampo, hanegada, precio1, precio2, codzonas, poligono, parcela, subparce) "
    
    If EsRectificativa Then
        SqlValues = " select " & DBSet(CodTipom, "T") & "," & DBSet(numfactu, "N") & "," & DBSet(Text1(1).Text, "F") & ", numlinea, codcampo, hanegada, precio1 * (-1), precio2 * (-1), codzonas, poligono, parcela, subparce "
        SqlValues = SqlValues & " from rrecibpozos_cam where codtipom = " & DBSet(vTipoM, "T") & " and numfactu = " & DBSet(NumFact, "N") & " and fecfactu = " & DBSet(FecFact, "F")
    Else
        SqlValues = " select " & DBSet(CodTipom, "T") & "," & DBSet(numfactu, "N") & "," & DBSet(Text1(1).Text, "F") & ", numlinea, codcampo, hanegada, "
        SqlValues = SqlValues & "round(precio1 * (1 + (" & DBSet(Text1(0).Text, "N") & "/100)),4), round(precio2 * (1 + (" & DBSet(Text1(0).Text, "N") & "/100)),4), codzonas, poligono, parcela, subparce "
        SqlValues = SqlValues & " from rrecibpozos_cam where codtipom = " & DBSet(vTipoM, "T") & " and numfactu = " & DBSet(NumFact, "N") & " and fecfactu = " & DBSet(FecFact, "F")
    End If
    
    
    conn.Execute SqlInsert & SqlValues
    
    ' lineas de factura rrecibpozos_hid
    Mens = "Insertando rrecibpozos_hid"
    
    SqlInsert = "insert into rrecibpozos_hid (codTipom , numfactu, fecfactu, numlinea, Hidrante, hanegada, nroorden) "
    
    SqlValues = " select " & DBSet(CodTipom, "T") & "," & DBSet(numfactu, "N") & "," & DBSet(Text1(1).Text, "F") & ", numlinea, Hidrante, hanegada, nroorden "
    SqlValues = SqlValues & " from rrecibpozos_hid where  codtipom = " & DBSet(vTipoM, "T") & " and numfactu = " & DBSet(NumFact, "N") & " and fecfactu = " & DBSet(FecFact, "F")
    
    conn.Execute SqlInsert & SqlValues
    
    B = vTipoMov.IncrementarContador(CodTipom)
    
    CrearFactura = B
    Exit Function

eCrearFactura:
    MuestraError Err.Number, "Crear Factura", Mens & vbCrLf & Err.Description
End Function

Private Function CalculoImpRecargo(TipoM As String, NumFact As String, FecFact As String) As Currency
Dim Sql As String
Dim vImporte As Currency
Dim vSeccion As CSeccion
Dim Rs As ADODB.Recordset

    vImporte = 0

' de momento comentado pq lo vamos a calcular sobre el importe de factura
'    Set vSeccion = New CSeccion
'    If vSeccion.LeerDatos(vParamAplic.Seccionhorto) Then
'        If vSeccion.AbrirConta Then
'            Sql = "SELECT sum(coalesce(impvenci,0) + coalesce(gastos,0) - coalesce(impcobro,0)) importe "
'            Sql = Sql & " FROM scobro INNER JOIN usuarios.stipom ON scobro.numserie = stipom.letraser "
'            Sql = Sql & " WHERE stipom.codtipom = " & DBSet(Tipom, "T")
'            Sql = Sql & " and scobro.codfaccl = " & DBSet(NumFact, "N")
'            Sql = Sql & " and scobro.fecfaccl = " & DBSet(FecFact, "F")
'
'            Set Rs = New ADODB.Recordset
'            Rs.Open Sql, ConnConta, adOpenForwardOnly, adLockPessimistic, adCmdText
'            If Not Rs.EOF Then
'                vImporte = DBLet(Rs.Fields(0).Value, "N")
'            End If
'        End If
'    End If
'    Set vSeccion = Nothing

    Sql = "SELECT sum(coalesce(totalfact,0))  from rrecibpozos "
    Sql = Sql & " where codtipom = " & DBSet(TipoM, "T")
    Sql = Sql & " and numfactu = " & DBSet(NumFact, "N")
    Sql = Sql & " and fecfactu = " & DBSet(FecFact, "F")
    
    vImporte = DevuelveValor(Sql)
    

    If ComprobarCero(Text1(0).Text) <> "0" Then
        vImporte = Round2(vImporte * ImporteFormateado(Text1(0).Text) / 100, 2)
    Else
        vImporte = ImporteFormateado(Text1(2).Text)
    End If

    CalculoImpRecargo = vImporte


End Function

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
            End Select
            
        CargarFacturas Text1(3).Text
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
    
End Sub


Private Sub InicializarListView()
'Inicializa las columnas del List view

    ListView1.ListItems.Clear
    
    ListView1.ColumnHeaders.Clear
    ListView1.ColumnHeaders.Add , , "Tipo", 800
    ListView1.ColumnHeaders.Add , , "Descripción", 2400
    ListView1.ColumnHeaders.Add , , "Factura", 1100
    ListView1.ColumnHeaders.Add , , "Fecha", 1100
    ListView1.ColumnHeaders.Add , , "Importe", 1600, 1
    
End Sub



Private Sub imgCheck_Click(Index As Integer)
Dim B As Boolean
Dim TotalArray As Long

    'En el listview3
    B = Index = 1
    For TotalArray = 1 To ListView1.ListItems.Count
    
        ListView1.ListItems(TotalArray).Checked = B
        
        'No dejamos marcar las facturas que sean negativas
        If Index = 1 And ListView1.ListItems(TotalArray).SubItems(4) < 0 Then ListView1.ListItems(TotalArray).Checked = False
        If (TotalArray Mod 50) = 0 Then DoEvents
    Next TotalArray
    CalcularTotales
    
End Sub


Private Sub BotonImprimir()
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String
    
Dim Sql As String
Dim Rs As ADODB.Recordset
    
    
    Sql = "select codtipom from rrecibpozos where imprimir = " & DBSet(vUsu.PC, "T")
    Sql = Sql & " and fecfactu = " & DBSet(Text1(1).Text, "F") & " group by 1 order by 1"
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not Rs.EOF
        
        
        cadFormula = ""
        CadParam = ""
        cadSelect = ""
        numParam = 0
        
        '===================================================
        '============ PARAMETROS ===========================
        Select Case DBLet(Rs!CodTipom)
            Case "RCP"
                indRPT = 46 'Impresion de recibos de consumo de pozos
                cadTitulo = "Reimpresión de Recibos Consumo"
            Case "RMP"
                indRPT = 47 'Impresion de recibos de mantenimiento de pozos
                cadTitulo = "Reimpresión de Recibos Mantenimiento"
            Case "RVP"
                indRPT = 47 'Impresion de recibos de contadores pozos
                cadTitulo = "Reimpresión de Recibos Contadores"
            Case "TAL"
                indRPT = 47 'Impresion de recibos de talla
                cadTitulo = "Reimpresión de Recibos Talla"
            Case "RMT"
                indRPT = 47 'Impresion de recibos de consumo a manta
                cadTitulo = "Reimpresión de Recibos Consumo Manta"
                
            '[Monica]14/01/2016: las rectificativas
            Case "RRC"
                indRPT = 46 ' impresion de recibos de consumo
                cadTitulo = "Reimpresión de Recibos Rect.Consumo"
            Case "RRM"
                indRPT = 47 'Impresion de recibos de mantenimiento de pozos
                cadTitulo = "Reimpresión de Recibos Rect.Mantenimiento"
            Case "RRV"
                indRPT = 47 'Impresion de recibos de contadores pozos
                cadTitulo = "Reimpresión de Recibos Rect.Contadores"
            Case "RTA"
                indRPT = 47 'Impresion de recibos de talla
                cadTitulo = "Reimpresión de Recibos Rect.Talla"
            Case "RRT"
                indRPT = 47 'Impresion de recibos de consumo a manta
                cadTitulo = "Reimpresión de Recibos Rect.Consumo Manta"
        End Select
        
        If Not PonerParamRPT(indRPT, CadParam, numParam, nomDocu) Then Exit Sub
        
        If DBLet(Rs!CodTipom) = "TAL" Then nomDocu = Replace(nomDocu, "Mto.", "Tal.")
        If DBLet(Rs!CodTipom) = "RVP" Then nomDocu = Replace(nomDocu, "Mto.", "Cont.")
        If DBLet(Rs!CodTipom) = "RMT" Then nomDocu = Replace(nomDocu, "Mto.", "Manta.")
          
        '[Monica]14/01/2016: las rectificativas
        If DBLet(Rs!CodTipom) = "RTA" Then nomDocu = Replace(nomDocu, "Mto.", "Tal.")
        If DBLet(Rs!CodTipom) = "RRV" Then nomDocu = Replace(nomDocu, "Mto.", "Cont.")
        If DBLet(Rs!CodTipom) = "RRM" Then nomDocu = Replace(nomDocu, "Mto.", "Manta.")
          
          
        'Nombre fichero .rpt a Imprimir
        frmImprimir.NombreRPT = nomDocu
        
        If vParamAplic.ContabilidadNueva And UCase(Mid(frmImprimir.NombreRPT, 1, 3)) = "ESC" Then frmImprimir.NombreRPT = Replace(frmImprimir.NombreRPT, ".rpt", "6.rpt")
            
        '===================================================
        '================= FORMULA =========================
        'Cadena para seleccion Nº de recibo
        '---------------------------------------------------
            
        'Fecha Factura
        devuelve = "{" & NombreTabla & ".fecfactu}=Date(" & Year(Text1(1).Text) & "," & Month(Text1(1).Text) & "," & Day(Text1(1).Text) & ")"
        If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
        devuelve = "fecfactu = " & DBSet(Text1(1).Text, "F")
        If Not AnyadirAFormula(cadSelect, devuelve) Then Exit Sub
        
        'Socio
        devuelve = "{" & NombreTabla & ".codsocio}=" & Val(Text1(3).Text)
        If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
        devuelve = "codsocio = " & Val(Text1(3).Text)
        If Not AnyadirAFormula(cadSelect, devuelve) Then Exit Sub
        
        ' quien ha generado las facturas
        If Not AnyadirAFormula(cadSelect, "rrecibpozos.imprimir=" & DBSet(vUsu.PC, "T")) Then Exit Sub
        If Not AnyadirAFormula(cadFormula, "{rrecibpozos.imprimir} = """ & vUsu.PC & """") Then Exit Sub
        
        If Not AnyadirAFormula(cadSelect, "rrecibpozos.codtipom=" & DBSet(Rs!CodTipom, "T")) Then Exit Sub
        If Not AnyadirAFormula(cadFormula, "{rrecibpozos.codtipom} = """ & Rs!CodTipom & """") Then Exit Sub
        
        
        
        
        If Not HayRegParaInforme(NombreTabla, cadSelect) Then Exit Sub
         
        With frmImprimir
              '[Monica]06/02/2012: añadido la siguientes 3 lineas para el envio por el outlook
                .outClaveNombreArchiv = "" 'Mid(Combo1(0).Text, 1, 3) & Format(Text1(0).Text, "0000000")
                .outCodigoCliProv = 0
                .outTipoDocumento = 100
                
                .FormulaSeleccion = cadFormula
                .OtrosParametros = CadParam
                .NumeroParametros = numParam
                .SoloImprimir = False
                .EnvioEMail = False
                .Opcion = 0
                .Titulo = cadTitulo '"Impresión de Recibos de Socios"
                
                '[Monica]11/09/2015: pasamos la contabilidad que es pq tenemos que imprimir que gastos de cobros tiene.
                If vParamAplic.Cooperativa = 10 Then
                    vParamAplic.NumeroConta = DevuelveValor("Select empresa_conta from rseccion where codsecci = " & vParamAplic.Seccionhorto)
                End If
                .ConSubInforme = True
                .Show vbModal
        End With
    
        If frmVisReport.EstaImpreso Then
            ActualizarRegistros "rrecibpozos", cadSelect
        End If
    
        Rs.MoveNext
   Wend
   Set Rs = Nothing
    
End Sub



Private Function CalcularTotales()
Dim Rs As ADODB.Recordset
Dim Sql As String
Dim importe1 As Currency
Dim importe2 As Currency
Dim vImporte As Currency
Dim I As Integer

    On Error Resume Next
    
    importe1 = 0
    importe2 = 0
    
    For I = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(I).Checked Then
            If ComprobarCero(Text1(0).Text) <> "0" Then
                vImporte = Round2(ListView1.ListItems(I).SubItems(4) * ImporteFormateado(Text1(0).Text) / 100, 2)
            Else
                vImporte = ImporteFormateado(Text1(2).Text)
            End If
            importe1 = importe1 + ListView1.ListItems(I).SubItems(4)
            
            importe2 = importe2 + ListView1.ListItems(I).SubItems(4) + vImporte
        End If
    Next I
    
    Text2(0).Text = Format(importe1, "###,###,##0.00")
    Text2(1).Text = Format(importe2, "###,###,##0.00")
    
    
    DoEvents
    

End Function




