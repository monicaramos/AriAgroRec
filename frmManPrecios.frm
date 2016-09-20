VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmManPrecios 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Precios"
   ClientHeight    =   6915
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8835
   Icon            =   "frmManPrecios.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6915
   ScaleWidth      =   8835
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "Cuenta Principal|N|N|0|1|cltebanc|ctaprpal|0||"
   Begin VB.Frame FrameAux1 
      BorderStyle     =   0  'None
      Height          =   3930
      Left            =   225
      TabIndex        =   22
      Top             =   2115
      Width           =   8340
      Begin VB.CheckBox Check1 
         Caption         =   "Igualar Precio Socio"
         Height          =   195
         Left            =   1680
         TabIndex        =   34
         Top             =   120
         Width           =   1845
      End
      Begin VB.TextBox txtaux1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   290
         Index           =   5
         Left            =   7335
         MaxLength       =   7
         TabIndex        =   31
         Tag             =   "Precio Socio|N|N|||rprecios_calidad|presocio|#0.0000||"
         Text            =   "pr.soci"
         Top             =   2925
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.TextBox txtaux1 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   290
         Index           =   4
         Left            =   1665
         MaxLength       =   3
         TabIndex        =   27
         Tag             =   "Contador|N|N|1|999|rprecios_calidad|contador|000|S|"
         Text            =   "con"
         Top             =   2925
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.TextBox txtaux1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   290
         Index           =   3
         Left            =   6615
         MaxLength       =   7
         TabIndex        =   30
         Tag             =   "Precio Coop|N|N|||rprecios_calidad|precoop|#0.0000||"
         Text            =   "pr.coop"
         Top             =   2925
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.TextBox txtaux1 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   290
         Index           =   2
         Left            =   945
         MaxLength       =   6
         TabIndex        =   26
         Tag             =   "Tipo Factura|N|N|0|3|rprecios_calidad|tipofact||S|"
         Text            =   "tipo"
         Top             =   2925
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.TextBox txtaux1 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   290
         Index           =   0
         Left            =   225
         MaxLength       =   6
         TabIndex        =   25
         Tag             =   "Código Variedad|N|N|1|999999|rprecios_calidad|codvarie|000000|S|"
         Text            =   "var"
         Top             =   2925
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.TextBox txtaux1 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   290
         Index           =   1
         Left            =   2385
         MaxLength       =   2
         TabIndex        =   28
         Tag             =   "Calidad|N|N|||rprecios_calidad|codcalid|00|S|"
         Text            =   "cal"
         Top             =   2925
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.CommandButton cmdAux 
         Appearance      =   0  'Flat
         Caption         =   "+"
         Height          =   315
         Index           =   1
         Left            =   3060
         TabIndex        =   24
         ToolTipText     =   "Buscar calidad"
         Top             =   2925
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.TextBox txtAux2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   1
         Left            =   3240
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   23
         Text            =   "Nombre calidad"
         Top             =   2925
         Visible         =   0   'False
         Width           =   3285
      End
      Begin MSComctlLib.Toolbar ToolAux 
         Height          =   390
         Index           =   1
         Left            =   60
         TabIndex        =   29
         Top             =   30
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   688
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   3
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Nuevo"
               Object.Tag             =   "2"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Modificar"
               Object.Tag             =   "2"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Eliminar"
               Object.Tag             =   "2"
            EndProperty
         EndProperty
      End
      Begin MSAdodcLib.Adodc AdoAux 
         Height          =   375
         Index           =   1
         Left            =   3720
         Top             =   480
         Visible         =   0   'False
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
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
         Caption         =   "AdoAux(1)"
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
      Begin MSDataGridLib.DataGrid DataGridAux 
         Bindings        =   "frmManPrecios.frx":000C
         Height          =   3195
         Index           =   1
         Left            =   45
         TabIndex        =   32
         Top             =   450
         Width           =   8160
         _ExtentX        =   14393
         _ExtentY        =   5636
         _Version        =   393216
         AllowUpdate     =   0   'False
         BorderStyle     =   0
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            AllowFocus      =   0   'False
            AllowRowSizing  =   0   'False
            AllowSizing     =   0   'False
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.Image imgAyuda 
         Height          =   240
         Index           =   1
         Left            =   3540
         MousePointer    =   4  'Icon
         Tag             =   "-1"
         ToolTipText     =   "Ayuda"
         Top             =   90
         Width           =   240
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1440
      Index           =   0
      Left            =   240
      TabIndex        =   11
      Top             =   495
      Width           =   8340
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   5
         Left            =   2970
         MaxLength       =   7
         TabIndex        =   5
         Tag             =   "Precio Industria|N|S|||rprecios|precioindustria|#0.0000||"
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   4
         Left            =   4245
         MaxLength       =   30
         TabIndex        =   6
         Tag             =   "Texto|T|N|||rprecios|textoper|||"
         Text            =   "123456789012345678901234567890"
         Top             =   1080
         Width           =   3825
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   3
         Left            =   1665
         MaxLength       =   10
         TabIndex        =   4
         Tag             =   "Fecha Fin|F|S|||rprecios|fechafin|dd/mm/yyyy||"
         Text            =   "1234567890"
         Top             =   1080
         Width           =   1110
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   2
         Left            =   225
         MaxLength       =   10
         TabIndex        =   3
         Tag             =   "Fecha Inicio|F|N|||rprecios|fechaini|dd/mm/yyyy||"
         Text            =   "1234567890"
         Top             =   1080
         Width           =   1200
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   0
         Left            =   5625
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Tag             =   "Tipo Factura|N|N|0|1|rprecios|tipofact||S|"
         Top             =   450
         Width           =   1485
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   225
         MaxLength       =   6
         TabIndex        =   0
         Tag             =   "Código Variedad|N|N|1|999999|rprecios|codvarie|000000|S|"
         Top             =   450
         Width           =   1095
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   1350
         MaxLength       =   40
         TabIndex        =   17
         Top             =   450
         Width           =   4050
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   1
         Left            =   7215
         MaxLength       =   3
         TabIndex        =   2
         Tag             =   "Contador|N|N|1|999|rprecios|contador|000|S|"
         Top             =   480
         Width           =   855
      End
      Begin VB.Image imgAyuda 
         Height          =   240
         Index           =   0
         Left            =   6600
         MousePointer    =   4  'Icon
         Tag             =   "-1"
         ToolTipText     =   "Ayuda"
         Top             =   210
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Pr.Industria"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   1
         Left            =   2970
         TabIndex        =   33
         Top             =   810
         Width           =   1005
      End
      Begin VB.Label Label8 
         Caption         =   "Texto"
         Height          =   255
         Left            =   4245
         TabIndex        =   21
         Top             =   810
         Width           =   705
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   1
         Left            =   2520
         Picture         =   "frmManPrecios.frx":0024
         ToolTipText     =   "Buscar fecha"
         Top             =   810
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   0
         Left            =   1170
         Picture         =   "frmManPrecios.frx":00AF
         ToolTipText     =   "Buscar fecha"
         Top             =   810
         Width           =   240
      End
      Begin VB.Label Label22 
         Caption         =   "Fecha Fin"
         Height          =   255
         Left            =   1665
         TabIndex        =   20
         Top             =   810
         Width           =   975
      End
      Begin VB.Label Label18 
         Caption         =   "Fecha Inicio"
         Height          =   255
         Left            =   225
         TabIndex        =   19
         Top             =   810
         Width           =   930
      End
      Begin VB.Label Label9 
         Caption         =   "Tipo Factura"
         Height          =   255
         Left            =   5625
         TabIndex        =   18
         Top             =   255
         Width           =   1035
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   990
         ToolTipText     =   "Buscar Variedad"
         Top             =   180
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Variedad"
         Height          =   255
         Left            =   270
         TabIndex        =   13
         Top             =   225
         Width           =   690
      End
      Begin VB.Label Label1 
         Caption         =   "Contador"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   7215
         TabIndex        =   12
         Top             =   255
         Width           =   1005
      End
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Index           =   0
      Left            =   195
      TabIndex        =   9
      Top             =   6210
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
         TabIndex        =   10
         Top             =   180
         Width           =   2655
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   7335
      TabIndex        =   8
      Top             =   6300
      Width           =   1035
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   6120
      TabIndex        =   7
      Top             =   6300
      Width           =   1035
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   3690
      Top             =   6345
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
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
      Height          =   360
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   8835
      _ExtentX        =   15584
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   21
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Buscar"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Ver todos"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Nuevo"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Modificar"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Eliminar"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Generación Masiva"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Copiar Variedad"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir"
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Primero"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Anterior"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Siguiente"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Último"
            ImageIndex      =   9
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   195
         Index           =   0
         Left            =   7155
         TabIndex        =   16
         Top             =   90
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   7335
      TabIndex        =   14
      Top             =   6300
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Menu mnOpciones 
      Caption         =   "&Opciones"
      Begin VB.Menu mnBuscar 
         Caption         =   "&Buscar"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnVerTodos 
         Caption         =   "&Ver Todos"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnBarra1 
         Caption         =   "-"
      End
      Begin VB.Menu mnNuevo 
         Caption         =   "&Nuevo"
         HelpContextID   =   2
         Shortcut        =   ^N
      End
      Begin VB.Menu mnModificar 
         Caption         =   "&Modificar"
         HelpContextID   =   2
         Shortcut        =   ^M
      End
      Begin VB.Menu mnEliminar 
         Caption         =   "&Eliminar"
         HelpContextID   =   2
         Shortcut        =   ^E
      End
      Begin VB.Menu mnBarra3 
         Caption         =   "-"
      End
      Begin VB.Menu mnGeneracionMasiva 
         Caption         =   "&Generacion Masiva"
         Shortcut        =   ^G
      End
      Begin VB.Menu mnCopia 
         Caption         =   "Copia a otra Variedad"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnBarra4 
         Caption         =   "-"
      End
      Begin VB.Menu mnImprimir 
         Caption         =   "&Imprimir"
         Shortcut        =   ^I
      End
      Begin VB.Menu mnBarra2 
         Caption         =   "-"
      End
      Begin VB.Menu mnSalir 
         Caption         =   "&Salir"
         Shortcut        =   ^S
      End
   End
End
Attribute VB_Name = "frmManPrecios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' +-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+
' +-+- Autor: CÈSAR                    -+-+
' +-+- Menú: General-Clientes-Clientes -+-+
' +-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+

' +-+-+-+- DISSENY +-+-+-+-
' 1. Posar tots els controls al formulari
' 2. Posar els index correlativament
' 3. Si n'hi han botons de buscar repasar el ToolTipText
' 4. Alliniar els camps numérics a la dreta i el resto a l'esquerra
' 5. Posar els TAGs
' (si es INTEGER: si PK => mínim 1; si no PK => mínim 0; màxim => 99; format => 00)
' (si es DECIMAL; mínim => 0; màxim => 99.99; format => #,###,###,##0.00)
' (si es DATE; format => dd/mm/yyyy)
' 6. Posar els MAXLENGTHs
' 7. Posar els TABINDEXs

Option Explicit

'Dim T1 As Single

Public DatosADevolverBusqueda As String    'Tindrà el nº de text que vol que torne, empipat
Public Event DatoSeleccionado(CadenaSeleccion As String)

Public Tipo As Byte ' 0 = variedades que no son del grupo 5 ni 6
                    ' 1 = variedades del grupo 5 (almazara)
                    ' 2 = variedades del grupo 6 (bodega)

Private CadB1 As String
' *** declarar els formularis als que vaig a cridar ***
Private WithEvents frmB As frmBuscaGrid
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmC As frmCal 'calendario fecha
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmC1 As frmCal 'calendario fecha
Attribute frmC1.VB_VarHelpID = -1
Private WithEvents frmZ As frmZoom  'Zoom para campos Text
Attribute frmZ.VB_VarHelpID = -1
Private WithEvents frmGen As frmGeneraPrecios 'Generacion de precios masiva
Attribute frmGen.VB_VarHelpID = -1

Private WithEvents frmVar As frmComVar 'variedades
Attribute frmVar.VB_VarHelpID = -1
Private WithEvents frmCalid As frmManCalidades 'calidades
Attribute frmCalid.VB_VarHelpID = -1

' *****************************************************


Private Modo As Byte
'*************** MODOS ********************
'   0.-  Formulari net sense cap camp ple
'   1.-  Preparant per a fer la búsqueda
'   2.-  Ja tenim registres i els anem a recorrer
'        ,podem editar-los Edició del camp
'   3.-  Inserció de nou registre
'   4.-  Modificar
'   5.-  Manteniment Llinies

'+-+-Variables comuns a tots els formularis+-+-+

Dim ModoLineas As Byte
'1.- Afegir,  2.- Modificar,  3.- Borrar,  0.-Passar control a Llínies

Dim NumTabMto As Integer 'Indica quin nº de Tab està en modo Mantenimient
Dim TituloLinea As String 'Descripció de la llínia que està en Mantenimient
Dim PrimeraVez As Boolean

Private CadenaConsulta As String 'SQL de la taula principal del formulari
Private Ordenacion As String
Private NombreTabla As String  'Nom de la taula
Private NomTablaLineas As String 'Nom de la Taula de llínies del Mantenimient en que estem

Private kCampo As Integer
'-------------------------------------------------------------------------
Private HaDevueltoDatos As Boolean

'Private VieneDeBuscar As Boolean
'Per a quan torna 2 poblacions en el mateix codi Postal. Si ve de pulsar prismatic
'de búsqueda posar el valor de població seleccionada i no tornar a recuperar de la Base de Datos

Dim btnPrimero As Byte 'Variable que indica el nº del Botó PrimerRegistro en la Toolbar1
'Dim CadAncho() As Boolean  'array, per a quan cridem al form de llínies
Dim indice As Byte 'Index del text1 on es poses els datos retornats des d'atres Formularis de Mtos
Dim CadB As String

Dim vSeccion As CSeccion
Dim b As Boolean

Private BuscaChekc As String

'GENERALES PARA PASARLE A CRYSTAL REPORT
Private cadFormula As String 'Cadena con la FormulaSelection para Crystal Report
Private CadParam As String 'Cadena con los parametros para Crystal Report
Private numParam As Byte 'Numero de parametros que se pasan a Crystal Report
Private cadSelect As String 'Cadena para comprobar si hay datos antes de abrir Informe
Private cadTitulo As String 'Titulo para la ventana frmImprimir
Private cadNombreRPT As String 'Nombre del informe

Dim VarieAnt As String


Private Sub cmdAceptar_Click()

    Screen.MousePointer = vbHourglass
    
    On Error GoTo Error1
    
    Select Case Modo
        Case 1  'BÚSQUEDA
            HacerBusqueda
            
        Case 3 'INSERTAR
            If DatosOk Then
                If InsertarDesdeForm2(Me, 1) Then
                    ' *** canviar o llevar el WHERE, repasar codEmpre ****
                    Data1.RecordSource = "Select * from " & NombreTabla & Ordenacion
                    'Data1.RecordSource = "Select * from " & NombreTabla & " where codempre = " & codEmpre & Ordenacion
                    ' ***************************************************************
                    PosicionarData
                    PonerCampos
                End If
            Else
                ModoLineas = 0
            End If
            
        Case 4  'MODIFICAR
            If DatosOk Then
                Modificar
'                If ModificaDesdeFormulario2(Me, 1) Then
'                    TerminaBloquear
'                    PosicionarData
'                    CargaGrid 1, True
'                End If
            Else
                ModoLineas = 0
            End If
        ' *** si n'hi han llínies ***
        Case 5 'LLÍNIES
            Select Case ModoLineas
                Case 1 'afegir llínia
                    InsertarLinea
                Case 2 'modificar llínies
                    ModificarLinea
                    PosicionarData
                    SumaTotalPorcentajes
            End Select
        ' **************************
'            If NumTabMto = 1 Then
'                If Not vSeccion Is Nothing Then
'                    vSeccion.CerrarConta
'                    Set vSeccion = Nothing
'                End If
'            End If
    
    End Select
    Screen.MousePointer = vbDefault
    
    
Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub



Private Sub cmdAux_Click(Index As Integer)
    TerminaBloquear
    Select Case Index
        Case 1 'Calidades de la variedad de cabecera
            Set frmCalid = New frmManCalidades
            frmCalid.DatosADevolverBusqueda = "0|1|2|3|"
            frmCalid.CodigoActual = txtAux1(1).Text
            frmCalid.ParamVariedad = txtAux1(0).Text
            frmCalid.Show vbModal
            Set frmCalid = Nothing
            PonerFoco txtAux1(1)

    End Select
    If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
End Sub

' *** si n'hi han combos a la capçalera ***
Private Sub Combo1_GotFocus(Index As Integer)
    If Modo = 1 Then Combo1(Index).BackColor = vbYellow
End Sub

Private Sub Combo1_LostFocus(Index As Integer)
    If Combo1(Index).BackColor = vbYellow Then Combo1(Index).BackColor = vbWhite
    
    ' ofertamos el siguiente contador para la variedad / tipo
    If Text1(0).Text <> "" And Modo = 3 Then
        Text1(1).Text = SugerirCodigoSiguienteStr("rprecios", "contador", "codvarie = " & Text1(0).Text & " and tipofact = " & Combo1(0).ListIndex)
    End If
    
    '[Monica]10/04/2013: Añadido pq se quedaba bloqueado el precio de industria
    '[Monica]27/01/2016: Añadido el precio de complementaria unica
    If Tipo = 0 Then
        Text1(5).Enabled = (Combo1(0).ListIndex = 2) Or (Combo1(0).ListIndex = 4) Or (Combo1(0).ListIndex = 5) Or (Combo1(0).ListIndex = 6) Or Modo = 1
    Else
        Text1(5).Enabled = True
    End If
    
End Sub

Private Sub Combo1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
    If PrimeraVez Then PrimeraVez = False
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Form_Unload(Cancel As Integer)
    CheckValueGuardar Me.Name, Me.chkVistaPrevia(0).Value
    Screen.MousePointer = vbDefault
    If Modo = 4 Or Modo = 5 Then TerminaBloquear
End Sub

Private Sub Form_Load()
Dim i As Integer

    PrimeraVez = True
    
    ' ICONETS DE LA BARRA
    btnPrimero = 18 'index del botó "primero"
    With Me.Toolbar1
        .HotImageList = frmPpal.imgListComun_OM
        .DisabledImageList = frmPpal.imgListComun_BN
        .ImageList = frmPpal.imgListComun
        'l'1 i el 2 son separadors
        .Buttons(3).Image = 1   'Buscar
        .Buttons(4).Image = 2   'Totss
        'el 5 i el 6 son separadors
        .Buttons(7).Image = 3   'Insertar
        .Buttons(8).Image = 4   'Modificar
        .Buttons(9).Image = 5   'Borrar
        'el 10 i el 11 son separadors
        .Buttons(11).Image = 17  ' generacion masiva
        .Buttons(12).Image = 16  ' copia a otra variedad
        
        .Buttons(14).Image = 10  'Imprimir
        .Buttons(15).Image = 11  'Eixir
        'el 13 i el 14 son separadors
        .Buttons(btnPrimero).Image = 6  'Primer
        .Buttons(btnPrimero + 1).Image = 7 'Anterior
        .Buttons(btnPrimero + 2).Image = 8 'Següent
        .Buttons(btnPrimero + 3).Image = 9 'Últim
    End With
    
    ' ******* si n'hi han llínies *******
    'ICONETS DE LES BARRES ALS TABS DE LLÍNIA
    For i = 1 To ToolAux.Count
        With Me.ToolAux(i)
            .HotImageList = frmPpal.imgListComun_OM16
            .DisabledImageList = frmPpal.imgListComun_BN16
            .ImageList = frmPpal.imgListComun16
            .Buttons(1).Image = 3   'Insertar
            .Buttons(2).Image = 4   'Modificar
            .Buttons(3).Image = 5   'Borrar
        End With
    Next i
    ' ***********************************
    
    For i = 0 To imgAyuda.Count - 1
        imgAyuda(i).Picture = frmPpal.ImageListB.ListImages(10).Picture
    Next i
    
    If Tipo <> 0 Then
        imgAyuda(0).Enabled = False
        imgAyuda(0).visible = False
    End If
    
    'cargar IMAGES de busqueda
    For i = 0 To Me.imgBuscar.Count - 1
        Me.imgBuscar(i).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next i
'    'carga IMAGES de mail
'    For i = 0 To Me.imgMail.Count - 1
'        Me.imgMail(i).Picture = frmPpal.imgListImages16.ListImages(2).Picture
'    Next i
    'IMAGES para zoom
'    For I = 0 To Me.imgZoom.Count - 1
'        Me.imgZoom(I).Picture = frmPpal.imgListImages16.ListImages(3).Picture
'    Next I
    
    
'    ' *** si n'hi han tabs, per a que per defecte sempre es pose al 1r***
'    Me.SSTab1.Tab = 0
'    ' *******************************************************************
    
    LimpiarCampos   'Neteja els camps TextBox
'    ' ******* si n'hi han llínies *******
'    DataGridAux(0).ClearFields
    DataGridAux(1).ClearFields
    
    '*** canviar el nom de la taula i l'ordenació de la capçalera ***
    NombreTabla = "rprecios"
    Ordenacion = " ORDER BY codvarie"
    '************************************************
    
    'Mirem com està guardat el valor del check
    chkVistaPrevia(0).Value = CheckValueLeer(Name)
    
    ' dependiendo desde donce se llame el formulario seleccionamos registros segun variedades
    Select Case Tipo
        Case 0  ' variedades que no sean de almazara ni bodega
            CadB1 = " rprecios.codvarie not in (select codvarie from variedades inner join productos on "
            CadB1 = CadB1 & " variedades.codprodu = productos.codprodu and (productos.codgrupo = 5 or productos.codgrupo = 6))"
            
            '[Monica]25/06/2012: solo quatretonda puede ver los precios de almazara en horto
            If vParamAplic.Cooperativa = 7 Then
                CadB1 = " rprecios.codvarie not in (select codvarie from variedades inner join productos on "
                CadB1 = CadB1 & " variedades.codprodu = productos.codprodu and (productos.codgrupo = 6))"
            End If
            
        Case 1 ' variedades de almazara
            CadB1 = " rprecios.codvarie in (select codvarie from variedades inner join productos on "
            CadB1 = CadB1 & " variedades.codprodu = productos.codprodu and productos.codgrupo = 5)"
            
            Me.Caption = "Precios de Almazara"
            
        Case 2 ' variedades de bodega
            CadB1 = " rprecios.codvarie in (select codvarie from variedades inner join productos on "
            CadB1 = CadB1 & " variedades.codprodu = productos.codprodu and productos.codgrupo = 6)"
            
            Me.Caption = "Precios de Bodega"
    End Select
    
    If Tipo > 0 Then Label1(1).Caption = "Precio"
    
    'ASignamos un SQL al DATA1
    Data1.ConnectionString = conn
    '***** canviar el nom de la PK de la capçalera; repasar codEmpre *************
    Data1.RecordSource = "Select * from " & NombreTabla & " where codvarie=-1" & " and " & CadB1
    
    Data1.Refresh
       
    ' ******* si n'hi han llinies en datagrid *******
'    ReDim CadAncho(DataGridAux.Count) 'redimensione l'array a la quantitat de datagrids
'    CadAncho(0) = False
'    CadAncho(1) = False
'    CadAncho(2) = False
'    CadAncho(4) = False
    
    ModoLineas = 0
       
    ' **** si n'hi ha algun frame que no te datagrids ***
'    CargaFrame 3, False
    ' *************************************************
         
    ' *** si n'hi han combos (capçalera o llínies) ***
    CargaCombo
    ' ************************************************
    
    If DatosADevolverBusqueda = "" Then
        PonerModo 0
    Else
        PonerModo 1 'búsqueda
        ' *** posar de groc els camps visibles de la clau primaria de la capçalera ***
        Text1(0).BackColor = vbYellow 'codclien
        ' ****************************************************************************
    End If
End Sub


Private Sub LimpiarCampos()
Dim i As Integer

    On Error Resume Next
    
    limpiar Me   'Mètode general: Neteja els controls TextBox
    lblIndicador.Caption = ""
'    Me.chkAbonos(0).Value = 0
    
    For i = 0 To Combo1.Count - 1
        Combo1(i).ListIndex = -1
'        Me.chkAbonos(I).Value = 0
    Next i
    
    ' *** si n'hi han combos a la capçalera ***
    ' *****************************************

    If Err.Number <> 0 Then Err.Clear
End Sub

Private Sub LimpiarCamposLin(FrameAux As String)
    On Error Resume Next
    
    LimpiarLin Me, FrameAux  'Mètode general: Neteja els controls TextBox
    lblIndicador.Caption = ""

    If Err.Number <> 0 Then Err.Clear
End Sub

'----------------------------------------------------------------
'----------------------------------------------------------------
'   En PONERMODO s'habiliten, o no, els diversos camps del
'   formulari en funció del modo en que anem a treballar
Private Sub PonerModo(Kmodo As Byte, Optional indFrame As Integer)
Dim i As Integer, NumReg As Byte
Dim b As Boolean

    On Error GoTo EPonerModo
 
    'Actualisa Iconos Insertar,Modificar,Eliminar
    'ActualizarToolbar Modo, Kmodo
    Modo = Kmodo
    PonerIndicador lblIndicador, Modo, ModoLineas
    
    BuscaChekc = ""
       
    'Modo 2. N'hi han datos i estam visualisant-los
    '=========================================
    'Posem visible, si es formulari de búsqueda, el botó "Regresar" quan n'hi han datos
    If DatosADevolverBusqueda <> "" Then
        cmdRegresar.visible = (Modo = 2)
    Else
        cmdRegresar.visible = False
    End If
    
    '=======================================
    b = (Modo = 2)
    'Posar Fleches de desplasament visibles
    NumReg = 1
    If Not Data1.Recordset.EOF Then
        If Data1.Recordset.RecordCount > 1 Then NumReg = 2 'Només es per a saber que n'hi ha + d'1 registre
    End If
    DesplazamientoVisible Me.Toolbar1, btnPrimero, b, NumReg
    
    '---------------------------------------------
    b = Modo <> 0 And Modo <> 2
    cmdCancelar.visible = b
    cmdAceptar.visible = b
       
    'Bloqueja els camps Text1 si no estem modificant/Insertant Datos
    'Si estem en Insertar a més neteja els camps Text1
    BloquearText1 Me, Modo
    BloquearCombo Me, Modo
    
    ' **** si n'hi han imagens de buscar en la capçalera *****
    BloquearImgBuscar Me, Modo, ModoLineas
    For i = 0 To imgFec.Count - 1
        BloquearImgFec Me, i, Modo
    Next i
'    BloquearImgFec Me, 25, Modo, ModoLineas
    BloquearImgZoom Me, Modo, ModoLineas
    ' ********************************************************
    ' *** si n'hi han llínies i imagens de buscar que no estiguen als grids ******
    'Llínies Departaments
    b = (Modo = 5 And (ModoLineas = 1 Or ModoLineas = 2) And (NumTabMto = 0))
    ' ****************************************************************************
            
    chkVistaPrevia(0).Enabled = (Modo <= 2)
    
    PonerLongCampos

    If (Modo < 2) Or (Modo = 3) Then
        CargaGrid 1, False
    End If
    
    b = (Modo = 4) Or (Modo = 2)
    

    DataGridAux(1).Enabled = b
    
    'clasificacion
    b = (Modo = 5) And (NumTabMto = 1) 'And (ModoLineas <> 3)
    For i = 1 To txtAux1.Count - 1
        BloquearTxt txtAux1(i), Not b
    Next i
    b = (Modo = 5) And (NumTabMto = 1) And ModoLineas = 2
    BloquearTxt txtAux1(1), b
    BloquearBtn cmdAux(1), b
     '-----------------------------
    PonerModoOpcionesMenu (Modo) 'Activar opcions menú según modo
    PonerOpcionesMenu   'Activar opcions de menú según nivell
                        'de permisos de l'usuari

EPonerModo:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub

Private Sub PonerLongCampos()
'Modificar el MaxLength del campo en funcion de si es modo de búsqueda o no
'para los campos que permitan introducir criterios más largos del tamaño del campo
'Si maxlength=6 para codprove, en modo busqueda Maxlenth=13 'el doble + 1
    
    'para los TEXT1
    PonerLongCamposGnral Me, Modo, 1
End Sub

Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
End Sub

Private Sub PonerModoOpcionesMenu(Modo)
'Actives unes Opcions de Menú i Toolbar según el modo en que estem
Dim b As Boolean, bAux As Boolean
Dim i As Byte
    
    'Barra de CAPÇALERA
    '------------------------------------------
    'b = (Modo = 2 Or Modo = 0 Or Modo = 1)
    b = (Modo = 2 Or Modo = 0)
    'Buscar
    Toolbar1.Buttons(3).Enabled = b
    Me.mnBuscar.Enabled = b
    'Vore Tots
    Toolbar1.Buttons(4).Enabled = b
    Me.mnVerTodos.Enabled = b
    
    'Insertar
    Toolbar1.Buttons(7).Enabled = b
    Me.mnNuevo.Enabled = b
    
    b = (Modo = 2 And Data1.Recordset.RecordCount > 0) 'And Not DeConsulta
    'Modificar
    Toolbar1.Buttons(8).Enabled = b
    Me.mnModificar.Enabled = b
    'eliminar
    Toolbar1.Buttons(9).Enabled = b
    Me.mnEliminar.Enabled = b
'    'generar
'    Toolbar1.Buttons(11).Enabled = b
'    'copiar
'    Toolbar1.Buttons(12).Enabled = b
'    'Imprimir
'    Toolbar1.Buttons(14).Enabled = b
    
    ' *** si n'hi han llínies que tenen grids (en o sense tab) ***
    b = (Modo = 3 Or Modo = 4 Or Modo = 2)
    For i = 1 To ToolAux.Count
        ToolAux(i).Buttons(1).Enabled = b
        If b Then bAux = (b And Me.AdoAux(i).Recordset.RecordCount > 0)
        ToolAux(i).Buttons(2).Enabled = bAux
        ToolAux(i).Buttons(3).Enabled = bAux
    Next i
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


Private Sub Desplazamiento(Index As Integer)
'Botons de Desplaçament; per a desplaçar-se pels registres de control Data
    If Data1.Recordset.EOF Then Exit Sub
    DesplazamientoData Data1, Index
    PonerCampos
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
Dim Sql As String
Dim Tabla As String
    
    ' ********* si n'hi han tabs, dona igual si en datagrid o no ***********
    Select Case Index
       Case 1 ' calidades
            Tabla = "rprecios_calidad"
            Sql = "SELECT rprecios_calidad.codvarie, rprecios_calidad.tipofact, rprecios_calidad.contador, rprecios_calidad.codcalid, rcalidad.nomcalid, rprecios_calidad.precoop, rprecios_calidad.presocio "
            Sql = Sql & " FROM " & Tabla & " INNER JOIN rcalidad ON rprecios_calidad.codvarie = rcalidad.codvarie "
            Sql = Sql & " and rprecios_calidad.codcalid = rcalidad.codcalid "
            If enlaza Then
                Sql = Sql & ObtenerWhereCab(True)
            Else
                Sql = Sql & " WHERE rprecios_calidad.codvarie = -1"
            End If
            Sql = Sql & " ORDER BY " & Tabla & ".codcalid "
            
            
    End Select
    ' ********************************************************************************
    
    MontaSQLCarga = Sql
End Function

Private Sub frmB_Selecionado(CadenaDevuelta As String)
    Dim Aux As String
    
    If CadenaDevuelta <> "" Then
        HaDevueltoDatos = True
        Screen.MousePointer = vbHourglass
        'Sabem quins camps son els que mos torna
        'Creem una cadena consulta i posem els datos
        CadB = ""
        Aux = ValorDevueltoFormGrid(Text1(0), CadenaDevuelta, 1)
        CadB = Aux
        Aux = ValorDevueltoFormGrid(Combo1(0), CadenaDevuelta, 2)
        CadB = CadB & " and " & Aux
        Aux = ValorDevueltoFormGrid(Text1(1), CadenaDevuelta, 3)
        CadB = CadB & " and " & Aux
        
        '   Com la clau principal es única, en posar el sql apuntant
        '   al valor retornat sobre la clau ppal es suficient
        ' *** canviar o llevar el WHERE; repasar codEmpre ***
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " " & Ordenacion
        'CadenaConsulta = "select * from " & NombreTabla & " WHERE codempre = " & codEmpre & " AND " & CadB & " " & Ordenacion
        ' **********************************
        PonerCadenaBusqueda
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub frmC_Selec(vFecha As Date)
Dim indice As Byte
    indice = CByte(Me.cmdAux(0).Tag + 2)
    txtAux1(indice).Text = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub frmC1_Selec(vFecha As Date)
Dim indice As Byte
    indice = CByte(Me.imgFec(0).Tag)
    Text1(indice).Text = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub frmCalid_DatoSeleccionado(CadenaSeleccion As String)
    txtAux1(0).Text = RecuperaValor(CadenaSeleccion, 1) 'codigo variedad
    txtAux1(1).Text = RecuperaValor(CadenaSeleccion, 3) 'codigo calidad
    txtAux2(1).Text = RecuperaValor(CadenaSeleccion, 4) 'nombre calidad
End Sub


Private Sub frmVar_DatoSeleccionado(CadenaSeleccion As String)
    Text1(0).Text = RecuperaValor(CadenaSeleccion, 1) 'codvariedad
    FormateaCampo Text1(0)
    Text2(0).Text = RecuperaValor(CadenaSeleccion, 2) 'nombre variedad
End Sub

Private Sub frmZ_Actualizar(vCampo As String)
     Text1(indice).Text = vCampo
End Sub

Private Sub imgAyuda_Click(Index As Integer)
Dim vCadena As String
    Select Case Index
        Case 0
'            '[Monica]27/01/2016: distinguimos entre catadau y el resto pq tiene un tipo nuevo de liquidacion complementaria unica
'            If vParamAplic.Cooperativa <> 0 Then
'                vCadena = "Cuando el Tipo de Factura es de Industria, Anticipo Genérico o " & vbCrLf & _
'                          "Anticipo de Retirada el precio se introduce en el Precio Industria." & vbCrLf & vbCrLf & _
'                          "En Anticipo Genérico y Anticipo de Retirada es un precio único para" & vbCrLf & _
'                          "todas las calidades, ya que se realiza sobre entradas clasificadas " & vbCrLf & _
'                          "como pendientes de clasificar (Histórico y Clasificadas)" & vbCrLf & _
'                          "" & vbCrLf & vbCrLf
'            Else
                vCadena = "Cuando el Tipo de Factura es de Industria, Anticipo Genérico,  " & vbCrLf & _
                          "Anticipo de Retirada o Complem.Única el precio se introduce en el " & vbCrLf & _
                          "Precio Industria." & vbCrLf & vbCrLf & _
                          "En Anticipo Genérico y Anticipo de Retirada es un precio único para" & vbCrLf & _
                          "todas las calidades, ya que se realiza sobre entradas clasificadas " & vbCrLf & _
                          "como pendientes de clasificar (Histórico y Clasificadas)" & vbCrLf & vbCrLf & _
                          "En la complementaria única tambien es un precio único para todas " & vbCrLf & _
                          "las calidades." & vbCrLf & _
                          "" & vbCrLf & vbCrLf
'            End If
                      
        Case 1
           ' "____________________________________________________________"
            vCadena = "Cuando esté marcado e insertemos/modifiquemos precios de calidades" & vbCrLf & _
                      "de entradas de horto, el precio de cooperativa se copia en el precio" & vbCrLf & _
                      "de socio y se situa el cursor en el boton de Aceptar." & vbCrLf & vbCrLf & _
                      "En caso contrario se pide el precio de cooperativa y el precio de socio." & vbCrLf & vbCrLf
    End Select
    MsgBox vCadena, vbInformation, "Descripción de Ayuda"

End Sub

Private Sub imgFec_Click(Index As Integer)
       
       Screen.MousePointer = vbHourglass
       
       Dim esq As Long
       Dim dalt As Long
       Dim menu As Long
       Dim obj As Object
    
       Set frmC1 = New frmCal
        
       esq = imgFec(Index).Left
       dalt = imgFec(Index).Top
        
       Set obj = imgFec(Index).Container
    
       While imgFec(Index).Parent.Name <> obj.Name
            esq = esq + obj.Left
            dalt = dalt + obj.Top
            Set obj = obj.Container
       Wend
        
       menu = Me.Height - Me.ScaleHeight 'ací tinc el heigth del menú i de la toolbar
    
       frmC1.Left = esq + imgFec(Index).Parent.Left + 30
       frmC1.Top = dalt + imgFec(Index).Parent.Top + imgFec(Index).Height + menu - 40
    
       
       frmC1.NovaData = Now
       Select Case Index
            Case 0, 1
                indice = Index + 2
       End Select
       
       Me.imgFec(0).Tag = indice
       
       PonerFormatoFecha Text1(indice)
       If Text1(indice).Text <> "" Then frmC1.NovaData = CDate(Text1(indice).Text)
    
       Screen.MousePointer = vbDefault
       frmC1.Show vbModal
       Set frmC1 = Nothing
       PonerFoco Text1(indice)
    
End Sub

Private Sub imgMail_Click(Index As Integer)
''Abrir Outlook para enviar e-mail
'Dim dirMail As String
'
''    If Modo = 0 Then Exit Sub
'    Screen.MousePointer = vbHourglass
'
'    Select Case Index
'        Case 0: dirMail = Text1(12).Text
'        Case 1: dirMail = txtAux(11).Text
'    End Select
'
'    If LanzaMailGnral(dirMail) Then espera 2
'    Screen.MousePointer = vbDefault
End Sub

Private Sub imgZoom_Click(Index As Integer)
    
    Set frmZ = New frmZoom

    Select Case Index
        Case 0
            indice = 21
            frmZ.pTitulo = "Observaciones del Campo"
            frmZ.pValor = Text1(indice).Text
            frmZ.pModo = Modo
            frmZ.Show vbModal
            Set frmZ = Nothing
            PonerFoco Text1(indice)
    End Select
            
End Sub

Private Sub mnBuscar_Click()
    BotonBuscar
End Sub

Private Sub mnCopia_click()
    
    Set frmGen = New frmGeneraPrecios
    frmGen.OpcionGenerar = 1
    frmGen.Show vbModal
    Set frmGen = Nothing

End Sub

Private Sub mnEliminar_Click()
    BotonEliminar
End Sub

Private Sub mnImprimir_Click()
Dim cadena As String

'    Select Case tipo
'        Case 0  ' variedades que no sean de almazara ni bodega
'            Cadena = "{productos.codgrupo} <> 5 and {productos.codgrupo} <> 6"
'        Case 1 ' variedades de almazara
'            Cadena = "{productos.codgrupo} = 5"
'        Case 2 ' variedades de bodega
'            Cadena = "{productos.codgrupo} = 6"
'    End Select
'
'    With frmImprimir
'        .FormulaSeleccion = Cadena
'        .OtrosParametros = "pEmpresa='" & vEmpresa.nomempre & "'|"
'        .NumeroParametros = 1
'        .SoloImprimir = False
'        .EnvioEMail = False
'        .Titulo = "Listado de Precios"
'        .NombreRPT = "rManPrecios.rpt"
'        .ConSubInforme = True
'        .Opcion = 0
'        .Show vbModal
'    End With

    frmListado.OpcionListado = 44
    frmListado.Show vbModal



End Sub

Private Sub mnModificar_Click()
    ' *** repasar el nom de l'adodc, l'index del Field i el camp que te la PK ***
    'El registre de codi 0 no es pot Modificar ni Eliminar
    If EsCodigoCero(CStr(Data1.Recordset.Fields(0).Value), FormatoCampo(Text1(0))) Then Exit Sub
    ' ***************************************************************************
    
    If BLOQUEADesdeFormulario2(Me, Data1, 1) Then BotonModificar
End Sub

Private Sub mnNuevo_Click()
    BotonAnyadir
End Sub

Private Sub mnSalir_Click()
    Unload Me
End Sub


Private Sub mnGeneracionMasiva_Click()
Dim Sql As String
Dim Rs As ADODB.Recordset
    
    Set frmGen = New frmGeneraPrecios
    frmGen.OpcionGenerar = 0
    frmGen.Show vbModal
    Set frmGen = Nothing
    
End Sub


Private Sub mnVerTodos_Click()
    BotonVerTodos
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.Index
        Case 3  'Búscar
           mnBuscar_Click
        Case 4  'Tots
            mnVerTodos_Click
        Case 7  'Nou
            mnNuevo_Click
        Case 8  'Modificar
            mnModificar_Click
        Case 9  'Borrar
            mnEliminar_Click
        Case 11 ' Generacion masiva
            mnGeneracionMasiva_Click
        Case 12 ' Copia a otra variedad
            mnCopia_click
        Case 14 'Imprimir
'            AbrirListado (10)
            mnImprimir_Click
        Case 15    'Eixir
            mnSalir_Click
            
        Case btnPrimero To btnPrimero + 3 'Fleches Desplaçament
            Desplazamiento (Button.Index - btnPrimero)
    End Select
End Sub

Private Sub BotonBuscar()
Dim i As Integer
' ***** Si la clau primaria de la capçalera no es Text1(0), canviar-ho en <=== *****
    If Modo <> 1 Then
        LimpiarCampos
        PonerModo 1
        PonerFoco Text1(0) ' <===
        Text1(0).BackColor = vbYellow ' <===
        ' *** si n'hi han combos a la capçalera ***
    Else
        HacerBusqueda
        If Data1.Recordset.EOF Then
            Text1(kCampo).Text = ""
            Text1(kCampo).BackColor = vbYellow
            PonerFoco Text1(kCampo)
        End If
    End If
' ******************************************************************************
End Sub

Private Sub HacerBusqueda()

    CadB = ObtenerBusqueda2(Me, BuscaChekc, 1)
    If CadB <> "" Then
        CadB = CadB & " and " & CadB1
    Else
        CadB = CadB & CadB1
    End If
    
    If chkVistaPrevia(0) = 1 Then
        MandaBusquedaPrevia CadB
    ElseIf CadB <> "" Then
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " " & Ordenacion
        PonerCadenaBusqueda
    Else
        ' *** foco al 1r camp visible de la capçalera que siga clau primaria ***
        PonerFoco Text1(0)
        ' **********************************************************************
    End If
End Sub


Private Sub MandaBusquedaPrevia(CadB As String)
    Dim Cad As String
    Dim NombreTabla1 As String
        
    'Cridem al form
    ' **************** arreglar-ho per a vore lo que es desije ****************
    ' NOTA: el total d'amples de ParaGrid, ha de sumar 100
    Cad = ""
    Cad = Cad & "Código|rprecios.codvarie|N|000000|9·"
    Cad = Cad & "Variedad|variedades.nomvarie|T||40·"
    Cad = Cad & "Tipo|rprecios.tipofact|N|000000|9·"
    Cad = Cad & "Contador|rprecios.contador|N|00000|9·"
    Cad = Cad & "Fecha Inicio|rprecios.fechaini|F||15·"
    Cad = Cad & "Fecha Fin|rprecios.fechafin|F||15·"
    
    
    NombreTabla1 = "(rprecios inner join variedades on rprecios.codvarie = variedades.codvarie)"
    
    If Cad <> "" Then
        Screen.MousePointer = vbHourglass
        Set frmB = New frmBuscaGrid
        frmB.vCampos = Cad
        frmB.vTabla = NombreTabla1
        
        frmB.vSQL = CadB
        
        HaDevueltoDatos = False
        frmB.vDevuelve = "0|2|3|" '*** els camps que volen que torne ***
        frmB.vTitulo = "Precios" ' ***** repasa açò: títol de BuscaGrid *****
        frmB.vSelElem = 0

        frmB.Show vbModal
        Set frmB = Nothing
        'Si ha posat valors i tenim que es formulari de búsqueda llavors
        'tindrem que tancar el form llançant l'event
        If HaDevueltoDatos Then
            If (Not Data1.Recordset.EOF) And DatosADevolverBusqueda <> "" Then _
                cmdRegresar_Click
        Else   'de ha retornat datos, es a decir NO ha retornat datos
            PonerFoco Text1(kCampo)
        End If
    End If
End Sub


Private Sub cmdRegresar_Click()
Dim Cad As String
Dim Aux As String
Dim i As Integer
Dim j As Integer

    If Data1.Recordset.EOF Then
        MsgBox "Ningún registro devuelto.", vbExclamation
        Exit Sub
    End If
    
    Cad = ""
    i = 0
    Do
        j = i + 1
        i = InStr(j, DatosADevolverBusqueda, "|")
        If i > 0 Then
            Aux = Mid(DatosADevolverBusqueda, j, i - j)
            j = Val(Aux)
            Cad = Cad & Text1(j).Text & "|"
        End If
    Loop Until i = 0
    RaiseEvent DatoSeleccionado(Cad)
    Unload Me
End Sub


Private Sub PonerCadenaBusqueda()
    Screen.MousePointer = vbHourglass
    On Error GoTo EEPonerBusq
    
    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    
    If Data1.Recordset.RecordCount <= 0 Then
        MsgBox "No hay ningún registro en la tabla " & NombreTabla, vbInformation
        Screen.MousePointer = vbDefault
        Exit Sub
    Else
        PonerModo 2
        'Data1.Recordset.MoveLast
        Data1.Recordset.MoveFirst
        PonerCampos
    End If
    
    Screen.MousePointer = vbDefault
    Exit Sub

EEPonerBusq:
    MuestraError Err.Number, "PonerCadenaBusqueda"
    PonerModo 0
    Screen.MousePointer = vbDefault
End Sub


Private Sub BotonVerTodos()
'Vore tots
    LimpiarCampos 'Neteja els Text1
    CadB = ""
    
    PonerModo 0
    
    If chkVistaPrevia(0).Value = 1 Then
        MandaBusquedaPrevia CadB1
    Else
        ' *** canviar o llevar, si cal, el WHERE; repasar codEmpre ***
        CadenaConsulta = "Select * from " & NombreTabla & " where " & CadB1 & Ordenacion
        'CadenaConsulta = "Select * from " & NombreTabla & " where codempre = " & codEmpre & Ordenacion
        ' ******************************************
        PonerCadenaBusqueda
        ' *** si n'hi han llínies sense grids ***
'        CargaFrame 0, True
        ' ************************************
    End If
End Sub


Private Sub BotonAnyadir()

    LimpiarCampos 'Huida els TextBox
    PonerModo 3
    
    ' ****** Valors per defecte a l'afegir, repasar si n'hi ha
    ' codEmpre i quins camps tenen la PK de la capçalera *******
'    Text1(0).Text = SugerirCodigoSiguienteStr("rcampos", "codcampo")
    FormateaCampo Text1(0)
       
    PonerFoco Text1(0) '*** 1r camp visible que siga PK ***
    ' ***********************************************************
    
    ' *** si n'hi han camps de descripció a la capçalera ***
    PosarDescripcions
    ' ******************************************************

    ' *** si n'hi han tabs, em posicione al 1r ***
'    Me.SSTab1.Tab = 0
    ' ********************************************
End Sub


Private Sub BotonModificar()

    PonerModo 4

    ' *** bloquejar els camps visibles de la clau primaria de la capçalera ***
    BloquearTxt Text1(0), True
    ' *************************************************************************
    
    VarieAnt = Text1(2).Text
    ' *** foco al 1r camp visible que NO siga clau primaria ***
    PonerFoco Text1(1)
    ' *********************************************************
End Sub


Private Sub BotonEliminar()
Dim Cad As String

    On Error GoTo EEliminar

    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub

    ' *** repasar el nom de l'adodc, l'index del Field i el camp que te la PK ***
    'El registre de codi 0 no es pot Modificar ni Eliminar
    If EsCodigoCero(CStr(Data1.Recordset.Fields(0).Value), FormatoCampo(Text1(0))) Then Exit Sub
    ' ***************************************************************************

    ' *************** canviar la pregunta ****************
    Cad = "¿Seguro que desea eliminar el Precio de la Variedad?"
    Cad = Cad & vbCrLf & "Variedad: " & Format(Data1.Recordset.Fields(0), FormatoCampo(Text1(0)))
    Cad = Cad & vbCrLf & "Tipo    : "
    Select Case Data1.Recordset.Fields(1)
        Case 0
            Cad = Cad & "Anticipo"
        Case 1
            Cad = Cad & "Liquidacion"
        Case 2
            Cad = Cad & "Industria directa"
        Case 3
            Cad = Cad & "Complementaria"
        Case 4
            Cad = Cad & "Anticipo Genérico"
        Case 5
            Cad = Cad & "Anticipo Retirada"
        Case 6
            Cad = Cad & "Complementaria única"
    End Select
    Cad = Cad & vbCrLf & "Contador: " & Data1.Recordset.Fields(2)
    ' **************************************************************************
    
    'borrem
    If MsgBox(Cad, vbQuestion + vbYesNo) = vbYes Then
        On Error GoTo EEliminar
        Screen.MousePointer = vbHourglass
        NumRegElim = Data1.Recordset.AbsolutePosition
        If Not eliminar Then
            Screen.MousePointer = vbDefault
            Exit Sub
        ' **** repasar si es diu Data1 l'adodc de la capçalera ***
        ElseIf SituarDataTrasEliminar(Data1, NumRegElim) Then
        ' ********************************************************
            PonerCampos
        Else
            LimpiarCampos
            PonerModo 0
        End If
    End If
    Screen.MousePointer = vbDefault
    
EEliminar:
    Screen.MousePointer = vbDefault
    If Err.Number > 0 Then MuestraError Err.Number, "Eliminar Cliente", Err.Description
End Sub


Private Sub PonerCampos()
Dim i As Integer
Dim CodPobla As String, desPobla As String
Dim CPostal As String, desProvi As String, desPais As String

    If Data1.Recordset.EOF Then Exit Sub
    PonerCamposForma2 Me, Data1, 1 'opcio=1: posa el format o els camps de la capçalera
    
    ' *** si n'hi han llínies en datagrids ***
    For i = 1 To DataGridAux.Count ' - 1
        If i <> 3 Then
            CargaGrid i, True
            If Not AdoAux(i).Recordset.EOF Then _
                PonerCamposForma2 Me, AdoAux(i), 2, "FrameAux" & i
        End If
    Next i
    ' *******************************************

    ' *** si n'hi han llínies sense datagrid ***
    ' ***************************************
    
    ' ************* configurar els camps de les descripcions de la capçalera *************
'    txtAux2(22).Text = PonerNombreDeCod(txtAux(22), "poblacio", "despobla", "codpobla", "N")

    PosarDescripcions

    PonerClasificacionGrafica

    VisualizaClasificacion


    ' ********************************************************************************
    
    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
    
    PonerModoOpcionesMenu (Modo)
    PonerOpcionesMenu   'Activar opcions de menú según nivell
                        'de permisos de l'usuari
    
End Sub


Private Sub cmdCancelar_Click()
Dim i As Integer
Dim V

    Select Case Modo
        Case 1, 3 'Búsqueda, Insertar
                LimpiarCampos
                If Data1.Recordset.EOF Then
                    PonerModo 0
                Else
                    PonerModo 2
                    PonerCampos
                End If
                ' *** foco al primer camp visible de la capçalera ***
                PonerFoco Text1(0)
                ' ***************************************************

        Case 4  'Modificar
                TerminaBloquear
                PonerModo 2
                PonerCampos
                ' *** primer camp visible de la capçalera ***
                PonerFoco Text1(0)
                ' *******************************************
                
                
        Case 5 'LLÍNIES
            Select Case ModoLineas
                Case 1 'afegir llínia
                    ModoLineas = 0
                    ' *** les llínies que tenen datagrid (en o sense tab) ***
                    If NumTabMto = 0 Or NumTabMto = 1 Or NumTabMto = 2 Or NumTabMto = 4 Then
                        DataGridAux(NumTabMto).AllowAddNew = False
                        ' **** repasar si es diu Data1 l'adodc de la capçalera ***
                        'If BLOQUEADesdeFormulario2(Me, Data1, 1) Then BotonModificar 'Modificar
                        ' ********************************************************
                        LLamaLineas NumTabMto, ModoLineas 'ocultar txtAux
                        DataGridAux(NumTabMto).Enabled = True
                        DataGridAux(NumTabMto).SetFocus

                        ' *** si n'hi han camps de descripció dins del grid, els neteje ***
                        'txtAux2(2).text = ""
                        ' *****************************************************************

                        ' ***  bloquejar i huidar els camps que estan fora del datagrid ***
                        Select Case NumTabMto
                            Case 0 'cuentas bancarias
                                'BotonModificar
'                                BloquearTxt txtaux(11), True
'                                BloquearTxt txtaux(12), True
                            Case 1 'secciones
                                For i = 0 To txtAux1.Count - 1
                                    txtAux1(i).Text = ""
                                    BloquearTxt txtAux1(i), True
                                Next i
                                txtAux2(1).Text = ""
                                BloquearTxt txtAux2(1), True
'                            Case 2 'telefonos
'                                For I = 0 To txtAux.Count
'                                    BloquearTxt txtAux(I), True
'                                Next I
                        End Select
                    ' *** els tabs que no tenen datagrid ***
                    ElseIf NumTabMto = 3 Then
                        If BLOQUEADesdeFormulario2(Me, Data1, 1) Then BotonModificar 'Modificar
                        CargaFrame 3, True
                    End If

                    ' *** si n'hi han tabs ***
                    SituarTab (NumTabMto)
                    'SSTab1.Tab = 1
                    'SSTab2.Tab = NumTabMto
                    ' ************************

                    If Not AdoAux(NumTabMto).Recordset.EOF Then
                        AdoAux(NumTabMto).Recordset.MoveFirst
                    End If

                Case 2 'modificar llínies
                    ModoLineas = 0

                    ' *** si n'hi han tabs ***
                    SituarTab (NumTabMto)
                    'SSTab1.Tab = 1
                    'SSTab2.Tab = NumTabMto
                    ' ***********************

                    PonerModo 4
                    If Not AdoAux(NumTabMto).Recordset.EOF Then
                        ' *** l'Index de Fields es el que canvie de la PK de llínies ***
                        V = AdoAux(NumTabMto).Recordset.Fields(1) 'el 2 es el nº de llinia
                        AdoAux(NumTabMto).Recordset.Find (AdoAux(NumTabMto).Recordset.Fields(1).Name & " =" & V)
                        ' ***************************************************************
                    End If

                    ' ***  bloquejar els camps fora dels grids ***

                    ' ***  bloquejar els camps fora dels grids ***
                    LLamaLineas NumTabMto, ModoLineas 'ocultar txtAux
            End Select

'            If NumTabMto = 1 Then
'                If Not vSeccion Is Nothing Then
'                    vSeccion.CerrarConta
'                    Set vSeccion = Nothing
'                End If
'            End If
            
            SumaTotalPorcentajes

            PosicionarData
            
            TerminaBloquear

            ' *** si n'hi han llínies en grids i camps fora d'estos ***
            If Not AdoAux(NumTabMto).Recordset.EOF Then
                DataGridAux_RowColChange NumTabMto, 1, 1
            Else
                LimpiarCamposFrame NumTabMto
            End If
            ' *********************************************************
    End Select
End Sub


Private Function DatosOk() As Boolean
Dim b As Boolean
Dim Sql As String
Dim Cad As String
'Dim Datos As String

    On Error GoTo EDatosOK

    DatosOk = False
    b = CompForm2(Me, 1)
    If Not b Then Exit Function
    
    ' *** canviar els arguments de la funcio, el mensage i repasar si n'hi ha codEmpre ***
    If (Modo = 3) Then 'insertar
        'comprobar si existe ya el cod. del campo clave primaria
        Sql = DevuelveDesdeBDNew(cAgro, "rprecios", "codvarie", "codvarie", Text1(0).Text, "N", , "tipofact", Combo1(0).ListIndex, "N", "contador", Text1(1).Text, "N")
        If Sql <> "" Then
            MsgBox "Ya existe el contador para este variedad. Revise.", vbExclamation
            b = False
        End If

'--monica(27/07/2009): debemos permitir introducir para un mismo periodo la misma variedad
'        ' controlamos que no me inserten para la misma variedad un mismo periodo de fechas
'        If b Then
'            Sql = "select count(*) from rprecios where codvarie = " & DBSet(Text1(0).Text, "N")
'            Sql = Sql & " and tipofact = " & DBSet(Combo1(0).ListIndex, "N")
'            Sql = Sql & " and fechaini = " & DBSet(Text1(2).Text, "F")
'            Sql = Sql & " and fechafin = " & DBSet(Text1(3).Text, "F")
'
'            If TotalRegistros(Sql) > 0 Then
'                MsgBox "Ya existe un registro para esta variedad rango de fechas. Revise.", vbExclamation
'                b = False
'            End If
'        End If
    End If
    
    '[Monica] 24/09/2009 Si es catadau controlamos que no se puedan solapar fechas
    ' para el caso de Valsur dejamos que se solapen pq obligo a introducir el rango exacto y cojo el maximo contador
    '[Monica]29/02/2012: Natural era la cooperativa 0 junto con Catadau ahora es la 9
    If vParamAplic.Cooperativa = 0 Or vParamAplic.Cooperativa = 7 Or vParamAplic.Cooperativa = 9 Or vParamAplic.Cooperativa = 5 Or vParamAplic.Cooperativa = 2 Or vParamAplic.Cooperativa = 14 Or vParamAplic.Cooperativa = 16 Then ' caso de Catadau y de quatretonda y de Castelduc
        If b And (Modo = 3 Or Modo = 4) And Combo1(0).ListIndex <= 1 Then ' solo anticipos y liquidacion
                                                                          ' para el caso de complementaria obligo a meter el rango exacto
'            If Modo = 4 Then And (Text1(2).Text <> Data1.Recordset!FechaIni Or Text1(3).Text <> Data1.Recordset!FechaFin) Then
                                                                          
                b = ComprobacionRangoFechas(Text1(0).Text, CStr(Combo1(0).ListIndex), Text1(1).Text, Text1(2).Text, Text1(3).Text)
        
                If b = False Then
                    MsgBox "El rango de fechas se solapa con otro registro del mismo tipo de esta variedad. Revise.", vbExclamation
                End If
'            End If
        End If
    End If
    
    ' ************************************************************************************
    DatosOk = b
    
EDatosOK:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function


Private Sub PosicionarData()
Dim Cad As String, Indicador As String

    ' *** canviar-ho per tota la PK de la capçalera, no llevar els () ***
    Cad = "(codvarie=" & Text1(0).Text & " and tipofact = " & Combo1(0).ListIndex & " and contador = " & Text1(1).Text & " )"
    ' ***************************************
    
    ' *** gastar SituarData o SituarDataMULTI depenent de si la PK es simple o composta ***
    'If SituarDataMULTI(Data1, cad, Indicador) Then
    If SituarDataMULTI(Data1, Cad, Indicador) Then
        If ModoLineas <> 1 Then PonerModo 2
        lblIndicador.Caption = Indicador
    Else
       LimpiarCampos
       PonerModo 0
    End If
    ' ***********************************************************************************
End Sub

Private Function eliminar() As Boolean
Dim vWhere As String

    On Error GoTo FinEliminar

    conn.BeginTrans
    ' ***** canviar el nom de la PK de la capçalera, repasar codEmpre *******
    vWhere = " WHERE codvarie=" & Data1.Recordset!codvarie
    vWhere = vWhere & " and tipofact = " & Data1.Recordset!TipoFact
    vWhere = vWhere & " and contador = " & Data1.Recordset!Contador
        ' ***********************************************************************
        
    ' ***** elimina les llínies ****
    conn.Execute "DELETE FROM rprecios_calidad " & vWhere

'    ' *******************************
'    'Eliminar la CAPÇALERA
    conn.Execute "Delete from " & NombreTabla & vWhere
       
FinEliminar:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Eliminar"
        conn.RollbackTrans
        eliminar = False
    Else
        conn.CommitTrans
        eliminar = True
    End If
End Function


Private Sub Text1_GotFocus(Index As Integer)
    kCampo = Index
    ConseguirFoco Text1(Index), Modo
    
    If Index = 5 Then
        If Tipo = 0 Then
            '[Monica]27/01/2016: incluido tambien el tipo complementaria unica
            Text1(Index).Enabled = (Combo1(0).ListIndex = 2) Or (Combo1(0).ListIndex = 4) Or (Combo1(0).ListIndex = 5) Or (Combo1(0).ListIndex = 6)
        Else
            Text1(Index).Enabled = True
        End If
    End If
    
End Sub

Private Sub Text1_LostFocus(Index As Integer)
Dim cadMen As String
Dim Nuevo As Boolean

    If Not PerderFocoGnral(Text1(Index), Modo) Then Exit Sub

    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    Select Case Index
        Case 0 'Variedad
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), "variedades", "nomvarie")
                If Text2(Index).Text = "" Then
                    cadMen = "No existe la Variedad: " & Text1(Index).Text & vbCrLf
                    cadMen = cadMen & "¿Desea crearla?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmVar = New frmComVar
                        frmVar.DatosADevolverBusqueda = "0|1|"
'                        frmVar.NuevoCodigo = Text1(Index).Text
                        Text1(Index).Text = ""
                        TerminaBloquear
                        frmVar.Show vbModal
                        Set frmVar = Nothing
                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
                    Else
                        Text1(Index).Text = ""
                    End If
                    PonerFoco Text1(Index)
                Else
                    Select Case Tipo
                        Case 0  'cualquier tipo
                            '[Monica]30/11/2011: en quatretonda se anticipa y liquida almazara como si fuera horto
                            If vParamAplic.Cooperativa = 7 Then
                                If EsVariedadGrupo6(Text1(Index)) Then
                                    MsgBox "Esta variedad es del grupo de bodega. Reintroduzca.", vbExclamation
                                    PonerFoco Text1(Index)
                                End If
                            Else
                                If EsVariedadGrupo5(Text1(Index)) Then
                                    MsgBox "Esta variedad es del grupo de almazara. Reintroduzca.", vbExclamation
                                    PonerFoco Text1(Index)
                                Else
                                    If EsVariedadGrupo6(Text1(Index)) Then
                                        MsgBox "Esta variedad es del grupo de bodega. Reintroduzca.", vbExclamation
                                        PonerFoco Text1(Index)
                                    End If
                                End If
                            End If
                        Case 1  'almazara
                            If Not EsVariedadGrupo5(Text1(Index)) Then
                                MsgBox "Esta variedad no es del grupo de almazara. Reintroduzca.", vbExclamation
                                PonerFoco Text1(Index)
                            End If
                        Case 2  'bodega
                            If Not EsVariedadGrupo6(Text1(Index)) Then
                                MsgBox "Esta variedad no es del grupo de bodega. Reintroduzca.", vbExclamation
                                PonerFoco Text1(Index)
                            End If
                    End Select
                End If
            Else
                Text2(Index).Text = ""
            End If
        
        Case 2, 3 ' fechas de inicio y fin
            '[Monica]28/08/2013: comprobamos que la fecha esté en la campaña, antes no se comprobaba
            If PonerFormatoFecha(Text1(Index), True) Then
                If Text1(2).Text <> "" And Text1(3).Text <> "" Then
                    If CDate(Text1(2).Text) > CDate(Text1(3).Text) Then
                        MsgBox "La Fecha Inicio debe ser inferior a la Fecha Fin. Revise", vbExclamation
                        Text1(Index).Text = ""
                        PonerFoco Text1(Index)
                    End If
                End If
            End If
        Case 1 'contador
            PonerFormatoEntero Text1(Index)
                
        Case 4 'texto
            If Text1(Index).Text <> "" Then Text1(Index).Text = UCase(Text1(Index).Text)
            
        Case 5 'precio de industria si se trata de industria directa
            PonerFormatoDecimal Text1(Index), 8
            
    End Select
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = teclaBuscar Then
        If Modo = 1 Or Modo = 3 Or Modo = 4 Then
            Select Case Index
                Case 0: KEYBusqueda KeyAscii, 0 'variedad
            End Select
        End If
    Else
        If Index <> 21 Then KEYpress KeyAscii
    End If
End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Alvançar/Retrocedir els camps en les fleches de desplaçament del teclat.
    KEYdown KeyCode
End Sub

Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 0, cerrar
    If cerrar Then Unload Me
End Sub


Private Sub KEYBusqueda(KeyAscii As Integer, indice As Integer)
    KeyAscii = 0
    imgBuscar_Click (indice)
End Sub

' **** si n'hi han camps de descripció a la capçalera ****
Private Sub PosarDescripcions()
Dim NomEmple As String

    On Error GoTo EPosarDescripcions

    Text2(0).Text = PonerNombreDeCod(Text1(0), "variedades", "nomvarie", "codvarie", "N")
    
'    PonerDatosPartida
'
'    If Text1(5).Text <> "" Then
'        Text2(5).Text = Round2(ImporteSinFormato(Text1(5).Text) / cFaneca, 4)
'        PonerFormatoDecimal Text2(5), 7
'    End If
'
'    If Text1(6).Text <> "" Then
'        Text2(6).Text = Round2(ImporteSinFormato(Text1(6).Text) / cFaneca, 4)
'        PonerFormatoDecimal Text2(6), 7
'    End If
'
'    If Text1(7).Text <> "" Then
'        Text2(7).Text = Round2(ImporteSinFormato(Text1(7).Text) / cFaneca, 4)
'        PonerFormatoDecimal Text2(7), 7
'    End If
    
EPosarDescripcions:
    If Err.Number <> 0 Then MuestraError Err.Number, "Poniendo descripciones", Err.Description
End Sub
' ************************************************************


'************* LLINIES: ****************************
Private Sub ToolAux_ButtonClick(Index As Integer, ByVal Button As MSComctlLib.Button)
'-- pon el bloqueo aqui
    'If BLOQUEADesdeFormulario2(Me, Data1, 1) Then
    Select Case Button.Index
        Case 1
            BotonAnyadirLinea Index
        Case 2
            BotonModificarLinea Index
        Case 3
            BotonEliminarLinea Index
        Case Else
    End Select
    'End If
End Sub


Private Sub BotonEliminarLinea(Index As Integer)
Dim Sql As String
Dim vWhere As String
Dim eliminar As Boolean

    On Error GoTo Error2

    ModoLineas = 3 'Posem Modo Eliminar Llínia
    
    If Modo = 4 Then 'Modificar Capçalera
        cmdAceptar_Click
        If ModoLineas = 0 Then Exit Sub
    End If
       
    NumTabMto = Index
    PonerModo 5, Index

    If AdoAux(Index).Recordset.EOF Then Exit Sub
    If Not SepuedeBorrar(Index) Then Exit Sub
    NumTabMto = Index
    eliminar = False
   
    vWhere = ObtenerWhereCab(True)
    
    ' ***** independentment de si tenen datagrid o no,
    ' canviar els noms, els formats i el DELETE *****
    Select Case Index
        Case 1 'calidad
            Sql = "¿Seguro que desea eliminar la calidad?"
            Sql = Sql & vbCrLf & "Calidad: " & AdoAux(Index).Recordset!codcalid & " - " & AdoAux(Index).Recordset!nomcalid
            If MsgBox(Sql, vbQuestion + vbYesNo) = vbYes Then
                eliminar = True
                Sql = "DELETE FROM rprecios_calidad "
                Sql = Sql & vWhere & " and codcalid = " & DBLet(AdoAux(Index).Recordset!codcalid, "N")
            End If
        
    End Select

    If eliminar Then
        NumRegElim = AdoAux(Index).Recordset.AbsolutePosition
        TerminaBloquear
        conn.Execute Sql
        ' *** si n'hi han tabs sense datagrid, posar l'If ***
        If Index <> 3 Then _
            CargaGrid Index, True
        ' ***************************************************
        If Not SituarDataTrasEliminar(AdoAux(Index), NumRegElim, True) Then
'            PonerCampos
            
        End If
        ' *** si n'hi han tabs sense datagrid ***
        If Index = 3 Then CargaFrame 3, True
        ' ***************************************
        If BLOQUEADesdeFormulario2(Me, Data1, 1) Then BotonModificar
        ' *** si n'hi han tabs ***
        SituarTab (NumTabMto)
        ' ************************
    End If
    
    ModoLineas = 0
    PosicionarData
    
    Exit Sub
Error2:
    Screen.MousePointer = vbDefault
    MuestraError Err.Number, "Eliminando linea", Err.Description
End Sub


Private Sub BotonAnyadirLinea(Index As Integer)
Dim NumF As String
Dim vWhere As String, vTabla As String
Dim anc As Single
Dim i As Integer

    ModoLineas = 1 'Posem Modo Afegir Llínia

    If (Modo = 3) Or (Modo = 4) Then 'Insertar o Modificar Capçalera
        cmdAceptar_Click
        If ModoLineas = 0 Then Exit Sub
    End If

    NumTabMto = Index
    PonerModo 5, Index

    ' *** bloquejar la clau primaria de la capçalera ***
    BloquearTxt Text1(0), True
    ' **************************************************

    ' *** posar el nom del les distintes taules de llínies ***
    Select Case Index
        Case 1: vTabla = "rprecios_calidad"
    End Select
    ' ********************************************************

    vWhere = ObtenerWhereCab(False)

    Select Case Index
         Case 1   'clasificacion
            ' *** canviar la clau primaria de les llínies,
            'pasar a "" si no volem que mos sugerixca res a l'afegir ***
'            If Index <> 4 Then ' *** els index als que no volem sugerir-li un codi ***
'                NumF = SugerirCodigoSiguienteStr(vTabla, "codsecci", vWhere)
'            Else
'                NumF = ""
'            End If
            ' ***************************************************************

            AnyadirLinea DataGridAux(Index), AdoAux(Index)

            anc = DataGridAux(Index).Top
            If DataGridAux(Index).Row < 0 Then
                anc = anc + 210
            Else
                anc = anc + DataGridAux(Index).RowTop(DataGridAux(Index).Row) + 5
            End If

            LLamaLineas Index, ModoLineas, anc

            Select Case Index
                ' *** valor per defecte a l'insertar i formateig de tots els camps ***
                Case 1 'calidades
                    For i = 0 To txtAux1.Count - 1
                        txtAux1(i).Text = ""
                    Next i
                    txtAux1(0).Text = Text1(0).Text 'codvariedad
                    txtAux1(2).Text = Combo1(0).ListIndex  'tipo de factura
                    txtAux1(4).Text = Text1(1).Text 'contador
                    
                    txtAux1(1).Text = "" 'calidad
                    txtAux2(1).Text = ""
                    PonerFoco txtAux1(1)

            End Select


'        ' *** si n'hi han llínies sense datagrid ***
'        Case 3
'            LimpiarCamposLin "FrameAux3"
'            txtaux(42).Text = text1(0).Text 'codclien
'            txtaux(43).Text = vSesion.Empresa
'            Me.cmbAux(28).ListIndex = 0
'            Me.cmbAux(29).ListIndex = 1
'            PonerFoco txtaux(25)
'        ' ******************************************
    End Select
End Sub


Private Sub BotonModificarLinea(Index As Integer)
    Dim anc As Single
    Dim i As Integer
    Dim j As Integer

    If AdoAux(Index).Recordset.EOF Then Exit Sub
    If AdoAux(Index).Recordset.RecordCount < 1 Then Exit Sub

    ModoLineas = 2 'Modificar llínia

    If Modo = 4 Then 'Modificar Capçalera
        cmdAceptar_Click
        If ModoLineas = 0 Then Exit Sub
    End If

    NumTabMto = Index
    PonerModo 5, Index
    ' *** bloqueje la clau primaria de la capçalera ***
    BloquearTxt Text1(0), True
    ' *********************************

    Select Case Index
        Case 0, 1 ' *** pose els index de llínies que tenen datagrid (en o sense tab) ***
            If DataGridAux(Index).Bookmark < DataGridAux(Index).FirstRow Or DataGridAux(Index).Bookmark > (DataGridAux(Index).FirstRow + DataGridAux(Index).VisibleRows - 1) Then
                i = DataGridAux(Index).Bookmark - DataGridAux(Index).FirstRow
                DataGridAux(Index).Scroll 0, i
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
        Case 1 'calidades
            txtAux1(0).Text = DataGridAux(Index).Columns(0).Text
            txtAux1(4).Text = DataGridAux(Index).Columns(2).Text
            txtAux1(2).Text = DataGridAux(Index).Columns(1).Text
            
            txtAux1(1).Text = DataGridAux(Index).Columns(3).Text 'calidad
            txtAux2(1).Text = DataGridAux(Index).Columns(4).Text ' nombre calidad
            txtAux1(3).Text = DataGridAux(Index).Columns(5).Text 'precio cooperativa
            txtAux1(5).Text = DataGridAux(Index).Columns(6).Text 'precio socio
            
    End Select

    LLamaLineas Index, ModoLineas, anc

    ' *** foco al 1r camp visible de les llinies en grids que no siga PK (en o sense tab) ***
    Select Case Index
        Case 1 'calidades
            PonerFoco txtAux1(3)
    End Select
    ' ***************************************************************************************
End Sub


Private Sub LLamaLineas(Index As Integer, xModo As Byte, Optional alto As Single)
Dim jj As Integer
Dim b As Boolean

    ' *** si n'hi han tabs sense datagrid posar el If ***
    If Index <> 3 Then DeseleccionaGrid DataGridAux(Index)
    ' ***************************************************

    b = (xModo = 1 Or xModo = 2) 'Insertar o Modificar Llínies
    Select Case Index
        Case 1 'calidad
            For jj = 1 To txtAux1.Count - 1
                If jj <> 0 And jj <> 2 And jj <> 4 Then
                    txtAux1(jj).visible = b
                    txtAux1(jj).Top = alto
                End If
            Next jj
            
            txtAux2(1).visible = b
            txtAux2(1).Top = alto

            For jj = 1 To cmdAux.Count
                cmdAux(jj).visible = b
                cmdAux(jj).Top = txtAux1(3).Top
                cmdAux(jj).Height = txtAux1(3).Height
            Next jj
    End Select
End Sub



Private Function DatosOkLlin(nomframe As String) As Boolean
Dim Rs As ADODB.Recordset
Dim Sql As String
Dim b As Boolean
Dim Cant As Integer
Dim Mens As String
Dim vFact As Byte, vDocum As Byte

    DatosOkLlin = True
    
    On Error GoTo EDatosOKLlin

    Mens = ""
    DatosOkLlin = False

    b = CompForm2(Me, 2, nomframe) 'Comprovar formato datos ok
    If Not b Then Exit Function
    
    If b And (Modo = 5 And ModoLineas = 1) Then  'insertar
        'comprobar si existe ya el cod. de la calidad para ese campo
        Sql = ""
'        SQL = DevuelveDesdeBDNew(cAgro, "rprecios_calidad", "codcalid", "codvarie", txtaux1(0).Text, "N", , "tipofact", txtaux1(2).Text, "N", "codcalid", txtaux1(1).Text, "N")
        If Sql <> "" Then
            MsgBox "Ya existe la calidad. Revise.", vbExclamation
            PonerFoco txtAux1(1)
            b = False
        End If
    End If
    
'
'    ' *** si cal fer atres comprovacions a les llínies (en o sense tab) ***
'    Select Case NumTabMto
'        Case 0  'CUENTAS BANCARIAS
'            SQL = "SELECT COUNT(ctaprpal) FROM cltebanc "
'            SQL = SQL & ObtenerWhereCab(True) & " AND ctaprpal=1"
'            If ModoLineas = 2 Then SQL = SQL & " AND numlinea<> " & AdoAux(NumTabMto).Recordset!numlinea
'            Set RS = New ADODB.Recordset
'            RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'            Cant = IIf(Not RS.EOF, RS.Fields(0).Value, 0)
'
'            RS.Close
'            Set RS = Nothing
''yo
''            'no n'hi ha cap conter principal i ha seleccionat que no
''            If (Cant = 0) And (cmbAux(10).ItemData(cmbAux(10).ListIndex) = 0) Then
''                Mens = "Debe una haber una cuenta principal"
''            ElseIf (cmbAux(10).ItemData(cmbAux(10).ListIndex) = 1) And (cmbAux(9).ItemData(cmbAux(9).ListIndex) = 0) Then
''                Mens = "Debe seleccionar que esta cuenta está activa si desea que sea la principal"
''            End If
'
''            'No puede haber más de una cuenta principal
''            If cant > 0 And (cmbAux(10).ItemData(cmbAux(10).ListIndex) = 1) Then
''                Mens = "No puede haber más de una cuenta principal."
''            End If
''yo
''            'No pueden haber registros con el mismo: codbanco-codsucur-digcontr-ctabanc
''            If Mens = "" Then
''                SQL = "SELECT count(codclien) FROM cltebanc "
''                SQL = SQL & " WHERE codclien=" & text1(0).Text & " AND codempre= " & vSesion.Empresa
''                If ModoLineas = 2 Then SQL = SQL & " AND numlinea<> " & AdoAux(NumTabMto).Recordset!numlinea
''                SQL = SQL & " AND codnacio=" & cmbAux(0).ItemData(cmbAux(0).ListIndex)
''                SQL = SQL & " AND codbanco=" & txtaux(3).Text & " AND codsucur=" & txtaux(4).Text
''                SQL = SQL & " AND digcontr='" & txtaux(5).Text & "' AND ctabanco='" & txtaux(6).Text & "'"
''                Set RS = New ADODB.Recordset
''                RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
''                Cant = IIf(Not RS.EOF, RS.Fields(0).Value, 0)
''                If Cant > 0 Then
''                    Mens = "Ya Existe la cuenta bancaria: " & cmbAux(0).List(cmbAux(0).ListIndex) & " - " & txtaux(3).Text & "-" & txtaux(4).Text & "-" & txtaux(5).Text & "-" & txtaux(6).Text
''                End If
''                RS.Close
''                Set RS = Nothing
''            End If
''
''            If Mens <> "" Then
''                Screen.MousePointer = vbNormal
''                MsgBox Mens, vbExclamation
''                DatosOkLlin = False
''                'PonerFoco txtAux(3)
''                Exit Function
''            End If
''
'    End Select
'    ' ******************************************************************************
    DatosOkLlin = b

EDatosOKLlin:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function


Private Function SepuedeBorrar(ByRef Index As Integer) As Boolean
    SepuedeBorrar = False
    
'    ' *** si cal comprovar alguna cosa abans de borrar ***
'    Select Case Index
'        Case 0 'cuentas bancarias
'            If AdoAux(Index).Recordset!ctaprpal = 1 Then
'                MsgBox "No puede borrar una Cuenta Principal. Seleccione antes otra cuenta como Principal", vbExclamation
'                Exit Function
'            End If
'    End Select
'    ' ****************************************************
    
    SepuedeBorrar = True
End Function

' *** si n'hi han formularis de buscar codi a les llínies ***
Private Sub imgBuscar_Click(Index As Integer)
    TerminaBloquear
    
    Select Case Index
        Case 0 'situacion
            Set frmVar = New frmComVar
'            frmVar.DeConsulta = True
            frmVar.DatosADevolverBusqueda = "0|1|"
'            frmVar.CodigoActual = Text1(2).Text
            frmVar.Show vbModal
            Set frmVar = Nothing
            PonerFoco Text1(2)
        
    End Select
    
    If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
End Sub


' *********************************************************************************
Private Sub DataGridAux_RowColChange(Index As Integer, LastRow As Variant, ByVal LastCol As Integer)
'Dim I As Byte
'
'    If ModoLineas <> 1 Then
'        Select Case Index
'            Case 0 'telefonos
'                If DataGridAux(Index).Columns.Count > 2 Then
'                    For I = 5 To txtAux.Count - 1
'                        txtAux(I).Text = DataGridAux(Index).Columns(I).Text
'                    Next I
'                    Me.chkAbonos(1).Value = DataGridAux(Index).Columns(17).Text
'
'                End If
'            Case 1 'secciones
'                If DataGridAux(Index).Columns.Count > 2 Then
'                    txtAux2(4).Text = ""
'                    txtAux2(5).Text = ""
'                    txtAux2(0).Text = ""
'                    Set vSeccion = New CSeccion
'                    If vSeccion.LeerDatos(AdoAux(1).Recordset!codsecci) Then
'                        If vSeccion.AbrirConta Then
'                            If DBLet(AdoAux(1).Recordset!codmaccli, "T") <> "" Then
'                                txtAux2(4).Text = DevuelveDesdeBDNew(cConta, "cuentas", "nommacta", "codmacta", AdoAux(1).Recordset!codmaccli, "T")
'                            End If
'                            If DBLet(AdoAux(1).Recordset!codmacpro, "T") <> "" Then
'                                txtAux2(5).Text = DevuelveDesdeBDNew(cConta, "cuentas", "nommacta", "codmacta", AdoAux(1).Recordset!codmacpro, "T")
'                            End If
'                            txtAux2(0).Text = DevuelveDesdeBDNew(cConta, "tiposiva", "nombriva", "codigiva", AdoAux(1).Recordset!CodIVA, "N")
'                            vSeccion.CerrarConta
'                        End If
'                    End If
'                    Set vSeccion = Nothing
'                End If
'        End Select
'    End If
End Sub

' ***** si n'hi han varios nivells de tabs *****
Private Sub SituarTab(numTab As Integer)
    On Error Resume Next
    
'    If numTab = 0 Then
'        SSTab1.Tab = 2
'    ElseIf numTab = 1 Then
'        SSTab1.Tab = 1
'    End If
    
    If Err.Number <> 0 Then Err.Clear
End Sub
' **********************************************


Private Sub CargaFrame(Index As Integer, enlaza As Boolean)
'Dim tip As Integer
'Dim I As Byte
'
'    AdoAux(Index).ConnectionString = Conn
'    AdoAux(Index).RecordSource = MontaSQLCarga(Index, enlaza)
'    AdoAux(Index).CursorType = adOpenDynamic
'    AdoAux(Index).LockType = adLockPessimistic
'    AdoAux(Index).Refresh
'
'    If Not AdoAux(Index).Recordset.EOF Then
'        PonerCamposForma2 Me, AdoAux(Index), 2, "FrameAux" & Index
'        ' *** si n'hi han tabs sense datagrids, li pose els valors als camps ***
'        If (Index = 3) Then 'datos facturacion
'            tip = AdoAux(Index).Recordset!tipclien
'            If (tip = 1) Then 'persona
'                txtAux2(27).Text = AdoAux(Index).Recordset!ape_raso & "," & AdoAux(Index).Recordset!Nom_Come
'            ElseIf (tip = 2) Then 'empresa
'                txtAux2(27).Text = AdoAux(Index).Recordset!Nom_Come
'            End If
'            txtAux2(28).Text = DBLet(AdoAux(Index).Recordset!desforpa, "T")
'            txtAux2(29).Text = DBLet(AdoAux(Index).Recordset!desrutas, "T")
'            'txtAux2(31).Text = DBLet(AdoAux(Index).Recordset!comision, "T") & " %"
'            txtAux2(32).Text = DBLet(AdoAux(Index).Recordset!nomrapel, "T")
'            'Descripcion cuentas contables de la Contabilidad
'            For I = 35 To 38
'                txtAux2(I).Text = PonerNombreDeCod(txtAux(I), "cuentas", "nommacta", "codmacta", , cConta)
'            Next I
'        End If
'        ' ************************************************************************
'    Else
'        ' *** si n'hi han tabs sense datagrids, li pose els valors als camps ***
'        NetejaFrameAux "FrameAux3" 'neteja només lo que te TAG
'        txtAux2(0).Text = ""
'        txtAux2(1).Text = ""
'
''        txtaux2(27).Text = ""
''        txtaux2(28).Text = ""
''        txtaux2(29).Text = ""
'        'txtAux2(31).Text = ""
''        txtaux2(32).Text = ""
''        For i = 35 To 38
''            txtaux2(i).Text = ""
''        Next i
'        ' **********************************************************************
'    End If
End Sub

' *** si n'hi han tabs sense datagrids ***
Private Sub NetejaFrameAux(nom_frame As String)
Dim Control As Object
    
    For Each Control In Me.Controls
        If (Control.Tag <> "") Then
            If (Control.Container.Name = nom_frame) Then
                If TypeOf Control Is TextBox Then
                    Control.Text = ""
                ElseIf TypeOf Control Is ComboBox Then
                    Control.ListIndex = -1
                End If
            End If
        End If
    Next Control

End Sub
' ****************************************


Private Sub CargaGrid(Index As Integer, enlaza As Boolean)
Dim b As Boolean
Dim i As Byte
Dim tots As String

    On Error GoTo ECarga

    tots = MontaSQLCarga(Index, enlaza)

    b = DataGridAux(Index).Enabled
    DataGridAux(Index).Enabled = False
    
    AdoAux(Index).ConnectionString = conn
    AdoAux(Index).RecordSource = MontaSQLCarga(Index, enlaza)
    AdoAux(Index).CursorType = adOpenDynamic
    AdoAux(Index).LockType = adLockPessimistic
    DataGridAux(Index).ScrollBars = dbgNone
    AdoAux(Index).Refresh
    Set DataGridAux(Index).DataSource = AdoAux(Index)
    
    DataGridAux(Index).AllowRowSizing = False
    DataGridAux(Index).RowHeight = 290
    If PrimeraVez Then
        DataGridAux(Index).ClearFields
        DataGridAux(Index).ReBind
        DataGridAux(Index).Refresh
    End If

    For i = 0 To DataGridAux(Index).Columns.Count - 1
        DataGridAux(Index).Columns(i).AllowSizing = False
    Next i
    
    CargaGridGnral Me.DataGridAux(Index), Me.AdoAux(Index), tots, PrimeraVez
    
    
    'DataGridAux(Index).Enabled = b
'    PrimeraVez = False
    
    Select Case Index
        Case 1 'clasificacion segun la calidad
            'si es visible|control|tipo campo|nombre campo|ancho control|
            tots = "N||||0|;N||||0|;N||||0|;S|txtaux1(1)|T|Cód.|800|;S|cmdAux(1)|B|||;" 'codsocio,codsecci
            tots = tots & "S|txtAux2(1)|T|Nombre|3870|;"
            tots = tots & "S|txtaux1(3)|T|Pr.Coop.|1300|;"
            tots = tots & "S|txtaux1(5)|T|Pr.Socio|1300|;"
            arregla tots, DataGridAux(Index), Me
        
            DataGridAux(Index).Columns(4).Alignment = dbgLeft
            DataGridAux(Index).Columns(5).Alignment = dbgRight
            DataGridAux(Index).Columns(6).Alignment = dbgRight
            
            
            b = (Modo = 4) And ((ModoLineas = 1) Or (ModoLineas = 2))
'            BloquearTxt txtAux(14), Not b
'            BloquearTxt txtAux(15), Not b

            If (enlaza = True) And (Not AdoAux(Index).Recordset.EOF) Then 'per a que pose els valors de les arees de text la primera volta
'                txtAux2(4).Text = PonerNombreCuenta(txtaux1(4), Modo)
'                txtAux2(5).Text = PonerNombreCuenta(txtaux1(5), Modo)
'                txtAux2(0).Text = PonerNombreDeCod(txtaux1(6), "tiposiva", "nombriva", "codigiva", "N", cConta)
'                If VisualizaClasificacion Then
'                    PonerClasificacionGrafica
'
''                    SumaTotalPorcentajes
'                End If
            Else
                For i = 0 To 4
                    txtAux1(i).Text = ""
                Next i
                txtAux2(1).Text = ""
'                Me.MSChart1.visible = False
            End If
    End Select
    DataGridAux(Index).ScrollBars = dbgAutomatic
      
    ' **** si n'hi han llínies en grids i camps fora d'estos ****
    If Not AdoAux(Index).Recordset.EOF Then
        DataGridAux_RowColChange Index, 1, 1
    Else
        LimpiarCamposFrame Index
    End If
    ' **********************************************************
      
ECarga:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid: " & DataGridAux(Index).Tag, Err.Description
End Sub


Private Sub InsertarLinea()
'Inserta registre en les taules de Llínies
Dim nomframe As String
Dim b As Boolean

    On Error Resume Next

    ' *** posa els noms del frames, tant si son de grid com si no ***
    Select Case NumTabMto
        Case 1: nomframe = "FrameAux1" 'clasificacion
    End Select
    ' ***************************************************************
    
    If DatosOkLlin(nomframe) Then
        TerminaBloquear
        If InsertarDesdeForm2(Me, 2, nomframe) Then
            ' *** si n'hi ha que fer alguna cosa abas d'insertar
            ' *************************************************
            b = BLOQUEADesdeFormulario2(Me, Data1, 1)
            
            '++monica: en caso de estar insertando seccion y que no existan las
            'cuentas contables hacemos esto para que las inserte en contabilidad.
'            If NumTabMto = 1 Then
'               txtAux2(4).Text = PonerNombreCuenta(txtaux1(4), 3, Text1(0))
'               txtAux2(5).Text = PonerNombreCuenta(txtaux1(5), 3, Text1(0))
'            End If
            
            Select Case NumTabMto
                Case 0, 1 ' *** els index de les llinies en grid (en o sense tab) ***
                    CargaGrid NumTabMto, True
                    If b Then BotonAnyadirLinea NumTabMto
'                Case 3 ' *** els index dels tabs que NO tenen grid ***
'                    CargaFrame 3, True
'                    If b Then BotonModificar
'                    ModoLineas = 0
'                LLamaLineas NumTabMto, 0
            End Select
           
            SituarTab (NumTabMto)
        End If
    End If
End Sub


Private Sub ModificarLinea()
'Modifica registre en les taules de Llínies
Dim nomframe As String
Dim V As Integer
Dim Cad As String
    On Error Resume Next

    ' *** posa els noms del frames, tant si son de grid com si no ***
    Select Case NumTabMto
'        Case 0: nomframe = "FrameAux0" 'telefonos
        Case 1: nomframe = "FrameAux1" 'secciones
    End Select
    ' **************************************************************

    If DatosOkLlin(nomframe) Then
        TerminaBloquear
        If ModificaDesdeFormulario2(Me, 2, nomframe) Then
            ' *** si cal que fer alguna cosa abas d'insertar ***
            If NumTabMto = 0 Then
            End If
            ' ******************************************************
            If BLOQUEADesdeFormulario2(Me, Data1, 1) Then BotonModificar
            ModoLineas = 0

            If NumTabMto <> 3 Then
                V = AdoAux(NumTabMto).Recordset.Fields(4) 'el 2 es el nº de llinia
                CargaGrid NumTabMto, True
            End If

            ' *** si n'hi han tabs ***
            SituarTab (NumTabMto)

            ' *** si n'hi han tabs que no tenen datagrid, posar el if ***
            If NumTabMto <> 3 Then
                DataGridAux(NumTabMto).SetFocus
                AdoAux(NumTabMto).Recordset.Find (AdoAux(NumTabMto).Recordset.Fields(3).Name & " =" & V)
            End If
            ' ***********************************************************

            LLamaLineas NumTabMto, 0
            
        End If
    End If
        
End Sub


Private Sub Modificar()
'Modifica registre en les taules de Llínies
Dim nomframe As String
Dim V As Integer
Dim Cad As String
Dim Sql As String

    On Error GoTo eModificar

    conn.BeginTrans

    ' modificamos los datos del campo
    If ModificaDesdeFormulario2(Me, 1) Then
        TerminaBloquear
        PosicionarData
        CargaGrid 1, True
        VisualizaClasificacion
    End If
    
    conn.CommitTrans
    Exit Sub

eModificar:
    conn.RollbackTrans
    MuestraError Err.Number, "Modificando lineas"

End Sub




Private Function ObtenerWhereCab(conW As Boolean) As String
Dim vWhere As String
    
    vWhere = ""
    If conW Then vWhere = " WHERE "
    ' *** canviar-ho per la clau primaria de la capçalera ***
    vWhere = vWhere & " rprecios_calidad.codvarie=" & Val(Text1(0).Text)
    vWhere = vWhere & " and rprecios_calidad.tipofact = " & Val(Combo1(0).ListIndex)
    vWhere = vWhere & " and rprecios_calidad.contador = " & Val(Text1(1).Text)
    ' *******************************************************
    
    ObtenerWhereCab = vWhere
End Function

' *** neteja els camps dels tabs de grid que
'estan fora d'este, i els camps de descripció ***
Private Sub LimpiarCamposFrame(Index As Integer)
'Dim I As Integer
'    On Error Resume Next
'
'    Select Case Index
'        Case 0 'telefonos
'            For I = 0 To txtAux.Count - 1
'                txtAux(I).Text = ""
'            Next I
'        Case 1 'secciones
'            For I = 0 To txtaux1.Count - 1
'                txtaux1(I).Text = ""
'            Next I
'    End Select
'
'    If Err.Number <> 0 Then Err.Clear
End Sub
' ***********************************************


'' ### [DavidV] 26/04/2006: Activar/desactivar la rueda del ratón.
'Private Sub DataGridAux_GotFocus(Index As Integer)
'  WheelHook DataGridAux(Index)
'End Sub
'Private Sub DataGridAux_LostFocus(Index As Integer)
'  WheelUnHook
'End Sub

' ********* si n'hi han combos a la capçalera ************
Private Sub CargaCombo()
Dim Ini As Integer
Dim Fin As Integer
Dim i As Integer

    ' *** neteje els combos, els pose valor i seleccione el valor per defecte ***
    For i = 0 To Combo1.Count - 1
        Combo1(i).Clear
    Next i
    
    'tipo de precios
    Combo1(0).AddItem "Anticipo"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 0
    Combo1(0).AddItem "Liquidacion"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 1
    
    ' solo hay industria directa y complementaria en horto
    If Tipo = 0 Then
        Combo1(0).AddItem "Industria Directa"
        Combo1(0).ItemData(Combo1(0).NewIndex) = 2
        Combo1(0).AddItem "Complementaria"
        Combo1(0).ItemData(Combo1(0).NewIndex) = 3
        Combo1(0).AddItem "Anticipo Genérico"
        Combo1(0).ItemData(Combo1(0).NewIndex) = 4
        Combo1(0).AddItem "Anticipo Retirada"
        Combo1(0).ItemData(Combo1(0).NewIndex) = 5
        '[Monica]27/01/2016: liquidacion complementaria sobre kilos normales / venta campo y kilos aportados
'        If vParamAplic.Cooperativa = 0 Then
'            Combo1(0).AddItem "Complem.Única"
'            Combo1(0).ItemData(Combo1(0).NewIndex) = 6
'        End If
    End If
End Sub


Private Sub txtAux1_LostFocus(Index As Integer)
Dim cadMen As String
Dim Nuevo As Boolean

    If Not PerderFocoGnral(txtAux1(Index), Modo) Then Exit Sub

    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub


    ' ******* configurar el LostFocus dels camps de llínies (dins i fora grid) ********
    Select Case Index
        Case 1 ' calidad
            If PonerFormatoEntero(txtAux1(Index)) Then
                txtAux2(Index).Text = PonerNombreDeCod(txtAux1(Index), "rcalidad", "nomcalid", "codcalid", "N", , "codvarie", txtAux1(0).Text, "N")
                If txtAux2(Index).Text = "" Then
                    cadMen = "No existe la Calidad: " & txtAux1(Index).Text & vbCrLf
                    cadMen = cadMen & "¿Desea crearla?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmCalid = New frmManCalidades
                        frmCalid.DatosADevolverBusqueda = "0|1|"
                        frmCalid.NuevoCodigo = txtAux1(Index).Text
                        txtAux1(Index).Text = ""
                        TerminaBloquear
                        frmCalid.Show vbModal
                        Set frmCalid = Nothing
                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
                    Else
                        txtAux1(Index).Text = ""
                    End If
                    PonerFoco txtAux1(Index)
                End If
            Else
                txtAux2(Index).Text = ""
            End If

        Case 3, 5 'precios coop y socio
            If PonerFormatoDecimal(txtAux1(Index), 8) Then
                '[Monica]02/05/2012: sustituyo la siguiente instruccion para Picassent por un check para todos
                '[Monica]22/06/2011: si estan insertando en picassent quiere que se repita el valor del primer precio en el segundo
                'If vParamAplic.Cooperativa = 2 And Index = 3 And txtaux1(5).Text = "" And ModoLineas = 1 Then txtaux1(5).Text = txtaux1(3).Text
                If Me.Check1.Value = 1 And Index = 3 Then
                    txtAux1(5).Text = txtAux1(3).Text
                    cmdAceptar.SetFocus
                End If
                
                If Index = 5 Then cmdAceptar.SetFocus
            End If

    End Select

    ' ******************************************************************************
End Sub

Private Sub txtAux1_GotFocus(Index As Integer)
   If Not txtAux1(Index).MultiLine Then ConseguirFocoLin txtAux1(Index)
End Sub

Private Sub txtAux1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Not txtAux1(Index).MultiLine Then KEYdown KeyCode
End Sub

Private Sub txtAux1_KeyPress(Index As Integer, KeyAscii As Integer)
    If Not txtAux1(Index).MultiLine Then
        If KeyAscii = teclaBuscar Then
            If Modo = 5 And (ModoLineas = 1 Or ModoLineas = 2) Then
                Select Case Index
                    Case 1: KEYBusqueda KeyAscii, 1 'calidad
                End Select
            End If
        Else
            KEYpress KeyAscii
        End If
    End If
End Sub


Private Sub PonerDatosPartida()
Dim Zona As String
Dim OtroCampo As String
Dim CodPobla As String

'    Zona = ""
'    Text3(3).Text = ""
'    Text4(3).Text = ""
'    Text5(3).Text = ""
'
'    OtroCampo = "codpobla"
'    Zona = DevuelveDesdeBDNew(cAgro, "rpartida", "codzonas", "codparti", Text1(3), "N", OtroCampo)
'    Text3(3).Text = Zona
'
'    If Zona <> "" Then
'        Text4(3).Text = DevuelveDesdeBDNew(cAgro, "rzonas", "nomzonas", "codzonas", Zona, "N")
'        If OtroCampo <> "" Then
'            CodPobla = OtroCampo
'            If CodPobla <> "" Then Text5(3).Text = DevuelveDesdeBDNew(cAgro, "rpueblos", "despobla", "codpobla", CodPobla, "T")
'        End If
'    End If

End Sub


Private Sub LlamarImprimir()
    With frmImprimir
        .FormulaSeleccion = cadFormula
        .OtrosParametros = CadParam
        .NumeroParametros = numParam
        .SoloImprimir = False
        .EnvioEMail = False
        .Titulo = cadTitulo
        .NombreRPT = cadNombreRPT
        .Opcion = 0
        .Show vbModal
    End With
End Sub

Private Sub InicializarVbles()
    cadFormula = ""
    cadSelect = ""
    CadParam = ""
    numParam = 0
End Sub

Private Sub PonerClasificacionGrafica()
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim i As Integer
Dim arrData()
Dim TotalPorc As Currency
   
'    SQL = "select count(*) from rcampos_clasif, rcalidad where rcampos_clasif.codcampo = " & Data1.Recordset!codCampo
'    SQL = SQL & " and rcampos_clasif.codvarie = rcalidad.codvarie "
'    SQL = SQL & " and rcampos_clasif.codcalid = rcalidad.codcalid "
'
'    I = TotalRegistros(SQL)
'
'    MSChart1.visible = True
'    If I = 0 Then
'        MSChart1.visible = False
'        Exit Sub
'    End If
'
'    ReDim arrData(I - 1, 2)
'
'    SQL = "select rcampos_clasif.muestra, rcalidad.nomcalid from rcampos_clasif, rcalidad where rcampos_clasif.codcampo = " & Data1.Recordset!codCampo
'    SQL = SQL & " and rcampos_clasif.codvarie = rcalidad.codvarie "
'    SQL = SQL & " and rcampos_clasif.codcalid = rcalidad.codcalid "
'    SQL = SQL & " order by rcampos_clasif.codcalid "
'
'    Set RS = New ADODB.Recordset
'    RS.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
'
'    I = 0
'    TotalPorc = 0
'    While Not RS.EOF
'        arrData(I, 1) = DBLet(RS!nomcalid, "T") '& " " & DBLet(Rs!muestra, "N")
'        arrData(I, 2) = DBLet(RS!muestra, "N")
'
'        TotalPorc = TotalPorc + DBLet(RS!muestra, "N")
'
'        I = I + 1
'
'        RS.MoveNext
'    Wend
'
'    Set RS = Nothing
'
'    MSChart1.ChartData = arrData
'    MSChart1.ColumnLabel = "Porcentaje Total : " & TotalPorc & "%"
'
''    arrData(0, 1) = "Ene"   ' Establece las etiquetas en la primera serie.
''    arrData(1, 1) = "Feb"
''    arrData(2, 1) = "Mar"
''
''    arrData(0, 2) = 8
''    arrData(1, 2) = 4
''    arrData(2, 2) = 0.3
End Sub


Private Sub SumaTotalPorcentajes()
Dim Sql As String
Dim i As Currency
Dim Rs As ADODB.Recordset
   
'    SQL = "select sum(muestra) from rcampos_clasif where rcampos_clasif.codcampo = " & Data1.Recordset!codCampo
'
'    Set RS = New ADODB.Recordset
'    RS.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
'
'    I = 0
'    If Not RS.EOF Then
'        I = DBLet(RS.Fields(0).Value, "N")
'    End If
'
'    If I <> 100 Then
'        MsgBox "La suma de muestras es " & I & ". Debe de ser 100%. Revise.", vbExclamation
'    End If

End Sub


Private Function VisualizaClasificacion() As Boolean
Dim Sql As String


'    SQL = ""
'    SQL = DevuelveDesdeBDNew(cAgro, "variedades", "tipoclasifica", "codvarie", Data1.Recordset!CodVarie, "N")
'
'    SSTab1.TabEnabled(1) = (SQL = "0")
'    SSTab1.TabVisible(1) = (SQL = "0")
'
'    VisualizaClasificacion = (SQL = "0")

End Function


