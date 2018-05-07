VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmFVARConceptos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Conceptos de Facturas"
   ClientHeight    =   7140
   ClientLeft      =   45
   ClientTop       =   30
   ClientWidth     =   12945
   Icon            =   "frmFVARConceptos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7140
   ScaleWidth      =   12945
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameBotonGnral 
      Height          =   705
      Left            =   120
      TabIndex        =   30
      Top             =   30
      Width           =   3585
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   210
         TabIndex        =   31
         Top             =   180
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   10
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
               Object.Width           =   1e-4
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Buscar"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Ver Todos"
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Imprimir"
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Object.Visible         =   0   'False
               Object.ToolTipText     =   "Salir"
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Object.Visible         =   0   'False
               Style           =   3
            EndProperty
         EndProperty
      End
   End
   Begin VB.CommandButton btnBuscar 
      Appearance      =   0  'Flat
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   4
      Left            =   8895
      MaskColor       =   &H00000000&
      TabIndex        =   27
      ToolTipText     =   "Buscar Cta.Contable"
      Top             =   3900
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   10
      Left            =   8040
      MaxLength       =   10
      TabIndex        =   4
      Tag             =   "Cuenta Contable Prov.|T|N|||fvarconce|codmacpr|||"
      Top             =   3900
      Width           =   870
   End
   Begin VB.TextBox txtAux2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   1
      Left            =   3900
      Locked          =   -1  'True
      MaxLength       =   40
      TabIndex        =   26
      Text            =   "Nombre seccion"
      Top             =   3900
      Visible         =   0   'False
      Width           =   2865
   End
   Begin VB.CommandButton btnBuscar 
      Appearance      =   0  'Flat
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   3
      Left            =   3660
      MaskColor       =   &H00000000&
      TabIndex        =   25
      ToolTipText     =   "Buscar Sección"
      Top             =   3900
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.CommandButton btnBuscar 
      Appearance      =   0  'Flat
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   2
      Left            =   10770
      MaskColor       =   &H00000000&
      TabIndex        =   22
      ToolTipText     =   "Buscar Centro Coste"
      Top             =   3900
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.TextBox txtAux 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   8
      Left            =   10170
      MaxLength       =   4
      TabIndex        =   6
      Tag             =   "Centro Coste|T|S|||fvarconce|codccost|||"
      Text            =   "Ccos"
      Top             =   3900
      Width           =   555
   End
   Begin VB.CommandButton btnBuscar 
      Appearance      =   0  'Flat
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   1
      Left            =   9930
      MaskColor       =   &H00000000&
      TabIndex        =   17
      ToolTipText     =   "Buscar Tipo Iva"
      Top             =   3885
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Frame Frame2 
      Height          =   975
      Left            =   105
      TabIndex        =   14
      Top             =   5430
      Width           =   12600
      Begin VB.TextBox txtAux 
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
         Index           =   11
         Left            =   3945
         MaxLength       =   30
         TabIndex        =   28
         Top             =   495
         Width           =   3810
      End
      Begin VB.TextBox txtAux 
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
         Index           =   9
         Left            =   10500
         MaxLength       =   30
         TabIndex        =   23
         Top             =   495
         Width           =   1995
      End
      Begin VB.TextBox txtAux 
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
         Index           =   7
         Left            =   9900
         MaxLength       =   30
         TabIndex        =   20
         Top             =   495
         Width           =   555
      End
      Begin VB.TextBox txtAux 
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
         Index           =   5
         Left            =   7800
         MaxLength       =   30
         TabIndex        =   18
         Top             =   495
         Width           =   2055
      End
      Begin VB.TextBox txtAux 
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
         Index           =   6
         Left            =   150
         MaxLength       =   30
         TabIndex        =   15
         Top             =   510
         Width           =   3750
      End
      Begin VB.Label Label5 
         Caption         =   "Cuenta Contable Proveedor"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3930
         TabIndex        =   29
         Top             =   240
         Width           =   2925
      End
      Begin VB.Label Label4 
         Caption         =   "Centro Coste"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   10500
         TabIndex        =   24
         Top             =   240
         Width           =   1905
      End
      Begin VB.Label Label3 
         Caption         =   "%Iva"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   9900
         TabIndex        =   21
         Top             =   225
         Width           =   525
      End
      Begin VB.Label Label2 
         Caption         =   "Tipo de Iva"
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
         Left            =   7800
         TabIndex        =   19
         Top             =   225
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Cuenta Contable Cliente"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   135
         TabIndex        =   16
         Top             =   255
         Width           =   2385
      End
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   3
      Left            =   6870
      MaxLength       =   10
      TabIndex        =   3
      Tag             =   "Cuenta Contable Cliente|T|N|||fvarconce|codmacta|||"
      Text            =   "CtaContabl"
      Top             =   3915
      Width           =   870
   End
   Begin VB.CommandButton btnBuscar 
      Appearance      =   0  'Flat
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   290
      Index           =   0
      Left            =   7725
      MaskColor       =   &H00000000&
      TabIndex        =   13
      ToolTipText     =   "Buscar Cta.Contable"
      Top             =   3915
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   0
      Left            =   135
      MaxLength       =   3
      TabIndex        =   0
      Tag             =   "Codigo|N|N|1|999|fvarconce|codconce|000|S|"
      Text            =   "Cod"
      Top             =   3930
      Width           =   555
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   1
      Left            =   810
      MaxLength       =   30
      TabIndex        =   1
      Tag             =   "Nombre|T|N|||fvarconce|nomconce|||"
      Text            =   "Nombre"
      Top             =   3930
      Width           =   2295
   End
   Begin VB.TextBox txtAux 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   4
      Left            =   9390
      MaxLength       =   2
      TabIndex        =   5
      Tag             =   "Tipo Iva|N|N|0|99|fvarconce|tipoiva|00||"
      Text            =   "Iv"
      Top             =   3885
      Width           =   555
   End
   Begin VB.TextBox txtAux 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   2
      Left            =   3150
      MaxLength       =   3
      TabIndex        =   2
      Tag             =   "Sección|N|N|0|999|fvarconce|codsecci|000||"
      Text            =   "C"
      Top             =   3915
      Width           =   555
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10515
      TabIndex        =   7
      Top             =   6585
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11640
      TabIndex        =   8
      Top             =   6585
      Visible         =   0   'False
      Width           =   1095
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmFVARConceptos.frx":000C
      Height          =   4545
      Left            =   120
      TabIndex        =   11
      Top             =   810
      Width           =   12565
      _ExtentX        =   22172
      _ExtentY        =   8017
      _Version        =   393216
      AllowUpdate     =   0   'False
      BorderStyle     =   0
      HeadLines       =   1
      RowHeight       =   19
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
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
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11640
      TabIndex        =   12
      Top             =   6585
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Index           =   1
      Left            =   120
      TabIndex        =   9
      Top             =   6480
      Width           =   2385
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
         Height          =   255
         Left            =   45
         TabIndex        =   10
         Top             =   210
         Width           =   2295
      End
   End
   Begin MSAdodcLib.Adodc adodc1 
      Height          =   330
      Left            =   4410
      Top             =   45
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
   Begin MSComctlLib.Toolbar ToolbarAyuda 
      Height          =   330
      Left            =   12240
      TabIndex        =   32
      Top             =   180
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   1
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Ayuda"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnOpciones 
      Caption         =   "&Opciones"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu mnBuscar 
         Caption         =   "&Buscar"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnVerTodos 
         Caption         =   "&Ver todos"
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
      Begin VB.Menu mnBarra2 
         Caption         =   "-"
      End
      Begin VB.Menu mnImprimir 
         Caption         =   "&Imprimir"
         Shortcut        =   ^I
      End
      Begin VB.Menu mnBarra3 
         Caption         =   "-"
      End
      Begin VB.Menu mnSalir 
         Caption         =   "&Salir"
         Shortcut        =   ^S
      End
   End
End
Attribute VB_Name = "frmFVARConceptos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' +-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-
' +-+- Autor: MANOLO (Se lo copia)                          +-+-
' +-+- Menú: Bancos Propios (con un par)                    +-+-
' +-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-

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

Private Const IdPrograma = 6019



Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)

'codi per al registe que s'afegix al cridar des d'altre formulari.
'Obrir en modo Insertar i tornar datos del registre insertat
Public NuevoCodigo As String

Public DeConsulta As Boolean
Public CodigoActual As String


' *** adrede: per a quan busque suplements/desconters des de frmViagrc ***
Public ExpedBusca As Long
Public TipoSuplem As Integer
' *********************************************************************

' *** declarar els formularis als que vaig a cridar ***
'Private WithEvents frmB As frmBuscaGrid

Private CadenaConsulta As String
Private CadB As String

' ### [Monica] 08/09/2006
Private WithEvents frmCtas As frmCtasConta 'cuentas contables
Attribute frmCtas.VB_VarHelpID = -1
Private WithEvents frmTIva As frmTipIVAConta 'tipos de via de contabilidad
Attribute frmTIva.VB_VarHelpID = -1
Private WithEvents frmCCos As frmCCosConta 'centros de coste de contabilidad
Attribute frmCCos.VB_VarHelpID = -1
Private WithEvents frmSec As frmManSeccion 'secciones de ariagro
Attribute frmSec.VB_VarHelpID = -1

Dim Indice As Byte 'Index del text1 on es poses els datos retornats des d'atres Formularis de Mtos

Private kCampo As Integer

Dim vSeccion As CSeccion

Dim Modo As Byte
'----------- MODOS --------------------------------
'   0.-  Formulari net sense cap camp ple
'   1.-  Preparant per a fer la búsqueda
'   2.-  Ja tenim registres i els anem a recorrer
'        ,podem editar-los Edició del camp
'   3.-  Inserció de nou registre
'   4.-  Modificar
'--------------------------------------------------

Private Sub PonerModo(vModo)
Dim B As Boolean
Dim I As Integer
    
    Modo = vModo
'    PonerIndicador lblIndicador, Modo
    
    B = (Modo = 2)
    If B Then
        PonerContRegIndicador
    Else
        PonerIndicador lblIndicador, Modo
    End If
    
    ' **** posar tots els controls (botons inclosos) que siguen del Grid
    For I = 0 To 4
        txtAux(I).visible = Not B
        txtAux(I).BackColor = vbWhite
    Next I
    txtAux(8).visible = Not B ' centro de coste
    txtAux(10).visible = Not B ' centro de coste
    
    For I = 0 To btnBuscar.Count - 1
        btnBuscar(I).visible = (Modo = 3 Or Modo = 4 Or Modo = 1)
        btnBuscar(I).Enabled = (Modo = 3 Or Modo = 4 Or Modo = 1)
    Next I
    ' **************************************************
    txtAux2(1).visible = (Modo = 3 Or Modo = 4)
    txtAux2(1).Enabled = (Modo = 3 Or Modo = 4)
    
    ' **** si n'hi han camps fora del grid, bloquejar-los ****
    For I = 5 To 7
        BloquearTxt txtAux(I), True
    Next I
    BloquearTxt txtAux(9), True ' nombre de centro de coste
    BloquearTxt txtAux(11), True ' nombre de cuenta de proveedor
    
    ' ********************************************************

    cmdAceptar.visible = Not B
    cmdCancelar.visible = Not B
    DataGrid1.Enabled = B
    
    'Si es retornar
    If DatosADevolverBusqueda <> "" Then cmdRegresar.visible = B
    
    PonerLongCampos
    PonerModoOpcionesMenu 'Activar/Desact botons de menu según Modo
    PonerOpcionesMenu 'Activar/Desact botons de menu según permissos de l'usuari
    
    ' *** bloquejar tota la PK quan estem en modificar  ***
    BloquearTxt txtAux(0), (Modo = 4) 'codconce
    
    BloquearImgBuscar Me, Modo

End Sub

Private Sub PonerModoOpcionesMenu()
'Activa/Desactiva botons de la toolbar i del menu, según el modo en que estiguem
Dim B As Boolean

    ' *** adrede: per a que no es puga fer res si estic cridant des de frmViagrc ***

    B = (Modo = 2) And ExpedBusca = 0
    'Busqueda
    Toolbar1.Buttons(5).Enabled = B
    Me.mnBuscar.Enabled = B
    'Vore Tots
    Toolbar1.Buttons(6).Enabled = B
    Me.mnVerTodos.Enabled = B
    
    'Insertar
    Toolbar1.Buttons(1).Enabled = B And Not DeConsulta
    Me.mnNuevo.Enabled = B And Not DeConsulta
    
    B = (B And adodc1.Recordset.RecordCount > 0) And Not DeConsulta And ExpedBusca = 0
    'Modificar
    Toolbar1.Buttons(2).Enabled = B
    Me.mnModificar.Enabled = B

    'Eliminar
    Toolbar1.Buttons(3).Enabled = B
    Me.mnEliminar.Enabled = B
    'Imprimir
    Toolbar1.Buttons(8).Enabled = B
    Me.mnImprimir.Enabled = B

    ' ******************************************************************************
End Sub

Private Sub BotonAnyadir()
Dim NumF As String
Dim anc As Single
Dim I As Integer
    
    CargaGrid 'primer de tot carregue tot el grid
    CadB = ""
    '********* canviar taula i camp; repasar codEmpre ************
    If (DatosADevolverBusqueda <> "") And NuevoCodigo <> "" Then
        NumF = NuevoCodigo
    Else
        NumF = SugerirCodigoSiguienteStr("fvarconce", "codconce")
        'NumF = SugerirCodigoSiguienteStr("sdexpgrp", "codsupdt", "codempre=" & vSesion.Empresa)
        'NumF = ""
    End If
    '***************************************************************
    'Situem el grid al final
    Modo = 3
    
    AnyadirLinea DataGrid1, adodc1

    anc = DataGrid1.Top
    If DataGrid1.Row < 0 Then
        anc = anc + 240
    Else
        anc = anc + DataGrid1.RowTop(DataGrid1.Row) + 5
    End If
    
    ' *** valors per defecte a l'afegir (dins i fora del grid); repasar codEmpre ***
    txtAux(0).Text = NumF
    For I = 1 To 9
        txtAux(I).Text = ""
    Next I
    txtAux2(1).Text = ""
    
    ' **************************************************
    
    LLamaLineas anc, 3
       
    ' *** posar el foco ***
    If (DatosADevolverBusqueda <> "") And NuevoCodigo <> "" Then
        PonerFoco txtAux(1) '**** 1r camp visible que NO siga PK ****
    Else
        PonerFoco txtAux(0) '**** 1r camp visible que siga PK ****
    End If
    ' ******************************************************
End Sub

Private Sub BotonVerTodos()
    CargaGrid ""
    CadB = ""
    PonerModo 2
End Sub

Private Sub BotonBuscar()
    Dim I As Integer
    
    ' *** canviar per la PK (no posar codempre si està a Form_Load) ***
    CargaGrid "codconce = -1"
    '*******************************************************************************

    ' *** canviar-ho pels valors per defecte al buscar (dins i fora del grid);
    For I = 0 To 7
        txtAux(I).Text = ""
    Next I
    txtAux2(1).Text = ""
    
    LLamaLineas DataGrid1.Top + 240, 1
    
    ' *** posar el foco al 1r camp visible que siga PK ***
    PonerFoco txtAux(0)
    ' ***************************************************************
End Sub

Private Sub BotonModificar()
    Dim anc As Single
    Dim I As Integer

    Screen.MousePointer = vbHourglass
    
    If DataGrid1.Bookmark < DataGrid1.FirstRow Or DataGrid1.Bookmark > (DataGrid1.FirstRow + DataGrid1.VisibleRows - 1) Then
        I = DataGrid1.Bookmark - DataGrid1.FirstRow
        DataGrid1.Scroll 0, I
        DataGrid1.Refresh
    End If
    
    If DataGrid1.Row < 0 Then
        anc = 320
    Else
        anc = DataGrid1.RowTop(DataGrid1.Row) + DataGrid1.Top '545
    End If

    ' *** asignar als controls del grid, els valors de les columnes ***
    txtAux(0).Text = DataGrid1.Columns(0).Text
    txtAux(1).Text = DataGrid1.Columns(1).Text
    txtAux(2).Text = ComprobarCero(Trim(DataGrid1.Columns(2).Text))
    txtAux2(1).Text = DataGrid1.Columns(3).Text
    txtAux(3).Text = DataGrid1.Columns(4).Text
    txtAux(4).Text = ComprobarCero(Trim(DataGrid1.Columns(6).Text))
    txtAux(8).Text = DataGrid1.Columns(7).Text
    
    
    Set vSeccion = New CSeccion
    If vSeccion.LeerDatos(txtAux(2)) Then
        txtAux2(1).Text = vSeccion.Nombre
        If vSeccion.AbrirConta Then
        
        End If
    End If
    
    ' ********************************************************

    LLamaLineas anc, 4 'modo 4
   
    ' *** foco al 1r camp visible que NO siga clau primaria ***
    PonerFoco txtAux(1)
    ' *********************************************************
    Screen.MousePointer = vbDefault
End Sub

Private Sub LLamaLineas(alto As Single, xModo As Byte)
Dim I As Integer

    DeseleccionaGrid Me.DataGrid1
    PonerModo xModo

    ' *** posar el Top a tots els controls del grid (botons també) ***
    'Me.imgFec(2).Top = alto
    For I = 0 To 4
        txtAux(I).Top = alto
    Next I
    txtAux(8).Top = alto ' centro de coste
    txtAux(10).Top = alto ' centro de coste
    
    txtAux2(1).Top = alto
    For I = 0 To btnBuscar.Count - 1
        btnBuscar(I).Top = alto
    Next I
    ' ***************************************************
End Sub

Private Sub BotonEliminar()
Dim Sql As String
Dim temp As Boolean

    On Error GoTo Error2
    
    'Certes comprovacions
    If adodc1.Recordset.EOF Then Exit Sub
'    If Not SepuedeBorrar Then Exit Sub
    
    ' *** repasar el nom de l'adodc, l'index del Field i el camp que te la PK ***
    'El registre de codi 0 no es pot Modificar ni Eliminar
    If EsCodigoCero(CStr(adodc1.Recordset.Fields(1).Value), FormatoCampo(txtAux(1))) Then Exit Sub
    ' ***************************************************************************
    
    '*** canviar la pregunta, els noms dels camps i el DELETE; repasar codEmpre ***
    Sql = "¿Seguro que desea eliminar el Concepto?"
    'SQL = SQL & vbCrLf & "Código: " & Format(adodc1.Recordset.Fields(0), "000")
    Sql = Sql & vbCrLf & "Código: " & adodc1.Recordset.Fields(0)
    Sql = Sql & vbCrLf & "Nombre: " & adodc1.Recordset.Fields(1)
    
    If MsgBox(Sql, vbQuestion + vbYesNo) = vbYes Then
        'N'hi ha que eliminar
        NumRegElim = adodc1.Recordset.AbsolutePosition
        Sql = "Delete from fvarconce where codconce = " & adodc1.Recordset!codConce
        
        conn.Execute Sql
        CargaGrid CadB
'        If CadB <> "" Then
'            CargaGrid CadB
'            lblIndicador.Caption = "RESULTADO BUSQUEDA"
'        Else
'            CargaGrid ""
'            lblIndicador.Caption = ""
'        End If
        temp = SituarDataTrasEliminar(adodc1, NumRegElim, True)
        PonerModoOpcionesMenu
        adodc1.Recordset.Cancel
    End If
    Exit Sub
    
Error2:
    Screen.MousePointer = vbDefault
    MuestraError Err.Number, "Eliminando registro", Err.Description
End Sub

Private Sub PonerLongCampos()
'Modificar el MaxLength del campo en funcion de si es modo de búsqueda o no
'para los campos que permitan introducir criterios más largos del tamaño del campo
'Si maxlength=6 para codprove, en modo busqueda Maxlenth=13 'el doble + 1
    
    'para los txtAux
    PonerLongCamposGnral Me, Modo, 3
End Sub

Private Sub btnBuscar_Click(Index As Integer)
    TerminaBloquear
     Select Case Index
        Case 0 'Cuentas Contables (de contabilidad)
            If txtAux(2).Text = "" Then Exit Sub
            
            Set vSeccion = New CSeccion
            If vSeccion.LeerDatos(txtAux(2).Text) Then
                If vSeccion.AbrirConta Then
                    Indice = 3
                    Set frmCtas = New frmCtasConta
                    frmCtas.NumDigit = 0 'DevuelveDesdeBDNew(cConta, "empresa", "numdigi" & vEmpresa.DigitosUltimoNivel, "", "", "")
'                    frmCtas.CadBusqueda = DevuelveDesdeBDNew(cConta, "parametros", "grupovta", "", "", "T")
                    frmCtas.DatosADevolverBusqueda = "0|1|"
                    frmCtas.CodigoActual = txtAux(Indice).Text
                    frmCtas.Show vbModal
                    Set frmCtas = Nothing
                    PonerFoco txtAux(Indice)
                    
                    vSeccion.CerrarConta
                End If
            End If
            Set vSeccion = Nothing
            
        Case 4 'Cuentas Contables (de contabilidad) de compras
            If txtAux(2).Text = "" Then Exit Sub
            
            Set vSeccion = New CSeccion
            If vSeccion.LeerDatos(txtAux(2).Text) Then
                If vSeccion.AbrirConta Then
                    Indice = 10
                    Set frmCtas = New frmCtasConta
                    frmCtas.NumDigit = 0 'DevuelveDesdeBDNew(cConta, "empresa", "numdigi" & vEmpresa.DigitosUltimoNivel, "", "", "")
'                    frmCtas.CadBusqueda = DevuelveDesdeBDNew(cConta, "parametros", "grupovta", "", "", "T")
                    frmCtas.DatosADevolverBusqueda = "0|1|"
                    frmCtas.CodigoActual = txtAux(Indice).Text
                    frmCtas.Show vbModal
                    Set frmCtas = Nothing
                    PonerFoco txtAux(Indice)
                    
                    vSeccion.CerrarConta
                End If
            End If
            Set vSeccion = Nothing
            
            
            
            
        Case 1 'Tipos de Iva (de contabilidad)
            If txtAux(2).Text = "" Then Exit Sub
            
            Set vSeccion = New CSeccion
            If vSeccion.LeerDatos(txtAux(2).Text) Then
                If vSeccion.AbrirConta Then
                    Indice = 4
                    Set frmTIva = New frmTipIVAConta
                    frmTIva.DatosADevolverBusqueda = "0|1|2|"
                    frmTIva.CodigoActual = txtAux(Indice).Text
                    frmTIva.Show vbModal
                    Set frmTIva = Nothing
                    PonerFoco txtAux(Indice)
                    
                    vSeccion.CerrarConta
                End If
            End If
            Set vSeccion = Nothing
            
        Case 2 'Centros de coste de contabilidad
            If txtAux(2).Text = "" Then Exit Sub
            
            Set vSeccion = New CSeccion
            If vSeccion.LeerDatos(txtAux(2).Text) Then
                If vSeccion.AbrirConta Then

                    Indice = 8
                    Set frmCCos = New frmCCosConta
                    frmCCos.DatosADevolverBusqueda = "0|1|"
                    frmCCos.CodigoActual = txtAux(Indice).Text
                    frmCCos.Show vbModal
                    Set frmCCos = Nothing
                    PonerFoco txtAux(Indice)
            
                    vSeccion.CerrarConta
                End If
            End If
            Set vSeccion = Nothing
            
            
        Case 3:  ' seccion
        
            Set frmSec = New frmManSeccion
            frmSec.DatosADevolverBusqueda = "0|1|"
            frmSec.CodigoActual = txtAux(1).Text
            frmSec.Show vbModal
            Set frmSec = Nothing
            PonerFoco txtAux(2)
       
            
    End Select
    
    If Modo = 4 Then BLOQUEADesdeFormulario2 Me, adodc1, 1

End Sub

Private Sub cmdAceptar_Click()
Dim I As Long

    Select Case Modo
        Case 1  'BUSQUEDA
            CadB = ObtenerBusqueda(Me)
            If CadB <> "" Then
                CargaGrid CadB
                PonerModo 2
'                lblIndicador.Caption = "RESULTADO BUSQUEDA"
                PonerFocoGrid Me.DataGrid1
            End If
        
        Case 3 'INSERTAR
            If DatosOK Then
                'If InsertarDesdeForm(Me) Then
                If InsertarDesdeForm2(Me, 0) Then
                    CargaGrid
                    If (DatosADevolverBusqueda <> "") And NuevoCodigo <> "" Then
                        cmdCancelar_Click
                        If Not adodc1.Recordset.EOF Then
                            ' *** filtrar per tota la PK; repasar codEmpre **
                            'adodc1.Recordset.Filter = "codempre = " & txtAux(0).Text & " AND codsupdt = " & txtAux(1).Text
                            adodc1.Recordset.Filter = "codsecci = " & txtAux(0).Text
                            ' ****************************************************
                        End If
                        cmdRegresar_Click
                    Else
                        If Not vSeccion Is Nothing Then
                            vSeccion.CerrarConta
                            Set vSeccion = Nothing
                        End If

                        BotonAnyadir
                    End If
                    CadB = ""
                End If
            End If
            
        Case 4 'MODIFICAR
            If DatosOK Then
                'If ModificaDesdeFormulario(Me) Then
                If ModificaDesdeFormulario2(Me, 0) Then
                    I = adodc1.Recordset.AbsolutePosition
                    TerminaBloquear
                    PonerModo 2
                    CargaGrid CadB
'                    If CadB <> "" Then
'                        CargaGrid CadB
'                        lblIndicador.Caption = "RESULTADO BUSQUEDA"
'                    Else
'                        CargaGrid
'                        lblIndicador.Caption = ""
'                    End If
                    If Not vSeccion Is Nothing Then
                        vSeccion.CerrarConta
                        Set vSeccion = Nothing
                    End If
                    adodc1.Recordset.Move I - 1
                    PonerFocoGrid Me.DataGrid1
                End If
            End If
            
    End Select
End Sub

Private Sub cmdCancelar_Click()
'On Error Resume Next

    Select Case Modo
        Case 3 'INSERTAR
            DataGrid1.AllowAddNew = False
            'CargaGrid
            If Not adodc1.Recordset.EOF Then adodc1.Recordset.MoveFirst
        Case 4 'MODIFICAR
            TerminaBloquear
        Case 1 'BUSQUEDA
            CargaGrid CadB
    End Select
    
    If Not adodc1.Recordset.EOF Then
        CargaForaGrid
    Else
        LimpiarCampos
    End If
    
    PonerModo 2
'    If CadB <> "" Then
'        lblIndicador.Caption = "RESULTADO BUSQUEDA"
'    Else
'        lblIndicador.Caption = ""
'    End If
    PonerFocoGrid Me.DataGrid1
'    If Err.Number <> 0 Then Err.Clear
End Sub

Private Sub cmdRegresar_Click()
Dim cad As String
Dim I As Integer
Dim J As Integer
Dim Aux As String

    If adodc1.Recordset.EOF Then
        MsgBox "Ningún registro devuelto.", vbExclamation
        Exit Sub
    End If
    cad = ""
    I = 0
    Do
        J = I + 1
        I = InStr(J, DatosADevolverBusqueda, "|")
        If I > 0 Then
            Aux = Mid(DatosADevolverBusqueda, J, I - J)
            J = Val(Aux)
            cad = cad & adodc1.Recordset.Fields(J) & "|"
        End If
    Loop Until I = 0
    ' *** adrede: per a tornar el TipoSuplem ***
    ' cad = cad & TipoSuplem & "|"
    ' ******************************************
    RaiseEvent DatoSeleccionado(cad)
    Unload Me
End Sub

Private Sub DataGrid1_DblClick()
    If cmdRegresar.visible Then cmdRegresar_Click
End Sub


Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
    'Posem el foco
    If (DatosADevolverBusqueda <> "") And NuevoCodigo <> "" Then
        PonerFoco txtAux(1)
    End If
End Sub


Private Sub Form_Load()
Dim I As Integer

    'Icono del formulario
    Me.Icon = frmPpal.Icon

    '******* repasar si n'hi ha botó d'imprimir o no******
    With Me.Toolbar1
        .HotImageList = frmPpal.imgListComun_OM
        .DisabledImageList = frmPpal.imgListComun_BN
        .ImageList = frmPpal.imgListComun
        'el 1 es separadors
        'el 4 i el 5 son separadors
        .Buttons(1).Image = 3   'Insertar
        .Buttons(2).Image = 4   'Modificar
        .Buttons(3).Image = 5   'Borrar
        
        .Buttons(5).Image = 1   'Buscar
        .Buttons(6).Image = 2   'Tots
        'el 9 i el 10 son separadors
        .Buttons(8).Image = 10  'Imprimir
    End With
    '*****************************************************

    ' La Ayuda
    With Me.ToolbarAyuda
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 12
    End With

    ' *** SI N'HI HAN COMBOS ***
    ' CargaCombo 0
    ' **************************
    
    '****************** canviar la consulta *********************************+
    CadenaConsulta = "SELECT fvarconce.codconce, fvarconce.nomconce, fvarconce.codsecci, nomsecci, fvarconce.codmacta, fvarconce.codmacpr, fvarconce.tipoiva, fvarconce.codccost "
    CadenaConsulta = CadenaConsulta & " FROM fvarconce inner join rseccion on fvarconce.codsecci = rseccion.codsecci "
    '************************************************************************
    
    CadB = ""
    CargaGrid
    
    ' ****** Si n'hi han camps fora del grid ******
    CargaForaGrid
    ' *********************************************
    
    If (DatosADevolverBusqueda <> "") And NuevoCodigo <> "" Then
        BotonAnyadir
    Else
        PonerModo 2
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Screen.MousePointer = vbDefault
End Sub

Private Sub ToolbarAyuda_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1
            LanzaVisorMimeDocumento Me.hWnd, DireccionAyuda & IdPrograma & ".html"
    End Select
End Sub


Private Sub frmCCos_DatoSeleccionado(CadenaSeleccion As String)
'Centro de coste de la Contabilidad
    txtAux(Indice).Text = RecuperaValor(CadenaSeleccion, 1) 'codccost
    txtAux(9).Text = RecuperaValor(CadenaSeleccion, 2) 'nomcoste
End Sub

' ### [Monica] 08/09/2006
Private Sub frmCtas_DatoSeleccionado(CadenaSeleccion As String)
'Cuentas contables de la Contabilidad
    txtAux(Indice).Text = RecuperaValor(CadenaSeleccion, 1) 'codmacta
    If Indice = 3 Then
        txtAux(6).Text = RecuperaValor(CadenaSeleccion, 2) 'des macta
    Else
        txtAux(11).Text = RecuperaValor(CadenaSeleccion, 2) 'des macta
    End If
End Sub



Private Sub imgMail_Click(Index As Integer)
    If Index = 0 Then
        If txtAux(16).Text <> "" Then
            LanzaMailGnral txtAux(16).Text
        End If
    End If
End Sub

Private Sub imgWeb_Click(Index As Integer)
    If Modo = 0 Then Exit Sub
    Screen.MousePointer = vbHourglass
    If LanzaHomeGnral(txtAux(15).Text) Then espera 2
    Screen.MousePointer = vbDefault
End Sub

Private Sub frmSec_DatoSeleccionado(CadenaSeleccion As String)
    txtAux(2).Text = RecuperaValor(CadenaSeleccion, 1) 'codigo de seccion
    txtAux2(1).Text = RecuperaValor(CadenaSeleccion, 2) 'nombre
End Sub

Private Sub frmTIva_DatoSeleccionado(CadenaSeleccion As String)
'Tipos de Iva de la Contabilidad
    txtAux(Indice).Text = RecuperaValor(CadenaSeleccion, 1) 'codigiva
    txtAux(5).Text = RecuperaValor(CadenaSeleccion, 2) 'nombriva
    txtAux(7).Text = RecuperaValor(CadenaSeleccion, 3) 'Porcentaje iva
End Sub

Private Sub mnBuscar_Click()
    BotonBuscar
End Sub

Private Sub mnEliminar_Click()
    BotonEliminar
End Sub

Private Sub mnImprimir_Click()
    printNou
End Sub

Private Sub mnModificar_Click()
    'Comprobaciones
    If adodc1.Recordset.EOF Then Exit Sub
    If adodc1.Recordset.RecordCount < 1 Then Exit Sub
    
    ' *** repasar el nom de l'adodc, l'index del Field i el camp que te la PK ***
    'El registre de codi 0 no es pot Modificar ni Eliminar
    If EsCodigoCero(CStr(adodc1.Recordset.Fields(1).Value), FormatoCampo(txtAux(1))) Then Exit Sub
    ' ***************************************************************************
    
    
    'Prepara para modificar
    '-----------------------
    If BLOQUEADesdeFormulario2(Me, adodc1, 1) Then BotonModificar
End Sub

Private Sub mnNuevo_Click()
    BotonAnyadir
End Sub

Private Sub mnSalir_Click()
    Unload Me
End Sub

Private Sub mnVerTodos_Click()
    BotonVerTodos
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1
                mnNuevo_Click
        Case 2
                mnModificar_Click
        Case 3
                mnEliminar_Click
        Case 5
                mnBuscar_Click
        Case 6
                mnVerTodos_Click
        Case 8 'Imprimir
                mnImprimir_Click
    End Select
End Sub

Private Sub CargaGrid(Optional vSQL As String)
    Dim I As Integer
    Dim Sql As String
    Dim tots As String
    
    If vSQL <> "" Then
        Sql = CadenaConsulta & " WHERE " & vSQL  ' ### [Monica] 08/09/2006: antes habia AND
    Else
        Sql = CadenaConsulta
    End If
    '********************* canviar el ORDER BY *********************++
    'SQL = SQL & " ORDER BY codempre, codsupdt"
    Sql = Sql & " ORDER BY codconce"
    '**************************************************************++
    
    CargaGridGnral Me.DataGrid1, Me.adodc1, Sql, False
       
       
    ' *** posar només els controls del grid ***
    tots = "S|txtAux(0)|T|Código|850|;S|txtAux(1)|T|Denominación|3300|;S|txtAux(2)|T|Seccion|1000|;S|btnBuscar(3)|B|||;"
    tots = tots & "S|txtAux2(1)|T|Nombre|2150|;S|txtAux(3)|T|Cta.Ventas|1400|;S|btnBuscar(0)|B|||;"
    tots = tots & "S|txtAux(10)|T|Cta.Compras|1400|;S|btnBuscar(4)|B|||;S|txtAux(4)|T|Iva|1000|;"
    tots = tots & "S|btnBuscar(1)|B|||;S|txtAux(8)|T|C.Coste|900|;S|btnBuscar(2)|B|||;"
    
    arregla tots, DataGrid1, Me, 350
    
    DataGrid1.ScrollBars = dbgAutomatic
    ' **********************************************************
    
    ' *** alliniar les columnes que siguen numèriques a la dreta ***
    DataGrid1.Columns(2).Alignment = dbgCenter
    DataGrid1.Columns(4).Alignment = dbgCenter
    DataGrid1.Columns(5).Alignment = dbgCenter
    ' *****************************
    
    
    ' *** Si n'hi han camps fora del grid ***
    If Not adodc1.Recordset.EOF Then
        CargaForaGrid
    Else
        LimpiarCampos
    End If
    ' **************************************
End Sub

Private Sub txtaux_GotFocus(Index As Integer)
    ConseguirFoco txtAux(Index), Modo
End Sub

Private Sub txtAux_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Alvançar/Retrocedir els camps en les fleches de desplaçament del teclat.
    KEYdown KeyCode
End Sub

Private Sub txtaux_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = teclaBuscar Then
        If Modo = 1 Or Modo = 3 Or Modo = 4 Then
            Select Case Index
                Case 2: KEYBusqueda KeyAscii, 3 'codsecci
                Case 3: KEYBusqueda KeyAscii, 0 'cta contable
                Case 10: KEYBusqueda KeyAscii, 4 'cta contable proveedor
                Case 4: KEYBusqueda KeyAscii, 1 'iva
                Case 8: KEYBusqueda KeyAscii, 2 'centro de coste
            End Select
        End If
    Else
        KEYpress KeyAscii
    End If
    
End Sub

Private Sub KEYBusqueda(KeyAscii As Integer, Indice As Integer)
    KeyAscii = 0
    btnBuscar_Click (Indice)
End Sub


Private Sub txtAux_LostFocus(Index As Integer)
Dim cadMen As String

    If Not PerderFocoGnral(txtAux(Index), Modo) Then Exit Sub
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    
    '*** configurar el LostFocus dels camps (de dins i de fora del grid) ***
    Select Case Index
        Case 0
            PonerFormatoEntero txtAux(Index)
        
        Case 1
            If txtAux(Index).Text = "" Then Exit Sub
            txtAux(Index).Text = UCase(txtAux(Index).Text)
            
        Case 2 'base de datos de contabilidad
'            PonerFormatoEntero txtAux(Index)
'            If txtAux(Index).Text <> "" Then
'                txtAux2(1).Text = PonerNombreDeCod(txtAux(2), "rseccion", "nomsecci", "codsecci", "N")
'                txtAux(Index).Text = ""
'                PonerFoco txtAux(Index)
'            End If
                If PonerFormatoEntero(txtAux(Index)) Then
                    If Not vSeccion Is Nothing Then Set vSeccion = Nothing
                
                    Set vSeccion = New CSeccion
                    If vSeccion.LeerDatos(txtAux(Index)) Then
                        txtAux2(1).Text = vSeccion.Nombre
                        If vSeccion.AbrirConta Then
                        
                        End If
                    Else
                        Set vSeccion = Nothing
                        cadMen = "No existe la Sección: " & txtAux(Index).Text & vbCrLf
                        cadMen = cadMen & "¿Desea crearla?" & vbCrLf
                        If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                            Set frmSec = New frmManSeccion
                            frmSec.DatosADevolverBusqueda = "0|1|"
                            frmSec.NuevoCodigo = txtAux(Index).Text
                            txtAux(Index).Text = ""
                            TerminaBloquear
                            frmSec.Show vbModal
                            Set frmSec = Nothing
                            If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Me.adodc1, 1
                        Else
                            txtAux(Index).Text = ""
                        End If
                    End If
                Else
                    txtAux(Index).Text = ""
                End If
        
        Case 3 'cuenta contable
            If txtAux(Index).Text = "" Then Exit Sub
            
            If Not vSeccion Is Nothing Then
                txtAux(6).Text = PonerNombreCuenta(txtAux(Index), Modo)
                If txtAux(6).Text = "" Then PonerFoco txtAux(Index)
            End If
            
            
        Case 10 'cuenta contable de compras
            If txtAux(Index).Text = "" Then Exit Sub
            
            If Not vSeccion Is Nothing Then
                txtAux(11).Text = PonerNombreCuenta(txtAux(Index), Modo)
                If txtAux(11).Text = "" Then PonerFoco txtAux(Index)
            End If
            
'            If AbrirConexionContaFac(vParamAplic.UsuarioContaFac, vParamAplic.PasswordContaFac, txtAux(2).Text) Then
'                Set vEmpresaFac = New CempresaFac
'                If vEmpresaFac.LeerNiveles Then
'                    txtAux(6) = PonerNombreCuenta(txtAux(3), Modo, , CByte(txtAux(2).Text), True)
'                    'DevuelveDesdeBDNewFac("cuentas", "nommacta", "codmacta", txtAux(3).Text, "T")
'                End If
'                Set vEmpresaFac = Nothing
'                CerrarConexionContaFac
'            End If
        
        Case 4 'tipo de iva
            If txtAux(Index).Text = "" Then Exit Sub
            PonerFormatoEntero txtAux(Index)
            If txtAux(2).Text <> "" Then
                If Not vSeccion Is Nothing Then
                     txtAux(5) = DevuelveDesdeBDNew(cConta, "tiposiva", "nombriva", "codigiva", txtAux(4).Text, "N")
                     txtAux(7) = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", txtAux(4).Text, "N")
                End If
            End If
'                If AbrirConexionContaFac(vParamAplic.UsuarioContaFac, vParamAplic.PasswordContaFac, txtAux(2).Text) Then
'                    Set vEmpresaFac = New CempresaFac
'                    If vEmpresaFac.LeerNiveles Then
'                        txtAux(5) = DevuelveDesdeBDNewFac("tiposiva", "nombriva", "codigiva", txtAux(4).Text, "N")
'                        txtAux(7) = DevuelveDesdeBDNewFac("tiposiva", "porceiva", "codigiva", txtAux(4).Text, "N")
'                    End If
'                    Set vEmpresaFac = Nothing
'                    CerrarConexionContaFac
'                End If
            
            
       Case 8 ' centro de coste
            If txtAux(Index).Text = "" Then Exit Sub
            txtAux(Index).Text = UCase(txtAux(Index).Text)
            If txtAux(2).Text <> "" Then
                If Not vSeccion Is Nothing Then
                     txtAux(9) = DevuelveDesdeBDNew(cConta, "cabccost", "nomccost", "codccost", txtAux(8).Text, "T")
                     If txtAux(9).Text = "" Then
                          MsgBox "No existe el centro de coste. Reintroduzca.", vbExclamation
                          PonerFoco txtAux(8)
                     End If
                End If
            End If


'                If AbrirConexionContaFac(vParamAplic.UsuarioContaFac, vParamAplic.PasswordContaFac, txtAux(2).Text) Then
'                    Set vEmpresaFac = New CempresaFac
'                    If vEmpresaFac.LeerNiveles Then
'                         txtAux(9) = DevuelveDesdeBDNewFac("cabccost", "nomccost", "codccost", txtAux(8).Text, "T")
'                         If txtAux(9).Text = "" Then
'                            MsgBox "No existe el centro de coste. Reintroduzca.", vbExclamation
'                            PonerFoco txtAux(8)
'                         End If
'                    End If
'                    Set vEmpresaFac = Nothing
'                    CerrarConexionContaFac
'                End If
       
    End Select
    '**************************************************************************
End Sub


Private Function DatosOK() As Boolean
Dim Datos As String
Dim B As Boolean
' *** només per ad este manteniment ***
Dim Rs As Recordset
Dim cad As String
'Dim exped As String
' *************************************

    B = CompForm(Me)
    If Not B Then Exit Function


    If B And (Modo = 3) Then
        'Estem insertant
        'aço es com posar: select codvarie from svarie where codvarie = txtAux(0)
        'la N es pa dir que es numèric
         
        ' *** canviar els arguments de la funcio, el mensage i repasar si n'hi ha codEmpre ***
        Datos = DevuelveDesdeBD("codconce", "fvarconce", "codconce", txtAux(0).Text, "N")
'       Datos = DevuelveDesdeBDNew(1, "sdexpgrp", "codsupdt", "codsupdt", txtAux(1).Text, "N", "", "codempre", CStr(vSesion.Empresa), "N")
         
        If Datos <> "" Then
            MsgBox "Ya existe el Código de Concepto: " & txtAux(0).Text, vbExclamation
            B = False
            PonerFoco txtAux(1) '*** posar el foco al 1r camp visible de la PK de la capçalera ***
            Exit Function
        End If
        '*************************************************************************************
    End If

    ' *** Si cal fer atres comprovacions ***
    'comprobamos que la cta contable sea de gastos
    Dim CtaGasto As String
'    If AbrirConexionContaFac(vParamAplic.UsuarioContaFac, vParamAplic.PasswordContaFac, txtAux(2).Text) Then
'        Set vEmpresaFac = New CempresaFac
'        If vEmpresaFac.LeerNiveles Then
''[Monica]13/12/2011: quitamos el control de que la cuenta contable sea una cuenta de ventas
''            CtaGasto = DevuelveDesdeBDNewFac("parametros", "grupovta", "", "", "T")
''            If CtaGasto <> Mid(txtAux(3).Text, 1, 1) Then
''                MsgBox "La cuenta introducida no es de Ventas. Reintroduzca.", vbExclamation
''                b = False
''                PonerFoco txtAux(3)
''            End If
'            If b And vEmpresaFac.TieneAnalitica Then
'                If txtAux(8).Text = "" Then
'                    MsgBox "La Contabilidad tiene Analítica debe introducir el Centro de Coste.", vbExclamation
'                    b = False
'                    PonerFoco txtAux(8)
'                Else
'
'                End If
'            End If
'        End If
'        Set vEmpresaFac = Nothing
'        CerrarConexionContaFac
'    Else
'        b = False
'        txtAux(2).Text = ""
'        PonerFoco txtAux(2)
'    End If
'
    
        
    ' *********************************************

    DatosOK = B
End Function


Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
    PonerOpcionesMenuGeneralNew Me
End Sub


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me

End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Dim I As Byte

    If Modo <> 3 And Modo <> 4 Then 'Modificar
        CargaForaGrid
    Else
        For I = 0 To txtAux.Count - 1
            txtAux(I).Text = ""
        Next I
    End If
    
    PonerContRegIndicador
    
End Sub

Private Sub CargaForaGrid()
        If adodc1.Recordset.EOF Then Exit Sub
        
        ' *** posar als camps de fora del grid el valor de la columna corresponent ***
        Set vSeccion = New CSeccion
        If vSeccion.LeerDatos(Me.adodc1.Recordset.Fields(2).Value) Then
            If vSeccion.AbrirConta Then
                txtAux(6).Text = DevuelveDesdeBDNew(cConta, "cuentas", "nommacta", "codmacta", Me.adodc1.Recordset.Fields(4).Value, "T")
                txtAux(11).Text = DevuelveDesdeBDNew(cConta, "cuentas", "nommacta", "codmacta", Me.adodc1.Recordset.Fields(5).Value, "T")
                txtAux(5).Text = DevuelveDesdeBDNew(cConta, "tiposiva", "nombriva", "codigiva", Me.adodc1.Recordset.Fields(6).Value, "N")
                txtAux(7).Text = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", Me.adodc1.Recordset.Fields(6).Value, "N")
                txtAux(9).Text = ""
                If DBLet(Me.adodc1.Recordset.Fields(7).Value, "T") <> "" Then
                    txtAux(9).Text = DevuelveDesdeBDNew(cConta, "cabccost", "nomccost", "codccost", Me.adodc1.Recordset.Fields(6).Value, "T")
                End If
            
                vSeccion.CerrarConta
            End If
        End If
        Set vSeccion = Nothing
            
        
        ' **********************************************************************
 End Sub

Private Sub LimpiarCampos()
Dim I As Integer
On Error Resume Next

    ' *** posar a huit tots els camps de fora del grid ***
    For I = 5 To 7
        txtAux(I).Text = ""
    Next I
    txtAux(11).Text = ""
    ' ****************************************************
'    text2(12).Text = "" ' el nombre de la cuenta contable la ponemos a cero

    If Err.Number <> 0 Then Err.Clear
End Sub

Private Sub PonerContRegIndicador()
'si estamos en modo ver registros muestra el numero de registro en el que estamos
'situados del total de registros mostrados: 1 de 24
Dim cadReg As String

    If (Modo = 2 Or Modo = 0) Then
        cadReg = PonerContRegistros(Me.adodc1)
        If CadB = "" Then
            lblIndicador.Caption = cadReg
        Else
            lblIndicador.Caption = "BUSQUEDA: " & cadReg
        End If
    End If
End Sub

Private Sub printNou()
    With frmImprimir2
        .cadTabla2 = "fvarconce"
        .Informe2 = "rFVARConceptos.rpt"
        If CadB <> "" Then
            '.cadRegSelec = Replace(SQL2SF(CadB), "clientes", "clientes_1")
            .cadRegSelec = SQL2SF(CadB)
        Else
            .cadRegSelec = ""
        End If
        ' *** repasar el nom de l'adodc ***
        '.cadRegActua = Replace(POS2SF(Data1, Me), "clientes", "clientes_1")
        .cadRegActua = POS2SF(adodc1, Me)
        ' *** repasar codEmpre ***
        .cadTodosReg = ""
        '.cadTodosReg = "{sbanco.codbanpr} = " & vSesion.Empresa
        ' *** repasar si li pose ordre o no ****
        .OtrosParametros2 = "pEmpresa='" & vEmpresa.nomempre & "'|pOrden={fvarconce.codconce}|"
        '.OtrosParametros2 = "pEmpresa='" & vEmpresa.NomEmpre & "'|"
        ' *** posar el nº de paràmetres que he posat en OtrosParametros2 ***
        '.NumeroParametros2 = 1
        .NumeroParametros2 = 2
        ' ******************************************************************
        .MostrarTree2 = False
        .InfConta2 = False
        .ConSubInforme2 = False

        .Show vbModal
    End With
End Sub

'' ### [DavidV] 26/04/2006: Activar/desactivar la rueda del ratón.
'Private Sub DataGrid1_GotFocus()
'  WheelHook DataGrid1
'End Sub
'Private Sub DataGrid1_Lostfocus()
'  WheelUnHook
'End Sub

