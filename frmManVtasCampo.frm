VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmManVtasCampo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ventas Campo"
   ClientHeight    =   9825
   ClientLeft      =   45
   ClientTop       =   30
   ClientWidth     =   14775
   Icon            =   "frmManVtasCampo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9825
   ScaleWidth      =   14775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameBotonGnral 
      Height          =   705
      Left            =   90
      TabIndex        =   28
      Top             =   0
      Width           =   3585
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   210
         TabIndex        =   29
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
   Begin VB.Frame FrameBotonGnral2 
      Height          =   705
      Left            =   3750
      TabIndex        =   26
      Top             =   0
      Width           =   2595
      Begin MSComctlLib.Toolbar Toolbar2 
         Height          =   330
         Left            =   210
         TabIndex        =   27
         Top             =   180
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   4
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Generar Facturas"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Deshacer Facturaci�n"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Anticipos sin Entradas"
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Rec�lculo Importe"
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
      Left            =   10560
      TabIndex        =   25
      Top             =   210
      Visible         =   0   'False
      Width           =   1605
   End
   Begin VB.TextBox txtAux2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
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
      Height          =   330
      Index           =   2
      Left            =   6270
      TabIndex        =   24
      Top             =   4920
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.Frame Frame2 
      Caption         =   "Totales"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   4020
      TabIndex        =   19
      Top             =   9045
      Width           =   6465
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
         Index           =   3
         Left            =   990
         Locked          =   -1  'True
         TabIndex        =   21
         Text            =   "Text2"
         Top             =   210
         Width           =   1755
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
         Index           =   4
         Left            =   4080
         Locked          =   -1  'True
         TabIndex        =   20
         Text            =   "Text2"
         Top             =   210
         Width           =   1785
      End
      Begin VB.Label Label1 
         Caption         =   "IMPORTE"
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
         Left            =   3090
         TabIndex        =   23
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "K.NETO"
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
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   1335
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
      Height          =   350
      Index           =   2
      Left            =   9810
      MaskColor       =   &H00000000&
      TabIndex        =   18
      ToolTipText     =   "Buscar Fecha"
      Top             =   4920
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
      Index           =   6
      Left            =   10020
      MaxLength       =   13
      TabIndex        =   7
      Tag             =   "Imp.Entrada|N|S|||rhisfruta|impentrada|##,###,##0.00||"
      Text            =   "Imp.Entrada"
      Top             =   4920
      Width           =   915
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
      Index           =   1
      Left            =   2970
      MaxLength       =   6
      TabIndex        =   1
      Tag             =   "C�digo Socio|N|N|0|999999|rhisfruta|codsocio|000000|N|"
      Text            =   "Socio"
      Top             =   4920
      Width           =   800
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
      Height          =   350
      Index           =   1
      Left            =   3780
      MaskColor       =   &H00000000&
      TabIndex        =   17
      ToolTipText     =   "Buscar Socio"
      Top             =   4920
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.TextBox txtAux2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
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
      Height          =   330
      Index           =   1
      Left            =   3990
      TabIndex        =   16
      Top             =   4920
      Visible         =   0   'False
      Width           =   1485
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
      Index           =   4
      Left            =   8520
      MaxLength       =   7
      TabIndex        =   5
      Tag             =   "Albar�n|N|N|||rhisfruta|numalbar|0000000|S|"
      Text            =   "albaran"
      Top             =   4920
      Width           =   540
   End
   Begin VB.ComboBox Combo1 
      Appearance      =   0  'Flat
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
      ItemData        =   "frmManVtasCampo.frx":000C
      Left            =   7080
      List            =   "frmManVtasCampo.frx":000E
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Tag             =   "Recolecci�n|N|N|0|3|rhisfruta|recolect|||"
      Top             =   4920
      Width           =   735
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
      Index           =   3
      Left            =   7830
      MaxLength       =   7
      TabIndex        =   4
      Tag             =   "Kgs.Neto|N|N|||rhisfruta|kilosnet|###,##0||"
      Text            =   "kg.neto"
      Top             =   4920
      Width           =   630
   End
   Begin VB.TextBox txtAux2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
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
      Height          =   330
      Index           =   0
      Left            =   1380
      TabIndex        =   15
      Top             =   4920
      Visible         =   0   'False
      Width           =   1515
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
      Height          =   350
      Index           =   0
      Left            =   1170
      MaskColor       =   &H00000000&
      TabIndex        =   14
      ToolTipText     =   "Buscar Variedad"
      Top             =   4920
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
      Index           =   5
      Left            =   9120
      MaxLength       =   10
      TabIndex        =   6
      Tag             =   "Fecha Albaran|F|S|||rhisfruta|fecalbar|dd/mm/yyyy||"
      Text            =   "Fec.Albara"
      Top             =   4920
      Width           =   645
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
      Left            =   12405
      TabIndex        =   8
      Top             =   9180
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
      Left            =   13560
      TabIndex        =   9
      Top             =   9180
      Visible         =   0   'False
      Width           =   1095
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
      Index           =   2
      Left            =   5520
      MaxLength       =   8
      TabIndex        =   2
      Tag             =   "Campo|N|N|||rhisfruta|codcampo|00000000||"
      Text            =   "campo"
      Top             =   4920
      Width           =   780
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
      Left            =   330
      MaxLength       =   6
      TabIndex        =   0
      Tag             =   "C�digo Variedad|N|N|0|999999|rhisfruta|codvarie|000000|N|"
      Text            =   "Var"
      Top             =   4920
      Width           =   800
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmManVtasCampo.frx":0010
      Height          =   8145
      Left            =   120
      TabIndex        =   12
      Top             =   780
      Width           =   14520
      _ExtentX        =   25612
      _ExtentY        =   14367
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
      Left            =   13560
      TabIndex        =   13
      Top             =   9195
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Index           =   1
      Left            =   120
      TabIndex        =   10
      Top             =   9105
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
         TabIndex        =   11
         Top             =   180
         Width           =   2295
      End
   End
   Begin MSAdodcLib.Adodc adodc1 
      Height          =   375
      Left            =   7110
      Top             =   180
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
      Left            =   14190
      TabIndex        =   30
      Top             =   150
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
         Enabled         =   0   'False
         HelpContextID   =   2
         Shortcut        =   ^N
         Visible         =   0   'False
      End
      Begin VB.Menu mnModificar 
         Caption         =   "&Modificar"
         HelpContextID   =   2
         Shortcut        =   ^M
      End
      Begin VB.Menu mnEliminar 
         Caption         =   "&Eliminar"
         Enabled         =   0   'False
         HelpContextID   =   2
         Shortcut        =   ^E
         Visible         =   0   'False
      End
      Begin VB.Menu mnBarra2 
         Caption         =   "-"
      End
      Begin VB.Menu mnGenerarFactura 
         Caption         =   "&Generar Factura"
         Shortcut        =   ^G
      End
      Begin VB.Menu mnDeshacerFactura 
         Caption         =   "&Deshacer Factura"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnBarra3 
         Caption         =   "-"
      End
      Begin VB.Menu mnAnticipo 
         Caption         =   "&Anticipo sin Entradas"
      End
      Begin VB.Menu mnRecalculo 
         Caption         =   "&Rec�lculo Importe"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnBarra4 
         Caption         =   "-"
      End
      Begin VB.Menu mnSalir 
         Caption         =   "&Salir"
         Shortcut        =   ^S
      End
   End
End
Attribute VB_Name = "frmManVtasCampo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' +-+-+-+-+-+-+-+-+-+-+-+-
' +-+- Autor: MANOLO  +-+-
' +-+-+-+-+-+-+-+-+-+-+-+-

' **************** PER A QUE FUNCIONE EN UN ATRE MANTENIMENT ********************
' 0. Posar-li l'atribut Datasource a "adodc1" del Datagrid1. Canviar el Caption
'    del formulari
' 1. Canviar els TAGs i els Maxlength de TextAux(0) i TextAux(1)
' 2. En PonerModo(vModo) repasar els indexs del botons, per si es canvien
' 3. En la funci� BotonAnyadir() canviar la taula i el camp per a SugerirCodigoSiguienteStr
' 4. En la funci� BotonBuscar() canviar el nom de la clau primaria
' 5. En la funci� BotonEliminar() canviar la pregunta, les descripcions de la
'    variable SQL i el contingut del DELETE
' 6. En la funci� PonerLongCampos() posar els camps als que volem canviar el MaxLength quan busquem
' 7. En Form_Load() repasar la barra d'iconos (per si es vol canviar alg�n) i
'    canviar la consulta per a vore tots els registres
' 8. En Toolbar1_ButtonClick repasar els indexs de cada bot� per a que corresponguen
' 9. En la funci� CargaGrid canviar l'ORDER BY (normalment per la clau primaria);
'    canviar adem�s els noms dels camps, el format i si fa falta la cantitat;
'    repasar els index dels botons modificar i eliminar.
'    NOTA: si en Form_load ya li he posat clausula WHERE, canviar el `WHERE` de
'    `SQL = CadenaConsulta & " WHERE " & vSQL` per un `AND`
' 10. En txtAux_LostFocus canviar el mensage i el format del camp
' 11. En la funci� DatosOk() canviar els arguments de DevuelveDesdeBD i el mensage
'    en cas d'error
' 12. En la funci� SepuedeBorrar() canviar les comprovacions per a vore si es pot
'    borrar el registre
' *******************************SI N'HI HA COMBO*******************************
' 0. Comprovar que en el SQL de Form_Load() es fa�a refer�ncia a la taula del Combo
' 1. Pegar el Combo1 al  costat dels TextAux. Canviar-li el TAG
' 2. En BotonModificar() canviar el camp del Combo
' 3. En CargaCombo() canviar la consulta i els noms del camps, o posar els valor
'    a ma si no es llig de cap base de datos els valors del Combo

Option Explicit

Private Const IdPrograma = 5003

Public DatosADevolverBusqueda As String    'Tendra el n� de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)

Public ParamVariedad As String

'codi per al registe que s'afegix al cridar des d'atre formulari.
'Obrir en modo Insertar i tornar datos del registre insertat
Public NuevoCodigo As String
Public CodigoActual As String
Public DeConsulta As Boolean

Private CadenaConsulta As String
Private CadB As String

Private WithEvents frmVar As frmComVar    ' Variedades
Attribute frmVar.VB_VarHelpID = -1
Private WithEvents frmSoc As frmManSocios ' Socios
Attribute frmSoc.VB_VarHelpID = -1
Private WithEvents frmC1 As frmCal 'calendario fecha
Attribute frmC1.VB_VarHelpID = -1


' utilizado para buscar por checks
Private BuscaChekc As String

Dim Modo As Byte
'----------- MODOS --------------------------------
'   0.-  Formulari net sense cap camp ple
'   1.-  Preparant per a fer la b�squeda
'   2.-  Ja tenim registres i els anem a recorrer
'        ,podem editar-los Edici� del camp
'   3.-  Inserci� de nou registre
'   4.-  Modificar
'--------------------------------------------------
Dim PrimeraVez As Boolean
Dim Indice As Byte 'Index del text1 on es poses els datos retornats des d'atres Formularis de Mtos
Dim i As Integer
Dim cadSelGrid As String

Private Sub PonerModo(vModo)
Dim B As Boolean

    Modo = vModo
    BuscaChekc = ""
    
    B = (Modo = 2)
    If B Then
        PonerContRegIndicador lblIndicador, adodc1, CadB
    Else
        PonerIndicador lblIndicador, Modo
    End If
    
    For i = 0 To txtAux.Count - 1
        BloquearTxt txtAux(i), (Modo = 0 Or Modo = 2)
    Next i
    
    For i = 0 To txtAux.Count - 1
        txtAux(i).visible = Not B
        txtAux(i).BackColor = vbWhite
    Next i
    
    txtAux2(0).visible = Not B
    txtAux2(1).visible = Not B
    txtAux2(2).visible = Not B
    btnBuscar(0).visible = Not B
    btnBuscar(1).visible = Not B
    btnBuscar(2).visible = Not B
    Combo1(0).visible = Not B

    CmdAceptar.visible = Not B
    CmdCancelar.visible = Not B
    DataGrid1.Enabled = B
    
    'Si es regresar
    If DatosADevolverBusqueda <> "" Then cmdRegresar.visible = B
    
    PonerLongCampos
    PonerModoOpcionesMenu 'Activar/Desact botones de menu segun Modo
    PonerOpcionesMenu  'En funcion del usuario
    'Si estamos modo Modificar bloquear todo excepto el importe de anticipo
    For i = 0 To txtAux.Count - 1
        If i <> 6 Then txtAux(i).Enabled = (Modo <> 4)
    Next i
    For i = 0 To 2
        BloquearBtn Me.btnBuscar(i), (Modo = 4)
    Next i
    Combo1(0).Enabled = (Modo <> 4)
    
End Sub


Private Sub PonerModoOpcionesMenu()
'Activa/Desactiva botones del la toobar y del menu, segun el modo en que estemos
Dim B As Boolean

    B = (Modo = 2)
    'Busqueda
    Toolbar1.Buttons(5).Enabled = B
    Me.mnBuscar.Enabled = B
    'Ver Todos
    Toolbar1.Buttons(6).Enabled = B
    Me.mnVerTodos.Enabled = B
    
    
    'Factura de anticipo sin entrada
    Toolbar2.Buttons(3).Enabled = B
    Me.mnAnticipo.Enabled = B Or (Modo = 0)
    'recalculo de importe
    Toolbar2.Buttons(4).Enabled = B
    Me.mnRecalculo.Enabled = B Or (Modo = 0)
    
    
    
    B = (B And adodc1.Recordset.RecordCount > 0) And Not DeConsulta
    'Modificar
    Toolbar1.Buttons(2).Enabled = B
    Me.mnModificar.Enabled = B
    
    'Eliminar
    Toolbar1.Buttons(3).Enabled = False 'B
    Me.mnEliminar.Enabled = False 'B
    'Insertar
    Toolbar1.Buttons(1).Enabled = False ' B And Not DeConsulta
    Me.mnNuevo.Enabled = False 'B And Not DeConsulta
    'Impresion
    Toolbar1.Buttons(8).Enabled = False ' B And Not DeConsulta
    
    'Generar Factura
    Toolbar2.Buttons(1).Enabled = B
    Me.mnGenerarFactura.Enabled = B
    'Deshacer Factura
    Toolbar2.Buttons(2).Enabled = B
    Me.mnDeshacerFactura.Enabled = B
    
    
End Sub

Private Sub BotonAnyadir()
    Dim NumF As String
    Dim anc As Single
    
'    CargaGrid 'primer de tot carregue tot el grid
'    CadB = ""
'    '******************** canviar taula i camp **************************
'    If (DatosADevolverBusqueda <> "") And NuevoCodigo <> "" Then
'        NumF = NuevoCodigo
'    Else
'        NumF = SugerirCodigoSiguienteStr("productos", "codprodu")
'    End If
'    '********************************************************************
    'Situamos el grid al final
    AnyadirLinea DataGrid1, adodc1
         
    anc = DataGrid1.Top
    If DataGrid1.Row < 0 Then
        anc = anc + 240
    Else
        anc = anc + DataGrid1.RowTop(DataGrid1.Row) + 5
    End If
    txtAux(0).Text = NumF
    FormateaCampo txtAux(0)
    For i = 1 To txtAux.Count - 1
        txtAux(i).Text = ""
    Next i
    txtAux2(2).Text = ""
    Combo1(0).ListIndex = -1

    LLamaLineas anc, 3 'Pone el form en Modo=3, Insertar
       
    'Ponemos el foco
    PonerFoco txtAux(0)
End Sub

Private Sub BotonVerTodos()
    CadB = ""
    CargaGrid ""
    PonerModo 2
End Sub

Private Sub BotonBuscar()
    ' ***************** canviar per la clau primaria ********
    CargaGrid "rhisfruta.numalbar = -1"
    '*******************************************************************************
    'Buscar
    For i = 0 To txtAux.Count - 1
        txtAux(i).Text = ""
    Next i
    txtAux2(0).Text = ""
    txtAux2(1).Text = ""
    txtAux2(2).Text = ""
    Combo1(0).ListIndex = -1
    LLamaLineas DataGrid1.Top + 240, 1 'Pone el form en Modo=1, Buscar
    PonerFoco txtAux(0)
End Sub

Private Sub BotonModificar()
    Dim anc As Single
    Dim i As Integer
    
    Screen.MousePointer = vbHourglass
    
    If DataGrid1.Bookmark < DataGrid1.FirstRow Or DataGrid1.Bookmark > (DataGrid1.FirstRow + DataGrid1.VisibleRows - 1) Then
        i = DataGrid1.Bookmark - DataGrid1.FirstRow
        DataGrid1.Scroll 0, i
        DataGrid1.Refresh
    End If
    
    If DataGrid1.Row < 0 Then
        anc = 320
    Else
        anc = DataGrid1.RowTop(DataGrid1.Row) + DataGrid1.Top '545 '1025 '545
    End If

    'Llamamos al form
    txtAux(0).Text = DataGrid1.Columns(0).Text
    txtAux2(0).Text = DataGrid1.Columns(1).Text
    txtAux(1).Text = DataGrid1.Columns(2).Text
    txtAux2(1).Text = DataGrid1.Columns(3).Text
    txtAux(2).Text = DataGrid1.Columns(4).Text
    txtAux2(2).Text = DataGrid1.Columns(5).Text
    
    txtAux(3).Text = DataGrid1.Columns(8).Text
    txtAux(4).Text = DataGrid1.Columns(9).Text
    txtAux(5).Text = DataGrid1.Columns(10).Text
    txtAux(6).Text = DataGrid1.Columns(11).Text
    
    ' ***** canviar-ho pel nom del camp del combo *********
    i = adodc1.Recordset!Recolect
    ' *****************************************************
    PosicionarCombo Me.Combo1(0), i
    
    LLamaLineas anc, 4 'Pone el form en Modo=4, Modificar
   
    'Como es modificar
    PonerFoco txtAux(6)
    Screen.MousePointer = vbDefault
End Sub

Private Sub LLamaLineas(alto As Single, xModo As Byte)
    DeseleccionaGrid Me.DataGrid1
    PonerModo xModo
    
    'Fijamos el ancho
    For i = 0 To txtAux.Count - 1
        txtAux(i).Top = alto
    Next i
    
    ' ### [Monica] 12/09/2006
    txtAux2(0).Top = alto
    txtAux2(1).Top = alto
    txtAux2(2).Top = alto
    btnBuscar(0).Top = alto - 15
    btnBuscar(1).Top = alto - 15
    btnBuscar(2).Top = alto - 15
    Combo1(0).Top = alto - 15
    
End Sub


Private Sub BotonEliminar()
Dim Sql As String
Dim temp As Boolean

    On Error GoTo Error2
    'Ciertas comprobaciones
    If adodc1.Recordset.EOF Then Exit Sub
'    If Not SepuedeBorrar Then Exit Sub
        
    ' ### [Monica] 26/09/2006 dejamos modificar y eliminar el codigo 0
    ' *** repasar el nom de l'adodc, l'index del Field i el camp que te la PK ***
    'El registre de codi 0 no es pot Modificar ni Eliminar
    'If EsCodigoCero(CStr(adodc1.Recordset.Fields(0).Value), FormatoCampo(txtAux(0))) Then Exit Sub
    ' ***************************************************************************
    
    '*************** canviar els noms i el DELETE **********************************
    Sql = "�Seguro que desea eliminar la Calidad?"
    Sql = Sql & vbCrLf & "Variedad: " & adodc1.Recordset.Fields(0) & " " & adodc1.Recordset.Fields(1)
    Sql = Sql & vbCrLf & "C�digo: " & adodc1.Recordset.Fields(2)
    Sql = Sql & vbCrLf & "Descripci�n: " & adodc1.Recordset.Fields(3)
    
    If MsgBox(Sql, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        NumRegElim = adodc1.Recordset.AbsolutePosition
        Sql = "Delete from rcalidad where codvarie=" & adodc1.Recordset!Codvarie
        Sql = Sql & " and codcalid = " & adodc1.Recordset!codcalid
        conn.Execute Sql
        CargaGrid CadB
'        If CadB <> "" Then
'            CargaGrid CadB
'            lblIndicador.Caption = "BUSQUEDA: " & PonerContRegistros(Me.adodc1)
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
'Modificar el MaxLength del campo en funcion de si es modo de b�squeda o no
'para los campos que permitan introducir criterios m�s largos del tama�o del campo
'Si maxlength=6 para codprove, en modo busqueda Maxlenth=13 'el doble + 1
    
    'para los txtAux
    PonerLongCamposGnral Me, Modo, 3
End Sub

Private Sub btnBuscar_Click(Index As Integer)
 TerminaBloquear
    
    Select Case Index
        Case 0 'variedades de comercial
            Indice = Index
            Set frmVar = New frmComVar
            frmVar.DatosADevolverBusqueda = "0|1|"
            frmVar.CodigoActual = txtAux(Indice).Text
            frmVar.Show vbModal
            Set frmVar = Nothing
            PonerFoco txtAux(Indice)
        
        Case 1 'socios
            Indice = Index
            Set frmSoc = New frmManSocios
            frmSoc.DatosADevolverBusqueda = "0|1|"
            frmSoc.Show vbModal
            Set frmSoc = Nothing
            PonerFoco txtAux(Indice)
    
        Case 2 ' fecha
           Screen.MousePointer = vbHourglass
           
           Dim esq As Long
           Dim dalt As Long
           Dim menu As Long
           Dim obj As Object
        
           Set frmC1 = New frmCal
            
           esq = btnBuscar(Index).Left
           dalt = btnBuscar(Index).Top
            
           Set obj = btnBuscar(Index).Container
        
           While btnBuscar(Index).Parent.Name <> obj.Name
                esq = esq + obj.Left
                dalt = dalt + obj.Top
                Set obj = obj.Container
           Wend
            
           menu = Me.Height - Me.ScaleHeight 'ac� tinc el heigth del men� i de la toolbar
        
           frmC1.Left = esq + btnBuscar(Index).Parent.Left + 30
           frmC1.Top = dalt + btnBuscar(Index).Parent.Top + btnBuscar(Index).Height + menu - 40
        
           
           frmC1.NovaData = Now
           Select Case Index
                Case 2
                    Indice = 5
           End Select
           
           Me.btnBuscar(2).Tag = Indice
           
           PonerFormatoFecha txtAux(Indice)
           If txtAux(Indice).Text <> "" Then frmC1.NovaData = CDate(txtAux(Indice).Text)
        
           Screen.MousePointer = vbDefault
           frmC1.Show vbModal
           Set frmC1 = Nothing
           PonerFoco txtAux(Indice)
          'Para la fecha de la navegacion
           If Index = 5 And txtAux(5).Text <> "" Then
                btnBuscar(2).Tag = txtAux(5).Text
           End If
            
        
    End Select
    
    If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Me.adodc1, 1
End Sub

Private Sub frmC1_Selec(vFecha As Date)
Dim Indice As Byte
    
    Indice = CByte(Me.btnBuscar(2).Tag)
    txtAux(Indice).Text = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub imgFec_Click(Index As Integer)
       
End Sub
Private Sub chkAux_Click(Index As Integer)
    If Modo = 1 Then
        'Buscqueda
        If InStr(1, BuscaChekc, "chkAux(" & Index & ")") = 0 Then BuscaChekc = BuscaChekc & "chkAux(" & Index & ")|"
    End If
End Sub

Private Sub chkAux_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub cmdAceptar_Click()
    Dim i As Long

    Select Case Modo
        Case 1 'BUSQUEDA
            CadB = ObtenerBusqueda(Me, BuscaChekc)
            If CadB <> "" Then
                CargaGrid CadB
                PonerModo 2
'                lblIndicador.Caption = "BUSQUEDA: " & PonerContRegistros(Me.adodc1)
                PonerFocoGrid Me.DataGrid1
                '[Monica]23/09/2011: solo para Picassent, si hay anticipos sin entradas mostrar un aviso de las facturas en
                '                    donde aparecen
                If vParamAplic.Cooperativa = 2 Or vParamAplic.Cooperativa = 16 Then
                    MostrarFacturasAnticiposSinKilos CadB
                End If
                
            End If
            
        Case 3 'INSERTAR
            If DatosOK Then
                If InsertarDesdeForm(Me) Then
                    CargaGrid CadB
                    If (DatosADevolverBusqueda <> "") And NuevoCodigo <> "" Then
                        cmdCancelar_Click
'                        If Not adodc1.Recordset.EOF Then adodc1.Recordset.MoveLast
                        If Not adodc1.Recordset.EOF Then
                            SituarDataMULTI adodc1, "codvarie = " & txtAux(0) & " and codcalid = " & txtAux(1), "" ' Find (adodc1.Recordset.Fields(2).Name & " =" & NuevoCodigo)
                        End If
                        cmdRegresar_Click
                    Else
                        BotonAnyadir
                    End If
                    CadB = ""
                End If
            End If
            
        Case 4 'MODIFICAR
            If DatosOK Then
                If ModificaDesdeFormulario(Me) Then
                    TerminaBloquear
                    i = adodc1.Recordset.Fields(9)
                    PonerModo 2
                    CargaGrid CadB
'                    If CadB <> "" Then
'                        CargaGrid CadB
'                        lblIndicador.Caption = "BUSQUEDA: " & PonerContRegistros(Me.adodc1)
'                    Else
'                        CargaGrid
'                        lblIndicador.Caption = ""
'                    End If
                    adodc1.Recordset.Find (adodc1.Recordset.Fields(9).Name & " =" & i)
                    PonerFocoGrid Me.DataGrid1
                End If
            End If
    End Select
End Sub

Private Sub cmdCancelar_Click()
    On Error Resume Next
    
    Select Case Modo
        Case 1 'b�squeda
            CargaGrid CadB
        Case 3 'insertar
            DataGrid1.AllowAddNew = False
            'CargaGrid
            If Not adodc1.Recordset.EOF Then adodc1.Recordset.MoveFirst
        Case 4 'modificar
            TerminaBloquear
    End Select
    
    PonerModo 2
    
'    If CadB <> "" Then
'        lblIndicador.Caption = "BUSQUEDA: " & PonerContRegistros(Me.adodc1)
''    Else
''        lblIndicador.Caption = ""
'    End If
    
    PonerFocoGrid Me.DataGrid1
    If Err.Number <> 0 Then Err.Clear
End Sub

Private Sub cmdRegresar_Click()
Dim cad As String
Dim i As Integer
Dim J As Integer
Dim Aux As String

    If adodc1.Recordset.EOF Then
        MsgBox "Ning�n registro devuelto.", vbExclamation
        Exit Sub
    End If
    cad = ""
    i = 0
    Do
        J = i + 1
        i = InStr(J, DatosADevolverBusqueda, "|")
        If i > 0 Then
            Aux = Mid(DatosADevolverBusqueda, J, i - J)
            J = Val(Aux)
            cad = cad & adodc1.Recordset.Fields(J) & "|"
        End If
    Loop Until i = 0
    RaiseEvent DatoSeleccionado(cad)
    Unload Me
End Sub

Private Sub Combo1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Combo1_GotFocus(Index As Integer)
    If Modo = 1 Then Combo1(Index).BackColor = vbLightBlue
End Sub

Private Sub Combo1_LostFocus(Index As Integer)
    If Combo1(Index).BackColor = vbLightBlue Then Combo1(Index).BackColor = vbWhite
End Sub


Private Sub DataGrid1_DblClick()
    If cmdRegresar.visible Then cmdRegresar_Click
End Sub

Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If Modo = 2 Then PonerContRegIndicador lblIndicador, adodc1, CadB
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault

    If PrimeraVez Then
        PrimeraVez = False
        If (DatosADevolverBusqueda <> "") And NuevoCodigo <> "" Then
            BotonAnyadir
        Else
            PonerModo 2
             If Me.CodigoActual <> "" Then
                SituarData Me.adodc1, "codvarie=" & CodigoActual, "", True
            End If
        End If
    End If
End Sub

Private Sub Form_Load()

    'Icono del formulario
    Me.Icon = frmPpal.Icon

    PrimeraVez = True

    With Me.Toolbar1
        .HotImageList = frmPpal.imgListComun_OM
        .DisabledImageList = frmPpal.imgListComun_BN
        .ImageList = frmPpal.imgListComun
        'el 1 es separadors
        .Buttons(5).Image = 1   'Buscar
        .Buttons(6).Image = 2   'Todos
        'el 4 i el 5 son separadors
        .Buttons(1).Image = 3   'Insertar
        .Buttons(2).Image = 4   'Modificar
        .Buttons(3).Image = 5   'Borrar
        
        .Buttons(8).Image = 10  'Imprimir
    End With

    With Me.Toolbar2
        .HotImageList = frmPpal.imgListComun_OM
        .DisabledImageList = frmPpal.imgListComun_BN
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 26  'generar factura
        .Buttons(2).Image = 32  ' deshacer facturacion
        .Buttons(3).Image = 25 ' generar factura de anticipo sin entradas
        .Buttons(4).Image = 31 ' reparto de importe segun kilos
    End With

    ' La Ayuda
    With Me.ToolbarAyuda
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 12
    End With


    '## A mano
'    chkVistaPrevia.Value = CheckValueLeer(Name)
    
    CargaCombo
    
    '****************** canviar la consulta ************************************************
    CadenaConsulta = "SELECT rhisfruta.codvarie, variedades.nomvarie, rhisfruta.codsocio, "
    CadenaConsulta = CadenaConsulta & " rsocios.nomsocio, rhisfruta.codcampo, rcampos.nrocampo, rhisfruta.recolect, "
    CadenaConsulta = CadenaConsulta & " CASE rhisfruta.recolect WHEN 0 THEN ""Cooperativa"" WHEN 1 THEN ""Socio"" END, "
    CadenaConsulta = CadenaConsulta & " rhisfruta.kilosnet, rhisfruta.numalbar, rhisfruta.fecalbar, rhisfruta.impentrada "
    CadenaConsulta = CadenaConsulta & " FROM rhisfruta, variedades, rsocios, rcampos "
    CadenaConsulta = CadenaConsulta & " WHERE rhisfruta.codvarie = variedades.codvarie "
    CadenaConsulta = CadenaConsulta & " and rhisfruta.codsocio = rsocios.codsocio "
    CadenaConsulta = CadenaConsulta & " and rhisfruta.tipoentr = 1 "
    CadenaConsulta = CadenaConsulta & " and rhisfruta.codcampo = rcampos.codcampo "
    '***************************************************************************************
    
    
    CadB = "numalbar = -1 "
    CargaGrid CadB
    
'    If (DatosADevolverBusqueda <> "") And NuevoCodigo <> "" Then
'        BotonAnyadir
'    Else
'        PonerModo 2
'    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
'    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
    Screen.MousePointer = vbDefault
    If Modo = 4 Then TerminaBloquear
End Sub


Private Sub frmSoc_DatoSeleccionado(CadenaSeleccion As String)
    txtAux(Indice).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000")
    txtAux2(Indice).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmVar_DatoSeleccionado(CadenaSeleccion As String)
'Variedad comercial
    txtAux(Indice).Text = RecuperaValor(CadenaSeleccion, 1) 'codvarie
    txtAux2(Indice).Text = RecuperaValor(CadenaSeleccion, 2)  'nombre variedad
End Sub

Private Sub mnAnticipo_Click()
    AbrirListadoAnticipos (16)
End Sub

Private Sub mnBuscar_Click()
    BotonBuscar
End Sub

Private Sub mnDeshacerFactura_Click()
    AbrirListadoAnticipos (7)
End Sub

Private Sub mnEliminar_Click()
    BotonEliminar
End Sub

Private Sub mnGenerarFactura_Click()
Dim Sql As String

    Sql = CadB
    AbrirListadoAnticipos (6)
    CargaGrid Sql
    
End Sub

Private Sub mnModificar_Click()
    'Comprobaciones
    '--------------
    If adodc1.Recordset.EOF Then Exit Sub
    
    If adodc1.Recordset.RecordCount < 1 Then Exit Sub
    
    ' ### [Monica] 26/09/2006 dejamos modificar y eliminar el codigo 0
    ' *** repasar el nom de l'adodc, l'index del Field i el camp que te la PK ***
    'El registre de codi 0 no es pot Modificar ni Eliminar
    'If EsCodigoCero(CStr(adodc1.Recordset.Fields(0).Value), FormatoCampo(txtAux(0))) Then Exit Sub
    
    
    'Preparamos para modificar
    '-------------------------
    If BLOQUEADesdeFormulario2(Me, adodc1, 1) Then BotonModificar
End Sub

Private Sub mnNuevo_Click()
    BotonAnyadir
End Sub

Private Sub mnRecalculo_Click()
    frmListAnticipos.OpcionListado = 17
    frmListAnticipos.Show vbModal
    CargaGrid CadB
End Sub

Private Sub mnSalir_Click()
    Unload Me
End Sub

Private Sub mnVerTodos_Click()
    BotonVerTodos
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 5
                mnBuscar_Click
        Case 6
                mnVerTodos_Click
        Case 1
                mnNuevo_Click
        Case 2
                mnModificar_Click
        Case 3
                mnEliminar_Click
    End Select
End Sub




Private Sub CargaGrid(Optional vSQL As String)
    Dim Sql As String
    Dim tots As String
    
'    adodc1.ConnectionString = Conn
    If vSQL <> "" Then
        Sql = CadenaConsulta & " AND " & vSQL
    Else
        Sql = CadenaConsulta
    End If
    
    cadSelGrid = vSQL
    'If vSQL <> "" Then cadSelGrid = cadSelGrid & vSQL
    
'    If ParamVariedad <> "" Then SQL = SQL & " and rcalidad.codvarie = " & ParamVariedad
    
    '********************* canviar el ORDER BY *********************++
    Sql = Sql & " ORDER BY rhisfruta.codvarie, rhisfruta.codsocio, rhisfruta.codcampo, rhisfruta.numalbar"
    '**************************************************************++
    
    CargaGridGnral Me.DataGrid1, Me.adodc1, Sql, PrimeraVez
    
    ' *******************canviar els noms i si fa falta la cantitat********************
    tots = "S|txtAux(0)|T|C�digo|1000|;S|btnBuscar(0)|B|||;S|txtAux2(0)|T|Variedad|1500|;"
    tots = tots & "S|txtAux(1)|T|Socio|1000|;S|btnBuscar(1)|B|||;S|txtAux2(1)|T|Nombre|2300|;"
    tots = tots & "S|txtAux(2)|T|Campo|1100|;S|txtAux2(2)|T|Orden|700|;"
    tots = tots & "N||||0|;S|Combo1(0)|C|Tipo|1320|;S|txtAux(3)|T|Peso Neto|1200|;S|txtAux(4)|T|Albar�n|1000|;"
    tots = tots & "S|txtAux(5)|T|Fecha|1400|;S|btnBuscar(2)|B|||;S|txtAux(6)|T|Imp.Entrada|1400|;"
    
    arregla tots, DataGrid1, Me, 350
    
    DataGrid1.ScrollBars = dbgAutomatic
    DataGrid1.Columns(0).Alignment = dbgLeft
    DataGrid1.Columns(2).Alignment = dbgLeft
    DataGrid1.Columns(4).Alignment = dbgLeft
    
    CalcularTotales
End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1 ' generar factura
                mnGenerarFactura_Click
        Case 2 ' dehacer facturacion
                mnDeshacerFactura_Click
        Case 3 ' generacion de factura de anticipo sin entradas
                mnAnticipo_Click
        Case 4 ' recalculo de importe
                mnRecalculo_Click
    End Select
End Sub

Private Sub ToolbarAyuda_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1
            LanzaVisorMimeDocumento Me.hWnd, DireccionAyuda & IdPrograma & ".html"
    End Select
End Sub

Private Sub txtaux_GotFocus(Index As Integer)
    ConseguirFoco txtAux(Index), Modo
End Sub


Private Sub txtAux_LostFocus(Index As Integer)
    If Not PerderFocoGnral(txtAux(Index), Modo) Then Exit Sub
    
    Select Case Index
        Case 0 'codigo de variedad
            If txtAux(Index).Text = "" Then Exit Sub
            txtAux2(Index).Text = PonerNombreDeCod(txtAux(Index), "variedades", "nomvarie", "codvarie", "N")
        
        Case 1 'codigo de socio
            If txtAux(Index).Text = "" Then Exit Sub
            txtAux2(Index).Text = PonerNombreDeCod(txtAux(Index), "rsocios", "nomsocio", "codsocio", "N")
            
        Case 2 ' codigo de campo
            PonerFormatoEntero txtAux(Index)
        
        Case 3 ' peso neto
            PonerFormatoEntero txtAux(Index)
        
        Case 4 ' albaran
            PonerFormatoEntero txtAux(Index)
        
        Case 5 ' fecha de albaran
            '[Monica]28/08/2013: comprobamos que la fecha est� en la campa�a
            PonerFormatoFecha txtAux(Index), True
        
        Case 6 ' importe de entrada
            PonerFormatoDecimal txtAux(Index), 3
    End Select
    
End Sub

Private Function DatosOK() As Boolean
'Dim Datos As String
Dim B As Boolean
Dim Sql As String
Dim Mens As String

    B = CompForm(Me)
    If Not B Then Exit Function
    
    If B And (Modo = 4) Then
'        SQL = "select count(*) from rcalidad where codvarie = " & DBSet(txtAux(0).Text, "N")
'        SQL = SQL & " and codcalid <> " & DBSet(txtAux(1).Text, "N")
'        SQL = SQL & " and posicion = " & DBSet(txtAux(4).Text, "N")
'
'        If TotalRegistros(SQL) <> 0 Then
'            MsgBox "La posici�n de esta calidad est� asignada a otra. Revise.", vbExclamation
'            PonerFoco txtAux(4)
'            b = False
'        End If
    End If
  
    DatosOK = B
End Function


Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
    PonerOpcionesMenuGeneralNew Me
End Sub

'' ### [DavidV] 26/04/2006: Activar/desactivar la rueda del rat�n.
'Private Sub DataGrid1_GotFocus()
'  WheelHook DataGrid1
'End Sub
'Private Sub DataGrid1_Lostfocus()
'  WheelUnHook
'End Sub

Private Sub txtaux_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = teclaBuscar Then
        If Modo = 1 Or Modo = 3 Or Modo = 4 Then
            Select Case Index
                Case 0: KEYBusqueda KeyAscii, 0 'variedad
                Case 1: KEYBusqueda KeyAscii, 1 'socio
                Case 5: KEYBusqueda KeyAscii, 2 ' fecha
            End Select
        End If
    Else
        KEYpress KeyAscii
    End If
End Sub

Private Sub txtAux_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Alvan�ar/Retrocedir els camps en les fleches de despla�ament del teclat.
    KEYdown KeyCode
End Sub

Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me

End Sub

Private Sub KEYBusqueda(KeyAscii As Integer, Indice As Integer)
    KeyAscii = 0
    btnBuscar_Click (Indice)
End Sub

Private Sub KEYFecha(KeyAscii As Integer, Indice As Integer)
    KeyAscii = 0
    imgFec_Click (Indice)
End Sub

Private Sub CargaCombo()

    On Error GoTo ErrCarga
    
    ' Tipo de recoleccion
    Combo1(0).Clear
    Combo1(0).AddItem "Cooperativa"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 0
    Combo1(0).AddItem "Socio"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 1
    
ErrCarga:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargar datos combo.", Err.Description
End Sub



Private Sub CalcularTotales()
'calcula la cantidad total y el importe total para los
'registros mostrados de cada art�culo
Dim Sql As String
Dim Rs As ADODB.Recordset
    
    On Error GoTo ErrTotales
'    If cadSelGrid = "" Then Exit Sub
    
    Sql = "SELECT sum(impentrada) as totImporte,sum(kilosnet) as totKilos from rhisfruta "
    Sql = Sql & " where rhisfruta.tipoentr = 1 "
    If cadSelGrid <> "" Then Sql = Sql & " and " & cadSelGrid

    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not Rs.EOF Then
        Text2(4).Text = DBLet(Rs!totImporte, "N")
        Text2(4).Text = Format(Text2(4).Text, FormatoImporte)
        Text2(3).Text = DBLet(Rs!TotKilos, "N")
        Text2(3).Text = Format(Text2(3).Text, "###,###,###,##0")
    End If
    
    Rs.Close
    Set Rs = Nothing
    
    Exit Sub
    
ErrTotales:
    MuestraError Err.Number, "Calcular totales.", Err.Description
End Sub


Private Sub MostrarFacturasAnticiposSinKilos(CadB As String)
Dim Sql As String
Dim Facturas As String
Dim cadWHERE As String
Dim Rs As ADODB.Recordset

    '[Monica]23/05/2018: solo tenia en cuenta las facturas de anticipos sin kilos de FAC, no las CAT. CORREGIDO
    Sql = "select distinct rfactsoc.numfactu, rfactsoc.fecfactu, rfactsoc.codtipom from rfactsoc inner join rfactsoc_variedad on rfactsoc.codtipom = rfactsoc_variedad.codtipom and  rfactsoc.numfactu = rfactsoc_variedad.numfactu and rfactsoc.fecfactu = rfactsoc_variedad.fecfactu "
    Sql = Sql & " where rfactsoc.codtipom in ('FAC','CAT') and " & Replace(Replace(CadB, "rhisfruta.codsocio", "rfactsoc.codsocio"), "rhisfruta.codvarie", "rfactsoc_variedad.codvarie")
    Sql = Sql & " and rfactsoc_variedad.kilosnet = 0"
    If TotalRegistrosConsulta(Sql) <> 0 Then
        Set Rs = New ADODB.Recordset
        Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
   
        Facturas = "Facturas de Anticipos sin Kilos Entrados: " & vbCrLf & vbCrLf
        Facturas = ""
        While Not Rs.EOF
            Facturas = Facturas & "(" & DBSet(Rs.Fields(2).Value, "T") & "," & DBSet(Rs.Fields(0).Value, "N") & "," & DBSet(Rs.Fields(1).Value, "F") & "),"
        
            Rs.MoveNext
        Wend
        
        Set Rs = Nothing
        
        cadWHERE = "(codtipom, numfactu, fecfactu) in (" & Mid(Facturas, 1, Len(Facturas) - 1) & ")"
        
        frmMensajes.cadWHERE = cadWHERE
        frmMensajes.OpcionMensaje = 31
        frmMensajes.Show vbModal
    End If

End Sub

