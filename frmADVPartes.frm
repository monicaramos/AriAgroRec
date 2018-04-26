VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmADVPartes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Entrada de Partes"
   ClientHeight    =   9855
   ClientLeft      =   45
   ClientTop       =   4035
   ClientWidth     =   14700
   Icon            =   "frmADVPartes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9855
   ScaleWidth      =   14700
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameBotonGnral 
      Height          =   705
      Left            =   90
      TabIndex        =   105
      Top             =   135
      Width           =   3585
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   210
         TabIndex        =   106
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
      Left            =   3705
      TabIndex        =   103
      Top             =   135
      Width           =   1785
      Begin MSComctlLib.Toolbar Toolbar5 
         Height          =   330
         Left            =   210
         TabIndex        =   104
         Top             =   180
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   4
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Confirmación"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Cuadrilla"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Asignación de Precios"
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Inserción de Gastos"
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame FrameDesplazamiento 
      Height          =   705
      Left            =   5565
      TabIndex        =   101
      Top             =   135
      Width           =   2415
      Begin MSComctlLib.Toolbar ToolbarDes 
         Height          =   330
         Left            =   240
         TabIndex        =   102
         Top             =   180
         Width           =   1965
         _ExtentX        =   3466
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   4
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Primero"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Anterior"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Siguiente"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Último"
               ImageIndex      =   9
            EndProperty
         EndProperty
      End
   End
   Begin VB.CheckBox chkVistaPrevia 
      Caption         =   "Vista previa"
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
      Left            =   12105
      TabIndex        =   100
      Top             =   270
      Width           =   1605
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4740
      Left            =   90
      TabIndex        =   61
      Top             =   4395
      Width           =   14460
      _ExtentX        =   25506
      _ExtentY        =   8361
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   6
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Artículos"
      TabPicture(0)   =   "frmADVPartes.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FrameAux1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Cuadrilla"
      TabPicture(1)   =   "frmADVPartes.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label1(6)"
      Tab(1).Control(1)=   "Frame3"
      Tab(1).Control(2)=   "FrameAux0"
      Tab(1).Control(3)=   "Text2(1)"
      Tab(1).ControlCount=   4
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
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
         Index           =   1
         Left            =   -63630
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   97
         Text            =   "Text2"
         Top             =   4020
         Width           =   1800
      End
      Begin VB.Frame FrameAux0 
         BorderStyle     =   0  'None
         Height          =   2805
         Left            =   -74760
         TabIndex        =   86
         Top             =   1140
         Width           =   14015
         Begin VB.TextBox txtAux1 
            Alignment       =   1  'Right Justify
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
            Height          =   315
            Index           =   6
            Left            =   9180
            MaxLength       =   10
            TabIndex        =   96
            Tag             =   "Importe|N|N|||advpartes_trabajador|importel|###,###,##0.00||"
            Text            =   "Importe"
            Top             =   2130
            Visible         =   0   'False
            Width           =   1305
         End
         Begin VB.TextBox txtAux1 
            Alignment       =   1  'Right Justify
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
            Height          =   315
            Index           =   5
            Left            =   7800
            MaxLength       =   10
            TabIndex        =   95
            Tag             =   "Precio|N|N|||advpartes_trabajador|precio|#,##0.0000||"
            Text            =   "Precio"
            Top             =   2130
            Visible         =   0   'False
            Width           =   1305
         End
         Begin VB.TextBox txtAux1 
            Alignment       =   1  'Right Justify
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
            Height          =   315
            Index           =   4
            Left            =   6390
            MaxLength       =   15
            TabIndex        =   94
            Tag             =   "Horas|N|N|||advpartes_trabajador|horas|###,##0.00||"
            Text            =   "horas"
            Top             =   2130
            Visible         =   0   'False
            Width           =   1305
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
            Height          =   300
            Index           =   1
            Left            =   2250
            MaskColor       =   &H00000000&
            TabIndex        =   99
            ToolTipText     =   "Buscar trabajador"
            Top             =   2130
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.TextBox txtAux1 
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
            Height          =   315
            Index           =   3
            Left            =   2550
            MaxLength       =   40
            TabIndex        =   92
            Text            =   "Nombre"
            Top             =   2130
            Visible         =   0   'False
            Width           =   3735
         End
         Begin VB.TextBox txtAux1 
            Alignment       =   1  'Right Justify
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
            Height          =   315
            Index           =   2
            Left            =   1650
            MaxLength       =   6
            TabIndex        =   91
            Tag             =   "Cod.Trabajador|N|N|||advpartes_trabajador|codtraba|000000||"
            Text            =   "Trabaj"
            Top             =   2130
            Visible         =   0   'False
            Width           =   555
         End
         Begin VB.TextBox txtAux1 
            Alignment       =   1  'Right Justify
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
            Height          =   315
            Index           =   1
            Left            =   1020
            MaxLength       =   6
            TabIndex        =   90
            Tag             =   "Linea|N|N|||advpartes_trabajador|numlinea|000000|S|"
            Text            =   "Linea"
            Top             =   2130
            Visible         =   0   'False
            Width           =   555
         End
         Begin VB.TextBox txtAux1 
            Alignment       =   1  'Right Justify
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
            Height          =   315
            Index           =   0
            Left            =   360
            MaxLength       =   7
            TabIndex        =   89
            Tag             =   "Num. Parte|N|N|||advpartes_trabajador|numparte|000000|S|"
            Text            =   "Parte"
            Top             =   2130
            Visible         =   0   'False
            Width           =   555
         End
         Begin MSComctlLib.Toolbar ToolAux 
            Height          =   390
            Index           =   0
            Left            =   0
            TabIndex        =   87
            Top             =   0
            Width           =   1110
            _ExtentX        =   1958
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
         Begin MSDataGridLib.DataGrid DataGrid1 
            Bindings        =   "frmADVPartes.frx":0044
            Height          =   2025
            Left            =   0
            TabIndex        =   88
            Top             =   450
            Width           =   13160
            _ExtentX        =   23204
            _ExtentY        =   3572
            _Version        =   393216
            AllowUpdate     =   0   'False
            AllowArrows     =   -1  'True
            ColumnHeaders   =   -1  'True
            HeadLines       =   1
            RowHeight       =   19
            FormatLocked    =   -1  'True
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
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
         Begin MSAdodcLib.Adodc Adoaux 
            Height          =   330
            Index           =   0
            Left            =   1215
            Top             =   0
            Visible         =   0   'False
            Width           =   1425
            _ExtentX        =   2514
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
      End
      Begin VB.Frame Frame3 
         Enabled         =   0   'False
         Height          =   675
         Left            =   -74760
         TabIndex        =   78
         Top             =   420
         Width           =   13205
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
            Index           =   34
            Left            =   10995
            MaxLength       =   9
            TabIndex        =   82
            Tag             =   "Nro.Horas|N|S|0|999999|advpartes|nrohoras|###,##0.00||"
            Top             =   210
            Width           =   1065
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
            Index           =   33
            Left            =   8295
            MaxLength       =   7
            TabIndex        =   81
            Tag             =   "Nro.Hombres|N|S|0|999999|advpartes|nrohombres|###,##0||"
            Top             =   210
            Width           =   1065
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
            Index           =   32
            Left            =   2535
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   79
            Text            =   "Text2"
            Top             =   210
            Width           =   4005
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
            Index           =   32
            Left            =   1575
            MaxLength       =   6
            TabIndex        =   80
            Tag             =   "Cuadrilla|N|S|0|999999|advpartes|codcuadrilla|000000|N|"
            Text            =   "Text1"
            Top             =   210
            Width           =   900
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000013&
            Height          =   315
            Index           =   35
            Left            =   1620
            MaxLength       =   7
            TabIndex        =   93
            Text            =   "Text1 7"
            Top             =   210
            Width           =   885
         End
         Begin VB.Label Label1 
            Caption         =   "Nro. Horas "
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
            Index           =   5
            Left            =   9810
            TabIndex        =   85
            Top             =   270
            Width           =   1320
         End
         Begin VB.Label Label1 
            Caption         =   "Nro. Hombres "
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
            Left            =   6855
            TabIndex        =   84
            Top             =   270
            Width           =   1440
         End
         Begin VB.Image imgBuscar 
            Height          =   240
            Index           =   3
            Left            =   1230
            ToolTipText     =   "Buscar Cuadrilla"
            Top             =   255
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Cuadrilla"
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
            Left            =   180
            TabIndex        =   83
            Top             =   255
            Width           =   960
         End
      End
      Begin VB.Frame FrameAux1 
         BorderStyle     =   0  'None
         Caption         =   "Artículos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   4155
         Left            =   30
         TabIndex        =   62
         Top             =   420
         Width           =   14285
         Begin VB.TextBox txtAux 
            Alignment       =   1  'Right Justify
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
            Height          =   315
            Index           =   4
            Left            =   2430
            MaxLength       =   3
            TabIndex        =   74
            Tag             =   "Almacen|N|N|||advpartes_lineas|codalmac|000||"
            Text            =   "Alm"
            Top             =   2250
            Visible         =   0   'False
            Width           =   420
         End
         Begin VB.TextBox Text2 
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
            Index           =   16
            Left            =   1755
            Locked          =   -1  'True
            MaxLength       =   60
            TabIndex        =   73
            Text            =   "Text2 wwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwqa"
            Top             =   3570
            Width           =   8430
         End
         Begin VB.TextBox txtAux 
            Alignment       =   1  'Right Justify
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
            Height          =   315
            Index           =   10
            Left            =   9630
            MaxLength       =   2
            TabIndex        =   72
            Tag             =   "CodIva|N|N|||advpartes_lineas|codigiva|00||"
            Text            =   "Codiva"
            Top             =   2250
            Visible         =   0   'False
            Width           =   555
         End
         Begin VB.TextBox txtAux 
            Alignment       =   1  'Right Justify
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
            Height          =   315
            Index           =   9
            Left            =   8820
            MaxLength       =   12
            TabIndex        =   71
            Tag             =   "Importe|N|N|||advpartes_lineas|importel|#,###,##0.00||"
            Text            =   "importe"
            Top             =   2250
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.TextBox txtAux 
            Alignment       =   1  'Right Justify
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
            Height          =   315
            Index           =   8
            Left            =   6270
            MaxLength       =   12
            TabIndex        =   70
            Tag             =   "Dosis Habitual|N|S|||advpartes_lineas|dosishab|###,##0.000||"
            Text            =   "Dosis"
            Top             =   2250
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.TextBox txtAux 
            Alignment       =   1  'Right Justify
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
            Height          =   315
            Index           =   7
            Left            =   7980
            MaxLength       =   12
            TabIndex        =   69
            Tag             =   "Precio|N|N|||advpartes_lineas|preciove|###,##0.0000||"
            Text            =   "precio"
            Top             =   2250
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.TextBox txtAux 
            Alignment       =   1  'Right Justify
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
            Height          =   315
            Index           =   3
            Left            =   1740
            MaxLength       =   12
            TabIndex        =   68
            Tag             =   "Num.Linea|N|N|||advpartes_lineas|numlinea|000|S|"
            Text            =   "Linea"
            Top             =   2250
            Visible         =   0   'False
            Width           =   600
         End
         Begin VB.TextBox txtAux 
            Alignment       =   1  'Right Justify
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
            Height          =   315
            Index           =   6
            Left            =   7110
            MaxLength       =   12
            TabIndex        =   67
            Tag             =   "Cantidad|N|N|||advpartes_lineas|cantidad|###,##0.000||"
            Text            =   "cantidad"
            Top             =   2250
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.TextBox txtAux 
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
            Height          =   315
            Index           =   5
            Left            =   3645
            MaxLength       =   16
            TabIndex        =   66
            Tag             =   "Artículo|T|N|||advpartes_lineas|codartic||N|"
            Text            =   "articulo"
            Top             =   2250
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.TextBox txtAux 
            Alignment       =   1  'Right Justify
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
            Height          =   315
            Index           =   1
            Left            =   990
            MaxLength       =   12
            TabIndex        =   65
            Tag             =   "Num.Parte|N|N|||advpartes_lineas|numparte|0000000|S|"
            Text            =   "NumParte"
            Top             =   2250
            Visible         =   0   'False
            Width           =   690
         End
         Begin VB.TextBox Text2 
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
            Height          =   315
            Index           =   0
            Left            =   4905
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   64
            Text            =   "Nombre articulo"
            Top             =   2250
            Width           =   1200
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
            Height          =   300
            Index           =   0
            Left            =   4680
            MaskColor       =   &H00000000&
            TabIndex        =   63
            ToolTipText     =   "Buscar Artículo ADV"
            Top             =   2250
            Visible         =   0   'False
            Width           =   195
         End
         Begin MSComctlLib.Toolbar ToolAux 
            Height          =   390
            Index           =   1
            Left            =   240
            TabIndex        =   75
            Top             =   75
            Width           =   1110
            _ExtentX        =   1958
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
         Begin MSDataGridLib.DataGrid DataGrid3 
            Bindings        =   "frmADVPartes.frx":0059
            Height          =   2760
            Left            =   240
            TabIndex        =   76
            Top             =   570
            Width           =   13155
            _ExtentX        =   23204
            _ExtentY        =   4868
            _Version        =   393216
            AllowUpdate     =   0   'False
            AllowArrows     =   -1  'True
            ColumnHeaders   =   -1  'True
            HeadLines       =   1
            RowHeight       =   19
            FormatLocked    =   -1  'True
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
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
         Begin MSAdodcLib.Adodc Adoaux 
            Height          =   330
            Index           =   1
            Left            =   1455
            Top             =   75
            Visible         =   0   'False
            Width           =   1425
            _ExtentX        =   2514
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
            Caption         =   "Ampliación Línea"
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
            Index           =   35
            Left            =   405
            TabIndex        =   77
            Top             =   3615
            Width           =   1335
         End
      End
      Begin VB.Label Label1 
         Caption         =   "TOTAL CUADRILLA"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   6
         Left            =   -65625
         TabIndex        =   98
         Top             =   4080
         Width           =   2010
      End
   End
   Begin VB.Frame FrameFactura 
      Caption         =   "Totales"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   3510
      Left            =   7425
      TabIndex        =   27
      Top             =   855
      Width           =   7125
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
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
         Index           =   12
         Left            =   5175
         MaxLength       =   15
         TabIndex        =   49
         Text            =   "Text1 7"
         Top             =   570
         Width           =   1830
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   9
         Left            =   5250
         MaxLength       =   15
         TabIndex        =   48
         Text            =   "Text1 7"
         Top             =   570
         Width           =   1275
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
         Index           =   31
         Left            =   4395
         MaxLength       =   5
         TabIndex        =   47
         Text            =   "Text1 7"
         Top             =   1605
         Width           =   705
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
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
         Index           =   30
         Left            =   5175
         MaxLength       =   15
         TabIndex        =   46
         Text            =   "Text1 7"
         Top             =   1605
         Width           =   1830
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
         Index           =   29
         Left            =   4395
         MaxLength       =   5
         TabIndex        =   45
         Text            =   "Text1 7"
         Top             =   1965
         Width           =   705
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
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
         Index           =   28
         Left            =   5175
         MaxLength       =   15
         TabIndex        =   44
         Text            =   "Text1 7"
         Top             =   1965
         Width           =   1830
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
         Index           =   27
         Left            =   4410
         MaxLength       =   5
         TabIndex        =   43
         Text            =   "Text1 7"
         Top             =   2340
         Width           =   705
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
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
         Index           =   26
         Left            =   5175
         MaxLength       =   15
         TabIndex        =   42
         Text            =   "Text1 7"
         Top             =   2340
         Width           =   1830
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0FF&
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
         Index           =   25
         Left            =   4410
         MaxLength       =   15
         TabIndex        =   41
         Text            =   "Text1 7"
         Top             =   2790
         Width           =   2580
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
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
         Index           =   24
         Left            =   2940
         MaxLength       =   15
         TabIndex        =   40
         Text            =   "Text1 7"
         Top             =   2340
         Width           =   1395
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
         Index           =   22
         Left            =   2250
         MaxLength       =   5
         TabIndex        =   39
         Text            =   "Text1 7"
         Top             =   2340
         Width           =   660
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
         Index           =   23
         Left            =   675
         MaxLength       =   15
         TabIndex        =   38
         Text            =   "Text1 7"
         Top             =   2340
         Width           =   1530
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
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
         Index           =   20
         Left            =   2940
         MaxLength       =   15
         TabIndex        =   37
         Text            =   "Text1 7"
         Top             =   1965
         Width           =   1395
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
         Index           =   18
         Left            =   2235
         MaxLength       =   5
         TabIndex        =   36
         Text            =   "Text1 7"
         Top             =   1965
         Width           =   660
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
         Index           =   19
         Left            =   675
         MaxLength       =   15
         TabIndex        =   35
         Text            =   "Text1 7"
         Top             =   1980
         Width           =   1530
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
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
         Index           =   16
         Left            =   2940
         MaxLength       =   15
         TabIndex        =   34
         Text            =   "Text1 7"
         Top             =   1605
         Width           =   1395
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
         Index           =   14
         Left            =   2235
         MaxLength       =   5
         TabIndex        =   33
         Text            =   "Text1 7"
         Top             =   1605
         Width           =   660
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
         Index           =   15
         Left            =   675
         MaxLength       =   15
         TabIndex        =   32
         Text            =   "Text1 7"
         Top             =   1605
         Width           =   1530
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
         Index           =   13
         Left            =   90
         MaxLength       =   5
         TabIndex        =   31
         Text            =   "Text1 7"
         Top             =   1605
         Width           =   525
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
         Index           =   17
         Left            =   90
         MaxLength       =   5
         TabIndex        =   30
         Text            =   "Text1 7"
         Top             =   1965
         Width           =   525
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
         Index           =   21
         Left            =   90
         MaxLength       =   5
         TabIndex        =   29
         Text            =   "Text1 7"
         Top             =   2340
         Width           =   525
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   10
         Left            =   5220
         MaxLength       =   15
         TabIndex        =   28
         Text            =   "Text1 7"
         Top             =   570
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Importe Recargo"
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
         Index           =   16
         Left            =   5205
         TabIndex        =   59
         Top             =   1350
         Width           =   1785
      End
      Begin VB.Label Label1 
         Caption         =   "% Rec"
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
         Index           =   15
         Left            =   4395
         TabIndex        =   58
         Top             =   1350
         Width           =   720
      End
      Begin VB.Label Label1 
         Caption         =   "% IVA"
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
         Index           =   41
         Left            =   2235
         TabIndex        =   57
         Top             =   1350
         Width           =   630
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "TOTAL PARTE"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   39
         Left            =   2745
         TabIndex        =   56
         Top             =   2820
         Width           =   1515
      End
      Begin VB.Label Label1 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   36
         Left            =   11880
         TabIndex        =   55
         Top             =   2160
         Width           =   135
      End
      Begin VB.Label Label1 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   37
         Left            =   6075
         TabIndex        =   54
         Top             =   1065
         Width           =   135
      End
      Begin VB.Label Label1 
         Caption         =   "Importe IVA"
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
         Index           =   33
         Left            =   2940
         TabIndex        =   53
         Top             =   1350
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Base Imponible"
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
         Index           =   13
         Left            =   690
         TabIndex        =   52
         Top             =   1320
         Width           =   1530
      End
      Begin VB.Label Label1 
         Caption         =   "Cod."
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
         Index           =   12
         Left            =   120
         TabIndex        =   51
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Importe Bruto"
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
         Index           =   7
         Left            =   5220
         TabIndex        =   50
         Top             =   300
         Width           =   1755
      End
      Begin VB.Line Line1 
         X1              =   5175
         X2              =   6975
         Y1              =   975
         Y2              =   975
      End
   End
   Begin VB.Frame Frame2 
      Height          =   3480
      Left            =   90
      TabIndex        =   14
      Top             =   870
      Width           =   7305
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
         Index           =   6
         Left            =   2865
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   22
         Text            =   "Text2"
         Top             =   1290
         Width           =   4140
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
         Index           =   6
         Left            =   1905
         MaxLength       =   6
         TabIndex        =   4
         Tag             =   "Socio|N|N|0|999999|advpartes|codsocio|000000||"
         Text            =   "Text1"
         Top             =   1290
         Width           =   945
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
         Left            =   1905
         MaxLength       =   10
         TabIndex        =   1
         Tag             =   "Fecha Parte|F|N|||advpartes|fechapar|dd/mm/yyyy|N|"
         Top             =   390
         Width           =   1305
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Facturar"
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
         Index           =   0
         Left            =   5820
         TabIndex        =   2
         Tag             =   "Facturar S/N|N|N|||advpartes|factursn|0||"
         Top             =   390
         Width           =   1365
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
         Index           =   4
         Left            =   1905
         MaxLength       =   8
         TabIndex        =   5
         Tag             =   "Codigo Campo|N|N|0|99999999|advpartes|codcampo|00000000||"
         Text            =   "Text1"
         Top             =   1695
         Width           =   945
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
         Index           =   4
         Left            =   2865
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   20
         Text            =   "Text2"
         Top             =   1695
         Width           =   4140
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
         Left            =   2865
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   18
         Text            =   "Text2"
         Top             =   885
         Width           =   4140
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Index           =   2
         Left            =   210
         MaxLength       =   200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   8
         Tag             =   "Observaciones|T|S|||advpartes|observac|||"
         Top             =   2790
         Width           =   6810
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
         Left            =   1905
         MaxLength       =   4
         TabIndex        =   3
         Tag             =   "Cod.Tratamiento|T|N|||advpartes|codtrata|||"
         Text            =   "Text1"
         Top             =   885
         Width           =   945
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000013&
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
         Left            =   210
         MaxLength       =   7
         TabIndex        =   0
         Tag             =   "NºParte|N|S|||advpartes|numparte|0000000|S|"
         Text            =   "Text1 7"
         Top             =   390
         Width           =   1110
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   5
         Left            =   2970
         MaxLength       =   10
         TabIndex        =   24
         Top             =   2850
         Width           =   1065
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
         Index           =   7
         Left            =   1905
         MaxLength       =   7
         TabIndex        =   6
         Tag             =   "Litros Previstos|N|N|0|999999|advpartes|litrospre|###,##0||"
         Top             =   2100
         Width           =   1245
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
         Index           =   8
         Left            =   5760
         MaxLength       =   7
         TabIndex        =   7
         Tag             =   "Litros Reales|N|N|0|999999|advpartes|litrosrea|###,##0||"
         Top             =   2115
         Width           =   1245
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   360
         Index           =   11
         Left            =   5805
         MaxLength       =   10
         TabIndex        =   60
         Top             =   2115
         Width           =   1065
      End
      Begin VB.Label Label1 
         Caption         =   "Litros Previstos "
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
         Left            =   240
         TabIndex        =   26
         Top             =   2160
         Width           =   1620
      End
      Begin VB.Label Label1 
         Caption         =   "Litros Reales"
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
         Index           =   1
         Left            =   4410
         TabIndex        =   25
         Top             =   2160
         Width           =   1425
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
         Index           =   17
         Left            =   225
         TabIndex        =   23
         Top             =   1335
         Width           =   540
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   2
         Left            =   1605
         ToolTipText     =   "Buscar Socio"
         Top             =   1335
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   1
         Left            =   1605
         ToolTipText     =   "Buscar Campo"
         Top             =   1740
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Campo"
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
         Index           =   8
         Left            =   225
         TabIndex        =   21
         Top             =   1740
         Width           =   900
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha "
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
         Left            =   1905
         TabIndex        =   19
         Top             =   135
         Width           =   585
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   0
         Left            =   2940
         Picture         =   "frmADVPartes.frx":006E
         ToolTipText     =   "Buscar fecha"
         Top             =   150
         Width           =   240
      End
      Begin VB.Image imgZoom 
         Height          =   240
         Index           =   0
         Left            =   1890
         ToolTipText     =   "Zoom descripción"
         Top             =   2520
         Width           =   240
      End
      Begin VB.Label Label29 
         Caption         =   "Observaciones"
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
         Left            =   225
         TabIndex        =   17
         Top             =   2550
         Width           =   1440
      End
      Begin VB.Label Label1 
         Caption         =   "Tratamiento"
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
         Left            =   225
         TabIndex        =   16
         Top             =   930
         Width           =   1245
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   1605
         ToolTipText     =   "Buscar Tratamiento"
         Top             =   915
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "NºParte"
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
         Left            =   225
         TabIndex        =   15
         Top             =   135
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Height          =   475
      Index           =   0
      Left            =   90
      TabIndex        =   12
      Top             =   9225
      Width           =   2175
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
         TabIndex        =   13
         Top             =   135
         Width           =   1785
      End
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
      Left            =   13500
      TabIndex        =   10
      Top             =   9315
      Width           =   1065
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
      Left            =   12330
      TabIndex        =   9
      Top             =   9315
      Width           =   1065
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
      Left            =   13500
      TabIndex        =   11
      Top             =   9315
      Visible         =   0   'False
      Width           =   1065
   End
   Begin MSAdodcLib.Adodc Data3 
      Height          =   330
      Left            =   3000
      Top             =   1080
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
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   240
      Top             =   7890
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
   Begin MSAdodcLib.Adodc Data2 
      Height          =   330
      Left            =   180
      Top             =   7950
      Visible         =   0   'False
      Width           =   1425
      _ExtentX        =   2514
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
      Left            =   14145
      TabIndex        =   107
      Top             =   210
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
         Caption         =   "&Ver Todos"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnBarra1 
         Caption         =   "-"
      End
      Begin VB.Menu mnNuevo 
         Caption         =   "&Nuevo"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnModificar 
         Caption         =   "&Modificar"
         Shortcut        =   ^M
      End
      Begin VB.Menu mnEliminar 
         Caption         =   "&Eliminar"
         Shortcut        =   ^E
      End
      Begin VB.Menu mnBarra2 
         Caption         =   "-"
      End
      Begin VB.Menu mnConfirmacion 
         Caption         =   "&Confirmación"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnCuadrilla 
         Caption         =   "Cuadrilla"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnPrecios 
         Caption         =   "Asignacion Precios"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnInsercionGastos 
         Caption         =   "Inserción Gastos"
         Shortcut        =   ^G
      End
      Begin VB.Menu mnImprimir 
         Caption         =   "&Imprimir"
         HelpContextID   =   2
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
Attribute VB_Name = "frmADVPartes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'========== VBLES PUBLICAS ====================
Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)

Public Albaran As String  ' venimos de albaranes para ver las facturas donde aparece el albaran

'========== VBLES PRIVADAS ====================
Private WithEvents frmB As frmBuscaGrid 'Form para busquedas
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmZ As frmZoom  'Zoom para campos Text
Attribute frmZ.VB_VarHelpID = -1
Private WithEvents frmArt As frmADVArticulos 'Form Mto de Articulos de adv
Attribute frmArt.VB_VarHelpID = -1
Private WithEvents frmTra As frmADVTratamientos 'Form Mto de Tratamientos de adv
Attribute frmTra.VB_VarHelpID = -1
Private WithEvents frmTra2 As frmADVTrataMoi  'Form Mto de Tipos de venta
Attribute frmTra2.VB_VarHelpID = -1
Private WithEvents frmMens As frmMensajes  ' form de mensajes
Attribute frmMens.VB_VarHelpID = -1
Private WithEvents frmSoc As frmManSocios 'Form Mto de Socios
Attribute frmSoc.VB_VarHelpID = -1
Private WithEvents frmCua As frmManCuadrillas ' form de cuadrillas
Attribute frmCua.VB_VarHelpID = -1
Private WithEvents frmTra1 As frmManTraba 'Form Mto de Trabajadores
Attribute frmTra1.VB_VarHelpID = -1

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

Dim ModificaLineas As Byte
'1.- Añadir,  2.- Modificar,  3.- Borrar,  0.-Pasar control a Lineas

Dim CodTipoMov As String
'Codigo tipo de movimiento en función del valor en la tabla de parámetros: stipom

Dim TituloLinea As String 'Descripcion de la linea que estamos en Mantenimiento

Dim PrimeraVez As Boolean

Dim EsCabecera As Boolean


Dim EsDeVarios As Boolean
'Si el cliente mostrado es de Varios o No

'SQL de la tabla principal del formulario
Private CadenaConsulta As String

Private Ordenacion As String
Private NombreTabla As String  'Nombre de la tabla o de la
Private NomTablaLineas As String 'Nombre de la Tabla de lineas
Private kCampo As Integer
'-------------------------------------------------------------------------
Private HaDevueltoDatos As Boolean

Dim btnPrimero As Byte
'Variable que indica el número del Boton  PrimerRegistro en la Toolbar1


Private VieneDeBuscar As Boolean
'Para cuando devuelve dos poblaciones con el mismo codigo Postal. Si viene de pulsar prismatico
'de busqueda poner el valor de poblacion seleccionado y no volver a recuperar de la Base de Datos
Private HaCambiadoCP As Boolean
'Para saber si tras haber vuelto de prismaticos ha cambiado el valor del CPostal
Dim Indice As Byte

Dim TipoFactura As Byte

Dim Confirmacion As Boolean
Dim CampoAnt As Long
Dim LitrosAnt As Long
Dim CuadrillaAnt As Long

Dim ModoCuadrilla As Boolean
Dim numTab As Byte

Dim UniCajas As Long

Private BuscaChekc As String

Private Sub btnBuscar_Click(Index As Integer)
    TerminaBloquear
    Select Case Index
        Case 0 'Articulos
            Set frmArt = New frmADVArticulos
            frmArt.DatosADevolverBusqueda = "0|1|"
            frmArt.CodigoActual = txtAux(5).Text
            frmArt.Show vbModal
            Set frmArt = Nothing
            PonerFoco txtAux(5)
    
        Case 1 ' trabajadores
            Set frmTra1 = New frmManTraba
            frmTra1.DatosADevolverBusqueda = "0|2|"
            frmTra1.Show vbModal
            Set frmTra1 = Nothing
            PonerFoco txtAux1(2)
         

    End Select
    If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1

End Sub

Private Sub Check1_Click(Index As Integer)
    If Modo = 1 Then
        'Buscqueda
        If InStr(1, BuscaChekc, "check1(" & Index & ")") = 0 Then BuscaChekc = BuscaChekc & "check1(" & Index & ")|"
    End If
End Sub


Private Sub Check1_GotFocus(Index As Integer)
    PonerFocoChk Me.Check1(Index)
End Sub

Private Sub Check1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub cmdAceptar_Click()
Dim I As Integer

    Screen.MousePointer = vbHourglass
    On Error GoTo Error1

    Select Case Modo
        Case 1  'BUSQUEDA
            HacerBusqueda
            ModoCuadrilla = False
        Case 3  'AÑADIR
            If DatosOK Then InsertarCabecera

        Case 4  'MODIFICAR
            If DatosOK Then
                If ModoCuadrilla Then
                    If ModificaCabeceraCuadrilla Then
                        espera 0.2
                        TerminaBloquear
                        PosicionarData
                        PonerCampos
                        PonerCamposLineas
                    End If
                Else
                    If ModificaCabecera Then
                        espera 0.2
                        TerminaBloquear
                        PosicionarData
                        PonerCampos
                        PonerCamposLineas
                        
                        CalcularDatosAlbaran
                    End If
                End If
            End If
            
         Case 5 'InsertarModificar LINEAS
            Select Case ModificaLineas
                Case 1 'afegir llínia
                    InsertarLinea
                Case 2 'modificar llínies
                    Select Case numTab
                        Case 1
                            If ModificarLinea Then PosicionarData
                        Case 0
                            If ModificarLineaCuadrilla Then PosicionarData
                    End Select
            End Select
    End Select
    Screen.MousePointer = vbDefault
    
    Confirmacion = False

    
Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub cmdCancelar_Click()
    Select Case Modo
        Case 0, 1, 3 'Busqueda, Insertar
            LimpiarCampos
            'Poner los grid sin apuntar a nada
            LimpiarDataGrids
            PonerModo 0
            LLamaLineas Modo, 0, "DataGrid2"
            PonerFoco Text1(3)
            
        Case 4  'Modificar
            TerminaBloquear
            PonerModo 2
            PonerCampos
            PonerFoco Text1(3)
            
        Case 5 'Lineas Detalle
            TerminaBloquear
            
            Select Case numTab
                Case 1
                    If ModificaLineas = 1 Then 'INSERTAR
                        ModificaLineas = 0
                        DataGrid3.AllowAddNew = False
                        If Not Adoaux(1).Recordset.EOF Then Adoaux(1).Recordset.MoveFirst
                    End If
                    ModificaLineas = 0
                    LLamaLineas Modo, 0, "DataGrid3"
                    PonerModo 2
                    DataGrid3.Enabled = True
                    If Not Data1.Recordset.EOF Then _
                        Me.lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
            
                    'Habilitar las opciones correctas del menu segun Modo
                    PonerModoOpcionesMenu (Modo)
                    PonerOpcionesMenu 'Habilitar las opciones correctas del menu segun Nivel de Acceso
                    DataGrid3.Enabled = True
                    PonerFocoGrid DataGrid3
                Case 0
                    If ModificaLineas = 1 Then 'INSERTAR
                        ModificaLineas = 0
                        DataGrid1.AllowAddNew = False
                        If Not Adoaux(0).Recordset.EOF Then Adoaux(0).Recordset.MoveFirst
                    End If
                    ModificaLineas = 0
                    LLamaLineas Modo, 0, "DataGrid1"
                    PonerModo 2
                    DataGrid1.Enabled = True
                    If Not Data1.Recordset.EOF Then _
                        Me.lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
            
                    'Habilitar las opciones correctas del menu segun Modo
                    PonerModoOpcionesMenu (Modo)
                    PonerOpcionesMenu 'Habilitar las opciones correctas del menu segun Nivel de Acceso
                    DataGrid1.Enabled = True
                    PonerFocoGrid DataGrid1
                
           End Select
    End Select
    Confirmacion = False
    ModoCuadrilla = False
    
End Sub

Private Sub BotonAnyadir()
    LimpiarCampos 'Huida els TextBox
    
    
    PonerModo 3
    ' ****** Valors per defecte a l'afegir, repasar si n'hi ha
    ' codEmpre i quins camps tenen la PK de la capçalera *******
'    Text1(0).Text = SugerirCodigoSiguienteStr("palets", "numpalet")
'    FormateaCampo Text1(0)
'    cmbAux(0).ListIndex = -1
    
    Text1(1).Text = Format(Now, "dd/mm/yyyy")
    Check1(0).Value = 1
    
    'los litros reales pasan a ser 0
    Text1(8).Text = 0
    Text1(35).Tag = ""
    
    '[Monica]18/05/2012
    If vParamAplic.Cooperativa = 3 Then
        ' valores por defecto
        Text1(4).Text = 0 ' codigo de campo no puede ser nulo
        Text1(7).Text = 0 ' litros previstos
        Text1(8).Text = 0 ' litros reales
    End If
    
    LimpiarDataGrids
    
    PonerFoco Text1(1) '*** 1r camp visible que siga PK ***
    
    ' *** si n'hi han camps de descripció a la capçalera ***
    'PosarDescripcions
End Sub


Private Sub BotonBuscar()
Dim anc As Single

    'Buscar
    If Modo <> 1 Then
        LimpiarCampos
        'Poner los grid sin apuntar a nada
        LimpiarDataGrids
        PonerModo 1
        
'        'poner los txtaux para buscar por lineas de albaran
'        anc = DataGrid2.Top
'        If DataGrid2.Row < 0 Then
'            anc = anc + 440
'        Else
'            anc = anc + DataGrid2.RowTop(DataGrid2.Row) + 20
'        End If
'        LLamaLineas Modo, anc, "DataGrid2"
        
        
        'Si pasamos el control aqui lo ponemos en amarillo
        PonerFoco Text1(0)
        Text1(0).BackColor = vbLightBlue
    Else
        HacerBusqueda
        If Data1.Recordset.EOF Then
            Text1(kCampo).Text = ""
            Text1(kCampo).BackColor = vbLightBlue
            PonerFoco Text1(kCampo)
        End If
    End If
End Sub


Private Sub BotonVerTodos()
    If chkVistaPrevia.Value = 1 Then
        EsCabecera = True
        MandaBusquedaPrevia ""
    Else
        LimpiarCampos
        LimpiarDataGrids
        CadenaConsulta = "Select advpartes.* "
        CadenaConsulta = CadenaConsulta & "from " & NombreTabla
'        CadenaConsulta = CadenaConsulta & " WHERE scafac.codtipom='" & CodTipoMov & "'"
        PonerCadenaBusqueda
    End If
End Sub


Private Sub Desplazamiento(Index As Integer)
'Botones de Desplazamiento de la Toolbar
'Para desplazarse por los registros de control Data
    If Data1.Recordset.EOF Then Exit Sub
    DesplazamientoData Data1, Index, True
    PonerCampos
End Sub


Private Sub BotonModificar()
Dim DeVarios As Boolean

    'Añadiremos el boton de aceptar y demas objetos para insertar
    PonerModo 4
    
    CampoAnt = Text1(4).Text
    LitrosAnt = Text1(7).Text
    
    
    Text1(35).Tag = "NºParte|N|S|||advpartes|numparte|0000000|S|"
    
    PonerFoco Text1(4) '*** 1r camp visible que siga PK ***
    
    If Confirmacion Then
        PonerFoco Text1(8)
'    Else
''       [Monica] 30/09/2009: solo dejo modificar si no esta confirmado i.e. sin litros reales
'        If DBLet(Data1.Recordset!LitrosRea, "N") <> 0 Then
'            PonerFoco Text1(2)
'        End If
    End If
    
End Sub


Private Sub BotonModificarLinea(Index As Integer)
'Modificar una linea
Dim vWhere As String
Dim anc As Single
Dim J As Byte

    On Error GoTo eModificarLinea


    ModificaLineas = 2 'Modificar

    'Si no estaba modificando lineas salimos
    'Es decir, si estaba insertando linea no podemos hacer otra cosa
    If ModificaLineas = 1 Then  '1= Insertar
        TerminaBloquear
        Exit Sub
    End If
    
    numTab = Index
    
'--monica
'    If Data2.Recordset.EOF Then
'        TerminaBloquear
'        Exit Sub
'    End If
    
    PonerModo 5, Index
    
    Select Case Index
        Case 1
    
            vWhere = ObtenerWhereCP(False)
            vWhere = vWhere & " and numlinea=" & Adoaux(1).Recordset!numlinea
            If Not BloqueaRegistro("advpartes_lineas", vWhere) Then
                TerminaBloquear
                Exit Sub
            End If
            If DataGrid3.Bookmark < DataGrid3.FirstRow Or DataGrid3.Bookmark > (DataGrid3.FirstRow + DataGrid3.VisibleRows - 1) Then
                J = DataGrid3.Bookmark - DataGrid3.FirstRow
                DataGrid3.Scroll 0, J
                DataGrid3.Refresh
            End If
            
        '    anc = ObtenerAlto(Me.DataGrid1)
            anc = DataGrid3.Top
            If DataGrid3.Row < 0 Then
                anc = anc + 210
            Else
                anc = anc + DataGrid3.RowTop(DataGrid3.Row) + 10
            End If
        
            txtAux(1).Text = DataGrid3.Columns(0).Text ' parte
            txtAux(3).Text = DataGrid3.Columns(1).Text ' linea
            txtAux(4).Text = DataGrid3.Columns(2).Text ' almacen
            txtAux(5).Text = DataGrid3.Columns(3).Text ' articulo
            Text2(0).Text = DataGrid3.Columns(4).Text ' nombre del articulo
            txtAux(8).Text = DataGrid3.Columns(5).Text ' dtolinea
            txtAux(6).Text = DataGrid3.Columns(6).Text ' cantidad
            txtAux(7).Text = DataGrid3.Columns(7).Text ' precio
            txtAux(9).Text = DataGrid3.Columns(8).Text ' importe
            Text2(16).Text = DataGrid3.Columns(9).Text ' ampliacion
            txtAux(10).Text = DataGrid3.Columns(10).Text ' codigo de iva
            
        
            BloquearTxt txtAux(4), True
            BloquearTxt txtAux(5), True
        '    BloquearTxt txtAux(7), True
            BloquearTxt txtAux(9), True
            txtAux(4).Enabled = False
            txtAux(5).Enabled = False
        '    txtAux(7).Enabled = False
            txtAux(9).Enabled = False
            
            BloquearTxt txtAux(6), False
            BloquearTxt txtAux(7), False
            BloquearTxt txtAux(8), False
            
            BloquearBtn Me.btnBuscar(0), True
            
            LLamaLineas ModificaLineas, anc, "DataGrid3"
            
            'Añadiremos el boton de aceptar y demas objetos para insertar
            Me.lblIndicador.Caption = "MODIFICAR"
            PonerModoOpcionesMenu (Modo)
            PonerOpcionesMenu 'Habilitar las opciones correctas del menu segun Nivel de Acceso
            DataGrid3.Enabled = True
            
            PonerFoco txtAux(8)
            Me.DataGrid3.Enabled = False

        Case 0 ' advpartes_trabajadores
        
            vWhere = ObtenerWhereCP(False)
            vWhere = vWhere & " and numlinea=" & Adoaux(0).Recordset!numlinea
            If Not BloqueaRegistro("advpartes_lineas", vWhere) Then
                TerminaBloquear
                Exit Sub
            End If
            If DataGrid1.Bookmark < DataGrid1.FirstRow Or DataGrid1.Bookmark > (DataGrid1.FirstRow + DataGrid1.VisibleRows - 1) Then
                J = DataGrid1.Bookmark - DataGrid1.FirstRow
                DataGrid1.Scroll 0, J
                DataGrid1.Refresh
            End If
            
        '    anc = ObtenerAlto(Me.DataGrid1)
            anc = DataGrid1.Top
            If DataGrid1.Row < 0 Then
                anc = anc + 210
            Else
                anc = anc + DataGrid1.RowTop(DataGrid1.Row) + 10
            End If
        
            txtAux1(0).Text = DataGrid1.Columns(0).Text ' parte
            txtAux1(1).Text = DataGrid1.Columns(1).Text ' linea
            txtAux1(2).Text = DataGrid1.Columns(2).Text ' trabajador
            txtAux1(3).Text = DataGrid1.Columns(3).Text ' nombre del trabajador
            txtAux1(4).Text = DataGrid1.Columns(4).Text ' horas
            txtAux1(5).Text = DataGrid1.Columns(5).Text ' precio
            txtAux1(6).Text = DataGrid1.Columns(6).Text ' importe
            
        
            BloquearTxt txtAux1(4), False
            BloquearTxt txtAux1(5), False
            BloquearTxt txtAux1(6), False
            
            BloquearBtn Me.btnBuscar(1), True
            
            LLamaLineas ModificaLineas, anc, "DataGrid1"
            
            'Añadiremos el boton de aceptar y demas objetos para insertar
            Me.lblIndicador.Caption = "MODIFICAR"
            PonerModoOpcionesMenu (Modo)
            PonerOpcionesMenu 'Habilitar las opciones correctas del menu segun Nivel de Acceso
            DataGrid1.Enabled = True
            
            PonerFoco txtAux1(4)
            Me.DataGrid1.Enabled = False
            
    End Select
    
eModificarLinea:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub LLamaLineas(xModo As Byte, Optional alto As Single, Optional grid As String)
Dim jj As Integer
Dim B As Boolean
    
    Select Case grid
        Case "DataGrid3"
            DeseleccionaGrid Me.DataGrid3
            B = (xModo = 1 Or xModo = 2)
            For jj = 4 To 9
                txtAux(jj).Height = DataGrid3.RowHeight - 10
                txtAux(jj).Top = alto + 5
                txtAux(jj).visible = B
                txtAux(jj).Enabled = B
            Next jj
            
            '[Monica]18/05/2012: la dosis habitual no la sacamos en lso albaranes
            If vParamAplic.Cooperativa = 3 Then
                ' la sacamos pq son los bultos
                txtAux(8).visible = B
                txtAux(8).Enabled = B
            End If
            
            txtAux(9).Enabled = False
            
            Text2(0).Height = DataGrid3.RowHeight - 10
            Text2(0).Top = alto + 5
            Text2(0).visible = B
           
            btnBuscar(0).Height = DataGrid3.RowHeight - 10
            btnBuscar(0).Top = alto + 5
            btnBuscar(0).visible = B
        
        Case "DataGrid1"
            DeseleccionaGrid Me.DataGrid1
            B = (xModo = 1 Or xModo = 2)
            For jj = 2 To 2
                txtAux1(jj).Height = DataGrid1.RowHeight - 10
                txtAux1(jj).Top = alto + 5
                txtAux1(jj).visible = B
                txtAux1(jj).Enabled = B
            Next jj
            For jj = 3 To 3
                txtAux1(jj).Height = DataGrid1.RowHeight - 10
                txtAux1(jj).Top = alto + 5
                txtAux1(jj).visible = B
            Next jj
            For jj = 4 To 6
                BloquearTxt txtAux1(jj), False
                txtAux1(jj).Height = DataGrid1.RowHeight - 10
                txtAux1(jj).Top = alto + 5
                txtAux1(jj).visible = B
                txtAux1(jj).Enabled = B
            Next jj
            BloquearTxt txtAux1(3), True
            
            btnBuscar(1).Height = DataGrid1.RowHeight - 10
            btnBuscar(1).Top = alto + 5
            btnBuscar(1).visible = B

    End Select
End Sub


Private Sub BotonEliminar()
'Eliminar Registro de la Cabecera: Tabla de Facturas (scafac)
' y los registros correspondientes de las tablas cab. albaranes (scafac1)
' y las lineas de la factura (slifac)
Dim cad As String
'Dim vTipoMov As CTiposMov

    On Error GoTo EEliminar

    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub
    
'    'solo se puede modificar la factura si no esta contabilizada
'    If FactContabilizada Then Exit Sub
    
    cad = "Cabecera de Albaranes." & vbCrLf
    cad = cad & "-------------------------------------      " & vbCrLf & vbCrLf
    cad = cad & "Va a eliminar el Parte:            "
    cad = cad & vbCrLf & "Nº Parte:  " & Format(Text1(0).Text, "0000000")
    cad = cad & vbCrLf & "Fecha:  " & Format(Text1(1).Text, "dd/mm/yyyy")

    cad = cad & vbCrLf & vbCrLf & " ¿Desea Eliminarlo? "

    'Borramos
    If MsgBox(cad, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
'        On Error GoTo EEliminar
        Screen.MousePointer = vbHourglass
        NumRegElim = Data1.Recordset.AbsolutePosition
'        NumPedElim = Data1.Recordset.Fields(1).Value
        If Not Eliminar Then
            Screen.MousePointer = vbDefault
            Exit Sub
        ElseIf SituarDataTrasEliminar(Data1, NumRegElim) Then
            PonerCampos
        Else
            LimpiarCampos
            'Poner los grid sin apuntar a nada
            LimpiarDataGrids
            PonerModo 0
        End If
        
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
    
EEliminar:
    Screen.MousePointer = vbDefault
    MuestraError Err.Number, "Eliminar Parte", Err.Description
End Sub

Private Sub cmdRegresar_Click()
'Este es el boton Cabecera
Dim cad As String

    'Quitar lineas y volver a la cabecera
    If Modo = 5 Then  'modo 5: Mantenimientos Lineas
        PonerModo 2
        DataGrid3.Enabled = True
        If Not Data1.Recordset.EOF Then _
            Me.lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount

    Else 'Se llama desde algún Prismatico de otro Form al Mantenimiento de Trabajadores
        If Data1.Recordset.EOF Then
            MsgBox "Ningún registro devuelto.", vbExclamation
            Exit Sub
        End If
        cad = Data1.Recordset.Fields(0) & "|"
        cad = cad & Data1.Recordset.Fields(1) & "|"
        RaiseEvent DatoSeleccionado(cad)
        Unload Me
    End If
End Sub

Private Sub DataGrid3_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If Not Adoaux(1).Recordset.EOF And ModificaLineas = 2 Then
        Text2(16).Text = DBLet(Adoaux(1).Recordset!ampliaci, "T")
    End If
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
     
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Form_Load()
Dim I As Integer

    'Icono del formulario
    Me.Icon = frmPpal.Icon

     'Icono de busqueda
    For kCampo = 0 To Me.imgBuscar.Count - 1
        Me.imgBuscar(kCampo).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next kCampo

    ' ICONITOS DE LA BARRA
    btnPrimero = 16
'    With Me.Toolbar1
'        .HotImageList = frmPpal.imgListComun_OM
'        .DisabledImageList = frmPpal.imgListComun_BN
'        .ImageList = frmPpal.imgListComun
'        .Buttons(1).Image = 1   'Botón Buscar
'        .Buttons(2).Image = 2   'Botón Todos
'        .Buttons(4).Image = 3   'Insertar
'        .Buttons(5).Image = 4   'Modificar
'        .Buttons(6).Image = 5   'Borrar
'        .Buttons(8).Image = 26  'Confirmacion
'        .Buttons(9).Image = 19  'Cuadrilla
'        .Buttons(10).Image = 19  'Asignacion de precios
'        .Buttons(11).Image = 33  'Insercion de gastos
'        .Buttons(12).Image = 10  'Impresión el parte
'        .Buttons(14).Image = 11  'Salir
'        .Buttons(btnPrimero).Image = 6  'Primero
'        .Buttons(btnPrimero + 1).Image = 7 'Anterior
'        .Buttons(btnPrimero + 2).Image = 8 'Siguiente
'        .Buttons(btnPrimero + 3).Image = 9 'Último
'    End With
    
    With Me.Toolbar1
        .HotImageList = frmPpal.imgListComun_OM
        .DisabledImageList = frmPpal.imgListComun_BN
        .ImageList = frmPpal.imgListComun
        'l'1 i el 2 son separadors
        .Buttons(5).Image = 1   'Buscar
        .Buttons(6).Image = 2   'Totss
        'el 5 i el 6 son separadors
        .Buttons(1).Image = 3   'Insertar
        .Buttons(2).Image = 4   'Modificar
        .Buttons(3).Image = 5   'Borrar
        'el 10  son separadors
        .Buttons(8).Image = 10  'Imprimir
    End With
    
    With Me.Toolbar5
        .HotImageList = frmPpal.imgListComun_OM
        .DisabledImageList = frmPpal.imgListComun_BN
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 26  'Confirmacion
        .Buttons(2).Image = 19  'Cuadrilla
        .Buttons(3).Image = 19  'Asignacion de precios
        .Buttons(4).Image = 33  'Insercion de gastos
    End With
    
    ' desplazamiento
    With Me.ToolbarDes
        .HotImageList = frmPpal.imgListComun_OM
        .DisabledImageList = frmPpal.imgListComun_BN
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 6
        .Buttons(2).Image = 7
        .Buttons(3).Image = 8
        .Buttons(4).Image = 9
    End With
    
    ' La Ayuda
    With Me.ToolbarAyuda
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 12
    End With
    
    
    ' ******* si n'hi han llínies *******
    'ICONETS DE LES BARRES ALS TABS DE LLÍNIA
    For kCampo = 0 To 1
        With Me.ToolAux(kCampo)
            .HotImageList = frmPpal.imgListComun_OM16
            .DisabledImageList = frmPpal.imgListComun_BN16
            .ImageList = frmPpal.imgListComun16
            .Buttons(1).Image = 3   'Insertar
            .Buttons(2).Image = 4   'Modificar
            .Buttons(3).Image = 5   'Borrar
        End With
    Next kCampo
   ' ***********************************
   'IMAGES para zoom
    For I = 0 To Me.imgZoom.Count - 1
        Me.imgZoom(I).Picture = frmPpal.imgListImages16.ListImages(3).Picture
    Next I
    
    LimpiarCampos   'Limpia los campos TextBox

    CodTipoMov = "PAR"
    VieneDeBuscar = False
    
        
    '## A mano
    NombreTabla = "advpartes"
    NomTablaLineas = "advpartes_lineas" 'Tabla lineas de tratamiento del parte
    Ordenacion = " ORDER BY advpartes.numparte"
    
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    
    CadenaConsulta = "select * from advpartes "
    If Albaran <> "" Then
        CadenaConsulta = CadenaConsulta & " where numparte = " & Albaran
    Else
        CadenaConsulta = CadenaConsulta & " where numparte = -1"
    End If
    
    
    Data1.ConnectionString = conn
    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
        
    '[Monica]18/05/2012: cambiamos la apariencia del formulario para Moixent
    If vParamAplic.Cooperativa = 3 Then
        Label1(28).Caption = "Albarán"
        Label1(0).Caption = "Tipo Vta"
        Label1(8).visible = False
        imgBuscar(1).visible = False
        imgBuscar(1).Enabled = False
        Text1(4).visible = False
        Text1(4).Enabled = False
        Text2(4).visible = False
        Text2(4).Enabled = False
        Label29.Top = 1740 '1515
        imgZoom(0).Top = 1740 '1515
        Text1(2).Top = 2060 '1830
        Text1(2).Height = 1115 '1200
        imgBuscar(0).ToolTipText = "Buscar Tipo de Venta"
        
        frmADVPartes.Caption = "Mantenimiento de Albaranes de Venta"
        SSTab1.TabEnabled(1) = False
        SSTab1.TabVisible(1) = False
        mnConfirmacion.Enabled = False
        mnConfirmacion.visible = False
        mnCuadrilla.Enabled = False
        mnCuadrilla.visible = False
        Me.Toolbar5.Buttons(1).Enabled = False
        Me.Toolbar5.Buttons(1).visible = False
        Me.Toolbar5.Buttons(2).Enabled = False
        Me.Toolbar5.Buttons(2).visible = False
        
        Text1(4).TabIndex = 100
'        Text1(7).TabIndex = 101
'        Text1(8).TabIndex = 102
        
        txtAux(8).Tag = "Bultos|N|S|||advpartes_lineas|dosishab|##,##0||"
        txtAux(8).TabIndex = 69
        txtAux(6).TabIndex = 70
        
    Else '[Monica]24/07/2012: la asignacion de precios es unicamente para moixent
        mnPrecios.Enabled = False
        mnPrecios.visible = False
        Me.Toolbar5.Buttons(3).Enabled = False
        Me.Toolbar5.Buttons(3).visible = False
    End If
        
        
        
        
    'Cargar el DataGrid de lineas de Revisiones inicialmente a nada DATA2
    PrimeraVez = True
    
'        CargaGrid DataGrid1, Data2, False
    'Poner los grid sin apuntar a nada
    LimpiarDataGrids
    
    SSTab1.Tab = 0
    
    If DatosADevolverBusqueda = "" Then
        If Albaran = "" Then
            PonerModo 0
        Else
            HacerBusqueda
'            SSTab1.Tab = 0
        End If
    Else
        BotonBuscar
    End If
    
    PrimeraVez = False
End Sub


Private Sub LimpiarCampos()
On Error Resume Next
    limpiar Me   'Metodo general: Limpia los controles TextBox
    lblIndicador.Caption = ""
    'Aqui va el especifico de cada form es
    '### a mano
    Me.Check1(0).Value = 0
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub Form_Unload(Cancel As Integer)
    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
    If Modo = 4 Then TerminaBloquear
End Sub


Private Sub frmArt_DatoSeleccionado(CadenaSeleccion As String)
'Articulos
    txtAux(5).Text = RecuperaValor(CadenaSeleccion, 1) 'codartic
    Text2(0).Text = RecuperaValor(CadenaSeleccion, 2) 'descripcion
    If txtAux(5) <> "" Then
        txtAux(7) = DevuelveDesdeBDNew(cAgro, "advartic", "preciove", "codartic", txtAux(5), "T")
        ' nos guardamos el codigo de iva del articulo
        txtAux(10) = DevuelveDesdeBDNew(cAgro, "advartic", "codigiva", "codartic", txtAux(5), "T")
    End If
End Sub

Private Sub frmB_Selecionado(CadenaDevuelta As String)
Dim CadB As String
Dim Aux As String
      
    If CadenaDevuelta <> "" Then
        HaDevueltoDatos = True
        Screen.MousePointer = vbHourglass
        CadB = ""
        Aux = ValorDevueltoFormGrid(Text1(0), CadenaDevuelta, 1)
        CadB = Aux
'        Aux = ValorDevueltoFormGrid(Text1(0), CadenaDevuelta, 2)
'        CadB = CadB & " and  " & Aux
'        Aux = ValorDevueltoFormGrid(Text1(1), CadenaDevuelta, 3)
'        CadB = CadB & " and " & Aux
        
        
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " " & Ordenacion
        PonerCadenaBusqueda
        Screen.MousePointer = vbDefault
    End If
    Screen.MousePointer = vbDefault
End Sub



Private Sub frmC_Selec(vFecha As Date)
    ' *** repasar si el camp es txtAux o Text1 ***
    If imgFec(0).Tag < 2 Then
        Text1(CByte(imgFec(0).Tag) + 1).Text = Format(vFecha, "dd/mm/yyyy") '<===
    Else
        Text1(CByte(imgFec(0).Tag) + 8).Text = Format(vFecha, "dd/mm/yyyy") '<===
    End If
    ' ********************************************
End Sub

Private Sub frmCua_DatoSeleccionado(CadenaSeleccion As String)
'cuadrilla
    Text1(32).Text = RecuperaValor(CadenaSeleccion, 1) 'codcuadrilla
    Text2(32).Text = RecuperaValor(CadenaSeleccion, 2) 'nombre
End Sub

Private Sub frmMens_DatoSeleccionado(CadenaSeleccion As String)
Dim cad As String

    If CadenaSeleccion = "" Then
        Text1(4).Text = 0
        Exit Sub
    End If

    Text1(4).Text = RecuperaValor(CadenaSeleccion, 1)
    PonerDatosCampo Text1(4)
    
'   [Monica]:20/02/2011 Si hubiera más de un campo seleccionado lo metemos en las observaciones
    If RecuperaValor(CadenaSeleccion, 2) <> "" Then
        If Text1(2).Text = "" Then
            cad = Mid(CadenaSeleccion, InStr(1, CadenaSeleccion, RecuperaValor(CadenaSeleccion, 2)), Len(CadenaSeleccion) - Len(RecuperaValor(CadenaSeleccion, 1)) - 2)
            Text1(2).Text = Replace(cad, "|", ", ")
        Else
            Text1(2).Text = Text1(2).Text & vbCrLf & Replace(Mid(CadenaSeleccion, InStr(1, CadenaSeleccion, RecuperaValor(CadenaSeleccion, 2)), -Len(RecuperaValor(CadenaSeleccion, 1)) - 2), "|", ", ")
        End If
    End If
End Sub

Private Sub frmSoc_DatoSeleccionado(CadenaSeleccion As String)
    Text1(Indice).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000") 'Cod socio
    Text2(Indice).Text = RecuperaValor(CadenaSeleccion, 2) 'Nombre del socio
    PonerFoco Text1(Indice)
End Sub

Private Sub frmTra_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de tratamientos
    Text1(Indice).Text = RecuperaValor(CadenaSeleccion, 1) 'Codigo
    Text2(Indice).Text = RecuperaValor(CadenaSeleccion, 2) 'Nombre
End Sub

Private Sub frmTra1_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de trabajadores
    txtAux1(2).Text = RecuperaValor(CadenaSeleccion, 1) 'Codigo
    txtAux1(3).Text = RecuperaValor(CadenaSeleccion, 2) 'Nombre
    
    If txtAux1(2).Text <> "" Then PonerPrecioHoraTrabajador txtAux1(2).Text
    
'    PonerDatosTrabajador txtAux1(2).Text
End Sub

Private Sub frmTra2_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de tipos de venta para Moixent
    Text1(Indice).Text = RecuperaValor(CadenaSeleccion, 1) 'Codigo
    Text2(Indice).Text = RecuperaValor(CadenaSeleccion, 2) 'Nombre
End Sub

Private Sub frmZ_Actualizar(vCampo As String)
     Text1(Indice).Text = vCampo
End Sub

Private Sub imgBuscar_Click(Index As Integer)

    If Modo = 2 Or Modo = 0 Then Exit Sub
    Screen.MousePointer = vbHourglass

    Select Case Index
        Case 2 'Cod. de socio
            Indice = 6
            PonerFoco Text1(Indice)
            Set frmSoc = New frmManSocios
            frmSoc.DatosADevolverBusqueda = "0|1|"
            frmSoc.Show vbModal
            Set frmSoc = Nothing
            PonerFoco Text1(Indice)
        
        Case 0 'Tratamiento
            Indice = 3
            PonerFoco Text1(Indice)
            
            '[Monica]18/05/2012:
            If vParamAplic.Cooperativa = 3 Then
                Set frmTra2 = New frmADVTrataMoi
                frmTra2.DatosADevolverBusqueda = "0|1|"
                frmTra2.Show vbModal
                Set frmTra2 = Nothing
            Else
                Set frmTra = New frmADVTratamientos
                frmTra.DatosADevolverBusqueda = "0|1|"
                frmTra.Show vbModal
                Set frmTra = Nothing
            End If
            PonerFoco Text1(Indice)
            
       Case 1 'Campo
            PonerCamposSocio
            PonerFoco Text1(4)
           
       Case 3 ' Cuadrilla
            Indice = 32
            PonerFoco Text1(Indice)
            Set frmCua = New frmManCuadrillas
            frmCua.DatosADevolverBusqueda = "0|1|"
            frmCua.vCondicion = "tipocuadrilla = 1"
            frmCua.Show vbModal
            Set frmCua = Nothing
            PonerFoco Text1(Indice)
            
           
    End Select
    
    Screen.MousePointer = vbDefault
End Sub

Private Sub imgFec_Click(Index As Integer)
    Dim esq As Long
    Dim dalt As Long
    Dim menu As Long
    Dim obj As Object

    Set frmC = New frmCal
    
    esq = imgFec(Index).Left
    dalt = imgFec(Index).Top
        
    Set obj = imgFec(Index).Container
      
      While imgFec(Index).Parent.Name <> obj.Name
            esq = esq + obj.Left
            dalt = dalt + obj.Top
            Set obj = obj.Container
      Wend
    
    menu = Me.Height - Me.ScaleHeight 'ací tinc el heigth del menú i de la toolbar

    frmC.Left = esq + imgFec(Index).Parent.Left + 30
    frmC.Top = dalt + imgFec(Index).Parent.Top + imgFec(Index).Height + menu - 40

    If Index < 2 Then
        imgFec(0).Tag = Index '<===
        ' *** repasar si el camp es txtAux o Text1 ***
        If Text1(Index + 1).Text <> "" Then frmC.NovaData = Text1(Index + 1).Text
    Else
        imgFec(0).Tag = Index '<===
        ' *** repasar si el camp es txtAux o Text1 ***
        If Text1(Index + 8).Text <> "" Then frmC.NovaData = Text1(Index + 8).Text
    End If
    ' ********************************************

    frmC.Show vbModal
    Set frmC = Nothing
    ' *** repasar si el camp es txtAux o Text1 ***
    If Index < 2 Then
        PonerFoco Text1(CByte(imgFec(0).Tag) + 1) '<===
    Else
        PonerFoco Text1(CByte(imgFec(0).Tag) + 8) '<===
    End If
    ' ********************************************
End Sub


Private Sub imgZoom_Click(Index As Integer)
    
    Set frmZ = New frmZoom

    If Index = 0 Then
        Indice = 2
        frmZ.pTitulo = "Observaciones del Albarán"
        frmZ.pValor = Text1(Indice).Text
        frmZ.pModo = Modo
    
        frmZ.Show vbModal
        Set frmZ = Nothing
            
        PonerFoco Text1(Indice)
    End If
End Sub

Private Sub mnBuscar_Click()
    SSTab1.Tab = 0
    ModoCuadrilla = True
    BotonBuscar
End Sub

Private Sub mnConfirmacion_Click()
    Confirmacion = True
    ModoCuadrilla = False
    SSTab1.Tab = 0
    If BLOQUEADesdeFormulario(Me) Then
        BotonModificar
    End If
End Sub

Private Sub mnCuadrilla_Click()

    If Data1.Recordset.EOF Then
        MsgBox "No existe parte para introducir datos de la cuadrilla. Revise.", vbExclamation
        Exit Sub
    End If
    
    ModoCuadrilla = True
    
    CuadrillaAnt = 0
    If Text1(32).Text <> "" Then CuadrillaAnt = CLng(Text1(32).Text)
    
    SSTab1.Tab = 1

    PonerModo 4
    Text1(35).Text = Text1(0).Text
    Text1(35).Tag = Text1(0).Tag
    
    PonerFoco Text1(32)

End Sub

Private Sub mnEliminar_Click()
    If Modo = 5 Then 'Eliminar lineas de Pedido
'         BotonEliminarLinea
    Else   'Eliminar Pedido
         BotonEliminar
    End If
End Sub

Private Sub mnImprimir_Click()
'Imprimir Factura
    
    If Data1.Recordset.EOF Then Exit Sub
    
    BotonImprimir
End Sub

Private Sub mnInsercionGastos_Click()
    ' insercion de gastos
    AbrirListadoADV 3
    
End Sub

Private Sub mnNuevo_Click()
    
    SSTab1.Tab = 0
    BotonAnyadir
End Sub


Private Sub mnModificar_Click()
    Confirmacion = False
    ModoCuadrilla = False
    
    SSTab1.Tab = 0
    
    If Modo = 5 Then 'Modificar lineas
        'bloquea la tabla cabecera de albaranes: scaalb
        If BLOQUEADesdeFormulario(Me) Then
            If BloqueaLineasAlb Then BotonModificarLinea (1)
        End If
         
    Else   'Modificar albaran
        'bloquea la tabla cabecera de albaranes: scaalb
        If BLOQUEADesdeFormulario(Me) Then
            BotonModificar
        End If
    End If
End Sub


Private Function BloqueaLineasAlb() As Boolean
'bloquea todas las lineas de la factura
Dim Sql As String

    On Error GoTo EBloqueaLin

    BloqueaLineasAlb = False
    'bloquear cabecera albaranes
    Sql = "select * FROM advpartes_lineas "
    Sql = Sql & ObtenerWhereCP(True) & " FOR UPDATE"
    conn.Execute Sql, , adCmdText
    BloqueaLineasAlb = True

EBloqueaLin:
    If Err.Number <> 0 Then BloqueaLineasAlb = False
End Function

Private Sub mnPrecios_Click()
    ' asignacion de precios segun el tratamiento ie el tipo de venta
    AbrirListadoADV 2
    
End Sub

Private Sub mnSalir_Click()
    Screen.MousePointer = vbDefault
    If (Modo = 5) Then 'Modo 5: Mto Lineas
        '1:Insertar linea, 2: Modificar
        If ModificaLineas = 1 Or ModificaLineas = 2 Then cmdCancelar_Click
        cmdRegresar_Click
        Exit Sub
    End If
    Unload Me
End Sub

Private Sub mnVerTodos_Click()
    BotonVerTodos
End Sub



Private Sub Text1_Change(Index As Integer)
    If Index = 9 Then HaCambiadoCP = True 'Cod. Postal
End Sub

'### A mano
'Los metodos del text tendran que estar
'Los descomentamos cuando esten puestos ya los controles
Private Sub Text1_GotFocus(Index As Integer)
    kCampo = Index
'    If Index = 9 Then HaCambiadoCP = False 'CPostal
'    If Index = 1 And Modo = 1 Then
'        SendKeys "{tab}"
'        Exit Sub
'    End If
    ConseguirFoco Text1(Index), Modo
End Sub


Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
    KEYdown KeyCode
End Sub


Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index <> 2 Or (Index = 2 And Text1(2).Text = "") Then KEYpress KeyAscii
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
Dim devuelve As String
Dim cadMen As String
Dim Sql As String
Dim Nregs As Long

        
    If Not PerderFocoGnral(Text1(Index), Modo) Then Exit Sub
        
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
          
    'Si queremos hacer algo ..
    Select Case Index
        Case 1 'Fecha albaran
            '[Monica]28/08/2013: comprobamos que la fecha esté en la campaña
            If Text1(Index).Text <> "" Then PonerFormatoFecha Text1(Index), True
            
        Case 6 'Socio
            If Modo = 1 Then Exit Sub
            
            If PonerFormatoEntero(Text1(Index)) Then
                If Modo = 1 Then
                    Text2(Index).Text = PonerNombreDeCod(Text1(Index), "rsocios", "nomsocio")
                Else
                    PonerDatosSocios (Text1(Index).Text)
                    If Text2(Index).Text = "" Then
                        cadMen = "No existe el Socio: " & Text1(Index).Text & vbCrLf
                        cadMen = cadMen & "¿Desea crearlo?" & vbCrLf
                        If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                            Set frmSoc = New frmManSocios
                            frmSoc.DatosADevolverBusqueda = "0|1|"
                            Text1(Index).Text = ""
                            TerminaBloquear
                            frmSoc.Show vbModal
                            Set frmSoc = Nothing
                            If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
                        Else
                            Text1(Index).Text = ""
                        End If
                        PonerFoco Text1(Index)
                        
                    Else
                        If EsSocioDeSeccion(Text1(Index).Text, vParamAplic.SeccionADV) Then
                            If EstaSocioDeAlta(Text1(Index)) Then
                                If vParamAplic.Cooperativa = 3 Then
                                Else
                                    PonerCamposSocio
                                End If
                            Else
                                MsgBox "El socio está dado de baja. Reintroduzca.", vbExclamation
                                Text1(Index).Text = ""
                                PonerFoco Text1(Index)
                            End If
                        Else
                            MsgBox "El socio no es de la sección de ADV. Reintroduzca.", vbExclamation
                            Text1(Index).Text = ""
                            PonerFoco Text1(Index)
                        End If
                    End If
               End If
            End If
                
            
        Case 3 ' Tratamiento
            If Modo = 1 Then Exit Sub
            
            If Text1(Index).Text <> "" Then
                Text1(Index).Text = UCase(Text1(Index).Text)
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), "advtrata", "nomtrata", "codtrata", "T")
                If Text2(Index).Text = "" Then
                    '[Monica]18/05/2012
                    If vParamAplic.Cooperativa = 3 Then
                        cadMen = "No existe el Tipo: " & Text1(Index).Text & vbCrLf
                    Else
                        cadMen = "No existe el Tratamiento: " & Text1(Index).Text & vbCrLf
                    End If
                    cadMen = cadMen & "¿Desea crearlo?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmTra = New frmADVTratamientos
                        frmTra.DatosADevolverBusqueda = "0|1|"
                        Text1(Index).Text = ""
                        TerminaBloquear
                        frmTra.Show vbModal
                        Set frmTra = Nothing
                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
                    Else
                        Text1(Index).Text = ""
                    End If
                    PonerFoco Text1(Index)
                End If
            End If
            
         Case 4 ' campo
            If PonerFormatoEntero(Text1(Index)) Then
                PonerDatosCampo Text1(Index)
            End If
         
         Case 7 'litros previstos
            If Modo = 1 Then Exit Sub
            If Text1(Index).Text <> "" Then PonerFormatoEntero Text1(Index)
            
         Case 8 'litros reales
            If Modo = 1 Then Exit Sub
            If PonerFormatoEntero(Text1(Index)) Then
                If Confirmacion Then cmdAceptar.SetFocus
            End If
            
         Case 32 ' codigo de cuadrilla
            If Text1(32).Text <> "" Then
                Text2(32).Text = DevuelveValor("select nomcapat from rcapataz inner join rcuadrilla on rcapataz.codcapat = rcuadrilla.codcapat and rcuadrilla.tipocuadrilla = 1 and rcuadrilla.codcuadrilla = " & DBSet(Text1(32).Text, "N"))
                If Text2(32).Text = "0" Then
                    MsgBox "No existe la cuadrilla o no es del tipo adv. Revise.", vbExclamation
                Else
                    ' dependiendo del modo insertamos o no los trabajadores
                    Sql = "select count(*) from advpartes_trabajador where numparte = " & DBSet(Text1(0).Text, "N")
                    If TotalRegistros(Sql) <> 0 Then
                        If Text1(32).Text <> CuadrillaAnt Then
                            MsgBox "Hay trabajadores en este parte. Se eliminaran y se añadiran los de esta cuadrilla", vbExclamation
                        End If
                    End If
                    ' si no hay numero de trabajadores introducidos metemos los de la cuadrilla
                    If Text1(33).Text = "" Then
                        Sql = "select count(*) from rcuadrilla_trabajador where codcuadrilla = " & DBSet(Text1(32).Text, "N")
                        Text1(33).Text = TotalRegistros(Sql)
                        If Text1(33).Text = "0" Then Text1(33).Text = ""
                    End If
                End If
            End If
        
        Case 33 ' nro de trabajadores
            PonerFormatoEntero Text1(Index)
            
        Case 34 ' nro de horas
            If PonerFormatoDecimal(Text1(Index), 3) Then PonerFocoBtn cmdAceptar
            
    End Select
End Sub


Private Sub HacerBusqueda()
Dim CadB As String
Dim cadAux As String
    
'    '--- Laura 12/01/2007
'    cadAux = Text1(5).Text
'    If Text1(4).Text <> "" Then Text1(5).Text = ""
'    '---
    
'    '--- Laura 12/01/2007
'    Text1(5).Text = cadAux
'    '---
    
'--monica
'    CadB = ObtenerBusqueda(Me)
'++monica
    If Albaran = "" Then
        CadB = ObtenerBusqueda2(Me, BuscaChekc, 1)
    Else
        CadB = "numalbar = " & Albaran & " "
    End If

    If chkVistaPrevia = 1 Then
        EsCabecera = True
        MandaBusquedaPrevia CadB
    ElseIf CadB <> "" Then
        'Se muestran en el mismo form
        CadenaConsulta = "select advpartes.* from " & NombreTabla
        CadenaConsulta = CadenaConsulta & " WHERE " & CadB & Ordenacion
        PonerCadenaBusqueda
    End If
End Sub


Private Sub MandaBusquedaPrevia(CadB As String)
'Carga el formulario frmBuscaGrid con los valores correspondientes
Dim cad As String
Dim tabla As String
Dim Titulo As String
Dim Desc As String, devuelve As String
    'Llamamos a al form
    '##A mano
    cad = ""
    cad = cad & "Nº.Parte|advpartes.numparte|N||15·"
    cad = cad & "Socio|advpartes.codsocio|N||10·" 'ParaGrid(Text1(3), 10, "Socio")
    cad = cad & "Nombre Socio|rsocios.nomsocio|N||60·"
    cad = cad & ParaGrid(Text1(1), 15, "F.Parte")
    tabla = NombreTabla & " INNER JOIN rsocios ON advpartes.codsocio=rsocios.codsocio "
    
    Titulo = "Partes"
    devuelve = "0|"
           
    If cad <> "" Then
        Screen.MousePointer = vbHourglass
        Set frmB = New frmBuscaGrid
        frmB.vCampos = cad
        frmB.vtabla = tabla
        frmB.vSQL = CadB
        HaDevueltoDatos = False
        '###A mano
        frmB.vDevuelve = "0|"
        frmB.vDevuelve = devuelve
        frmB.vTitulo = Titulo
        frmB.vSelElem = 0
'        frmB.vConexionGrid = cAgro  'Conexión a BD: Ariagro
        If Not EsCabecera Then frmB.Label1.FontSize = 11
'        frmB.vBuscaPrevia = chkVistaPrevia
        '#
        frmB.Show vbModal
        Set frmB = Nothing
'        If EsCabecera Then
'            PonerCadenaBusqueda
'            Text1(0).Text = Format(Text1(0).Text, "0000000")
'        End If
        'Si ha puesto valores y tenemos que es formulario de busqueda entonces
        'tendremos que cerrar el form lanzando el evento
        If HaDevueltoDatos Then
            If (Not Data1.Recordset.EOF) And DatosADevolverBusqueda <> "" Then _
                cmdRegresar_Click
        Else   'de ha devuelto datos, es decir NO ha devuelto datos
            PonerFoco Text1(kCampo)
        End If
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub PonerCadenaBusqueda()
    Screen.MousePointer = vbHourglass

    On Error GoTo EEPonerBusq

    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    If Data1.Recordset.RecordCount <= 0 Then
        MsgBox "No hay ningún registro en la tabla " & NombreTabla, vbInformation
        Screen.MousePointer = vbDefault
        If Modo = 1 Then
            PonerFoco Text1(kCampo)
'            Text1(0).BackColor = vbLightBlue
        End If
        Exit Sub
    Else
        Data1.Recordset.MoveFirst
        PonerModo 2
        '--monica
        'LLamaLineas Modo, 0, "DataGrid2"
        PonerCampos
    End If


    Screen.MousePointer = vbDefault
    Exit Sub
    
EEPonerBusq:
    MuestraError Err.Number, "PonerCadenaBusqueda"
    PonerModo 0
    Screen.MousePointer = vbDefault
End Sub


Private Sub PonerCamposLineas()
'Carga el grid de los AlbaranesxFactura, es decir, la tabla scafac1 de la factura seleccionada
Dim B As Boolean
Dim b2 As Boolean
Dim I As Integer

    On Error GoTo EPonerLineas

    If Data1.Recordset.EOF Then Exit Sub

    Screen.MousePointer = vbHourglass
    
    If Data1.Recordset.RecordCount > 0 Then
        CargaGrid DataGrid3, Adoaux(1), True
        CargaGrid DataGrid1, Adoaux(0), True
    Else
        CargaGrid DataGrid3, Adoaux(1), False
        CargaGrid DataGrid1, Adoaux(0), False
    End If
    If Not Adoaux(1).Recordset.EOF Then
        Text2(16).Text = DBLet(Adoaux(1).Recordset!ampliaci, "T")
    Else
        Text2(16).Text = ""
    End If
    
    Screen.MousePointer = vbDefault
    Exit Sub
    
EPonerLineas:
    MuestraError Err.Number, "PonerCamposLineas"
    PonerModo 2
    Screen.MousePointer = vbDefault
End Sub


Private Sub PonerCampos()
Dim BrutoFac As Single
Dim B As Boolean

    On Error Resume Next

    If Data1.Recordset.EOF Then Exit Sub
    
    B = PonerCamposForma2(Me, Data1, 2, "Frame2")
    B = PonerCamposForma2(Me, Data1, 2, "Frame3")
    
    Text1(35).Text = Text1(0).Text
    
    'poner descripcion campos
    Modo = 4
    
    Text2(6).Text = PonerNombreDeCod(Text1(6), "rsocios", "nomsocio", "codsocio", "N") 'socio
    Text2(3).Text = DevuelveDesdeBDNew(cAgro, "advtrata", "nomtrata", "codtrata", Text1(3), "T") 'tratamiento
    Text2(32).Text = ""
    If Text1(32).Text <> "" Then
        Text2(32).Text = DevuelveValor("select nomcapat from rcapataz inner join rcuadrilla on rcapataz.codcapat = rcuadrilla.codcapat and rcuadrilla.codcuadrilla = " & DBSet(Text1(32).Text, "N"))
    End If
    
    PonerDatosCampo Text1(4).Text
'    Text2(4).Text = DevuelveDesdeBDNew(cAgro, "destinos", "nomdesti", "codclien", Text1(3).Text, "N", , "coddesti", Text1(6).Text, "N")
    
'    Text2(18).Text = PonerNombreDeCod(Text1(16), "salmpr", "nomalmac", "codalmac", "N") 'almacen
    
    Modo = 2
    
    CalcularDatosAlbaran
    
    PonerCamposLineas 'Pone los datos de las tablas de lineas de Ofertas
    
    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
    
    PonerModoOpcionesMenu (Modo) 'Activar opciones de menu según modo
    PonerOpcionesMenu 'Activar opciones de menu según nivel de permisos del usuario

    If Err.Number <> 0 Then Err.Clear
End Sub


'----------------------------------------------------------------
'----------------------------------------------------------------
'   En PONERMODO se habilitan, o no, los diverso campos del
'   formulario en funcion del modo en k vayamos a trabajar
Private Sub PonerModo(Kmodo As Byte, Optional indFrame As Integer)
Dim I As Byte, NumReg As Byte
Dim B As Boolean

    On Error GoTo EPonerModo

    'Actualiza Iconos Insertar,Modificar,Eliminar
    '## No tiene el boton modificar y no utiliza la funcion general
'    ActualizarToolbar Modo, Kmodo
    BuscaChekc = ""
    
    Modo = Kmodo
    PonerIndicador lblIndicador, Modo
    
    'Modo 2. Hay datos y estamos visualizandolos
    '=========================================
    B = (Modo = 2)
    'Ponemos visible, si es formulario de busqueda, el boton regresar cuando hay datos
    If DatosADevolverBusqueda <> "" Or Albaran <> "" Then
        cmdRegresar.visible = B
    Else
        cmdRegresar.visible = False
    End If
        
    'Poner Flechas de desplazamiento visibles
    NumReg = 1
    If Not Data1.Recordset.EOF Then
        If Data1.Recordset.RecordCount > 1 Then NumReg = 2 'Solo es para saber q hay + de 1 registro
    End If
'    DesplazamientoVisible Me.Toolbar1, btnPrimero, B, NumReg
    DesplazamientoVisible B And Data1.Recordset.RecordCount > 1
    
          
    Frame3.Enabled = ModoCuadrilla
'    Frame4.Enabled = ModoCuadrilla
        
    'Bloquea los campos Text1 sino estamos modificando/Insertando Datos
    'Si estamos en Insertar además limpia los campos Text1
    'si estamos en modificar bloquea las compos que son clave primaria
    BloquearText1 Me, Modo
    For I = 9 To 31
        BloquearTxt Text1(I), Not (Modo = 1)
        Text1(I).Enabled = (Modo = 1)
    Next I
    
    B = (Modo <> 1)
    'Campos Nº Albarán bloqueado y en azul
    BloquearTxt Text1(0), B, True
    
    B = (Modo <> 1) And (Modo <> 3)
    
'    BloquearTxt Text1(1), b 'fechafactura
    BloquearTxt Text1(6), B 'socio
    BloquearTxt Text1(3), B  'tratamiento
'    BloquearTxt Text1(4), b 'campo lo puedo modificar
    
    If vParamAplic.Cooperativa = 3 And Modo = 4 Then
        BloquearTxt Text1(3), False   'tratamiento
        BloquearTxt Text1(6), False   'socio
    End If
    
    
    BloquearChk Me.Check1(0), (Modo = 0 Or Modo = 2)
    
    BloquearTxt Text1(8), (Modo <> 1)
    
    If Modo = 4 And Confirmacion Then
        For I = 0 To 7
            BloquearTxt Text1(I), True
        Next I
        BloquearTxt Text1(8), False
'    Else
'        '[Monica] 30/09/2009: Sólo dejamos modificar observaciones si no tiene litros reales
'        If Modo = 4 And Not Data1.Recordset.EOF Then
'            If DBLet(Data1.Recordset!LitrosRea, "N") <> 0 Then
'                For i = 0 To 8
'                    BloquearTxt Text1(i), True
'                Next i
'                BloquearTxt Text1(2), False
'            End If
'        End If
    End If
    
    If Modo = 4 And ModoCuadrilla Then
        For I = 0 To 7
            BloquearTxt Text1(I), True
        Next I
    End If
    
    Me.imgZoom(0).Enabled = Not (Modo = 0)
    
    'Si no es modo lineas Boquear los TxtAux
    For I = 1 To 1
        txtAux(I).visible = False
        BloquearTxt txtAux(I), True
    Next I
    For I = 3 To 10
        txtAux(I).visible = False
        BloquearTxt txtAux(I), True
    Next I
    
    For I = 0 To 0
        Text2(I).visible = ((Modo = 5) And (indFrame = 1))
        Text2(I).Enabled = False
    Next I
    
    For I = 0 To txtAux1.Count - 1
        txtAux1(I).visible = False
        txtAux1(I).Enabled = False
        'BloquearTxt txtAux1(i), True
    Next I
    
    
    BloquearTxt Text2(16), (Modo <> 5)
    
    BloquearBtn Me.btnBuscar(0), True
    
    '---------------------------------------------
    B = (Modo <> 0 And Modo <> 2)
    cmdCancelar.visible = B
    cmdAceptar.visible = B
    
    BloquearImgBuscar Me, Modo, ModificaLineas
    BloquearImgFec Me, 0, Modo
                    
                    
    '[Monica]18/05/2012: cambiamos la apariencia del formulario para Moixent
    If vParamAplic.Cooperativa = 3 Then
        imgBuscar(1).visible = False
        imgBuscar(1).Enabled = False
    End If
                    
                    
    ' si estamos modificando
    If Modo = 4 Then
        For I = 0 To 2
            If I <> 1 Then
                Me.imgBuscar(I).Enabled = False
                Me.imgBuscar(I).visible = False
            Else
                If Confirmacion Then 'Or DBLet(Data1.Recordset!LitrosRea, "N") <> 0 Then
                    Me.imgBuscar(I).Enabled = False
                    Me.imgBuscar(I).visible = False
                End If
            End If
        Next I
        
'        imgFec(0).Enabled = False
'        imgFec(0).visible = False
        If vParamAplic.Cooperativa = 3 Then
            Me.imgBuscar(0).Enabled = True
            Me.imgBuscar(0).visible = True
            Me.imgBuscar(2).Enabled = True
            Me.imgBuscar(2).visible = True
        End If
    End If
                    
    Me.chkVistaPrevia.Enabled = (Modo <= 2)
       
    ' *** si n'hi han llínies i algún tab que no te datagrid ***
    BloquearFrameAux Me, "FrameAux1", Modo, 1
    BloquearFrameAux Me, "FrameAux0", Modo, 0
    
    ' ***************************
    'Poner el tamaño de los campos. Si es modo Busqueda el MaxLength del campo
    'debe ser mayor para adminir intervalos de busqueda.
    PonerLongCampos
    PonerModoOpcionesMenu (Modo) 'Activar opciones de menu según modo
    PonerOpcionesMenu 'Activar opciones de menu según nivel de permisos del usuario
    
EPonerModo:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub DesplazamientoVisible(bol As Boolean)
    FrameDesplazamiento.visible = bol
    FrameDesplazamiento.Enabled = bol
End Sub


Private Sub PonerLongCampos()
'Modificar el MaxLength del campo en funcion de si es modo de búsqueda o no
'para los campos que permitan introducir criterios más largos del tamaño del campo
    PonerLongCamposGnral Me, Modo, 1
End Sub


Private Function DatosOK() As Boolean
'Comprobar que los datos de la cabecera son correctos antes de Insertar o Modificar
'la cabecera del Pedido
Dim B As Boolean
Dim Serie As String
Dim Sql As String

    On Error GoTo EDatosOK

    DatosOK = False
    
'    ComprobarDatosTotales
    If (Modo = 3 Or Modo = 4) And vParamAplic.Cooperativa = 3 Then
        Text1(4).Text = 0 ' codigo de campo no puede ser nulo
        Text1(7).Text = 0 ' litros previstos
        Text1(8).Text = 0 ' litros reales
    End If


    'comprobamos datos OK de la tabla scaalb
    B = CompForm2(Me, 2, "Frame2") ' , 1) 'Comprobar formato datos ok de la cabecera: opcion=1
    If Not B Then Exit Function
    
    B = CompForm2(Me, 2, "Frame3") ' , 1) 'Comprobar formato datos ok de la cabecera: opcion=1
    If Not B Then Exit Function
    
    If B And ModoCuadrilla Then
        If Text1(32).Text <> "" Then
            Sql = "select count(*) from rcuadrilla where codcuadrilla = " & DBSet(Text1(32).Text, "N") & " and tipocuadrilla = 1"
            If TotalRegistros(Sql) = 0 Then
                MsgBox "No existe la cuadrilla o no es del tipo adv. Revise.", vbExclamation
                B = False
            End If
        End If
    End If
    If B Then
        If Text1(7).Text = "" Then
            MsgBox "Los litros previstos no pueden ser nulos. Revise.", vbExclamation
            B = False
            PonerFoco Text1(7)
        End If
    End If
    
    If B And Confirmacion Then
        If ComprobarCero(Text1(8).Text) = "0" Then
            MsgBox "Los litros reales deben tener un valor. Revise.", vbExclamation
            B = False
            PonerFoco Text1(8)
        End If
    End If
    
    
    DatosOK = B
EDatosOK:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function


Private Function DatosOkLinea() As Boolean
Dim B As Boolean
Dim I As Byte

    On Error GoTo EDatosOkLinea

    DatosOkLinea = False
    B = True

    For I = 4 To 7
        If I = 4 Or I = 6 Or I = 7 Then
            If txtAux(I).Text = "" Then
                MsgBox "El campo " & txtAux(I).Tag & " no puede ser nulo", vbExclamation
                B = False
                PonerFoco txtAux(I)
                Exit Function
            End If
        End If
    Next I
            
    DatosOkLinea = B
    
EDatosOkLinea:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function


Private Sub Text2_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Index = 16 And KeyCode = 40 Then 'campo Amliacion Linea y Flecha hacia abajo
        PonerFocoBtn Me.cmdAceptar
    Else
        KEYdown KeyCode
    End If
End Sub

Private Sub Text2_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 16 And KeyAscii = 13 Then 'campo Amliacion Linea y ENTER
        PonerFocoBtn Me.cmdAceptar
    End If
End Sub

Private Sub ToolAux_ButtonClick(Index As Integer, ByVal Button As MSComctlLib.Button)

    If BloqueaRegistro(NombreTabla, "numparte= " & Data1.Recordset!Numparte) Then
'    If BLOQUEADesdeFormulario2(Me, Data1, 1) Then
        Select Case Button.Index
            Case 1
                BotonAnyadirLinea Index
            Case 2
                BotonModificarLinea Index
            Case 3
                BotonEliminarLinea Index
            Case Else
        End Select
    End If

End Sub


Private Sub BotonEliminarLinea(Index As Integer)
Dim cad As String
Dim Sql As String
Dim Mens As String
Dim B As Boolean

    On Error GoTo EEliminarLinea

    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub

    ' *** repasar el nom de l'adodc, l'index del Field i el camp que te la PK ***
    ' ### [Monica] 26/09/2006 dejamos modificar y eliminar el codigo 0
    'El registre de codi 0 no es pot Modificar ni Eliminar
    'If EsCodigoCero(CStr(Data1.Recordset.Fields(0).Value), FormatoCampo(text1(0))) Then Exit Sub
    ' ***************************************************************************

    B = True

    Select Case Index

        Case 1
            ' *************** canviar la pregunta ****************
            cad = "¿Seguro que desea eliminar el Artículo?"
            cad = cad & vbCrLf & "Parte: " & Adoaux(1).Recordset.Fields(0)
            cad = cad & vbCrLf & "Artículo: " & Adoaux(1).Recordset.Fields(3) & " - " & Adoaux(1).Recordset.Fields(4)
            
            If MsgBox(cad, vbQuestion + vbYesNo) = vbYes Then
                On Error GoTo EEliminarLinea
                Screen.MousePointer = vbHourglass
                NumRegElim = Adoaux(1).Recordset.AbsolutePosition
                
                If Not EliminarLinea Then
                    Screen.MousePointer = vbDefault
                    Exit Sub
                Else
                    CalcularDatosAlbaran
                    If SituarDataTrasEliminar(Adoaux(1), NumRegElim) Then
                        PonerCampos
                    Else
                        PonerCampos
        '                        LimpiarCampos
        '                        PonerModo 0
                    End If
                End If
            End If
        Case 0
             ' *************** canviar la pregunta ****************
            cad = "¿Seguro que desea eliminar el Trabajador del Parte?"
            cad = cad & vbCrLf & "Parte: " & Adoaux(0).Recordset.Fields(0)
            cad = cad & vbCrLf & "Trabajador: " & Adoaux(0).Recordset.Fields(2) & " - " & Adoaux(0).Recordset.Fields(3)
            
            If MsgBox(cad, vbQuestion + vbYesNo) = vbYes Then
                On Error GoTo EEliminarLinea
                Screen.MousePointer = vbHourglass
                NumRegElim = Adoaux(0).Recordset.AbsolutePosition
                
                If Not EliminarLineaTrab Then
                    Screen.MousePointer = vbDefault
                    Exit Sub
                Else
                    TerminaBloquear
                    If SituarDataTrasEliminar(Adoaux(0), NumRegElim) Then
                        PonerCampos
                    Else
                        PonerCampos
                    End If
                End If
            End If
       
        
   End Select
    
    
    Screen.MousePointer = vbDefault
       
EEliminarLinea:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Or Not B Then MuestraError Err.Number, "Eliminar Linea de Parte", Err.Description

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 5  'Buscar
            mnBuscar_Click
        Case 6  'Todos
            BotonVerTodos
        Case 1  'Añadir
            mnNuevo_Click
        Case 2  'Modificar
            mnModificar_Click
        Case 3  'Borrar
            mnEliminar_Click
'        Case 8  ' Confirmacion
'            mnConfirmacion_Click
'        Case 9  ' Cuadrilla
'            mnCuadrilla_Click
'        Case 10 ' Asignacion de precios
'            mnPrecios_Click
'        Case 11 ' insercion de gastos
'            mnInsercionGastos_Click
        Case 8  ' Impresion de albaran
            mnImprimir_Click
    End Select
End Sub


Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
End Sub


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me
End Sub


Private Sub ActualizarToolbar(Modo As Byte, Kmodo As Byte)
'Modo: Modo antiguo
'Kmodo: Modo que se va a poner

    If (Modo = 5) And (Kmodo <> 5) Then
        'El modo antigu era modificando las lineas
        'Luego hay que reestablecer los dibujitos y los TIPS
        '-- insertar
'        Toolbar1.Buttons(5).Image = 3
'        Toolbar1.Buttons(5).ToolTipText = "Nuevo Albaran"
        '-- Modificar
        Toolbar1.Buttons(5).Image = 4
        Toolbar1.Buttons(5).ToolTipText = "Modificar Factura"
        '-- eliminar
        Toolbar1.Buttons(6).Image = 5
        Toolbar1.Buttons(6).ToolTipText = "Eliminar Factura"
    End If
    If Kmodo = 5 Then
        'Ponemos nuevos dibujitos y tal y tal
        'Luego hay que reestablecer los dibujitos y los TIPS
        '-- insertar
'        Toolbar1.Buttons(5).Image = 12
'        Toolbar1.Buttons(5).ToolTipText = "Nueva linea"
        '-- Modificar
        Toolbar1.Buttons(5).Image = 13
        Toolbar1.Buttons(5).ToolTipText = "Modificar linea factura"
        '-- eliminar
        Toolbar1.Buttons(6).Image = 14
        Toolbar1.Buttons(6).ToolTipText = "Eliminar linea factura"
    End If
End Sub
    


Private Sub CargaGrid(ByRef vDataGrid As DataGrid, ByRef vData As Adodc, enlaza As Boolean)
Dim B As Boolean
Dim Opcion As Byte
Dim Sql As String

    On Error GoTo ECargaGrid

    B = DataGrid3.Enabled
    Select Case vDataGrid.Name
        Case "DataGrid3" 'envases
            Opcion = 1
        Case "DataGrid1" ' cuadrilla
            Opcion = 2
    End Select
    
    Sql = MontaSQLCarga(enlaza, Opcion)
    CargaGridGnral vDataGrid, vData, Sql, PrimeraVez
    
    vDataGrid.RowHeight = 270
    
    CargaGrid2 vDataGrid, vData
    vDataGrid.ScrollBars = dbgAutomatic
    
    B = (Modo = 5) And (ModificaLineas = 1 Or ModificaLineas = 2)
    vDataGrid.Enabled = Not B
    
    If vDataGrid.Name = "DataGrid1" Then CalcularTotales

    Exit Sub
    
ECargaGrid:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid", Err.Description
End Sub


Private Sub CargaGrid2(ByRef vDataGrid As DataGrid, ByRef vData As Adodc)
Dim tots As String
    
    On Error GoTo ECargaGrid

    Select Case vDataGrid.Name
        Case "DataGrid3" 'slialb lineas de envases
'select codtipom,numfactu,fecfactu,numlinea,codartic,sartic.nomartic,cantidad,precioar,dtolinea,importel,ampliaci,codigiva
           tots = "N||||0|;N||||0|;S|txtAux(4)|T|Alm|600|;"
           tots = tots & "S|txtAux(5)|T|Articulo|2000|;S|btnBuscar(0)|B|||;"
           
           '[Monica]18/05/2012
           If vParamAplic.Cooperativa = 3 Then
               tots = tots & "S|Text2(0)|T|Nombre|4000|;S|txtAux(8)|T|Bultos|1360|;S|txtAux(6)|T|Cantidad|1500|;"
           Else
               tots = tots & "S|Text2(0)|T|Nombre|4000|;S|txtAux(8)|T|Dosis Hab|1360|;S|txtAux(6)|T|Cantidad|1500|;"
           End If
       
           tots = tots & "S|txtAux(7)|T|Precio|1400|;S|txtAux(9)|T|Importe|1700|;N||||0|;N||||0|;"
           arregla tots, DataGrid3, Me, 350
    
    
        Case "DataGrid1" 'trabajadores de la cuadrilla
'select codtipom,numfactu,fecfactu,numlinea,codartic,sartic.nomartic,cantidad,precioar,dtolinea,importel,ampliaci,codigiva
           tots = "N||||0|;N||||0|;"
           tots = tots & "S|txtAux1(2)|T|Codigo|1200|;S|btnBuscar(1)|B|||;"
           tots = tots & "S|txtAux1(3)|T|Trabajador|6400|;S|txtAux1(4)|T|Horas|1500|;S|txtAux1(5)|T|Precio|1700|;"
           tots = tots & "S|txtAux1(6)|T|Importe|1750|;"
           arregla tots, DataGrid1, Me, 350
    End Select
    
    vDataGrid.HoldFields
    Exit Sub
    
ECargaGrid:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid", Err.Description
End Sub

Private Sub Toolbar5_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1  ' Confirmacion
            mnConfirmacion_Click
        Case 2  ' Cuadrilla
            mnCuadrilla_Click
        Case 3 ' Asignacion de precios
            mnPrecios_Click
        Case 4 ' insercion de gastos
            mnInsercionGastos_Click
    End Select
End Sub

Private Sub ToolbarDes_ButtonClick(ByVal Button As MSComctlLib.Button)
    Desplazamiento (Button.Index)
End Sub

Private Sub txtaux_GotFocus(Index As Integer)
    ConseguirFocoLin txtAux(Index)
End Sub

Private Sub txtAux_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
    KEYdown KeyCode
End Sub

Private Sub txtaux_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtAux_LostFocus(Index As Integer)
Dim cadMen As String
Dim Sql As String
Dim devuelve As String
Dim B As Boolean
Dim TipoDto As Byte
Dim vCstock As CStockADV
Dim OtroCampo As String


    'Quitar espacios en blanco
    If Not PerderFocoGnralLineas(txtAux(Index), ModificaLineas) Then Exit Sub
    
    Select Case Index
        Case 4 'almacen
            If txtAux(Index).Text <> "" Then PonerFormatoEntero txtAux(Index)
        
        Case 5 'articulo
            If txtAux(Index).Text = "" Then
                Exit Sub
            End If
        
            If txtAux(4).Text = "" Then
                MsgBox "Debe seleccionar un almacen.", vbInformation
                PonerFoco txtAux(4)
                Exit Sub
            End If
        
            devuelve = ""
            If ModificaLineas = 2 Then
                If Not Adoaux(1).Recordset.EOF Then devuelve = Adoaux(1).Recordset!codArtic
            End If
        
            If Not PonerArticulo(txtAux(5), Text2(0), txtAux(4).Text, CodTipoMov, ModificaLineas, devuelve) Then
                PonerFoco txtAux(Index)
            Else
                txtAux(7) = DevuelveDesdeBDNew(cAgro, "advartic", "preciove", "codartic", txtAux(5), "T", OtroCampo)
                txtAux(10).Text = DevuelveDesdeBDNew(cAgro, "advartic", "codigiva", "codartic", txtAux(5), "T")

                If vParamAplic.Cooperativa = 3 Then
                    UniCajas = DevuelveDesdeBDNew(cAgro, "advartic", "unicajas", "codartic", txtAux(5).Text, "T")
                    PonerFoco txtAux(8)
                End If

'--monica: preguntar a manolo
'                If Combo1(0).ListIndex = 1 Then
'                    txtAux(10).Text = vParamAplic.CodIvaExento
'                Else
'                    txtAux(10).Text = DevuelveDesdeBDNew(cAgro, "sartic", "codigiva", "codartic", txtAux(5), "T")
'                End If
            End If
        
        Case 6 ' Cantidad
            If PonerFormatoDecimal(txtAux(Index), 2) Then  'Tipo 1: Decimal(8,3)
            
'                'Comprobar si hay suficiente stock
'                Set vCstock = New CStockADV
'                If Not InicializarCStock(vCstock, "S") Then Exit Sub
'                If vCstock.MueveStock Then 'Comprobar si el articulo mueve stock: tiene control de stock y no es instalacion
'                  If Not vCstock.MoverStock Then
'                    PonerFoco txtAux(Index)
'                    Set vCstock = Nothing
'                    Exit Sub
'                  End If
'                End If
'
'                Set vCstock = Nothing
            End If
            
        Case 7 ' Precio
            If PonerFormatoDecimal(txtAux(Index), 11) Then   'Tipo 11:decimal(10,4)
                PonerFoco Text2(16)
            End If
            
        Case 8  'Dosis habitual
            ' en caso de ser albaranes de venta campo
            If vParamAplic.Cooperativa = 3 Then
                PonerFormatoEntero txtAux(Index)
                txtAux(6).Text = Round(UniCajas * ImporteSinFormato(ComprobarCero(txtAux(8).Text)), 2)
                PonerFoco txtAux(6)
            Else
                PonerFormatoDecimal txtAux(Index), 12 'Tipo 4: Decimal(6,3)
            End If
            
        Case 9 'Importe Linea
            PonerFormatoDecimal txtAux(Index), 3 'Tipo 3: Decimal(10,2)
            
    End Select
     If (Index = 6 Or Index = 7 Or Index = 9) Then 'Cant., Precio, Importe
'        If Trim(TxtAux(3).Text) = "" Or Trim(TxtAux(4).Text) = "" Then Exit Sub
'        If Trim(TxtAux(6).Text) = "" Or Trim(TxtAux(7).Text) = "" Then Exit Sub
        TipoDto = 0 'DevuelveDesdeBDNew(cAgro, "rsocios", "tipodtos", "codsocio", Text1(6).Text, "N")
                
        
        '[Monica]27/07/2012: segun el tipo de venta es por cantidad o por bulto se calcula el importe
        '                    añadido el if
        Dim TipoVenta As Integer
        If vParamAplic.Cooperativa = 3 Then
            TipoVenta = DevuelveValor("select tipoprecio from advpartes, advtrata where advpartes.codtrata = advtrata.codtrata and advpartes.numparte = " & DBSet(txtAux(1).Text, "N"))
            If TipoVenta = 0 Then
                txtAux(9).Text = CalcularImporte(txtAux(6).Text, txtAux(7).Text, 0, 0, TipoDto, 0)
            Else
                txtAux(9).Text = CalcularImporte(txtAux(8).Text, txtAux(7).Text, 0, 0, TipoDto, 0)
            End If
        Else
            txtAux(9).Text = CalcularImporte(txtAux(6).Text, txtAux(7).Text, 0, 0, TipoDto, 0)
        End If
        
        PonerFormatoDecimal txtAux(9), 3
    End If
    
End Sub




Private Function Eliminar() As Boolean
Dim Sql As String, LEtra As String, Sql2 As String
Dim B As Boolean
Dim vTipoMov As CTiposMov
Dim Mens As String

    On Error GoTo FinEliminar

    If Data1.Recordset.EOF Then Exit Function
        
    conn.BeginTrans
        
    B = EliminarStock

    If B Then
        'Eliminar en tablas de cabecera de albaran
        '------------------------------------------
        Sql = " " & ObtenerWhereCP(True)
        
        'Lineas de articulos (advpartes_lineas)
        conn.Execute "Delete from advpartes_lineas " & Sql
        
        'Lineas de trabajadores de la cuadrilla (advpartes_trabajador)
        conn.Execute "Delete from advpartes_trabajador " & Sql
        
        'Cabecera de factura
        conn.Execute "Delete from " & NombreTabla & Sql
        
        'Decrementar contador si borramos el ult. palet
        Set vTipoMov = New CTiposMov
        vTipoMov.DevolverContador CodTipoMov, Val(Text1(0).Text)
        Set vTipoMov = Nothing
        
        B = True
    End If
FinEliminar:
    If Err.Number <> 0 Or Not B Then
        MuestraError Err.Number, "Eliminar Parte", Err.Description & " " & Mens
        B = False
    End If
    If Not B Then
        conn.RollbackTrans
        Eliminar = False
    Else
        conn.CommitTrans
        Eliminar = True
    End If
End Function

Private Function EliminarLinea() As Boolean
Dim Sql As String, LEtra As String
Dim B As Boolean
Dim vTipoMov As CTiposMov
Dim Mens As String
Dim vCstock As CStockADV

    On Error GoTo FinEliminar

    B = False
    
            
    If Adoaux(1).Recordset.EOF Then Exit Function
        
    conn.BeginTrans
        
    Mens = ""
    
    'Eliminar en tablas de slialb
    '------------------------------------------
    Sql = " where numparte = " & Adoaux(1).Recordset.Fields(0)
    Sql = Sql & " and numlinea = " & Adoaux(1).Recordset.Fields(1)


     ' borramos el movimiento y aumentamos el stock
    Set vCstock = New CStockADV

    If Not InicializarCStock(vCstock, "E", , Adoaux(1).Recordset) Then Exit Function

     'en actualizar stock comprobamos si el articulo tiene control de stock
     B = vCstock.DevolverStock
     Set vCstock = Nothing

    'Lineas de variedades
    
    conn.Execute "Delete from advpartes_lineas " & Sql

    
FinEliminar:
    If Err.Number <> 0 Or Not B Then
        MuestraError Err.Number, "Eliminar Artículos del Parte ", Err.Description & " " & Mens
        B = False
    End If
    If Not B Then
        conn.RollbackTrans
        EliminarLinea = False
    Else
        conn.CommitTrans
        EliminarLinea = True
    End If
End Function


Private Function EliminarLineaTrab() As Boolean
Dim Sql As String, LEtra As String
Dim B As Boolean
Dim vTipoMov As CTiposMov
Dim Mens As String
Dim vCstock As CStockADV

    On Error GoTo FinEliminarTrab

    B = False
    
            
    If Adoaux(0).Recordset.EOF Then Exit Function
        
        
    Mens = ""
    
    'Eliminar en tablas de slialb
    '------------------------------------------
    Sql = " where numparte = " & Adoaux(0).Recordset.Fields(0)
    Sql = Sql & " and numlinea = " & Adoaux(0).Recordset.Fields(1)

    'Lineas de variedades
    
    conn.Execute "Delete from advpartes_trabajador " & Sql

    B = True
FinEliminarTrab:
    If Err.Number <> 0 Or Not B Then
        MuestraError Err.Number, "Eliminar Trabajadores del Parte ", Err.Description & " " & Mens
        B = False
    End If
    If Not B Then
        EliminarLineaTrab = False
    Else
        EliminarLineaTrab = True
    End If
End Function



Private Sub LimpiarDataGrids()
'Pone los Grids sin datos, apuntando a ningún registro
On Error Resume Next

'    CargaGrid DataGrid2, Me.Adoaux(1), False 'variedades
    CargaGrid DataGrid3, Me.Adoaux(1), False 'articulos de adv
    CargaGrid DataGrid1, Me.Adoaux(0), False 'articulos de adv
    
    
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub PosicionarData()
Dim Indicador As String
Dim vWhere As String

    If Not Data1.Recordset.EOF Then
        'Hay datos en el Data1 bien porque se ha hecho VerTodos o una Busqueda
         vWhere = "(" & ObtenerWhereCP(False) & ")"
         If SituarDataMULTI(Data1, vWhere, Indicador) Then
             PonerModo 2
             lblIndicador.Caption = Indicador
        Else
             LimpiarCampos
             'Poner los grid sin apuntar a nada
             LimpiarDataGrids
             PonerModo 0
         End If
    Else
        'El Data esta vacio, desde el modo de inicio se pulsa Insertar
        CadenaConsulta = "Select * from " & NombreTabla & ObtenerWhereCP(True) & Ordenacion
        PonerCadenaBusqueda
    End If
End Sub


Private Function ObtenerWhereCP(conWhere As Boolean) As String
Dim Sql As String

    On Error Resume Next
    
    Sql = "numparte= " & DBSet(Text1(0).Text, "N")
    If conWhere Then Sql = " WHERE " & Sql
    ObtenerWhereCP = Sql
    
    If Err.Number <> 0 Then MuestraError Err.Number, "Obteniendo cadena WHERE.", Err.Description
End Function


Private Function MontaSQLCarga(enlaza As Boolean, Opcion As Byte) As String
'--------------------------------------------------------------------
' MontaSQlCarga:
'   Basándose en la información proporcionada por el vector de campos
'   crea un SQl para ejecutar una consulta sobre la base de datos que los
'   devuelva.
' Si ENLAZA -> Enlaza con el data1
'           -> Si no lo cargamos sin enlazar a ningun campo
'--------------------------------------------------------------------
Dim Sql As String
    
    Select Case Opcion
'        Case 0  'variedades
''select codtipom,numfactu,fecfactu,numlinea,numalbar,numlinealbar,cantreal,cantfact,precibru,precinet,
''dtocom1,dtocom2,imporbru,impornet,codigiva
'            Sql = "SELECT facturas_variedad.codtipom,numfactu,fecfactu,facturas_variedad.numlinea,"
'            Sql = Sql & " facturas_variedad.numalbar,numlinealbar,cantreal,"
'            Sql = Sql & " cantfact,facturas_variedad.unidades, precibru, precinet, imporbru,impornet, fechaalb, matrirem, "
'            Sql = Sql & " destinos.nomdesti,variedades.nomvarie, forfaits.nomconfe, dtocom1, dtocom2, facturas_variedad.codigiva  "
'            Sql = Sql & " FROM facturas_variedad, albaran, albaran_variedad, variedades, forfaits, destinos " 'lineas de variedades de la factura
'            Sql = Sql & " WHERE facturas_variedad.numalbar = albaran.numalbar "
'            Sql = Sql & " and albaran.numalbar = albaran_variedad.numalbar "
'            Sql = Sql & " and facturas_variedad.numlinealbar = albaran_variedad.numlinea "
'            Sql = Sql & " and albaran_variedad.codvarie = variedades.codvarie "
'            Sql = Sql & " and albaran_variedad.codforfait = forfaits.codforfait "
'            Sql = Sql & " and albaran.codclien = destinos.codclien "
'            Sql = Sql & " and albaran.coddesti = destinos.coddesti "
'
        Case 1  'articulos
'select codtipom,numfactu,fecfactu,numlinea,codalmac,codartic,sartic.nomartic,cantidad,precioar,dtolinea,importel,ampliaci,codigiva
            If vParamAplic.Cooperativa = 3 Then
                Sql = "SELECT advpartes_lineas.numparte,advpartes_lineas.numlinea,advpartes_lineas.codalmac,advpartes_lineas.codartic,advartic.nomartic,dosishab,cantidad,"
            Else
                Sql = "SELECT advpartes_lineas.numparte,advpartes_lineas.numlinea,advpartes_lineas.codalmac,advpartes_lineas.codartic,advartic.nomartic,dosishab,cantidad,"
            End If
            Sql = Sql & "advpartes_lineas.preciove,importel,ampliaci,advpartes_lineas.codigiva"
            Sql = Sql & " FROM advpartes_lineas, advartic "
            Sql = Sql & " WHERE advpartes_lineas.codartic = advartic.codartic "
            
        Case 2  ' trabajadores de la cuadrilla
        'select numparte, numlinea, codtraba, nomtraba
            Sql = "SELECT advpartes_trabajador.numparte,advpartes_trabajador.numlinea,advpartes_trabajador.codtraba,straba.nomtraba,"
            Sql = Sql & " advpartes_trabajador.horas, advpartes_trabajador.precio, advpartes_trabajador.importel "
            Sql = Sql & " FROM advpartes_trabajador, straba "
            Sql = Sql & " WHERE advpartes_trabajador.codtraba = straba.codtraba "
        
    End Select
    
    If enlaza Then
        Sql = Sql & " and " & ObtenerWhereCP(False)
    Else
        Sql = Sql & " and numparte = -1"
    End If
    Sql = Sql & " ORDER BY numparte,numlinea"
    MontaSQLCarga = Sql
End Function


Private Sub PonerModoOpcionesMenu(Modo As Byte)
'Activas unas Opciones de Menu y Toolbar según el modo en que estemos
Dim B As Boolean, bAux As Boolean
Dim I As Integer

    B = ((Modo = 2) Or (Modo = 0)) And (Albaran = "") 'Or (Modo = 5 And ModificaLineas = 0)
    'Buscar
    Toolbar1.Buttons(5).Enabled = B
    Me.mnBuscar.Enabled = B
    'Vore Tots
    Toolbar1.Buttons(6).Enabled = B
    Me.mnVerTodos.Enabled = B
    'Añadir
    Toolbar1.Buttons(1).Enabled = B
    Me.mnModificar.Enabled = B
    
    B = (Modo = 2 And Data1.Recordset.RecordCount > 0) And (Albaran = "")
    'Modificar
    Toolbar1.Buttons(2).Enabled = B
    Me.mnModificar.Enabled = B
    'eliminar
    Toolbar1.Buttons(3).Enabled = B
    Me.mnEliminar.Enabled = B
    'Confirmacion
    Toolbar5.Buttons(1).Enabled = B
    Me.mnConfirmacion.Enabled = B
    'Cuadrilla
    Toolbar5.Buttons(2).Enabled = B
    Me.mnCuadrilla.Enabled = B
    'Asignacion de precios
    Toolbar5.Buttons(3).Enabled = (Albaran = "" And vParamAplic.Cooperativa = 3)
    Me.mnPrecios.Enabled = (Albaran = "" And vParamAplic.Cooperativa = 3)
    
    '[Monica]19/07/2013: nueva opcion de insercion de gastos solo para mogente
    'Insercion de Gastos
    Toolbar5.Buttons(4).Enabled = (Albaran = "" And vParamAplic.Cooperativa = 3)
    Me.mnInsercionGastos.Enabled = (Albaran = "" And vParamAplic.Cooperativa = 3)
    
    'Imprimir
    Toolbar1.Buttons(8).Enabled = B Or (Albaran <> "")
    Me.mnImprimir.Enabled = B Or (Albaran <> "")

    ' *** si n'hi han llínies que tenen grids (en o sense tab) ***
'++monica: si insertamos lo he quitado
'    b = (Modo = 3 Or Modo = 4 Or Modo = 2) And Not DeConsulta
    B = (Modo = 2) And (Albaran = "")
    For I = 1 To 1
        ToolAux(I).Buttons(1).Enabled = B
        
        If B Then
            Select Case I
              Case 0
                bAux = (B And Me.Adoaux(0).Recordset.RecordCount > 0)
              Case 1
                bAux = (B And Me.Adoaux(1).Recordset.RecordCount > 0)
            End Select
        End If
        ToolAux(I).Buttons(2).Enabled = bAux
        ToolAux(I).Buttons(3).Enabled = bAux
    Next I

    ' solo tenemos acceso a los trabajadores de la cuadrilla si estamos en modo cuadrilla
    B = (Modo = 2) And (Albaran = "")
    ToolAux(0).Buttons(1).Enabled = B
    
    bAux = (B And Me.Adoaux(0).Recordset.RecordCount > 0)
    ToolAux(0).Buttons(2).Enabled = bAux
    ToolAux(0).Buttons(3).Enabled = bAux
    


End Sub


Private Sub BotonImprimir()
Dim cadFormula As String
Dim CadParam As String
Dim numParam As Byte
Dim cadSelect As String 'select para insertar en tabla temporal
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String

    If Text1(0).Text = "" Then
    
        If vParamAplic.Cooperativa = 3 Then
            MsgBox "Debe seleccionar un Albarán para Imprimir.", vbInformation
        Else
            MsgBox "Debe seleccionar un Parte para Imprimir.", vbInformation
        End If
        Exit Sub
    End If
    
    cadFormula = ""
    CadParam = ""
    cadSelect = ""
    numParam = 0
    
    '===================================================
    '============ PARAMETROS ===========================
    indRPT = 31 'Impresion de partes
    If Not PonerParamRPT(indRPT, CadParam, numParam, nomDocu) Then Exit Sub
      
    'Nombre fichero .rpt a Imprimir
    frmImprimir.NombreRPT = nomDocu
        
    '===================================================
    '================= FORMULA =========================
    'Cadena para seleccion Nº de factura
    '---------------------------------------------------
    If Text1(0).Text <> "" Then
        'Nº Albaran
        devuelve = "{" & NombreTabla & ".numparte}=" & Val(Text1(0).Text)
        If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
        devuelve = "numparte = " & Val(Text1(0).Text)
        If Not AnyadirAFormula(cadSelect, devuelve) Then Exit Sub
    End If
    
    If Not HayRegParaInforme(NombreTabla, cadSelect) Then Exit Sub
     
     With frmImprimir
            .FormulaSeleccion = cadFormula
            .OtrosParametros = CadParam
            .NumeroParametros = numParam
            .SoloImprimir = False
            .EnvioEMail = False
            .Opcion = 0
            '[Monica]18/05/2012
            If vParamAplic.Cooperativa = 3 Then
                .Titulo = "Impresión de Albaranes de Venta"
                .NroCopias = 2
            Else
                .Titulo = "Impresión de Partes"
            End If
            .ConSubInforme = True
            .Show vbModal
    End With
End Sub

'Private Sub TxtAux3_GotFocus(Index As Integer)
'    ConseguirFoco txtAux3(Index), Modo
'End Sub
'
'Private Sub TxtAux3_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'    If Index <> 0 And KeyCode <> 38 Then KEYdown KeyCode
'End Sub
'
'Private Sub TxtAux3_KeyPress(Index As Integer, KeyAscii As Integer)
'    KEYpress KeyAscii
'End Sub
'
'Private Sub TxtAux3_LostFocus(Index As Integer)
'Dim TipoDto As Byte
'Dim ImpDto As String
'Dim Unidades As String
'Dim cantidad As String
'Dim cad As String
'
'    'Quitar espacios en blanco
'    If Not PerderFocoGnralLineas(txtAux3(Index), ModificaLineas) Then Exit Sub
'
'    Select Case Index
'        Case 4 'Albaran
'            If txtAux3(Index) <> "" Then PonerFormatoEntero txtAux3(Index)
'
'            CargarDatosAlbaran txtAux3(4).Text, txtAux3(5).Text
'
'        Case 5 'Linea de albaran
'            If txtAux3(Index) <> "" Then PonerFormatoEntero txtAux3(Index)
'
'            If txtAux3(4).Text <> "" And txtAux3(5).Text <> "" Then
'                If AlbaranFacturado(txtAux3(4).Text, txtAux3(5).Text) Then
'                    cad = "Esta línea de Albarán está facturada. " & vbCrLf & vbCrLf & "    ¿ Desea continuar ? "
'                    If MsgBox(cad, vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
'                        CargarDatosAlbaran txtAux3(4).Text, txtAux3(5).Text
'                    Else
'                        txtAux3(4).Text = ""
'                        txtAux3(5).Text = ""
'                    End If
'                Else
'                    CargarDatosAlbaran txtAux3(4).Text, txtAux3(5).Text
'                End If
'            End If
'
'            If txtAux3(4).Text = "" Or txtAux3(5).Text = "" Then
'                PonerFoco txtAux3(4)
'            Else
'                PonerFoco txtAux3(8)
'            End If
'
'        Case 8 'precio bruto
'            If txtAux3(Index).Text <> "" Then
'                If PonerFormatoDecimal(txtAux3(Index), 7) Then
'
'                    Select Case TipoFacturarForfaits(txtAux3(4).Text, txtAux3(5).Text)
'                        Case 0  'por unidades
'                            txtAux3(10).Text = Round2(CCur(ImporteSinFormato(txtAux3(Index).Text)) * CCur(ImporteSinFormato(txtAux3(15).Text)), 2)
'                            PonerFormatoDecimal txtAux3(10), 3
'                        Case 1  'por kilos
'                            txtAux3(10).Text = Round2(CCur(ImporteSinFormato(txtAux3(Index).Text)) * CCur(ImporteSinFormato(txtAux3(6).Text)), 2)
'                            PonerFormatoDecimal txtAux3(10), 3
'                        Case Else
'
'                    End Select
'
'                    cmdAceptar.SetFocus
'                Else
'                    Exit Sub
'                End If
'            End If
'
'        Case 10 'importe bruto
'            If txtAux3(Index).Text <> "" Then
'                If PonerFormatoDecimal(txtAux3(Index), 3) Then
'
'                    Select Case TipoFacturarForfaits(txtAux3(4).Text, txtAux3(5).Text)
'                        Case 0
'                            Unidades = ComprobarCero(txtAux3(15).Text)
'                            If CCur(Unidades) <> 0 Then
'                                txtAux3(8).Text = Round2(CCur(ImporteSinFormato(txtAux3(Index).Text)) / CCur(Unidades), 4)
'                            Else
'                                txtAux3(8).Text = 0
'                            End If
'                            PonerFormatoDecimal txtAux3(8), 7
'                        Case 1
'                            cantidad = ComprobarCero(txtAux3(6).Text)
'                            If CCur(cantidad) <> 0 Then
'                                txtAux3(8).Text = Round2(CCur(ImporteSinFormato(txtAux3(Index).Text)) / CCur(cantidad), 4)
'                            Else
'                                txtAux3(8).Text = 0
'                            End If
'                            PonerFormatoDecimal txtAux3(8), 7
'                        Case Else
'
'                    End Select
'
'                    cmdAceptar.SetFocus
'               Else
'                    Exit Sub
'               End If
'            End If
'    End Select
'
'If ((Index = 8 And txtAux3(Index).Text <> "") Or (Index = 10 And txtAux3(Index).Text <> "")) Then
'        Dim Campo2 As String
'        Campo2 = "nrodecprec"
'        TipoDto = DevuelveDesdeBDNew(cAgro, "clientes", "tipodtos", "codclien", Text1(3).Text, "N", Campo2)
'        Select Case TipoFacturarForfaits(txtAux3(4).Text, txtAux3(5).Text)
'            Case 0 ' unidades
''                ImpDto = CalcularImporteDto(txtAux3(15).Text, txtAux3(8).Text, txtAux3(0).Text, txtAux3(1).Text, txtAux3(2).Text, CStr(DBLet(Data1.Recordset!impdtoc, "N")), False)
''                txtAux3(11).Text = CalcularImporte(txtAux3(15).Text, txtAux3(8).Text, txtAux3(12).Text, txtAux3(13).Text, TipoDto, ImpDto)
'                Unidades = ComprobarCero(txtAux3(15).Text)
'                ImpDto = CalcularImporteDto(txtAux3(15).Text, CStr(CCur(ImporteSinFormato(txtAux3(10).Text)) / CCur(Unidades)), txtAux3(0).Text, txtAux3(1).Text, txtAux3(2).Text, CStr(DBLet(Data1.Recordset!impdtoc, "N")), False)
'                txtAux3(11).Text = CalcularImporte(txtAux3(15).Text, CStr(CCur(ImporteSinFormato(txtAux3(10).Text)) / CCur(Unidades)), txtAux3(12).Text, txtAux3(13).Text, TipoDto, ImpDto, txtAux3(10).Text)
'                PonerFormatoDecimal txtAux3(11), 1
'
'                'precio neto
'                If ComprobarCero(txtAux3(15).Text) <> "0" Then
'                    txtAux3(9).Text = Round2(CCur(ImporteSinFormato(txtAux3(11).Text)) / CCur(ImporteSinFormato(txtAux3(15).Text)), CCur(Campo2))
'                End If
'                PonerFormatoDecimal txtAux3(9), 7
'
'            Case 1 ' kilos
''                ImpDto = CalcularImporteDto(txtAux3(6).Text, txtAux3(8).Text, txtAux3(0).Text, txtAux3(1).Text, txtAux3(2).Text, CStr(DBLet(Data1.Recordset!impdtoc, "N")), False)
''                txtAux3(11).Text = CalcularImporte(txtAux3(6).Text, txtAux3(8).Text, txtAux3(12).Text, txtAux3(13).Text, TipoDto, ImpDto)
'                cantidad = ComprobarCero(txtAux3(6).Text)
'                ImpDto = CalcularImporteDto(txtAux3(6).Text, CStr(CCur(ImporteSinFormato(txtAux3(10).Text)) / CCur(cantidad)), txtAux3(0).Text, txtAux3(1).Text, txtAux3(2).Text, CStr(DBLet(Data1.Recordset!impdtoc, "N")), False)
'                txtAux3(11).Text = CalcularImporte(txtAux3(6).Text, CStr(CCur(ImporteSinFormato(txtAux3(10).Text)) / CCur(cantidad)), txtAux3(12).Text, txtAux3(13).Text, TipoDto, ImpDto, txtAux3(10).Text)
'                PonerFormatoDecimal txtAux3(11), 1
'
'                'precio neto
'                If ComprobarCero(txtAux3(6).Text) <> "0" Then
'                    txtAux3(9).Text = Round2(CCur(ImporteSinFormato(txtAux3(11).Text)) / CCur(ImporteSinFormato(txtAux3(6).Text)), CCur(Campo2))
'                End If
'                PonerFormatoDecimal txtAux3(9), 7
'
'            Case Else
'
'        End Select
'
'    End If
'
'End Sub

Private Function ModificaCabeceraCuadrilla() As Boolean
Dim Sql As String
Dim SqlValues As String
Dim Rs As ADODB.Recordset
Dim I As Long
Dim B As Boolean
Dim MenError As String

Dim cantidad As Currency
Dim Precio As Currency
Dim Importe As Currency


    On Error GoTo EModificarCab
    
    conn.BeginTrans
    
    MenError = "Repasando datos de Cuadrilla"
    
    If Text1(32).Text = "" Then Text1(32).Text = "0"
    If CLng(Text1(32).Text) <> CuadrillaAnt Then
        Sql = "select count(*) from advpartes_trabajador where numparte = " & DBSet(Text1(0).Text, "N")
        If TotalRegistros(Sql) <> 0 Then
            If MsgBox("Se van a eliminar los trabajadores actuales del parte y añadir los de la nueva cuadrilla." & "¿ Continuar ?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
                conn.Execute "delete from advpartes_trabajador where numparte = " & DBLet(Data1.Recordset!Numparte, "N")
            Else
                ModificaCabeceraCuadrilla = True
                conn.CommitTrans
                Exit Function
            End If
        End If
        
        Sql = "select codtraba from rcuadrilla_trabajador where codcuadrilla = " & DBSet(Text1(32).Text, "N")
        Set Rs = New ADODB.Recordset
        Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            
        Sql = "insert into advpartes_trabajador (numparte, numlinea, codtraba, horas, precio, importel) values "
        SqlValues = ""
        I = 0
        While Not Rs.EOF
            I = I + 1
            
            Precio = DevuelveValor("select impsalar from salarios inner join straba on salarios.codcateg = straba.codcateg where straba.codtraba = " & DBSet(Rs!CodTraba, "N"))
            
            cantidad = 0
            If Text1(34).Text <> "" Then cantidad = CCur(Text1(34).Text)
            
            Importe = Round2(Precio * cantidad, 2)
             
            SqlValues = SqlValues & "(" & Data1.Recordset!Numparte & "," & I & "," & DBSet(Rs!CodTraba, "N") & ","
            SqlValues = SqlValues & DBSet(cantidad, "N") & "," & DBSet(Precio, "N") & "," & DBSet(Importe, "N") & "),"
            
            Rs.MoveNext
        Wend
        If SqlValues <> "" Then
            conn.Execute Sql & Mid(SqlValues, 1, Len(SqlValues) - 1)
        End If
        Set Rs = Nothing
            
    End If
    B = ModificaDesdeFormulario2(Me, 2, "Frame3")
    
    Text1(35).Tag = ""
    

EModificarCab:
    If Err.Number <> 0 Then
        MenError = "Modificando Datos de Trabajadores de Partes." & vbCrLf & "----------------------------" & vbCrLf & MenError
        MuestraError Err.Number, MenError, Err.Description
        B = False
    End If
    If B Then
        ModificaCabeceraCuadrilla = True
        ModoCuadrilla = False
        conn.CommitTrans
    Else
        ModificaCabeceraCuadrilla = False
        conn.RollbackTrans
    End If

End Function

Private Function ModificaCabecera() As Boolean
Dim B As Boolean
Dim MenError As String

    On Error GoTo EModificarCab

            
    conn.BeginTrans
    
    B = ModificaDesdeFormulario2(Me, 2, "Frame2")
    
    MenError = "Recalculando Importes Netos de lineas"
    
    If Confirmacion Or CampoAnt <> CLng(Text1(4).Text) Then
        If B Then B = RecalcularImportes(Text1(0).Text, Text1(8).Text, MenError)
    Else
        If LitrosAnt <> CLng(Text1(7).Text) And ComprobarCero(Text1(8).Text) = 0 Then
            If B Then B = RecalcularImportes(Text1(0).Text, Text1(7).Text, MenError)
        End If
    End If
    

EModificarCab:
    If Err.Number <> 0 Then
        MenError = "Modificando Articulos ADV de Partes." & vbCrLf & "----------------------------" & vbCrLf & MenError
        MuestraError Err.Number, MenError, Err.Description
        B = False
    End If
    If B Then
        ModificaCabecera = True
        conn.CommitTrans
    Else
        ModificaCabecera = False
        conn.RollbackTrans
    End If
End Function


Private Sub InsertarCabecera()
Dim vTipoMov As CTiposMov 'Clase Tipo Movimiento
Dim Sql As String

    On Error GoTo EInsertarCab
    
    Set vTipoMov = New CTiposMov
    If vTipoMov.Leer(CodTipoMov) Then
        Text1(0).Text = vTipoMov.ConseguirContador(CodTipoMov)
        Sql = CadenaInsertarDesdeForm(Me)
        If Sql <> "" Then
            If InsertarOferta(Sql, vTipoMov) Then
                CadenaConsulta = "Select * from " & NombreTabla & ObtenerWhereCP(True) & Ordenacion
                PonerCadenaBusqueda
                PonerModo 2
                'Ponerse en Modo Insertar Lineas
'                BotonMtoLineas 0, "Variedades"
'                BotonAnyadirLinea 0
            End If
        End If
        Text1(0).Text = Format(Text1(0).Text, "0000000")
    End If
    Set vTipoMov = Nothing
    
EInsertarCab:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Function InsertarOferta(vSQL As String, vTipoMov As CTiposMov) As Boolean
Dim MenError As String
Dim bol As Boolean, Existe As Boolean
Dim cambiaSQL As Boolean
Dim devuelve As String
Dim Sql2 As String
Dim vCstock As CStockADV
Dim Rs As ADODB.Recordset
Dim Sql As String
Dim Precio As Currency
Dim Importe As Currency
Dim Tipo As Byte
Dim vHayReg As Byte
    On Error GoTo EInsertarOferta
    
    bol = True
    
    cambiaSQL = False
    'Comprobar si mientras tanto se incremento el contador de Albaranes
    'para ello vemos si existe una factura con ese contador y si existe la incrementamos
    Do
        devuelve = DevuelveDesdeBDNew(cAgro, NombreTabla, "numparte", "numparte", Text1(0), "N")
        If devuelve <> "" Then
            'Ya existe el contador incrementarlo
            Existe = True
            vTipoMov.IncrementarContador (CodTipoMov)
            Text1(0).Text = vTipoMov.ConseguirContador(CodTipoMov)
            cambiaSQL = True
        Else
            Existe = False
        End If
    Loop Until Not Existe
    If cambiaSQL Then vSQL = CadenaInsertarDesdeForm(Me)
    
    'Aqui empieza transaccion
    conn.BeginTrans
    
    MenError = "Error al insertar en la tabla Cabecera de Partes (" & NombreTabla & ")."
    conn.Execute vSQL, , adCmdText
    
    MenError = "Error al insertar en la tabla Lineas de Partes (" & NomTablaLineas & ")."
    
    Sql2 = "Select " & Text1(0).Text & ", numlinea, " & vParamAplic.AlmacenADV & " as codalmac ,"
    Sql2 = Sql2 & " advtrata_lineas.codartic, advtrata_lineas.dosishab, advtrata_lineas.cantidad, "
    Sql2 = Sql2 & " advartic.preciove, round(advtrata_lineas.cantidad * advartic.preciove) as importe, "
    Sql2 = Sql2 & ValorNulo & "," 'ampliacion
    Sql2 = Sql2 & " advartic.codigiva "
    Sql2 = Sql2 & " from advtrata_lineas, advartic "
    Sql2 = Sql2 & " where advtrata_lineas.codartic = advartic.codartic "
    Sql2 = Sql2 & " and advtrata_lineas.codtrata = " & DBSet(Text1(3).Text, "T")

    Set Rs = New ADODB.Recordset
    Rs.Open Sql2, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    
    ModificaLineas = 1 ' insertar lineas de advpartes
    
    vHayReg = 0
    
    
    While Not Rs.EOF And bol
        Set vCstock = New CStockADV
        
        txtAux(5).Text = DBLet(Rs!codArtic, "T")
        txtAux(4).Text = vParamAplic.AlmacenADV
        txtAux(6).Text = DBLet(Rs!cantidad, "N")
        txtAux(9).Text = DBLet(Rs!Importe, "N")
        
        If Not InicializarCStock(vCstock, "S", DBLet(Rs!numlinea, "N")) Then Exit Function
        
        Precio = DBLet(Rs!preciove, "N")
        Importe = DBLet(Rs!Importe, "N")
        
        If DatosOkLineaEnv(vCstock) Then 'Lineas de factura
            'Inserta en tabla "facturas_envases"
            Sql = "INSERT INTO advpartes_lineas "
            Sql = Sql & "(numparte,numlinea,codalmac,codartic,dosishab,cantidad,preciove,importel,ampliaci,codigiva) "
            Sql = Sql & "VALUES (" & DBSet(Text1(0).Text, "N") & ", " & DBSet(Rs!numlinea, "N") & ", " & DBSet(vParamAplic.AlmacenADV, "N") & ","
            Sql = Sql & DBSet(Rs!codArtic, "T") & ", "
            Sql = Sql & DBSet(Rs!dosishab, "N") & ", "
            Sql = Sql & DBSet(DBLet(Rs!cantidad, "N"), "N") & ", "
            '[Monica]14/02/2012: cambiamos el precio
'            Sql = Sql & DBSet(DBLet(Rs!preciove, "N"), "N") & ", "
'            Sql = Sql & DBSet(DBLet(Rs!Importe, "N"), "N") & ","
            Sql = Sql & DBSet(DBLet(Precio, "N"), "N") & ", "
            Sql = Sql & DBSet(DBLet(Importe, "N"), "N") & ","
            Sql = Sql & ValorNulo & ","
            Sql = Sql & DBSet(DBLet(Rs!Codigiva, "N"), "N") & ")"
         Else
            Exit Function
         End If
        
        If Sql <> "" Then
            'insertar la linea
            conn.Execute Sql
            
            'si hay control de stock para el articulo actualizar en salmac e insertar en smoval
            'en actualizar stock comprobamos si el articulo tiene control de stock
            bol = vCstock.ActualizarStock()
            
            vHayReg = 1
        End If
        
        Set vCstock = Nothing
        
        Rs.MoveNext
    Wend
    
    Set Rs = Nothing
    
    
    '[Monica]14/02/2012: de momento nada
'    If vHayReg = 1 And Trim(vParamAplic.CodArticADV) <> "" Then
'        Dim Max As Long
'        '[Monica]14/02/2012: en caso de que haya algun articulo de tipo producto en el parte introducimos la linea de mano de obra
'        Sql = "select count(*) from advpartes_lineas INNER JOIN advartic ON advpartes_lineas.codartic = advartic.codartic "
'        Sql = Sql & " where advartic.tipoprod = 0 and advpartes_lineas.numparte = " & DBSet(Text1(0).Text, "N")
'
'        If TotalRegistros(Sql) > 0 Then
'            Max = DevuelveValor("select if(max(numlinea) is null,0,max(numlinea)) from advpartes_lineas where numparte = " & DBSet(Text1(0).Text, "N"))
'        End If
'    End If
    
    ModificaLineas = 0
     
    MenError = "Error al actualizar el contador de la Factura."
    vTipoMov.IncrementarContador (CodTipoMov)
    
    '[Monica]30/09/2009
    ' tenemos que hacer el recalculo con los litros previstos cuando se inserta un parte
    If bol Then
        bol = RecalcularImportes(Text1(0).Text, Text1(7).Text, "Recalculando importes")
    End If
    
    
EInsertarOferta:
        If Err.Number <> 0 Then
            MenError = "Insertando Parte." & vbCrLf & "----------------------------" & vbCrLf & MenError
            MuestraError Err.Number, MenError, Err.Description
            bol = False
        End If
        If bol Then
            conn.CommitTrans
            InsertarOferta = True
        Else
            conn.RollbackTrans
            InsertarOferta = False
        End If
End Function


'Private Sub CargaForaGrid()
'    If DataGrid2.Columns.Count <= 2 Then Exit Sub
'    ' *** posar als camps de fora del grid el valor de la columna corresponent ***
'    Text3(0) = DataGrid2.Columns(12).Text    'Fecha
'    Text3(1) = DataGrid2.Columns(13).Text    'Matricula
'    Text3(2) = DataGrid2.Columns(14).Text    'Destino
'    Text3(3) = DataGrid2.Columns(15).Text   'Variedad
'    Text3(4) = DataGrid2.Columns(16).Text   'Confeccion
'    ' *** Si fora del grid n'hi han camps de descripció, posar-los valor ***
'    ' **********************************************************************
'End Sub

Private Sub InsertarLinea()
'Inserta registre en les taules de Llínies
Dim nomframe As String
Dim B As Boolean
Dim Mens As String

    On Error Resume Next

    ' *** posa els noms del frames, tant si son de grid com si no ***
'        Case 0: nomFrame = "FrameAux0" 'variedades
    Select Case numTab
        Case 1
            nomframe = "FrameAux1" 'envases
        Case 0
            nomframe = "FrameAux0"
    ' ***************************************************************
    End Select
    
    
    If DatosOkLlin(nomframe) Then
        TerminaBloquear
        
        Select Case numTab
            Case 1
                If InsertarLineaEnv(txtAux(3).Text) Then
                    CalcularDatosAlbaran
                    B = BloqueaRegistro("advpartes", "numparte = " & Data1.Recordset!Numparte)
                    CargaGrid DataGrid3, Adoaux(1), True
            
                    If B Then BotonAnyadirLinea 1
                End If
            Case 0
                If InsertarLineaTrab(txtAux1(1).Text) Then
                    B = BloqueaRegistro("advpartes", "numparte = " & Data1.Recordset!Numparte)
                    CargaGrid DataGrid1, Adoaux(0), True
            
                    If B Then BotonAnyadirLinea 0
                End If
        End Select
    End If
End Sub

Private Sub BotonAnyadirLinea(Index As Integer)
Dim NumF As String
Dim vWhere As String, vtabla As String
Dim anc As Single
Dim I As Integer
    
    ModificaLineas = 1 'Posem Modo Afegir Llínia
    
    If (Modo = 3) Or (Modo = 4) Then 'Insertar o Modificar Capçalera
        cmdAceptar_Click
        If ModificaLineas = 0 Then Exit Sub
    End If
       
    PonerModo 5, Index
    
    numTab = Index
    
    ' *** bloquejar la clau primaria de la capçalera ***
    BloquearTxt Text1(0), True
    
    Select Case Index
        Case 0 ' trabajadores
            ' *** posar el nom del les distintes taules de llínies ***
            vtabla = "advpartes_trabajador"
            ' ********************************************************
            
            vWhere = ObtenerWhereCab(False)
            
            NumF = SugerirCodigoSiguienteStr(vtabla, "numlinea", vWhere)
            ' ***************************************************************
        
            AnyadirLinea DataGrid1, Adoaux(0)
        
            anc = DataGrid1.Top
            If DataGrid1.Row < 0 Then
                anc = anc + 215 '210
            Else
                anc = anc + DataGrid1.RowTop(DataGrid1.Row) + 5
            End If
          
            LLamaLineas ModificaLineas, anc, "DataGrid1"
        
            LimpiarCamposLin "FrameAux0"
            txtAux1(0).Text = Text1(0).Text 'numparte
        '            txtAux(2).Text = Text1(1).Text 'fecfactu
            txtAux1(1).Text = NumF
            PonerFoco txtAux1(2)
            txtAux1(3).Text = ""
            BloquearTxt txtAux1(2), False
            
            BloquearBtn Me.btnBuscar(1), False
        
        Case 1 ' lineas de articulos
    
            ' **************************************************
        
            ' *** posar el nom del les distintes taules de llínies ***
            vtabla = "advpartes_lineas"
            ' ********************************************************
            
            vWhere = ObtenerWhereCab(False)
            
            NumF = SugerirCodigoSiguienteStr(vtabla, "numlinea", vWhere)
            ' ***************************************************************
        
            AnyadirLinea DataGrid3, Adoaux(1)
        
            anc = DataGrid3.Top
            If DataGrid3.Row < 0 Then
                anc = anc + 215 '210
            Else
                anc = anc + DataGrid3.RowTop(DataGrid3.Row) + 5
            End If
          
            LLamaLineas ModificaLineas, anc, "DataGrid3"
        
            LimpiarCamposLin "FrameAux1"
        '            txtAux(0).Text = Text1(6).Text 'codtipom
            txtAux(1).Text = Text1(0).Text 'numparte
        '            txtAux(2).Text = Text1(1).Text 'fecfactu
            txtAux(3).Text = NumF
            txtAux(4).Text = vParamAplic.AlmacenADV
            PonerFoco txtAux(5)
            For I = 0 To 0
                Text2(I).Text = ""
            Next I
            txtAux(10).Enabled = False
            txtAux(10).visible = False
            BloquearTxt txtAux(9), True
            BloquearTxt Text2(16), False
            BloquearBtn Me.btnBuscar(0), False
' ******************************************
    End Select
End Sub

Private Sub LimpiarCamposLin(FrameAux As String)
    On Error Resume Next
    
    LimpiarLin Me, FrameAux  'Mètode general: Neteja els controls TextBox
    lblIndicador.Caption = ""

    If Err.Number <> 0 Then Err.Clear
End Sub

Private Function ModificarLinea() As Boolean
'Modifica registre en les taules de Llínies
Dim nomframe As String
Dim V As Integer
Dim cad As String
Dim Sql As String
Dim vCstock As CStockADV
Dim B As Boolean
Dim Mens As String
    
    On Error GoTo eModificarLinea

    ModificarLinea = False
    Sql = ""

    ' *** posa els noms del frames, tant si son de grid com si no ***
    nomframe = "FrameAux1" 'articulos del parte
    ' **************************************************************

    If DatosOkLlin(nomframe) Then
        TerminaBloquear
        
        Set vCstock = New CStockADV
        If Not InicializarCStock(vCstock, "S", , Me.Adoaux(1).Recordset) Then Exit Function
        
        If DatosOkLineaEnv(vCstock) Then
            '#### LAURA 15/11/2006
            conn.BeginTrans
            
    '        Set vCStock = New CStock
            'iniciamos la clase con los valores anteriores para deshacer lo q insertamos antes

            B = InicializarCStock(vCstock, "E", , Me.Adoaux(1).Recordset)

            If B Then
                B = vCstock.DevolverStock 'eliminamos de smoval y devolvemos stock valores anteriores
                'ahora leemos los valores nuevos
                If B Then B = InicializarCStock(vCstock, "S", , Me.Adoaux(1).Recordset)
                'insertamos en smoval y actualizamos stock a los valores nuevos
                
                vCstock.cantidad = CSng(ComprobarCero(txtAux(6).Text))
                If B Then B = vCstock.ActualizarStock()
        
                'actualizar la linea de Albaran
                If B Then
                    Sql = "UPDATE advpartes_lineas Set codalmac = " & txtAux(4).Text & ", codartic=" & DBSet(txtAux(5).Text, "T") & ", "
                    Sql = Sql & "ampliaci=" & DBSet(Text2(16).Text, "T") & ", "
                    Sql = Sql & "cantidad= " & DBSet(txtAux(6).Text, "N") & ", "
                    Sql = Sql & "preciove= " & DBSet(txtAux(7).Text, "N") & ", " 'precio
                    Sql = Sql & "dosishab= " & DBSet(txtAux(8).Text, "N", "S") & ", " ' dosis habitual
                    Sql = Sql & "importel= " & DBSet(txtAux(9).Text, "N") & ", " 'Importe
                    Sql = Sql & "codigiva= " & DBSet(txtAux(10).Text, "N") & " " 'codigo de iva
                    Sql = Sql & Replace(ObtenerWhereCP(True), NombreTabla, "advpartes_lineas") & " AND numlinea=" & Adoaux(1).Recordset!numlinea
                    conn.Execute Sql
                End If
            End If
        End If
        Set vCstock = Nothing
        
        CalcularDatosAlbaran
        
        ModificaLineas = 0
        
        V = Adoaux(1).Recordset.Fields(1) 'el 2 es el nº de llinia
        CargaGrid DataGrid3, Adoaux(1), True
        CargaGrid DataGrid1, Adoaux(0), True

        DataGrid3.SetFocus
        Adoaux(1).Recordset.Find (Adoaux(1).Recordset.Fields(1).Name & " =" & V)

        LLamaLineas ModificaLineas, 0, "DataGrid3"
    End If
        
        
eModificarLinea:
    If Err.Number <> 0 Or Not B Then
        MuestraError Err.Number, "Modificar Lineas Factura" & vbCrLf & Err.Description & vbCrLf & Mens
        B = False
    End If
    
    If B Then
        conn.CommitTrans
        ModificarLinea = True
    Else
        conn.RollbackTrans
        ModificarLinea = False
    End If
    CargaGrid DataGrid3, Adoaux(1), True
    CargaGrid DataGrid1, Adoaux(0), True
    
    Adoaux(1).Recordset.Find (Adoaux(1).Recordset.Fields(1).Name & " =" & V)
End Function
        
        
Private Function ModificarLineaCuadrilla() As Boolean
'Modifica registre en les taules de Llínies
Dim nomframe As String
Dim V As Integer
Dim cad As String
Dim Sql As String
Dim vCstock As CStockADV
Dim B As Boolean
Dim Mens As String
    
    On Error GoTo eModificarLinea

    ModificarLineaCuadrilla = False
    Sql = ""

    ' *** posa els noms del frames, tant si son de grid com si no ***
    nomframe = "FrameAux0" 'trabajadores de la cuadrilla
    ' **************************************************************

    If DatosOkLlin(nomframe) Then
        TerminaBloquear
        
        ModificaDesdeFormulario2 Me, 2, "FrameAux0"
        
        ModificaLineas = 0
        
        V = Adoaux(0).Recordset.Fields(1) 'el 2 es el nº de llinia
        CargaGrid DataGrid1, Adoaux(0), True

        DataGrid1.SetFocus
        Adoaux(0).Recordset.Find (Adoaux(0).Recordset.Fields(1).Name & " =" & V)

        LLamaLineas ModificaLineas, 0, "DataGrid1"
    
        ModificarLineaCuadrilla = True
    
    End If
    
    Exit Function
        
eModificarLinea:
    If Err.Number <> 0 Or Not B Then
        MuestraError Err.Number, "Modificar Lineas Factura" & vbCrLf & Err.Description & vbCrLf & Mens
    End If
End Function
        
        

Private Function DatosOkLlin(nomframe As String) As Boolean
Dim Rs As ADODB.Recordset
Dim Sql As String
Dim B As Boolean
Dim Cant As Integer
Dim Mens As String
Dim vFact As Byte, vDocum As Byte
Dim Cliente As String

    DatosOkLlin = True
    
    On Error GoTo EDatosOKLlin

    Mens = ""
    DatosOkLlin = False

    B = CompForm2(Me, 2, nomframe) 'Comprovar formato datos ok
    If Not B Then Exit Function
    
    'en variedades comprobamos que el albaran introducido corresponde al cliente
    Select Case nomframe
        Case "FrameAux1"
            '++
            '[Monica]15/02/2011: Problema con el Alt+A
            If vParamAplic.Cooperativa <> 3 Then
                txtAux_LostFocus (6)
                txtAux_LostFocus (7)
                txtAux_LostFocus (8)
                txtAux_LostFocus (9)
            End If
            '++
    
        Case "FrameAux0"
        
    End Select
    
    DatosOkLlin = B

EDatosOKLlin:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function

Private Function ObtenerWhereCab(conW As Boolean) As String
Dim vWhere As String
    
    vWhere = ""
    If conW Then vWhere = " WHERE "
    ' *** canviar-ho per la clau primaria de la capçalera ***
    vWhere = vWhere & " numparte= " & Val(Text1(0).Text)
    ' *******************************************************
    
    ObtenerWhereCab = vWhere
End Function

'' ***** si n'hi han varios nivells de tabs *****
'Private Sub SituarTab(numTab As Integer)
'    On Error Resume Next
'
'    SSTab1.Tab = numTab
'
'    If Err.Number <> 0 Then Err.Clear
'End Sub

' **********************************************
Private Sub PonerDatosSocios(Codsocio As String, Optional nifSocio As String)
Dim vSocio As cSocio
Dim Observaciones As String
    
    On Error GoTo EPonerDatos
    
    If Codsocio = "" Then
        LimpiarDatosSocio
        Exit Sub
    End If

    Set vSocio = New cSocio
    
    'si se ha modificado el cliente volver a cargar los datos
    If vSocio.Existe(Codsocio) Then
        If vSocio.LeerDatos(Codsocio) Then
            Text1(6).Text = vSocio.Codigo
            FormateaCampo Text1(6)
            If (Modo = 3) Or (Modo = 4) Then
                Text2(6).Text = vSocio.Nombre  'Nom socio
            End If
            Observaciones = DBLet(vSocio.Observaciones, "T")
            If Trim(Observaciones) <> "" Then
                MsgBox Observaciones, vbInformation, "Observaciones del socio"
            End If
        End If
    Else
        LimpiarDatosSocio
    End If
    Set vSocio = Nothing

EPonerDatos:
    If Err.Number <> 0 Then MuestraError Err.Number, "Poner Datos Cliente", Err.Description
End Sub


Private Sub LimpiarDatosSocio()
Dim I As Byte

    Text1(2).Text = ""
    Text1(4).Text = ""
    Text1(7).Text = ""
    Text1(8).Text = "0"
    Text1(6).Text = ""
    

    Text2(3).Text = ""
    Text2(4).Text = ""
    Text2(6).Text = ""
End Sub
    

Private Function InsertarLineaEnv(numlinea As String) As Boolean
'Inserta un registro en la tabla de lineas de Albaranes: slialb
Dim Sql As String
Dim vWhere As String
Dim B As Boolean
Dim vCstock As CStockADV
Dim DentroTRANS As Boolean

    InsertarLineaEnv = False
    Sql = ""
    DentroTRANS = False
    
    'Conseguir el siguiente numero de linea
    vWhere = Replace(ObtenerWhereCP(False), NombreTabla, NomTablaLineas)
'    numlinea = SugerirCodigoSiguienteStr(NomTablaLineas, "numlinea", vWhere)
    
    Set vCstock = New CStockADV
    
    If Not InicializarCStock(vCstock, "S", CInt(numlinea)) Then Exit Function
    
    If DatosOkLineaEnv(vCstock) Then 'Lineas de factura
        'Inserta en tabla "facturas_envases"
        Sql = "INSERT INTO advpartes_lineas "
        Sql = Sql & "(numparte,numlinea,codalmac,codartic,dosishab,cantidad,preciove,importel,ampliaci,codigiva) "
        Sql = Sql & "VALUES (" & DBSet(txtAux(1).Text, "N") & ", " & numlinea & ", " & DBSet(txtAux(4).Text, "N") & ","
        Sql = Sql & DBSet(txtAux(5).Text, "T") & ", "
        Sql = Sql & DBSet(txtAux(8).Text, "N") & ", "
        Sql = Sql & DBSet(txtAux(6).Text, "N") & ", "
        Sql = Sql & DBSet(txtAux(7).Text, "N") & ", "
        Sql = Sql & DBSet(txtAux(9).Text, "N") & ","
        Sql = Sql & DBSet(Text2(16).Text, "T") & ","
        Sql = Sql & DBSet(txtAux(10).Text, "N") & ")"
     Else
        Exit Function
     End If
    
    If Sql <> "" Then
        On Error GoTo EInsertarLineaEnv
        conn.BeginTrans
        DentroTRANS = True
        
        'insertar la linea
        conn.Execute Sql
        
        'si hay control de stock para el articulo actualizar en salmac e insertar en smoval
        'en actualizar stock comprobamos si el articulo tiene control de stock
        B = vCstock.ActualizarStock()
        
    
    End If
    Set vCstock = Nothing
    
    If B Then
        conn.CommitTrans
        InsertarLineaEnv = True
    Else
        conn.RollbackTrans
         InsertarLineaEnv = False
    End If
    Exit Function
    
EInsertarLineaEnv:
    If Err.Number <> 0 Then
        InsertarLineaEnv = False
        If DentroTRANS Then conn.RollbackTrans
        MuestraError Err.Number, "Insertar Lineas Partes" & vbCrLf & Err.Description
'        b = False
    End If
'    If b Then
'        Conn.CommitTrans
'        InsertarLinea = True
'    Else
'        Conn.RollbackTrans
'         InsertarLinea = False
'    End If
End Function


Private Function InsertarLineaTrab(numlinea As String) As Boolean
'Inserta un registro en la tabla de lineas de Albaranes: slialb
Dim Sql As String
Dim vWhere As String
Dim B As Boolean
Dim DentroTRANS As Boolean

    On Error GoTo EInsertarLineaTrab

    InsertarLineaTrab = False
    Sql = ""
    
    'Conseguir el siguiente numero de linea
    vWhere = Replace(ObtenerWhereCP(False), NombreTabla, NomTablaLineas)
'    numlinea = SugerirCodigoSiguienteStr(NomTablaLineas, "numlinea", vWhere)
    
    'Inserta en tabla "facturas_envases"
    Sql = "INSERT INTO advpartes_trabajador "
    Sql = Sql & "(numparte,numlinea,codtraba,horas,precio,importel) "
    Sql = Sql & "VALUES (" & DBSet(txtAux1(0).Text, "N") & ", " & numlinea & ", " & DBSet(txtAux1(2).Text, "N") & ","
    Sql = Sql & DBSet(txtAux1(4).Text, "N") & "," & DBSet(txtAux1(5).Text, "N") & "," & DBSet(txtAux1(6).Text, "N") & ")"
    
    'insertar la linea
    conn.Execute Sql
        
    InsertarLineaTrab = True
    Exit Function
    
EInsertarLineaTrab:
    If Err.Number <> 0 Then
        InsertarLineaTrab = False
        MuestraError Err.Number, "Insertar Trabajador Partes" & vbCrLf & Err.Description
    End If
End Function




Private Function DatosOkLineaEnv(ByRef vCstock As CStockADV) As Boolean
Dim B As Boolean
Dim I As Byte
    
    On Error GoTo EDatosOkLineaEnv

    DatosOkLineaEnv = False
    B = True

    'Comprobar si se mueve mover stock (hay stock, o si no hay pero no control de stock)
    If vCstock.MueveStock Then
        B = vCstock.MoverStock
    End If
    DatosOkLineaEnv = B
    
EDatosOkLineaEnv:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function



Private Function EliminarStock() As Boolean
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim vCstock As CStockADV
Dim B As Boolean

    On Error GoTo eEliminarStock
    
    Sql = "select * from advpartes_lineas where " & ObtenerWhereCP(False)
    Set Rs = New ADODB.Recordset
    
    Rs.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    B = True
    While Not Rs.EOF And B
        Set vCstock = New CStockADV
        
        vCstock.cantidad = DBLet(Rs!cantidad, "N")
        vCstock.codAlmac = DBLet(Rs!codAlmac, "N")
        vCstock.codArtic = DBLet(Rs!codArtic, "T")
        vCstock.Documento = Format(DBLet(Rs!Numparte, "N"), "0000000")
        vCstock.DetaMov = DBLet(CodTipoMov, "T")
        vCstock.Fechamov = DBLet(Text1(1).Text, "F")
        vCstock.Importe = DBLet(Rs!ImporteL, "N")
        vCstock.LineaDocu = DBLet(Rs!numlinea, "N")
        vCstock.tipoMov = "E"
        
        B = vCstock.DevolverStock
        
        Rs.MoveNext
        
        Set vCstock = Nothing
    Wend

    Set Rs = Nothing

eEliminarStock:
    If Err.Number <> 0 Or Not B Then
        EliminarStock = False
    Else
        EliminarStock = True
    End If

End Function


Private Sub CalcularDatosAlbaran()
Dim I As Integer
Dim cadWHERE As String, Sql As String

    'Limpiar en el form los datos calculados de la factura
    'y volvemos a recalcular
    For I = 9 To 31
         Text1(I).Text = ""
    Next I
    
    'Comprobar que hay lineas de facturas_variedad para calcular totales
    cadWHERE = ObtenerWhereCP(False)
    Sql = "Select count(*) from " & NomTablaLineas & " Where " & Replace(cadWHERE, NombreTabla, NomTablaLineas)
    If RegistrosAListar(Sql) = 0 Then
        'Comprobar que hay lineas de facturas_envases para calcular totales
        Sql = "Select count(*) from advpartes_lineas Where " & Replace(cadWHERE, NombreTabla, "advpartes_lineas")
        If RegistrosAListar(Sql) = 0 Then Exit Sub
    End If
    
    
    If CalcularDatosAlbaranVenta(cadWHERE, NombreTabla, NomTablaLineas) Then
'        PosicionarData
'        PonerCampos
    Else
        MuestraError Err.Number, "Calculando Totales", Err.Description
    End If
'    Set vFactu = Nothing
End Sub

'
'##Monica
'
Private Function CalcularDatosAlbaranVenta(cadWHERE As String, NomTabla As String, NomTablaLin As String) As Boolean
'cadWhere: cad para la where de la SQL que selecciona las lineas del albaran o la factura
'nomTabla: nombre de la tabla de albaranes(scaalp) o de AlbaranesXFactura(scafpa)
'           segun llamemos desde recepcion de facturas o desde Hco de Facturas
Dim Rs As ADODB.Recordset
Dim I As Integer

Dim Sql As String
Dim cadAux As String
Dim cadAux1 As String

'Aqui vamos acumulando los totales
Dim TotBruto As Currency
Dim TotNeto As Currency
Dim TotImpIVA As Currency

Dim ImpAux As Currency
Dim impiva As Currency
Dim ImpREC As Currency
Dim ImpBImIVA As Currency 'Importe Base imponible a la que hay q aplicar el IVA

Dim vBruto As Currency
Dim vNeto As Currency

Dim exentoIVA As Boolean
Dim conDesplaz As Boolean
    
Dim BaseImp As Currency
Dim BaseIVA1 As Currency
Dim BaseIVA2 As Currency
Dim BaseIVA3 As Currency
    
Dim BrutoFac As Currency
    
Dim ImpIVA1 As Currency
Dim ImpIVA2 As Currency
Dim ImpIVA3 As Currency
    
Dim PorceIVA1 As Currency
Dim PorceIVA2 As Currency
Dim PorceIVA3 As Currency
    
Dim ImpREC1 As Currency
Dim ImpREC2 As Currency
Dim ImpREC3 As Currency
    
Dim PorceREC1 As Currency
Dim PorceREC2 As Currency
Dim PorceREC3 As Currency
    
Dim TipoIVA1 As Currency
Dim TipoIVA2 As Currency
Dim TipoIVA3 As Currency
    
Dim TotalFac As Currency

Dim IvaAnt As Integer
Dim cadWhere1 As String
    
Dim Nulo2 As String
Dim Nulo3 As String

Dim vSeccion As CSeccion

Dim EsFactADVInterna As Byte

    CalcularDatosAlbaranVenta = False
    On Error GoTo ECalcular

    BaseImp = 0
    BaseIVA1 = 0
    BaseIVA2 = 0
    BaseIVA3 = 0
    
    BrutoFac = 0
    
    ImpIVA1 = 0
    ImpIVA2 = 0
    ImpIVA3 = 0
    
    PorceIVA1 = 0
    PorceIVA2 = 0
    PorceIVA3 = 0
    
    ImpREC1 = 0
    ImpREC2 = 0
    ImpREC3 = 0
    
    PorceREC1 = 0
    PorceREC2 = 0
    PorceREC3 = 0
    
    TipoIVA1 = 0
    TipoIVA2 = 0
    TipoIVA3 = 0
    
    TotalFac = 0

    Sql = "select esfactadvinterna from advpartes inner join rsocios on advpartes.codsocio = rsocios.codsocio where " & cadWHERE
    EsFactADVInterna = DevuelveValor(Sql)


    'Agrupar el importe bruto por tipos de iva
    cadWhere1 = Replace(cadWHERE, "advpartes", "advpartes_lineas")
    
    ' si la facturacion es interna el codigo de iva es el exento de parametros
    ' sino es el del articulo
    If EsFactADVInterna = 1 Then
    
        Set vSeccion = New CSeccion
        If vSeccion.LeerDatos(vParamAplic.SeccionADV) Then
            If vSeccion.AbrirConta Then
                ' codigo de iva de facturas internas de adv
                Sql = DevuelveDesdeBDNew(cConta, "tiposiva", "nombriva", "codigiva", vParamAplic.CodIvaExeADV, "N")
                
                If Sql = "" Then
                    MsgBox "No está parametrizado el código de iva de socios con facturación interna o no existe en contabilidad. Revise.", vbExclamation
                    CalcularDatosAlbaranVenta = True
                    Set vSeccion = Nothing
                    Exit Function
                End If
            Else
                MsgBox "No está parametrizada la sección de adv en parámetros. Revise.", vbExclamation
                CalcularDatosAlbaranVenta = True
                Set vSeccion = Nothing
                Exit Function
            End If
        End If
        Set vSeccion = Nothing
    
        Sql = "SELECT " & vParamAplic.CodIvaExeADV & " as codigiva , sum(importel) as bruto"
    Else
        Sql = "SELECT advpartes_lineas.codigiva, sum(importel) as bruto"
    End If
    
    Sql = Sql & " FROM advpartes_lineas "
    Sql = Sql & " WHERE " & cadWhere1
    Sql = Sql & " GROUP BY 1 "
    Sql = Sql & " ORDER BY 1 "

    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText

    TotBruto = 0
    TotNeto = 0
    TotImpIVA = 0
    vBruto = 0
    vNeto = 0
    I = 1

    If Not Rs.EOF Then Rs.MoveFirst
    IvaAnt = Rs.Fields(0).Value
    While Not Rs.EOF
            
            IvaAnt = Rs.Fields(0).Value
            
            vBruto = Rs.Fields(1).Value
            TotBruto = TotBruto + vBruto
            ImpBImIVA = vBruto
        

            'Obtener el % de IVA
'            cadAux = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", CStr(IvaAnt), "N")
            cadAux = 0
            Set vSeccion = New CSeccion
            If vSeccion.LeerDatos(vParamAplic.SeccionADV) Then
                If vSeccion.AbrirConta Then
                    cadAux = DevuelveDesdeBDNew(cConta, "tiposiva", "porceiva", "codigiva", CStr(IvaAnt), "N")
                End If
                vSeccion.CerrarConta
            End If
            Set vSeccion = Nothing


            'aplicar el IVA a la base imponible de ese tipo
            impiva = CalcularPorcentaje(ImpBImIVA, CCur(cadAux), 2)
            
            'sumamos todos los IVAS para sumarselo a la base imponible total de la factura
            'los vamos acumulando
            TotImpIVA = TotImpIVA + impiva

'--monica:preguntar manolo
'            If CInt(Me.Combo1(0).ListIndex) = 2 Then  ' tipoivac 0=normal 1=exento 2=recargo equivalencia
'                'Obtener el % de RECARGO
'                cadAux1 = DevuelveDesdeBDNew(cConta, "tiposiva", "porcerec", "codigiva", CStr(IvaAnt), "N")
'
'                'aplicar el RECARGO a la base imponible de ese tipo
'                ImpREC = CalcularPorcentaje(ImpBImIVA, CCur(cadAux1), 2)
'
'                'sumamos todos los RECARGOS para sumarselo a la base imponible total de la factura
'                'los vamos acumulando
'                TotImpIVA = TotImpIVA + ImpREC
'            Else
                cadAux1 = "0"
                ImpREC = 0
'            End If


            Select Case I
                Case 1  'IVA 1
                    TipoIVA1 = IvaAnt 'RS!codigiva

                    BaseIVA1 = ImpBImIVA 'BASE IMPONIBLE

                    PorceIVA1 = cadAux '% de IVA

                    'Importe total con IVA
                    ImpIVA1 = impiva
                    
                    PorceREC1 = cadAux1 '% de REC

                    'Importe total con RECARGO
                    ImpREC1 = ImpREC

                Case 2  'IVA 2
                    TipoIVA2 = IvaAnt 'RS!codigiva

                    BaseIVA2 = ImpBImIVA 'BASE IMPONIBLE

                    PorceIVA2 = cadAux '% de IVA

                    'Importe total con IVA
                    ImpIVA2 = impiva

                    PorceREC2 = cadAux1 '% de REC

                    'Importe total con RECARGO
                    ImpREC2 = ImpREC
                Case 3  'IVA 3
                    TipoIVA3 = IvaAnt 'RS!codigiva

                    BaseIVA3 = ImpBImIVA 'BASE IMPONIBLE

                    PorceIVA3 = cadAux '% de IVA

                    'Importe total con IVA
                    ImpIVA3 = impiva
                    
                    PorceREC3 = cadAux1 '% de REC

                    'Importe total con RECARGO
                    ImpREC3 = ImpREC
            End Select
            
            
            I = I + 1
        
        
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing


    'Base Imponible
    BaseImp = TotBruto

    'TOTAL de la factura
    TotalFac = BaseImp + TotImpIVA

    'ACTUALIZAMOS EL ALBARAN (tabla albaranes)
    
    For I = 9 To 31
        Text1(I).Text = ""
    Next I
    
    If BaseImp <> 0 Then Text1(12).Text = BaseImp
    
    If BaseIVA1 <> 0 Then Text1(15).Text = Format(BaseIVA1, "###,###,##0.00")
    If ImpIVA1 <> 0 Then Text1(16).Text = Format(ImpIVA1, "###,###,##0.00")
    If ImpREC1 <> 0 Then Text1(30).Text = Format(ImpREC1, "###,###,##0.00")
    If TipoIVA1 <> 0 Then Text1(13).Text = TipoIVA1
    If PorceREC1 <> 0 Then Text1(31).Text = Format(PorceREC1, "##0.00")
    If PorceIVA1 <> 0 Then Text1(14).Text = Format(PorceIVA1, "##0.00")
    
    If BaseIVA2 <> 0 Then Text1(19).Text = Format(BaseIVA2, "###,###,##0.00")
    If ImpIVA2 <> 0 Then Text1(20).Text = Format(ImpIVA2, "###,###,##0.00")
    If ImpREC2 <> 0 Then Text1(28).Text = Format(ImpREC2, "###,###,##0.00")
    If TipoIVA2 <> 0 Then Text1(17).Text = TipoIVA2
    If PorceIVA2 <> 0 Then Text1(18).Text = Format(PorceIVA2, "##0.00")
    If PorceREC2 <> 0 Then Text1(29).Text = Format(PorceREC2, "##0.00")
    
    If BaseIVA3 <> 0 Then Text1(23).Text = Format(BaseIVA3, "###,###,##0.00")
    If ImpIVA3 <> 0 Then Text1(24).Text = Format(ImpIVA3, "###,###,##0.00")
    If ImpREC3 <> 0 Then Text1(26).Text = Format(ImpREC3, "###,###,##0.00")
    If TipoIVA3 <> 0 Then Text1(21).Text = TipoIVA3
    If PorceIVA3 <> 0 Then Text1(22).Text = Format(PorceIVA3, "##0.00")
    If PorceREC3 <> 0 Then Text1(27).Text = Format(PorceREC3, "##0.00")
    
    If TotBruto <> 0 Then Text1(10).Text = Format(TotBruto, "###,###,##0.00")
    If TotalFac <> 0 Then Text1(25).Text = Format(TotalFac, "###,###,##0.00")

    CalcularDatosAlbaranVenta = True

ECalcular:
    If Err.Number <> 0 Then
        CalcularDatosAlbaranVenta = False
    Else
        CalcularDatosAlbaranVenta = True
    End If
End Function


Private Sub PonerCamposSocio()
Dim cad As String
Dim Cad1 As String
Dim NumRegis As Long
Dim Rs As ADODB.Recordset


    If Text1(6).Text = "" Then Exit Sub
    
    If Not (Modo = 3 Or Modo = 4) Then Exit Sub

    cad = "rcampos.codsocio = " & DBSet(Text1(6).Text, "N") & " and rcampos.fecbajas is null"
     
    Cad1 = "select count(*) from rcampos where " & cad
     
    NumRegis = TotalRegistros(Cad1)
    
    If NumRegis = 0 Then Exit Sub
    If NumRegis = 1 Then
        Cad1 = "select codcampo from rcampos where " & cad
        Set Rs = New ADODB.Recordset
        Rs.Open Cad1, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not Rs.EOF Then
            Text1(4).Text = DBLet(Rs.Fields(0).Value)
            PonerDatosCampo Text1(4).Text
        End If
    Else
        Set frmMens = New frmMensajes
        frmMens.cadWHERE = " and " & cad '"rcampos.codsocio = " & NumCod & " and rcampos.fecbajas is null"
        frmMens.campo = Text1(4).Text
        frmMens.OpcionMensaje = 7 '6
        frmMens.Show vbModal
        Set frmMens = Nothing
    End If
End Sub


Private Sub PonerDatosCampo(campo As String)
Dim cad As String
Dim Cad1 As String
Dim NumRegis As Long
Dim Rs As ADODB.Recordset


    If campo = "" Then Exit Sub
    
'    If Not (Modo = 3 Or Modo = 4) Then Exit Sub

    cad = "rcampos.codcampo = " & DBSet(campo, "N") & " and rcampos.fecbajas is null"
     
    Cad1 = "select rcampos.codparti, rpartida.nomparti, rpartida.codzonas, rzonas.nomzonas, "
    Cad1 = Cad1 & " rpueblos.despobla from rcampos, rpartida, rzonas, rpueblos "
    Cad1 = Cad1 & " where " & cad
    Cad1 = Cad1 & " and rcampos.codparti = rpartida.codparti "
    Cad1 = Cad1 & " and rpartida.codzonas = rzonas.codzonas "
    Cad1 = Cad1 & " and rpartida.codpobla = rpueblos.codpobla "
     
    Set Rs = New ADODB.Recordset
    Rs.Open Cad1, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    Text2(4).Text = ""
    If Not Rs.EOF Then
        Text1(4).Text = campo
        PonerFormatoEntero Text1(4)
        Text2(4).Text = DBLet(Rs.Fields(1).Value, "T") ' nombre de partida
    End If
    
    Set Rs = Nothing
    
End Sub


Private Function RecalcularImportes(Numparte As String, LitrosRea As String, MenError As String) As Boolean
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim cantidad As Currency
Dim Importe As Currency
Dim vCArticuloADV As CArticuloADV
Dim vCstock As CStockADV
Dim B As Boolean

    On Error GoTo eRecalcularImportes


    RecalcularImportes = False

    B = True
    
    Sql = "select * from advpartes_lineas where numparte = " & DBSet(Numparte, "N")
    Sql = Sql & " order by numlinea "
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not Rs.EOF And B
        Set vCArticuloADV = New CArticuloADV
        
        If vCArticuloADV.LeerDatos(DBLet(Rs!codArtic, "T")) Then
            cantidad = DBLet(Rs!cantidad, "N")
            If vCArticuloADV.TipoProd = 0 Then ' solo en el caso de que sea producto
                cantidad = Round2(DBLet(Rs!dosishab, "N") / 1000 * CCur(LitrosRea), 3)
            End If
            Importe = Round2(cantidad * DBLet(Rs!preciove, "N"), 2)
        
            Set vCstock = New CStockADV
            
            txtAux(5).Text = Rs!codArtic
            txtAux(4).Text = Rs!codAlmac
            txtAux(7).Text = Rs!preciove
            txtAux(6).Text = cantidad
            
            txtAux(9).Text = Importe
            
            
            ModificaLineas = 2
            
            If Not InicializarCStock(vCstock, "S", , Rs) Then B = False
            
            If B Then B = InicializarCStock(vCstock, "E", , Rs)

            If B Then
                B = vCstock.DevolverStock 'eliminamos de advsmoval y devolvemos stock valores anteriores
                'ahora leemos los valores nuevos
                If B Then B = InicializarCStock(vCstock, "S", , Rs)
                'insertamos en smoval y actualizamos stock a los valores nuevos
'                txtAux(6).Text = Format(cantidad, "###,##0.000")
'                txtAux(9).Text = Format(Importe, "##,###,###0.00")
'                txtAux(7).Text = Format(DBLet(Rs!preciove, "N"), "###,##0.0000")
                
'                vCstock.cantidad = CSng(ComprobarCero(txtAux(6).Text))
                If B Then B = vCstock.ActualizarStock()
        
                'actualizar la linea de Albaran
                If B Then
                    Sql = "UPDATE advpartes_lineas Set codalmac = " & txtAux(4).Text & ", codartic=" & DBSet(txtAux(5).Text, "T") & ", "
                    Sql = Sql & "ampliaci=" & DBSet(Rs!ampliaci, "T") & ", "
                    Sql = Sql & "cantidad= " & DBSet(txtAux(6).Text, "N") & ", "
                    Sql = Sql & "preciove= " & DBSet(txtAux(7).Text, "N") & ", " 'precio
                    Sql = Sql & "dosishab= " & DBSet(Rs!dosishab, "N") & ", " ' dosis habitual
                    Sql = Sql & "importel= " & DBSet(txtAux(9).Text, "N") & ", " 'Importe
                    Sql = Sql & "codigiva= " & DBSet(Rs!Codigiva, "N") & " " 'codigo de iva
                    Sql = Sql & Replace(ObtenerWhereCP(True), NombreTabla, "advpartes_lineas") & " AND numlinea=" & DBLet(Rs!numlinea, "N")
                    conn.Execute Sql
                End If
            End If
            Set vCstock = Nothing
        
        
            ModificaLineas = 0
        
        
        
        
        End If
        
        Set vCArticuloADV = Nothing
        
        Rs.MoveNext
    Wend
    
    Set Rs = Nothing
    
    RecalcularImportes = B
    
    Exit Function
    
eRecalcularImportes:
    MenError = MenError & "Error en Recalcular Importes. " & vbCrLf & Err.Description
End Function


Private Function InicializarCStock(ByRef vCstock As CStockADV, TipoM As String, Optional numlinea As String, Optional Rs As ADODB.Recordset) As Boolean
    On Error Resume Next

    vCstock.tipoMov = TipoM
    vCstock.DetaMov = CodTipoMov 'Text1(6).Text
    vCstock.trabajador = CLng(Text1(6).Text) 'guardamos el socio del albaran
    vCstock.Documento = Format(Text1(0).Text, "0000000") 'Nº parte
    vCstock.Fechamov = Text1(1).Text 'Fecha del parte
    vCstock.campo = Text1(4).Text ' campo
    
    '1=Insertar, 2=Modificar
    If ModificaLineas = 1 Or (ModificaLineas = 2 And TipoM = "S") Then
        vCstock.codArtic = txtAux(5).Text
        vCstock.codAlmac = CInt(txtAux(4).Text)
        If ModificaLineas = 1 Then '1=Insertar
            vCstock.cantidad = CSng(ComprobarCero(txtAux(6).Text))
        Else '2=Modificar(Debe haber en stock la diferencia)
            vCstock.cantidad = CSng(ComprobarCero(txtAux(6).Text)) '- DBLet(RS!cantidad, "N")
        End If
        vCstock.Importe = CCur(ComprobarCero(txtAux(9).Text))
    Else
        vCstock.codArtic = DBLet(Rs!codArtic, "T")
        vCstock.codAlmac = DBLet(Rs!codAlmac, "N")
        vCstock.cantidad = CSng(DBLet(Rs!cantidad, "N"))
        vCstock.Importe = CCur(DBLet(Rs!ImporteL, "N"))
    End If
    If ModificaLineas = 1 Then
        vCstock.LineaDocu = CInt(ComprobarCero(numlinea))
    Else
        vCstock.LineaDocu = CInt(DBLet(Rs!numlinea, "N"))
    End If
    If Err.Number <> 0 Then
        MsgBox "No se han podido inicializar la clase para actualizar Stock", vbExclamation
        InicializarCStock = False
    Else
        InicializarCStock = True
    End If
End Function

Private Sub txtAux1_GotFocus(Index As Integer)
    ConseguirFocoLin txtAux1(Index)
End Sub

Private Sub txtAux1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
    KEYdown KeyCode
End Sub

Private Sub txtAux1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtAux1_LostFocus(Index As Integer)
Dim cadMen As String
Dim Sql As String
Dim devuelve As String
Dim B As Boolean
Dim TipoDto As Byte


    'Quitar espacios en blanco
    If Not PerderFocoGnralLineas(txtAux1(Index), ModificaLineas) Then Exit Sub
    
    Select Case Index
        Case 2 'trabajador
            If txtAux1(Index).Text = "" Then Exit Sub
            txtAux1(3).Text = PonerNombreDeCod(txtAux1(Index), "straba", "nomtraba", "codtraba", "N")
            
            If txtAux1(3).Text = "" Then
                cadMen = "No existe el Trabajador: " & txtAux1(Index).Text & vbCrLf
                cadMen = cadMen & "¿Desea crearlo?" & vbCrLf
                If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                    Set frmTra1 = New frmManTraba
                    frmTra1.DatosADevolverBusqueda = "0|1|"
                    txtAux1(Index).Text = ""
                    TerminaBloquear
                    frmTra1.Show vbModal
                    Set frmTra1 = Nothing
                    If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
                Else
                    txtAux1(Index).Text = ""
                End If
                PonerFoco txtAux1(Index)
            Else
                PonerPrecioHoraTrabajador txtAux1(2).Text
'                PonerDatosTrabajador txtAux1(Index).Text
'                PonerFocoBtn Me.cmdAceptar
            End If
            
        Case 4 'horas
            PonerFormatoDecimal txtAux1(Index), 3
            
        Case 5 'precio
            PonerFormatoDecimal txtAux1(Index), 7
       
        Case 6 'importe
            If PonerFormatoDecimal(txtAux1(Index), 1) Then
                PonerFocoBtn cmdAceptar
            End If
    End Select
    
    If Index = 2 And Index = 4 Or Index = 5 Or Index = 6 Then
        txtAux1(6).Text = CalcularImporte(txtAux1(4).Text, txtAux1(5).Text, "", "", 0, "0")
        PonerFormatoDecimal txtAux1(6), 1
    End If
    
End Sub

Private Sub PonerDatosTrabajador(Traba As String)
Dim Sql As String
Dim Rs As ADODB.Recordset

    On Error Resume Next

    Sql = "select niftraba, teltraba, movtraba from straba where codtraba = " & DBSet(Traba, "N")
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    txtAux1(4).Text = ""
    txtAux1(5).Text = ""
    txtAux1(6).Text = ""
    
    If Not Rs.EOF Then
        txtAux1(4).Text = DBLet(Rs!niftraba, "T")
        txtAux1(5).Text = DBLet(Rs!teltraba, "T")
        txtAux1(6).Text = DBLet(Rs!movtraba, "T")
    End If
    
    Set Rs = Nothing
    
End Sub


Private Sub PonerPrecioHoraTrabajador(Traba As String)
Dim Sql As String
Dim Rs As ADODB.Recordset

    On Error Resume Next

    Sql = "select impsalar from salarios inner join straba on salarios.codcateg = straba.codcateg where codtraba = " & DBSet(Traba, "N")
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    txtAux1(5).Text = ""
    
    If Not Rs.EOF Then
        txtAux1(5).Text = DBLet(Rs!impsalar, "N")
        PonerFormatoDecimal txtAux1(5), 7
    End If
    
    Set Rs = Nothing
    
End Sub


Private Sub CalcularTotales()
'calcula la cantidad total y el importe total para los
'registros mostrados de cada artículo
Dim Sql As String
Dim Rs As ADODB.Recordset
    
    On Error GoTo ErrTotales
'    If cadSelGrid = "" Then Exit Sub
    
    If Data1.Recordset.EOF Then Exit Sub
    
    
    Sql = "SELECT sum(importel) as totImporte from advpartes_trabajador "
    Sql = Sql & " where advpartes_trabajador.numparte = " & Data1.Recordset!Numparte

    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not Rs.EOF Then
        Text2(1).Text = DBLet(Rs!totImporte, "N")
        If ComprobarCero(Text2(1).Text) = 0 Then
            Text2(1).Text = ""
        Else
            Text2(1).Text = Format(Text2(1).Text, FormatoImporte)
        End If
        DoEvents
    End If
    
    Rs.Close
    Set Rs = Nothing
    
    Exit Sub
    
ErrTotales:
    MuestraError Err.Number, "Calcular totales.", Err.Description
End Sub





