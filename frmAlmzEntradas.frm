VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmAlmzEntradas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Entradas de Almazara"
   ClientHeight    =   7605
   ClientLeft      =   45
   ClientTop       =   30
   ClientWidth     =   10080
   Icon            =   "frmAlmzEntradas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7605
   ScaleWidth      =   10080
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameBotonGnral 
      Height          =   705
      Left            =   90
      TabIndex        =   76
      Top             =   135
      Width           =   3585
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   210
         TabIndex        =   77
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
      TabIndex        =   74
      Top             =   135
      Width           =   1335
      Begin MSComctlLib.Toolbar Toolbar5 
         Height          =   330
         Left            =   210
         TabIndex        =   75
         Top             =   180
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   2
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Cambiar Rendimiento"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Asignación de Precios Masiva"
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame FrameDesplazamiento 
      Height          =   705
      Left            =   5160
      TabIndex        =   72
      Top             =   135
      Width           =   2415
      Begin MSComctlLib.Toolbar ToolbarDes 
         Height          =   330
         Left            =   240
         TabIndex        =   73
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
      Index           =   0
      Left            =   7695
      TabIndex        =   71
      Top             =   270
      Width           =   1605
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   2625
      Left            =   90
      TabIndex        =   39
      Top             =   4185
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   4630
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   4
      TabHeight       =   520
      ForeColor       =   12582912
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Incidencias"
      TabPicture(0)   =   "frmAlmzEntradas.frx":000C
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "FrameAux1"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Gastos"
      TabPicture(1)   =   "frmAlmzEntradas.frx":0028
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "FrameAux2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.Frame FrameAux2 
         BorderStyle     =   0  'None
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
         Height          =   2070
         Left            =   90
         TabIndex        =   46
         Top             =   420
         Width           =   8830
         Begin VB.Frame FrameToolAux2 
            Height          =   600
            Left            =   45
            TabIndex        =   68
            Top             =   0
            Width           =   1500
            Begin MSComctlLib.Toolbar ToolAux 
               Height          =   330
               Index           =   2
               Left            =   135
               TabIndex        =   69
               Top             =   180
               Width           =   1110
               _ExtentX        =   1958
               _ExtentY        =   582
               ButtonWidth     =   609
               ButtonHeight    =   582
               AllowCustomize  =   0   'False
               Style           =   1
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
         End
         Begin VB.TextBox Text2 
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
            Height          =   360
            Index           =   5
            Left            =   6345
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   53
            Text            =   "Importe total"
            Top             =   135
            Width           =   1920
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
            Height          =   280
            Index           =   8
            Left            =   4545
            MaxLength       =   10
            TabIndex        =   52
            Tag             =   "Importe|N|S|||rhisfruta_gastos|importe|###,##0.00|N|"
            Text            =   "Importe"
            Top             =   1665
            Visible         =   0   'False
            Width           =   705
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
            Height          =   280
            Index           =   9
            Left            =   1830
            MaxLength       =   2
            TabIndex        =   51
            Tag             =   "Cod.Gasto|N|N|0|99|rhisfruta_gastos|codgasto|00||"
            Text            =   "Ga"
            Top             =   1665
            Visible         =   0   'False
            Width           =   705
         End
         Begin VB.CommandButton btnBuscar 
            Appearance      =   0  'Flat
            Caption         =   "+"
            Height          =   360
            Index           =   0
            Left            =   2565
            MaskColor       =   &H00000000&
            TabIndex        =   50
            ToolTipText     =   "Buscar Gasto"
            Top             =   1665
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.TextBox Text2 
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
            Height          =   280
            Index           =   7
            Left            =   2790
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   49
            Text            =   "Nombre gasto"
            Top             =   1665
            Visible         =   0   'False
            Width           =   1740
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
            Height          =   280
            Index           =   10
            Left            =   180
            MaxLength       =   7
            TabIndex        =   48
            Tag             =   "Num.Albaran|N|N|||rhisfruta_gastos|numalbar||S|"
            Text            =   "numalbar"
            Top             =   1665
            Visible         =   0   'False
            Width           =   855
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
            Height          =   280
            Index           =   11
            Left            =   1080
            MaxLength       =   6
            TabIndex        =   47
            Tag             =   "Linea|N|N|0|999999|rhisfruta_gastos|numlinea|000000|S|"
            Text            =   "linea"
            Top             =   1665
            Visible         =   0   'False
            Width           =   705
         End
         Begin MSDataGridLib.DataGrid DataGridAux 
            Bindings        =   "frmAlmzEntradas.frx":0044
            Height          =   1320
            Index           =   2
            Left            =   45
            TabIndex        =   54
            Top             =   630
            Width           =   8595
            _ExtentX        =   15161
            _ExtentY        =   2328
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
            Index           =   2
            Left            =   1440
            Top             =   90
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
         Begin VB.Label Label3 
            Caption         =   "TOTAL  GASTOS:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   4545
            TabIndex        =   55
            Top             =   180
            Width           =   1875
         End
      End
      Begin VB.Frame FrameAux1 
         BorderStyle     =   0  'None
         Caption         =   "Incidencias"
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
         Height          =   2160
         Left            =   -74910
         TabIndex        =   40
         Top             =   360
         Width           =   9010
         Begin VB.Frame FrameToolAux1 
            Height          =   600
            Left            =   45
            TabIndex        =   66
            Top             =   45
            Width           =   1185
            Begin MSComctlLib.Toolbar ToolAux 
               Height          =   330
               Index           =   1
               Left            =   135
               TabIndex        =   67
               Top             =   180
               Width           =   930
               _ExtentX        =   1640
               _ExtentY        =   582
               ButtonWidth     =   609
               ButtonHeight    =   582
               AllowCustomize  =   0   'False
               Style           =   1
               _Version        =   393216
               BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
                  NumButtons      =   3
                  BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                     Object.ToolTipText     =   "Nuevo"
                     Object.Tag             =   "2"
                  EndProperty
                  BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                     Enabled         =   0   'False
                     Object.Visible         =   0   'False
                     Object.ToolTipText     =   "Modificar"
                     Object.Tag             =   "2"
                  EndProperty
                  BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                     Object.ToolTipText     =   "Eliminar"
                     Object.Tag             =   "2"
                  EndProperty
               EndProperty
            End
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
            MaxLength       =   16
            TabIndex        =   44
            Tag             =   "Nro.Albarán|N|N|||rhisfruta_incidencia|numalbar|0000000|S|"
            Text            =   "codfor"
            Top             =   1590
            Visible         =   0   'False
            Width           =   555
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
            Left            =   1050
            MaxLength       =   4
            TabIndex        =   43
            Tag             =   "Incidencia|N|N|||rhisfruta_incidencia|codincid|0000|S|"
            Text            =   "inci"
            Top             =   1590
            Visible         =   0   'False
            Width           =   600
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
            Left            =   1950
            TabIndex        =   42
            Text            =   "nombre"
            Top             =   1590
            Visible         =   0   'False
            Width           =   2040
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
            Height          =   360
            Index           =   1
            Left            =   1725
            MaskColor       =   &H00000000&
            TabIndex        =   41
            ToolTipText     =   "Buscar Incidencia"
            Top             =   1590
            Visible         =   0   'False
            Width           =   195
         End
         Begin MSDataGridLib.DataGrid DataGridAux 
            Bindings        =   "frmAlmzEntradas.frx":0059
            Height          =   1320
            Index           =   1
            Left            =   45
            TabIndex        =   45
            Top             =   675
            Width           =   8685
            _ExtentX        =   15319
            _ExtentY        =   2328
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
         Begin MSAdodcLib.Adodc Adoaux 
            Height          =   375
            Index           =   1
            Left            =   1350
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
            Height          =   290
            Index           =   2
            Left            =   180
            MaxLength       =   4
            TabIndex        =   70
            Tag             =   "NumNota|N|N|||rhisfruta_incidencia|numnotac|0000|S|"
            Text            =   "inci"
            Top             =   270
            Visible         =   0   'False
            Width           =   600
         End
      End
   End
   Begin VB.Frame Frame2 
      Height          =   3195
      Index           =   0
      Left            =   90
      TabIndex        =   17
      Top             =   870
      Width           =   9885
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
         Index           =   10
         Left            =   5805
         TabIndex        =   62
         Top             =   1890
         Width           =   885
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
         Index           =   9
         Left            =   4395
         TabIndex        =   61
         Top             =   1890
         Width           =   795
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
         Index           =   8
         Left            =   2925
         TabIndex        =   60
         Top             =   1890
         Width           =   735
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
         Index           =   5
         Left            =   1395
         MaxLength       =   8
         TabIndex        =   59
         Tag             =   "Campo|N|S|||rhisfruta|codcampo|00000000||"
         Text            =   "12345678"
         Top             =   1890
         Width           =   1050
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
         Height          =   360
         Index           =   11
         Left            =   5715
         MaxLength       =   7
         TabIndex        =   2
         Tag             =   "Nro.Muestra|N|S|||rhisfruta|nromuestraalmz|0000000||"
         Text            =   "1234567"
         Top             =   300
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
         Index           =   10
         Left            =   8160
         MaxLength       =   12
         TabIndex        =   7
         Tag             =   "Rendimiento|N|S|||rhisfruta|prliquidalmz|###,##0.0000||"
         Top             =   315
         Width           =   1365
      End
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
         Index           =   2
         Left            =   8160
         TabIndex        =   33
         Top             =   2745
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
         Index           =   7
         Left            =   8160
         MaxLength       =   12
         TabIndex        =   12
         Tag             =   "Rendimiento|N|S|||rhisfruta|prestimado|###,##0.00||"
         Top             =   2340
         Width           =   1365
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
         Height          =   360
         Index           =   1
         Left            =   3330
         MaxLength       =   10
         TabIndex        =   1
         Tag             =   "Fecha albarán|F|N|||rhisfruta|fecalbar|dd/mm/yyyy||"
         Text            =   "1234567890"
         Top             =   300
         Width           =   1290
      End
      Begin VB.ComboBox Combo1 
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
         Left            =   8160
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Tag             =   "Tipo Entrada|N|N|0|3|rhisfruta|tipoentr||N|"
         Top             =   720
         Width           =   1350
      End
      Begin VB.ComboBox Combo1 
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
         Left            =   8160
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Tag             =   "Recolectado|N|N|0|1|rhisfruta|recolect||N|"
         Top             =   1125
         Width           =   1350
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
         Index           =   6
         Left            =   2415
         TabIndex        =   27
         Top             =   1500
         Width           =   4275
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
         Left            =   1395
         MaxLength       =   6
         TabIndex        =   5
         Tag             =   "Código Población|T|S|||rhisfruta|codpobla|||"
         Top             =   1500
         Width           =   960
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
         Left            =   2430
         TabIndex        =   26
         Text            =   "12345678901234567890"
         Top             =   1125
         Width           =   4245
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
         Height          =   480
         Index           =   8
         Left            =   180
         MaxLength       =   250
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Top             =   2595
         Width           =   6495
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
         Left            =   2430
         TabIndex        =   23
         Text            =   "12345678901234567890"
         Top             =   720
         Width           =   4245
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
         Left            =   1395
         MaxLength       =   6
         TabIndex        =   3
         Tag             =   "Variedad|N|N|0|999999|rhisfruta|codvarie|000000||"
         Text            =   "123456"
         Top             =   720
         Width           =   960
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
         Index           =   2
         Left            =   8160
         MaxLength       =   7
         TabIndex        =   11
         Tag             =   "Peso Neto|N|N|||rhisfruta|kilosnet|###,##0||"
         Top             =   1920
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
         Index           =   9
         Left            =   8160
         MaxLength       =   7
         TabIndex        =   10
         Tag             =   "Peso Bruto|N|N|||rhisfruta|kilosbru|###,##0||"
         Top             =   1515
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
         Index           =   4
         Left            =   1395
         MaxLength       =   6
         TabIndex        =   4
         Tag             =   "Nombre|N|N|||rhisfruta|codsocio|000000||"
         Text            =   "123456"
         Top             =   1125
         Width           =   960
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
         Height          =   360
         Index           =   0
         Left            =   1395
         MaxLength       =   7
         TabIndex        =   0
         Tag             =   "Nro.Albarán|N|S|||rhisfruta|numalbar|0000000|S|"
         Text            =   "1234567"
         Top             =   300
         Width           =   975
      End
      Begin VB.Label Label9 
         Caption         =   "Sub."
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
         Left            =   5295
         TabIndex        =   65
         Top             =   1890
         Width           =   705
      End
      Begin VB.Label Label8 
         Caption         =   "Parc."
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
         Left            =   3840
         TabIndex        =   64
         Top             =   1890
         Width           =   705
      End
      Begin VB.Label Label7 
         Caption         =   "Pol."
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
         Left            =   2445
         TabIndex        =   63
         Top             =   1920
         Width           =   570
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   4
         Left            =   1125
         ToolTipText     =   "Buscar Pueblo"
         Top             =   1920
         Width           =   240
      End
      Begin VB.Label Label5 
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
         Left            =   180
         TabIndex        =   58
         Top             =   1920
         Width           =   705
      End
      Begin VB.Label Label4 
         Caption         =   "N.Muestra"
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
         Left            =   4635
         TabIndex        =   57
         Top             =   315
         Width           =   1020
      End
      Begin VB.Label Label1 
         Caption         =   "Precio"
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
         Left            =   6765
         TabIndex        =   56
         Top             =   345
         Width           =   1185
      End
      Begin VB.Label Label1 
         Caption         =   "Litros Aceite"
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
         Left            =   6765
         TabIndex        =   34
         Top             =   2775
         Width           =   1365
      End
      Begin VB.Label Label1 
         Caption         =   "Rendimiento"
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
         Left            =   6765
         TabIndex        =   32
         Top             =   2385
         Width           =   1185
      End
      Begin VB.Label Label14 
         Caption         =   "Fecha"
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
         Left            =   2445
         TabIndex        =   31
         Top             =   315
         Width           =   615
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   0
         Left            =   3075
         Picture         =   "frmAlmzEntradas.frx":0071
         ToolTipText     =   "Buscar fecha"
         Top             =   300
         Width           =   240
      End
      Begin VB.Label Label2 
         Caption         =   "Tipo Entrada"
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
         Left            =   6765
         TabIndex        =   30
         Top             =   780
         Width           =   1440
      End
      Begin VB.Label Label11 
         Caption         =   "Recolectado"
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
         Left            =   6765
         TabIndex        =   29
         Top             =   1185
         Width           =   1350
      End
      Begin VB.Label Label23 
         Caption         =   "Proced."
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
         Left            =   180
         TabIndex        =   28
         Top             =   1530
         Width           =   795
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   3
         Left            =   1125
         ToolTipText     =   "Buscar Pueblo"
         Top             =   1530
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   1
         Left            =   1125
         ToolTipText     =   "Buscar Socio"
         Top             =   1155
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Albarán"
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
         TabIndex        =   25
         Top             =   315
         Width           =   1185
      End
      Begin VB.Label Label10 
         Caption         =   "Variedad"
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
         Left            =   180
         TabIndex        =   24
         Top             =   765
         Width           =   870
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   1125
         ToolTipText     =   "Buscar Variedad"
         Top             =   780
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Kilos Neto"
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
         Left            =   6765
         TabIndex        =   22
         Top             =   1965
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Kilos Bruto"
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
         Left            =   6765
         TabIndex        =   21
         Top             =   1545
         Width           =   1320
      End
      Begin VB.Label Label6 
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
         Index           =   1
         Left            =   180
         TabIndex        =   20
         Top             =   1125
         Width           =   690
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
         Left            =   180
         TabIndex        =   19
         Top             =   2310
         Width           =   1485
      End
      Begin VB.Image imgZoom 
         Height          =   240
         Index           =   0
         Left            =   1710
         ToolTipText     =   "Zoom descripción"
         Top             =   2325
         Width           =   240
      End
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   2715
      TabIndex        =   36
      Text            =   "12345678901234567890"
      Top             =   1605
      Width           =   3300
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   1860
      TabIndex        =   35
      Text            =   "12345678901234567890"
      Top             =   1920
      Width           =   4140
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Index           =   0
      Left            =   150
      TabIndex        =   15
      Top             =   6885
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
         TabIndex        =   16
         Top             =   180
         Width           =   2655
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
      Left            =   8910
      TabIndex        =   14
      Top             =   7035
      Width           =   1035
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
      Left            =   7770
      TabIndex        =   13
      Top             =   7035
      Width           =   1035
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   2130
      Top             =   4890
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
      Left            =   8910
      TabIndex        =   18
      Top             =   7035
      Visible         =   0   'False
      Width           =   1035
   End
   Begin MSComctlLib.Toolbar ToolbarAyuda 
      Height          =   330
      Left            =   9510
      TabIndex        =   78
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
   Begin VB.Label Label6 
      Caption         =   "Campo"
      Height          =   255
      Index           =   0
      Left            =   870
      TabIndex        =   38
      Top             =   1590
      Width           =   690
   End
   Begin VB.Image imgBuscar 
      Height          =   240
      Index           =   2
      Left            =   1590
      ToolTipText     =   "Buscar Campo"
      Top             =   1590
      Width           =   240
   End
   Begin VB.Label Label6 
      Caption         =   "Población"
      Height          =   255
      Index           =   2
      Left            =   870
      TabIndex        =   37
      Top             =   1935
      Width           =   900
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
      Begin VB.Menu mnCambiarRdto 
         Caption         =   "&Cambiar Rdto"
         HelpContextID   =   2
         Shortcut        =   ^R
      End
      Begin VB.Menu mnAsignacion 
         Caption         =   "&Asignación de Precios"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnBarra2 
         Caption         =   "-"
      End
      Begin VB.Menu mnImprimir 
         Caption         =   "&Imprimir"
         Shortcut        =   ^I
      End
      Begin VB.Menu mnSalir 
         Caption         =   "&Salir"
         Shortcut        =   ^S
      End
   End
End
Attribute VB_Name = "frmAlmzEntradas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' +-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+
' +-+- Autor: MANOLO                   -+-+
' +-+- Menú: CLIENTES                  -+-+
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

Private Const IdPrograma = 10008


'Dim T1 As Single

Public DatosADevolverBusqueda As String    'Tindrà el nº de text que vol que torne, empipat
Public Event DatoSeleccionado(CadenaSeleccion As String)
Public NuevoCodigo As String
Public CodigoActual As String
Public DeConsulta As Boolean

' *** declarar els formularis als que vaig a cridar ***
Private WithEvents frmB As frmBuscaGrid
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmC As frmCal 'calendario fecha
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmZ As frmZoom  'Zoom para campos Text
Attribute frmZ.VB_VarHelpID = -1
Private WithEvents frmMens As frmMensajes ' campos del socio
Attribute frmMens.VB_VarHelpID = -1

'Private WithEvents frmArt As frmManArtic 'articulos
Private WithEvents frmVar As frmComVar 'variedades de comercial
Attribute frmVar.VB_VarHelpID = -1
Private WithEvents frmSoc As frmManSocios 'socios
Attribute frmSoc.VB_VarHelpID = -1
Private WithEvents frmPob As frmManPueblos 'poblaciones
Attribute frmPob.VB_VarHelpID = -1
Private WithEvents frmTra As frmManTranspor 'tranportistas
Attribute frmTra.VB_VarHelpID = -1
Private WithEvents frmTar As frmManTarTra 'tarifas de transporte
Attribute frmTar.VB_VarHelpID = -1
Private WithEvents frmInc As frmManInciden 'incidencias
Attribute frmInc.VB_VarHelpID = -1
Private WithEvents frmCal As frmManCalidades 'calidades
Attribute frmCal.VB_VarHelpID = -1
Private WithEvents frmCam As frmManCampos 'campos
Attribute frmCam.VB_VarHelpID = -1
Private WithEvents frmGas As frmManConcepGasto 'Form Mto de conceptos de gastos
Attribute frmGas.VB_VarHelpID = -1

'
'*****************************************************
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

Private CadB1 As String


Private kCampo As Integer
'-------------------------------------------------------------------------
Private HaDevueltoDatos As Boolean

'Private VieneDeBuscar As Boolean
'Per a quan torna 2 poblacions en el mateix codi Postal. Si ve de pulsar prismatic
'de búsqueda posar el valor de població seleccionada i no tornar a recuperar de la Base de Datos

Dim Gastos As Boolean

Dim CodTipoMov As String
Dim NotaExistente As Boolean

Dim btnPrimero As Byte 'Variable que indica el nº del Botó PrimerRegistro en la Toolbar1
'Dim CadAncho() As Boolean  'array, per a quan cridem al form de llínies
Dim Indice As Byte 'Index del text1 on es poses els datos retornats des d'atres Formularis de Mtos
Dim CadB As String

Dim SinCampos As Boolean

Dim ObservacAnt As String

Private Sub btnBuscar_Click(Index As Integer)
    TerminaBloquear
    Select Case Index
        Case 1 'Incidencia
            Indice = 1
            Set frmInc = New frmManInciden
            frmInc.DatosADevolverBusqueda = "0|1|"
            frmInc.CodigoActual = txtAux(1).Text
            frmInc.Show vbModal
            Set frmInc = Nothing
            PonerFoco txtAux(1)
        
        Case 0 'Conceptos de gastos
            Indice = 9
            Set frmGas = New frmManConcepGasto
            frmGas.DatosADevolverBusqueda = "0|1|"
            frmGas.CodigoActual = txtAux(9).Text
            frmGas.Show vbModal
            Set frmGas = Nothing
            PonerFoco txtAux(9)
    End Select
    If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
End Sub




Private Sub cmdAceptar_Click()
    Screen.MousePointer = vbHourglass
    
    On Error GoTo Error1
    
    Select Case Modo
        Case 1  'BÚSQUEDA
            HacerBusqueda
            
        Case 3 'INSERTAR
            If DatosOK Then
                NotaExistente = False
                InsertarCabecera
                
                TerminaBloquear
                BloqueaRegistro "rhisfruta", "numalbar = " & Trim(Text1(0).Text)
                
                BotonAnyadirLinea 1

            Else
                ModoLineas = 0
            End If
            
        Case 4  'MODIFICAR
            If DatosOK Then
                If ModificaDesdeFormulario2(Me, 1) Then
                    ' si he cambiado observaciones las actualizo en rhisfruta_entradas
                    If ObservacAnt <> Text1(8).Text Then ActualizoObservaciones
                    TerminaBloquear
                    PosicionarData
                End If
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
            End Select
            If NumTabMto = 1 Then CalcularTotalGastos

        ' **************************
    End Select
    Screen.MousePointer = vbDefault
 
Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub

Private Sub Combo1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
'    If PrimeraVez Then PrimeraVez = False
    If PrimeraVez Then
        PrimeraVez = False
        If DatosADevolverBusqueda = "" Then
            PonerModo 0
        Else
            If (DatosADevolverBusqueda <> "") And NuevoCodigo <> "" Then
                BotonAnyadir
            Else
                PonerModo 1 'búsqueda
                ' *** posar de groc els camps visibles de la clau primaria de la capçalera ***
                Text1(0).BackColor = vbYellow 'nro nota
                ' ****************************************************************************
            End If
        End If
    End If

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Form_Unload(Cancel As Integer)
    CheckValueGuardar Me.Name, Me.chkVistaPrevia(0).Value
    If Modo = 4 Then TerminaBloquear
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
Dim I As Integer
    'Icono del formulario
    Me.Icon = frmPpal.Icon

    PrimeraVez = True
    
    ' ICONETS DE LA BARRA
    btnPrimero = 18 'index del botó "primero"
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
        .Buttons(1).Image = 31  'Expandir Añadir, Borrar y Modificar
        .Buttons(2).Image = 13  'Asignacion masiva de precios de liquidacion de albaranes
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
    For I = 1 To ToolAux.Count
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
    For I = 0 To Me.imgBuscar.Count - 1
        Me.imgBuscar(I).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next I
    
    SSTab1.Tab = 0
    
    '[Monica]09/11/2015: solo para el caso de ABN tengo la ayuda de campos
    imgBuscar(4).Enabled = (vParamAplic.Cooperativa = 1)
    imgBuscar(4).visible = (vParamAplic.Cooperativa = 1)
    Text2(8).visible = (vParamAplic.Cooperativa = 1)
    Text2(9).visible = (vParamAplic.Cooperativa = 1)
    Text2(10).visible = (vParamAplic.Cooperativa = 1)
    If vParamAplic.Cooperativa <> 1 Then Text1(5).Top = 2190
    
'    Me.imgBuscar(8).Picture = frmPpal.imgListComun.ListImages(21).Picture
   
    'carga IMAGES de mail
'    For i = 0 To Me.imgMail.Count - 1
'        Me.imgMail(i).Picture = frmPpal.imgListImages16.ListImages(2).Picture
'    Next i
    
    'IMAGES para zoom
    For I = 0 To Me.imgZoom.Count - 1
        Me.imgZoom(I).Picture = frmPpal.imgListImages16.ListImages(3).Picture
    Next I
    
    CodTipoMov = "NOC"
    
    ' el precio de liquidacion por albaran unicamente esta activo para castelduc
    Label1(4).Enabled = (vParamAplic.Cooperativa = 5)
    Label1(4).visible = (vParamAplic.Cooperativa = 5)
    Text1(10).Enabled = (vParamAplic.Cooperativa = 5)
    Text1(10).visible = (vParamAplic.Cooperativa = 5)
    
    '[Monica]22/10/2015: el campo lo utiliza ABN
    ' el campo unicamente está activo si es abn
    Label5.Enabled = (vParamAplic.Cooperativa = 1)
    Label5.visible = (vParamAplic.Cooperativa = 1)
    Text1(5).Enabled = (vParamAplic.Cooperativa = 1)
    Text1(5).visible = (vParamAplic.Cooperativa = 1)
    Label7.visible = (vParamAplic.Cooperativa = 1)
    Label8.visible = (vParamAplic.Cooperativa = 1)
    Label9.visible = (vParamAplic.Cooperativa = 1)
    
    ' *** si n'hi han combos (capçalera o llínies) ***
    CargaCombo
    
    LimpiarCampos   'Neteja els camps TextBox
    ' ******* si n'hi han llínies *******
    DataGridAux(1).ClearFields
    
    '*** canviar el nom de la taula i l'ordenació de la capçalera ***
    NombreTabla = "rhisfruta"
    Ordenacion = " ORDER BY numalbar"
    
    'Mirem com està guardat el valor del check
    chkVistaPrevia(0).Value = CheckValueLeer(Name)
    
    CadB1 = "rhisfruta.codvarie in (select codvarie from variedades, productos, grupopro where grupopro.codgrupo = 5 "
    CadB1 = CadB1 & " and variedades.codprodu = productos.codprodu and productos.codgrupo = grupopro.codgrupo ) "
    
    
    Data1.ConnectionString = conn
    '***** cambiar el nombre de la PK de la cabecera *************
    Data1.RecordSource = "Select * from " & NombreTabla & " where numalbar=-1"
    Data1.Refresh
       
    CargaGrid 1, False
    CargaGrid 2, False
       
    ModoLineas = 0
       
    
'    If DatosADevolverBusqueda = "" Then
'        PonerModo 0
'    Else
'        PonerModo 1 'búsqueda
'        ' *** posar de groc els camps visibles de la clau primaria de la capçalera ***
'        Text1(0).BackColor = vbYellow 'codforfait
'        ' ****************************************************************************
'    End If
End Sub

Private Sub LimpiarCampos()
    On Error Resume Next
    
    limpiar Me   'Mètode general: Neteja els controls TextBox
    lblIndicador.Caption = ""
    
    ' *** si n'hi han combos a la capçalera ***
    Me.Combo1(0).ListIndex = -1
    Me.Combo1(1).ListIndex = -1
    Me.Combo1(2).ListIndex = -1
    ' *****************************************
    Text2(5).Text = ""
    
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
Dim I As Integer, NumReg As Byte
Dim B As Boolean

    On Error GoTo EPonerModo
 
    'Actualisa Iconos Insertar,Modificar,Eliminar
    'ActualizarToolbar Modo, Kmodo
    Modo = Kmodo
    PonerIndicador lblIndicador, Modo, ModoLineas
       
    'Modo 2. N'hi han datos i estam visualisant-los
    '=========================================
    'Posem visible, si es formulari de búsqueda, el botó "Regresar" quan n'hi han datos
    If DatosADevolverBusqueda <> "" Then
        cmdRegresar.visible = (Modo = 2)
    Else
        cmdRegresar.visible = False
    End If
    
    Text1(5).Enabled = True
    Combo1(1).Enabled = True
    
    
    '=======================================
    B = (Modo = 2)
    'Posar Fleches de desplasament visibles
    NumReg = 1
    If Not Data1.Recordset.EOF Then
        If Data1.Recordset.RecordCount > 1 Then NumReg = 2 'Només es per a saber que n'hi ha + d'1 registre
    End If
'    DesplazamientoVisible Me.Toolbar1, btnPrimero, B, NumReg
    DesplazamientoVisible B And Data1.Recordset.RecordCount > 1


    '---------------------------------------------
    
    B = Modo <> 0 And Modo <> 2
    cmdCancelar.visible = B
    cmdAceptar.visible = B
       
    ' añado el tag para engañar para que lo ponga en amarillo
    Text1(8).Tag = "A"
    'Bloqueja els camps Text1 si no estem modificant/Insertant Datos
    'Si estem en Insertar a més neteja els camps Text1
    BloquearText1 Me, Modo
    Text1(8).Tag = ""
    
    ' no puedo buscar por observaciones pq está en la tabla rhisfruta_entradas
    ' y no en la tabla rhisfruta
    Text1(8).Enabled = Not (Modo = 0 Or Modo = 1 Or Modo = 2 Or Modo = 5)
    
    If Text1(8).Enabled Then
        Text1(8).BackColor = vbWhite
    Else
        Text1(8).BackColor = &H80000018 'groc
    End If
    
    '*** si n'hi han combos a la capçalera ***
    BloquearCombo Me, Modo
    '**************************
    
    ' **** si n'hi han imagens de buscar en la capçalera *****
    BloquearImgBuscar Me, Modo, ModoLineas
    BloquearImgZoom Me, Modo, ModoLineas
    BloquearImgFec Me, 0, Modo, ModoLineas
    
    If vParamAplic.Cooperativa <> 1 Then
        Me.imgBuscar(4).Enabled = False
        Me.imgBuscar(4).visible = False
    End If
    
    'Si no es modo lineas Boquear los TxtAux
    For I = 0 To 1
        txtAux(I).visible = False
        BloquearTxt txtAux(I), True
    Next I
    For I = 8 To 11
        txtAux(I).visible = False
        BloquearTxt txtAux(I), True
    Next I
    
    For I = 0 To btnBuscar.Count - 1
        btnBuscar(I).visible = False
        btnBuscar(I).Enabled = True
    Next I
    
'    BloquearTxt Text1(2), Not (Modo = 3 Or Modo = 4 Or Modo = 1)
    chkVistaPrevia(0).Enabled = (Modo <= 2)
    
    PonerLongCampos
    
    
    If (Modo < 2) Or (Modo = 3) Then
        CargaGrid 1, False
        CargaGrid 2, False
    End If
    
    
    B = (Modo = 4) Or (Modo = 2)
    DataGridAux(1).Enabled = B
    DataGridAux(2).Enabled = B
     
    btnBuscar(1).Enabled = (ModoLineas = 1) And (NumTabMto = 1)
    btnBuscar(1).visible = (ModoLineas = 1) And (NumTabMto = 1)
   
    For I = 8 To 9
        BloquearTxt txtAux(I), Not ((ModoLineas = 1) Or (ModoLineas = 2)) And (NumTabMto = 2)
        txtAux(I).visible = ((ModoLineas = 1) Or (ModoLineas = 2)) And (NumTabMto = 2)
    Next I
    btnBuscar(0).Enabled = ((ModoLineas = 1) Or (ModoLineas = 2)) And (NumTabMto = 2)
    btnBuscar(0).visible = ((ModoLineas = 1) Or (ModoLineas = 2)) And (NumTabMto = 2)
        
    
    ' ****** si n'hi han combos a la capçalera ***********************
     If (Modo = 0) Or (Modo = 2) Or (Modo = 5) Then
        Combo1(0).Enabled = False
        Combo1(0).BackColor = &H80000018 'groc
        Combo1(1).Enabled = False
        Combo1(1).BackColor = &H80000018 'groc
    ElseIf (Modo = 1) Or (Modo = 3) Or (Modo = 4) Then
        Combo1(0).Enabled = True
        Combo1(0).BackColor = &H80000005 'blanc
        Combo1(1).Enabled = True
        Combo1(1).BackColor = &H80000005 'blanc
    End If
    ' ****************************************************************
    
    PonerModoOpcionesMenu (Modo) 'Activar opcions menú según modo
    PonerOpcionesMenu   'Activar opcions de menú según nivell
                        'de permisos de l'usuari

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
'Si maxlength=6 para codprove, en modo busqueda Maxlenth=13 'el doble + 1
    
    'para los TEXT1
    PonerLongCamposGnral Me, Modo, 1
End Sub

Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
    PonerOpcionesMenuGeneralNew Me
End Sub

Private Sub PonerModoOpcionesMenu(Modo)
'Actives unes Opcions de Menú i Toolbar según el modo en que estem
Dim B As Boolean, bAux As Boolean
Dim I As Byte
    
    'Barra de CAPÇALERA
    '------------------------------------------
    'b = (Modo = 2 Or Modo = 0 Or Modo = 1)
    B = (Modo = 2 Or Modo = 0)
    'Buscar
    Toolbar1.Buttons(5).Enabled = B
    Me.mnBuscar.Enabled = B
    'Vore Tots
    Toolbar1.Buttons(6).Enabled = B
    Me.mnVerTodos.Enabled = B
    
    'Insertar
    Toolbar1.Buttons(1).Enabled = B And Not DeConsulta
    Me.mnNuevo.Enabled = B And Not DeConsulta
    B = (Modo = 2 And Data1.Recordset.RecordCount > 0) And Not DeConsulta
    'Modificar
    Toolbar1.Buttons(2).Enabled = B
    Me.mnModificar.Enabled = B
    'eliminar
    Toolbar1.Buttons(3).Enabled = B
    Me.mnEliminar.Enabled = B
    
    'cambio de rendimiento
    Toolbar5.Buttons(1).Enabled = True And Not DeConsulta
    Me.mnCambiarRdto.Enabled = True And Not DeConsulta
    
    'asignacion de precios massiva ( sólo esta visible para Castelduc que liquida por precio de albaran )
    Toolbar5.Buttons(2).Enabled = (vParamAplic.Cooperativa = 5) And Not DeConsulta
    Me.mnAsignacion.Enabled = (vParamAplic.Cooperativa = 5) And Not DeConsulta
    Toolbar5.Buttons(2).visible = (vParamAplic.Cooperativa = 5) And Not DeConsulta
    Me.mnAsignacion.visible = (vParamAplic.Cooperativa = 5) And Not DeConsulta
    
    'Imprimir
    Toolbar1.Buttons(8).Enabled = True And Not DeConsulta
    Me.mnImprimir.Enabled = True And Not DeConsulta
     
    ' *** si n'hi han llínies que tenen grids (en o sense tab) ***
'++monica: si insertamos lo he quitado
'    b = (Modo = 3 Or Modo = 4 Or Modo = 2) And Not DeConsulta
    B = (Modo = 2) And Not DeConsulta
    For I = 1 To 2
        ToolAux(I).Buttons(1).Enabled = B
        If B Then bAux = (B And Me.Adoaux(I).Recordset.RecordCount > 0)
        ToolAux(I).Buttons(2).Enabled = bAux
        ToolAux(I).Buttons(3).Enabled = bAux
    Next I
    
End Sub

Private Sub Desplazamiento(Index As Integer)
'Botons de Desplaçament; per a desplaçar-se pels registres de control Data
    If Data1.Recordset.EOF Then Exit Sub
    DesplazamientoData Data1, Index, True
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
Dim SQL As String
Dim tabla As String
    
    ' ********* si n'hi han tabs, dona igual si en datagrid o no ***********
    Select Case Index
        Case 1 'INCIDENCIAS
            SQL = "SELECT rhisfruta_incidencia.numalbar, rhisfruta_incidencia.codincid, rincidencia.nomincid "
            SQL = SQL & " FROM rhisfruta_incidencia, rincidencia "
            If enlaza Then
                SQL = SQL & ObtenerWhereCab(True)
            Else
                SQL = SQL & " WHERE rhisfruta_incidencia.numalbar = -1"
            End If
            SQL = SQL & " and rhisfruta_incidencia.codincid = rincidencia.codincid"
            SQL = SQL & " ORDER BY rhisfruta_incidencia.codincid"
        
        Case 2 'GASTOS
            SQL = "SELECT rhisfruta_gastos.numalbar, rhisfruta_gastos.numlinea, rhisfruta_gastos.codgasto, rconcepgasto.nomgasto, rhisfruta_gastos.importe "
            SQL = SQL & " FROM rhisfruta_gastos, rconcepgasto "
            SQL = SQL & " WHERE rhisfruta_gastos.codgasto = rconcepgasto.codgasto "
    
            If enlaza Then
                SQL = SQL & " and " & ObtenerWhereCP(False)
            Else
                SQL = SQL & " and rhisfruta_gastos.numalbar = -1"
            End If
            SQL = SQL & " ORDER BY rhisfruta_gastos.numlinea"
            
    End Select
    
    MontaSQLCarga = SQL
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
        '   Com la clau principal es única, en posar el sql apuntant
        '   al valor retornat sobre la clau ppal es suficient
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " " & Ordenacion
        ' **********************************
        PonerCadenaBusqueda
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub frmC_Selec(vFecha As Date)
Dim Indice As Byte
    Indice = CByte(Me.imgFec(0).Tag)
    Text1(Indice).Text = Format(vFecha, "dd/mm/yyyy")
End Sub


Private Sub frmCam_DatoSeleccionado(CadenaSeleccion As String)
'Campos
    Text1(Indice).Text = RecuperaValor(CadenaSeleccion, 1) 'codcampo
End Sub

Private Sub frmGas_DatoSeleccionado(CadenaSeleccion As String)
'Conceptos de Gastos
    txtAux(Indice).Text = RecuperaValor(CadenaSeleccion, 1) 'codgasto
    Text2(7).Text = RecuperaValor(CadenaSeleccion, 2) 'nomgasto
End Sub

Private Sub frmPob_DatoSeleccionado(CadenaSeleccion As String)
'Poblacion
    Text1(Indice).Text = RecuperaValor(CadenaSeleccion, 1) 'codpoblacion
    Text2(Indice).Text = RecuperaValor(CadenaSeleccion, 2) 'descripcion
End Sub

Private Sub frmInc_DatoSeleccionado(CadenaSeleccion As String)
'Incidencias
    txtAux(Indice).Text = RecuperaValor(CadenaSeleccion, 1) 'codincid
    txtAux2(Indice).Text = RecuperaValor(CadenaSeleccion, 2) 'descripcion
End Sub

Private Sub frmMens_DatoSeleccionado(CadenaSeleccion As String)
    Text1(5).Text = RecuperaValor(CadenaSeleccion, 1)
    PonerDatosCampo Text1(5)
    If vParamAplic.Cooperativa = 1 Then PonerPoligonoParcela Text1(5)
End Sub

Private Sub frmSoc_DatoSeleccionado(CadenaSeleccion As String)
'Socios
    Text1(Indice).Text = RecuperaValor(CadenaSeleccion, 1) 'codsocio
    Text2(Indice).Text = RecuperaValor(CadenaSeleccion, 2) 'descripcion
End Sub

Private Sub frmTar_DatoSeleccionado(CadenaSeleccion As String)
'Tarifas
    Text1(Indice).Text = RecuperaValor(CadenaSeleccion, 1) 'codtarifa
    Text2(Indice).Text = RecuperaValor(CadenaSeleccion, 2) 'descripcion
End Sub

Private Sub frmTra_DatoSeleccionado(CadenaSeleccion As String)
'Transportistas
    Text1(Indice).Text = RecuperaValor(CadenaSeleccion, 1) 'codtranspor
    Text2(Indice).Text = RecuperaValor(CadenaSeleccion, 2) 'descripcion
End Sub

Private Sub frmVar_DatoSeleccionado(CadenaSeleccion As String)
'Variedades
    Text1(3).Text = RecuperaValor(CadenaSeleccion, 1) 'codvariedad
    Text2(3).Text = RecuperaValor(CadenaSeleccion, 2) 'descripcion
End Sub

Private Sub frmZ_Actualizar(vCampo As String)
     Text1(Indice).Text = vCampo
End Sub

Private Sub imgFec_Click(Index As Integer)
   
   Screen.MousePointer = vbHourglass
   
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

   
   frmC.NovaData = Now
   Select Case Index
        Case 0
            Indice = 1
        Case 1
            Indice = 13
   End Select
   
   Me.imgFec(0).Tag = Indice
   
   PonerFormatoFecha Text1(Indice)
   If Text1(Indice).Text <> "" Then frmC.NovaData = CDate(Text1(Indice).Text)

   Screen.MousePointer = vbDefault
   frmC.Show vbModal
   Set frmC = Nothing
   PonerFoco Text1(Indice)

End Sub

Private Sub imgZoom_Click(Index As Integer)
    
    Set frmZ = New frmZoom

    If Index = 0 Then
        Indice = 8
        frmZ.pTitulo = "Observaciones de la Entrada"
        frmZ.pValor = Text1(Indice).Text
        frmZ.pModo = Modo
    
        frmZ.Show vbModal
        Set frmZ = Nothing
            
        PonerFoco Text1(Indice)
    End If
End Sub



Private Sub mnAsignacion_Click()
    frmAlmzAsigPrecios.Show vbModal
End Sub

Private Sub mnBuscar_Click()
Dim I As Integer
    BotonBuscar
    For I = 0 To Combo1.Count - 1
        Combo1(I).ListIndex = -1 'quan busque, per defecte no seleccione cap tipo de client
    Next I
End Sub

Private Sub mnCambiarRdto_Click()
Dim numalbar As String
    frmAlmzCambioRdto.Show vbModal
    Data1.Refresh
'    NumAlbar = Text1(0).Text
'    Text1(0).Text = ""
'    PosicionarData
'    Text1(0).Text = NumAlbar
'    PosicionarData
    PonerCampos
End Sub

Private Sub mnEliminar_Click()
    BotonEliminar
End Sub

Private Sub mnImprimir_Click()
    Screen.MousePointer = vbHourglass
    frmAlmzListEntradas.CadTag = Text1(0).Text
    frmAlmzListEntradas.OpcionListado = 2
    frmAlmzListEntradas.Show vbModal
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnModificar_Click()
    ' *** repasar el nom de l'adodc, l'index del Field i el camp que te la PK ***
    ' ### [Monica] 26/09/2006 dejamos modificar y eliminar el codigo 0
    'El registre de codi 0 no es pot Modificar ni Eliminar
    'If EsCodigoCero(CStr(Data1.Recordset.Fields(0).Value), FormatoCampo(Text1(0))) Then Exit Sub
    ' ***************************************************************************
    
    If BLOQUEADesdeFormulario2(Me, Data1, 1) Then BotonModificar
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
        Case 5  'Búscar
           mnBuscar_Click
        Case 6  'Tots
            mnVerTodos_Click
        Case 1  'Nou
            mnNuevo_Click
        Case 2  'Modificar
            mnModificar_Click
        Case 3  'Borrar
            mnEliminar_Click
'        Case 11 ' Cambio de Rendimiento
'            mnCambiarRdto_Click
'        Case 12 ' Asignacion de precios masiva
'            mnAsignacion_Click
        Case 8 'Imprimir
            mnImprimir_Click
    End Select
End Sub

Private Sub BotonBuscar()
Dim I As Integer
' ***** Si la clau primaria de la capçalera no es Text1(0), canviar-ho en <=== *****
    If Modo <> 1 Then
        LimpiarCampos
        PonerModo 1
        PonerFoco Text1(0) ' <===
        Text1(0).BackColor = vbYellow ' <===
        ' *** si n'hi han combos a la capçalera ***
        For I = 0 To Combo1.Count - 1
            Combo1(I).ListIndex = -1
        Next I
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
    
'    CadB = ObtenerBusqueda2(Me, 1)
    CadB = ObtenerBusqueda(Me, , , CadB1)
    
    If chkVistaPrevia(0) = 1 Then
        MandaBusquedaPrevia CadB & " and " & CadB1
    ElseIf CadB <> "" Then
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " and " & CadB1 & " " & Ordenacion
        PonerCadenaBusqueda
    Else
        ' *** foco al 1r camp visible de la capçalera que siga clau primaria ***
        PonerFoco Text1(0)
        ' **********************************************************************
    End If
End Sub

Private Sub MandaBusquedaPrevia(CadB As String)
    Dim cad As String
        
    'Cridem al form
    ' **************** arreglar-ho per a vore lo que es desije ****************
    ' NOTA: el total d'amples de ParaGrid, ha de sumar 100
    cad = ""
    cad = cad & ParaGrid(Text1(0), 12, "Albarán")
    cad = cad & ParaGrid(Text1(1), 15, "Fecha")
    cad = cad & "Socio|nomsocio|T||30·"
'    cad = cad & ParaGrid(text1(2), 60, "Descripción")
    cad = cad & "Variedad|nomvarie|T||28·"
    cad = cad & ParaGrid(Text1(6), 15, "Población")
    
    If cad <> "" Then
        
        Screen.MousePointer = vbHourglass
        Set frmB = New frmBuscaGrid
        frmB.vCampos = cad
        cad = "(" & NombreTabla & " inner join variedades on rhisfruta.codvarie = variedades.codvarie) inner join rsocios on rhisfruta.codsocio = rsocios.codsocio "
        frmB.vtabla = cad 'NombreTabla
        frmB.vSQL = Trim(CadB)
        HaDevueltoDatos = False
        frmB.vDevuelve = "0|" '*** els camps que volen que torne ***
        frmB.vTitulo = "Entradas Almazara" ' ***** repasa açò: títol de BuscaGrid *****
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
Dim cad As String
Dim Aux As String
Dim I As Integer
Dim J As Integer

    If Data1.Recordset.EOF Then
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
            cad = cad & Text1(J).Text & "|"
        End If
    Loop Until I = 0
    RaiseEvent DatoSeleccionado(cad)
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
    
    If chkVistaPrevia(0).Value = 1 Then
        MandaBusquedaPrevia CadB1 '""
    Else
        CadenaConsulta = "Select * from " & NombreTabla & " where " & CadB1 & Ordenacion
        PonerCadenaBusqueda
    End If
End Sub

Private Sub BotonAnyadir()
Dim NumF As String
Dim cTabla As String

    LimpiarCampos 'Huida els TextBox
    PonerModo 3
    
    ' ****** Valors per defecte a l'afegir, repasar si n'hi ha
    ' codEmpre i quins camps tenen la PK de la capçalera *******
'    text1(0).Text = SugerirCodigoSiguienteStr("forfaits", "codforfait")
'    FormateaCampo text1(0)
    '******************** canviar taula i camp **************************
    If (DatosADevolverBusqueda <> "") And NuevoCodigo <> "" Then
        NumF = NuevoCodigo
    Else
        NumF = ""
    End If
    '********************************************************************

    cTabla = "(rhisfruta inner join variedades on rhisfruta.codvarie = variedades.codvarie) " & _
             " inner join productos on variedades.codprodu = productos.codprodu and codgrupo = 5 "
    
    NumF = SugerirCodigoSiguienteStr(cTabla, "numalbar")
    Text1(0).Text = NumF
    
    PosicionarCombo Combo1(0), 0
    PosicionarCombo Combo1(1), 0
        
        
            
    Text1(0) = NumF
    PonerFoco Text1(0) '*** 1r camp visible que siga PK ***
    
    '[Monica]02/12/2014:
    Text1(1).Text = Format(Now, "dd/mm/yyyy")
    ' *** si n'hi han camps de descripció a la capçalera ***
    'PosarDescripcions

End Sub

Private Sub BotonModificar()

    Gastos = False

    PonerModo 4


    ObservacAnt = Text1(8).Text

    ' *** bloquejar els camps visibles de la clau primaria de la capçalera ***
    BloquearTxt Text1(0), True
    
    ' *** foco al 1r camp visible que NO siga clau primaria ***
    PonerFoco Text1(1)
End Sub

Private Sub BotonEliminar()
Dim cad As String

    On Error GoTo EEliminar

    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub

    ' *** repasar el nom de l'adodc, l'index del Field i el camp que te la PK ***
    ' ### [Monica] 26/09/2006 dejamos modificar y eliminar el codigo 0
    'El registre de codi 0 no es pot Modificar ni Eliminar
    'If EsCodigoCero(CStr(Data1.Recordset.Fields(0).Value), FormatoCampo(text1(0))) Then Exit Sub
    ' ***************************************************************************

    ' *************** canviar la pregunta ****************
    cad = "¿Seguro que desea eliminar la Entrada?"
    cad = cad & vbCrLf & "Número: " & Data1.Recordset.Fields(0)
    
    If MsgBox(cad, vbQuestion + vbYesNo) = vbYes Then
        On Error GoTo EEliminar
        Screen.MousePointer = vbHourglass
        NumRegElim = Data1.Recordset.AbsolutePosition
        If Not Eliminar Then
            Screen.MousePointer = vbDefault
            Exit Sub
        ElseIf SituarDataTrasEliminar(Data1, NumRegElim) Then
            PonerCampos
        Else
            LimpiarCampos
            PonerModo 0
        End If
    End If
    Screen.MousePointer = vbDefault
    
EEliminar:
    Screen.MousePointer = vbDefault
    If Err.Number > 0 Then MuestraError Err.Number, "Eliminar Proveedor", Err.Description
End Sub




Private Sub PonerCampos()
Dim I As Integer
Dim CodPobla As String, desPobla As String
Dim CPostal As String, desProvi As String, desPais As String

    If Data1.Recordset.EOF Then Exit Sub
    
    PonerCamposForma2 Me, Data1, 1 'opcio=1: posa el format o els camps de la capçalera
    
'    VisualizarDatosCampo Data1.Recordset!CodCampo
    
    Text1(8).Text = DevuelveDesdeBDNew(cAgro, "rhisfruta_entradas", "observac", "numalbar", Text1(0).Text, "N")
    
    ' *** si n'hi han llínies en datagrids ***
    'For i = 0 To DataGridAux.Count - 1
    For I = 1 To 2
        CargaGrid I, True
        If Not Adoaux(I).Recordset.EOF Then _
            PonerCamposForma2 Me, Adoaux(I), 2, "FrameAux" & I
    Next I

    
    ' ************* configurar els camps de les descripcions de la capçalera *************
    Text2(3).Text = PonerNombreDeCod(Text1(3), "variedades", "nomvarie")
    Text2(4).Text = PonerNombreDeCod(Text1(4), "rsocios", "nomsocio")
    If Text1(6).Text <> "" Then
        Text2(6).Text = PonerNombreDeCod(Text1(6), "rpueblos", "despobla") 'poblacion
    End If
    ' ********************************************************************************
    
    CalcularKilogrado
    
    '[Monica]09/11/2015: mostramos poligono y parcela
    If vParamAplic.Cooperativa = 1 Then
        PonerPoligonoParcela Text1(5).Text
    End If
    
    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
    
    PonerModoOpcionesMenu (Modo)
    PonerOpcionesMenu
End Sub

Private Sub cmdCancelar_Click()
Dim I As Integer
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

        Case 4  'Modificar
                TerminaBloquear
                PonerModo 2
                PonerCampos
                ' *** primer camp visible de la capçalera ***
                PonerFoco Text1(0)
        
        Case 5 'LLÍNIES
            Select Case ModoLineas
                Case 1 'afegir llínia
                    ModoLineas = 0
                    ' *** les llínies que tenen datagrid (en o sense tab) ***
                    If NumTabMto = 0 Or NumTabMto = 1 Or NumTabMto = 2 Or NumTabMto = 4 Then
                        DataGridAux(NumTabMto).AllowAddNew = False
                        ' **** repasar si es diu Data1 l'adodc de la capçalera ***
                        'If BLOQUEADesdeFormulario2(Me, Data1, 1) Then BotonModificar 'Modificar
                        LLamaLineas NumTabMto, ModoLineas 'ocultar txtAux
                        DataGridAux(NumTabMto).Enabled = True
                        DataGridAux(NumTabMto).SetFocus

                        ' *** si n'hi han camps de descripció dins del grid, els neteje ***
                        'txtAux2(2).text = ""

                    End If

'                    ' *** si n'hi han tabs ***
'                    SituarTab (NumTabMto + 1)

                    If Not Adoaux(NumTabMto).Recordset.EOF Then
                        Adoaux(NumTabMto).Recordset.MoveFirst
                    End If
                    

                Case 2 'modificar llínies
                    ModoLineas = 0
                    
                    ' *** si n'hi han tabs ***
'                    SituarTab (NumTabMto + 1)
                    LLamaLineas NumTabMto, ModoLineas 'ocultar txtAux
                    PonerModo 4
                    If Not Adoaux(NumTabMto).Recordset.EOF Then
                        ' *** l'Index de Fields es el que canvie de la PK de llínies ***
                        V = Adoaux(NumTabMto).Recordset.Fields(1) 'el 2 es el nº de llinia
                        Adoaux(NumTabMto).Recordset.Find (Adoaux(NumTabMto).Recordset.Fields(1).Name & " =" & V)
                        ' ***************************************************************
                    End If
            End Select
            
            PosicionarData
            
            ' *** si n'hi han llínies en grids i camps fora d'estos ***
            If Not Adoaux(NumTabMto).Recordset.EOF Then
                DataGridAux_RowColChange NumTabMto, 1, 1
            Else
                LimpiarCamposFrame NumTabMto
            End If
    End Select
End Sub

Private Function DatosOK() As Boolean
Dim B As Boolean
'Dim Datos As String

    On Error GoTo EDatosOK

    DatosOK = False
    
    B = CompForm2(Me, 1)
    If Not B Then Exit Function
    
    If Modo = 3 Then
        If ExisteCP(Text1(0)) Then B = False
    End If
    
    If Modo = 3 Or Modo = 4 Then
        '[Monica]09/11/2015: para el caso de abn nom comprobamos que el campo sea del socio y variedad introducida
        If vParamAplic.Cooperativa = 1 Then
        
        Else
            If Not EsCampoSocioVariedad(Text1(5).Text, Text1(4).Text, Text1(3).Text) Then
                MsgBox "El campo no es del socio o de la variedad introducida. Revise.", vbExclamation
                B = False
            End If
        End If
    End If
    ' ************************************************************************************
    DatosOK = B
    
EDatosOK:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function

Private Sub PosicionarData()
Dim cad As String, Indicador As String

    ' *** canviar-ho per tota la PK de la capçalera, no llevar els () ***
    cad = "(numalbar=" & DBSet(Text1(0).Text, "N") & ")"
    
    ' *** gastar SituarData o SituarDataMULTI depenent de si la PK es simple o composta ***
    'If SituarDataMULTI(Data1, cad, Indicador) Then
    If SituarData(Data1, cad, Indicador) Then
        If ModoLineas <> 1 Then PonerModo 2
        lblIndicador.Caption = Indicador
    Else
       LimpiarCampos
       PonerModo 0
    End If
End Sub

Private Function Eliminar() As Boolean
Dim vWhere As String

    On Error GoTo FinEliminar

    conn.BeginTrans
    ' ***** canviar el nom de la PK de la capçalera, repasar codEmpre *******
    vWhere = " WHERE numalbar=" & Data1.Recordset!numalbar
    
    ' ***** elimina les llínies ****
    
    conn.Execute "DELETE FROM rhisfruta_incidencia " & vWhere
        
    conn.Execute "DELETE FROM rhisfruta_entradas " & vWhere
        
    conn.Execute "DELETE FROM rhisfruta_gastos " & vWhere
        
    'Eliminar la CAPÇALERA
    conn.Execute "Delete from " & NombreTabla & vWhere
       
FinEliminar:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Eliminar"
        conn.RollbackTrans
        Eliminar = False
    Else
        conn.CommitTrans
        Eliminar = True
    End If
End Function

Private Sub Text1_GotFocus(Index As Integer)
    kCampo = Index
    ConseguirFoco Text1(Index), Modo
End Sub

' *** si n'hi han combos a la capçalera ***
Private Sub Combo1_GotFocus(Index As Integer)
    If Modo = 1 Then Combo1(Index).BackColor = vbYellow
  
End Sub

Private Sub Combo1_LostFocus(Index As Integer)
Dim B As Boolean

    If Combo1(Index).BackColor = vbYellow Then Combo1(Index).BackColor = vbWhite
    
End Sub

Private Sub Text1_LostFocus(Index As Integer)
Dim cadMen As String
Dim Nuevo As Boolean
Dim SQL As String

    If Not PerderFocoGnral(Text1(Index), Modo) Then Exit Sub
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    
    ' ***************** configurar els LostFocus dels camps de la capçalera *****************
    Select Case Index
        Case 0 'numero de nota
            PonerFormatoEntero Text1(Index)
        
        Case 3 'Variedad
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
                    If (Modo = 3 Or Modo = 4) And Not EsVariedadGrupo5(Text1(Index).Text) Then
                        MsgBox "Esta variedad no es del Grupo de Almazara. Revise.", vbExclamation
                        PonerFoco Text1(Index)
                    End If
                End If
            Else
                Text2(Index).Text = ""
            End If
            
        Case 4 'Socio
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index) = PonerNombreDeCod(Text1(Index), "rsocios", "nomsocio")
                If Text2(Index).Text = "" Then
                    cadMen = "No existe el Socio: " & Text1(Index).Text & vbCrLf
                    cadMen = cadMen & "¿Desea crearlo?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmSoc = New frmManSocios
                        frmSoc.DatosADevolverBusqueda = "0|1|"
'                        frmSoc.NuevoCodigo = Text1(Index).Text
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
                    If EstaSocioDeAlta(Text1(Index)) Then
                        PonerCamposSocioVariedad
                        '[Monica]15/01/2015: si el socio no tiene campos
                        If SinCampos Then
                            MsgBox "El socio no tiene campos dados de alta. Reintroduzca.", vbExclamation
                            Text1(Index).Text = ""
                            Text2(4).Text = ""
                            PonerFoco Text1(Index)
                        End If
                    Else
                        MsgBox "El socio está dado de baja. Reintroduzca.", vbExclamation
                        Text1(Index).Text = ""
                        Text2(4).Text = ""
                        PonerFoco Text1(Index)
                    End If
                End If
            Else
                Text2(Index).Text = ""
            End If
                
        Case 5 'campo
            If PonerFormatoEntero(Text1(Index)) Then
                If Modo = 1 Then Exit Sub
                SQL = ""
                SQL = DevuelveDesdeBDNew(cAgro, "rcampos", "codcampo", "codcampo", Text1(Index).Text, "N")
                If SQL = "" Then
                    cadMen = "No existe el Campo: " & Text1(Index).Text & vbCrLf
                    cadMen = cadMen & "¿Desea crearlo?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmCam = New frmManCampos
                        frmCam.DatosADevolverBusqueda = "0|1|"
'                        frmCamp.NuevoCodigo = Text1(Index).Text
                        Text1(Index).Text = ""
                        TerminaBloquear
                        frmCam.Show vbModal
                        Set frmCam = Nothing
                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
                    Else
                        Text1(Index).Text = ""
                    End If
                    PonerFoco Text1(Index)
                Else
                    If Not EstaCampoDeAlta(Text1(Index).Text) Then
                        MsgBox "El campo no está dado de alta. Reintroduzca.", vbExclamation
                        Text1(Index).Text = ""
                        PonerFoco Text1(Index)
                    Else
                        VisualizarDatosCampo (Text1(Index))
                    End If
                End If
            End If
        
        Case 6 'Depósito
               '¡¡¡¡OJO!!!! SE ACCEDE A LA TABLA DE CAPATACES
'            If PonerFormatoEntero(Text1(Index)) Then
            If Text1(Index).Text <> "" Then
                Text2(Index) = PonerNombreDeCod(Text1(Index), "rpueblos", "despobla")
                If Text2(Index).Text = "" Then
                    cadMen = "No existe el Pueblo: " & Text1(Index).Text & vbCrLf
                    cadMen = cadMen & "¿Desea crearlo?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmPob = New frmManPueblos
                        frmPob.DatosADevolverBusqueda = "0|1|"
                        frmPob.Caption = "Pueblos"
                        frmPob.NuevoCodigo = Text1(Index).Text
                        Text1(Index).Text = ""
                        TerminaBloquear
                        frmPob.Show vbModal
                        Set frmPob = Nothing
                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
                    Else
                        Text1(Index).Text = ""
                    End If
                    PonerFoco Text1(Index)
                End If
            Else
                Text2(Index).Text = ""
            End If
        
        Case 9, 2 'kilos brutos, kilosneto
            PonerFormatoEntero Text1(Index)
            
            CalcularKilogrado
       
        Case 1 ' formato fecha
            '[Monica]28/08/2013: comprobamos que la fecha esté en la campaña
            PonerFormatoFecha Text1(Index), True
       
        Case 7 ' grados
            If Modo = 1 Then Exit Sub
            PonerFormatoDecimal Text1(Index), 3
            
            CalcularKilogrado
            
        Case 10 ' precio de liquidacion de castelduc
            PonerFormatoDecimal Text1(10), 11
            
        Case 11 ' nro de muestra
            PonerFormatoEntero Text1(11)
    End Select
    
    If (Index = 5 Or Index = 3 Or Index = 4) And (Modo = 3 Or Modo = 4) Then
        '[Monica]09/11/2015: solo si es abn no miro si el campo es o no de la variedad
        If vParamAplic.Cooperativa <> 1 Then
            If Not EsCampoSocioVariedad(Text1(5).Text, Text1(4).Text, Text1(3).Text) Then
                MsgBox "El campo no es del socio o de la variedad introducida. Revise.", vbExclamation
            End If
        Else
            If Not EsCampoSocio(Text1(5).Text, Text1(4).Text) Then
                MsgBox "El campo no es del socio. Revise.", vbExclamation
            End If
        End If
    End If
    '[Monica]09/11/2015: solo en le caso de ABN
    If vParamAplic.Cooperativa = 1 Then
        If Text1(5).Text <> "" Then
            PonerPoligonoParcela (Text1(5).Text)
        End If
    End If
    
    ' ***************************************************************************
End Sub


Private Sub PonerPoligonoParcela(campo As String)
Dim SQL As String
Dim Rs As ADODB.Recordset

    Text2(8).Text = ""
    Text2(9).Text = ""
    Text2(10).Text = ""

    If campo = "" Then Exit Sub

    SQL = "select poligono, parcela, subparce from rcampos where codcampo =" & DBSet(campo, "N")
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not Rs.EOF Then
        Text2(8).Text = DBLet(Rs!Poligono)
        Text2(9).Text = DBLet(Rs!Parcela)
        Text2(10).Text = DBLet(Rs!SubParce)
    End If

    Set Rs = Nothing
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = teclaBuscar Then
        If Modo = 1 Or Modo = 3 Or Modo = 4 Then
            Select Case Index
                Case 1: KEYFecha KeyAscii, 0 'fecha
                Case 3: KEYBusqueda KeyAscii, 0 'variedad
                Case 4: KEYBusqueda KeyAscii, 1 'socio
                Case 6: KEYBusqueda KeyAscii, 3 'procedencia
                Case 5: KEYBusqueda KeyAscii, 4 'campo
                Case 8: KEYZoom KeyAscii, 0 'observaciones
            End Select
        End If
    Else
'        If Index <> 3 Or (Index = 3 And Text1(3).Text = "") Then KEYpress KeyAscii
        KEYpress KeyAscii
    End If
End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Alvançar/Retrocedir els camps en les fleches de desplaçament del teclat.
    KEYdown KeyCode
End Sub

Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me
End Sub

Private Sub KEYBusqueda(KeyAscii As Integer, Indice As Integer)
    KeyAscii = 0
    imgBuscar_Click (Indice)
End Sub

Private Sub KEYFecha(KeyAscii As Integer, Indice As Integer)
    KeyAscii = 0
    imgFec_Click (Indice)
End Sub

Private Sub KEYZoom(KeyAscii As Integer, Indice As Integer)
    KeyAscii = 0
    imgZoom_Click (Indice)
End Sub

Private Sub KEYBusquedaAux(KeyAscii As Integer, Indice As Integer)
    KeyAscii = 0
    btnBuscar_Click (Indice)
End Sub

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
Dim SQL As String
Dim vWhere As String
Dim Eliminar As Boolean

    On Error GoTo Error2

    ModoLineas = 3 'Posem Modo Eliminar Llínia
    
    If Modo = 4 Then 'Modificar Capçalera
        cmdAceptar_Click
        If ModoLineas = 0 Then Exit Sub
    End If
       
    NumTabMto = Index
    PonerModo 5, Index

    If Adoaux(Index).Recordset.EOF Then Exit Sub
    If Not SepuedeBorrar(Index) Then Exit Sub
    NumTabMto = Index
    Eliminar = False
   
    vWhere = ObtenerWhereCab(True)
    
    ' ***** independentment de si tenen datagrid o no,
    ' canviar els noms, els formats i el DELETE *****
    Select Case Index
        Case 1 'incidencias
            SQL = "¿Seguro que desea eliminar la Incidencia?"
            SQL = SQL & vbCrLf & "Nombre: " & Adoaux(Index).Recordset!nomincid
            If MsgBox(SQL, vbQuestion + vbYesNo) = vbYes Then
                Eliminar = True
                SQL = "DELETE FROM rhisfruta_incidencia "
                SQL = SQL & vWhere & " AND codincid= " & Adoaux(Index).Recordset!codincid
            End If
        
        Case 2
            SQL = "¿Seguro que desea eliminar el Gasto?"
            SQL = SQL & vbCrLf & "Nombre: " & Adoaux(Index).Recordset!NomGasto
            If MsgBox(SQL, vbQuestion + vbYesNo) = vbYes Then
                Eliminar = True
                SQL = "DELETE FROM rhisfruta_gastos "
                SQL = SQL & vWhere & " AND numlinea = " & Adoaux(Index).Recordset!numlinea
            End If
    End Select

    If Eliminar Then
        NumRegElim = Adoaux(Index).Recordset.AbsolutePosition
        TerminaBloquear
        conn.Execute SQL
        ' *** si n'hi han tabs sense datagrid, posar l'If ***
        CargaGrid Index, True
        If Not SituarDataTrasEliminar(Adoaux(Index), NumRegElim, True) Then
'            PonerCampos
            
        End If
        If BLOQUEADesdeFormulario2(Me, Data1, 1) Then BotonModificar
        ' *** si n'hi han tabs ***
'        SituarTab (NumTabMto + 1)
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
Dim vWhere As String, vtabla As String
Dim anc As Single
Dim I As Integer
    
    ModoLineas = 1 'Posem Modo Afegir Llínia
    
    If (Modo = 3) Or (Modo = 4) Then 'Insertar o Modificar Capçalera
        cmdAceptar_Click
        If ModoLineas = 0 Then Exit Sub
    End If
       
    NumTabMto = Index
    PonerModo 5, Index
    
    ' *** bloquejar la clau primaria de la capçalera ***
    BloquearTxt Text1(0), True

    ' *** posar el nom del les distintes taules de llínies ***
    Select Case Index
        Case 1: vtabla = "rhisfruta_incidencia"
        Case 2: vtabla = "rhisfruta_gastos"
    End Select
    
    vWhere = ObtenerWhereCab(False)
    
    Select Case Index
        Case 1, 2  ' *** pose els index dels tabs de llínies que tenen datagrid ***
            ' *** canviar la clau primaria de les llínies,
            'pasar a "" si no volem que mos sugerixca res a l'afegir ***
            
'            If Index = 0 Then NumF = SugerirCodigoSiguienteStr(vTabla, "numlinea", vWhere)

            AnyadirLinea DataGridAux(Index), Adoaux(Index)
    
            anc = DataGridAux(Index).Top
            If DataGridAux(Index).Row < 0 Then
                anc = anc
                If Index = 1 Then
                    anc = anc + 230
                Else
                    anc = anc + 240
                    
                End If
            Else
                anc = anc + DataGridAux(Index).RowTop(DataGridAux(Index).Row)
                If Index = 1 Then
                    anc = anc + 10
                Else
                    anc = anc + 30
                End If
            End If
            
            LLamaLineas Index, ModoLineas, anc
        
            Select Case Index
                ' *** valor per defecte a l'insertar i formateig de tots els camps ***
                Case 1 'incidencias
                    txtAux(0).Text = Text1(0).Text 'numalbar
                    txtAux(1).Text = "" 'NumF 'codcoste
                    txtAux(2).Text = "0" 'NumF 'numero de nota
                    txtAux2(1).Text = ""
                    For I = 1 To 1
                        BloquearTxt txtAux(I), False
                    Next I
                    BloquearbtnBuscar Me, Modo, ModoLineas, "FrameAux1"
                    PonerFoco txtAux(1)
                
                Case 2 'gastos
                    NumF = SugerirCodigoSiguienteStr(vtabla, "numlinea", vWhere)
                    
                    txtAux(10).Text = Text1(0).Text 'numalbar
                    txtAux(11).Text = NumF 'numlinea
                    txtAux(9).Text = "" 'codcoste
                    txtAux(8).Text = "" ' importe
                    Text2(7).Text = ""
                    BloquearTxt txtAux(8), False
                    BloquearTxt txtAux(9), False
                    
'                    BloquearTxt txtaux(12), False
                    BloquearbtnBuscar Me, Modo, ModoLineas, "FrameAux2"
                    PonerFoco txtAux(9)
                        End Select
            
    End Select
End Sub

Private Sub BotonModificarLinea(Index As Integer)
    Dim anc As Single
    Dim I As Integer
    Dim J As Integer
    
    If Adoaux(Index).Recordset.EOF Then Exit Sub
    If Adoaux(Index).Recordset.RecordCount < 1 Then Exit Sub
    
    ModoLineas = 2 'Modificar llínia
       
    If Modo = 4 Then 'Modificar Capçalera
        cmdAceptar_Click
        If ModoLineas = 0 Then Exit Sub
    End If
       
    NumTabMto = Index
    PonerModo 5, Index
    ' *** bloqueje la clau primaria de la capçalera ***
    BloquearTxt Text1(0), True
  
    Select Case Index
        Case 1, 2 ' *** pose els index de llínies que tenen datagrid (en o sense tab) ***
            If DataGridAux(Index).Bookmark < DataGridAux(Index).FirstRow Or DataGridAux(Index).Bookmark > (DataGridAux(Index).FirstRow + DataGridAux(Index).VisibleRows - 1) Then
                I = DataGridAux(Index).Bookmark - DataGridAux(Index).FirstRow
                DataGridAux(Index).Scroll 0, I
                DataGridAux(Index).Refresh
            End If
              
            anc = DataGridAux(Index).Top
            If DataGridAux(Index).Row < 0 Then
                anc = anc + 210
            Else
                anc = anc + DataGridAux(Index).RowTop(DataGridAux(Index).Row) + 10
            End If
    End Select
    
    Select Case Index
        ' *** valor per defecte al modificar dels camps del grid ***
        Case 1 'incidencias
            For J = 0 To 1
                txtAux(J).Text = DataGridAux(Index).Columns(J).Text
            Next J
            txtAux2(1).Text = DataGridAux(Index).Columns(2).Text
            For I = 1 To 1
                BloquearTxt txtAux(I), True
            Next I
            BloquearbtnBuscar Me, Modo, ModoLineas, "FrameAux1"
        
        Case 2 'gastos
            txtAux(10).Text = DataGridAux(Index).Columns(0).Text
            txtAux(11).Text = DataGridAux(Index).Columns(1).Text
            txtAux(9).Text = DataGridAux(Index).Columns(2).Text
            
            Text2(7).Text = DataGridAux(Index).Columns(3).Text
            
            txtAux(8).Text = DataGridAux(Index).Columns(4).Text
    End Select
    
    LLamaLineas Index, ModoLineas, anc
   
    ' *** foco al 1r camp visible de les llinies en grids que no siga PK (en o sense tab) ***
    Select Case Index
        Case 1 'incidencias
            PonerFoco txtAux(1)
        Case 2 'gastos
            PonerFoco txtAux(9)
    End Select
    ' ***************************************************************************************
End Sub

Private Sub LLamaLineas(Index As Integer, xModo As Byte, Optional alto As Single)
Dim jj As Integer
Dim B As Boolean

    ' *** si n'hi han tabs sense datagrid posar el If ***
    DeseleccionaGrid DataGridAux(Index)
       
    B = (xModo = 1 Or xModo = 2) 'Insertar o Modificar Llínies
    Select Case Index
        Case 1 'incidencias
            For jj = 1 To 1
                txtAux(jj).visible = B
                txtAux(jj).Top = alto
            Next jj
            txtAux2(1).visible = B
            txtAux2(1).Top = alto
            btnBuscar(1).visible = B
            btnBuscar(1).Top = alto
        Case 2 ' gastos
            For jj = 8 To 9
                txtAux(jj).visible = B
                txtAux(jj).Top = alto
            Next jj
            Text2(7).visible = B
            Text2(7).Top = alto
            btnBuscar(0).visible = B
            btnBuscar(0).Top = alto - 5
    End Select
End Sub

' ********* si n'hi han combos a la capçalera ************
Private Sub CargaCombo()
Dim Ini As Integer
Dim Fin As Integer
Dim I As Integer

    ' *** neteje els combos, els pose valor i seleccione el valor per defecte ***
    For I = 0 To Combo1.Count - 1
        Combo1(I).Clear
    Next I
    
    'tipo de entrada
    Combo1(0).AddItem "Normal"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 0
    Combo1(0).AddItem "V.Campo"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 1
    Combo1(0).AddItem "P.Integrado"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 2
    
    'recolectado por
    Combo1(1).AddItem "Cooperativa"
    Combo1(1).ItemData(Combo1(1).NewIndex) = 0
    Combo1(1).AddItem "Socio"
    Combo1(1).ItemData(Combo1(1).NewIndex) = 1
    
    
End Sub

Private Sub Toolbar5_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1 ' Cambio de Rendimiento
            mnCambiarRdto_Click
        Case 2 ' Asignacion de precios masiva
            mnAsignacion_Click
    End Select
End Sub

Private Sub ToolbarDes_ButtonClick(ByVal Button As MSComctlLib.Button)
    Desplazamiento (Button.Index)
End Sub

Private Sub txtAux_LostFocus(Index As Integer)
Dim cadMen As String
Dim Nuevo As Boolean
    
    If Not PerderFocoGnral(txtAux(Index), Modo) Then Exit Sub
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    

    ' ******* configurar el LostFocus dels camps de llínies (dins i fora grid) ********
    Select Case Index
        Case 1 ' codigo de incidencia
            If PonerFormatoEntero(txtAux(Index)) Then
                txtAux2(Index).Text = PonerNombreDeCod(txtAux(Index), "rincidencia", "nomincid")
                If txtAux2(Index).Text = "" Then
                    cadMen = "No existe el Código de Incidencia: " & txtAux(Index).Text & vbCrLf
                    cadMen = cadMen & "¿Desea crearla?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmInc = New frmManInciden
                        frmInc.DatosADevolverBusqueda = "0|1|"
                        frmInc.NuevoCodigo = txtAux(Index).Text
                        txtAux(Index).Text = ""
                        TerminaBloquear
                        frmInc.Show vbModal
                        Set frmInc = Nothing
                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
                    Else
                        txtAux(Index).Text = ""
                    End If
                    PonerFoco txtAux(Index)
                Else
                    cmdAceptar.SetFocus
                End If
            Else
                txtAux2(Index).Text = ""
            End If
            
        Case 9 ' nombre de gastos
            If txtAux(Index) <> "" Then
                Text2(7) = DevuelveDesdeBDNew(cAgro, "rconcepgasto", "nomgasto", "codgasto", txtAux(9), "N")
                If Text2(7).Text = "" Then
                    cadMen = "No existe el Concepto de Gasto: " & txtAux(Index).Text & vbCrLf
                    cadMen = cadMen & "¿Desea crearla?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmGas = New frmManConcepGasto
                        frmGas.DatosADevolverBusqueda = "0|1|"
                        frmGas.NuevoCodigo = txtAux(9).Text
                        TerminaBloquear
                        frmGas.Show vbModal
                        Set frmGas = Nothing
                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
                        PonerFoco txtAux(Index)
                    Else
                        txtAux(Index).Text = ""
                    End If
                    PonerFoco txtAux(Index)
                Else
                    If EsGastodeFactura(txtAux(9).Text) = True Then
                        MsgBox "Este concepto de gasto es de factura. Reintroduzca.", vbExclamation
                        PonerFoco txtAux(Index)
                    End If
                End If
            Else
                Text2(7).Text = ""
            End If
    
        Case 8 ' importe
            If txtAux(Index) <> "" Then
                If PonerFormatoDecimal(txtAux(Index), 3) Then cmdAceptar.SetFocus
            End If
        
            
            
    End Select
End Sub

Private Sub txtaux_GotFocus(Index As Integer)
   If Not txtAux(Index).MultiLine Then ConseguirFocoLin txtAux(Index)
End Sub


Private Sub txtAux_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Not txtAux(Index).MultiLine Then KEYdown KeyCode
End Sub

Private Sub txtaux_KeyPress(Index As Integer, KeyAscii As Integer)
    If Not txtAux(Index).MultiLine Then
        If KeyAscii = teclaBuscar Then
            If Modo = 5 And (ModoLineas = 1 Or ModoLineas = 2) Then
                Select Case Index
                    Case 9: KEYBusquedaAux KeyAscii, 0 'concepto de gasto
                    Case 1: KEYBusquedaAux KeyAscii, 1 'incidencia
                End Select
            End If
        Else
            KEYpress KeyAscii
        End If
    End If
End Sub

Private Function DatosOkLlin(nomframe As String) As Boolean
Dim Rs As ADODB.Recordset
Dim SQL As String
Dim B As Boolean
Dim Cant As Integer
Dim Mens As String
Dim vFact As Byte, vDocum As Byte

    On Error GoTo EDatosOKLlin

    Mens = ""
    DatosOkLlin = False
        
    B = CompForm2(Me, 2, nomframe) 'Comprovar formato datos ok
    If Not B Then Exit Function
    
    ' ******************************************************************************
    DatosOkLlin = B
    
EDatosOKLlin:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function

Private Function SepuedeBorrar(ByRef Index As Integer) As Boolean
    SepuedeBorrar = False
    
    ' *** si cal comprovar alguna cosa abans de borrar ***
'    Select Case Index
'        Case 0 'cuentas bancarias
'            If AdoAux(Index).Recordset!ctaprpal = 1 Then
'                MsgBox "No puede borrar una Cuenta Principal. Seleccione antes otra cuenta como Principal", vbExclamation
'                Exit Function
'            End If
'    End Select
    ' ****************************************************
    
    SepuedeBorrar = True
End Function

Private Sub imgBuscar_Click(Index As Integer)
    TerminaBloquear
    Indice = Index + 3
     Select Case Index
        Case 0 'variedades
            Set frmVar = New frmComVar
            frmVar.DatosADevolverBusqueda = "0|1|"
            frmVar.CodigoActual = Text1(3).Text
            frmVar.Show vbModal
            Set frmVar = Nothing
            PonerFoco Text1(3)
        Case 1 'socios
            Set frmSoc = New frmManSocios
            frmSoc.DatosADevolverBusqueda = "0|1|"
'            frmSoc.CodigoActual = Text1(4).Text
            frmSoc.Show vbModal
            Set frmSoc = Nothing
            PonerFoco Text1(4)
        Case 2 'campos
            Set frmCam = New frmManCampos
            frmCam.DatosADevolverBusqueda = "0|"
'            frmCam.CodigoActual = Text1(5).Text
            frmCam.Show vbModal
            Set frmCam = Nothing
            PonerFoco Text1(5)
        Case 3 'Poblacion
            Set frmPob = New frmManPueblos
            frmPob.Caption = "Pueblos"
            frmPob.DatosADevolverBusqueda = "0|1|"
            frmPob.CodigoActual = Text1(6).Text
            frmPob.Show vbModal
            Set frmPob = Nothing
            PonerFoco Text1(6)
            
        Case 4 '[Monica]09/11/2015: ayuda de campos de un socio solo para ABN
            Dim cad As String
            Dim Cad1 As String
            Dim NumRegis As Long

            If Text1(4).Text = "" Then Exit Sub

            cad = "rcampos.codsocio = " & DBSet(Text1(4).Text, "N") & " and rcampos.fecbajas is null"
            
            Cad1 = "select count(*) from rcampos where " & cad
             
            NumRegis = TotalRegistros(Cad1)
            If NumRegis = 0 Then
            
            Else
                Set frmMens = New frmMensajes
                frmMens.cadWHERE = " and " & cad '"rcampos.codsocio = " & NumCod & " and rcampos.fecbajas is null"
                frmMens.campo = Text1(5).Text
                frmMens.OpcionMensaje = 6
                frmMens.Show vbModal
                Set frmMens = Nothing
            End If
    
        
    End Select
    
    If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
End Sub

Private Sub DataGridAux_RowColChange(Index As Integer, LastRow As Variant, ByVal LastCol As Integer)
Dim I As Byte

    If ModoLineas <> 1 Then
        Select Case Index
            Case 0 'cuentas bancarias
                If DataGridAux(Index).Columns.Count > 2 Then
'                    txtAux(11).Text = DataGridAux(Index).Columns("direccio").Text
'                    txtAux(12).Text = DataGridAux(Index).Columns("observac").Text
                End If
                
            Case 1 'departamentos
                If DataGridAux(Index).Columns.Count > 2 Then
'                    txtAux(21).Text = DataGridAux(Index).Columns(5).Text
'                    txtAux(22).Text = DataGridAux(Index).Columns(6).Text
'                    txtAux(23).Text = DataGridAux(Index).Columns(8).Text
'                    txtAux(24).Text = DataGridAux(Index).Columns(15).Text
'                    txtAux2(22).Text = DataGridAux(Index).Columns(7).Text
                End If
                
        End Select
        
    Else 'vamos a Insertar
        Select Case Index
            Case 0 'cuentas bancarias
'                txtAux(11).Text = ""
'                txtAux(12).Text = ""
            Case 1 'departamentos
                For I = 21 To 24
'                   txtAux(i).Text = ""
                Next I
'               txtAux2(22).Text = ""
            Case 2 'Tarjetas
'               txtAux(50).Text = ""
'               txtAux(51).Text = ""
        End Select
    End If
End Sub

' ***** si n'hi han varios nivells de tabs *****
'Private Sub SituarTab(numTab As Integer)
'    On Error Resume Next
'
'    SSTab1.Tab = numTab
'
'    If Err.Number <> 0 Then Err.Clear
'End Sub
' **********************************************

Private Sub CargaFrame(Index As Integer, enlaza As Boolean)
Dim tip As Integer
Dim I As Byte

    Adoaux(Index).ConnectionString = conn
    Adoaux(Index).RecordSource = MontaSQLCarga(Index, enlaza)
    Adoaux(Index).CursorType = adOpenDynamic
    Adoaux(Index).LockType = adLockPessimistic
    Adoaux(Index).Refresh
    
    If Not Adoaux(Index).Recordset.EOF Then
        PonerCamposForma2 Me, Adoaux(Index), 2, "FrameAux" & Index
    Else
        ' *** si n'hi han tabs sense datagrids, li pose els valors als camps ***
        NetejaFrameAux "FrameAux3" 'neteja només lo que te TAG
    End If
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

Private Sub CargaGrid(Index As Integer, enlaza As Boolean)
Dim B As Boolean
Dim I As Byte
Dim tots As String

    On Error GoTo ECarga

    tots = MontaSQLCarga(Index, enlaza)

    CargaGridGnral Me.DataGridAux(Index), Me.Adoaux(Index), tots, PrimeraVez
    
    Select Case Index
        Case 1 'incidencias
            'si es visible|control|tipo campo|nombre campo|ancho control|
            tots = "N||||0|;S|txtAux(1)|T|Incidencia|1500|;S|btnBuscar(1)|B||195|;"
            tots = tots & "S|txtAux2(1)|T|Denominación|6600|;"

            arregla tots, DataGridAux(Index), Me, 350
            
            B = (Modo = 4) And ((ModoLineas = 1) Or (ModoLineas = 2))
        
        Case 2 'rhisfruta_gastos
'       SQL = SELECT rhisfruta_gastos.numalbar, rhisfruta_gastos.numlinea, rhisfruta_gastos.codgasto, rconcepgasto.nomgasto, rhisfruta_gastos.importe "
            tots = "N||||0|;N||||0|;"
            tots = tots & "S|txtAux(9)|T|Código|1200|;S|btnBuscar(0)|B|||;"
            tots = tots & "S|Text2(7)|T|Descripción|4800|;"
            tots = tots & "S|txtAux(8)|T|Importe|1950|;"
            
            arregla tots, DataGridAux(Index), Me, 350
            
            DataGridAux(Index).Columns(2).Alignment = dbgLeft
            DataGridAux(Index).Columns(3).Alignment = dbgLeft
            DataGridAux(Index).Columns(4).Alignment = dbgRight
            
    End Select
    
    DataGridAux(Index).ScrollBars = dbgAutomatic
    
    If Index = 2 Then CalcularTotalGastos
    ' **** si n'hi han llínies en grids i camps fora d'estos ****
'    If Not AdoAux(Index).Recordset.EOF Then
'        DataGridAux_RowColChange Index, 1, 1
'    Else
''        LimpiarCamposFrame Index
'    End If
ECarga:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid: " & DataGridAux(Index).Tag, Err.Description
End Sub

Private Sub InsertarLinea()
'Inserta registre en les taules de Llínies
Dim nomframe As String
Dim B As Boolean

    On Error Resume Next

    ' *** posa els noms del frames, tant si son de grid com si no ***
    Select Case NumTabMto
        Case 1: nomframe = "FrameAux1" 'incidencias
        Case 2: nomframe = "FrameAux2" 'gastos
    End Select
    
    If DatosOkLlin(nomframe) Then
        TerminaBloquear
        If InsertarDesdeForm2(Me, 2, nomframe) Then
            Select Case NumTabMto
                Case 1, 2
                    B = BLOQUEADesdeFormulario2(Me, Data1, 1)
                    CargaGrid NumTabMto, True
                    If B Then BotonAnyadirLinea NumTabMto
            End Select
           
        End If
    End If
End Sub

Private Function ModificarLinea() As Boolean
'Modifica registre en les taules de Llínies
Dim nomframe As String
Dim V As Integer
    
    On Error Resume Next

    ' *** posa els noms del frames, tant si son de grid com si no ***
    Select Case NumTabMto
        Case 1: nomframe = "FrameAux1" 'incidencias
        Case 2: nomframe = "FrameAux2" 'gastos
    End Select
    ModificarLinea = False
    If DatosOkLlin(nomframe) Then
        TerminaBloquear
        If ModificaDesdeFormulario2(Me, 2, nomframe) Then
            ModoLineas = 0
            Select Case NumTabMto
                Case 1
                    V = Adoaux(NumTabMto).Recordset.Fields(2) 'el 2 es el nº de llinia
                Case 2
                    V = Adoaux(NumTabMto).Recordset.Fields(2) 'el 2 es el nº de llinia
            End Select
            CargaGrid NumTabMto, True
            
            ' *** si n'hi han tabs ***
'            SituarTab (NumTabMto + 1)

            ' *** si n'hi han tabs que no tenen datagrid, posar el if ***
            PonerFocoGrid Me.DataGridAux(NumTabMto)
            Adoaux(NumTabMto).Recordset.Find (Adoaux(NumTabMto).Recordset.Fields(3).Name & " =" & V)
            
            LLamaLineas NumTabMto, 0
            ModificarLinea = True
        End If
    End If
End Function

Private Function ObtenerWhereCab(conW As Boolean) As String
Dim vWhere As String
    
    vWhere = ""
    If conW Then vWhere = " WHERE "
    ' *** canviar-ho per la clau primaria de la capçalera ***
    vWhere = vWhere & " numalbar=" & Me.Data1.Recordset!numalbar
    
    ObtenerWhereCab = vWhere
End Function

'' *** neteja els camps dels tabs de grid que
''estan fora d'este, i els camps de descripció ***
Private Sub LimpiarCamposFrame(Index As Integer)
    On Error Resume Next
 
'    Select Case Index
'        Case 0 'Cuentas Bancarias
'            txtAux(11).Text = ""
'            txtAux(12).Text = ""
'        Case 1 'Departamentos
'            txtAux(21).Text = ""
'            txtAux(22).Text = ""
'            txtAux2(22).Text = ""
'            txtAux(23).Text = ""
'            txtAux(24).Text = ""
'        Case 2 'Tarjetas
'            txtAux(50).Text = ""
'            txtAux(51).Text = ""
'        Case 4 'comisiones
'            txtAux2(2).Text = ""
'    End Select
'
    If Err.Number <> 0 Then Err.Clear
End Sub

'' ### [DavidV] 26/04/2006: Activar/desactivar la rueda del ratón.
'Private Sub DataGridAux_GotFocus(Index As Integer)
'  WheelHook DataGridAux(Index)
'End Sub
'Private Sub DataGridAux_LostFocus(Index As Integer)
'  WheelUnHook
'End Sub


Private Sub CalcularGastos()
Dim Rs As ADODB.Recordset
Dim SQL As String
Dim TotalEnvases As String
Dim TotalCostes As String
Dim Valor As Currency
Dim GasRecol As Currency
Dim GasAcarreo As Currency
Dim KilosTria As Long
Dim KilosNet As Long
Dim EurDesta As Currency
Dim EurRecol As Currency
Dim PrecAcarreo As Currency
Dim I As Integer

    On Error Resume Next
    
    GasRecol = 0
    GasAcarreo = 0
    
    If Combo1(0).ListIndex = 1 Then
        For I = 14 To 19
            Text1(I).Text = ""
        Next I
        Exit Sub
    End If
    
    
    SQL = "select eurdesta, eurecole from variedades where codvarie = " & DBSet(Text1(3).Text, "N")
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not Rs.EOF Then
        EurDesta = DBLet(Rs.Fields(0).Value, "N")
        EurRecol = DBLet(Rs.Fields(1).Value, "N")
    End If

    Set Rs = Nothing

'    Sql = "select sum(kilosnet) from rclasifica_clasif where numnotac = " & DBSet(Text1(0).Text, "N")
'    KilosNet = TotalRegistros(Sql)

    KilosNet = CLng(ImporteSinFormato(Text1(10).Text))

    'recolecta socio
    If Combo1(1).ListIndex = 1 Then
        SQL = "select sum(kilosnet) from rclasifica_clasif, rcalidad  where numnotac = " & DBSet(Text1(0).Text, "N")
        SQL = SQL & " and rclasifica_clasif.codvarie = rcalidad.codvarie "
        SQL = SQL & " and rclasifica_clasif.codcalid = rcalidad.codcalid "
        SQL = SQL & " and rcalidad.gastosrec = 1"
        
        KilosTria = TotalRegistros(SQL)
        
        GasRecol = Round2(KilosTria * EurRecol, 2)
    Else
    'recolecta cooperativa
        If Combo1(2).ListIndex = 0 Then
            'horas
            'gastosrecol = horas * personas * rparam.(costeshora + costesegso)
            GasRecol = Round2(HorasDecimal(Text1(18).Text) * CCur(Text1(19).Text) * (vParamAplic.CosteHora + vParamAplic.CosteSegSo), 2)
        Else
            'destajo
            GasRecol = Round2(KilosNet * EurDesta, 2)
        End If
    End If
    
'12/05/2009
'    If Text1(8).Text <> "" Then
'        sql = DevuelveDesdeBDNew(cAgro, "rtarifatra", "preciokg", "codtarif", Text1(8).Text, "N")
'        PrecAcarreo = CCur(sql)
'    Else
'        PrecAcarreo = 0
'    End If
'12/05/2009 cambiado por esto pq si que hay tarifa 0

    PrecAcarreo = 0
    SQL = ""
    SQL = DevuelveDesdeBDNew(cAgro, "rtarifatra", "preciokg", "codtarif", Text1(8).Text, "N")
    If SQL <> "" Then
        PrecAcarreo = CCur(SQL)
    End If
    
    GasAcarreo = Round2(PrecAcarreo * KilosNet, 2)
    
    Text1(16).Text = Format(GasRecol, "#,##0.00")
    Text1(15).Text = Format(GasAcarreo, "#,##0.00")
    

End Sub

'Private Function HorasDecimal(cantidad As String) As Currency
'Dim Entero As Long
'Dim vCantidad As String
'Dim vDecimal As String
'Dim vEntero As String
'Dim vHoras As Currency
'Dim J As Integer
'    HorasDecimal = 0
'
'    vCantidad = ImporteSinFormato(cantidad)
'
'    J = InStr(1, vCantidad, ",")
'
'    If J > 0 Then
'        vEntero = Mid(vCantidad, 1, J - 1)
'        vDecimal = Mid(vCantidad, J + 1, Len(vCantidad))
'    Else
'        vEntero = vCantidad
'        vDecimal = ""
'    End If
'
'    vHoras = (CLng(vEntero) * 60) + CLng(vDecimal)
'
'    HorasDecimal = Round2(vHoras / 60, 2)
'
'End Function


Private Sub PasarSigReg()
'Nos situamos en el siguiente registro
    If Me.DataGridAux(0).Bookmark < Me.Adoaux(0).Recordset.RecordCount Then
'        DataGridAux(0).Row = DataGridAux(0).Row + 1
        DataGridAux(0).Bookmark = DataGridAux(0).Bookmark + 1
        BotonModificarLinea 0
    ElseIf DataGridAux(0).Bookmark = Adoaux(0).Recordset.RecordCount Then
'        PonerFocoBtn Me.cmdAceptar
        BotonModificarLinea 0
    End If
End Sub


Private Sub PasarAntReg()
'Nos situamos en el siguiente registro
    If Me.DataGridAux(0).Bookmark > 1 Then
'        DataGridAux(0).Row = DataGridAux(0).Row + 1
        DataGridAux(0).Bookmark = DataGridAux(0).Bookmark - 1
        BotonModificarLinea 0
    ElseIf DataGridAux(0).Bookmark = 1 Then
'        PonerFocoBtn Me.cmdAceptar
        BotonModificarLinea 0
    End If
End Sub



Private Sub VisualizarDatosCampo(campo As String)
Dim cad As String
Dim Cad1 As String
Dim NumRegis As Long
Dim Rs As ADODB.Recordset


    If campo = "" Then Exit Sub
    
'    If Not (Modo = 3 Or Modo = 4) Then Exit Sub

    cad = "rcampos.codcampo = " & DBSet(campo, "N")
     
    Cad1 = "select rcampos.codparti, rpartida.nomparti, rpartida.codzonas, rzonas.nomzonas, "
    Cad1 = Cad1 & " rpueblos.despobla from rcampos, rpartida, rzonas, rpueblos "
    Cad1 = Cad1 & " where " & cad
    Cad1 = Cad1 & " and rcampos.codparti = rpartida.codparti "
    Cad1 = Cad1 & " and rpartida.codzonas = rzonas.codzonas "
    Cad1 = Cad1 & " and rpartida.codpobla = rpueblos.codpobla "
     
    Set Rs = New ADODB.Recordset
    Rs.Open Cad1, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    Text2(0).Text = ""
    Text2(1).Text = ""
    If Not Rs.EOF Then
        Text1(5).Text = campo
        PonerFormatoEntero Text1(5)
        Text2(0).Text = DBLet(Rs.Fields(1).Value, "T") ' nombre de partida
        Text2(1).Text = DBLet(Rs!desPobla, "T")        ' nombre de la poblacion
    End If
    
    Set Rs = Nothing
    
End Sub


Private Sub PonerCamposSocioVariedad()
Dim cad As String
Dim Cad1 As String
Dim NumRegis As Long
Dim Rs As ADODB.Recordset


    If Text1(3).Text = "" Or Text1(4).Text = "" Then Exit Sub
    
    If Not (Modo = 3 Or Modo = 4) Then Exit Sub

    cad = "rcampos.codsocio = " & DBSet(Text1(4).Text, "N") & " and rcampos.fecbajas is null"
    
    '[Monica]09/11/2015: solo para el caso de abn no miramos que los campos sean de la variedad
    If vParamAplic.Cooperativa <> 1 Then
        cad = cad & " and rcampos.codvarie = " & DBSet(Text1(3).Text, "N")
    End If
     
    Cad1 = "select count(*) from rcampos where " & cad
     
    NumRegis = TotalRegistros(Cad1)
    SinCampos = False
    If NumRegis = 0 Then
        '[Monica]15/01/2015: Si no hay campos dados de alta damos un error
        If vParamAplic.Cooperativa = 5 Then SinCampos = True
        Exit Sub
    End If
    If NumRegis = 1 Then
        Cad1 = "select codcampo from rcampos where " & cad
        Set Rs = New ADODB.Recordset
        Rs.Open Cad1, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not Rs.EOF Then
            Text1(5).Text = DBLet(Rs.Fields(0).Value)
            PonerDatosCampo Text1(5).Text
        End If
    Else
        Set frmMens = New frmMensajes
        frmMens.cadWHERE = " and " & cad '"rcampos.codsocio = " & NumCod & " and rcampos.fecbajas is null"
        frmMens.campo = Text1(5).Text
        frmMens.OpcionMensaje = 6
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
    
    Text1(5).Text = ""
    Text2(0).Text = ""
    Text2(1).Text = ""
    If Not Rs.EOF Then
        Text1(5).Text = campo
        PonerFormatoEntero Text1(5)
        Text2(0).Text = DBLet(Rs.Fields(1).Value, "T") ' nombre de partida
        Text2(1).Text = DBLet(Rs.Fields(4).Value, "T") ' descripcion de poblacion
    End If
    
    Set Rs = Nothing
    
End Sub

Private Sub InsertarCabecera()
Dim SQL As String
Dim actualiza As Boolean
        
    SQL = CadenaInsertarDesdeForm(Me)
    If InsertarOferta(SQL) Then
        CadenaConsulta = "Select * from " & NombreTabla & ObtenerWhereCP(True) & Ordenacion
        PonerCadenaBusqueda
        PonerModo 2
    End If
    
    Text1(0).Text = Format(Text1(0).Text, "0000000")
End Sub


Private Function InsertarOferta(vSQL As String) As Boolean
Dim MenError As String
Dim bol As Boolean, Existe As Boolean
Dim cambiaSQL As Boolean
Dim devuelve As String
Dim Sql2 As String

Dim Rs As ADODB.Recordset
Dim Sql3 As String
Dim cadMen As String

    On Error GoTo EInsertarOferta
    
    bol = True
    
    cambiaSQL = False
    If cambiaSQL Then vSQL = CadenaInsertarDesdeForm(Me)
    
    
    'Aqui empieza transaccion
    conn.BeginTrans
    MenError = "Error al insertar en la tabla Cabecera de Entradas (" & NombreTabla & ")."
    conn.Execute vSQL, , adCmdText
    
    cadMen = ""
    Sql3 = "select * from rhisfruta where numalbar = " & DBSet(Text1(0).Text, "N")
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql3, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not Rs.EOF Then
        bol = InsertarLineaClasificacion(Rs, cadMen, "")
        cadMen = "Insertando Línea de Clasificacion: " & cadMen
    End If
    
    Set Rs = Nothing
    
    MenError = MenError & cadMen
    
EInsertarOferta:
    If Err.Number <> 0 Or Not bol Then
        MenError = "Insertando Entrada." & vbCrLf & "----------------------------" & vbCrLf & MenError
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

Private Function ObtenerWhereCP(conWhere As Boolean) As String
Dim SQL As String

    On Error Resume Next
    
    SQL = " numalbar= " & Text1(0).Text
    If conWhere Then SQL = " WHERE " & SQL
    ObtenerWhereCP = SQL
    
    If Err.Number <> 0 Then MuestraError Err.Number, "Obteniendo cadena WHERE.", Err.Description
End Function


Private Function InsertarLineaClasificacion(ByRef Rs As ADODB.Recordset, cadErr As String, vCalidad As String) As Boolean
'Insertando en tabla conta.cabfact
Dim SQL As String
Dim Sql1 As String
Dim RS1 As ADODB.Recordset
Dim cad As String
Dim KilosMuestra As Currency
Dim TotalKilos As Currency
Dim Calidad As Currency
Dim Diferencia As Currency
Dim HayReg As Byte
Dim TipoClasif As Byte
Dim vTipoClasif As String
Dim vCalidDest As String
Dim CalidadClasif As String
Dim CalidadVC As String
Dim Hora As String

    On Error GoTo EInsertar
    
    Hora = DBLet(Rs!Fecalbar, "F") & " " & Format(Now, "HH:MM:SS")
    
    SQL = "insert into rhisfruta_entradas (numalbar, numnotac, fechaent, horaentr, kilosbru, "
    SQL = SQL & " numcajon, kilosnet, observac, prestimado) values ("
    SQL = SQL & DBSet(Rs!numalbar, "N") & ","
    SQL = SQL & "0,"  ' numero de nota --> no viene de ninguna nota.
    SQL = SQL & DBSet(Rs!Fecalbar, "F") & ","
    SQL = SQL & DBSet(Hora, "FH") & ","
    SQL = SQL & DBSet(Rs!KilosBru, "N") & ","
    SQL = SQL & "0," ' numero de cajones
    SQL = SQL & DBSet(Rs!KilosNet, "N") & ","
    SQL = SQL & DBSet(Text1(8).Text, "T") & ","
    SQL = SQL & DBSet(Rs!PrEstimado, "N") & ")"
    
    conn.Execute SQL
    
    
EInsertar:
    If Err.Number <> 0 Then
        InsertarLineaClasificacion = False
        cadErr = Err.Description
    Else
        InsertarLineaClasificacion = True
    End If
End Function


Private Sub ActualizoObservaciones()
Dim SQL As String

    SQL = "update rhisfruta_entradas set observac = " & DBSet(Text1(8).Text, "T")
    SQL = SQL & " where numalbar = " & DBSet(Text1(0).Text, "N")
    
    conn.Execute SQL
    
End Sub


Private Sub CalcularKilogrado()
    
    If Text1(2).Text <> "" And Text1(7).Text <> "" Then
        Text2(2).Text = Round2(ImporteSinFormato(Text1(2).Text) * ImporteSinFormato(Text1(7).Text) / 100, 0)
        Text2(2).Text = Format(Text2(2).Text, "###,###,##0")
    End If

End Sub

Private Sub CalcularTotalGastos()
Dim Gastos As Double

    If Data1.Recordset.EOF Then Exit Sub

    Gastos = DevuelveValor("select sum(importe) from rhisfruta_gastos where numalbar = " & Data1.Recordset.Fields(0))
    Text2(5).Text = Format(Gastos, "###,###,##0.00")
    
End Sub


Private Sub ToolbarAyuda_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1
            LanzaVisorMimeDocumento Me.hWnd, DireccionAyuda & IdPrograma & ".html"
    End Select
End Sub
