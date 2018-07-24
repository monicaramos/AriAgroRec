VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmManPartes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Partes de Campo"
   ClientHeight    =   8895
   ClientLeft      =   45
   ClientTop       =   4035
   ClientWidth     =   14670
   Icon            =   "frmManPartes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8895
   ScaleWidth      =   14670
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameBotonGnral 
      Height          =   705
      Left            =   90
      TabIndex        =   71
      Top             =   90
      Width           =   3090
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   210
         TabIndex        =   72
         Top             =   180
         Width           =   2685
         _ExtentX        =   4736
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
               Enabled         =   0   'False
               Object.Visible         =   0   'False
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
               Object.ToolTipText     =   "Informe dias trabajados"
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
      Left            =   3240
      TabIndex        =   69
      Top             =   90
      Width           =   1335
      Begin MSComctlLib.Toolbar Toolbar5 
         Height          =   330
         Left            =   210
         TabIndex        =   70
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
               Object.ToolTipText     =   "Traer Entradas"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Recalcular importes"
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame FrameDesplazamiento 
      Height          =   705
      Left            =   4620
      TabIndex        =   67
      Top             =   90
      Width           =   2415
      Begin MSComctlLib.Toolbar ToolbarDes 
         Height          =   330
         Left            =   240
         TabIndex        =   68
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
      Left            =   11880
      TabIndex        =   66
      Top             =   315
      Width           =   1605
   End
   Begin VB.Frame Frame2 
      Height          =   1065
      Left            =   90
      TabIndex        =   11
      Top             =   840
      Width           =   14425
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
         Left            =   10665
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Tag             =   "Tipo Parte|N|N|||rpartes|tipoparte|||"
         Top             =   450
         Width           =   1485
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
         Left            =   8805
         MaxLength       =   10
         TabIndex        =   3
         Tag             =   "Fecha Entrada|F|N|||rpartes|fecentrada|dd/mm/yyyy||"
         Top             =   450
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
         Index           =   2
         Left            =   2610
         MaxLength       =   6
         TabIndex        =   2
         Tag             =   "Cod.Cuadrilla|N|N|0|999999|rpartes|codcuadrilla|000000||"
         Text            =   "Text1"
         Top             =   450
         Width           =   1050
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
         Left            =   225
         MaxLength       =   7
         TabIndex        =   0
         Tag             =   "Nº Parte|N|S|||rpartes|nroparte|0000000|S|"
         Text            =   "nropart"
         Top             =   450
         Width           =   975
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Impreso"
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
         Left            =   12255
         TabIndex        =   5
         Top             =   480
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
         Index           =   1
         Left            =   1260
         MaxLength       =   10
         TabIndex        =   1
         Tag             =   "Fecha Parte|F|N|||rpartes|fechapar|dd/mm/yyyy||"
         Top             =   450
         Width           =   1260
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
         Index           =   2
         Left            =   3690
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   14
         Text            =   "Text2"
         Top             =   450
         Width           =   4695
      End
      Begin VB.Label Label21 
         Caption         =   "Tipo Parte"
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
         Left            =   10665
         TabIndex        =   56
         Top             =   165
         Width           =   1455
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   1
         Left            =   10065
         Picture         =   "frmManPartes.frx":000C
         ToolTipText     =   "Buscar fecha"
         Top             =   180
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Fec.Entrada"
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
         Left            =   8805
         TabIndex        =   52
         Top             =   180
         Width           =   1230
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
         Left            =   1305
         TabIndex        =   15
         Top             =   180
         Width           =   690
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   0
         Left            =   2250
         Picture         =   "frmManPartes.frx":0097
         ToolTipText     =   "Buscar fecha"
         Top             =   180
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
         Index           =   0
         Left            =   2610
         TabIndex        =   13
         Top             =   225
         Width           =   840
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   3510
         ToolTipText     =   "Buscar Variedad"
         Top             =   180
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Nº Parte"
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
         TabIndex        =   12
         Top             =   180
         Width           =   855
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6210
      Left            =   60
      TabIndex        =   16
      Top             =   1980
      Width           =   14535
      _ExtentX        =   25638
      _ExtentY        =   10954
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   4
      TabHeight       =   520
      ForeColor       =   9771019
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Reparto Trabajadores"
      TabPicture(0)   =   "frmManPartes.frx":0122
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame5"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame4"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "FrameAux0"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      Begin VB.Frame FrameAux0 
         Caption         =   "Gastos Generales Parte"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0095180B&
         Height          =   2580
         Left            =   7290
         TabIndex        =   33
         Top             =   3555
         Width           =   7100
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
            Height          =   315
            Index           =   1
            Left            =   1080
            MaxLength       =   6
            TabIndex        =   39
            Tag             =   "linea|N|N|0|999|rpartes_gastos|numlinea|000|S|"
            Text            =   "linea"
            Top             =   1665
            Visible         =   0   'False
            Width           =   705
         End
         Begin VB.TextBox txtAux2 
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
            Left            =   180
            MaxLength       =   7
            TabIndex        =   38
            Tag             =   "Num.Parte|N|N|||rpartes_gastos|nroparte||S|"
            Text            =   "numpart"
            Top             =   1665
            Visible         =   0   'False
            Width           =   855
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
            Index           =   6
            Left            =   2790
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   37
            Text            =   "Nombre Gasto"
            Top             =   1665
            Visible         =   0   'False
            Width           =   1740
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
            Left            =   2565
            MaskColor       =   &H00000000&
            TabIndex        =   36
            ToolTipText     =   "Buscar Gasto Nómina/Campo"
            Top             =   1665
            Visible         =   0   'False
            Width           =   195
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
            Height          =   315
            Index           =   2
            Left            =   1845
            MaxLength       =   2
            TabIndex        =   34
            Tag             =   "Gasto|N|N|0|99|rpartes_gastos|codgasto|00|S|"
            Text            =   "Ga"
            Top             =   1665
            Visible         =   0   'False
            Width           =   705
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
            Height          =   315
            Index           =   3
            Left            =   4545
            MaxLength       =   7
            TabIndex        =   35
            Tag             =   "Importe|N|S|||rpartes_gastos|importe|###,##0.00|N|"
            Text            =   "Importe"
            Top             =   1665
            Visible         =   0   'False
            Width           =   705
         End
         Begin MSComctlLib.Toolbar ToolAux 
            Height          =   390
            Index           =   1
            Left            =   90
            TabIndex        =   40
            Top             =   315
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
            Bindings        =   "frmManPartes.frx":013E
            Height          =   1680
            Left            =   90
            TabIndex        =   41
            Top             =   720
            Width           =   6780
            _ExtentX        =   11959
            _ExtentY        =   2963
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
            Left            =   1485
            Top             =   360
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
         Height          =   3075
         Left            =   90
         TabIndex        =   22
         Top             =   330
         Width           =   14265
         Begin VB.TextBox txtAux3 
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
            Left            =   5820
            MaxLength       =   7
            TabIndex        =   58
            Tag             =   "Cajas Recol|N|S|||rpartes_trabajador|numcajas|###,##0|N|"
            Text            =   "Cajas"
            Top             =   1170
            Visible         =   0   'False
            Width           =   540
         End
         Begin VB.TextBox txtAux3 
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
            Left            =   7200
            MaxLength       =   7
            TabIndex        =   27
            Tag             =   "Horasl|N|S|||rpartes_trabajador|horastra|#,##0.00|N|"
            Text            =   "Horas"
            Top             =   1170
            Visible         =   0   'False
            Width           =   540
         End
         Begin VB.TextBox txtAux3 
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
            Left            =   9000
            MaxLength       =   10
            TabIndex        =   55
            Tag             =   "Modificado|N|N|0|1|rpartes_trabajador|modificado|0||"
            Text            =   "Modificado"
            Top             =   1170
            Visible         =   0   'False
            Width           =   945
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
            Index           =   1
            Left            =   12150
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   53
            Text            =   "Text2"
            Top             =   180
            Width           =   1680
         End
         Begin VB.TextBox txtAux3 
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
            Left            =   855
            MaxLength       =   7
            TabIndex        =   47
            Tag             =   "Num.Linea|N|N|||rpartes_trabajador|numlinea|0000000|S|"
            Text            =   "lin"
            Top             =   1170
            Visible         =   0   'False
            Width           =   345
         End
         Begin VB.TextBox txtAux3 
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
            Left            =   4050
            MaxLength       =   6
            TabIndex        =   46
            Tag             =   "Variedad|N|S|||rpartes_trabajador|codvarie|000000|N|"
            Text            =   "Varied"
            Top             =   1170
            Visible         =   0   'False
            Width           =   630
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
            Left            =   4950
            MaxLength       =   30
            TabIndex        =   45
            Text            =   "Nom Varie"
            Top             =   1170
            Visible         =   0   'False
            Width           =   630
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
            Index           =   3
            Left            =   4725
            MaskColor       =   &H00000000&
            TabIndex        =   44
            ToolTipText     =   "Buscar Variedad"
            Top             =   1170
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
            Height          =   300
            Index           =   2
            Left            =   3105
            MaskColor       =   &H00000000&
            TabIndex        =   43
            ToolTipText     =   "Buscar Gasto Nómina/Campo"
            Top             =   1170
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
            Height          =   300
            Index           =   1
            Left            =   1755
            MaskColor       =   &H00000000&
            TabIndex        =   42
            ToolTipText     =   "Buscar Trabajador"
            Top             =   1170
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.TextBox txtAux3 
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
            Left            =   7935
            MaxLength       =   10
            TabIndex        =   28
            Tag             =   "Importe|N|N|||rpartes_trabajador|importe|###,##0.00||"
            Text            =   "Importe"
            Top             =   1170
            Visible         =   0   'False
            Width           =   945
         End
         Begin VB.TextBox txtAux3 
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
            Left            =   6525
            MaxLength       =   7
            TabIndex        =   26
            Tag             =   "Kilos Recol|N|S|||rpartes_trabajador|kilosrec|###,##0|N|"
            Text            =   "Kilos"
            Top             =   1170
            Visible         =   0   'False
            Width           =   540
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
            Index           =   4
            Left            =   3330
            MaxLength       =   30
            TabIndex        =   30
            Text            =   "Nom Gas"
            Top             =   1170
            Visible         =   0   'False
            Width           =   630
         End
         Begin VB.TextBox txtAux3 
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
            Left            =   2430
            MaxLength       =   30
            TabIndex        =   25
            Tag             =   "Cod.Gasto|N|S|||rpartes_trabajador|codgasto|00|N|"
            Text            =   "Gast"
            Top             =   1170
            Visible         =   0   'False
            Width           =   630
         End
         Begin VB.TextBox txtAux3 
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
            Left            =   180
            MaxLength       =   7
            TabIndex        =   29
            Tag             =   "Num.Parte|N|N|||rpartes_trabajador|nroparte|0000000|S|"
            Text            =   "nropart"
            Top             =   1170
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.TextBox txtAux3 
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
            Left            =   1170
            MaxLength       =   7
            TabIndex        =   24
            Tag             =   "Cod.Traba|N|N|||rpartes_trabajador|codtraba|0000000|N|"
            Text            =   "Trab"
            Top             =   1170
            Visible         =   0   'False
            Width           =   540
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
            Index           =   3
            Left            =   1800
            MaxLength       =   30
            TabIndex        =   23
            Text            =   "Nomtra"
            Top             =   1170
            Visible         =   0   'False
            Width           =   585
         End
         Begin MSComctlLib.Toolbar ToolAux 
            Height          =   390
            Index           =   0
            Left            =   60
            TabIndex        =   31
            Top             =   210
            Width           =   1470
            _ExtentX        =   2593
            _ExtentY        =   688
            ButtonWidth     =   609
            ButtonHeight    =   582
            AllowCustomize  =   0   'False
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   4
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
               BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Insertar cajones"
                  Object.Tag             =   "2"
               EndProperty
            EndProperty
         End
         Begin MSDataGridLib.DataGrid DataGrid2 
            Bindings        =   "frmManPartes.frx":0153
            Height          =   2370
            Left            =   75
            TabIndex        =   32
            Top             =   675
            Width           =   14070
            _ExtentX        =   24818
            _ExtentY        =   4180
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
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "TOTAL IMPORTE "
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0095180B&
            Height          =   255
            Index           =   2
            Left            =   10290
            TabIndex        =   54
            Top             =   225
            Width           =   1665
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Reparto Kilos"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0095180B&
         Height          =   2580
         Left            =   90
         TabIndex        =   17
         Top             =   3555
         Width           =   7145
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
            Left            =   5700
            MaxLength       =   12
            TabIndex        =   57
            Tag             =   "Horas|N|N|||rpartes_variedad|horastra|#,##0.00|N|"
            Text            =   "horastra"
            Top             =   1410
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
            Index           =   2
            Left            =   2070
            MaxLength       =   7
            TabIndex        =   49
            Tag             =   "Nota Campo|N|N|||rpartes_variedad|numnotac|0000000|N|"
            Text            =   "notacam"
            Top             =   1395
            Visible         =   0   'False
            Width           =   630
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
            Left            =   1170
            MaxLength       =   7
            TabIndex        =   48
            Tag             =   "Linea|N|N|||rpartes_variedad|numlinea|0000000|S|"
            Text            =   "linea"
            Top             =   1395
            Visible         =   0   'False
            Width           =   855
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
            Index           =   0
            Left            =   270
            MaxLength       =   7
            TabIndex        =   20
            Tag             =   "Num.Parte|N|N|||rpartes_variedad|nroparte|0000000|S|"
            Text            =   "parte"
            Top             =   1395
            Visible         =   0   'False
            Width           =   855
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
            Left            =   2880
            MaxLength       =   7
            TabIndex        =   19
            Tag             =   "Variedad|N|N|||rpartes_variedad|codvarie|000000|N|"
            Text            =   "codvari"
            Top             =   1395
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
            Index           =   4
            Left            =   3825
            MaxLength       =   15
            TabIndex        =   18
            Text            =   "nomvarie"
            Top             =   1380
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
            Index           =   5
            Left            =   4860
            MaxLength       =   12
            TabIndex        =   50
            Tag             =   "Kilos Rec|N|N|||rpartes_variedad|kilosrec|##,###,##0|N|"
            Text            =   "kilosrec"
            Top             =   1395
            Visible         =   0   'False
            Width           =   735
         End
         Begin MSDataGridLib.DataGrid DataGrid1 
            Bindings        =   "frmManPartes.frx":0168
            Height          =   1680
            Left            =   90
            TabIndex        =   21
            Top             =   720
            Width           =   6915
            _ExtentX        =   12197
            _ExtentY        =   2963
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
         Begin MSComctlLib.Toolbar ToolAux 
            Height          =   390
            Index           =   2
            Left            =   90
            TabIndex        =   51
            Top             =   315
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
      End
      Begin VB.Frame Frame5 
         Caption         =   "Resumen"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0095180B&
         Height          =   2565
         Left            =   7290
         TabIndex        =   59
         Top             =   3555
         Width           =   7100
         Begin VB.TextBox txtAux4 
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
            Left            =   1350
            MaxLength       =   15
            TabIndex        =   65
            Text            =   "nomvarie"
            Top             =   1500
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.TextBox txtAux4 
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
            Left            =   420
            MaxLength       =   7
            TabIndex        =   64
            Tag             =   "Variedad|N|N|||rpartes_variedad|codvarie|000000|N|"
            Text            =   "codvari"
            Top             =   1500
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.TextBox txtAux4 
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
            Left            =   3375
            MaxLength       =   7
            TabIndex        =   63
            Tag             =   "Kilos Recol|N|S|||rpartes_trabajador|kilosrec|###,##0|N|"
            Text            =   "Kilos"
            Top             =   1500
            Visible         =   0   'False
            Width           =   540
         End
         Begin VB.TextBox txtAux4 
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
            Left            =   4785
            MaxLength       =   10
            TabIndex        =   62
            Tag             =   "Importe|N|N|||rpartes_trabajador|importe|###,##0.00||"
            Text            =   "Importe"
            Top             =   1500
            Visible         =   0   'False
            Width           =   945
         End
         Begin VB.TextBox txtAux4 
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
            Left            =   2460
            MaxLength       =   7
            TabIndex        =   61
            Tag             =   "Cajas Recol|N|S|||rpartes_trabajador|numcajas|###,##0|N|"
            Text            =   "Cajas"
            Top             =   1500
            Visible         =   0   'False
            Width           =   540
         End
         Begin MSDataGridLib.DataGrid DataGrid4 
            Bindings        =   "frmManPartes.frx":017D
            Height          =   1680
            Left            =   120
            TabIndex        =   60
            Top             =   720
            Width           =   6795
            _ExtentX        =   11986
            _ExtentY        =   2963
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
      End
   End
   Begin VB.Frame Frame1 
      Height          =   475
      Index           =   0
      Left            =   90
      TabIndex        =   9
      Top             =   8310
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
         Left            =   240
         TabIndex        =   10
         Top             =   135
         Width           =   1755
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
      Left            =   13515
      TabIndex        =   7
      Top             =   8400
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
      Left            =   12345
      TabIndex        =   6
      Top             =   8415
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
      TabIndex        =   8
      Top             =   8415
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
      Left            =   900
      Top             =   7515
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
      Left            =   720
      Top             =   7605
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
   Begin MSAdodcLib.Adodc Data4 
      Height          =   330
      Left            =   0
      Top             =   0
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
   Begin MSComctlLib.Toolbar ToolbarAyuda 
      Height          =   330
      Left            =   13920
      TabIndex        =   73
      Top             =   255
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
         Enabled         =   0   'False
         Shortcut        =   ^M
         Visible         =   0   'False
      End
      Begin VB.Menu mnEliminar 
         Caption         =   "&Eliminar"
         Shortcut        =   ^E
      End
      Begin VB.Menu mnBarra2 
         Caption         =   "-"
      End
      Begin VB.Menu mnTraerEntradas 
         Caption         =   "&Traer Entradas"
         Shortcut        =   ^T
      End
      Begin VB.Menu mnRecalcularImportes 
         Caption         =   "&Recalcular Importes"
         HelpContextID   =   2
         Shortcut        =   ^R
      End
      Begin VB.Menu mnBarra3 
         Caption         =   "-"
      End
      Begin VB.Menu mnInforme 
         Caption         =   "&Informe Días Trabajados"
         Shortcut        =   ^I
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
Attribute VB_Name = "frmManPartes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const IdPrograma = 8007



'========== VBLES PUBLICAS ====================
Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)

Public NroParte As String  ' venimos de mantenimineto de socios

'Si se llama de la busqueda en el frmAlmMovimArticulos se accede
'a las tablas del Albaran o de Facturas de movimiento seleccionado (solo consulta)
Public hcoCodMovim As String 'cod. movim
Public hcoCodTipoM As String 'Codigo detalle de Movimiento(ALC)
Public hcoFechaMov As String 'fecha del movimiento

'========== VBLES PRIVADAS ====================
Private WithEvents frmB As frmBuscaGrid 'Form para busquedas
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmZ As frmZoom  'Zoom para campos Text
Attribute frmZ.VB_VarHelpID = -1
Private WithEvents frmGas As frmManCGastosNom 'Form Mto de conceptos de gastos nomina
Attribute frmGas.VB_VarHelpID = -1
Private WithEvents frmGas1 As frmManCGastosNom 'Form Mto de conceptos de gastos nomina
Attribute frmGas1.VB_VarHelpID = -1
Private WithEvents frmMens As frmMensajes ' mensajes para sacar campos
Attribute frmMens.VB_VarHelpID = -1

Private WithEvents frmTMP As frmManPartesTMP ' temporal para introducir los trabajadores
Attribute frmTMP.VB_VarHelpID = -1

Private WithEvents frmTra As frmManTraba 'Form Mto de trabajadores
Attribute frmTra.VB_VarHelpID = -1
Private WithEvents frmCua As frmManCuadrillas 'Form Mto de cuadrillas
Attribute frmCua.VB_VarHelpID = -1

Private WithEvents frmPartesCam As frmBasico2
Attribute frmPartesCam.VB_VarHelpID = -1


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
'Para saber en MandaBusquedaPrevia si busca en la tabla scapla o en la tabla sdirec
Dim NumTabMto As Integer 'Indica quin nº de Tab està en modo Mantenimient


Dim EsDeVarios As Boolean
'Si el cliente mostrado es de Varios o No

'SQL de la tabla principal del formulario
Private CadenaConsulta As String

' indicamos que las variedades que vamos a mostrar en el formulario no sean del grupo 6
Private CadB1 As String


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
Dim Facturas As String

Dim Cliente As String
Dim cadSelect As String

Private BuscaChekc As String

Private SePuedeModificar As Boolean

Dim KilosAnt As Long
Dim ImporteAnt As Currency


Private Sub btnBuscar_Click(Index As Integer)
    TerminaBloquear
    Select Case Index
        Case 0 'Conceptos de gastos
            Set frmGas1 = New frmManCGastosNom
            frmGas1.DatosADevolverBusqueda = "0|1|"
            frmGas1.CodigoActual = txtAux2(2).Text
            frmGas1.Show vbModal
            Set frmGas1 = Nothing
            PonerFoco txtAux2(2)
        Case 1 'Trabajadores
            Set frmTra = New frmManTraba
            frmTra.DatosADevolverBusqueda = "0|1|"
'            frmTra.CodigoActual = txtAux3(1).Text
            frmTra.Show vbModal
            Set frmTra = Nothing
            PonerFoco txtAux3(1)
        Case 2 'Conceptos de gastos
            Set frmGas = New frmManCGastosNom
            frmGas.DatosADevolverBusqueda = "0|1|"
            frmGas.CodigoActual = txtAux3(3).Text
            frmGas.Show vbModal
            Set frmGas = Nothing
            PonerFoco txtAux3(3)
    End Select
    If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1

End Sub



Private Sub Check1_Click(Index As Integer)
    If Modo = 1 Then
        'Buscqueda
        If InStr(1, BuscaChekc, "check1(" & Index & ")") = 0 Then BuscaChekc = BuscaChekc & "check1(" & Index & ")|"
    End If
End Sub

Private Sub Check1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub cmdAceptar_Click()
Dim i As Integer

    Screen.MousePointer = vbHourglass
    On Error GoTo Error1

    Select Case Modo
        Case 1  'BUSQUEDA
            HacerBusqueda

        Case 3  'AÑADIR
            If DatosOk Then InsertarCabecera

        Case 4  'MODIFICAR
            If DatosOk Then
                If ModificaCabecera Then
                    espera 0.2
                    TerminaBloquear
                    PosicionarData
'                    FormatoDatosTotales
'                    i = Data3.Recordset.AbsolutePosition
                    PonerCampos
                    PonerCamposLineas
'                    SituarDataPosicion Data3, CLng(i), ""
                End If
            End If
            
         Case 5 'InsertarModificar LINEAS
            Select Case ModificaLineas
                Case 1 'afegir llínia
                    InsertarLinea NumTabMto
                Case 2 'modificar llínies
                    ModificarLinea
                    PosicionarData
            End Select
    End Select
    Screen.MousePointer = vbDefault

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
            PonerFoco Text1(0)
            
        Case 4  'Modificar
            TerminaBloquear
            PonerModo 2
            PonerCampos
            PonerFoco Text1(0)
            
        Case 5 'Lineas Detalle
            TerminaBloquear
            
            Select Case NumTabMto
                Case 0
                    If ModificaLineas = 1 Then 'INSERTAR
                        ModificaLineas = 0
                        DataGrid2.AllowAddNew = False
                        If Not Data3.Recordset.EOF Then Data3.Recordset.MoveFirst
                    End If
                    ModificaLineas = 0
                    LLamaLineas Modo, 0, "DataGrid2"
                    PonerModo 2
                    DataGrid2.Enabled = True
                    If Not Data1.Recordset.EOF Then _
                        Me.lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
            
                    'Habilitar las opciones correctas del menu segun Modo
                    PonerModoOpcionesMenu (Modo)
                    PonerOpcionesMenu 'Habilitar las opciones correctas del menu segun Nivel de Acceso
                    DataGrid2.Enabled = True
                    PonerFocoGrid DataGrid2
                
                Case 1
                    If ModificaLineas = 1 Then 'INSERTAR
                        ModificaLineas = 0
                        DataGrid3.AllowAddNew = False
                        If Not Adoaux(0).Recordset.EOF Then Adoaux(0).Recordset.MoveFirst
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
                    
                Case 2
                    If ModificaLineas = 1 Then 'INSERTAR
                        ModificaLineas = 0
                        DataGrid1.AllowAddNew = False
                        If Not Data2.Recordset.EOF Then Data2.Recordset.MoveFirst
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
End Sub

Private Sub BotonAnyadir()

    LimpiarCampos 'Huida els TextBox
    PonerModo 3
    
    ' ****** Valors per defecte a l'afegir, repasar si n'hi ha
    ' codEmpre i quins camps tenen la PK de la capçalera *******
'    Text1(0).Text = SugerirCodigoSiguienteStr("palets", "numpalet")
'    FormateaCampo Text1(0)
    
    Text1(1).Text = Format(Now, "dd/mm/yyyy")
    Text1(3).Text = Format(Now, "dd/mm/yyyy")
    
    '[Monica]13/06/2013: por defecto los partes son a destajo
    Combo1(0).ListIndex = 0
    
    LimpiarDataGrids
    
    ' el campo de total de gastos tiene que estar limpio
    
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
        
        'poner los txtaux para buscar por lineas de albaran
        anc = DataGrid2.Top
        If DataGrid2.Row < 0 Then
            anc = anc + 440
        Else
            anc = anc + DataGrid2.RowTop(DataGrid2.Row) + 20
        End If
        LLamaLineas Modo, anc, "DataGrid2"
        
        
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
        MandaBusquedaPrevia CadB1
    Else
        LimpiarCampos
        LimpiarDataGrids
        CadenaConsulta = "Select rpartes.* "
        CadenaConsulta = CadenaConsulta & "from " & NombreTabla
        If CadB1 <> "" Then CadenaConsulta = CadenaConsulta & " where " & CadB1
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

'    'solo se puede modificar la factura si no esta contabilizada
'    If FactContabilizada Then
'        TerminaBloquear
'        Exit Sub
'    End If
    
    'Añadiremos el boton de aceptar y demas objetos para insertar
    PonerModo 4
    
    PonerFoco Text1(1) '*** 1r camp visible que siga PK ***
        
End Sub


Private Sub BotonModificarLinea(Index As Integer)
'Modificar una linea
Dim vWhere As String
Dim anc As Single
Dim J As Byte

    On Error GoTo eModificarLinea

    Select Case NumTabMto
        Case 0
            If Data3.Recordset.EOF Then
                TerminaBloquear
                Exit Sub
            End If
            SePuedeModificar = SePuedeModificarLinea(CStr(Data3.Recordset.Fields(0).Value), CStr(Data3.Recordset.Fields(1).Value))
        
        Case 1
            If Adoaux(0).Recordset.EOF Then
                TerminaBloquear
                Exit Sub
            End If
            
        Case 2
            If Data2.Recordset.EOF Then
                TerminaBloquear
                Exit Sub
            End If
    End Select
       
    ModificaLineas = 2
    
    PonerModo 5, Index
 
    Select Case NumTabMto
        Case 0 ' rpartes_trabajador
            vWhere = ObtenerWhereCP(False)
            If Not BloqueaRegistro("rpartes_trabajador", vWhere) Then
                TerminaBloquear
                Exit Sub
            End If
            If DataGrid2.Bookmark < DataGrid2.FirstRow Or DataGrid2.Bookmark > (DataGrid2.FirstRow + DataGrid2.VisibleRows - 1) Then
                J = DataGrid2.Bookmark - DataGrid2.FirstRow
                DataGrid2.Scroll 0, J
                DataGrid2.Refresh
            End If
            
        '    anc = ObtenerAlto(Me.DataGrid1)
            anc = DataGrid2.Top
            If DataGrid2.Row < 0 Then
                anc = anc + 210
            Else
                anc = anc + DataGrid2.RowTop(DataGrid2.Row) + 10
            End If
        
            txtAux3(0).Text = DataGrid2.Columns(0).Text
            txtAux3(6).Text = DataGrid2.Columns(1).Text
            txtAux3(1).Text = DataGrid2.Columns(2).Text
            Text2(3).Text = DataGrid2.Columns(3).Text
            txtAux3(2).Text = DataGrid2.Columns(4).Text
            Text2(4).Text = DataGrid2.Columns(5).Text
            txtAux3(5).Text = DataGrid2.Columns(6).Text
            Text2(0).Text = DataGrid2.Columns(7).Text
            txtAux3(9).Text = DataGrid2.Columns(8).Text
            txtAux3(4).Text = DataGrid2.Columns(11).Text
            
            txtAux3(3).Text = DataGrid2.Columns(9).Text
            txtAux3(8).Text = DataGrid2.Columns(10).Text
            
            
            '[Monica]01/03/2012: me guardo los kilos iniciales por si los modifica
            KilosAnt = CLng(txtAux3(3).Text)
            ImporteAnt = CCur(ComprobarCero(txtAux3(4).Text))
            
            ModificaLineas = 2 'Modificar
            LLamaLineas ModificaLineas, anc, "DataGrid2"
            
            'Añadiremos el boton de aceptar y demas objetos para insertar
            Me.lblIndicador.Caption = "MODIFICAR"
            PonerModoOpcionesMenu (Modo)
            PonerOpcionesMenu 'Habilitar las opciones correctas del menu segun Nivel de Acceso
            DataGrid2.Enabled = True
            
'            PonerBotonCabecera False
            If txtAux3(5).Text = "" Then
                PonerFoco txtAux3(4)
            Else
                PonerFoco txtAux3(3)
            End If
            
            Me.DataGrid2.Enabled = False
        
        Case 1 ' rpartes_gastos
            vWhere = ObtenerWhereCP(False)
            If Not BloqueaRegistro("rpartes_gastos", vWhere) Then
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
        
            For J = 0 To 2
                txtAux2(J).Text = DataGrid3.Columns(J).Text
            Next J
            Text2(6).Text = DataGrid3.Columns(3).Text
            
            txtAux2(3).Text = DataGrid3.Columns(4).Text
            
            ModificaLineas = 2 'Modificar
            LLamaLineas ModificaLineas, anc, "DataGrid3"
            
            'Añadiremos el boton de aceptar y demas objetos para insertar
            Me.lblIndicador.Caption = "MODIFICAR"
            PonerModoOpcionesMenu (Modo)
            PonerOpcionesMenu 'Habilitar las opciones correctas del menu segun Nivel de Acceso
            DataGrid3.Enabled = True
            
'            PonerBotonCabecera False
            PonerFoco txtAux2(3)
            Me.DataGrid3.Enabled = False
            
        Case 2 'rpartes_variedad
            vWhere = ObtenerWhereCP(False)
            If Not BloqueaRegistro("rpartes_variedad", vWhere) Then
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
        
            For J = 0 To 5
                txtAux(J).Text = DataGrid1.Columns(J).Text
            Next J
            
            ModificaLineas = 2 'Modificar
            LLamaLineas ModificaLineas, anc, "DataGrid1"
            
            'Añadiremos el boton de aceptar y demas objetos para insertar
            Me.lblIndicador.Caption = "MODIFICAR"
            PonerModoOpcionesMenu (Modo)
            PonerOpcionesMenu 'Habilitar las opciones correctas del menu segun Nivel de Acceso
            DataGrid1.Enabled = True
            
'            PonerBotonCabecera False
            PonerFoco txtAux(5)
            Me.DataGrid1.Enabled = False
            
    End Select
    
eModificarLinea:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub LLamaLineas(xModo As Byte, Optional alto As Single, Optional grid As String)
Dim jj As Integer
Dim B As Boolean
    
    Select Case grid
        Case "DataGrid1"
            DeseleccionaGrid Me.DataGrid1
            'PonerModo xModo + 1
    
            B = (xModo = 1 Or xModo = 2) And Modo = 5  'Insertar o Modificar Lineas
    
            For jj = 2 To 5
                txtAux(jj).Height = DataGrid1.RowHeight - 10
                txtAux(jj).Top = alto + 5
                txtAux(jj).visible = B
            Next jj
            
        Case "DataGrid2"
            DeseleccionaGrid Me.DataGrid2
            B = (xModo = 1 Or xModo = 2) And Modo = 5
            For jj = 1 To 5
                txtAux3(jj).Height = DataGrid2.RowHeight
                txtAux3(jj).Top = alto + 10 '- 210 '200
                txtAux3(jj).visible = B
            Next jj
            For jj = 8 To 9
                txtAux3(jj).Height = DataGrid2.RowHeight
                txtAux3(jj).Top = alto + 10 '- 210 '200
                txtAux3(jj).visible = B
            Next jj
            For jj = 1 To 2
                btnBuscar(jj).Height = DataGrid3.RowHeight - 10
                btnBuscar(jj).Top = alto + 5
                btnBuscar(jj).visible = B
            Next jj
            For jj = 3 To 4
                Text2(jj).Height = DataGrid2.RowHeight - 10
                Text2(jj).Top = alto + 5
                Text2(jj).visible = B
            Next jj
            Text2(0).Height = DataGrid2.RowHeight - 10
            Text2(0).Top = alto + 5
            Text2(0).visible = B
            
            
        Case "DataGrid3"
            DeseleccionaGrid Me.DataGrid3
            B = (xModo = 1 Or xModo = 2) And Modo = 5
             For jj = 2 To 3
                txtAux2(jj).Height = DataGrid3.RowHeight - 10
                txtAux2(jj).Top = alto + 5
                txtAux2(jj).visible = B
            Next jj
            btnBuscar(0).Height = DataGrid3.RowHeight - 10
            btnBuscar(0).Top = alto + 5
            btnBuscar(0).visible = B
            Text2(6).Height = DataGrid3.RowHeight - 10
            Text2(6).Top = alto + 5
            Text2(6).visible = B
            
            btnBuscar(0).Top = alto + 5
            btnBuscar(0).visible = B
    
        
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
    
    cad = "Cabecera de Partes." & vbCrLf
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
        
'        'Devolvemos contador, si no estamos actualizando
'        Set vTipoMov = New CTiposMov
'        vTipoMov.DevolverContador CodTipoMov, NumPedElim
'        Set vTipoMov = Nothing
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
        DataGrid2.Enabled = True
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

Private Sub DataGrid2_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Dim i As Byte

'    If LastCol = -1 Then Exit Sub

    'Datos de la tabla albaran_calibres
    If Not Data3.Recordset.EOF Then
        'Datos de la tabla rhisfruta_incidencia
        CargaGrid DataGrid1, Data2, True
    Else
        'Datos de la tabla rhisfruta_incidencia
        CargaGrid DataGrid1, Data2, False
    End If
    
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
    
    'Viene de DblClick en frmAlmMovimArticulos y carga el form con los valores
    If NroParte <> "" And Not Data1.Recordset.EOF Then PonerCadenaBusqueda
    
'    PonerCadenaBusqueda
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub Form_Load()
Dim i As Integer
    'Icono del formulario
    Me.Icon = frmPpal.Icon

    'Icono de busqueda
    For kCampo = 0 To Me.imgBuscar.Count - 1
        Me.imgBuscar(kCampo).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next kCampo

    ' ICONITOS DE LA BARRA
    btnPrimero = 15
'    With Me.Toolbar1
'        .HotImageList = frmPpal.imgListComun_OM
'        .DisabledImageList = frmPpal.imgListComun_BN
'        .ImageList = frmPpal.imgListComun
'        .Buttons(1).Image = 1   'Botón Buscar
'        .Buttons(2).Image = 2   'Botón Todos
'        .Buttons(4).Image = 3   'Insertar
'        .Buttons(5).Image = 4   'Modificar
'        .Buttons(6).Image = 5   'Borrar
'        .Buttons(8).Image = 33  'Traer entradas
'        .Buttons(9).Image = 31  'Recalcular importes
'        .Buttons(11).Image = 10 'Informe de dias trabajados
'        .Buttons(13).Image = 11  'Salir
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
        .Buttons(1).Image = 33  'Traer entradas
        .Buttons(2).Image = 31  'Recalcular importes
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
    For kCampo = 0 To ToolAux.Count - 1
        With Me.ToolAux(kCampo)
            .HotImageList = frmPpal.imgListComun_OM16
            .DisabledImageList = frmPpal.imgListComun_BN16
            .ImageList = frmPpal.imgListComun16
            .Buttons(1).Image = 3   'Insertar
            .Buttons(2).Image = 4   'Modificar
            .Buttons(3).Image = 5   'Borrar
            
            If kCampo = 0 Then .Buttons(4).Image = 16
        End With
    Next kCampo
   ' ***********************************
   'IMAGES para zoom
    
    LimpiarCampos   'Limpia los campos TextBox
    
    CodTipoMov = "PAC" 'PArtes Campo
    VieneDeBuscar = False
    
    '[Monica]13/06/2013: partes a destajo
    CargaCombo
            
    '## A mano
    NombreTabla = "rpartes"
    NomTablaLineas = "rpartes_gastos" 'Tabla de entradas
    Ordenacion = " ORDER BY rpartes.nroparte"
    
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    
    
    CadenaConsulta = "select * from rpartes "
    If NroParte <> "" Then
        CadenaConsulta = CadenaConsulta & " where nroparte = " & DBSet(NroParte, "N")
    Else
        CadenaConsulta = CadenaConsulta & " where nroparte = -1"
    End If
    
    Data1.ConnectionString = conn
    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
        
    'Cargar el DataGrid de lineas de Revisiones inicialmente a nada DATA2
    PrimeraVez = True
    
'        CargaGrid DataGrid1, Data2, False
    'Poner los grid sin apuntar a nada
    
    '[Monica]10/10/2016: si hay resumen por variedad de cajas, no mostramos los gastos generales
    FrameAux0.visible = (vParamAplic.HayResumenCajas = 0)
    FrameAux0.Enabled = (vParamAplic.HayResumenCajas = 0)
    Frame5.visible = (vParamAplic.HayResumenCajas = 1)
    Frame5.Enabled = (vParamAplic.HayResumenCajas = 1)
    
    LimpiarDataGrids
    
    SSTab1.Tab = 0
    
    If DatosADevolverBusqueda <> "" Then
        Text1(0).Text = DatosADevolverBusqueda
        HacerBusqueda
        SSTab1.Tab = 1
    Else
        PonerModo 0
    End If
    
    ToolAux(0).Buttons(4).visible = (vParamAplic.HayResumenCajas = 1)
    ToolAux(0).Buttons(4).Enabled = (vParamAplic.HayResumenCajas = 1)
    
'    If DatosADevolverBusqueda = "" Then
'        If numalbar = "" Then
'            PonerModo 0
'        Else
'            Text1(0).Text = numalbar
'            HacerBusqueda
''            SSTab1.Tab = 1
'        End If
'    Else
'        BotonBuscar
'    End If
    
    If vParamAplic.Cooperativa = 16 Then Combo1(0).ListIndex = 0
    
    PrimeraVez = False
End Sub


Private Sub LimpiarCampos()
Dim i As Integer

    On Error Resume Next
    
    limpiar Me   'Metodo general: Limpia los controles TextBox
    lblIndicador.Caption = ""
    'Aqui va el especifico de cada form es
    '### a mano
    Me.Check1(0).Value = 0
    Me.Combo1(0).ListIndex = -1
    
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub Form_Unload(Cancel As Integer)
    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
    If Modo = 4 Or Modo = 5 Then TerminaBloquear
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
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " " & Ordenacion
        PonerCadenaBusqueda
        Screen.MousePointer = vbDefault
    End If
    Screen.MousePointer = vbDefault
End Sub



Private Sub frmC_Selec(vFecha As Date)
    ' *** repasar si el camp es txtAux o Text1 ***
    Text1(CByte(imgFec(0).Tag)).Text = Format(vFecha, "dd/mm/yyyy") '<===
    ' ********************************************
End Sub


Private Sub frmCua_DatoSeleccionado(CadenaSeleccion As String)
Dim SQL As String
    Text1(2).Text = RecuperaValor(CadenaSeleccion, 1)
    
    SQL = "select nomcapat from rcapataz where codcapat = " & DBSet(RecuperaValor(CadenaSeleccion, 2), "N")
    Text2(2).Text = DevuelveValor(SQL)

End Sub

Private Sub frmGas_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de Conceptos de gastos
    txtAux3(2).Text = Format(RecuperaValor(CadenaSeleccion, 1), "00") 'Cod concepto de gasto
    Text2(4).Text = RecuperaValor(CadenaSeleccion, 2) 'Descripcion
End Sub

Private Sub frmGas1_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de Conceptos de gastos
    txtAux2(2).Text = Format(RecuperaValor(CadenaSeleccion, 1), "00") 'Cod concepto de gasto
    Text2(6).Text = RecuperaValor(CadenaSeleccion, 2) 'Descripcion
End Sub


Private Sub frmMens_DatoSeleccionado(CadenaSeleccion As String)
    If CadenaSeleccion <> "" Then
        cadSelect = " codtraba in (" & CadenaSeleccion & ")"
    Else
        cadSelect = " codtraba = -1 "
    End If
End Sub

Private Sub frmPartesCam_DatoSeleccionado(CadenaSeleccion As String)
Dim CadB As String

    If CadenaSeleccion <> "" Then
        LimpiarCampos
    
        Text1(0).Text = RecuperaValor(CadenaSeleccion, 1) 'numero de parte
        
        CadB = ObtenerBusqueda2(Me, BuscaChekc, 1)
    
        If CadB <> "" Then
            'Se muestran en el mismo form
            CadenaConsulta = "select * from rpartes WHERE " & CadB & " " & Ordenacion
            PonerCadenaBusqueda
        End If
    End If
End Sub

Private Sub frmTMP_DatoSeleccionado(CadenaSeleccion As String)
    If CadenaSeleccion <> "" Then
    
    End If
End Sub

Private Sub frmTra_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de Trabajadores
    txtAux3(1).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000") 'Cod trabajador
    Text2(3).Text = RecuperaValor(CadenaSeleccion, 2) 'Descripcion
End Sub

Private Sub frmZ_Actualizar(vCampo As String)
     Text1(Indice).Text = vCampo
End Sub

Private Sub imgBuscar_Click(Index As Integer)
    If Modo = 2 Or Modo = 0 Then Exit Sub
    Screen.MousePointer = vbHourglass

    Select Case Index
        Case 0 'Cuadrilla
            Indice = 2
            PonerFoco Text1(Indice)
            Set frmCua = New frmManCuadrillas
            frmCua.DatosADevolverBusqueda = "0|1|"
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

    
    ' *** repasar si el camp es txtAux o Text1 ***
    Select Case Index
        Case 0
            imgFec(0).Tag = 1 '<===
            If Text1(1).Text <> "" Then frmC.NovaData = Text1(1).Text
        Case 1
            imgFec(0).Tag = 3 '<===
            If Text1(3).Text <> "" Then frmC.NovaData = Text1(3).Text
    End Select
    ' ********************************************

    frmC.Show vbModal
    Set frmC = Nothing
    ' *** repasar si el camp es txtAux o Text1 ***
    PonerFoco Text1(CByte(imgFec(0).Tag)) '<===
    ' ********************************************
End Sub

Private Sub mnBuscar_Click()
    BotonBuscar
End Sub

Private Sub mnEliminar_Click()
    If Modo = 5 Then 'Eliminar lineas de Pedido
'         BotonEliminarLinea
    Else   'Eliminar Pedido
         BotonEliminar
    End If
End Sub

Private Sub mnInforme_Click()
Dim Capataz As String
Dim SQL As String
    
    If vParamAplic.Cooperativa = 16 Then
    
        frmListNomina.OpcionListado = 40
        frmListNomina.txtCodigo(74) = Text1(0).Text
        frmListNomina.txtCodigo(75) = Text1(0).Text
        frmListNomina.txtCodigo(70) = Text1(1).Text
        frmListNomina.txtCodigo(71) = Text1(1).Text
        
        SQL = DevuelveValor("select codcapat from rcuadrilla where codcuadrilla = " & DBSet(Text1(2).Text, "N"))
        If SQL = "0" Then SQL = ""
        Capataz = SQL
        
        frmListNomina.txtCodigo(72) = Capataz
        frmListNomina.txtNombre(72) = DevuelveValor("select nomcapat from rcapataz where codcapat = " & Capataz)
        frmListNomina.txtCodigo(73) = Capataz
        frmListNomina.txtNombre(73) = frmListNomina.txtNombre(72)
        
        frmListNomina.Show vbModal
    
        'AbrirListadoNominas (40) ' impresion de parte de trabajo
    Else
        AbrirListadoNominas (39) 'Informe dias trabajados
    End If
    
End Sub

Private Sub mnRecalcularImportes_Click()
' Recalcular Importes
    If Data1.Recordset.EOF Then Exit Sub
        
    BotonRecalcularImportes

End Sub

Private Sub BotonRecalcularImportes()
Dim cad As String
Dim SQL As String
Dim Sql2 As String
Dim Sql3 As String
Dim Capataz As Long

    If RecalcularImportes Then
        CargaGrid DataGrid1, Data2, True
        CargaGrid DataGrid2, Data3, True
        If vParamAplic.HayResumenCajas = 1 Then CargaGrid DataGrid4, Data4, True
        PonerModo 2
    End If
    
End Sub


Private Sub mnTraerEntradas_Click()
    
    If Data1.Recordset.EOF Then Exit Sub
        
    BotonTraerEntradas
End Sub


Private Sub BotonTraerEntradas()
Dim cad As String
Dim SQL As String
Dim Sql2 As String
Dim Sql3 As String
Dim Capataz As Long

    If TraerEntradas Then
        CargaGrid DataGrid1, Data2, True
        CargaGrid DataGrid2, Data3, True
        
        PonerModo 2
    End If
    
End Sub


Private Function TraerEntradas() As Boolean
Dim cad As String
Dim SQL As String
Dim Sql2 As String
Dim Sql3 As String
Dim Sql4 As String
Dim Capataz As Long
Dim NroTrabajadores As Long
Dim KilosTrab As Long
Dim Precio As Currency
Dim ImporteTrab As Currency
Dim Rs2 As ADODB.Recordset
Dim Rs As ADODB.Recordset

Dim NumF As String

    On Error GoTo eBotonTraerEntradas

    TraerEntradas = False

    If EstaPartePagado(Data1.Recordset.Fields(0).Value) Then
        MsgBox "Sobre este Parte ya se ha hecho la impresión del recibo. No se permite realizar la operación.", vbExclamation
        Exit Function
    End If

    If EstaParteenHoras(Data1.Recordset.Fields(0).Value) Then
        If MsgBox("Este Parte ya se ha traspasado a horas. ¿ Desea Continuar ?", vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then
            Exit Function
        End If
    End If

    cad = "Se va a proceder a traer los kilos recolectados en la fecha de entrada."
    
    If Data2.Recordset.RecordCount <> 0 Then
        cad = cad & vbCrLf & "Perderá los registros que actualmente tiene en Reparto de Kilos."
    End If
    cad = cad & vbCrLf & "  ¿ Desea Continuar ? "
    
    If MsgBox(cad, vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
        
        conn.BeginTrans
        
        ' borramos los registros de horas que hubieran
        SQL = "delete from horas where nroparte = " & Data1.Recordset.Fields(0).Value
        
        conn.Execute SQL
        
        ' borramos los registros de rpartes_trabajador que hubieran
        ' son los que tienen codconcep = 0
        SQL = "delete from rpartes_trabajador where nroparte = " & Data1.Recordset.Fields(0).Value
        SQL = SQL & " and automatico = 1 "
        
        conn.Execute SQL
    
        cad = "select codcapat from rcuadrilla where codcuadrilla = " & Data1.Recordset.Fields(2).Value
        Capataz = DevuelveValor(cad)
        
        ' borramos anteriormente los registros de rpartes_variedad que hubieran
        SQL = "delete from rpartes_variedad where nroparte = " & Data1.Recordset.Fields(0).Value
        conn.Execute SQL
        
        ' insertamos en rpartes_variedad
        SQL = "select " & Data1.Recordset.Fields(0).Value & ",rhisfruta_entradas.numnotac, rhisfruta.codvarie, rhisfruta_entradas.horastra, "
        
        '[Monica]26/07/2010 si es picassent sacamos los kilostra
        '[Monica]19/11/2010 si es alzira tambien sacamos los kilostra
        If vParamAplic.Cooperativa = 2 Or vParamAplic.Cooperativa = 4 Or vParamAplic.Cooperativa = 16 Then
            SQL = SQL & " rhisfruta_entradas.kilostra kilosnet " ' en kilosnet tenemos los kilostra
        Else
            SQL = SQL & " rhisfruta_entradas.kilosnet kilosnet "
        End If
        '[Monica]27/12/2016: añado las cajas de todas las tablas
        SQL = SQL & ", rhisfruta_entradas.numcajon "
        
        SQL = SQL & " from rhisfruta_entradas, rhisfruta "
        SQL = SQL & " where fechaent = " & DBSet(Data1.Recordset.Fields(3), "F") ' fecentrada
        SQL = SQL & " and codcapat = " & DBSet(Capataz, "N")
        SQL = SQL & " and rhisfruta.numalbar = rhisfruta_entradas.numalbar "
        
        '[Monica]26/07/2010: añadido si es picassent tenemos que sacar los kilos sin clasificar
        If vParamAplic.Cooperativa = 2 Or vParamAplic.Cooperativa = 16 Then
            SQL = SQL & " union "
            SQL = SQL & "select " & Data1.Recordset.Fields(0).Value & ",rclasifica.numnotac, rclasifica.codvarie, rclasifica.horastra,"
            SQL = SQL & " rclasifica.kilostra kilosnet " ' los kilostra
            '[Monica]27/12/2016: añado las cajas de todas las tablas
            SQL = SQL & ", rclasifica.numcajon "
            
            SQL = SQL & " from rclasifica "
            SQL = SQL & " where fechaent = " & DBSet(Data1.Recordset.Fields(3), "F") ' fecentrada
            SQL = SQL & " and codcapat = " & DBSet(Capataz, "N")
        End If
        
        '[Monica]25/09/2012: añadido si es Catadau tenemos que sacar los kilos sin clasificar
        If vParamAplic.Cooperativa = 0 Or vParamAplic.Cooperativa = 18 Then
            SQL = SQL & " union "
            SQL = SQL & "select " & Data1.Recordset.Fields(0).Value & ",rclasifica.numnotac, rclasifica.codvarie, rclasifica.horastra, "
            SQL = SQL & " rclasifica.kilosnet kilosnet " ' los kilostra
            '[Monica]27/12/2016: añado las cajas de todas las tablas
            SQL = SQL & ", rclasifica.numcajon "
            SQL = SQL & " from rclasifica "
            SQL = SQL & " where fechaent = " & DBSet(Data1.Recordset.Fields(3), "F") ' fecentrada
            SQL = SQL & " and codcapat = " & DBSet(Capataz, "N")
        End If
        
        '[Monica]12/12/2012: añadido si es Alzira tenemos que sacar los kilos sin clasificar pero netos
        If vParamAplic.Cooperativa = 4 Then
            SQL = SQL & " union "
            SQL = SQL & "select " & Data1.Recordset.Fields(0).Value & ",rclasifica.numnotac, rclasifica.codvarie, rclasifica.horastra, "
            SQL = SQL & " rclasifica.kilosnet kilosnet " ' los kilostra
            '[Monica]27/12/2016: añado las cajas de todas las tablas
            SQL = SQL & ", rclasifica.numcajon "
            SQL = SQL & " from rclasifica "
            SQL = SQL & " where fechaent = " & DBSet(Data1.Recordset.Fields(3), "F") ' fecentrada
            SQL = SQL & " and codcapat = " & DBSet(Capataz, "N")
        End If
        
        SQL = SQL & " order by 1, 2, 3 "
        
        Set Rs = New ADODB.Recordset
        Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        cad = ""
        NumF = ""
        While Not Rs.EOF
            ' solo si no está insertada la nota de campo en otro parte la insertamos de nuevas
            SQL = "select count(*) from rpartes_variedad where nroparte <> " & Data1.Recordset.Fields(0).Value
            SQL = SQL & " and numnotac = " & DBSet(Rs!NumNotac, "N")
            
            If TotalRegistros(SQL) = 0 Then
                NumF = SugerirCodigoSiguienteStr("rpartes_variedad", "numlinea", "nroparte = " & Data1.Recordset.Fields(0))
                cad = "insert into rpartes_variedad (nroparte, numlinea, numnotac, codvarie, horastra, kilosrec, numcajon) values "
                cad = cad & "(" & Data1.Recordset.Fields(0) & "," & DBSet(NumF, "N") & "," & DBSet(Rs!NumNotac, "N") & ","
                cad = cad & DBSet(Rs!codvarie, "N") & "," & DBSet(Rs!horastra, "N") & "," & DBSet(Rs!KilosNet, "N") & "," & DBSet(Rs!Numcajon, "N") & ")"
                
                conn.Execute cad
            End If
                
            Rs.MoveNext
        Wend
        
        Set Rs = Nothing
        
        If NumF = "" Then
            MsgBox "No existen entradas para esta fecha o están en otro parte de trabajo", vbExclamation
        End If
        
        TraerEntradas = True
        conn.CommitTrans
        
        
    End If
    
    Exit Function
    
eBotonTraerEntradas:
    conn.RollbackTrans
    MuestraError Err.Number, "Traer Entradas", Err.Description
End Function

Function SecuenciaTrabajadores() As String
Dim SQL As String
Dim Rs As ADODB.Recordset

    On Error GoTo eSecuenciaTrabajadores


    SQL = "select distinct codtraba from rpartes_trabajador where nroparte = " & Data1.Recordset.Fields(0).Value
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    SQL = ""
    While Not Rs.EOF
        SQL = SQL & DBSet(Rs!CodTraba, "N") & ","
        
        Rs.MoveNext
    Wend
    Set Rs = Nothing
    
    If SQL <> "" Then SQL = Mid(SQL, 1, Len(SQL) - 1)
    
    SecuenciaTrabajadores = SQL
    Exit Function
    
eSecuenciaTrabajadores:
    MuestraError Err.Number, "Secuencia Trabajadores", Err.Description
End Function


Function RecalcularImportes() As Boolean
Dim SQL As String
Dim Sql2 As String
Dim Sql4 As String
Dim Rs As ADODB.Recordset
Dim Rs2 As ADODB.Recordset
Dim Capataz As Long
Dim NroTrabajadores As Long
Dim KilosTrab As Long
Dim Precio As Currency
Dim ImporteTrab As Currency
Dim cad As String
Dim NumF As Long
Dim PlusCapataz As Currency
        
Dim KilosRec As Long
Dim KilosInicio As Long
Dim NroTrabajadores2 As Long

Dim SqlNue As String
Dim Sql5 As String
Dim ImporteVariedad As Currency
Dim ImpTot As Currency
Dim UltLinea As Integer
        
Dim Importe As Currency
        
    On Error GoTo eRecalcularImportes
    
    
    RecalcularImportes = False
    
    If EstaPartePagado(Data1.Recordset.Fields(0).Value) Then
        MsgBox "Sobre este Parte ya se ha hecho la impresión del recibo. No se permite realizar la operación.", vbExclamation
        Exit Function
    End If
    
    If EstaParteenHoras(Data1.Recordset.Fields(0).Value) Then
        If MsgBox("Este Parte ya se ha traspasado a horas. ¿ Desea Continuar ?", vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then
            Exit Function
        End If
    End If
    
    
    '[Monica]14/06/2013: dependiendo de si el parte es por horas
    If Combo1(0).ListIndex = 1 Then
        RecalcularImportes = RecalcularImportesHoras
        Exit Function
    End If
    
    
    cad = "Se va a proceder a recalcular los importes por trabajador según los kilos recolectados y "
    cad = cad & "los gastos generales introducidos. "
    cad = cad & vbCrLf & "         ¿ Desea Continuar ? "
    If MsgBox(cad, vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
        'mostramos cuales son los trabajadores de la cuadrilla que han de seleccionar para hacer el reparto
        
'        SQL = "select codcapat from rpartes, rcuadrilla where rpartes.codcuadrilla = " & Data1.Recordset.Fields(2).Value
'        SQL = SQL & " and rpartes.codcuadrilla = rcuadrilla.codcuadrilla "
'
'        Capataz = DevuelveValor(SQL)
        
        
        If vParamAplic.Cooperativa = 16 Then
            'miramos para que lo recalcule entre lo que haya en rpartes_trabajador
            cad = "select count(*) from rpartes_trabajador where nroparte = " & Data1.Recordset.Fields(0).Value
            If TotalRegistros(cad) <> 0 Then
                vvTrabajadores = SecuenciaTrabajadores '"select distinct codtraba from rpartes_trabajador where nroparte = " & Data1.Recordset.Fields(0).Value & ""
                cadSelect = "codtraba in (" & vvTrabajadores & ")"
            Else
                vvTrabajadores = ""
            
                Set frmTMP = New frmManPartesTMP
                frmTMP.ParamVariedad = Text1(0).Text
                frmTMP.FechaParte = Text1(1).Text
                frmTMP.SoloTrabajador = 1
                frmTMP.Show vbModal
                Set frmTMP = Nothing
                
                If vvTrabajadores <> "" Then
                    cadSelect = "codtraba in (" & vvTrabajadores & ")"
                Else
                    cadSelect = "codtraba = -1"
                End If
            End If
        Else
            Set frmMens = New frmMensajes
            
            frmMens.OpcionMensaje = 22
            frmMens.campo = Data1.Recordset.Fields(2).Value
            frmMens.Show vbModal
            
            Set frmMens = Nothing
        End If
        
        ' vemos cuantos trabajadores hay en la cuadrilla para realizar los calculos
        '[Monica]30/09/2016: para el caso de coopic no tiene que ser de la cuadrilla sino los seleccionados
        If vParamAplic.Cooperativa = 16 Then
            Sql4 = "select count(*) from straba where (1=1) "
            Sql4 = Sql4 & " and " & cadSelect
            Sql4 = Sql4 & " order by 1 "
        Else
            Sql4 = "select count(*) from rcuadrilla_trabajador where codcuadrilla = " & Data1.Recordset.Fields(2)
            Sql4 = Sql4 & " and " & cadSelect
            Sql4 = Sql4 & " order by 1 "
        End If
        NroTrabajadores = TotalRegistros(Sql4)
        
        If NroTrabajadores = 0 Then
            If vParamAplic.Cooperativa = 16 Then
                MsgBox "No se han seleccionado trabajadores. No se ha ejecutado el proceso.", vbExclamation
            Else
                MsgBox "No se han seleccionado trabajadores de la cuadrilla. No se ha ejecutado el proceso.", vbExclamation
            End If
            RecalcularImportes = True
            Exit Function
        End If
        
        conn.BeginTrans
        
        ' borramos los registros de rpartes_trabajador que hubieran
        ' son los que tienen codconcep = 0
        SQL = "delete from rpartes_trabajador where nroparte = " & Data1.Recordset.Fields(0).Value
        SQL = SQL & " and automatico = 1 "
        '[Monica]01/03/2012: excluimos del borrado los que he modificado
        SQL = SQL & " and modificado = 0 "
        conn.Execute SQL
        
        
        ' insertamos en rpartes_trabajador: todos los trabajadores de la cuadrilla
        ' con todas las variedades de las entradas
        SQL = "select codvarie, sum(kilosrec) as kilosrec from rpartes_variedad where nroparte = " & Data1.Recordset.Fields(0).Value
        SQL = SQL & " group by 1 order by 1 "
        
        Set Rs = New ADODB.Recordset
        Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        While Not Rs.EOF
            '[Monica]01/03/2012: los kilos deben ser menos los de los registros modificados
            Sql4 = "select sum(kilosrec) from rpartes_trabajador where nroparte = " & Data1.Recordset.Fields(0).Value
            Sql4 = Sql4 & " and codvarie = " & DBSet(Rs!codvarie, "N")
            KilosInicio = DevuelveValor(Sql4)
            KilosRec = DBLet(Rs!KilosRec, "N") - KilosInicio
        
            '[Monica]01/03/2012: vemos cuantos trabajadores hemos modificado para no realizar el prorrateo
            Sql4 = "select count(*) from rpartes_trabajador where nroparte = " & Data1.Recordset.Fields(0).Value
            Sql4 = Sql4 & " and codvarie = " & DBSet(Rs!codvarie, "N")
            Sql4 = Sql4 & " and modificado = 1 "
            
            NroTrabajadores2 = NroTrabajadores - TotalRegistros(Sql4)
        
            '[Monica]29/02/2012: Para el caso de Catadau el precio es eursegsoc
            If vParamAplic.Cooperativa = 0 Or vParamAplic.Cooperativa = 18 Then
                Precio = DevuelveValor("select eursegsoc from variedades where codvarie = " & DBSet(Rs!codvarie, "N"))
            Else
                Precio = DevuelveValor("select eurdesta from variedades where codvarie = " & DBSet(Rs!codvarie, "N"))
            End If
            ' si es Picassent
            If vParamAplic.Cooperativa = 2 Or vParamAplic.Cooperativa = 0 Or vParamAplic.Cooperativa = 16 Or vParamAplic.Cooperativa = 18 Then
                KilosTrab = Round(KilosRec / NroTrabajadores2, 0)
                ImporteTrab = Round2(KilosRec / NroTrabajadores2 * Precio, 2)
            Else
            ' si es Alzira
                KilosTrab = Round(KilosRec / NroTrabajadores2, 0)
                ImporteTrab = Round2(KilosTrab * Precio, 2)
                
            End If
            
            '[Monica]30/09/2016: para el caso de coopic solo trabajadores de la cuadrilla
            If vParamAplic.Cooperativa = 16 Then
                Sql4 = "select codtraba from straba where (1=1) "
                Sql4 = Sql4 & " and " & cadSelect
                Sql4 = Sql4 & " order by 1 "
            Else
                Sql4 = "select codtraba from rcuadrilla_trabajador where codcuadrilla = " & Data1.Recordset.Fields(2)
                Sql4 = Sql4 & " and " & cadSelect
                Sql4 = Sql4 & " order by 1 "
            End If
            Set Rs2 = New ADODB.Recordset
            Rs2.Open Sql4, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            
            While Not Rs2.EOF
                '[Monica]01/03/2012: no insertamos los modificados
                Sql4 = "select count(*) from rpartes_trabajador where codtraba = " & DBSet(Rs2!CodTraba, "N")
                Sql4 = Sql4 & " and codvarie = " & DBSet(Rs!codvarie, "N") & " and modificado = 1 "
                Sql4 = Sql4 & " and nroparte = " & Data1.Recordset.Fields(0).Value
                If TotalRegistros(Sql4) = 0 Then
                
                    NumF = SugerirCodigoSiguienteStr("rpartes_trabajador", "numlinea", "nroparte = " & Data1.Recordset.Fields(0))
                    
                    Sql2 = "insert into rpartes_trabajador (nroparte, numlinea, codtraba, codvarie, kilosrec, importe, automatico) values "
                    Sql2 = Sql2 & "(" & Data1.Recordset.Fields(0).Value & "," & DBSet(NumF, "N") & ","
                    Sql2 = Sql2 & DBSet(Rs2!CodTraba, "N") & ","
                    Sql2 = Sql2 & DBSet(Rs!codvarie, "N") & ","
                    Sql2 = Sql2 & DBSet(KilosTrab, "N") & "," & DBSet(ImporteTrab, "N") & ",1)"
                    
                    conn.Execute Sql2
                
                End If
                
                Rs2.MoveNext
            Wend
        
            Set Rs2 = Nothing
                
            '[Monica]07/01/2014: para el caso de Alzira el precio va a depender de la calidad de los kilos recolectados,
            '                    se calculará tanto para los que se han modificado como para los q no se han modificado.
            If vParamAplic.Cooperativa = 4 Then
                SqlNue = "select sum(importe) from ("
                SqlNue = SqlNue & "select rhisfruta_clasif.codcalid, round(rcalidad.eurreccoop * sum(kilosnet),2) importe from rhisfruta_clasif, rcalidad "
                SqlNue = SqlNue & " where rhisfruta_clasif.codvarie = rcalidad.codvarie and rhisfruta_clasif.codcalid = rcalidad.codcalid and "
                SqlNue = SqlNue & " rhisfruta_clasif.codvarie = " & DBSet(Rs!codvarie, "N") & " and rhisfruta_clasif.numalbar in "
                SqlNue = SqlNue & " (select rhisfruta_entradas.numalbar "
                SqlNue = SqlNue & " from rhisfruta_entradas, rhisfruta "
                SqlNue = SqlNue & " where fechaent = " & DBSet(Data1.Recordset.Fields(3), "F") ' fecentrada
                SqlNue = SqlNue & " and codcapat = " & Data1.Recordset.Fields(2) ' capataz
                SqlNue = SqlNue & " and rhisfruta.numalbar = rhisfruta_entradas.numalbar) "
                SqlNue = SqlNue & " group by 1 "
                SqlNue = SqlNue & ") aaaa "
            
'                SqlNue = "select rhisfruta_clasif.codcalid, rcalidad.eurreccoop, sum(kilosnet) kilosnet from rhisfruta_clasif, rcalidad "
'                SqlNue = SqlNue & " where rhisfruta_clasif.codvarie = rcalidad.codvarie and rhisfruta_clasif.codcalid = rcalidad.codcalid and rhisfruta_clasif.numalbar in "
'                SqlNue = SqlNue & " (select rhisfruta_entradas.numalbar "
'                SqlNue = SqlNue & " from rhisfruta_entradas, rhisfruta "
'                SqlNue = SqlNue & " where fechaent = " & DBSet(Data1.Recordset.Fields(3), "F") ' fecentrada
'                SqlNue = SqlNue & " and codcapat = " & Data1.Recordset.Fields(2) ' capataz
'                SqlNue = SqlNue & " and rhisfruta.numalbar = rhisfruta_entradas.numalbar) "
'
'                Set Rs2 = New ADODB.Recordset
'                Rs2.Open SqlNue, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'                While Not Rs2.EOF
'
'
'                    Rs2.MoveNext
'                Wend
'                Set Rs2 = Nothing
                
                ImporteVariedad = DevuelveValor(SqlNue)
                
                Sql4 = "select sum(kilosrec) from rpartes_trabajador where nroparte = " & Data1.Recordset.Fields(0).Value
                Sql4 = Sql4 & " and codvarie = " & DBSet(Rs!codvarie, "N")
                KilosInicio = DevuelveValor(Sql4)
                
                ' A continuacion se prorratea el importe por los kilos de cada trabajador
                SqlNue = "select numlinea, codtraba, kilosrec from rpartes_trabajador where nroparte = " & Data1.Recordset.Fields(0).Value
                SqlNue = SqlNue & " and codvarie = " & DBSet(Rs!codvarie, "N")
                SqlNue = SqlNue & " order by 1"
                
                ImpTot = 0
                
                Set Rs2 = New ADODB.Recordset
                Rs2.Open SqlNue, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                While Not Rs2.EOF
                    ImporteTrab = Round2(DBLet(Rs2!KilosRec, "N") * ImporteVariedad / KilosInicio, 2)
                    ImpTot = ImpTot + ImporteTrab
                    UltLinea = DBLet(Rs2!numlinea, "N")
            
                    Sql5 = "update rpartes_trabajador set importe = " & DBSet(ImporteTrab, "N")
                    Sql5 = Sql5 & " where nroparte = " & Data1.Recordset.Fields(0).Value & " and codtraba = " & DBSet(Rs2!CodTraba, "N")
                    Sql5 = Sql5 & " and codvarie = " & DBSet(Rs!codvarie, "N")
                    Sql5 = Sql5 & " and numlinea = " & DBSet(Rs2!numlinea, "N")
                    
                    conn.Execute Sql5
                    
                    Rs2.MoveNext
                Wend
                
                Set Rs2 = Nothing
                
'                'en el ultimo dejo la diferencia
'                If ImpTot <> ImporteVariedad Then
'                    Sql5 = "update rpartes_trabajador set importe = importe + " & DBSet(ImporteVariedad - ImpTot, "N")
'                    Sql5 = Sql5 & " where nroparte = " & Data1.Recordset.Fields(0).Value
'                    Sql5 = Sql5 & " and numlinea = " & DBSet(UltLinea, "N")
'
'                    conn.Execute Sql5
'                End If
            End If
            '[Monica]07/01/2014: hasta aqui el recalculo de importes por calidad
            
            
            Rs.MoveNext
        Wend
        
        Set Rs = Nothing
        
        ' insertamos en rpartes_trabajador: todos los trabajadores de la cuadrilla
        ' con todas los gastos prorrateados del parte
        SQL = "select numlinea, codgasto, importe from rpartes_gastos where nroparte=" & Data1.Recordset.Fields(0).Value
        
        Set Rs = New ADODB.Recordset
        Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        While Not Rs.EOF
        
            If vParamAplic.Cooperativa = 16 Then
                Sql4 = "select codtraba from straba where (1=1) "
                Sql4 = Sql4 & " and " & cadSelect
                Sql4 = Sql4 & " order by 1 "
            
            Else
        
                Sql4 = "select codtraba from rcuadrilla_trabajador where codcuadrilla = " & Data1.Recordset.Fields(2)
                Sql4 = Sql4 & " and " & cadSelect
                Sql4 = Sql4 & " order by 1 "
            End If
            
            Set Rs2 = New ADODB.Recordset
            Rs2.Open Sql4, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            
            While Not Rs2.EOF
                ImporteTrab = Round(Rs!Importe / NroTrabajadores, 2)
                    
                NumF = SugerirCodigoSiguienteStr("rpartes_trabajador", "numlinea", "nroparte = " & Data1.Recordset.Fields(0))
                
                Sql2 = "insert into rpartes_trabajador (nroparte, numlinea, codtraba, codgasto, importe, automatico) values "
                Sql2 = Sql2 & "(" & Data1.Recordset.Fields(0).Value & "," & DBSet(NumF, "N") & ","
                Sql2 = Sql2 & DBSet(Rs2!CodTraba, "N") & ","
                Sql2 = Sql2 & DBSet(Rs!Codgasto, "N") & ","
                Sql2 = Sql2 & DBSet(ImporteTrab, "N") & ",1)"
                
                conn.Execute Sql2
                
                Rs2.MoveNext
            Wend
            Set Rs2 = Nothing
            Rs.MoveNext
        Wend
        
        
        ' insertamos en rpartes_trabajador: el plus del capataz
        SQL = "select codtraba, pluscapataz from straba where " & cadSelect
        Set Rs = New ADODB.Recordset
        Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        While Not Rs.EOF
            If DBLet(Rs!PlusCapataz, "N") <> 0 Then
                NumF = SugerirCodigoSiguienteStr("rpartes_trabajador", "numlinea", "nroparte = " & Data1.Recordset.Fields(0))
        
                Sql2 = "insert into rpartes_trabajador (nroparte, numlinea, codtraba, codgasto, importe, automatico) values "
                Sql2 = Sql2 & "(" & Data1.Recordset.Fields(0).Value & "," & DBSet(NumF, "N") & ","
                Sql2 = Sql2 & DBSet(Rs!CodTraba, "N") & ","
                Sql2 = Sql2 & ValorNulo & ","
                
                If vParamAplic.Cooperativa = 16 Then
                    Dim NroCapataces As Integer
                    SQL = "select count(*) from straba where " & cadSelect & " and pluscapataz <> 0"
                    
                    NroCapataces = DevuelveValor(SQL)
                    If NroCapataces > 0 Then
                        Importe = Round2((DBLet(Rs!PlusCapataz, "N") * (NroTrabajadores - NroCapataces)) / NroCapataces, 2)
                    
                        Sql2 = Sql2 & DBSet(Importe, "N") & ",1)"
                    Else
                        Sql2 = Sql2 & "0,1)"
                    End If
                Else
                    Sql2 = Sql2 & DBSet(Rs!PlusCapataz, "N") & ",1)"
                End If
                
                conn.Execute Sql2
            
            End If
            Rs.MoveNext
        Wend
        Set Rs = Nothing
        
        conn.CommitTrans
        RecalcularImportes = True
    End If
        
    Exit Function

    

eRecalcularImportes:
    conn.RollbackTrans
    MuestraError Err.Number, "Recalcular Importes", Err.Description
End Function

Private Sub mnNuevo_Click()
    BotonAnyadir
End Sub


Private Sub mnModificar_Click()

'    If Data1.Recordset!impreso = 1 Then
'        If MsgBox("Este albarán está facturado y/o cobrado. ¿ Desea continuar ?", vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then
'            Exit Sub
'        End If
'    End If

    If Modo = 5 Then 'Modificar lineas
        'bloquea la tabla cabecera de factura: scafac
        If BLOQUEADesdeFormulario(Me) Then
            'bloquear la tabla cabecera de albaranes de la factura: scafac1
            If BloqueaAlbxFac Then
                If BloqueaLineasFac Then BotonModificarLinea (NumTabMto)
            End If
        End If
         
    Else   'Modificar Pedido
        'bloquea la tabla cabecera de factura: scafac
        If BLOQUEADesdeFormulario(Me) Then
            'bloquear la tabla cabecera de albaranes de la factura: scafac1
            BotonModificar
        End If
    End If
End Sub


Private Function BloqueaAlbxFac() As Boolean
'bloquea todos los albaranes de la factura
Dim SQL As String

    On Error GoTo EBloqueaAlb
    
    BloqueaAlbxFac = False
    'bloquear cabecera albaranes x factura
    SQL = "select * FROM scafac1 "
    SQL = SQL & ObtenerWhereCP(True) & " FOR UPDATE"
    conn.Execute SQL, , adCmdText
    BloqueaAlbxFac = True

EBloqueaAlb:
    If Err.Number <> 0 Then BloqueaAlbxFac = False
End Function


Private Function BloqueaLineasFac() As Boolean
'bloquea todas las lineas de la factura
Dim SQL As String

    On Error GoTo EBloqueaLin

    BloqueaLineasFac = False
    'bloquear cabecera albaranes x factura
    SQL = "select * FROM slifac "
    SQL = SQL & ObtenerWhereCP(True) & " FOR UPDATE"
    conn.Execute SQL, , adCmdText
    BloqueaLineasFac = True

EBloqueaLin:
    If Err.Number <> 0 Then BloqueaLineasFac = False
End Function

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
    If Index = 3 Then 'codigo de cliente
        Cliente = Text1(Index).Text
    End If
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
Dim devuelve As String
Dim cadMen As String
Dim SQL As String

        
    If Not PerderFocoGnral(Text1(Index), Modo) Then Exit Sub
        
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
          
    'Si queremos hacer algo ..
    Select Case Index
        Case 1, 3 'Fecha parte / fecha de entradas
            '[Monica]28/08/2013: comprobamos que la fecha esté en la campaña
            If Text1(Index).Text <> "" Then PonerFormatoFecha Text1(Index), True
            
        Case 2 'cuadrilla
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), "rcuadrilla", "codcapat")
                If Modo = 1 Then Exit Sub
                If Text2(Index).Text = "" Then
                    cadMen = "No existe la Cuadrilla: " & Text1(Index).Text & vbCrLf
                    cadMen = cadMen & "¿Desea crearla?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmCua = New frmManCuadrillas
                        frmCua.DatosADevolverBusqueda = "0|1|"
                        Text1(Index).Text = ""
                        TerminaBloquear
                        frmCua.Show vbModal
                        Set frmCua = Nothing
                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
                    Else
                        Text1(Index).Text = ""
                    End If
                    PonerFoco Text1(Index)
                Else
                    SQL = "select nomcapat from rcapataz where codcapat = " & DBSet(Text2(Index).Text, "N")
                    Text2(Index) = DevuelveValor(SQL)
                End If
            End If
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
'    CadB = ObtenerBusqueda(Me)
    CadB = ObtenerBusqueda2(Me, BuscaChekc, 1)
    
    If chkVistaPrevia = 1 Then
        EsCabecera = True
        If CadB1 <> "" Then CadB = CadB & " and " & CadB1
        MandaBusquedaPrevia CadB
    ElseIf CadB <> "" Then
        'Se muestran en el mismo form
        CadenaConsulta = "select rpartes.* from " & NombreTabla & " LEFT JOIN rpartes_variedad ON rpartes.nroparte=rpartes_variedad.nroparte "
        CadenaConsulta = CadenaConsulta & " WHERE " & CadB
        If CadB1 <> "" Then CadenaConsulta = CadenaConsulta & " and " & CadB1
        CadenaConsulta = CadenaConsulta & " GROUP BY rpartes.nroparte " & Ordenacion
'        CadenaConsulta = "select palets.* from " & NombreTabla
'        CadenaConsulta = CadenaConsulta & " WHERE " & CadB & " GROUP BY palets.numpalet " & Ordenacion
        PonerCadenaBusqueda
    End If
End Sub


Private Sub MandaBusquedaPrevia(CadB As String)
'Carga el formulario frmBuscaGrid con los valores correspondientes
Dim cad As String
Dim Tabla As String
Dim Titulo As String
Dim Desc As String, devuelve As String
    'Llamamos a al form
    '##A mano
'    cad = ""
'    cad = cad & "Parte|rpartes.nroparte|N|0000000|11·"
'    cad = cad & "Fecha|rpartes.fechapar|F||14·"
'    cad = cad & "Cuadrilla|rpartes.codcuadrilla|N|000000|10·"
'    cad = cad & "Capataz|rcuadrilla.codcapat|N|0000|10·"
'    cad = cad & "Nombre|rcapataz.nomcapat|N||55·"
'
''    Cad = Cad & "Cod|rhisfruta.codvarie|N||7·" 'ParaGrid(Text1(3), 10, "Cliente")
''    Cad = Cad & "Nombre|variedades.nomvarie|N||20·"
''    Cad = Cad & "Socio|rhisfruta.codsocio|N||10·" 'ParaGrid(Text1(3), 10, "Cliente")
''    Cad = Cad & "Nombre|rsocios.nomsocio|N||28·"
''    Cad = Cad & "Campo|rhisfruta.codcampo|N||10·"
'
'    tabla = NombreTabla & " INNER JOIN rcuadrilla ON rpartes.codcuadrilla=rcuadrilla.codcuadrilla "
'    tabla = "(" & tabla & ") INNER JOIN rcapataz ON rcuadrilla.codcapat=rcapataz.codcapat "
'
'    Titulo = "Partes de Campos"
'    devuelve = "0|"
'
'    If cad <> "" Then
'        Screen.MousePointer = vbHourglass
'        Set frmB = New frmBuscaGrid
'        frmB.vCampos = cad
'        frmB.vtabla = tabla
'        frmB.vSQL = CadB
'        HaDevueltoDatos = False
'        '###A mano
'        frmB.vDevuelve = "0|1|"
'        frmB.vDevuelve = devuelve
'        frmB.vTitulo = Titulo
'        frmB.vSelElem = 0
''        frmB.vConexionGrid = cAgro  'Conexión a BD: Ariagro
'        If Not EsCabecera Then frmB.Label1.FontSize = 11
''        frmB.vBuscaPrevia = chkVistaPrevia
'        '#
'        frmB.Show vbModal
'        Set frmB = Nothing
''        If EsCabecera Then
''            PonerCadenaBusqueda
''            Text1(0).Text = Format(Text1(0).Text, "0000000")
''        End If
'        'Si ha puesto valores y tenemos que es formulario de busqueda entonces
'        'tendremos que cerrar el form lanzando el evento
'        If HaDevueltoDatos Then
'            If (Not Data1.Recordset.EOF) And DatosADevolverBusqueda <> "" Then _
'                cmdRegresar_Click
'        Else   'de ha devuelto datos, es decir NO ha devuelto datos
'            PonerFoco Text1(kCampo)
'        End If
'    End If
'    Screen.MousePointer = vbDefault


    Set frmPartesCam = New frmBasico2
    
    AyudaPartesCampo frmPartesCam
    
    Set frmPartesCam = Nothing
    



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
'            Text1(0).BackColor = vblightblue
        End If
        Exit Sub
    Else
        Data1.Recordset.MoveFirst
        PonerModo 2
'        LLamaLineas Modo, 0, "DataGrid2"
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
Dim i As Integer


    On Error GoTo EPonerLineas

    Screen.MousePointer = vbHourglass
    
    If Data1.Recordset.RecordCount > 0 Then
        CargaGrid DataGrid2, Data3, True ' rpartes_trabajador
        CargaGrid DataGrid1, Data2, True ' rpartes_variedad
        CargaGrid DataGrid3, Adoaux(0), True ' rpartes_gastos
        CargaGrid DataGrid4, Data4, True ' rpartes_variedad
        
    Else
        CargaGrid DataGrid2, Data3, False ' rpartes_trabajador
        CargaGrid DataGrid1, Data2, False ' rpartes_variedad
        CargaGrid DataGrid3, Adoaux(0), False ' rpartes_gastos
        CargaGrid DataGrid4, Data4, False ' rpartes_gastos
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
Dim SQL As String

    On Error Resume Next

    If Data1.Recordset.EOF Then Exit Sub
    B = PonerCamposForma2(Me, Data1, 2, "Frame2")
    
'    FormatoDatosTotales
    Text2(2).Text = PonerNombreDeCod(Text1(2), "rcuadrilla", "codcapat")
    SQL = "select nomcapat from rcapataz where codcapat = " & DBSet(Text2(2).Text, "N")
    Text2(2).Text = DevuelveValor(SQL)
    
    Modo = 2
    
    PonerCamposLineas 'Pone los datos de las tablas de lineas
    
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
Dim i As Byte, NumReg As Byte
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
    If DatosADevolverBusqueda <> "" Or NroParte <> "" Then
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
    
        
    'Bloquea los campos Text1 sino estamos modificando/Insertando Datos
    'Si estamos en Insertar además limpia los campos Text1
    'si estamos en modificar bloquea las compos que son clave primaria
    BloquearText1 Me, Modo
    BloquearCombo Me, Modo
    Me.Check1(0).Enabled = (Modo = 1)
    
    B = (Modo <> 1)
    'Campos Nº Parte bloqueado y en azul
    BloquearTxt Text1(0), B, True
'    BloquearTxt Text1(3), b 'referencia
    
    'Si no es modo lineas Boquear los TxtAux
    For i = 0 To txtAux.Count - 1
        txtAux(i).visible = False
        BloquearTxt txtAux(i), True
    Next i
    For i = 0 To btnBuscar.Count - 1
        btnBuscar(i).visible = False
        btnBuscar(i).Enabled = True
    Next i
    For i = 0 To txtAux3.Count - 1
        txtAux3(i).visible = False
        BloquearTxt txtAux3(i), True
    Next i
    For i = 0 To txtAux2.Count - 1
        txtAux2(i).visible = False
        BloquearTxt txtAux2(i), True
    Next i
    
    For i = 2 To 3
        BloquearTxt txtAux2(i), Not (Modo = 5 And NumTabMto = 1)
        txtAux2(i).Enabled = (Modo = 5 And NumTabMto = 1)
    Next i
    
    Text2(6).visible = False
    Text2(6).Enabled = False
    
    For i = 1 To 4
        If (i = 1 Or i = 2) Then
            BloquearTxt txtAux3(i), Not (Modo = 5 And ((SePuedeModificar And ModificaLineas = 2) Or ModificaLineas = 1) And NumTabMto = 0)
            txtAux3(i).Enabled = (Modo = 5 And ((SePuedeModificar And ModificaLineas = 2) Or ModificaLineas = 1) And NumTabMto = 0)
        End If
        If i = 3 Then
            BloquearTxt txtAux3(i), Not (Modo = 5 And (Not SePuedeModificar And ModificaLineas = 2) And NumTabMto = 0)
            txtAux3(i).Enabled = (Modo = 5 And (Not SePuedeModificar And ModificaLineas = 2) And NumTabMto = 0)
        End If
        If i = 4 Then
            BloquearTxt txtAux3(i), Not (Modo = 5) ' And ((SePuedeModificar And ModificaLineas = 2) Or ModificaLineas = 1))
            txtAux3(i).Enabled = (Modo = 5) ' And ((SePuedeModificar And ModificaLineas = 2) Or ModificaLineas = 1))
        End If
    Next i
    
    '[Monica]17/06/2013
    BloquearTxt txtAux3(8), Not (Modo = 5 And (Not SePuedeModificar And ModificaLineas = 2) And NumTabMto = 0)
    txtAux3(8).Enabled = (Modo = 5 And (Not SePuedeModificar And ModificaLineas = 2) And NumTabMto = 0)
    
    '[Monica]23/12/2016
    BloquearTxt txtAux3(9), Not (Modo = 5 And (Not SePuedeModificar And ModificaLineas = 2) And NumTabMto = 0)
    txtAux3(9).Enabled = (Modo = 5 And (Not SePuedeModificar And ModificaLineas = 2) And NumTabMto = 0)
    
    
    For i = 1 To 2
        btnBuscar(i).visible = (Modo = 5 And ((SePuedeModificar And ModificaLineas = 2) Or ModificaLineas = 1) And NumTabMto = 0)
        btnBuscar(i).Enabled = (Modo = 5 And ((SePuedeModificar And ModificaLineas = 2) Or ModificaLineas = 1) And NumTabMto = 0)
    Next i
    
    BloquearTxt txtAux(5), Not (Modo = 5 And (ModificaLineas = 1 Or ModificaLineas = 2) And NumTabMto = 2)
    txtAux(5).Enabled = (Modo = 5 And (ModificaLineas = 1 Or ModificaLineas = 2) And NumTabMto = 2)
    BloquearTxt txtAux(2), Not (Modo = 5 And (ModificaLineas = 1) And NumTabMto = 2)
    txtAux(2).Enabled = (Modo = 5 And (ModificaLineas = 1) And NumTabMto = 2)
    
    '---------------------------------------------
    B = (Modo <> 0 And Modo <> 2)
    cmdCancelar.visible = B
    cmdAceptar.visible = B
    
    BloquearImgBuscar Me, Modo, ModificaLineas
    BloquearImgFec Me, 0, Modo
    
    Text1(2).Enabled = (Modo = 1 Or Modo = 3)
    imgBuscar(0).Enabled = (Modo = 1 Or Modo = 3)
    imgBuscar(0).visible = (Modo = 1 Or Modo = 3)
    
                    
    Me.chkVistaPrevia.Enabled = (Modo <= 2)
       
    ' *** si n'hi han llínies i algún tab que no te datagrid ***
    Select Case NumTabMto
        Case 0
            BloquearFrameAux Me, "FrameAux0", Modo, NumTabMto
        Case 1
            BloquearFrameAux Me, "Frame3", Modo, NumTabMto
        Case 2
            BloquearFrameAux Me, "Frame4", Modo, NumTabMto
    End Select
    
    If indFrame = 1 Then
        txtAux2(2).Enabled = (ModificaLineas = 1) And (NumTabMto = 1)
        txtAux2(2).visible = (ModificaLineas = 1) And (NumTabMto = 1)
        btnBuscar(0).Enabled = (ModificaLineas = 1) And (NumTabMto = 1)
        btnBuscar(0).visible = (ModificaLineas = 1) And (NumTabMto = 1)
    End If
        
        
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


Private Function DatosOk() As Boolean
'Comprobar que los datos de la cabecera son correctos antes de Insertar o Modificar
'la cabecera del Pedido
Dim B As Boolean

    On Error GoTo EDatosOK

    DatosOk = False
    
'    ComprobarDatosTotales

    'comprobamos datos OK de la tabla scafac
    B = CompForm2(Me, 2, "Frame2") ' , 1) 'Comprobar formato datos ok de la cabecera: opcion=1
    If Not B Then Exit Function
    
    DatosOk = B
    
EDatosOK:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function


Private Function DatosOkLinea() As Boolean
Dim B As Boolean
Dim i As Byte

    On Error GoTo EDatosOkLinea

    DatosOkLinea = False
    B = True

    For i = 0 To txtAux.Count - 1
        If i = 4 Or i = 6 Or i = 7 Then
            If txtAux(i).Text = "" Then
                MsgBox "El campo " & txtAux(i).Tag & " no puede ser nulo", vbExclamation
                B = False
                PonerFoco txtAux(i)
                Exit Function
            End If
        End If
    Next i
            
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
    
    If EstaPartePagado(Data1.Recordset.Fields(0).Value) Then
        MsgBox "Sobre este Parte ya se ha hecho la impresión del recibo. No se permite realizar la operación.", vbExclamation
        Exit Sub
    End If
    
    If EstaParteenHoras(Data1.Recordset.Fields(0).Value) Then
        If MsgBox("Este Parte ya se ha traspasado a horas. ¿ Desea Continuar ?", vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then
            Exit Sub
        End If
    End If
    
    If BloqueaRegistro(NombreTabla, "nroparte = " & Data1.Recordset!NroParte) Then
'    If BLOQUEADesdeFormulario2(Me, Data1, 1) Then
        Select Case Index
            Case 0 'rpartes_trabajador
                NumTabMto = 0
                Select Case Button.Index
                    Case 1
                        BotonAnyadirLinea Index
                    Case 2
                        BotonModificarLinea Index
                    Case 3
                        BotonEliminarLinea Index
                    Case 4
                        IntroducirCajones
                    Case Else
                End Select
            
            Case 1 'rpartes_gastos
                NumTabMto = 1
                Select Case Button.Index
                    Case 1
                        BotonAnyadirLinea Index
                    Case 2
                        BotonModificarLinea Index
                    Case 3
                        BotonEliminarLinea Index
                    Case Else
                End Select
                
            Case 2 'rpartes_variedad
                NumTabMto = 2
                Select Case Button.Index
                    Case 1
                        BotonAnyadirLinea Index
                    Case 2
                        BotonModificarLinea Index
                    Case 3
                        BotonEliminarLinea Index
                    Case Else
                End Select
            
                
        End Select
        
    End If

End Sub

Private Sub IntroducirCajones()
Dim frmPartesCajas As frmManPartesCajas

    If Text1(0).Text = "" Then Exit Sub
    
    If Me.Data2.Recordset.Fields.Count = 0 Then Exit Sub

    Set frmPartesCajas = New frmManPartesCajas

    frmPartesCajas.NroParte = Text1(0).Text
    frmPartesCajas.Show vbModal

    Set frmPartesCajas = Nothing
    
    RepartoxCajas
    
    TerminaBloquear
    PosicionarData
    
    CargaGrid DataGrid1, Data2, True
    CargaGrid DataGrid2, Data3, True
    If vParamAplic.HayResumenCajas = 1 Then CargaGrid DataGrid4, Data4, True
    PonerModo 2
    

End Sub

Private Sub RepartoxCajas()
Dim SQL As String
Dim Sql2 As String
Dim TotalCajas As Currency
Dim TotalImporte As Currency
Dim TotalKilos As Long

Dim TImporte As Currency
Dim TKilos As Long

Dim Rs As ADODB.Recordset
Dim Rs2 As ADODB.Recordset

Dim NumLin As Integer
Dim Importe As Currency
Dim Kilos As Long

    
    SQL = "select codvarie, sum(coalesce(numcajas,0)) numcajas, sum(coalesce(kilosrec,0)) kilos, sum(coalesce(importe,0)) importe from rpartes_trabajador where nroparte = " & DBSet(Text1(0).Text, "N")
    SQL = SQL & " group by 1 order by 1"
    
    Sql2 = "select sum(numcajas) from (" & SQL & ") aaaaa"
    If DevuelveValor(Sql2) = 0 Then Exit Sub
    
    
    
    Set Rs = New ADODB.Recordset
    
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not Rs.EOF
        TotalCajas = DBLet(Rs.Fields(1))
        TotalKilos = DBLet(Rs.Fields(2))
        TotalImporte = DBLet(Rs.Fields(3))
    
        TImporte = 0
        TKilos = 0
        
        Sql2 = "select * from rpartes_trabajador where nroparte = " & DBSet(Text1(0).Text, "N")
        Sql2 = Sql2 & " and codvarie = " & DBSet(Rs!codvarie, "N") & " order by numlinea "
        
        Set Rs2 = New ADODB.Recordset
        
        Rs2.Open Sql2, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not Rs2.EOF
            Importe = Round2(DBLet(Rs2!NumCajas, "N") * TotalImporte / TotalCajas, 2)
            Kilos = Round2(DBLet(Rs2!NumCajas, "N") * TotalKilos / TotalCajas, 0)
            
            SQL = "update rpartes_trabajador set importe = " & DBSet(Importe, "N")
            SQL = SQL & " , kilosrec = " & DBSet(Kilos, "N")
            SQL = SQL & " where nroparte = " & DBSet(Text1(0).Text, "N")
            SQL = SQL & " and numlinea = " & DBSet(Rs2!numlinea, "N")
            
            conn.Execute SQL
            
            TImporte = TImporte + Importe
            TKilos = TKilos + Kilos
        
            NumLin = DBLet(Rs2!numlinea, "N")
            
            Rs2.MoveNext
        Wend
        Set Rs2 = Nothing
        
        If TImporte <> TotalImporte Or TKilos <> TotalKilos Then
            SQL = "update rpartes_trabajador set importe = importe + " & DBSet(TotalImporte - TImporte, "N")
            SQL = SQL & ", kilosrec = kilosrec + " & DBSet(TotalKilos - TKilos, "N")
            SQL = SQL & " where nroparte = " & DBSet(Text1(0).Text, "N")
            SQL = SQL & " and numlinea = " & DBSet(NumLin, "N")
            
            conn.Execute SQL
        End If
    
        Rs.MoveNext
    
    Wend
    
    Set Rs = Nothing
    



End Sub

Private Sub BotonEliminarLinea(Index As Integer)
Dim cad As String
Dim SQL As String

    On Error GoTo EEliminarLinea

    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub

    ' *** repasar el nom de l'adodc, l'index del Field i el camp que te la PK ***
    ' ### [Monica] 26/09/2006 dejamos modificar y eliminar el codigo 0
    'El registre de codi 0 no es pot Modificar ni Eliminar
    'If EsCodigoCero(CStr(Data1.Recordset.Fields(0).Value), FormatoCampo(text1(0))) Then Exit Sub
    ' ***************************************************************************

    Select Case Index
        Case 0 'gastos individuales
            'comprobamos que la linea de gastos trabajador no es de kilos
            If Not SePuedeModificarLinea(CStr(Data3.Recordset.Fields(0).Value), CStr(Data3.Recordset.Fields(1).Value)) Then
                cad = "¿Seguro que desea eliminar la Linea?"
                cad = cad & vbCrLf & "Parte: " & Data3.Recordset.Fields(0)
            Else
                ' *************** canviar la pregunta ****************
                cad = "¿Seguro que desea eliminar el Gasto Individual?"
                cad = cad & vbCrLf & "Parte: " & Data3.Recordset.Fields(0)
            End If
        
            
            If MsgBox(cad, vbQuestion + vbYesNo) = vbYes Then
                On Error GoTo EEliminarLinea
                Screen.MousePointer = vbHourglass
                NumRegElim = Data3.Recordset.AbsolutePosition
                
                If Not EliminarLinea(Index) Then
                    Screen.MousePointer = vbDefault
                    Exit Sub
                Else
                    If SituarDataTrasEliminar(Data3, NumRegElim, True) Then
                        PonerCampos
                    Else
                        LimpiarCampos
                        PonerModo 0
                    End If
                End If
            End If
            Screen.MousePointer = vbDefault
       Case 1 'gastos del parte
            ' *************** canviar la pregunta ****************
            cad = "¿Seguro que desea eliminar el Gasto del Parte?"
            cad = cad & vbCrLf & "Parte: " & Adoaux(0).Recordset.Fields(0)
            cad = cad & vbCrLf & "Código: " & Adoaux(0).Recordset.Fields(2) & "-" & Adoaux(0).Recordset.Fields(3)
            
            If MsgBox(cad, vbQuestion + vbYesNo) = vbYes Then
                On Error GoTo EEliminarLinea
                Screen.MousePointer = vbHourglass
                NumRegElim = Adoaux(0).Recordset.AbsolutePosition
                TerminaBloquear
                SQL = "delete from rpartes_gastos where nroparte = " & Adoaux(0).Recordset.Fields(0)
                SQL = SQL & " and numlinea = " & Adoaux(0).Recordset.Fields(1)
                conn.Execute SQL
                
                SituarDataTrasEliminar Adoaux(0), NumRegElim
                
                CargaGrid DataGrid3, Adoaux(0), True
'                SSTab1.Tab = 1

            End If
            Screen.MousePointer = vbDefault
       
        Case 2 ' variedades
            ' *************** canviar la pregunta ****************
            cad = "¿Seguro que desea eliminar los kilos de la Nota?"
            cad = cad & vbCrLf & "Parte: " & Data2.Recordset.Fields(0)
            cad = cad & vbCrLf & "Nota: " & Data2.Recordset.Fields(2) & "-" & Data2.Recordset.Fields(4)
            
            If MsgBox(cad, vbQuestion + vbYesNo) = vbYes Then
                On Error GoTo EEliminarLinea
                Screen.MousePointer = vbHourglass
                NumRegElim = Data2.Recordset.AbsolutePosition
                TerminaBloquear
                SQL = "delete from rpartes_variedad where nroparte = " & Data2.Recordset.Fields(0)
                SQL = SQL & " and numlinea = " & Data2.Recordset.Fields(1)
                conn.Execute SQL
                
                SituarDataTrasEliminar Data2, NumRegElim
                
                CargaGrid DataGrid1, Data2, True
'                SSTab1.Tab = 1

            End If
            Screen.MousePointer = vbDefault
        
       
    End Select
       
EEliminarLinea:
    Screen.MousePointer = vbDefault
    If Err.Number > 0 Then MuestraError Err.Number, "Eliminar Linea de Albarán", Err.Description

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
            
'        Case 8  ' Traer entradads
'            mnTraerEntradas_Click
'
'        Case 9  ' Recalcular importes
'            mnRecalcularImportes_Click
            
        Case 8 ' Informe de dias trabajados
            mnInforme_Click
            
        Case 13   'Salir
            mnSalir_Click
        Case btnPrimero To btnPrimero + 3 'Flechas Desplazamiento
            Desplazamiento (Button.Index - btnPrimero)
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


Private Sub PonerBotonCabecera(B As Boolean)
'Pone el boton de Regresar a la Cabecera si pasamos a MAntenimiento de Lineas
'o Pone los botones de Aceptar y cancelar en Insert,update o delete lineas
    On Error Resume Next

    Me.cmdAceptar.visible = Not B
    Me.cmdCancelar.visible = Not B
    Me.cmdRegresar.visible = B
    Me.cmdRegresar.Caption = "Cabecera"
    If B Then
        Me.lblIndicador.Caption = "Líneas " & TituloLinea
        PonerFocoBtn Me.cmdRegresar
    End If
    'Habilitar las opciones correctas del menu segun Modo
    PonerModoOpcionesMenu (Modo)
    PonerOpcionesMenu 'Habilitar las opciones correctas del menu segun Nivel de Acceso
    DataGrid2.Enabled = Not B
    
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub CargaGrid(ByRef vDataGrid As DataGrid, ByRef vData As Adodc, enlaza As Boolean)
Dim B As Boolean
Dim Opcion As Byte
Dim SQL As String

    On Error GoTo ECargaGrid

    B = DataGrid1.Enabled
    Select Case vDataGrid.Name
        Case "DataGrid1"
            Opcion = 1
        Case "DataGrid2"
            Opcion = 2
        Case "DataGrid3" 'clasificacion
            Opcion = 3
        Case "DataGrid4" 'clasificacion
            Opcion = 4
    End Select
    
    SQL = MontaSQLCarga(enlaza, Opcion)
    CargaGridGnral vDataGrid, vData, SQL, PrimeraVez
    
    vDataGrid.RowHeight = 270
    
    CargaGrid2 vDataGrid, vData
    vDataGrid.ScrollBars = dbgAutomatic
    
     B = (Modo = 5) And (ModificaLineas = 1 Or ModificaLineas = 2)
     vDataGrid.Enabled = Not B
    
    
    
    If Opcion = 1 And Data1.Recordset.RecordCount > 0 Then
        Text2(1).Text = DevuelveValor("select sum(importe) from rpartes_trabajador where nroparte = " & Data1.Recordset.Fields(0).Value)
        Text2(1).Text = Format(Text2(1).Text, "###,###,##0.00")
    End If
    
    Exit Sub
    
ECargaGrid:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid", Err.Description
End Sub


Private Sub CargaGrid2(ByRef vDataGrid As DataGrid, ByRef vData As Adodc)
Dim tots As String
    
    On Error GoTo ECargaGrid

    Select Case vDataGrid.Name
        Case "DataGrid1" 'rpartes_variedad
'           SQL = "SELECT nroparte, codvarie, nomvarie, kilosrec
            tots = "N||||0|;N||||0|;S|txtAux(2)|T|Nota|1000|;"
            tots = tots & "S|txtAux(3)|T|Codigo|900|;"
            If vParamAplic.Cooperativa = 16 Then
                tots = tots & "S|txtAux(4)|T|Nombre Variedad|1900|;S|txtAux(6)|T|Cajas|1000|;S|txtAux(5)|T|Kilos|1500|;"
            Else
                tots = tots & "S|txtAux(4)|T|Nombre Variedad|1900|;S|txtAux(6)|T|Horas|1000|;S|txtAux(5)|T|Kilos|1500|;"
            End If
            arregla tots, DataGrid1, Me, 350
         
         Case "DataGrid2" 'rpartes_trabajador
'           SQL = "SELECT nroparte, numlinea, codtraba, nomtraba, codcoste, nomcoste, importe
            tots = "N||||0|;N||||0|;"
            tots = tots & "S|txtAux3(1)|T|Codigo|950|;S|btnBuscar(1)|B|||;"
            tots = tots & "S|Text2(3)|T|Nombre Trabajador|3000|;S|txtAux3(2)|T|Gasto|800|;S|btnBuscar(2)|B|||;S|Text2(4)|T|Descripcion Gasto|2435|;"
            tots = tots & "S|txtAux3(5)|T|Código|850|;S|Text2(0)|T|Variedad|1650|;"
            tots = tots & "S|txtAux3(9)|T|Cajas|930|;S|txtAux3(3)|T|KilosRec|1000|;S|txtAux3(8)|T|Horas|720|;S|txtAux3(4)|T|Importe|1150|;N||||0|;"
            arregla tots, DataGrid2, Me, 350
            
         Case "DataGrid3" 'rpartes_gastos (gastos generales del parte)
'       SQL = SELECT nroparte, numlinea, codgasto, nomgasto, importe
            tots = "N||||0|;N||||0|;"
            tots = tots & "S|txtAux2(2)|T|Codigo|900|;S|btnBuscar(0)|B|||;"
            tots = tots & "S|Text2(6)|T|Descripción Gasto|3700|;"
            tots = tots & "S|txtAux2(3)|T|Importe|1570|;"
            arregla tots, DataGrid3, Me, 350
         
         Case "DataGrid4" 'rpartes_gastos (gastos generales del parte)
'       SQL = SELECT nroparte, numlinea, codgasto, nomgasto, importe
            tots = ""
            tots = tots & "S|txtAux4(0)|T|Codigo|900|;"
            tots = tots & "S|txtAux4(1)|T|Nombre Variedad|1900|;"
            tots = tots & "S|txtAux4(2)|T|Cajas|900|;"
            tots = tots & "S|txtAux4(3)|T|Kilos|1100|;"
            tots = tots & "S|txtAux4(4)|T|Importe|1400|;"
            arregla tots, DataGrid4, Me, 350
    
    End Select
    
    vDataGrid.HoldFields
    Exit Sub
    
ECargaGrid:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid", Err.Description
End Sub


Private Sub Toolbar5_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1  ' Traer entradads
            mnTraerEntradas_Click
        
        Case 2  ' Recalcular importes
            mnRecalcularImportes_Click
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
Dim Nota As String
Dim campo2 As String
Dim Variedad As String
Dim Capataz As String
Dim CapatazParte  As String
Dim numalbar As String
Dim cad As String
Dim SQL As String
Dim Rs As ADODB.Recordset


    'Quitar espacios en blanco
    If Not PerderFocoGnralLineas(txtAux2(Index), ModificaLineas) Then Exit Sub

    Select Case Index
        Case 2 ' numero de nota
            If txtAux(Index) <> "" Then
                ' comprobamos que el numero de nota existe y nos traemos la variedad y los kilos
                
                numalbar = ""
                numalbar = DevuelveDesdeBDNew(cAgro, "rhisfruta_entradas", "numalbar", "numnotac", txtAux(Index), "N")
                
                If numalbar = "" Then
                    MsgBox "Este nro.de nota no existe. Reintroduzca.", vbExclamation
                    PonerFoco txtAux(Index)
                Else
                    If TotalRegistros("select count(*) from rpartes_variedad where nroparte = " & Data1.Recordset.Fields(0) & " and numnotac = " & txtAux(Index).Text) <> 0 Then
                        MsgBox "Esta nota ya ha sido introducida en el parte. Revise.", vbExclamation
                        PonerFoco txtAux(Index)
                    Else
                    
                        SQL = "select codcapat, kilosnet, fechaent from rhisfruta_entradas where numnotac = " & DBSet(txtAux(Index), "N")
                        
                        Set Rs = New ADODB.Recordset
                        Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                        
                        If Not Rs.EOF Then
                            Capataz = DBLet(Rs!codcapat, "N")
                        
                            cad = "select codcapat from rcuadrilla where codcuadrilla = " & Data1.Recordset.Fields(2).Value
                            CapatazParte = DevuelveValor(cad)
                            If CStr(Capataz) <> CStr(CapatazParte) Or DBLet(Rs!FechaEnt, "F") <> Data1.Recordset.Fields(3) Then
                                If MsgBox("Esta nota de campo no es del capataz del parte o no es de la fecha de entrada. ¿ Desea continuar ?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
                                    txtAux(3).Text = DevuelveDesdeBDNew(cAgro, "rhisfruta", "codvarie", "numalbar", numalbar, "N")
                                    txtAux(4).Text = PonerNombreDeCod(txtAux(3), "variedades", "nomvarie", "codvarie", "N")
                                
                                    txtAux(5).Text = DBLet(Rs!KilosNet, "N")
                                End If
                            Else
                                txtAux2(Index).Text = ""
                                txtAux(3).Text = ""
                                txtAux(4).Text = ""
                                txtAux(5).Text = ""
                                PonerFoco txtAux2(Index)
                            End If
                        End If
                    
                        Set Rs = Nothing
                    End If
                End If
            Else
                txtAux(3).Text = ""
                txtAux(4).Text = ""
                txtAux(5).Text = ""
            End If


        Case 5 ' kilos
            If PonerFormatoEntero(txtAux(Index)) Then cmdAceptar.SetFocus
    
    End Select

End Sub

Private Sub txtAux2_GotFocus(Index As Integer)
    ConseguirFocoLin txtAux2(Index)
End Sub

Private Sub txtAux2_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
    KEYdown KeyCode
End Sub

Private Sub txtAux2_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtAux2_LostFocus(Index As Integer)
Dim cadMen As String
Dim SQL As String

    'Quitar espacios en blanco
    If Not PerderFocoGnralLineas(txtAux2(Index), ModificaLineas) Then Exit Sub

    Select Case Index
        Case 2 ' codigo de gasto
            If PonerFormatoEntero(txtAux2(Index)) Then
                Text2(6) = DevuelveDesdeBDNew(cAgro, "rconcepgastonom", "nomgasto", "codgasto", txtAux2(Index), "N")
                If Text2(6).Text = "" Then
                    cadMen = "No existe el Concepto de Gasto: " & txtAux2(Index).Text & vbCrLf
                    cadMen = cadMen & "¿Desea crearlo?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmGas = New frmManCGastosNom
                        frmGas.DatosADevolverBusqueda = "0|1|"
                        frmGas.NuevoCodigo = txtAux2(Index).Text
                        TerminaBloquear
                        frmGas.Show vbModal
                        Set frmGas = Nothing
                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
                        PonerFoco txtAux2(Index)
                    Else
                        txtAux2(Index).Text = ""
                    End If
                    PonerFoco txtAux2(Index)
                End If
            Else
                Text2(6).Text = ""
            End If


        Case 3 ' importe
            If txtAux2(Index) <> "" Then
                If PonerFormatoDecimal(txtAux2(Index), 3) Then cmdAceptar.SetFocus
            End If
    
    End Select
    
End Sub




Private Function Eliminar() As Boolean
Dim SQL As String, LEtra As String, Sql2 As String
Dim B As Boolean
Dim vTipoMov As CTiposMov
Dim Mens As String

    On Error GoTo FinEliminar

    If Data1.Recordset.EOF Then Exit Function
        
    conn.BeginTrans
        
    'Eliminar en tablas de cabecera de albaran
    '------------------------------------------
    SQL = " " & ObtenerWhereCP(True)
    
    'Lineas de clasificacion (rpartes_gastos)
    conn.Execute "Delete from rpartes_gastos " & SQL
    
    'Lineas de incidencias de notas (rpartes_trabajador)
    conn.Execute "Delete from rpartes_trabajador " & SQL
    
    'Lineas de entradas (rpartes_variedad)
    conn.Execute "Delete from rpartes_variedad " & SQL

    'Cabecera de partes (rpartes)
    conn.Execute "Delete from " & NombreTabla & SQL
    
    'Decrementar contador si borramos el ult. palet
    Set vTipoMov = New CTiposMov
    vTipoMov.DevolverContador "PAC", Val(Text1(0).Text)
    Set vTipoMov = Nothing
    
    B = True
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

Private Function EliminarLinea(Indice As Integer) As Boolean
Dim SQL As String, LEtra As String
Dim B As Boolean
Dim vTipoMov As CTiposMov
Dim Mens As String

    On Error GoTo FinEliminar

    If Data3.Recordset.EOF Then Exit Function
        
    conn.BeginTrans
        
    Mens = ""

    Select Case Indice
        Case 0
            'Lineas de gastos individuales del trabajador (rpartes_trabajador)
            SQL = " where nroparte = " & Data3.Recordset.Fields(0)
            SQL = SQL & " and numlinea  = " & Data3.Recordset.Fields(1)
            
            conn.Execute "Delete from rpartes_trabajador " & SQL
        Case 1
            'Eliminar en tablas de rpartes_gastos
            '------------------------------------------
            SQL = " where nroparte = " & Data3.Recordset.Fields(0)
            SQL = SQL & " and numlinea = " & Data3.Recordset.Fields(1)
            
            conn.Execute "Delete from rpartes_gastos " & SQL
    End Select
    
FinEliminar:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Eliminar linea Parte ", Err.Description & " " & Mens
        conn.RollbackTrans
        EliminarLinea = False
    Else
        conn.CommitTrans
        EliminarLinea = True
    End If
End Function

Private Sub LimpiarDataGrids()
'Pone los Grids sin datos, apuntando a ningún registro
On Error Resume Next

    CargaGrid DataGrid2, Data3, False 'entradas e incidencias
    CargaGrid DataGrid1, Data2, False
    CargaGrid DataGrid3, Me.Adoaux(0), False 'clasificacion
    CargaGrid DataGrid4, Data4, False 'clasificacion
    
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
Dim SQL As String

    On Error Resume Next
    
    SQL = " nroparte= " & Text1(0).Text 'Data1.Recordset!numalbar 'Text1(0).Text
    If conWhere Then SQL = " WHERE " & SQL
    ObtenerWhereCP = SQL
    
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
Dim SQL As String
    
    Select Case Opcion
    Case 1  ' rpartes_variedad
        SQL = "SELECT rpartes_variedad.nroparte,rpartes_variedad.numlinea,rpartes_variedad.numnotac, "
        If vParamAplic.Cooperativa = 16 Then
            SQL = SQL & " rpartes_variedad.codvarie, variedades.nomvarie, rpartes_variedad.numcajon, rpartes_variedad.kilosrec "
        Else
            SQL = SQL & " rpartes_variedad.codvarie, variedades.nomvarie, rpartes_variedad.horastra, rpartes_variedad.kilosrec "
        End If
        SQL = SQL & " FROM rpartes_variedad, variedades WHERE rpartes_variedad.codvarie = variedades.codvarie "
        
    Case 2  ' rpartes_trabajador
        SQL = "SELECT rpartes_trabajador.nroparte, rpartes_trabajador.numlinea, rpartes_trabajador.codtraba, straba.nomtraba, rpartes_trabajador.codgasto, "
        SQL = SQL & " rconcepgastonom.nomgasto, rpartes_trabajador.codvarie, variedades.nomvarie, rpartes_trabajador.numcajas, rpartes_trabajador.kilosrec, rpartes_trabajador.horastra, rpartes_trabajador.importe, rpartes_trabajador.modificado "
        SQL = SQL & " FROM ((rpartes_trabajador INNER JOIN straba ON rpartes_trabajador.codtraba = straba.codtraba) "
        SQL = SQL & " LEFT JOIN rconcepgastonom ON rpartes_trabajador.codgasto = rconcepgastonom.codgasto) "
        SQL = SQL & " LEFT JOIN variedades ON rpartes_trabajador.codvarie = variedades.codvarie"
        SQL = SQL & " WHERE (1=1)"
        
    Case 3  ' rpartes_gastos
        SQL = "SELECT rpartes_gastos.nroparte, rpartes_gastos.numlinea, rpartes_gastos.codgasto, rconcepgastonom.nomgasto, rpartes_gastos.importe "
        SQL = SQL & " FROM rpartes_gastos, rconcepgastonom "
        SQL = SQL & " WHERE rpartes_gastos.codgasto = rconcepgastonom.codgasto "
        
    Case 4  ' resumen
        SQL = "select rpartes_trabajador.codvarie, variedades.nomvarie, sum(numcajas) cajas, sum(kilosrec) kilos , sum(importe) importe"
        SQL = SQL & "  from rpartes_trabajador, variedades "
        SQL = SQL & " where rpartes_trabajador.codvarie = variedades.codvarie "
        
        
    End Select
    
    If enlaza Then
        SQL = SQL & " and " & ObtenerWhereCP(False)
    Else
        SQL = SQL & " and nroparte = -1"
    End If
    Select Case Opcion
        Case 1
            SQL = SQL & " ORDER BY nroparte, numlinea "
        Case 2
            SQL = SQL & " ORDER BY nroparte, codtraba, codvarie "
        Case 3
            SQL = SQL & " ORDER BY nroparte, numlinea "
        Case 4
            SQL = SQL & " group by 1, 2 "
            SQL = SQL & " ORDER BY 1 "
            
    End Select
    MontaSQLCarga = SQL
End Function


Private Sub PonerModoOpcionesMenu(Modo As Byte)
'Activas unas Opciones de Menu y Toolbar según el modo en que estemos
Dim B As Boolean, bAux As Boolean
Dim i As Integer

    B = ((Modo = 2) Or (Modo = 0)) And (NroParte = "") 'Or (Modo = 5 And ModificaLineas = 0)
    'Buscar
    Toolbar1.Buttons(5).Enabled = B
    Me.mnBuscar.Enabled = B
    'Vore Tots
    Toolbar1.Buttons(6).Enabled = B
    Me.mnVerTodos.Enabled = B
    'Añadir
    Toolbar1.Buttons(1).Enabled = B
    Me.mnModificar.Enabled = B
    
    
    B = (Modo = 2 And Data1.Recordset.RecordCount > 0) And (NroParte = "") 'And Not (Check1(0).Value = 1)
    'Modificar
    Toolbar1.Buttons(2).Enabled = B
    Me.mnModificar.Enabled = B
    'eliminar
    Toolbar1.Buttons(3).Enabled = B
    Me.mnEliminar.Enabled = B
    
    'Traer entradas
    Toolbar5.Buttons(1).Enabled = (Modo = 2) And Data1.Recordset.RecordCount > 0
    Me.mnTraerEntradas.Enabled = (Modo = 2) And Data1.Recordset.RecordCount > 0
    'Recalcular Importes
    Toolbar5.Buttons(2).Enabled = (Modo = 2) And Data1.Recordset.RecordCount > 0 And Data2.Recordset.RecordCount > 0
    Me.mnRecalcularImportes.Enabled = (Modo = 2) And Data1.Recordset.RecordCount > 0 And Data2.Recordset.RecordCount > 0



    ' *** si n'hi han llínies que tenen grids (en o sense tab) ***
'++monica: si insertamos lo he quitado
'    b = (Modo = 3 Or Modo = 4 Or Modo = 2) And Not DeConsulta
    B = (Modo = 2) And NroParte = "" 'And Not Check1(0).Value = 1
    For i = 0 To ToolAux.Count - 1
        ToolAux(i).Buttons(1).Enabled = B
        
        If B Then
            Select Case i
              Case 0
                bAux = (B And Me.Data3.Recordset.RecordCount > 0)
              Case 1
                bAux = (B And Me.Adoaux(0).Recordset.RecordCount > 0)
              Case 2
                bAux = (B And Me.Data2.Recordset.RecordCount > 0)
            End Select
        End If
        ToolAux(i).Buttons(2).Enabled = bAux
        ToolAux(i).Buttons(3).Enabled = bAux
        '[Monica]10/10/2016: nuevo boton para insertar los cajones
        If i = 0 Then ToolAux(i).Buttons(4).Enabled = bAux
    Next i


End Sub


Private Sub BotonImprimir()
Dim cadFormula As String
Dim CadParam As String
Dim numParam As Byte
Dim cadSelect As String 'select para insertar en tabla temporal
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String
Dim NroCopias As Integer

    If Text1(0).Text = "" Then
        MsgBox "Debe seleccionar un Parte para Imprimir.", vbInformation
        Exit Sub
    End If
    
    cadFormula = ""
    CadParam = ""
    cadSelect = ""
    numParam = 0
    
    '===================================================
    '============ PARAMETROS ===========================
    indRPT = 22 'Impresion de Albaran de clasificacion
    If Not PonerParamRPT(indRPT, CadParam, numParam, nomDocu) Then Exit Sub
      
    'Nombre fichero .rpt a Imprimir
    frmImprimir.NombreRPT = nomDocu
        
    '===================================================
    '================= FORMULA =========================
    'Cadena para seleccion Nº de albaran
    '---------------------------------------------------
    If Text1(0).Text <> "" Then
        'Nº palet
        devuelve = "{" & NombreTabla & ".numalbar}=" & Val(Text1(0).Text)
        If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Sub
        devuelve = "numalbar = " & Val(Text1(0).Text)
        If Not AnyadirAFormula(cadSelect, devuelve) Then Exit Sub
    End If
    
    CadParam = CadParam & "pDuplicado=1|"
    numParam = numParam + 1
    
    
    If Not HayRegParaInforme(NombreTabla, cadSelect) Then Exit Sub
     
    With frmImprimir
            .FormulaSeleccion = cadFormula
            .OtrosParametros = CadParam
            .NumeroParametros = numParam
            .SoloImprimir = False
            .EnvioEMail = False
            .Opcion = 0
            .Titulo = "Impresión de Parte de Campo"
            .ConSubInforme = True
            .Show vbModal
    End With

End Sub

Private Sub TxtAux3_GotFocus(Index As Integer)
    ConseguirFoco txtAux3(Index), Modo
End Sub

Private Sub TxtAux3_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Index <> 0 And KeyCode <> 38 Then KEYdown KeyCode
End Sub

Private Sub TxtAux3_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub TxtAux3_LostFocus(Index As Integer)
Dim cadMen As String
Dim Precio As Currency
Dim ImporteTrab As Currency

    'Quitar espacios en blanco
    If Not PerderFocoGnralLineas(txtAux3(Index), ModificaLineas) Then Exit Sub

    Select Case Index
        Case 1 ' codigo de trabajador
            If txtAux3(Index) <> "" Then
                Text2(3) = DevuelveDesdeBDNew(cAgro, "straba", "nomtraba", "codtraba", txtAux3(Index), "N")
                If Text2(3).Text = "" Then
                    cadMen = "No existe el Trabajador: " & txtAux3(Index).Text & vbCrLf
                    cadMen = cadMen & "¿Desea crearlo?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmTra = New frmManTraba
                        frmTra.DatosADevolverBusqueda = "0|1|"
                        TerminaBloquear
                        frmTra.Show vbModal
                        Set frmTra = Nothing
                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
                        PonerFoco txtAux3(Index)
                    Else
                        txtAux3(Index).Text = ""
                    End If
                    PonerFoco txtAux3(Index)
                Else
                    '[Monica]28/10/2015: si el trabajador esta de baja no le dejamos salir
                    If TrabajadorDeBaja(txtAux3(Index).Text, Text1(1).Text) Then
                        MsgBox "El trabajador está de baja. Revise.", vbExclamation
                        txtAux3(Index).Text = ""
                        PonerFoco txtAux3(Index)
                    End If
                End If
            Else
                txtAux3(Index).Text = ""
            End If
        
    
    
        Case 2 ' codigo de gasto
            If txtAux3(Index) <> "" Then
                Text2(4) = DevuelveDesdeBDNew(cAgro, "rconcepgastonom", "nomgasto", "codgasto", txtAux3(Index), "N")
                If Text2(4).Text = "" Then
                    cadMen = "No existe el Concepto de Gasto: " & txtAux3(Index).Text & vbCrLf
                    cadMen = cadMen & "¿Desea crearlo?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmGas = New frmManCGastosNom
                        frmGas.DatosADevolverBusqueda = "0|1|"
                        frmGas.NuevoCodigo = txtAux3(3).Text
                        TerminaBloquear
                        frmGas.Show vbModal
                        Set frmGas = Nothing
                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
                        PonerFoco txtAux3(3)
                    Else
                        txtAux3(Index).Text = ""
                    End If
                    PonerFoco txtAux3(Index)
                End If
            Else
                txtAux3(Index).Text = ""
            End If

        Case 3 '[Monica]01/03/2012:
               '    si me modifican los kilos he de calcular el importe=kilos*precio
            If Combo1(0).ListIndex = 0 Then
                '[Monica]17/06/2013: añadida la condicion de si es un parte a destajo
                If vParamAplic.Cooperativa = 0 Or vParamAplic.Cooperativa = 18 Then
                    Precio = DevuelveValor("select eursegsoc from variedades where codvarie = " & DBSet(txtAux3(5).Text, "N"))
                Else
                    Precio = DevuelveValor("select eurdesta from variedades where codvarie = " & DBSet(txtAux3(5).Text, "N"))
                End If
                ImporteTrab = Round2(ComprobarCero(txtAux3(Index).Text) * Precio, 2)
                txtAux3(4).Text = Format(ImporteTrab, "###,##0.00")
            End If
            
        Case 8 ' horas
            'solo en el caso de que sea un parte a horas
            PonerFormatoDecimal txtAux3(Index), 10
            If Combo1(0).ListIndex = 1 Then
                Precio = DevuelveValor("select impsalar  from salarios inner join straba on salarios.codcateg = straba.codcateg where straba.codtraba = " & DBSet(txtAux3(1).Text, "N"))
                ImporteTrab = Round2(HorasDecimal(ComprobarCero(txtAux3(Index).Text)) * Precio, 2)
                txtAux3(4).Text = Format(ImporteTrab, "###,##0.00")
            End If
        
        Case 4 ' importe
            If txtAux3(Index) <> "" Then
                If PonerFormatoDecimal(txtAux3(Index), 3) Then cmdAceptar.SetFocus
            End If
    End Select
    
    
End Sub


Private Function ModificaCabecera() As Boolean
Dim B As Boolean
Dim MenError As String

    On Error GoTo EModificarCab

    conn.BeginTrans
    
    
    B = ModificaDesdeFormulario2(Me, 2, "Frame2")

EModificarCab:
    If Err.Number <> 0 Or Not B Then
        MenError = "Modificando Parte." & vbCrLf & "----------------------------" & vbCrLf & MenError
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
Dim SQL As String

    On Error GoTo EInsertarCab
    
    Set vTipoMov = New CTiposMov
    If vTipoMov.Leer(CodTipoMov) Then
        Text1(0).Text = vTipoMov.ConseguirContador(CodTipoMov)
        SQL = CadenaInsertarDesdeForm(Me)
        If SQL <> "" Then
            If InsertarOferta(SQL, vTipoMov) Then
                CadenaConsulta = "Select * from " & NombreTabla & ObtenerWhereCP(True) & Ordenacion
                PonerCadenaBusqueda
                PonerModo 2
                'Ponerse en Modo Insertar Lineas
'                BotonMtoLineas 0, "Variedades"
'                BotonAnyadirLinea
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

    On Error GoTo EInsertarOferta
    
    bol = True
    
    cambiaSQL = False
    'Comprobar si mientras tanto se incremento el contador de Albaranes
    'para ello vemos si existe una oferta con ese contador y si existe la incrementamos
    Do
        devuelve = DevuelveDesdeBDNew(cAgro, NombreTabla, "nroparte", "nroparte", Text1(0).Text, "N")
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
    
    MenError = "Error al actualizar el contador del Parte."
    vTipoMov.IncrementarContador (CodTipoMov)
    
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


Private Sub InsertarLinea(Index As Integer)
'Inserta registre en les taules de Llínies
Dim nomframe As String
Dim B As Boolean

    On Error Resume Next

    ' *** posa els noms del frames, tant si son de grid com si no ***
    Select Case Index
        Case 0: nomframe = "Frame3" 'gastos individuales
        Case 1: nomframe = "FrameAux0" 'gastos generales
        Case 2: nomframe = "Frame4" ' variedades
    End Select
    ' ***************************************************************
    
    If DatosOkLlin(nomframe) Then
        TerminaBloquear
        If InsertarDesdeForm2(Me, 2, nomframe) Then
            ' *** si n'hi ha que fer alguna cosa abas d'insertar
            If NumTabMto = 0 Then
'yo                'si ha seleccionat "cuenta principal", actualise totes les atres a "no"
'                If (cmbAux(10).ItemData(cmbAux(10).ListIndex) = 1) Then
'                    ActualisaCtaprpal (txtaux(2).Text)
'                End If
            End If
            ' *************************************************
            B = BloqueaRegistro("rpartes", "nroparte = " & Data1.Recordset!NroParte)
            Select Case Index
                Case 0  ' *** els index de les llinies en grid (en o sense tab) ***
                    CargaGrid DataGrid2, Data3, True
                    If B Then BotonAnyadirLinea NumTabMto
'                LLamaLineas NumTabMto, 0
                Case 1
                    CargaGrid DataGrid3, Adoaux(0), True
                    If B Then BotonAnyadirLinea NumTabMto
                Case 2
                    CargaGrid DataGrid1, Data2, True
                    If B Then BotonAnyadirLinea NumTabMto
                
            End Select
'            SSTab1.Tab = NumTabMto
        End If
    End If
End Sub

Private Sub BotonAnyadirLinea(Index As Integer)
Dim NumF As String
Dim vWhere As String, vtabla As String
Dim anc As Single
Dim i As Integer
    
    ModificaLineas = 1 'Posem Modo Afegir Llínia
    
    If (Modo = 3) Or (Modo = 4) Then 'Insertar o Modificar Capçalera
        cmdAceptar_Click
        If ModificaLineas = 0 Then Exit Sub
    End If
       
'    NumTabMto = Index
'    If Index = 2 Then NumTabMto = 3
    
    PonerModo 5, Index
    
    ' *** bloquejar la clau primaria de la capçalera ***
    BloquearTxt Text1(0), True
    ' **************************************************

    ' *** posar el nom del les distintes taules de llínies ***
    Select Case NumTabMto
        Case 0: vtabla = "rpartes_trabajador" ' gastos individuales
        Case 1: vtabla = "rpartes_gastos" ' gastos  generales
        Case 2: vtabla = "rpartes_variedad" ' notas/variedad kilos
    End Select
    ' ********************************************************
    
    vWhere = ObtenerWhereCab(False)
    
    Select Case NumTabMto
        Case 0 ' *** pose els index dels tabs de llínies que tenen datagrid ***
            ' *** canviar la clau primaria de les llínies,
            'pasar a "" si no volem que mos sugerixca res a l'afegir ***
            NumF = SugerirCodigoSiguienteStr(vtabla, "numlinea", vWhere)
            ' ***************************************************************

            AnyadirLinea DataGrid2, Data3
    
            anc = DataGrid2.Top
            If DataGrid2.Row < 0 Then
                anc = anc + 220
            Else
                anc = anc + DataGrid2.RowTop(DataGrid2.Row) + 5
            End If
            
            LLamaLineas ModificaLineas, anc, "DataGrid2"
        
            LimpiarCamposLin "Frame3"
            
            txtAux3(0).Text = Text1(0).Text 'nroparte
            txtAux3(6).Text = NumF ' numero de linea
            txtAux3(1).Text = "" ' codtraba
            txtAux3(2).Text = "" ' codconce
            txtAux3(3).Text = "0" ' kilos
            txtAux3(4).Text = "" ' importe
            txtAux3(5).Text = "" ' codvarie
            Text2(3).Text = "" ' nomtraba
            Text2(4).Text = "" ' nomconce
            Text2(0).Text = "" ' nomvarie
            txtAux3(7).Text = "0" ' modificado
            txtAux3(8).Text = "0" ' modificado
            txtAux3(9).Text = "0" ' modificado
            BloquearTxt txtAux3(1), False
'                    BloquearTxt txtaux(12), False
            PonerFoco txtAux3(1)
                    
                    
        Case 1 ' *** pose els index dels tabs de llínies que tenen datagrid ***
            ' *** canviar la clau primaria de les llínies,
            'pasar a "" si no volem que mos sugerixca res a l'afegir ***
            NumF = SugerirCodigoSiguienteStr(vtabla, "numlinea", vWhere)
            ' ***************************************************************

            AnyadirLinea DataGrid3, Adoaux(0)
    
            anc = DataGrid3.Top
            If DataGrid3.Row < 0 Then
                anc = anc + 220
            Else
                anc = anc + DataGrid3.RowTop(DataGrid3.Row) + 5
            End If
            
            LLamaLineas ModificaLineas, anc, "DataGrid3"
        
            LimpiarCamposLin "FrameAux0"
            
            txtAux2(0).Text = Text1(0).Text 'nroparte
            txtAux2(1).Text = NumF
            txtAux2(2).Text = ""
            Text2(6).Text = ""
            txtAux2(3).Text = ""
            
            BloquearTxt txtAux2(3), False
            PonerFoco txtAux2(2)
        
        Case 2 ' *** pose els index dels tabs de llínies que tenen datagrid ***
            ' *** canviar la clau primaria de les llínies,
            'pasar a "" si no volem que mos sugerixca res a l'afegir ***
            NumF = SugerirCodigoSiguienteStr(vtabla, "numlinea", vWhere)
            ' ***************************************************************

            AnyadirLinea DataGrid1, Data2
    
            anc = DataGrid1.Top
            If DataGrid1.Row < 0 Then
                anc = anc + 220
            Else
                anc = anc + DataGrid1.RowTop(DataGrid1.Row) + 5
            End If
            
            LLamaLineas ModificaLineas, anc, "DataGrid1"
        
            LimpiarCamposLin "Frame4"
            
            txtAux(0).Text = Text1(0).Text 'nroparte
            txtAux(1).Text = NumF
            txtAux(2).Text = ""
            txtAux(3).Text = ""
            txtAux(4).Text = ""
            txtAux(5).Text = ""
            txtAux(6).Text = 0
            
            BloquearTxt txtAux(2), False
            BloquearTxt txtAux(3), True
            BloquearTxt txtAux(5), False
            
            PonerFoco txtAux(2)
                    
    End Select
End Sub

Private Sub LimpiarCamposLin(FrameAux As String)
    On Error Resume Next
    
    LimpiarLin Me, FrameAux  'Mètode general: Neteja els controls TextBox
    lblIndicador.Caption = ""

    If Err.Number <> 0 Then Err.Clear
End Sub

Private Sub ModificarLinea()
'Modifica registre en les taules de Llínies
Dim nomframe As String
Dim V As Integer
Dim cad As String
    On Error Resume Next

    ' *** posa els noms del frames, tant si son de grid com si no ***
    Select Case NumTabMto
        Case 0:
            nomframe = "Frame3"    'rpartes_trabajador
            
            '[Monica]01/03/2012: si me han modificado los kilos o importe lo pongo que está modificado
            '                    solo en el caso de lineas que sean de variedad (automaticas)
            txtAux3(7).Text = "0"
            If ComprobarCero(txtAux3(5).Text) <> 0 And (CLng(txtAux3(3).Text) <> KilosAnt Or CCur(ComprobarCero(txtAux3(4).Text)) <> ImporteAnt) Then
                txtAux3(7).Text = "1"
            End If
            
        Case 1:
            nomframe = "FrameAux0" 'clasificacion
        Case 2:
            nomframe = "Frame4" 'rpartes_variedad
    End Select
    ' **************************************************************

    If DatosOkLlin(nomframe) Then
        TerminaBloquear
        If ModificaDesdeFormulario2(Me, 2, nomframe) Then
            ' *** si cal que fer alguna cosa abas d'insertar ***
            If NumTabMto = 0 Then
'yo                'si ha seleccionat "cuenta principal", actualise totes les atres a "no"
'                If (cmbAux(10).ItemData(cmbAux(10).ListIndex) = 1) Then
'                    ActualisaCtaprpal (txtaux(2).Text)
'                End If
            End If
            ' ******************************************************
'            If BLOQUEADesdeFormulario2(Me, Data1, 1) Then BotonModificar
            ModificaLineas = 0

            Select Case NumTabMto
                Case 1

                    V = Adoaux(0).Recordset.Fields(1) 'el 2 es el nº de llinia
                    CargaGrid DataGrid3, Adoaux(0), True

                    ' *** si n'hi han tabs ***
'                    SSTab1.Tab = 1

                    DataGrid3.SetFocus
                    Adoaux(0).Recordset.Find (Adoaux(0).Recordset.Fields(1).Name & " =" & V)

                    LLamaLineas ModificaLineas, 0, "DataGrid3"
                Case 0
                    V = Data3.Recordset.Fields(1) 'el 2 es el nº de llinia
                    CargaGrid DataGrid2, Data3, True

                    ' *** si n'hi han tabs ***
'                    SSTab1.Tab = 1

                    DataGrid2.SetFocus
                    Data3.Recordset.Find (Data3.Recordset.Fields(1).Name & " =" & V)

                    LLamaLineas ModificaLineas, 0, "DataGrid2"
                Case 2
                    V = Data2.Recordset.Fields(1) 'el 2 es el nº de llinia
                    CargaGrid DataGrid1, Data3, True

                    ' *** si n'hi han tabs ***
'                    SSTab1.Tab = 1

                    DataGrid1.SetFocus
                    Data2.Recordset.Find (Data2.Recordset.Fields(1).Name & " =" & V)

                    LLamaLineas ModificaLineas, 0, "DataGrid1"
            End Select
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

    DatosOkLlin = True
    
    On Error GoTo EDatosOKLlin

    Mens = ""
    DatosOkLlin = False

    B = CompForm2(Me, 2, nomframe) 'Comprovar formato datos ok
    If Not B Then Exit Function
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
    DatosOkLlin = B

EDatosOKLlin:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function

Private Function ObtenerWhereCab(conW As Boolean) As String
Dim vWhere As String
    
    vWhere = ""
    If conW Then vWhere = " WHERE "
    ' *** canviar-ho per la clau primaria de la capçalera ***
    vWhere = vWhere & " nroparte= " & Val(Text1(0).Text)
    ' *******************************************************
    
    ObtenerWhereCab = vWhere
End Function

' ***** si n'hi han varios nivells de tabs *****
Private Sub SituarTab(numTab As Integer)
'    On Error Resume Next
'
'    If numTab = 0 Or numTab = 1 Or numTab = 2 Or numTab = 3 Then
'        SSTab1.Tab = 2
'    ElseIf numTab = 4 Then
'        SSTab1.Tab = 2
'    End If
'
'    If Err.Number <> 0 Then Err.Clear
End Sub
' **********************************************


Private Function SePuedeModificarLinea(Parte As String, Linea As String) As Boolean
Dim SQL As String
Dim Valor As Variant
Dim Rs As ADODB.Recordset

    SePuedeModificarLinea = False
    
    SQL = "select codvarie, codgasto from rpartes_trabajador where nroparte = " & DBSet(Parte, "N")
    SQL = SQL & " and numlinea = " & DBSet(Linea, "N")
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    If Not Rs.EOF Then
        ' antes RS.Fields(0).Value > 0
        If IsNull(Rs.Fields(0).Value) Or DBLet(Rs.Fields(0).Value, "N") = 0 Then
            If Not IsNull(Rs.Fields(1).Value) Then SePuedeModificarLinea = True  'Solo es para saber que hay registros que mostrar
        End If
    End If
    Rs.Close
    Set Rs = Nothing
    
End Function


Private Function EstaPartePagado(Numparte) As Boolean
Dim SQL As String

    SQL = "select count(*) from horas where nroparte = " & DBLet(Numparte, "N")
    SQL = SQL & " and fecharec is not null"

    EstaPartePagado = (TotalRegistros(SQL) <> 0)
    
End Function


Private Function EstaParteenHoras(Numparte) As Boolean
Dim SQL As String

    SQL = "select count(*) from horas where nroparte = " & DBLet(Numparte, "N")

    EstaParteenHoras = (TotalRegistros(SQL) <> 0)
    
End Function

' ********* si n'hi han combos a la capçalera ************
Private Sub CargaCombo()
Dim Ini As Integer
Dim Fin As Integer
Dim i As Integer

    ' *** neteje els combos, els pose valor i seleccione el valor per defecte ***
    For i = 0 To Combo1.Count - 1
        Combo1(i).Clear
    Next i
    
    Combo1(0).AddItem "a Destajo"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 0
    Combo1(0).AddItem "por Horas"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 1
    
    
End Sub


'[Monica]14/06/2013: nueva funcion para calcular importes y kilos segun las horas por trabajador pidiendo que trabajadores
'                    y el nro de horas que ha trabajado en el parte

Private Function RecalcularImportesHoras() As Boolean
Dim cad As String
Dim SQL As String
Dim Sql2 As String
Dim Sql4 As String
Dim NroHoras As Currency
Dim PrecioHora As Currency
Dim vImporte As Currency
Dim vHoras As Currency
Dim tHoras As Currency
Dim NroTrabajadores As Integer
Dim i As Integer
Dim Rs2 As ADODB.Recordset
Dim Rs As ADODB.Recordset

Dim TotalKilos As Long
Dim TotalHoras As Currency
Dim HorasVarie As Currency
Dim THorasVarie As Currency

Dim HorasT As Currency
Dim HorasTrab As Currency
Dim ImporteTrab As Currency

Dim KilosInicio As Long
Dim KilosRec As Long

Dim SqlPre As String
Dim Precio As Currency
Dim NumF As Long
Dim KilosTrab As Long


    RecalcularImportesHoras = True

    
    cad = "Se va a proceder a recalcular los importes por trabajador según las horas. "
'    cad = cad & "los gastos generales introducidos. "
    cad = cad & vbCrLf & "         ¿ Desea Continuar ? "
    If MsgBox(cad, vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
        'mostramos cuales son los trabajadores de la cuadrilla que han de seleccionar para hacer el reparto
        
        SQL = "delete from rpartes_trabajador where nroparte = " & Data1.Recordset.Fields(0).Value
        conn.Execute SQL

        Sql2 = "delete from tmpliquidacion where codusu = " & vUsu.Codigo
        conn.Execute Sql2
        
        
        Set frmTMP = New frmManPartesTMP
        frmTMP.ParamVariedad = Text1(0).Text
        frmTMP.Show vbModal
        Set frmTMP = Nothing
    
        'hacemos el reparto de kilos segun las horas que hay introducidas
        SQL = "select count(*) from tmpliquidacion where codusu = " & vUsu.Codigo
        If TotalRegistros(SQL) = 0 Then
            MsgBox "No hay trabajadores a repartir. No se ha realizado el proceso.", vbExclamation
            RecalcularImportesHoras = True
            Exit Function
        Else
            SQL = "select codvarie, sum(kilosrec) as kilosrec from rpartes_variedad where nroparte = " & Data1.Recordset.Fields(0).Value
            SQL = SQL & " group by 1 order by 1 "
            
            TotalKilos = DevuelveValor("select sum(kilosrec) from (" & SQL & ") aaaaa ")
            
            TotalHoras = DevuelveValor("select sum(gastos) from tmpliquidacion where codusu = " & vUsu.Codigo)
            TotalHoras = HorasDecimal(CStr(Horas(CStr(TotalHoras))))
            
            Set Rs = New ADODB.Recordset
            Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            THorasVarie = 0
            While Not Rs.EOF
                HorasVarie = Round2(DBLet(Rs!KilosRec, "N") * TotalHoras / TotalKilos, 2)
'                If TotalKilos <> DBLet(RS!KilosRec, "N") Then
'                    HorasVarie = HorasDecimal(CStr(HorasVarie))
'                End If
'                THorasVarie = THorasVarie + HorasVarie
                
                Sql4 = "select codvarie, gastos from tmpliquidacion where codusu = " & DBSet(vUsu.Codigo, "N")
                Sql4 = Sql4 & " order by 1 "

                Set Rs2 = New ADODB.Recordset
                Rs2.Open Sql4, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

                While Not Rs2.EOF
                    HorasT = DBLet(Rs2!Gastos, "N") 'horas totales del trabajador
                    HorasT = HorasDecimal(CStr(HorasT))
                    HorasTrab = 0
                    If TotalHoras <> 0 Then
                        HorasTrab = Round2(HorasT * HorasVarie / TotalHoras, 2)
'                        HorasTrab = DecimalHoras(CStr(HorasTrab))
                    End If
                    
                    KilosTrab = 0
                    If HorasVarie <> 0 Then
                        KilosTrab = Round2(HorasTrab * DBLet(Rs!KilosRec, "N") / HorasVarie, 0)
                    End If
                   
                    SqlPre = "select impsalar from straba inner join salarios on straba.codcateg = salarios.codcateg "
                    SqlPre = SqlPre & " where straba.codtraba = " & DBSet(Rs2!codvarie, "N") ' trabajador
                    
                    Precio = DevuelveValor(SqlPre)
                    
                    ImporteTrab = Round2(HorasTrab * Precio, 2)

                    '[Monica]01/03/2012: no insertamos los modificados
                    Sql4 = "select count(*) from rpartes_trabajador where codtraba = " & DBSet(Rs2!codvarie, "N") ' codigo de trabajador
                    Sql4 = Sql4 & " and codvarie = " & DBSet(Rs!codvarie, "N")
                    Sql4 = Sql4 & " and nroparte = " & Data1.Recordset.Fields(0).Value
                    If TotalRegistros(Sql4) = 0 Then

                        NumF = SugerirCodigoSiguienteStr("rpartes_trabajador", "numlinea", "nroparte = " & Data1.Recordset.Fields(0))

                        Sql2 = "insert into rpartes_trabajador (nroparte, numlinea, codtraba, codvarie, horastra, kilosrec, importe, automatico) values "
                        Sql2 = Sql2 & "(" & Data1.Recordset.Fields(0).Value & "," & DBSet(NumF, "N") & ","
                        ' este es el trabajador
                        Sql2 = Sql2 & DBSet(Rs2!codvarie, "N") & ","
                        ' esta la variedad
                        Sql2 = Sql2 & DBSet(Rs!codvarie, "N") & "," & DBSet(DecimalHoras(HorasTrab), "N") & ","
                        Sql2 = Sql2 & DBSet(KilosTrab, "N") & "," & DBSet(ImporteTrab, "N") & ",1)"
'Decimalhoras(ImporteSinFormato(CStr(Format(HorasTrab, "##,##0.00"))))
                        conn.Execute Sql2

                    End If

                    Rs2.MoveNext
                Wend
                
                Set Rs2 = Nothing
                
                Rs.MoveNext
            
            Wend
            Set Rs = Nothing
                
        End If
    End If
    Exit Function
    
eRecalcularImportes:
    RecalcularImportesHoras = False
    MuestraError Err.Number, "Recalcular Importes", Err.Description
End Function


Private Sub ToolbarAyuda_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1
            LanzaVisorMimeDocumento Me.hWnd, DireccionAyuda & IdPrograma & ".html"
    End Select
End Sub
