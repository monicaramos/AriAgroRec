VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmManClasAuto 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Clasificación Automática"
   ClientHeight    =   7800
   ClientLeft      =   45
   ClientTop       =   30
   ClientWidth     =   13245
   Icon            =   "frmManClasAuto.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7800
   ScaleWidth      =   13245
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameBotonGnral 
      Height          =   705
      Left            =   30
      TabIndex        =   44
      Top             =   0
      Width           =   3585
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   210
         TabIndex        =   45
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
      Left            =   3660
      TabIndex        =   42
      Top             =   0
      Width           =   1305
      Begin MSComctlLib.Toolbar Toolbar2 
         Height          =   330
         Left            =   210
         TabIndex        =   43
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
               Enabled         =   0   'False
               Object.ToolTipText     =   "Traspaso desde el Calibrador"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Actualizar Entradas"
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
      Index           =   0
      Left            =   7620
      TabIndex        =   41
      Top             =   180
      Visible         =   0   'False
      Width           =   1605
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
      Height          =   360
      Index           =   3
      Left            =   4620
      Locked          =   -1  'True
      MaxLength       =   7
      TabIndex        =   39
      Tag             =   "Kilos Neto|N|S|||rclasifauto_clasif|kiloscal|###,##0.00||"
      Text            =   "neto"
      Top             =   7230
      Width           =   1400
   End
   Begin VB.Frame Frame2 
      Height          =   3105
      Index           =   0
      Left            =   30
      TabIndex        =   13
      Top             =   750
      Width           =   13205
      Begin VB.TextBox Text1 
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
         Height          =   360
         Index           =   8
         Left            =   10410
         MaxLength       =   7
         TabIndex        =   7
         Tag             =   "Kilos Destrio|N|N|||rclasifauto|kilosdes|###,##0||"
         Text            =   "destrio"
         Top             =   570
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.TextBox Text1 
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
         Height          =   360
         Index           =   7
         Left            =   9690
         MaxLength       =   7
         TabIndex        =   6
         Tag             =   "Kilos Pequeño|N|N|||rclasifauto|kilospeq|###,##0||"
         Text            =   "pequeño"
         Top             =   570
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.TextBox Text1 
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
         Height          =   360
         Index           =   6
         Left            =   8910
         MaxLength       =   7
         TabIndex        =   5
         Tag             =   "Kilos Podridos|N|N|||rclasifauto|kilospod|###,##0||"
         Text            =   "podrido"
         Top             =   570
         Visible         =   0   'False
         Width           =   720
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
         Index           =   4
         Left            =   6450
         MaskColor       =   &H00000000&
         TabIndex        =   37
         ToolTipText     =   "Buscar Campo"
         Top             =   600
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
         Height          =   360
         Index           =   3
         Left            =   3840
         MaskColor       =   &H00000000&
         TabIndex        =   36
         ToolTipText     =   "Buscar Socio"
         Top             =   600
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
         Height          =   360
         Index           =   1
         Left            =   1590
         MaskColor       =   &H00000000&
         TabIndex        =   35
         ToolTipText     =   "Buscar Variedad"
         Top             =   600
         Visible         =   0   'False
         Width           =   195
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
         Left            =   6600
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Tag             =   "Situación|N|N|0|8|rclasifauto|situacion|||"
         Top             =   570
         Visible         =   0   'False
         Width           =   2280
      End
      Begin VB.TextBox Text1 
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
         Height          =   360
         Index           =   5
         Left            =   5670
         MaxLength       =   8
         TabIndex        =   3
         Tag             =   "Nombre|N|N|||rclasifauto|codcampo|00000000||"
         Text            =   "12345678"
         Top             =   600
         Visible         =   0   'False
         Width           =   885
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
         Height          =   360
         Index           =   4
         Left            =   3960
         TabIndex        =   31
         Text            =   "12345678901234567890"
         Top             =   600
         Visible         =   0   'False
         Width           =   1680
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
         Height          =   360
         Index           =   3
         Left            =   1620
         TabIndex        =   24
         Text            =   "12345678901234567890"
         Top             =   600
         Visible         =   0   'False
         Width           =   1620
      End
      Begin VB.TextBox Text1 
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
         Height          =   360
         Index           =   3
         Left            =   960
         MaxLength       =   6
         TabIndex        =   1
         Tag             =   "Variedad|N|S|0|999999|rclasifauto|codvarie|000000||"
         Text            =   "123456"
         Top             =   600
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.TextBox Text1 
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
         Height          =   360
         Index           =   1
         Left            =   11220
         MaxLength       =   7
         TabIndex        =   8
         Tag             =   "Peso Neto|N|N|||rclasifauto|kilosnet|###,##0||"
         Text            =   "neto"
         Top             =   570
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.TextBox Text1 
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
         Height          =   360
         Index           =   4
         Left            =   3270
         MaxLength       =   6
         TabIndex        =   2
         Tag             =   "Nombre|N|S|||rclasifauto|codsocio|000000||"
         Text            =   "123456"
         Top             =   600
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.TextBox Text1 
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
         Height          =   360
         Index           =   0
         Left            =   150
         MaxLength       =   7
         TabIndex        =   0
         Tag             =   "Nro.Nota|N|N|||rclasifauto|numnotac|0000000|S|"
         Text            =   "1234567"
         Top             =   600
         Visible         =   0   'False
         Width           =   750
      End
      Begin MSDataGridLib.DataGrid DataGridAux 
         Bindings        =   "frmManClasAuto.frx":000C
         Height          =   2760
         Index           =   1
         Left            =   120
         TabIndex        =   38
         Top             =   210
         Width           =   12990
         _ExtentX        =   22913
         _ExtentY        =   4868
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
      Begin MSAdodcLib.Adodc AdoAux 
         Height          =   375
         Index           =   1
         Left            =   8130
         Top             =   210
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
         Caption         =   "AdoAux(0)"
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
      Begin VB.Label Label6 
         Caption         =   "Campo"
         Height          =   255
         Index           =   0
         Left            =   3990
         TabIndex        =   32
         Top             =   270
         Width           =   615
      End
      Begin VB.Label Label10 
         Caption         =   "Variedad"
         Height          =   255
         Left            =   990
         TabIndex        =   25
         Top             =   300
         Width           =   675
      End
      Begin VB.Label Label1 
         Caption         =   "Kilos Neto"
         Height          =   255
         Index           =   2
         Left            =   5580
         TabIndex        =   23
         Top             =   270
         Width           =   795
      End
      Begin VB.Label Label6 
         Caption         =   "Socio"
         Height          =   255
         Index           =   1
         Left            =   3240
         TabIndex        =   22
         Top             =   300
         Width           =   690
      End
      Begin VB.Label Label29 
         Caption         =   "Situación"
         Height          =   255
         Left            =   4710
         TabIndex        =   16
         Top             =   270
         Width           =   1005
      End
      Begin VB.Label Label1 
         Caption         =   "Nro.Nota"
         Height          =   255
         Index           =   0
         Left            =   180
         TabIndex        =   14
         Top             =   315
         Width           =   945
      End
   End
   Begin VB.Frame FrameAux0 
      Caption         =   "Clasificación"
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
      Height          =   3180
      Left            =   30
      TabIndex        =   17
      Top             =   3900
      Width           =   6675
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
         Height          =   360
         Index           =   3
         Left            =   4740
         MaxLength       =   7
         TabIndex        =   30
         Tag             =   "Kilos Neto|N|S|||rclasifauto_clasif|kiloscal|###,##0.00||"
         Text            =   "neto"
         Top             =   2565
         Visible         =   0   'False
         Width           =   1080
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
         Index           =   2
         Left            =   3420
         MaskColor       =   &H00000000&
         TabIndex        =   33
         ToolTipText     =   "Buscar Calidad"
         Top             =   2565
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
         Height          =   360
         Index           =   2
         Left            =   3060
         MaxLength       =   2
         TabIndex        =   29
         Tag             =   "Calidad|N|N|||rclasifauto_clasif|codcalid|00|S|"
         Text            =   "Ca"
         Top             =   2565
         Visible         =   0   'False
         Width           =   375
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
         Height          =   360
         Index           =   2
         Left            =   3645
         TabIndex        =   28
         Text            =   "Calidad"
         Top             =   2565
         Visible         =   0   'False
         Width           =   1005
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
         Index           =   0
         Left            =   945
         MaskColor       =   &H00000000&
         TabIndex        =   27
         ToolTipText     =   "Buscar Envase"
         Top             =   2565
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
         Height          =   360
         Index           =   0
         Left            =   1155
         TabIndex        =   26
         Top             =   2565
         Visible         =   0   'False
         Width           =   1815
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
         Height          =   360
         Index           =   1
         Left            =   495
         MaxLength       =   6
         TabIndex        =   19
         Tag             =   "Variedad|N|N|||rclasifauto_clasif|codvarie|000000|S|"
         Text            =   "Var"
         Top             =   2565
         Visible         =   0   'False
         Width           =   405
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
         Height          =   360
         Index           =   0
         Left            =   30
         MaxLength       =   16
         TabIndex        =   18
         Tag             =   "Nro.Nota|N|N|||rclasifauto_clasif|numnotac|0000000|S|"
         Text            =   "nota"
         Top             =   2565
         Visible         =   0   'False
         Width           =   375
      End
      Begin MSComctlLib.Toolbar ToolAux 
         Height          =   390
         Index           =   0
         Left            =   135
         TabIndex        =   20
         Top             =   225
         Visible         =   0   'False
         Width           =   1200
         _ExtentX        =   2117
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
         Index           =   0
         Left            =   3720
         Top             =   225
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
         Caption         =   "AdoAux(0)"
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
         Bindings        =   "frmManClasAuto.frx":0024
         Height          =   2760
         Index           =   0
         Left            =   120
         TabIndex        =   21
         Top             =   300
         Width           =   6210
         _ExtentX        =   10954
         _ExtentY        =   4868
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
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Index           =   0
      Left            =   45
      TabIndex        =   11
      Top             =   7065
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
      Left            =   12045
      TabIndex        =   10
      Top             =   7140
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
      Left            =   10950
      TabIndex        =   9
      Top             =   7140
      Width           =   1035
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
      Left            =   12060
      TabIndex        =   15
      Top             =   7140
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   2
      Left            =   3750
      MaxLength       =   250
      TabIndex        =   34
      Top             =   810
      Width           =   2205
   End
   Begin MSComctlLib.Toolbar ToolbarAyuda 
      Height          =   330
      Left            =   12660
      TabIndex        =   46
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
   Begin VB.Label Label2 
      Caption         =   "TOTAL :"
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
      Index           =   103
      Left            =   3510
      TabIndex        =   40
      Top             =   7260
      Width           =   1005
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
         HelpContextID   =   2
         Shortcut        =   ^E
      End
      Begin VB.Menu mnBarra2 
         Caption         =   "-"
      End
      Begin VB.Menu mnTraspaso 
         Caption         =   "&Traspaso desde el Calibrador"
         Shortcut        =   ^T
      End
      Begin VB.Menu mnActualizar 
         Caption         =   "&Actualizar clasificación"
         Shortcut        =   ^A
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
Attribute VB_Name = "frmManClasAuto"
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

'Dim T1 As Single
Private Const IdPrograma = 4010

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

'Private WithEvents frmArt As frmManArtic 'articulos
Private WithEvents frmVar As frmComVar 'variedades de comercial
Attribute frmVar.VB_VarHelpID = -1
Private WithEvents frmSoc As frmManSocios 'socios
Attribute frmSoc.VB_VarHelpID = -1
Private WithEvents frmCap As frmManCapataz 'capataces
Attribute frmCap.VB_VarHelpID = -1
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

Private kCampo As Integer
'-------------------------------------------------------------------------
Private HaDevueltoDatos As Boolean

'Private VieneDeBuscar As Boolean
'Per a quan torna 2 poblacions en el mateix codi Postal. Si ve de pulsar prismatic
'de búsqueda posar el valor de població seleccionada i no tornar a recuperar de la Base de Datos

Dim Gastos As Boolean

Dim btnPrimero As Byte 'Variable que indica el nº del Botó PrimerRegistro en la Toolbar1
'Dim CadAncho() As Boolean  'array, per a quan cridem al form de llínies
Dim Indice As Byte 'Index del text1 on es poses els datos retornats des d'atres Formularis de Mtos
Dim CadB As String

Private Sub btnBuscar_Click(Index As Integer)
    TerminaBloquear
    Select Case Index
        Case 1 ' variedades
            Indice = Index + 2
            Set frmVar = New frmComVar
            frmVar.DatosADevolverBusqueda = "0|1|"
            frmVar.CodigoActual = Text1(3).Text
            frmVar.Show vbModal
            Set frmVar = Nothing
            PonerFoco Text1(3)
        
        Case 2 'calidades
            Indice = Index
            Set frmCal = New frmManCalidades
            frmCal.DatosADevolverBusqueda = "2|3|"
            frmCal.CodigoActual = txtAux(2).Text
            frmCal.Show vbModal
            Set frmCal = Nothing
            PonerFoco txtAux(2)
    
        Case 3 'socios
            Indice = Index + 1
            Set frmSoc = New frmManSocios
            frmSoc.DatosADevolverBusqueda = "0|1|"
'            frmSoc.CodigoActual = Text1(4).Text
            frmSoc.Show vbModal
            Set frmSoc = Nothing
            PonerFoco Text1(4)
            
        Case 4 'campos
            Indice = Index + 1
            Set frmCam = New frmManCampos
            frmCam.DatosADevolverBusqueda = "0|"
'            frmCam.CodigoActual = Text1(5).Text
            frmCam.Show vbModal
            Set frmCam = Nothing
            PonerFoco Text1(5)
    
    
    End Select
    
    If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Adoaux(1), 1
End Sub


Private Sub cmdAceptar_Click()
Dim I As Long

    Screen.MousePointer = vbHourglass
    
    On Error GoTo Error1
    
    Select Case Modo
        Case 1  'BÚSQUEDA
'            HacerBusqueda
            CadB = ObtenerBusqueda(Me)
            If CadB <> "" Then
                CargaGrid 1, True, CadB
                PonerModo 2
'                lblIndicador.Caption = "BUSQUEDA: " & PonerContRegistros(Me.adodc1)
                PonerFocoGrid Me.DataGridAux(1)
            End If
            
        Case 3 'INSERTAR
            If DatosOK Then
                If InsertarDesdeForm2(Me, 1) Then
'                    text2(9).Text = PonerNombreCuenta(text1(9), Modo, text1(0).Text)
        
                    Adoaux(1).RecordSource = "Select * from " & NombreTabla & Ordenacion
                    PosicionarData
                End If
            Else
                ModoLineas = 0
            End If
            
        Case 4  'MODIFICAR
            If DatosOK Then
                If ModificaDesdeFormulario1(Me, 1) Then
                    TerminaBloquear
                    I = Adoaux(1).Recordset.Fields(0)
                    PonerModo 2
                    CargaGrid 1, True, CadB
                    Adoaux(1).Recordset.Find (Adoaux(1).Recordset.Fields(0).Name & " =" & I)
                    PonerFocoGrid Me.DataGridAux(1)
                    
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
                    If ModificarLinea Then
'                        PosicionarData
'                        PasarSigReg
                    Else
'                        PonerFoco txtAux(12)
                    End If
            End Select
            'nuevo calculamos los totales de lineas
            CalcularTotales
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
        If (DatosADevolverBusqueda <> "") And NuevoCodigo <> "" Then
            BotonAnyadir
        Else
            PonerModo 2
            If Me.CodigoActual <> "" Then
                SituarData Me.Adoaux(1), "numnotac=" & CodigoActual, "", True
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
'    btnPrimero = 16 'index del botó "primero"
'    With Me.Toolbar1
'        .HotImageList = frmPpal.imgListComun_OM
'        .DisabledImageList = frmPpal.imgListComun_BN
'        .ImageList = frmPpal.imgListComun
'        'l'1 i el 2 son separadors
'        .Buttons(3).Image = 1   'Buscar
'        .Buttons(4).Image = 2   'Totss
'        'el 5 i el 6 son separadors
'        .Buttons(7).Image = 3   'Insertar
'        .Buttons(8).Image = 4   'Modificar
'        .Buttons(9).Image = 5   'Borrar
'        .Buttons(11).Image = 34 'Traspaso desde el calibrador
'        'el 10 i el 11 son separadors
'        .Buttons(12).Image = 33 'Actualizar la clasificacion
'        .Buttons(13).Image = 11  'Eixir
'        'el 13 i el 14 son separadors
'        .Buttons(btnPrimero).Image = 6  'Primer
'        .Buttons(btnPrimero + 1).Image = 7 'Anterior
'        .Buttons(btnPrimero + 2).Image = 8 'Següent
'        .Buttons(btnPrimero + 3).Image = 9 'Últim
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
        'el 10 i el 11 son separadors
'        .Buttons(11).Image = 26    'tarar tractor
'        .Buttons(12).Image = 24  'paletizacion
        .Buttons(8).Image = 10  'Imprimir
    End With
    
    With Me.Toolbar2
        .HotImageList = frmPpal.imgListComun_OM
        .DisabledImageList = frmPpal.imgListComun_BN
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 34  'Traspaso desde el calibrador
        .Buttons(2).Image = 33  'Actualizar la clasificacion
    End With
    
    ' La Ayuda
    With Me.ToolbarAyuda
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 12
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
    
    CargaCombo
    
    LimpiarCampos   'Neteja els camps TextBox
    ' ******* si n'hi han llínies *******
    DataGridAux(0).ClearFields
    DataGridAux(1).ClearFields
    
    '*** canviar el nom de la taula i l'ordenació de la capçalera ***
    NombreTabla = "rclasifauto"
    Ordenacion = " ORDER BY numnotac"
    
'    'Mirem com està guardat el valor del check
'    chkVistaPrevia(0).Value = CheckValueLeer(Name)
'
'    AdoAux(1).ConnectionString = conn
'    '***** cambiar el nombre de la PK de la cabecera *************
'    AdoAux(1).RecordSource = "Select * from " & NombreTabla & " where numnotac=-1"
'    AdoAux(1).Refresh
       
    CargaGrid 1, False
    CargaGrid 0, False
       
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
    ' *****************************************
    Combo1(0).ListIndex = -1

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
'    PonerIndicador lblIndicador, Modo, ModoLineas
    
       
    'Modo 2. N'hi han datos i estam visualisant-los
    '=========================================
    'Posem visible, si es formulari de búsqueda, el botó "Regresar" quan n'hi han datos
    If DatosADevolverBusqueda <> "" Then
        cmdRegresar.visible = (Modo = 2)
    Else
        cmdRegresar.visible = False
    End If
    
    B = (Modo = 2) Or (Modo = 0)
    
    If B Then
        PonerContRegIndicador
    Else
        PonerIndicador lblIndicador, Modo, ModoLineas
    End If
    CmdAceptar.visible = Not B
    cmdCancelar.visible = Not B
    
    For I = 0 To Text1.Count - 1
        Text1(I).visible = Not B
    Next I
    
    Text2(3).visible = Not B
    Text2(4).visible = Not B
    btnBuscar(1).visible = Not B
    btnBuscar(3).visible = Not B
    btnBuscar(4).visible = Not B
    Combo1(0).visible = Not B
    
    
    '=======================================
'    b = (Modo = 2)
'    'Posar Fleches de desplasament visibles
'    NumReg = 1
'    If Not adoaux(1).Recordset.EOF Then
'        If adoaux(1).Recordset.RecordCount > 1 Then NumReg = 2 'Només es per a saber que n'hi ha + d'1 registre
'    End If
'    DesplazamientoVisible Me.Toolbar1, btnPrimero, b, NumReg
    '---------------------------------------------
    
    B = Modo <> 0 And Modo <> 2
    cmdCancelar.visible = B
    CmdAceptar.visible = B
       
'    'Bloqueja els camps Text1 si no estem modificant/Insertant Datos
'    'Si estem en Insertar a més neteja els camps Text1
    BloquearText1 Me, Modo
'    BloquearCombo Me, Modo
'    '*** si n'hi han combos a la capçalera ***
'    Combo1(0).Enabled = (Modo = 1)
'    '**************************
    
    ' ***** bloquejar tots els controls visibles de la clau primaria de la capçalera ***
'    If Modo = 4 Then _
'        BloquearTxt Text1(0), True 'si estic en  modificar, bloqueja la clau primaria
    ' **********************************************************************************
    
    ' **** si n'hi han imagens de buscar en la capçalera *****
    BloquearImgBuscar Me, Modo, ModoLineas
    BloquearImgZoom Me, Modo, ModoLineas
    BloquearImgFec Me, 0, Modo, ModoLineas
    
    ' ********************************************************
        
    chkVistaPrevia(0).Enabled = (Modo <= 2)
    
    PonerLongCampos
    
    
    If (Modo < 2) Or (Modo = 3) Then
        CargaGrid 0, False
        CargaGrid 1, False
    End If
    
    B = (Modo = 4) Or (Modo = 2)
    DataGridAux(0).Enabled = B
    DataGridAux(1).Enabled = (Modo = 2)
    
    Text1(0).Enabled = (Modo = 1)
    Combo1(0).Enabled = (Modo = 1)
    ' ****** si n'hi han combos a la capçalera ***********************
    ' ****************************************************************
    
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
    Toolbar1.Buttons(1).Enabled = False 'B And Not DeConsulta
    Me.mnNuevo.Enabled = False 'B And Not DeConsulta
    
    B = (Modo = 2 And Me.Adoaux(1).Recordset.RecordCount > 0) And Not DeConsulta
    'Modificar
    Toolbar1.Buttons(2).Enabled = B
    Me.mnModificar.Enabled = B
    'eliminar
    Toolbar1.Buttons(3).Enabled = B
    Me.mnEliminar.Enabled = B
    'actualizar
    Toolbar2.Buttons(2).Enabled = B
    Me.mnActualizar.Enabled = B
    
    
    'Traspaso desde el calibrador
    'Toolbar1.Buttons(11).Enabled = True And Not DeConsulta
    'Imprimir
    Toolbar1.Buttons(8).Enabled = False '(B Or Modo = 0)
'    Toolbar1.Buttons(12).Enabled = True And Not DeConsulta
       
    ' *** si n'hi han llínies que tenen grids (en o sense tab) ***
'++monica: si insertamos lo he quitado
'    b = (Modo = 3 Or Modo = 4 Or Modo = 2) And Not DeConsulta
    B = (Modo = 2) And Not DeConsulta
    For I = 0 To ToolAux.Count - 1
        ToolAux(I).Buttons(1).Enabled = B
        If B Then bAux = (B And Me.Adoaux(I).Recordset.RecordCount > 0)
        ToolAux(I).Buttons(2).Enabled = bAux
        ToolAux(I).Buttons(3).Enabled = bAux
    Next I
    
End Sub

'Private Sub Desplazamiento(Index As Integer)
''Botons de Desplaçament; per a desplaçar-se pels registres de control Data
'    If adoaux(1).Recordset.EOF Then Exit Sub
'    DesplazamientoData adoaux(1), Index
'    PonerCampos
'End Sub

Private Function MontaSQLCarga(Index As Integer, enlaza As Boolean, Optional CadB As String) As String
'--------------------------------------------------------------------
' MontaSQlCarga:
'   Basant-se en la informació proporcionada pel vector de camps
'   crea un SQl per a executar una consulta sobre la base de datos que els
'   torne.
' Si ENLAZA -> Enlaça en el adoaux(1)
'           -> Si no el carreguem sense enllaçar a cap camp
'--------------------------------------------------------------------
Dim SQL As String
Dim tabla As String
    
    ' ********* si n'hi han tabs, dona igual si en datagrid o no ***********
    Select Case Index
               
        Case 0 'CLASIFICACION
            SQL = "SELECT rclasifauto_clasif.numnotac, rclasifauto_clasif.codvarie, rclasifauto_clasif.codcalid,"
            SQL = SQL & " rcalidad.nomcalid, rclasifauto_clasif.kiloscal "
            SQL = SQL & " from rclasifauto_clasif left join rcalidad on rclasifauto_clasif.codcalid = rcalidad.codcalid "
            SQL = SQL & " and rclasifauto_clasif.codvarie = rcalidad.codvarie "
            
            If enlaza Then
                SQL = SQL & ObtenerWhereCab(True)
            Else
                SQL = SQL & " WHERE rclasifauto_clasif.numnotac = -1"
            End If
'            SQL = SQL & " and rclasifauto_clasif.codcalid = rcalidad.codcalid "
'            SQL = SQL & " and rclasifauto_clasif.codvarie = rcalidad.codvarie "
            SQL = SQL & " ORDER BY rclasifauto_clasif.codvarie, rclasifauto_clasif.codcalid"
               
        Case 1 ' ENTRADAS DE CABECERA
            
    ' 0 = sin error
    ' 1 = No existe calidad
    ' 2 = No existe nro nota
    ' 3 = Tipo de clasificacion incorrecta
    ' 4 = Kilos netos diferentes
    ' 5 = No hay destrio
    ' 6 = Socios Diferentes
    ' 7 = Campos Diferentes
    ' 8 = Variedades Diferentes
            
            
            SQL = "select rclasifauto.numnotac, rclasifauto.codvarie, variedades.nomvarie, "
            SQL = SQL & "rclasifauto.codsocio, rsocios.nomsocio, rclasifauto.codcampo, rclasifauto.situacion,"
            SQL = SQL & "CASE rclasifauto.situacion WHEN 0 THEN ""SIN ERROR"" WHEN 1 THEN ""NO EXISTE CALIDAD"" "
            SQL = SQL & " WHEN 2 THEN ""NO EXISTE NRO.NOTA"" WHEN 3 THEN ""DESTRIO SUPERIOR AL 50%"" "
            SQL = SQL & " WHEN 4 THEN ""KILOS NETOS DIFERENTES"" WHEN 5 THEN ""NO HAY DESTRIO"" "
            SQL = SQL & " WHEN 6 THEN ""SOCIOS DIFERENTES"" WHEN 7 THEN ""CAMPOS DIFERENTES"" "
            SQL = SQL & " WHEN 8 THEN ""VARIEDADES DIFERENTES"" END, "
            SQL = SQL & " rclasifauto.kilospod, rclasifauto.kilospeq, rclasifauto.kilosdes, rclasifauto.kilosnet "
            SQL = SQL & " from (rclasifauto left join variedades on rclasifauto.codvarie = variedades.codvarie) "
            SQL = SQL & " left join rsocios on rclasifauto.codsocio = rsocios.codsocio"
            
            If enlaza Then
                SQL = SQL & " WHERE 1=1 "
                If CadB <> "" Then
                    SQL = SQL & " and " & CadB
                End If
            Else
                SQL = SQL & " WHERE rclasifauto.numnotac = -1"
            End If
            SQL = SQL & " ORDER BY rclasifauto.numnotac"
    
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
    Text1(Indice).Text = Format(vFecha, "dd/mm/yyyy") 'fecha clasificacion
End Sub

Private Sub frmCal_DatoSeleccionado(CadenaSeleccion As String)
'Calidad
    txtAux(Indice).Text = RecuperaValor(CadenaSeleccion, 1) 'codcalid
    txtAux2(Indice).Text = RecuperaValor(CadenaSeleccion, 2) 'descripcion
End Sub

Private Sub frmCam_DatoSeleccionado(CadenaSeleccion As String)
'Campos
    Text1(Indice).Text = RecuperaValor(CadenaSeleccion, 1) 'codcampo
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

Private Sub frmVar_DatoSeleccionado(CadenaSeleccion As String)
'Variedades
    Text1(3).Text = RecuperaValor(CadenaSeleccion, 1) 'codvariedad
    Text2(3).Text = RecuperaValor(CadenaSeleccion, 2) 'descripcion
End Sub

Private Sub frmZ_Actualizar(vCampo As String)
     Text1(Indice).Text = vCampo
End Sub


Private Sub imgFec_Click(Index As Integer)
'    Dim esq As Long
'    Dim dalt As Long
'    Dim menu As Long
'    Dim obj As Object
'
'    Set frmC = New frmCal
'
'    esq = imgFec(Index).Left
'    dalt = imgFec(Index).Top
'
'    Set obj = imgFec(Index).Container
'
'    While imgFec(Index).Parent.Name <> obj.Name
'          esq = esq + obj.Left
'          dalt = dalt + obj.Top
'          Set obj = obj.Container
'    Wend
'
'    menu = Me.Height - Me.ScaleHeight 'ací tinc el heigth del menú i de la toolbar
'
'    frmC.Left = esq + imgFec(Index).Parent.Left + 30
'    frmC.Top = dalt + imgFec(Index).Parent.Top + imgFec(Index).Height + menu - 40
'
'    Select Case Index
'        Case 0
'            indice = Index + 6
'    End Select
'
'    imgFec(0).Tag = indice '<===
'    ' *** repasar si el camp es txtAux o Text1 ***
'    If Text1(indice).Text <> "" Then frmC.NovaData = Text1(indice).Text
'    ' ********************************************
'
'    frmC.Show vbModal
'    Set frmC = Nothing
'    ' *** repasar si el camp es txtAux o Text1 ***
'    PonerFoco Text1(indice) '<===
'    ' ********************************************
End Sub

Private Sub imgZoom_Click(Index As Integer)
    
    Set frmZ = New frmZoom

    If Index = 0 Then
        Indice = 2
        frmZ.pTitulo = "Observaciones de la Clasificación"
        frmZ.pValor = Text1(Indice).Text
        frmZ.pModo = Modo
    
        frmZ.Show vbModal
        Set frmZ = Nothing
            
        PonerFoco Text1(Indice)
    End If
End Sub

Private Sub mnActualizar_Click()
    BotonActualizar
End Sub

Private Sub mnBuscar_Click()
Dim I As Integer
    BotonBuscar
End Sub

Private Sub mnEliminar_Click()
    BotonEliminar
End Sub

Private Sub mnImprimir_Click()
'    Screen.MousePointer = vbHourglass
'    frmListConfeccion.Show vbModal
'    Screen.MousePointer = vbDefault
End Sub

Private Sub mnModificar_Click()
    ' *** repasar el nom de l'adodc, l'index del Field i el camp que te la PK ***
    ' ### [Monica] 26/09/2006 dejamos modificar y eliminar el codigo 0
    'El registre de codi 0 no es pot Modificar ni Eliminar
    'If EsCodigoCero(CStr(adoaux(1).Recordset.Fields(0).Value), FormatoCampo(Text1(0))) Then Exit Sub
    ' ***************************************************************************
    
    If BLOQUEADesdeFormulario2(Me, Adoaux(1), 1) Then BotonModificar
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
        Case 12 'Actualizar entradas
            mnActualizar_Click
        Case 13    'Eixir
            mnSalir_Click
            
'        Case btnPrimero To btnPrimero + 3 'Fleches Desplaçament
'            Desplazamiento (Button.Index - btnPrimero)
    End Select
End Sub

Private Sub BotonBuscar()
Dim I As Integer

' ***** Si la clau primaria de la capçalera no es Text1(0), canviar-ho en <=== *****
    If Modo <> 1 Then
        LimpiarCampos
        LLamaLineas 1, 1, DataGridAux(1).Top + 206 'Pone el form en Modo=1, Buscar
        PonerFoco Text1(0) ' <===
        Text1(0).BackColor = vbYellow ' <===
        ' *** si n'hi han combos a la capçalera ***
    Else
        HacerBusqueda
        If Adoaux(1).Recordset.EOF Then
            Text1(kCampo).Text = ""
            Text1(kCampo).BackColor = vbYellow
            PonerFoco Text1(kCampo)
        End If
    End If
' ******************************************************************************
End Sub

Private Sub BotonActualizar()
Dim SQL As String
Dim B As Boolean

    SQL = "select count(*) from rclasifauto where situacion <> 0"
    
    If TotalRegistros(SQL) <> 0 Then
        MsgBox "Hay entradas con error. Revise.", vbExclamation
    Else
        B = False
        Select Case vParamAplic.Cooperativa
            Case 0 ' Catadau
                B = ActualizarEntradasCatadau
            Case 1 ' Valsur
                B = ActualizarEntradasValsur
            Case 4 ' Alzira
                B = ActualizarEntradasAlzira
            Case 5 ' Castelduc
                B = ActualizarEntradasCastelduc
            '[Monica]29/02/2012: Natural era la cooperativa 0 junto con Catadau ahora es la 9
            Case 9 ' Natural
                B = ActualizarEntradasCatadau
            Case 16 'COOPIC
                B = ActualizarEntradasCoopic
            
        End Select
        If B Then
            MsgBox "Proceso realizado correctamente", vbExclamation
            BotonVerTodos
        End If
    End If
        
End Sub






Private Sub HacerBusqueda()
    
    CadB = ObtenerBusqueda2(Me, 1)
    
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
    Dim cad As String
        
    'Cridem al form
    ' **************** arreglar-ho per a vore lo que es desije ****************
    ' NOTA: el total d'amples de ParaGrid, ha de sumar 100
    cad = ""
    cad = cad & ParaGrid(Text1(0), 20, "Código")
    cad = cad & ParaGrid(Text1(1), 50, "Confección")
'    cad = cad & ParaGrid(text1(2), 60, "Descripción")
    cad = cad & "Variedad|nomvarie|T||30·"
    If cad <> "" Then
        
        Screen.MousePointer = vbHourglass
        Set frmB = New frmBuscaGrid
        frmB.vCampos = cad
        cad = NombreTabla & " inner join variedades on forfaits.codvarie = variedades.codvarie "
        frmB.vtabla = cad 'NombreTabla
        frmB.vSQL = CadB
        HaDevueltoDatos = False
        frmB.vDevuelve = "0|1|2|" '*** els camps que volen que torne ***
        frmB.vTitulo = "Forfaits" ' ***** repasa açò: títol de BuscaGrid *****
        frmB.vSelElem = 1

        frmB.Show vbModal
        Set frmB = Nothing
        'Si ha posat valors i tenim que es formulari de búsqueda llavors
        'tindrem que tancar el form llançant l'event
        If HaDevueltoDatos Then
            If (Not Adoaux(1).Recordset.EOF) And DatosADevolverBusqueda <> "" Then _
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

    If Adoaux(1).Recordset.EOF Then
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
    
    Adoaux(1).RecordSource = CadenaConsulta
    Adoaux(1).Refresh
    
    If Adoaux(1).Recordset.RecordCount <= 0 Then
        MsgBox "No hay ningún registro en la tabla " & NombreTabla, vbInformation
        Screen.MousePointer = vbDefault
        Exit Sub
    Else
        PonerModo 2
        'adoaux(1).Recordset.MoveLast
        Adoaux(1).Recordset.MoveFirst
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
''Vore tots
'    LimpiarCampos 'Neteja els Text1
'    CadB = ""
'
'    If chkVistaPrevia(0).Value = 1 Then
'        MandaBusquedaPrevia ""
'    Else
'        CadenaConsulta = "Select * from " & NombreTabla & Ordenacion
'        PonerCadenaBusqueda
'    End If
    
    CadB = ""
    CargaGrid 1, True, CadB
    PonerModo 2
    If Adoaux(1).Recordset.EOF Then
        CargaGrid 0, False
    Else
        CargaGrid 0, True
    End If
    
    
End Sub

Private Sub BotonAnyadir()
Dim NumF As String
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


       
    Text1(0) = NumF
    PonerFoco Text1(0) '*** 1r camp visible que siga PK ***
    
    ' *** si n'hi han camps de descripció a la capçalera ***
    'PosarDescripcions

End Sub

Private Sub BotonModificar()

    Dim anc As Single
    Dim I As Integer
    
    Screen.MousePointer = vbHourglass
    
    If DataGridAux(1).Bookmark < DataGridAux(1).FirstRow Or DataGridAux(1).Bookmark > (DataGridAux(1).FirstRow + DataGridAux(1).VisibleRows - 1) Then
        I = DataGridAux(1).Bookmark - DataGridAux(1).FirstRow
        DataGridAux(1).Scroll 0, I
        DataGridAux(1).Refresh
    End If
    
    If DataGridAux(1).Row < 0 Then
        anc = 320
    Else
        anc = DataGridAux(1).RowTop(DataGridAux(1).Row) + 210 '545
    End If

    'Llamamos al form
    Text1(0).Text = DataGridAux(1).Columns(0).Text
    Text1(3).Text = DataGridAux(1).Columns(1).Text
    Text1(4).Text = DataGridAux(1).Columns(3).Text
    Text1(5).Text = DataGridAux(1).Columns(5).Text
    Text1(1).Text = DataGridAux(1).Columns(11).Text
    Text1(6).Text = DataGridAux(1).Columns(8).Text
    Text1(7).Text = DataGridAux(1).Columns(9).Text
    Text1(8).Text = DataGridAux(1).Columns(10).Text
    
    ' ***** canviar-ho pel nom del camp del combo *********
    I = Adoaux(1).Recordset!Situacion
    ' *****************************************************
    PosicionarCombo Me.Combo1(0), I
'    For j = 0 To Combo1.ListCount - 1
'        If Combo1.ItemData(j) = i Then
'            Combo1.ListIndex = j
'            Exit For
'        End If
'    Next j

    LLamaLineas 1, 4, anc 'Pone el form en Modo=4, Modificar
   
    'Como es modificar
    PonerFoco Text1(3)
    Screen.MousePointer = vbDefault
End Sub

Private Sub BotonEliminar()
Dim cad As String

    On Error GoTo EEliminar

    'Ciertas comprobaciones
    If Adoaux(1).Recordset.EOF Then Exit Sub

    ' *** repasar el nom de l'adodc, l'index del Field i el camp que te la PK ***
    ' ### [Monica] 26/09/2006 dejamos modificar y eliminar el codigo 0
    'El registre de codi 0 no es pot Modificar ni Eliminar
    'If EsCodigoCero(CStr(adoaux(1).Recordset.Fields(0).Value), FormatoCampo(text1(0))) Then Exit Sub
    ' ***************************************************************************

    ' *************** canviar la pregunta ****************
    cad = "¿Seguro que desea eliminar la Clasificación?"
    cad = cad & vbCrLf & "Número: " & Adoaux(1).Recordset.Fields(0)
    
    If MsgBox(cad, vbQuestion + vbYesNo) = vbYes Then
        On Error GoTo EEliminar
        Screen.MousePointer = vbHourglass
        NumRegElim = Adoaux(1).Recordset.AbsolutePosition
        If Not Eliminar Then
            Screen.MousePointer = vbDefault
            Exit Sub
        ElseIf SituarDataTrasEliminar(Adoaux(1), NumRegElim, True) Then
'            PonerCampos
            CargaGrid 1, True, CadB
            SituarDataTrasEliminar Adoaux(1), NumRegElim, True
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

    If Adoaux(1).Recordset.EOF Then Exit Sub
    
    PonerCamposForma2 Me, Adoaux(1), 1 'opcio=1: posa el format o els camps de la capçalera
    
    ' *** si n'hi han llínies en datagrids ***
    'For i = 0 To DataGridAux.Count - 1
    CargaGrid I, True
    If Not Adoaux(I).Recordset.EOF Then _
        PonerCamposForma2 Me, Adoaux(I), 2, "FrameAux" & I

    
    ' ************* configurar els camps de les descripcions de la capçalera *************
    Text2(3).Text = PonerNombreDeCod(Text1(3), "variedades", "nomvarie")
    Text2(4).Text = PonerNombreDeCod(Text1(4), "rsocios", "nomsocio")
'    Text2(5).Text = PonerNombreDeCod(Text1(8), "rcampos", "nomcapac")
    ' ********************************************************************************
    
    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = Adoaux(1).Recordset.AbsolutePosition & " de " & Adoaux(1).Recordset.RecordCount
    
    PonerModoOpcionesMenu (Modo)
    PonerOpcionesMenu
End Sub

Private Sub cmdCancelar_Click()
Dim I As Integer
Dim V

    Select Case Modo
        Case 1, 3 'Búsqueda, Insertar
                LimpiarCampos
                If Adoaux(1).Recordset.EOF Then
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
                        ' **** repasar si es diu adoaux(1) l'adodc de la capçalera ***
                        'If BLOQUEADesdeFormulario2(Me, adoaux(1), 1) Then BotonModificar 'Modificar
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
                        V = Adoaux(NumTabMto).Recordset.Fields(2) 'el 2 es el nº de llinia
                        Adoaux(NumTabMto).Recordset.Find (Adoaux(NumTabMto).Recordset.Fields(2).Name & " =" & V)
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
Dim SQL As String
Dim Rs As ADODB.Recordset

'Dim Datos As String

    On Error GoTo EDatosOK

    DatosOK = False
    
    B = CompForm2(Me, 1)
    If Not B Then Exit Function
    
    ' *** canviar els arguments de la funcio, el mensage i repasar si n'hi ha codEmpre ***
    If Modo = 4 And B Then
        ' comprobamos los datos modificados con la not de entrada
        SQL = "select codsocio, codcampo, codvarie, kilosnet from rclasifica where numnotac = "
        SQL = SQL & DBSet(Text1(0).Text, "N")
        
        Set Rs = New ADODB.Recordset
        Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
    'situacion:
    ' 0 = sin error
    ' 1 = No existe calidad
    ' 2 = No existe nro nota
    ' 3 = Tipo de clasificacion incorrecta
    ' 4 = Kilos netos diferentes
    ' 5 = No hay destrio
    ' 6 = Socios Diferentes
    ' 7 = Campos Diferentes
    ' 8 = Variedades Diferentes
        
        
        If Not Rs.EOF Then
            ' no coinciden socios
            If CLng(DBLet(Rs.Fields(0).Value, "N")) <> CLng(Text1(4).Text) Then
                MsgBox "El socio de la entrada no se corresponde con el de la clasificación automática.Revise", vbExclamation
                Combo1(0).ListIndex = 6
                B = False
            End If
            ' no coinciden campos
            If B And CLng(DBLet(Rs.Fields(1).Value, "N")) <> CLng(Text1(5).Text) Then
                MsgBox "El campo de la entrada no se corresponde con el de la clasificación automática.Revise", vbExclamation
                Combo1(0).ListIndex = 7
                B = False
            End If
            ' no coinciden variedades
            If B And CLng(DBLet(Rs.Fields(2).Value, "N")) <> CLng(Text1(3).Text) Then
                MsgBox "La variedad de la entrada no se corresponde con la de la clasificación automática.Revise", vbExclamation
                Combo1(0).ListIndex = 8
                B = False
            End If
            ' no coinciden kilos netos
            If B And CLng(DBLet(Rs.Fields(3).Value, "N")) <> CLng(Text1(1).Text) Then
                MsgBox "El peso neto de la entrada no se corresponde con el de la clasificación automática.Revise", vbExclamation
                Combo1(0).ListIndex = 4
                B = False
            End If
            
            If B Then
                ' la entrada es correcta
                Combo1(0).ListIndex = 0
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
    cad = "(numnotac=" & DBSet(Text1(0).Text, "N") & ")"
    
    ' *** gastar SituarData o SituarDataMULTI depenent de si la PK es simple o composta ***
    'If SituarDataMULTI(adoaux(1), cad, Indicador) Then
    If SituarData(Adoaux(1), cad, Indicador) Then
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
    vWhere = " WHERE numnotac=" & Adoaux(1).Recordset!numnotac
    
    ' ***** elimina les llínies ****
    conn.Execute "DELETE FROM rclasifauto_clasif " & vWhere
        
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

Private Sub Text1_LostFocus(Index As Integer)
Dim cadMen As String
Dim Nuevo As Boolean

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
                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Adoaux(1), 1
                    Else
                        Text1(Index).Text = ""
                    End If
                    PonerFoco Text1(Index)
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
                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Adoaux(1), 1
                    Else
                        Text1(Index).Text = ""
                    End If
                    PonerFoco Text1(Index)
                End If
            Else
                Text2(Index).Text = ""
            End If
                
'        Case 6 'Fecha
'            PonerFormatoFecha Text1(Index)
        
        Case 1 'kilos
            PonerFormatoEntero Text1(Index)
        
        Case 6, 7, 8 ' podrido,merma,destrio
            PonerFormatoEntero Text1(Index)
        
        Case 5 'campo
            PonerFormatoEntero Text1(Index)
        
    End Select
    ' ***************************************************************************
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = teclaBuscar Then
        If Modo = 1 Or Modo = 3 Or Modo = 4 Then
            Select Case Index
                Case 5: KEYBusqueda KeyAscii, 2 'campo
                Case 4: KEYBusqueda KeyAscii, 1 'socio
                Case 3: KEYBusqueda KeyAscii, 0 'variedad
                Case 6: KEYFecha KeyAscii, 0    'fecha de clasificacion
            End Select
        End If
    Else
        If Index <> 3 Or (Index = 3 And Text1(3).Text = "") Then KEYpress KeyAscii
    End If
End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Alvançar/Retrocedir els camps en les fleches de desplaçament del teclat.
    KEYdown KeyCode
End Sub

Private Sub KEYFecha(KeyAscii As Integer, Indice As Integer)
    KeyAscii = 0
    imgFec_Click (Indice)
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



'************* LLINIES: ****************************
Private Sub ToolAux_ButtonClick(Index As Integer, ByVal Button As MSComctlLib.Button)
'-- pon el bloqueo aqui
    'If BLOQUEADesdeFormulario2(Me, adoaux(1), 1) Then
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
        Case 0 'calidades
            SQL = "¿Seguro que desea eliminar la Calidad?"
            SQL = SQL & vbCrLf & "Calidad: " & Adoaux(Index).Recordset!codcalid
            If MsgBox(SQL, vbQuestion + vbYesNo) = vbYes Then
                Eliminar = True
                SQL = "DELETE FROM rclasifauto_clasif "
                SQL = SQL & vWhere & " AND codvarie= " & Adoaux(Index).Recordset!codvarie
                SQL = SQL & " and codcalid= " & Adoaux(Index).Recordset!codcalid
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
        If BLOQUEADesdeFormulario2(Me, Adoaux(1), 1) Then BotonModificar
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
        Case 0: vtabla = "rclasifica_clasif"
        Case 1: vtabla = "rclasifica_incidencia"
    End Select
    
    vWhere = ObtenerWhereCab(False)
    
    Select Case Index
        Case 0, 1 ' *** pose els index dels tabs de llínies que tenen datagrid ***
            ' *** canviar la clau primaria de les llínies,
            'pasar a "" si no volem que mos sugerixca res a l'afegir ***
            
'            If Index = 0 Then NumF = SugerirCodigoSiguienteStr(vTabla, "numlinea", vWhere)

            AnyadirLinea DataGridAux(Index), Adoaux(Index)
    
            anc = DataGridAux(Index).Top
            If DataGridAux(Index).Row < 0 Then
                anc = anc + 210
            Else
                anc = anc + DataGridAux(Index).RowTop(DataGridAux(Index).Row) + 5
            End If
            
            LLamaLineas Index, ModoLineas, anc
        
            Select Case Index
                ' *** valor per defecte a l'insertar i formateig de tots els camps ***
                Case 0 'calidades
                    txtAux(0).Text = Text1(0).Text 'numnotac
                    txtAux(1).Text = Text1(3).Text 'codvarie
                    txtAux(2).Text = ""
                    txtAux2(2).Text = ""
                    txtAux(3).Text = ""
                    txtAux(4).Text = ""
                    BloquearTxt txtAux(2), False
                    BloquearTxt txtAux(3), False
                    BloquearbtnBuscar Me, Modo, ModoLineas, "FrameAux0"
                    Me.btnBuscar(0).Enabled = False
                    Me.btnBuscar(0).visible = False
                    PonerFoco txtAux(2)
                Case 1 'incidencias
                    txtAux(8).Text = Text1(0).Text 'numnotac
                    txtAux(9).Text = "" 'NumF 'codcoste
                    txtAux2(9).Text = ""
                    For I = 9 To 9
                        BloquearTxt txtAux(I), False
                    Next I
                    BloquearbtnBuscar Me, Modo, ModoLineas, "FrameAux1"
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
        Case 0, 1 ' *** pose els index de llínies que tenen datagrid (en o sense tab) ***
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
        Case 0 ' muestra
        
            For J = 0 To 1
                txtAux(J).Text = DataGridAux(Index).Columns(J).Text
            Next J
'            txtAux2(0).Text = DataGridAux(Index).Columns(2).Text
            txtAux(2).Text = DataGridAux(Index).Columns(2).Text
            txtAux2(2).Text = DataGridAux(Index).Columns(3).Text
            txtAux(3).Text = DataGridAux(Index).Columns(4).Text
            BloquearTxt txtAux(0), True
            BloquearbtnBuscar Me, Modo, ModoLineas, "FrameAux0"
            
    End Select
    
    LLamaLineas Index, ModoLineas, anc
   
    ' *** foco al 1r camp visible de les llinies en grids que no siga PK (en o sense tab) ***
    Select Case Index
        Case 0 'muestras
            PonerFoco txtAux(3)
    End Select
    ' ***************************************************************************************
End Sub

Private Sub LLamaLineas(Index As Integer, xModo As Byte, Optional alto As Single)
Dim jj As Integer
Dim B As Boolean

    ' *** si n'hi han tabs sense datagrid posar el If ***
    DeseleccionaGrid DataGridAux(Index)
    PonerModo xModo
       
    Select Case Index
        Case 0 'muestras
            B = (xModo = 1 Or xModo = 2) 'Insertar o Modificar Llínies
            For jj = 2 To 3
                txtAux(jj).visible = B
                txtAux(jj).Top = alto
            Next jj
            For jj = 2 To 2
                txtAux2(jj).visible = B
                txtAux2(jj).Top = alto
            Next jj
            btnBuscar(2).visible = B
            btnBuscar(2).Top = alto
       Case 1 ' entradas
            B = (xModo = 1 Or xModo = 4)
            Text1(0).visible = B
            Text1(0).Top = alto
            Text1(3).visible = B
            Text1(3).Top = alto
            Text1(4).visible = B
            Text1(4).Top = alto
            Text1(5).visible = B
            Text1(5).Top = alto
            Text1(1).visible = B
            Text1(1).Top = alto
            Text1(6).visible = B
            Text1(6).Top = alto
            Text1(7).visible = B
            Text1(7).Top = alto
            Text1(8).visible = B
            Text1(8).Top = alto
            Text2(3).visible = B
            Text2(3).Top = alto
            Text2(4).visible = B
            Text2(4).Top = alto
            btnBuscar(1).visible = B
            btnBuscar(1).Top = alto
            btnBuscar(3).visible = B
            btnBuscar(3).Top = alto
            btnBuscar(4).visible = B
            btnBuscar(4).Top = alto
            Combo1(0).visible = B
            Combo1(0).Top = alto
            
    End Select
End Sub


Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button
        Case 2 'Actualizar entradas
            mnActualizar_Click
    End Select
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
        Case 2 ' codigo de calidad
            If txtAux(Index) <> "" Then
                txtAux2(2).Text = PonerNombreDeCod(txtAux(Index), "rcalidad", "nomcalid")
                If txtAux2(2).Text = "" Then
                    cadMen = "No existe la Calidad: " & txtAux(Index).Text & vbCrLf
                    cadMen = cadMen & "¿Desea crearla?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmCal = New frmManCalidades
                        frmCal.DatosADevolverBusqueda = "0|1|"
                        frmCal.NuevoCodigo = txtAux(Index).Text
                        txtAux(Index).Text = ""
                        TerminaBloquear
                        frmCal.Show vbModal
                        Set frmCal = Nothing
                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Adoaux(1), 1
                    Else
                        txtAux(Index).Text = ""
                    End If
                    PonerFoco txtAux(Index)
                End If
            Else
                txtAux2(2).Text = ""
            End If
        
        Case 9 ' codigo de incidencia
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
                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Adoaux(1), 1
                    Else
                        txtAux(Index).Text = ""
                    End If
                    PonerFoco txtAux(Index)
                Else
                    CmdAceptar.SetFocus
                End If
            Else
                txtAux2(Index).Text = ""
            End If
        
        Case 2 ' codigo de calidad
            PonerFormatoEntero txtAux(Index)
            
        Case 3  ' muestra
            If txtAux(Index).Text <> "" Then PonerFormatoDecimal txtAux(Index), 3
            
        Case 4 ' kilosnetos
            PonerFormatoEntero txtAux(Index)
            
            CmdAceptar.SetFocus
    End Select
End Sub

Private Sub txtaux_GotFocus(Index As Integer)
   If Not txtAux(Index).MultiLine Then ConseguirFocoLin txtAux(Index)
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
    End Select
    
    If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Adoaux(1), 1
End Sub

Private Sub DataGridAux_RowColChange(Index As Integer, LastRow As Variant, ByVal LastCol As Integer)
Dim I As Byte

    Select Case Index
        Case 1 ' entradas
            PonerContRegIndicador
            CargaGrid 0, True
    End Select
    
'
'    If ModoLineas <> 1 Then
'        Select Case Index
'            Case 0 'cuentas bancarias
'                If DataGridAux(Index).Columns.Count > 2 Then
''                    txtAux(11).Text = DataGridAux(Index).Columns("direccio").Text
''                    txtAux(12).Text = DataGridAux(Index).Columns("observac").Text
'                End If
'
'            Case 1 'departamentos
'                If DataGridAux(Index).Columns.Count > 2 Then
''                    txtAux(21).Text = DataGridAux(Index).Columns(5).Text
''                    txtAux(22).Text = DataGridAux(Index).Columns(6).Text
''                    txtAux(23).Text = DataGridAux(Index).Columns(8).Text
''                    txtAux(24).Text = DataGridAux(Index).Columns(15).Text
''                    txtAux2(22).Text = DataGridAux(Index).Columns(7).Text
'                End If
'
'        End Select
'
'    Else 'vamos a Insertar
'        Select Case Index
'            Case 0 'cuentas bancarias
''                txtAux(11).Text = ""
''                txtAux(12).Text = ""
'            Case 1 'departamentos
'                For I = 21 To 24
''                   txtAux(i).Text = ""
'                Next I
''               txtAux2(22).Text = ""
'            Case 2 'Tarjetas
''               txtAux(50).Text = ""
''               txtAux(51).Text = ""
'        End Select
'    End If
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

Private Sub CargaGrid(Index As Integer, enlaza As Boolean, Optional CadB As String)
Dim B As Boolean
Dim I As Byte
Dim tots As String

    On Error GoTo ECarga

    tots = MontaSQLCarga(Index, enlaza, CadB)

    CargaGridGnral Me.DataGridAux(Index), Me.Adoaux(Index), tots, PrimeraVez
    
    Select Case Index
        Case 0 'clasificacion
            'si es visible|control|tipo campo|nombre campo|ancho control|
            tots = "N||||0|;" 'numnotac
            tots = tots & "N|txtAux(1)|T|Variedad|800|;" '"N|btnBuscar(0)|B|||;N|txtAux2(0)|T|Nombre|2000|;"
            tots = tots & "S|txtAux(2)|T|Calidad|1000|;S|btnBuscar(2)|B|||;S|txtAux2(2)|T|Nombre|3200|;"
            tots = tots & "S|txtAux(3)|T|Muestra|1400|;"
            
            arregla tots, DataGridAux(Index), Me, 350
            
'            DataGridAux(0).Columns(3).Alignment = dbgLeft
'            DataGridAux(0).Columns(5).NumberFormat = "###,##0"
'            DataGridAux(0).Columns(5).Alignment = dbgRight
        
            B = (Modo = 4) And ((ModoLineas = 1) Or (ModoLineas = 2))
            
            CalcularTotales
    
        Case 1 'entradas
            'si es visible|control|tipo campo|nombre campo|ancho control|
            tots = "S|Text1(0)|T|Nota|950|;" 'numnotac
            tots = tots & "S|Text1(3)|T|Codigo|800|;S|btnBuscar(1)|B|||;S|Text2(3)|T|Variedad|1100|;"
            tots = tots & "S|Text1(4)|T|Socio|850|;S|btnBuscar(3)|B|||;S|Text2(4)|T|Nombre|2100|;"
            tots = tots & "S|Text1(5)|T|Campo|1000|;S|btnBuscar(4)|B|||;N|||||;S|Combo1(0)|C|Situación|2100|;"
            tots = tots & "S|Text1(6)|T|Podrido|900|;"
            tots = tots & "S|Text1(7)|T|Pequeño|920|;"
            tots = tots & "S|Text1(8)|T|Destrio|900|;"
            tots = tots & "S|Text1(1)|T|Neto|800|;"
            
            arregla tots, DataGridAux(Index), Me, 350
            
            DataGridAux(1).Columns(1).Alignment = dbgLeft
            DataGridAux(1).Columns(3).Alignment = dbgLeft
            DataGridAux(1).Columns(5).Alignment = dbgLeft
            DataGridAux(1).Columns(8).NumberFormat = "###,##0"
            DataGridAux(1).Columns(8).Alignment = dbgRight
        
            B = (Modo = 4) And ((ModoLineas = 1) Or (ModoLineas = 2))
    
    
    
    End Select
    
    DataGridAux(Index).ScrollBars = dbgAutomatic
      
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
        Case 0: nomframe = "FrameAux0" 'envases
        Case 1: nomframe = "FrameAux1" 'costes
    End Select
    
    If DatosOkLlin(nomframe) Then
        TerminaBloquear
        If InsertarDesdeForm2(Me, 2, nomframe) Then
            B = BLOQUEADesdeFormulario2(Me, Adoaux(1), 1)
            Select Case NumTabMto
                Case 0, 1 ' *** els index de les llinies en grid (en o sense tab) ***
                     CargaGrid NumTabMto, True
                    If B Then BotonAnyadirLinea NumTabMto
            End Select
           
'            SituarTab (NumTabMto + 1)
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
        Case 0: nomframe = "FrameAux0" 'envases
        Case 1: nomframe = "FrameAux1" 'costes
    End Select
    ModificarLinea = False
    If DatosOkLlin(nomframe) Then
        TerminaBloquear
        If ModificaDesdeFormulario2(Me, 2, nomframe) Then
            ModoLineas = 0
            Select Case NumTabMto
                Case 0
                    V = Adoaux(NumTabMto).Recordset.Fields(3) 'el 2 es el nº de llinia
                Case 1
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
    vWhere = vWhere & " numnotac=" & Me.Adoaux(1).Recordset!numnotac
    
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




Private Sub CargaCombo()
Dim I As Integer

    ' *** neteje els combos, els pose valor i seleccione el valor per defecte ***
    For I = 0 To Combo1.Count - 1
        Combo1(I).Clear
    Next I
    'situacion:
    ' 0 = sin error
    ' 1 = No existe calidad
    ' 2 = No existe nro nota
    ' 3 = Destrio Superior al 50%
    ' 4 = Kilos netos diferentes
    ' 5 = No hay destrio
    ' 6 = Socios Diferentes
    ' 7 = Campos Diferentes
    ' 8 = Variedades Diferentes
    
    Combo1(0).AddItem "SIN ERROR"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 0
    Combo1(0).AddItem "NO EXISTE CALIDAD"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 1
    Combo1(0).AddItem "NO EXISTE NRO.NOTA"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 2
    Combo1(0).AddItem "DESTRIO SUPERIOR AL 50%"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 3
    Combo1(0).AddItem "KILOS NETOS DIFERENTES"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 4
    Combo1(0).AddItem "NO HAY DESTRIO"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 5
    Combo1(0).AddItem "SOCIOS DIFERENTES"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 6
    Combo1(0).AddItem "CAMPOS DIFERENTES"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 7
    Combo1(0).AddItem "VARIEDADES DIFERENTES"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 8

End Sub

Private Sub PonerContRegIndicador()
'si estamos en modo ver registros muestra el numero de registro en el que estamos
'situados del total de registros mostrados: 1 de 24
Dim cadReg As String

    If (Modo = 2 Or Modo = 0) Then
        cadReg = PonerContRegistros(Me.Adoaux(1))
        If CadB = "" Then
            lblIndicador.Caption = cadReg
        Else
            lblIndicador.Caption = "BUSQUEDA: " & cadReg
        End If
    End If
End Sub



Private Function ActualizarEntradasValsur() As Boolean
Dim Rs As ADODB.Recordset
Dim Rs2 As ADODB.Recordset
Dim I As Integer
Dim SQL As String
Dim Sql2 As String
Dim Sql3 As String

Dim FactCorrDest As Currency
Dim CalDestrio As Currency  ' calidad de destrio de la variedad
Dim CalDestrio2 As Currency ' segunda calidad de destrio
Dim KilDestrio As Currency
Dim KilMuestra As Currency
Dim KilosTot As Currency

Dim KilosDes As Currency    ' kilos destrio de la entrada automatica
Dim KilosPod As Currency    ' kilos podridos de la entrada automatica
Dim KilosNet As Currency    ' kilos netos de la entrada automatica
Dim KilosPeq As Currency    ' kilos pequeños de la entrada automatica
Dim KilosDes2 As Currency   ' 20% de los kilos de destrio de la entrada automatica para la calidad de destrio 2
Dim KilosDes3 As Currency   ' kilos de destrio de la clasificacion

Dim Kilos As Currency

Dim KilosEntrada As Currency  ' kilos netos de la entrada (rclasifica)

Dim UltCalidad As Currency
Dim PrimCalidad As Currency

Dim KilosNeto As String
Dim TotalKilos As String

    On Error GoTo eActualizarEntradasValsur

    conn.BeginTrans
    
    ActualizarEntradasValsur = False
    
    SQL = "select * from rclasifauto order by numnotac"
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not Rs.EOF
        KilosDes = DBLet(Rs!KilosDes, "N")
        KilosPod = DBLet(Rs!KilosPod, "N")
        KilosNet = DBLet(Rs!KilosNet, "N")
        KilosPeq = DBLet(Rs!KilosPeq, "N")
    
        Sql2 = "update rclasifauto_clasif set kiloscal = kiloscal * (" & DBSet((KilosPeq - KilosDes - KilosPod), "N")
        Sql2 = Sql2 & ") / " & DBSet(Rs!KilosNet, "N")
        Sql2 = Sql2 & " where numnotac = " & DBSet(Rs!numnotac, "N")
        Sql2 = Sql2 & " and codvarie = " & DBSet(Rs!codvarie, "N")
        
        conn.Execute Sql2
        
'de momento comentado pq no tienen 2da calidad de destrio
        KilosDes2 = 0
'        KilosDes2 = Round2(KilosDes * 20 / 100, 0)
        
        ' calidad de destrio
        Sql2 = "select codcalid from rcalidad where codvarie = " & DBSet(Rs!codvarie, "N")
        Sql2 = Sql2 & " and tipcalid = 1 "
        CalDestrio = DevuelveValor(Sql2)

'de momento comentado pq no tienen 2da calidad de destrio
'        ' segunda calidad de destrio
'        Sql2 = "select codcalid from rcalidad where codvarie = " & DBSet(Rs!CodVarie, "N")
'        Sql2 = Sql2 & " and tipcalid = 2 "
'        CalDestrio2 = DevuelveValor(Sql2)
        
        If CalDestrio <> 0 Then
            Sql2 = "update rclasifauto_clasif set kilocal = kiloscal + " & DBSet(KilosPod + KilosDes2, "N")
            Sql2 = Sql2 & " where numnotac = " & DBSet(Rs!numnotac, "N") & " and codvarie = " & DBSet(Rs!codvarie, "N")
            Sql2 = Sql2 & " and codcalid = " & DBSet(CalDestrio, "N")
            
            conn.Execute Sql2
        End If
'de momento comentado pq no tienen 2da calidad de destrio
'        If CalDestrio2 <> 0 Then
'            Sql2 = "update rclasifauto_clasif set kilocal = kiloscal + " & DBSet(KilosPod - KilosDes2, "N")
'            Sql2 = Sql2 & " where numnotac = " & DBSet(Rs!numnotac, "N") & " and codvarie = " & DBSet(Rs!CodVarie, "N")
'            Sql2 = Sql2 & " and codcalid = " & DBSet(CalDestrio2, "N")
'
'            conn.Execute Sql2
'        End If
    
        ' kilos de la entrada
        Sql2 = "select kilosnet from rclasifica where numnotac = " & DBSet(Rs!numnotac, "N")
        KilosEntrada = DevuelveValor(Sql2)
    
        KilMuestra = KilosPeq
        
        If KilMuestra <> 0 Then
            Sql2 = "select * from rclasifauto_clasif where numnotac = " & DBSet(Rs!numnotac, "N")
            Sql2 = Sql2 & " and codvarie = " & DBSet(Rs!codvarie, "N")
            Sql2 = Sql2 & " order by codcalid "
        
            Set Rs2 = New ADODB.Recordset
            Rs2.Open Sql2, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            
            KilosTot = 0
            While Not Rs2.EOF
                UltCalidad = Rs2!codcalid
            
                Kilos = Round2(KilosEntrada * DBLet(Rs2!KilosCal, "N") / KilMuestra, 0)
                KilosTot = KilosTot + Kilos
            
                SQL = "select count(*) from rclasifica_clasif where numnotac = " & DBSet(Rs!numnotac, "N")
                SQL = SQL & " and codvarie = " & DBSet(Rs!codvarie, "N")
                SQL = SQL & " and codcalid = " & DBSet(Rs2!codcalid, "N")
                
                If TotalRegistros(SQL) = 0 Then
                    SQL = "insert into rclasifica_clasif (numnotac, codvarie, codcalid, muestra, kilosnet) "
                    SQL = SQL & " values (" & DBSet(Rs!numnotac, "N") & "," & DBSet(Rs!codvarie, "N")
                    SQL = SQL & "," & DBSet(Rs2!codcalid, "N") & "," & DBSet(Rs!KilosCal, "N")
                    SQL = SQL & "," & DBSet(Kilos, "N") & ")"
                    
                    conn.Execute SQL
                Else
                    SQL = "update rclasifica_clasif set muestra = " & DBSet(Rs2!KilosCal, "N") & ","
                    SQL = SQL & " kilosnet = " & DBSet(Kilos, "N")
                    SQL = SQL & " where numnotac = " & DBSet(Rs!numnotac, "N")
                    SQL = SQL & " and codvarie = " & DBSet(Rs!codvarie, "N")
                    SQL = SQL & " and codcalid = " & DBSet(Rs2!codcalid, "N")
                
                    conn.Execute SQL
                End If
                
                Rs2.MoveNext
            Wend
            
            Set Rs2 = Nothing
            
            ' borramos las lineas de clasificacion que no tienen calidad
            SQL = "delete from rclasifica_clasif where numnotac = " & DBSet(Rs!numnotac, "N")
            SQL = SQL & " and muestra is null "
            
            conn.Execute SQL
            
            
            ' si la diferencia es positiva se suma a la ultima calidad
            If KilosEntrada - KilosTot > 0 Then
                SQL = "update rclasifica_clasif set kilosnet = kilosnet + " & DBSet(KilosEntrada - KilosTot, "N")
                SQL = SQL & " where numnotac = " & DBSet(Rs!numnotac, "N")
                SQL = SQL & " and codvarie = " & DBSet(Rs!codvarie, "N")
                SQL = SQL & " and codcalid = " & DBSet(UltCalidad, "N")
                
                conn.Execute SQL
            Else
            ' si es negativa a la primera que no deje el importe negqativo
                SQL = "select min(codcalid) from rclasifica_clasif "
                SQL = SQL & " where numnotac = " & DBSet(Rs!numnotac, "N")
                SQL = SQL & " and codvarie = " & DBSet(Rs!codvarie, "N")
                SQL = SQL & " and kiloscal >= " & DBSet(KilosEntrada - KilosTot, "N")
                
                PrimCalidad = DevuelveValor(SQL)
                
                SQL = "update rclasifica_clasif set kilosnet = kilosnet + " & DBSet(KilosEntrada - KilosTot, "N")
                SQL = SQL & " where numnotac = " & DBSet(Rs!numnotac, "N")
                SQL = SQL & " and codvarie = " & DBSet(Rs!codvarie, "N")
                SQL = SQL & " and codcalid = " & DBSet(PrimCalidad, "N")
                
                conn.Execute SQL
            End If
        End If
        
        If CalDestrio <> 0 Then
            ' factor de correccion de destrio
            Sql2 = "select facorrde from variedades where codvarie = " & DBSet(Rs!codvarie, "N")
            FactCorrDest = DevuelveValor(Sql2)
            
            KilDestrio = Round2(FactCorrDest * KilosEntrada / 100, 0)
            
            Sql3 = "select kiloscal from rclasifica_clasif where numnotac = " & DBSet(Rs!numnotac, "N")
            Sql3 = Sql3 & " and codvarie = " & DBSet(Rs!codvarie, "N") & " and codcalid = "
            Sql3 = Sql3 & DBSet(CalDestrio, "N")
            
            KilosDes3 = DevuelveValor(Sql3)
            
            KilosNet = KilosEntrada - KilosDes3
        
            KilosDes3 = KilosDes3 + KilDestrio
            
            Sql3 = "update rclasifica_clasif set kilosnet = kilosnet + " & DBSet(KilosDes3, "N")
            Sql3 = Sql3 & " where numnotac = " & DBSet(Rs!numnotac, "N")
            Sql3 = Sql3 & " and codvarie = " & DBSet(Rs!codvarie, "N")
            Sql3 = Sql3 & " and codcalid = " & DBSet(CalDestrio, "N")
            conn.Execute Sql3
            
            ' el resto de calidades
            Sql3 = "update rclasifica_clasif set kilosnet = kilosnet - round(kilosnet *"
            Sql3 = Sql3 & DBSet(KilDestrio, "N") & " / " & DBSet(KilosNet, "N") & "0) "
            Sql3 = Sql3 & " where numnotac = " & DBSet(Rs!numnotac, "N")
            Sql3 = Sql3 & " and codvarie = " & DBSet(Rs!codvarie, "N")
            Sql3 = Sql3 & " and codcalid <> " & CalDestrio
            conn.Execute Sql3
            
            Sql3 = "select sum(kilosnet) from rclasifica_clasif where numnotac = " & DBSet(Rs!numnotac, "N")
            Sql3 = Sql3 & " and codvarie " & DBSet(Rs!codvarie, "N")
            
            TotalKilos = DevuelveValor(Sql3)
            
            If KilosEntrada - TotalKilos > 0 Then
                ' si la diferencia es positiva va a la ultima calidad
                Sql3 = "update rclasifica_clasif set kilosnet = kilosnet + " & DBSet(KilosEntrada - TotalKilos, "N")
                Sql3 = Sql3 & " where numnotac = " & DBSet(Rs!numnotac, "N")
                Sql3 = Sql3 & " and codvarie = " & DBSet(Rs!codvarie, "N")
                Sql3 = Sql3 & " and codcalid = " & DBSet(UltCalidad, "N")
                conn.Execute Sql3
            End If
            If KilosEntrada - TotalKilos < 0 Then
                ' si la diferencia es negativa va a la primera calidad que se pueda
                Sql3 = "select min(codcalid) from rclasifica_clasif "
                Sql3 = Sql3 & " where numnotac = " & DBSet(Rs!numnotac, "N")
                Sql3 = Sql3 & " and codvarie = " & DBSet(Rs!codvarie, "N")
                Sql3 = Sql3 & " and kiloscal >= " & DBSet(KilosEntrada - TotalKilos, "N")
                
                PrimCalidad = DevuelveValor(SQL)
                
                Sql3 = "update rclasifica_clasif set kilosnet = kilosnet + " & DBSet(KilosEntrada - TotalKilos, "N")
                Sql3 = Sql3 & " where numnotac = " & DBSet(Rs!numnotac, "N")
                Sql3 = Sql3 & " and codvarie = " & DBSet(Rs!codvarie, "N")
                Sql3 = Sql3 & " and codcalid = " & DBSet(PrimCalidad, "N")
                
                conn.Execute SQL
            End If
            
            ' para todas las calidades que no sean de destrio con kilos negativos --> se ponen a 0
            Sql3 = "update rclasifica_clasif set kilosnet = 0 where numnotac = " & DBSet(Rs!numnotac, "N")
            Sql3 = Sql3 & " and codvarie = " & DBSet(Rs!codvarie, "N")
            Sql3 = Sql3 & " and codcalid <> " & DBSet(CalDestrio, "N")
            Sql3 = Sql3 & " and kilosnet < 0 "
            conn.Execute Sql3
            
            ' para la calidad de destrio, si kilos > kilos muestreados --> se pone kilos muestreados
            Sql3 = "update rclasifica_clasif set kilosnet = " & DBSet(KilosEntrada, "N")
            Sql3 = Sql3 & " where numnotac = " & DBSet(Rs!numnotac, "N")
            Sql3 = Sql3 & " and codvarie = " & DBSet(Rs!codvarie, "N")
            Sql3 = Sql3 & " and codcalid = " & DBSet(CalDestrio, "N")
            Sql3 = Sql3 & " and kilosnet > " & DBSet(KilosEntrada, "N")
            conn.Execute Sql3
            
            ' para todas las calidades
            Sql3 = "update rclasifica_clasif set muestra = round(kilosnet * " & DBSet(KilMuestra, "N")
            Sql3 = Sql3 & " / " & DBSet(KilosEntrada, "N") & ",2)"
            Sql3 = Sql3 & " where numnotac = " & DBSet(Rs!numnotac, "N")
            Sql3 = Sql3 & " and codvarie = " & DBSet(Rs!codvarie, "N")
            conn.Execute Sql3
        Else
            ' no existe la calidad de destrio damos un error
            MsgBox "No existe calidad de destrio para la variedad " & DBLet(Rs!codvarie, "N") & ". Revise.", vbExclamation
            Exit Function
        End If
        
        SQL = "delete from rclasifauto_clasif where numnotac = " & DBSet(Rs!numnotac, "N")
        SQL = SQL & " and codvarie = " & DBSet(Rs!codvarie, "N")
        conn.Execute SQL

        SQL = "delete from rclasifauto where numnotac = " & DBSet(Rs!numnotac, "N")
        SQL = SQL & " and codvarie = " & DBSet(Rs!codvarie, "N")
        conn.Execute SQL
    
        Rs.MoveNext
    Wend

    Set Rs = Nothing

    ActualizarEntradasValsur = True
    conn.CommitTrans
    Exit Function


eActualizarEntradasValsur:
    If Err.Number <> 0 Then
        conn.RollbackTrans
        MuestraError Err.Number, "Actualizar entradas", Err.Description
    End If
End Function

Private Function ActualizarEntradasCatadau() As Boolean
Dim Rs As ADODB.Recordset
Dim Rs2 As ADODB.Recordset
Dim RsGastos As ADODB.Recordset
Dim I As Integer
Dim SQL As String
Dim Sql2 As String

Dim KilosNet As Currency
Dim FactCorrDest As Currency
Dim CalDestrio As Currency
Dim CalPodrido As Currency
Dim KilDestrio As Currency
Dim KilMuestra As Currency
Dim KilPodrido As Currency
Dim KilosTot As Currency
Dim Kilos As Currency

Dim UltCalidad As Currency
Dim PrimCalidad As Currency

Dim B As Boolean
Dim cadErr As String
Dim EntClasif As String

Dim KilosDestrioMerma As Currency

    On Error GoTo eActualizarEntradasCatadau

    conn.BeginTrans
    
    ActualizarEntradasCatadau = False
    
    SQL = "select * from rclasifauto order by numnotac"
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    B = True
    
    EntClasif = ""
    While Not Rs.EOF And B
    
        If EntradaClasificada(DBLet(Rs!numnotac)) Then
            EntClasif = EntClasif & DBLet(Rs!numnotac) & ", "
        Else
        
            ' kilos de la entrada
            Sql2 = "select kilosnet from rclasifica where numnotac = " & DBSet(Rs!numnotac, "N")
            KilosNet = DevuelveValor(Sql2)
            
            ' si hay kilos de destrio recalculamos
            KilDestrio = CCur(DBLet(Rs!KilosDes, "N")) + CCur(DBLet(Rs!KilosPeq, "N"))
            KilPodrido = CCur(DBLet(Rs!KilosPod, "N"))
            If KilDestrio <> 0 Then
                ' factor de correccion de destrio
                Sql2 = "select facorrde from variedades where codvarie = " & DBSet(Rs!codvarie, "N")
                FactCorrDest = DevuelveValor(Sql2)
    
                ' calidad de destrio
                Sql2 = "select codcalid from rcalidad where codvarie = " & DBSet(Rs!codvarie, "N")
                Sql2 = Sql2 & " and tipcalid = 1 "
                CalDestrio = DevuelveValor(Sql2)
                
                If CalDestrio = 0 Then
                    ' no existe la calidad de destrio damos un error
                    MsgBox "No existe calidad de destrio para la variedad " & DBLet(Rs!codvarie, "N") & ". Revise.", vbExclamation
                    Exit Function
                End If
                
                ' actualizamos el muestreo de la calidad de destrio
                Sql2 = "select count(*) from rclasifauto_clasif where numnotac = " & DBSet(Rs!numnotac, "N")
                Sql2 = Sql2 & " and codvarie = " & DBSet(Rs!codvarie, "N")
                Sql2 = Sql2 & " and codcalid = " & DBSet(CalDestrio, "N")
                
                If TotalRegistros(Sql2) = 0 Then
                    ' si en la clasificacion no hay calidad de destrio, la creamos
                    Sql2 = "insert into rclasifauto_clasif (numnotac, codvarie, codcalid, kiloscal) values ("
                    Sql2 = Sql2 & DBSet(Rs!numnotac, "N") & "," & DBSet(Rs!codvarie, "N") & ","
                    Sql2 = Sql2 & DBSet(CalDestrio, "N") & "," & DBSet(KilDestrio, "N") & ")"
                    
                    conn.Execute Sql2
                Else
                    ' si en la clasificacion hay calidad de destrio, la actualizamos
                    Sql2 = "update rclasifauto_clasif set kiloscal = kiloscal + " & DBSet(KilDestrio, "N")
                    Sql2 = Sql2 & " where numnotac = " & DBSet(Rs!numnotac, "N")
                    Sql2 = Sql2 & " and codvarie = " & DBSet(Rs!codvarie, "N")
                    Sql2 = Sql2 & " and codcalid = " & DBSet(CalDestrio, "N")
                
                    conn.Execute Sql2
                End If
                
                ' multiplicamos los kilos de destrio por el factor de correccion
                Sql2 = "update rclasifauto_clasif set kiloscal = round(kiloscal * " & DBSet(FactCorrDest, "N") & ",2)"
                Sql2 = Sql2 & " where numnotac = " & DBSet(Rs!numnotac, "N")
                Sql2 = Sql2 & " and codvarie = " & DBSet(Rs!codvarie, "N")
                Sql2 = Sql2 & " and codcalid = " & DBSet(CalDestrio, "N")
                
                conn.Execute Sql2
            End If
        
            If KilPodrido <> 0 Then
                ' factor de correccion de podrido o mermas distinto del de destrio
                Sql2 = "select facorrme from variedades where codvarie = " & DBSet(Rs!codvarie, "N")
                FactCorrDest = DevuelveValor(Sql2)
    
                ' calidad de podrido o merma
                Sql2 = "select codcalid from rcalidad where codvarie = " & DBSet(Rs!codvarie, "N")
                Sql2 = Sql2 & " and tipcalid = 3 "
                CalPodrido = DevuelveValor(Sql2)
                
                If CalPodrido = 0 Then
                    ' no existe la calidad de podrido o merma damos un error
                    MsgBox "No existe calidad de podrido o merma para la variedad " & DBLet(Rs!codvarie, "N") & ". Revise.", vbExclamation
                    Exit Function
                End If
                
                ' actualizamos el muestreo de la calidad de podrido o merma
                Sql2 = "select count(*) from rclasifauto_clasif where numnotac = " & DBSet(Rs!numnotac, "N")
                Sql2 = Sql2 & " and codvarie = " & DBSet(Rs!codvarie, "N")
                Sql2 = Sql2 & " and codcalid = " & DBSet(CalPodrido, "N")
                
                If TotalRegistros(Sql2) = 0 Then
                    ' si en la clasificacion no hay calidad de podrido o merma, la creamos
                    Sql2 = "insert into rclasifauto_clasif (numnotac, codvarie, codcalid, kiloscal) values ("
                    Sql2 = Sql2 & DBSet(Rs!numnotac, "N") & "," & DBSet(Rs!codvarie, "N") & ","
                    Sql2 = Sql2 & DBSet(CalPodrido, "N") & "," & DBSet(KilPodrido, "N") & ")"
                    
                    conn.Execute Sql2
                Else
                    ' si en la clasificacion hay calidad de podrido o merma, la actualizamos
                    Sql2 = "update rclasifauto_clasif set kiloscal = kiloscal + " & DBSet(KilPodrido, "N")
                    Sql2 = Sql2 & " where numnotac = " & DBSet(Rs!numnotac, "N")
                    Sql2 = Sql2 & " and codvarie = " & DBSet(Rs!codvarie, "N")
                    Sql2 = Sql2 & " and codcalid = " & DBSet(CalPodrido, "N")
                
                    conn.Execute Sql2
                End If
                
                ' multiplicamos los kilos de podrido/merma por el factor de correccion
                Sql2 = "update rclasifauto_clasif set kiloscal = round(kiloscal * " & DBSet(FactCorrDest, "N") & ",2)"
                Sql2 = Sql2 & " where numnotac = " & DBSet(Rs!numnotac, "N")
                Sql2 = Sql2 & " and codvarie = " & DBSet(Rs!codvarie, "N")
                Sql2 = Sql2 & " and codcalid = " & DBSet(CalPodrido, "N")
                
                conn.Execute Sql2
        
            End If
        
        
            '[Monica]14/10/2011: añadido la variable KilosDestrioMerma : Kilos que no se prorratean (de destrio y de merma)
            Sql2 = "select sum(kiloscal) from rclasifauto_clasif where numnotac = " & DBSet(Rs!numnotac, "N")
            Sql2 = Sql2 & " and codvarie = " & DBSet(Rs!codvarie, "N")
            Sql2 = Sql2 & " and codcalid in (select codcalid from rcalidad where codvarie = " & DBSet(Rs!codvarie, "N")
            Sql2 = Sql2 & " and tipcalid in (1,3)) " ' muestras que sean de destrio y de merma
            KilosDestrioMerma = DevuelveValor(Sql2)
            
        
            Sql2 = "select sum(kiloscal) from rclasifauto_clasif where numnotac = " & DBSet(Rs!numnotac, "N")
            Sql2 = Sql2 & " and codvarie = " & DBSet(Rs!codvarie, "N")
            '[Monica]14/10/2011: añadimos en este punto que no sean calidades de destrio ni de merma
            Sql2 = Sql2 & " and codcalid not in (select codcalid from rcalidad where codvarie = " & DBSet(Rs!codvarie, "N")
            Sql2 = Sql2 & " and tipcalid in (1,3)) " ' muestras que no sean de destrio ni de merma
            
            KilMuestra = DevuelveValor(Sql2)
            If KilMuestra <> 0 Then
                Sql2 = "select * from rclasifauto_clasif where numnotac = " & DBSet(Rs!numnotac, "N")
                Sql2 = Sql2 & " and codvarie = " & DBSet(Rs!codvarie, "N")
                Sql2 = Sql2 & " order by codcalid "
            
                Set Rs2 = New ADODB.Recordset
                Rs2.Open Sql2, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                
                
                KilosTot = 0
                While Not Rs2.EOF
                    UltCalidad = Rs2!codcalid
                
                    '[Monica]14/10/2011: modificacion de dejar los kilos sin prorratear
                    If EsCalidadDestrio(CStr(DBLet(Rs!codvarie, "N")), CStr(DBLet(Rs2!codcalid, "N"))) Or _
                       EsCalidadMerma(CStr(DBLet(Rs!codvarie, "N")), CStr(DBLet(Rs2!codcalid, "N"))) Then
                        
                        Kilos = Round2(DBLet(Rs2!KilosCal, "N"), 0)
                        
                    Else
                        Kilos = Round2((KilosNet - KilosDestrioMerma) * DBLet(Rs2!KilosCal, "N") / KilMuestra, 0)
                        
                    End If
'                   Kilos = Round2(KilosNet * DBLet(Rs2!KilosCal, "N") / KilMuestra, 0)           antes estaba así
                    KilosTot = KilosTot + Kilos
                
                    SQL = "select count(*) from rclasifica_clasif where numnotac = " & DBSet(Rs!numnotac, "N")
                    SQL = SQL & " and codvarie = " & DBSet(Rs!codvarie, "N")
                    SQL = SQL & " and codcalid = " & DBSet(Rs2!codcalid, "N")
                    
                    If TotalRegistros(SQL) = 0 Then
                        SQL = "insert into rclasifica_clasif (numnotac, codvarie, codcalid, muestra, kilosnet) "
                        SQL = SQL & " values (" & DBSet(Rs!numnotac, "N") & "," & DBSet(Rs!codvarie, "N")
                        SQL = SQL & "," & DBSet(Rs2!codcalid, "N") & "," & DBSet(Rs2!KilosCal, "N")
                        SQL = SQL & "," & DBSet(Kilos, "N") & ")"
                        
                        conn.Execute SQL
                    Else
                        SQL = "update rclasifica_clasif set muestra = " & DBSet(Rs2!KilosCal, "N") & ","
                        SQL = SQL & " kilosnet = " & DBSet(Kilos, "N")
                        SQL = SQL & " where numnotac = " & DBSet(Rs!numnotac, "N")
                        SQL = SQL & " and codvarie = " & DBSet(Rs!codvarie, "N")
                        SQL = SQL & " and codcalid = " & DBSet(Rs2!codcalid, "N")
                    
                        conn.Execute SQL
                    End If
                    
                    Rs2.MoveNext
                Wend
                
                Set Rs2 = Nothing
                
                ' si la diferencia es positiva se suma a la ultima calidad
                If KilosNet - KilosTot > 0 Then
                    SQL = "update rclasifica_clasif set kilosnet = kilosnet + " & DBSet(KilosNet - KilosTot, "N")
                    SQL = SQL & " where numnotac = " & DBSet(Rs!numnotac, "N")
                    SQL = SQL & " and codvarie = " & DBSet(Rs!codvarie, "N")
                    SQL = SQL & " and codcalid = " & DBSet(UltCalidad, "N")
                    
                    conn.Execute SQL
                Else
                ' si es negativa a la primera
                    SQL = "select min(codcalid) from rclasifica_clasif "
                    SQL = SQL & " where numnotac = " & DBSet(Rs!numnotac, "N")
                    SQL = SQL & " and codvarie = " & DBSet(Rs!codvarie, "N")
                    SQL = SQL & " and kilosnet >= " & DBSet((KilosNet - KilosTot) * (-1), "N")
                    
                    PrimCalidad = DevuelveValor(SQL)
                    
                    SQL = "update rclasifica_clasif set kilosnet = kilosnet + " & DBSet(KilosNet - KilosTot, "N")
                    SQL = SQL & " where numnotac = " & DBSet(Rs!numnotac, "N")
                    SQL = SQL & " and codvarie = " & DBSet(Rs!codvarie, "N")
                    SQL = SQL & " and codcalid = " & DBSet(PrimCalidad, "N")
                    
                    conn.Execute SQL
                End If
            End If
        
            SQL = "delete from rclasifica_clasif where numnotac = " & DBSet(Rs!numnotac, "N")
            SQL = SQL & " and codvarie = " & DBSet(Rs!codvarie, "N") & " and kilosnet is null "
            conn.Execute SQL
            
            SQL = "delete from rclasifauto_clasif where numnotac = " & DBSet(Rs!numnotac, "N")
            SQL = SQL & " and codvarie = " & DBSet(Rs!codvarie, "N")
            conn.Execute SQL
            
            SQL = "delete from rclasifauto where numnotac = " & DBSet(Rs!numnotac, "N")
            SQL = SQL & " and codvarie = " & DBSet(Rs!codvarie, "N")
            conn.Execute SQL
            
            '++ 20-05-2009: calculamos los gastos de recoleccion para la entrada clasificadaç
            SQL = "select * from rclasifica where numnotac = " & DBSet(Rs!numnotac, "N")
            
            Set RsGastos = New ADODB.Recordset
            RsGastos.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            
            If Not RsGastos.EOF Then
                cadErr = "Actualizando Gastos"
                B = ActualizarGastos(RsGastos, cadErr)
            End If
            
            Set RsGastos = Nothing
            '++
        
        End If
        
        
        Rs.MoveNext
    Wend

    Set Rs = Nothing
    
    
    If EntClasif <> "" Then
        MsgBox "Las siguientes notas no han sido actualizadas, porque tenían clasificacion. Revise." & _
            vbCrLf & vbCrLf & Mid(EntClasif, 1, Len(EntClasif) - 2), vbExclamation
    End If

    If B Then
        ActualizarEntradasCatadau = True
        conn.CommitTrans
        Exit Function
    End If

eActualizarEntradasCatadau:
    If Err.Number <> 0 Or Not B Then
        conn.RollbackTrans
        MuestraError Err.Number, "Actualizar entradas", Err.Description & cadErr
    End If
End Function



Private Function ActualizarEntradasAlzira() As Boolean
Dim Rs As ADODB.Recordset
Dim Rs2 As ADODB.Recordset
Dim RsGastos As ADODB.Recordset
Dim I As Integer
Dim SQL As String
Dim Sql2 As String

Dim KilosNet As Currency
Dim FactCorrDest As Currency
Dim CalDestrio As Currency
Dim CalPodrido As Currency
Dim KilDestrio As Currency
Dim KilMuestra As Currency
Dim KilPodrido As Currency
Dim KilosTot As Currency
Dim Kilos As Currency

Dim UltCalidad As Currency
Dim PrimCalidad As Currency

Dim B As Boolean
Dim cadErr As String

    On Error GoTo eActualizarEntradasAlzira

    conn.BeginTrans
    
    ActualizarEntradasAlzira = False
    
    SQL = "select * from rclasifauto order by numnotac"
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    B = True
    
    While Not Rs.EOF And B
    
        ' kilos de la entrada
        Sql2 = "select kilosnet from rclasifica where numnotac = " & DBSet(Rs!numnotac, "N")
        KilosNet = DevuelveValor(Sql2)
        
'**** DE MOMENTO NO CALCULAMOS EL FACTOR DE CORRECCION SOBRE LOS KILOS DE DESTRIO
'**** NI SOBRE LOS KILOS DE MERMA : PREGUNTAR A MANOLO.???????
        
'        ' si hay kilos de destrio recalculamos
'        KilDestrio = CCur(DBLet(Rs!KilosDes, "N"))
'        KilPodrido = CCur(DBLet(Rs!KilosPod, "N"))
'        If KilDestrio <> 0 Then
'            ' factor de correccion de destrio
'            Sql2 = "select facorrde from variedades where codvarie = " & DBSet(Rs!CodVarie, "N")
'            FactCorrDest = DevuelveValor(Sql2)
'
'            ' calidad de destrio
'            Sql2 = "select codcalid from rcalidad where codvarie = " & DBSet(Rs!CodVarie, "N")
'            Sql2 = Sql2 & " and tipcalid = 1 "
'            CalDestrio = DevuelveValor(Sql2)
'
'            If CalDestrio = 0 Then
'                ' no existe la calidad de destrio damos un error
'                MsgBox "No existe calidad de destrio para la variedad " & DBLet(Rs!CodVarie, "N") & ". Revise.", vbExclamation
'                Exit Function
'            End If
'
'            ' multiplicamos los kilos de destrio por el factor de correccion
'            Sql2 = "update rclasifauto_clasif set kiloscal = round(kiloscal * " & DBSet(FactCorrDest, "N") & ",2)"
'            Sql2 = Sql2 & " where numnotac = " & DBSet(Rs!numnotac, "N")
'            Sql2 = Sql2 & " and codvarie = " & DBSet(Rs!CodVarie, "N")
'            Sql2 = Sql2 & " and codcalid = " & DBSet(CalDestrio, "N")
'
'            conn.Execute Sql2
'        End If
'
'        If KilPodrido <> 0 Then
'            ' factor de correccion de podrido o mermas distinto del de destrio
'            Sql2 = "select facorrme from variedades where codvarie = " & DBSet(Rs!CodVarie, "N")
'            FactCorrDest = DevuelveValor(Sql2)
'
'            ' calidad de podrido o merma
'            Sql2 = "select codcalid from rcalidad where codvarie = " & DBSet(Rs!CodVarie, "N")
'            Sql2 = Sql2 & " and tipcalid = 3 "
'            CalPodrido = DevuelveValor(Sql2)
'
'            If CalPodrido = 0 Then
'                ' no existe la calidad de podrido o merma damos un error
'                MsgBox "No existe calidad de podrido o merma para la variedad " & DBLet(Rs!CodVarie, "N") & ". Revise.", vbExclamation
'                Exit Function
'            End If
'
'            ' actualizamos el muestreo de la calidad de podrido o merma
'            Sql2 = "select count(*) from rclasifauto_clasif where numnotac = " & DBSet(Rs!numnotac, "N")
'            Sql2 = Sql2 & " and codvarie = " & DBSet(Rs!CodVarie, "N")
'            Sql2 = Sql2 & " and codcalid = " & DBSet(CalPodrido, "N")
'
'            If TotalRegistros(Sql2) = 0 Then
'                ' si en la clasificacion no hay calidad de podrido o merma, la creamos
'                Sql2 = "insert into rclasifauto_clasif (numnotac, codvarie, codcalid, kiloscal) values ("
'                Sql2 = Sql2 & DBSet(Rs!numnotac, "N") & "," & DBSet(Rs!CodVarie, "N") & ","
'                Sql2 = Sql2 & DBSet(CalPodrido, "N") & "," & DBSet(KilPodrido, "N") & ")"
'
'                conn.Execute Sql2
'            Else
'                ' si en la clasificacion hay calidad de podrido o merma, la actualizamos
'                Sql2 = "update rclasifauto_clasif set kiloscal = kiloscal + " & DBSet(KilPodrido, "N")
'                Sql2 = Sql2 & " where numnotac = " & DBSet(Rs!numnotac, "N")
'                Sql2 = Sql2 & " and codvarie = " & DBSet(Rs!CodVarie, "N")
'                Sql2 = Sql2 & " and codcalid = " & DBSet(CalPodrido, "N")
'
'                conn.Execute Sql2
'            End If
'
'            ' multiplicamos los kilos de podrido/merma por el factor de correccion
'            Sql2 = "update rclasifauto_clasif set kiloscal = round(kiloscal * " & DBSet(FactCorrDest, "N") & ",2)"
'            Sql2 = Sql2 & " where numnotac = " & DBSet(Rs!numnotac, "N")
'            Sql2 = Sql2 & " and codvarie = " & DBSet(Rs!CodVarie, "N")
'            Sql2 = Sql2 & " and codcalid = " & DBSet(CalPodrido, "N")
'
'            conn.Execute Sql2
'
'        End If
    
        Sql2 = "select sum(kiloscal) from rclasifauto_clasif where numnotac = " & DBSet(Rs!numnotac, "N")
        Sql2 = Sql2 & " and codvarie = " & DBSet(Rs!codvarie, "N")
        
        KilMuestra = DevuelveValor(Sql2)
        If KilMuestra <> 0 Then
            Sql2 = "select * from rclasifauto_clasif where numnotac = " & DBSet(Rs!numnotac, "N")
            Sql2 = Sql2 & " and codvarie = " & DBSet(Rs!codvarie, "N")
            Sql2 = Sql2 & " order by codcalid "
        
            Set Rs2 = New ADODB.Recordset
            Rs2.Open Sql2, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            
            
            KilosTot = 0
            While Not Rs2.EOF
                UltCalidad = Rs2!codcalid
            
                Kilos = Round2(KilosNet * DBLet(Rs2!KilosCal, "N") / KilMuestra, 0)
                KilosTot = KilosTot + Kilos
            
                SQL = "select count(*) from rclasifica_clasif where numnotac = " & DBSet(Rs!numnotac, "N")
                SQL = SQL & " and codvarie = " & DBSet(Rs!codvarie, "N")
                SQL = SQL & " and codcalid = " & DBSet(Rs2!codcalid, "N")
                
                If TotalRegistros(SQL) = 0 Then
                    SQL = "insert into rclasifica_clasif (numnotac, codvarie, codcalid, muestra, kilosnet) "
                    SQL = SQL & " values (" & DBSet(Rs!numnotac, "N") & "," & DBSet(Rs!codvarie, "N")
                    SQL = SQL & "," & DBSet(Rs2!codcalid, "N") & "," & DBSet(Rs2!KilosCal, "N")
                    SQL = SQL & "," & DBSet(Kilos, "N") & ")"
                    
                    conn.Execute SQL
                Else
                    SQL = "update rclasifica_clasif set muestra = " & DBSet(Rs2!KilosCal, "N") & ","
                    SQL = SQL & " kilosnet = " & DBSet(Kilos, "N")
                    SQL = SQL & " where numnotac = " & DBSet(Rs!numnotac, "N")
                    SQL = SQL & " and codvarie = " & DBSet(Rs!codvarie, "N")
                    SQL = SQL & " and codcalid = " & DBSet(Rs2!codcalid, "N")
                
                    conn.Execute SQL
                End If
                
                Rs2.MoveNext
            Wend
            
            Set Rs2 = Nothing
            
            ' si la diferencia es positiva se suma a la ultima calidad
            If KilosNet - KilosTot > 0 Then
                SQL = "update rclasifica_clasif set kilosnet = kilosnet + " & DBSet(KilosNet - KilosTot, "N")
                SQL = SQL & " where numnotac = " & DBSet(Rs!numnotac, "N")
                SQL = SQL & " and codvarie = " & DBSet(Rs!codvarie, "N")
                SQL = SQL & " and codcalid = " & DBSet(UltCalidad, "N")
                
                conn.Execute SQL
            Else
            ' si es negativa a la primera
                SQL = "select min(codcalid) from rclasifica_clasif "
                SQL = SQL & " where numnotac = " & DBSet(Rs!numnotac, "N")
                SQL = SQL & " and codvarie = " & DBSet(Rs!codvarie, "N")
                SQL = SQL & " and kilosnet >= " & DBSet((KilosNet - KilosTot) * (-1), "N")
                
                PrimCalidad = DevuelveValor(SQL)
                
                SQL = "update rclasifica_clasif set kilosnet = kilosnet + " & DBSet(KilosNet - KilosTot, "N")
                SQL = SQL & " where numnotac = " & DBSet(Rs!numnotac, "N")
                SQL = SQL & " and codvarie = " & DBSet(Rs!codvarie, "N")
                SQL = SQL & " and codcalid = " & DBSet(PrimCalidad, "N")
                
                conn.Execute SQL
            End If
        End If
    
        SQL = "delete from rclasifica_clasif where numnotac = " & DBSet(Rs!numnotac, "N")
        SQL = SQL & " and codvarie = " & DBSet(Rs!codvarie, "N") & " and kilosnet is null "
        conn.Execute SQL
        
        SQL = "delete from rclasifauto_clasif where numnotac = " & DBSet(Rs!numnotac, "N")
        SQL = SQL & " and codvarie = " & DBSet(Rs!codvarie, "N")
        conn.Execute SQL
        
        SQL = "delete from rclasifauto where numnotac = " & DBSet(Rs!numnotac, "N")
        SQL = SQL & " and codvarie = " & DBSet(Rs!codvarie, "N")
        conn.Execute SQL
        
        '++ 20-05-2009: calculamos los gastos de recoleccion para la entrada clasificadaç
        SQL = "select * from rclasifica where numnotac = " & DBSet(Rs!numnotac, "N")
        
        Set RsGastos = New ADODB.Recordset
        RsGastos.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        If Not RsGastos.EOF Then
            cadErr = "Actualizando Gastos"
            B = ActualizarGastos(RsGastos, cadErr)
        End If
        
        Set RsGastos = Nothing
        '++
    
        Rs.MoveNext
    Wend

    Set Rs = Nothing

    If B Then
        ActualizarEntradasAlzira = True
        conn.CommitTrans
        Exit Function
    End If

eActualizarEntradasAlzira:
    If Err.Number <> 0 Or Not B Then
        conn.RollbackTrans
        MuestraError Err.Number, "Actualizar entradas", Err.Description & cadErr
    End If
End Function


Private Function ActualizarEntradasCastelduc() As Boolean
Dim Rs As ADODB.Recordset
Dim Rs2 As ADODB.Recordset
Dim RsGastos As ADODB.Recordset
Dim I As Integer
Dim SQL As String
Dim Sql2 As String

Dim KilosNet As Currency
Dim FactCorrDest As Currency
Dim CalDestrio As Currency
Dim CalPodrido As Currency
Dim KilDestrio As Currency
Dim KilMuestra As Currency
Dim KilPodrido As Currency
Dim KilosTot As Currency
Dim Kilos As Currency

Dim UltCalidad As Currency
Dim PrimCalidad As Currency

Dim B As Boolean
Dim cadErr As String

Dim EntClasif As String

    On Error GoTo eActualizarEntradasCastelduc

    conn.BeginTrans
    
    ActualizarEntradasCastelduc = False
    
    SQL = "select * from rclasifauto order by numnotac"
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    B = True
    EntClasif = ""
    While Not Rs.EOF And B
        If EntradaClasificada(DBLet(Rs!numnotac)) Then
            EntClasif = EntClasif & DBLet(Rs!numnotac) & ", "
        Else
        
            ' kilos de la entrada
            Sql2 = "select kilosnet from rclasifica where numnotac = " & DBSet(Rs!numnotac, "N")
            KilosNet = DevuelveValor(Sql2)
            
        
            Sql2 = "select sum(kiloscal) from rclasifauto_clasif where numnotac = " & DBSet(Rs!numnotac, "N")
            Sql2 = Sql2 & " and codvarie = " & DBSet(Rs!codvarie, "N")
            
            KilMuestra = DevuelveValor(Sql2)
            
            
            
            If KilMuestra <> 0 Then
                Sql2 = "select * from rclasifauto_clasif where numnotac = " & DBSet(Rs!numnotac, "N")
                Sql2 = Sql2 & " and codvarie = " & DBSet(Rs!codvarie, "N")
                Sql2 = Sql2 & " order by codcalid "
            
                Set Rs2 = New ADODB.Recordset
                Rs2.Open Sql2, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                
                
                KilosTot = 0
                
                '[Monica]25/07/2016
                Sql2 = "select sum(kiloscal) from rclasifauto_clasif, rcalidad where numnotac = " & DBSet(Rs!numnotac, "N")
                Sql2 = Sql2 & " and rclasifauto_clasif.codvarie = " & DBSet(Rs!codvarie, "N")
                Sql2 = Sql2 & " and rclasifauto_clasif.codcalid = rcalidad.codcalid and rcalidad.tipcalid = 1 "
                
                KilDestrio = DevuelveValor("select kilosdes from rclasifauto where numnotac = " & DBSet(Rs!numnotac, "N"))
                
                While Not Rs2.EOF
                    '[Monica] 04/06/2010
                    ' comprobamos si es la calidad de destrio a la que le ponemos el total de kilos
                    SQL = "select count(*) from rcalidad where codvarie = " & DBSet(Rs2!codvarie, "N")
                    SQL = SQL & " and codcalid = " & DBSet(Rs2!codcalid, "N")
                    SQL = SQL & " and tipcalid = 1 "
                    
                    If TotalRegistros(SQL) > 0 Then
                        Kilos = DBLet(Rs2!KilosCal, "N")
                        KilosTot = KilosTot + Kilos
                    Else
                        UltCalidad = Rs2!codcalid
                        '[Monica]25/07/2016: la regla de 3 es sobre los kilos de muestra sin los de destrio
                        Kilos = Round2((KilosNet - KilDestrio) * DBLet(Rs2!KilosCal, "N") / (KilMuestra - KilDestrio), 0)
                        KilosTot = KilosTot + Kilos
                    End If
                    
                    '[Monica] 04/06/2010
                    SQL = "select count(*) from rclasifica_clasif where numnotac = " & DBSet(Rs!numnotac, "N")
                    SQL = SQL & " and codvarie = " & DBSet(Rs!codvarie, "N")
                    SQL = SQL & " and codcalid = " & DBSet(Rs2!codcalid, "N")
                    
                    If TotalRegistros(SQL) = 0 Then
                        SQL = "insert into rclasifica_clasif (numnotac, codvarie, codcalid, muestra, kilosnet) "
                        SQL = SQL & " values (" & DBSet(Rs!numnotac, "N") & "," & DBSet(Rs!codvarie, "N")
                        SQL = SQL & "," & DBSet(Rs2!codcalid, "N") & "," & DBSet(Rs2!KilosCal, "N")
                        SQL = SQL & "," & DBSet(Kilos, "N") & ")"
                        
                        conn.Execute SQL
                    Else
                        SQL = "update rclasifica_clasif set muestra = " & DBSet(Rs2!KilosCal, "N") & ","
                        SQL = SQL & " kilosnet = " & DBSet(Kilos, "N")
                        SQL = SQL & " where numnotac = " & DBSet(Rs!numnotac, "N")
                        SQL = SQL & " and codvarie = " & DBSet(Rs!codvarie, "N")
                        SQL = SQL & " and codcalid = " & DBSet(Rs2!codcalid, "N")
                    
                        conn.Execute SQL
                    End If
                    
                    Rs2.MoveNext
                Wend
                
                Set Rs2 = Nothing
                
'[Monica]22/07/2016: problema que le dio en albaricoques
' si hay diferencia no hacemos nada pq meten en el calibrador un cajon no la entrada completa como en melocotones
                ' si la diferencia es positiva se suma a la ultima calidad
                If KilosNet - KilosTot > 0 Then
                    SQL = "update rclasifica_clasif set kilosnet = kilosnet + " & DBSet(KilosNet - KilosTot, "N")
                    SQL = SQL & " where numnotac = " & DBSet(Rs!numnotac, "N")
                    SQL = SQL & " and codvarie = " & DBSet(Rs!codvarie, "N")
                    SQL = SQL & " and codcalid = " & DBSet(UltCalidad, "N")

                    conn.Execute SQL
                Else
                ' si es negativa a la primera
                    SQL = "select min(codcalid) from rclasifica_clasif "
                    SQL = SQL & " where numnotac = " & DBSet(Rs!numnotac, "N")
                    SQL = SQL & " and codvarie = " & DBSet(Rs!codvarie, "N")
                    SQL = SQL & " and kilosnet >= " & DBSet((KilosNet - KilosTot) * (-1), "N")

                    PrimCalidad = DevuelveValor(SQL)

                    SQL = "update rclasifica_clasif set kilosnet = kilosnet + " & DBSet(KilosNet - KilosTot, "N")
                    SQL = SQL & " where numnotac = " & DBSet(Rs!numnotac, "N")
                    SQL = SQL & " and codvarie = " & DBSet(Rs!codvarie, "N")
                    SQL = SQL & " and codcalid = " & DBSet(PrimCalidad, "N")

                    conn.Execute SQL
                End If
            End If
        
            SQL = "delete from rclasifica_clasif where numnotac = " & DBSet(Rs!numnotac, "N")
            SQL = SQL & " and codvarie = " & DBSet(Rs!codvarie, "N") & " and kilosnet is null "
            conn.Execute SQL
            
            SQL = "delete from rclasifauto_clasif where numnotac = " & DBSet(Rs!numnotac, "N")
            SQL = SQL & " and codvarie = " & DBSet(Rs!codvarie, "N")
            conn.Execute SQL
            
            SQL = "delete from rclasifauto where numnotac = " & DBSet(Rs!numnotac, "N")
            SQL = SQL & " and codvarie = " & DBSet(Rs!codvarie, "N")
            conn.Execute SQL
            
            '++ 20-05-2009: calculamos los gastos de recoleccion para la entrada clasificadaç
            SQL = "select * from rclasifica where numnotac = " & DBSet(Rs!numnotac, "N")
            
            Set RsGastos = New ADODB.Recordset
            RsGastos.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            
            If Not RsGastos.EOF Then
                cadErr = "Actualizando Gastos"
                B = ActualizarGastos(RsGastos, cadErr)
            End If
            
            Set RsGastos = Nothing
            '++
        End If
        Rs.MoveNext
            
    Wend

    If EntClasif <> "" Then
        MsgBox "Las siguientes notas no han sido actualizadas, porque tenían clasificacion. Revise." & _
            vbCrLf & vbCrLf & Mid(EntClasif, 1, Len(EntClasif) - 2), vbExclamation
    End If

    Set Rs = Nothing

    If B Then
        ActualizarEntradasCastelduc = True
        conn.CommitTrans
        Exit Function
    End If

eActualizarEntradasCastelduc:
    If Err.Number <> 0 Or Not B Then
        conn.RollbackTrans
        MuestraError Err.Number, "Actualizar entradas", Err.Description & cadErr
    End If
End Function




Private Function EntradaClasificada(Nota As Long) As Boolean
Dim SQL As String

    EntradaClasificada = False
    
    SQL = "select sum(kilosnet) from rclasifica_clasif where numnotac = " & DBSet(Nota, "N")
    
    EntradaClasificada = (DevuelveValor(SQL) <> 0)

End Function


Private Sub CalcularTotales()
Dim Importe  As Long
Dim Rs As ADODB.Recordset
Dim SQL As String
    
    If Me.Adoaux(0).Recordset.EOF Then Exit Sub

    SQL = "select sum(kiloscal) from rclasifauto_clasif where numnotac = " & Me.Adoaux(0).Recordset!numnotac

    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    Importe = 0
    If Not Rs.EOF Then
        If Rs.Fields(0).Value <> 0 Then Importe = DBLet(Rs.Fields(0).Value, "N") 'Solo es para saber que hay registros que mostrar
    End If
    Rs.Close
    Set Rs = Nothing

    txtAux2(3).Text = Format(Importe, "##,###,##0.00")

End Sub


Private Function ActualizarEntradasCoopic() As Boolean
Dim Rs As ADODB.Recordset
Dim Rs2 As ADODB.Recordset
Dim rs3 As ADODB.Recordset
Dim RsGastos As ADODB.Recordset
Dim I As Integer
Dim SQL As String
Dim Sql2 As String
Dim Sql3 As String

Dim KilosNet As Currency
Dim FactCorrDest As Currency
Dim CalDestrio As Currency
Dim CalPodrido As Currency
Dim KilDestrio As Currency
Dim KilMuestra As Currency
Dim KilPodrido As Currency
Dim KilosTot As Currency
Dim Kilos As Currency
Dim KilosEntrada As Currency  ' kilos netos de la entrada (rclasifica)

Dim UltCalidad As Currency
Dim PrimCalidad As Currency

Dim B As Boolean
Dim cadErr As String


    On Error GoTo eActualizarEntradasCoopic

    conn.BeginTrans
    
    ActualizarEntradasCoopic = False
    
'***********************
'*********************** REPARTIMOS LOS KILOS SEGUN LO MUESTREADO
    SQL = "select  codsocio, codcampo, codvarie, fechacla from rclasifauto  "
    SQL = SQL & " group by 1,2,3,4"
    SQL = SQL & " order by 1,2,3,4"

    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
    While Not Rs.EOF
        ' obtenemos el codcampo de rclasifica pq en rclasifauto en picassent llevamos el antiguo nro de campo ej:1001
'        Sql = "select codcampo from rcampos where codsocio = " & DBSet(Rs!Codsocio, "N")
'        Sql = Sql & " and codvarie = " & DBSet(Rs!codvarie, "N")
'        Sql = Sql & " and nrocampo = " & DBSet(Rs!codcampo, "N")
'
'        campo = DevuelveValor(Sql)
    
        Sql2 = "select rclasifica.numnotac, rclasifica.codvarie, sum(rclasifica_clasif.kilosnet) "
        Sql2 = Sql2 & " from rclasifica left join rclasifica_clasif on rclasifica.numnotac = rclasifica_clasif.numnotac "
        Sql2 = Sql2 & " where rclasifica.codsocio = " & DBSet(Rs!Codsocio, "N")
        Sql2 = Sql2 & " and rclasifica.codvarie = " & DBSet(Rs!codvarie, "N")
        Sql2 = Sql2 & " and rclasifica.fechaent = " & DBSet(Rs!fechacla, "F")
        Sql2 = Sql2 & " and rclasifica.codcampo = " & DBSet(Rs!codcampo, "N")
        Sql2 = Sql2 & " group by 1,2 "
        Sql2 = Sql2 & " having sum(rclasifica_clasif.kilosnet) = 0 or sum(rclasifica_clasif.kilosnet) is null"
        Sql2 = Sql2 & " order by 1,2 "
    
        Set Rs2 = New ADODB.Recordset
        Rs2.Open Sql2, conn, adOpenDynamic, adLockPessimistic, adCmdText
    
        If Rs2.EOF Then ' entrada no existe
            Sql2 = "update rclasifauto set situacion = 2 "
            Sql2 = Sql2 & " where rclasifauto.codsocio = " & DBSet(Rs!Codsocio, "N")
            Sql2 = Sql2 & " and rclasifauto.codvarie = " & DBSet(Rs!codvarie, "N")
            Sql2 = Sql2 & " and rclasifauto.fechacla = " & DBSet(Rs!fechacla, "F")
            Sql2 = Sql2 & " and rclasifauto.codcampo = " & DBSet(Rs!codcampo, "N")
            
            conn.Execute Sql2
        Else
            While Not Rs2.EOF
                Sql3 = "select sum(kiloscal) from rclasifauto_clasif "
                Sql3 = Sql3 & " where rclasifauto_clasif.codsocio = " & DBSet(Rs!Codsocio, "N")
                Sql3 = Sql3 & " and rclasifauto_clasif.codvarie = " & DBSet(Rs!codvarie, "N")
                Sql3 = Sql3 & " and rclasifauto_clasif.fechacla = " & DBSet(Rs!fechacla, "F")
                Sql3 = Sql3 & " and rclasifauto_clasif.codcampo = " & DBSet(Rs!codcampo, "N")
                
                KilMuestra = DevuelveValor(Sql3)
            
                If KilMuestra <> 0 Then
            
                    ' kilos de la entrada
                    Sql2 = "select kilosnet from rclasifica where numnotac = " & DBSet(Rs2!numnotac, "N")
                    KilosEntrada = DevuelveValor(Sql2)
                
                    Sql3 = "select codcalid, sum(kiloscal) kiloscal from rclasifauto_clasif "
                    Sql3 = Sql3 & " where rclasifauto_clasif.codsocio = " & DBSet(Rs!Codsocio, "N")
                    Sql3 = Sql3 & " and rclasifauto_clasif.codvarie = " & DBSet(Rs!codvarie, "N")
                    Sql3 = Sql3 & " and rclasifauto_clasif.fechacla = " & DBSet(Rs!fechacla, "F")
                    Sql3 = Sql3 & " and rclasifauto_clasif.codcampo = " & DBSet(Rs!codcampo, "N")
                    Sql3 = Sql3 & " group by 1 "
                    Sql3 = Sql3 & " order by 1 "
                
                    Set rs3 = New ADODB.Recordset
                    rs3.Open Sql3, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                    
                    KilosTot = 0
                    While Not rs3.EOF
                        UltCalidad = rs3!codcalid
                            
                        Kilos = Round2(KilosEntrada * DBLet(rs3!KilosCal, "N") / KilMuestra, 0)
                        KilosTot = KilosTot + Kilos
                    
                        SQL = "select count(*) from rclasifica_clasif where numnotac = " & DBSet(Rs2!numnotac, "N")
                        SQL = SQL & " and codvarie = " & DBSet(Rs!codvarie, "N")
                        SQL = SQL & " and codcalid = " & DBSet(rs3!codcalid, "N")
                        
                        If TotalRegistros(SQL) = 0 Then
                            SQL = "insert into rclasifica_clasif (numnotac, codvarie, codcalid, muestra, kilosnet) "
                            SQL = SQL & " values (" & DBSet(Rs2!numnotac, "N") & "," & DBSet(Rs!codvarie, "N")
                            SQL = SQL & "," & DBSet(rs3!codcalid, "N") & "," & DBSet(rs3!KilosCal, "N")
                            SQL = SQL & "," & DBSet(Kilos, "N") & ")"
                            
                            conn.Execute SQL
                        Else
                            SQL = "update rclasifica_clasif set muestra = " & DBSet(rs3!KilosCal, "N") & ","
                            SQL = SQL & " kilosnet = " & DBSet(Kilos, "N")
                            SQL = SQL & " where numnotac = " & DBSet(Rs2!numnotac, "N")
                            SQL = SQL & " and codvarie = " & DBSet(Rs!codvarie, "N")
                            SQL = SQL & " and codcalid = " & DBSet(rs3!codcalid, "N")
                        
                            conn.Execute SQL
                        End If
                        
                        rs3.MoveNext
                    Wend
                
                    Set rs3 = Nothing
                    
                    ' borramos las lineas de clasificacion que no tienen calidad
                    SQL = "delete from rclasifica_clasif where numnotac = " & DBSet(Rs2!numnotac, "N")
                    SQL = SQL & " and muestra is null "
                    
                    conn.Execute SQL
                    
                    ' si la diferencia es positiva se suma a la ultima calidad
                    If KilosEntrada - KilosTot > 0 Then
                        SQL = "update rclasifica_clasif set kilosnet = kilosnet + " & DBSet(KilosEntrada - KilosTot, "N")
                        SQL = SQL & " where numnotac = " & DBSet(Rs2!numnotac, "N")
                        SQL = SQL & " and codvarie = " & DBSet(Rs!codvarie, "N")
                        SQL = SQL & " and codcalid = " & DBSet(UltCalidad, "N")
                        
                        conn.Execute SQL
                    Else
                    ' si es negativa a la primera que no deje el importe negqativo
                        SQL = "select min(codcalid) from rclasifica_clasif "
                        SQL = SQL & " where numnotac = " & DBSet(Rs2!numnotac, "N")
                        SQL = SQL & " and codvarie = " & DBSet(Rs!codvarie, "N")
                        SQL = SQL & " and kilosnet >= " & DBSet(KilosEntrada - KilosTot, "N")
                        
                        PrimCalidad = DevuelveValor(SQL)
                        
                        SQL = "update rclasifica_clasif set kilosnet = kilosnet + " & DBSet(KilosEntrada - KilosTot, "N")
                        SQL = SQL & " where numnotac = " & DBSet(Rs2!numnotac, "N")
                        SQL = SQL & " and codvarie = " & DBSet(Rs!codvarie, "N")
                        SQL = SQL & " and codcalid = " & DBSet(PrimCalidad, "N")
                        
                        conn.Execute SQL
                    End If
                
                End If
                
                Rs2.MoveNext
            Wend
        
            Set Rs2 = Nothing
        
            SQL = "delete from rclasifauto_clasif where codcampo = " & DBSet(Rs!codcampo, "N")
            SQL = SQL & " and codvarie = " & DBSet(Rs!codvarie, "N")
            SQL = SQL & " and codsocio = " & DBSet(Rs!Codsocio, "N")
            SQL = SQL & " and fechacla = " & DBSet(Rs!fechacla, "F")
            conn.Execute SQL
    
            SQL = "delete from rclasifauto where codcampo = " & DBSet(Rs!codcampo, "N")
            SQL = SQL & " and codvarie = " & DBSet(Rs!codvarie, "N")
            SQL = SQL & " and codsocio = " & DBSet(Rs!Codsocio, "N")
            SQL = SQL & " and fechacla = " & DBSet(Rs!fechacla, "F")
            conn.Execute SQL
            
        End If
        Rs.MoveNext
    Wend
        
    Set Rs = Nothing

    ActualizarEntradasCoopic = True
    conn.CommitTrans
    Exit Function

eActualizarEntradasCoopic:
    If Err.Number <> 0 Or Not B Then
        conn.RollbackTrans
        MuestraError Err.Number, "Actualizar entradas", Err.Description & cadErr
    End If
End Function
