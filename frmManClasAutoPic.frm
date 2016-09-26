VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmManClasAutoPic 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Clasificación Automática"
   ClientHeight    =   7155
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12165
   Icon            =   "frmManClasAutoPic.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7155
   ScaleWidth      =   12165
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameAux2 
      Caption         =   "Plagas"
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
      Height          =   3000
      Left            =   5580
      TabIndex        =   40
      Top             =   3510
      Width           =   6415
      Begin VB.TextBox txtAux1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   290
         Index           =   6
         Left            =   5970
         MaxLength       =   6
         TabIndex        =   56
         Tag             =   "Ordinal|N|N|||rclasifauto_plagas|ordinal|0000000|S|"
         Text            =   "ordinal"
         Top             =   2580
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtAux2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   3570
         TabIndex        =   53
         Top             =   2580
         Visible         =   0   'False
         Width           =   2340
      End
      Begin VB.TextBox txtAux1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   290
         Index           =   4
         Left            =   4530
         MaxLength       =   10
         TabIndex        =   49
         Tag             =   "Fecha|F|N|||rclasifauto_plagas|fechacla|dd/mm/yyyy|S|"
         Text            =   "fecha"
         Top             =   2550
         Visible         =   0   'False
         Width           =   390
      End
      Begin VB.TextBox txtAux1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   290
         Index           =   2
         Left            =   3570
         MaxLength       =   10
         TabIndex        =   48
         Tag             =   "Socio|N|N|||rclasifauto_plagas|codsocio|000000|S|"
         Text            =   "socio"
         Top             =   2550
         Visible         =   0   'False
         Width           =   390
      End
      Begin VB.TextBox txtAux1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   290
         Index           =   0
         Left            =   30
         MaxLength       =   16
         TabIndex        =   46
         Tag             =   "Nro.Nota|N|N|||rclasifauto_plagas|numnotac|0000000|S|"
         Text            =   "nota"
         Top             =   2565
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtAux1 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   290
         Index           =   5
         Left            =   495
         MaxLength       =   6
         TabIndex        =   45
         Tag             =   "Plaga|N|N|||rclasifauto_plagas|codplaga|00|S|"
         Text            =   "Pla"
         Top             =   2565
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.TextBox txtAux2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Index           =   5
         Left            =   1155
         TabIndex        =   44
         Top             =   2565
         Visible         =   0   'False
         Width           =   2340
      End
      Begin VB.TextBox txtAux1 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   290
         Index           =   1
         Left            =   3060
         MaxLength       =   10
         TabIndex        =   43
         Tag             =   "Variedad|N|N|||rclasifauto_plagas|codvarie|000000|S|"
         Text            =   "Va"
         Top             =   2565
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton btnBuscar 
         Appearance      =   0  'Flat
         Caption         =   "+"
         Height          =   300
         Index           =   5
         Left            =   900
         MaskColor       =   &H00000000&
         TabIndex        =   42
         ToolTipText     =   "Buscar Calidad"
         Top             =   2550
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.TextBox txtAux1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   290
         Index           =   3
         Left            =   4080
         MaxLength       =   10
         TabIndex        =   41
         Tag             =   "Campo|N|N|||rclasifauto_plagas|codcampo|00000000|S|"
         Text            =   "campo"
         Top             =   2550
         Visible         =   0   'False
         Width           =   390
      End
      Begin MSComctlLib.Toolbar ToolAux 
         Height          =   390
         Index           =   2
         Left            =   120
         TabIndex        =   47
         Top             =   255
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
      Begin MSAdodcLib.Adodc AdoAux 
         Height          =   375
         Index           =   2
         Left            =   3150
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
         Caption         =   "AdoAux(2)"
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
         Bindings        =   "frmManClasAutoPic.frx":000C
         Height          =   2310
         Index           =   2
         Left            =   120
         TabIndex        =   50
         Top             =   600
         Width           =   6100
         _ExtentX        =   10769
         _ExtentY        =   4075
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
   End
   Begin VB.TextBox txtAux2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   3
      Left            =   3630
      Locked          =   -1  'True
      MaxLength       =   7
      TabIndex        =   37
      Tag             =   "Kilos Neto|N|S|||rclasifauto_clasif|kiloscal|###,##0.00||"
      Text            =   "neto"
      Top             =   6720
      Width           =   1400
   End
   Begin VB.Frame FrameAux1 
      Height          =   2985
      Left            =   30
      TabIndex        =   9
      Top             =   510
      Width           =   11965
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   9
         Left            =   11010
         MaxLength       =   6
         TabIndex        =   57
         Tag             =   "Observac|T|S|||rclasifauto|observac|||"
         Text            =   "observ"
         Top             =   660
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   8
         Left            =   10440
         MaxLength       =   6
         TabIndex        =   54
         Tag             =   "Ordinal|N|N|0|999999|rclasifauto|ordinal|000000|S|"
         Text            =   "ordina"
         Top             =   660
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   7
         Left            =   9540
         MaxLength       =   6
         TabIndex        =   52
         Tag             =   "%Destrio|N|S|||rclasifauto|porcdest|##0.00||"
         Text            =   "%"
         Top             =   630
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   6
         Left            =   8790
         MaxLength       =   7
         TabIndex        =   51
         Tag             =   "Kilos Manuales|N|S|||rclasifauto|kilospeq|###,##0||"
         Text            =   "kilos"
         Top             =   630
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   1
         Left            =   6630
         MaxLength       =   10
         TabIndex        =   39
         Tag             =   "Fecha Clasif|F|N|||rclasifauto|fechacla|dd/mm/yyyy|S|"
         Text            =   "fecha"
         Top             =   630
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.CommandButton btnBuscar 
         Appearance      =   0  'Flat
         Caption         =   "+"
         Height          =   300
         Index           =   4
         Left            =   6450
         MaskColor       =   &H00000000&
         TabIndex        =   35
         ToolTipText     =   "Buscar Campo"
         Top             =   600
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.CommandButton btnBuscar 
         Appearance      =   0  'Flat
         Caption         =   "+"
         Height          =   300
         Index           =   3
         Left            =   3840
         MaskColor       =   &H00000000&
         TabIndex        =   34
         ToolTipText     =   "Buscar Socio"
         Top             =   600
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.CommandButton btnBuscar 
         Appearance      =   0  'Flat
         Caption         =   "+"
         Height          =   300
         Index           =   1
         Left            =   1590
         MaskColor       =   &H00000000&
         TabIndex        =   33
         ToolTipText     =   "Buscar Variedad"
         Top             =   600
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.ComboBox Combo1 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   0
         Left            =   7380
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Tag             =   "Situación|N|N|0|8|rclasifauto|situacion|||"
         Top             =   630
         Visible         =   0   'False
         Width           =   2280
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   5
         Left            =   5670
         MaxLength       =   8
         TabIndex        =   3
         Tag             =   "Nombre|N|N|||rclasifauto|codcampo|00000000|S|"
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
         Height          =   285
         Index           =   4
         Left            =   3960
         TabIndex        =   29
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
         Height          =   285
         Index           =   3
         Left            =   1620
         TabIndex        =   22
         Text            =   "12345678901234567890"
         Top             =   600
         Visible         =   0   'False
         Width           =   1620
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   3
         Left            =   960
         MaxLength       =   6
         TabIndex        =   1
         Tag             =   "Variedad|N|S|0|999999|rclasifauto|codvarie|000000|S|"
         Text            =   "123456"
         Top             =   600
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   4
         Left            =   3270
         MaxLength       =   6
         TabIndex        =   2
         Tag             =   "Nombre|N|S|||rclasifauto|codsocio|000000|S|"
         Text            =   "123456"
         Top             =   600
         Visible         =   0   'False
         Width           =   675
      End
      Begin MSDataGridLib.DataGrid DataGridAux 
         Bindings        =   "frmManClasAutoPic.frx":0024
         Height          =   2610
         Index           =   1
         Left            =   120
         TabIndex        =   36
         Top             =   210
         Width           =   11665
         _ExtentX        =   20585
         _ExtentY        =   4604
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
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   0
         Left            =   150
         MaxLength       =   7
         TabIndex        =   0
         Tag             =   "Numnotac|N|S|0|999999|rclasifauto|numnotac|000000|S|"
         Text            =   "1234567"
         Top             =   600
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.Label Label6 
         Caption         =   "Campo"
         Height          =   255
         Index           =   0
         Left            =   3990
         TabIndex        =   30
         Top             =   270
         Width           =   615
      End
      Begin VB.Label Label10 
         Caption         =   "Variedad"
         Height          =   255
         Left            =   990
         TabIndex        =   23
         Top             =   300
         Width           =   675
      End
      Begin VB.Label Label1 
         Caption         =   "Kilos Neto"
         Height          =   255
         Index           =   2
         Left            =   5580
         TabIndex        =   21
         Top             =   270
         Width           =   795
      End
      Begin VB.Label Label6 
         Caption         =   "Socio"
         Height          =   255
         Index           =   1
         Left            =   3240
         TabIndex        =   20
         Top             =   300
         Width           =   690
      End
      Begin VB.Label Label29 
         Caption         =   "Situación"
         Height          =   255
         Left            =   4710
         TabIndex        =   14
         Top             =   270
         Width           =   1005
      End
      Begin VB.Label Label1 
         Caption         =   "Nro.Nota"
         Height          =   255
         Index           =   0
         Left            =   180
         TabIndex        =   10
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
      Height          =   3000
      Left            =   30
      TabIndex        =   15
      Top             =   3510
      Width           =   5505
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   290
         Index           =   4
         Left            =   4650
         MaxLength       =   7
         TabIndex        =   55
         Tag             =   "Ordinal|N|N|||rclasifauto_clasif|ordinal|000000|S|"
         Text            =   "ordinal"
         Top             =   2550
         Visible         =   0   'False
         Width           =   1080
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   290
         Index           =   3
         Left            =   4740
         MaxLength       =   7
         TabIndex        =   28
         Tag             =   "Kilos Neto|N|S|||rclasifauto_clasif|kiloscal|###,##0||"
         Text            =   "neto"
         Top             =   2565
         Visible         =   0   'False
         Width           =   1080
      End
      Begin VB.CommandButton btnBuscar 
         Appearance      =   0  'Flat
         Caption         =   "+"
         Height          =   300
         Index           =   2
         Left            =   3420
         MaskColor       =   &H00000000&
         TabIndex        =   31
         ToolTipText     =   "Buscar Calidad"
         Top             =   2565
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.TextBox txtAux 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   290
         Index           =   2
         Left            =   3060
         MaxLength       =   2
         TabIndex        =   27
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
         Height          =   285
         Index           =   2
         Left            =   3645
         TabIndex        =   26
         Text            =   "Calidad"
         Top             =   2565
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.CommandButton btnBuscar 
         Appearance      =   0  'Flat
         Caption         =   "+"
         Height          =   300
         Index           =   0
         Left            =   945
         MaskColor       =   &H00000000&
         TabIndex        =   25
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
         Height          =   285
         Index           =   0
         Left            =   1155
         TabIndex        =   24
         Top             =   2565
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox txtAux 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   290
         Index           =   1
         Left            =   495
         MaxLength       =   6
         TabIndex        =   17
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
         Height          =   290
         Index           =   0
         Left            =   30
         MaxLength       =   16
         TabIndex        =   16
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
         TabIndex        =   18
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
         Bindings        =   "frmManClasAutoPic.frx":003C
         Height          =   2610
         Index           =   0
         Left            =   120
         TabIndex        =   19
         Top             =   300
         Width           =   5220
         _ExtentX        =   9208
         _ExtentY        =   4604
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
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Index           =   0
      Left            =   45
      TabIndex        =   7
      Top             =   6525
      Width           =   2355
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
         TabIndex        =   8
         Top             =   180
         Width           =   1935
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   10905
      TabIndex        =   6
      Top             =   6690
      Width           =   1035
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   9810
      TabIndex        =   5
      Top             =   6690
      Width           =   1035
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   12165
      _ExtentX        =   21458
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   19
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
            Enabled         =   0   'False
            Object.Visible         =   0   'False
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
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Generacion"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Actualizar Entradas"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Primero"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Anterior"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Siguiente"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Último"
            ImageIndex      =   9
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Enabled         =   0   'False
         Height          =   195
         Index           =   0
         Left            =   7830
         TabIndex        =   13
         Top             =   60
         Visible         =   0   'False
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   10920
      TabIndex        =   11
      Top             =   6660
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   2
      Left            =   3750
      MaxLength       =   250
      TabIndex        =   32
      Top             =   600
      Width           =   2205
   End
   Begin VB.Label Label2 
      Caption         =   "TOTAL :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   103
      Left            =   2730
      TabIndex        =   38
      Top             =   6750
      Width           =   1005
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
      Begin VB.Menu mnGeneracion 
         Caption         =   "&Generación"
         Shortcut        =   ^G
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
Attribute VB_Name = "frmManClasAutoPic"
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
Dim indice As Byte 'Index del text1 on es poses els datos retornats des d'atres Formularis de Mtos
Dim CadB As String

Private Sub btnBuscar_Click(Index As Integer)
    TerminaBloquear
    Select Case Index
        Case 1 ' variedades
            indice = Index + 2
            Set frmVar = New frmComVar
            frmVar.DatosADevolverBusqueda = "0|1|"
            frmVar.CodigoActual = Text1(3).Text
            frmVar.Show vbModal
            Set frmVar = Nothing
            PonerFoco Text1(3)
        
        Case 2 'calidades
            indice = Index
            Set frmCal = New frmManCalidades
            frmCal.DatosADevolverBusqueda = "2|3|"
            frmCal.CodigoActual = txtAux(2).Text
            frmCal.Show vbModal
            Set frmCal = Nothing
            PonerFoco txtAux(2)
    
        Case 3 'socios
            indice = Index + 1
            Set frmSoc = New frmManSocios
            frmSoc.DatosADevolverBusqueda = "0|1|"
'            frmSoc.CodigoActual = Text1(4).Text
            frmSoc.Show vbModal
            Set frmSoc = Nothing
            PonerFoco Text1(4)
            
        Case 4 'campos
            indice = Index + 1
            Set frmCam = New frmManCampos
            frmCam.DatosADevolverBusqueda = "0|"
'            frmCam.CodigoActual = Text1(5).Text
            frmCam.Show vbModal
            Set frmCam = Nothing
            PonerFoco Text1(5)
    
        Case 5 ' incidencias = plagas
            indice = Index + 1
            Set frmInc = New frmManInciden
            frmInc.DatosADevolverBusqueda = "0|1|2|"
'            frmCam.CodigoActual = Text1(5).Text
            frmInc.Show vbModal
            Set frmInc = Nothing
            PonerFoco txtAux1(5)
        
    End Select
    
    If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Adoaux(1), 1
End Sub


Private Sub cmdAceptar_Click()
Dim i As Long

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
            If DatosOk Then
                If InsertarDesdeForm2(Me, 1) Then
'                    text2(9).Text = PonerNombreCuenta(text1(9), Modo, text1(0).Text)
        
                    Adoaux(1).RecordSource = "Select * from " & NombreTabla & Ordenacion
                    PosicionarData
                End If
            Else
                ModoLineas = 0
            End If
            
        Case 4  'MODIFICAR
            If DatosOk Then
                If ModificaDesdeFormulario2(Me, 2, "FrameAux1") Then
                    TerminaBloquear
                    i = Adoaux(1).Recordset.Fields(0)
                    PonerModo 2
                    CargaGrid 1, True, CadB
                    Adoaux(1).Recordset.Find (Adoaux(1).Recordset.Fields(0).Name & " =" & i)
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
Dim i As Integer

    PrimeraVez = True
    
    ' ICONETS DE LA BARRA
    btnPrimero = 16 'index del botó "primero"
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
        .Buttons(11).Image = 34 'Traspaso desde el calibrador
        'el 10 i el 11 son separadors
        .Buttons(12).Image = 33 'Actualizar la clasificacion
        .Buttons(13).Image = 11  'Eixir
        'el 13 i el 14 son separadors
        .Buttons(btnPrimero).Image = 6  'Primer
        .Buttons(btnPrimero + 1).Image = 7 'Anterior
        .Buttons(btnPrimero + 2).Image = 8 'Següent
        .Buttons(btnPrimero + 3).Image = 9 'Últim
    End With
    
    ' ******* si n'hi han llínies *******
    'ICONETS DE LES BARRES ALS TABS DE LLÍNIA
    For i = 0 To ToolAux.Count
        If i <> 1 Then
            With Me.ToolAux(i)
                .HotImageList = frmPpal.imgListComun_OM16
                .DisabledImageList = frmPpal.imgListComun_BN16
                .ImageList = frmPpal.imgListComun16
                .Buttons(1).Image = 3   'Insertar
                .Buttons(2).Image = 4   'Modificar
                .Buttons(3).Image = 5   'Borrar
            End With
        End If
    Next i
    ' ***********************************
    
    CargaCombo
    
    LimpiarCampos   'Neteja els camps TextBox
    ' ******* si n'hi han llínies *******
    DataGridAux(0).ClearFields
    DataGridAux(1).ClearFields
    DataGridAux(2).ClearFields
    
    '*** canviar el nom de la taula i l'ordenació de la capçalera ***
    NombreTabla = "rclasifauto"
    Ordenacion = " ORDER BY codvarie,codsocio,codcampo"
    
'    'Mirem com està guardat el valor del check
'    chkVistaPrevia(0).Value = CheckValueLeer(Name)
'
'    AdoAux(1).ConnectionString = conn
'    '***** cambiar el nombre de la PK de la cabecera *************
'    AdoAux(1).RecordSource = "Select * from " & NombreTabla & " where numnotac=-1"
'    AdoAux(1).Refresh
       
    CargaGrid 1, False
    CargaGrid 0, False
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
Dim i As Integer, NumReg As Byte
Dim b As Boolean

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
    
    b = (Modo = 2) Or (Modo = 0) Or (Modo = 5)
    
    If b Then
        PonerContRegIndicador
    Else
        PonerIndicador lblIndicador, Modo, ModoLineas
    End If
    cmdAceptar.visible = Not b
    cmdCancelar.visible = Not b
    
    For i = 0 To Text1.Count - 1
        Text1(i).visible = Not b
    Next i
    
    Text2(3).visible = Not b
    Text2(4).visible = Not b
    btnBuscar(1).visible = (Modo = 1) 'Not b And Not Modo = 4
    btnBuscar(3).visible = (Modo = 1) 'Not b And Not Modo = 4
    btnBuscar(4).visible = (Modo = 1) 'Not b And Not Modo = 4
    btnBuscar(5).visible = Not b And Not Modo = 4
    Combo1(0).visible = Not b
    
    
    '=======================================
'    b = (Modo = 2)
'    'Posar Fleches de desplasament visibles
'    NumReg = 1
'    If Not adoaux(1).Recordset.EOF Then
'        If adoaux(1).Recordset.RecordCount > 1 Then NumReg = 2 'Només es per a saber que n'hi ha + d'1 registre
'    End If
'    DesplazamientoVisible Me.Toolbar1, btnPrimero, b, NumReg
    '---------------------------------------------
    
    b = Modo <> 0 And Modo <> 2
    cmdCancelar.visible = b
    cmdAceptar.visible = b
       
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
        CargaGrid 2, False
    End If
    
    b = (Modo = 4) Or (Modo = 2)
    DataGridAux(0).Enabled = b
    DataGridAux(1).Enabled = (Modo = 2)
    DataGridAux(2).Enabled = b
    
    
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
    Toolbar1.Buttons(7).Enabled = b And Not DeConsulta
    Me.mnNuevo.Enabled = b And Not DeConsulta
    
    b = (Modo = 2 And Me.Adoaux(1).Recordset.RecordCount > 0) And Not DeConsulta
    'Modificar
    Toolbar1.Buttons(8).Enabled = b
    Me.mnModificar.Enabled = b
    'eliminar
    Toolbar1.Buttons(9).Enabled = b
    Me.mnEliminar.Enabled = b
    'actualizar
    Toolbar1.Buttons(12).Enabled = b
    Me.mnActualizar.Enabled = b
    
    
    'Traspaso desde el calibrador
    'Toolbar1.Buttons(11).Enabled = True And Not DeConsulta
    'Imprimir
    'Toolbar1.Buttons(12).Enabled = (b Or Modo = 0)
    'Toolbar1.Buttons(12).Enabled = True And Not DeConsulta
       
    ' *** si n'hi han llínies que tenen grids (en o sense tab) ***
'++monica: si insertamos lo he quitado
'    b = (Modo = 3 Or Modo = 4 Or Modo = 2) And Not DeConsulta
    b = (Modo = 2) And Not DeConsulta
    For i = 0 To ToolAux.Count - 1
        If i <> 1 Then
            ToolAux(i).Buttons(1).Enabled = b
            If b Then bAux = (b And Me.Adoaux(i).Recordset.RecordCount > 0)
            ToolAux(i).Buttons(2).Enabled = bAux
            ToolAux(i).Buttons(3).Enabled = bAux
        End If
    Next i
    
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
Dim Sql As String
Dim Tabla As String
    
    ' ********* si n'hi han tabs, dona igual si en datagrid o no ***********
    Select Case Index
               
        Case 0 'CLASIFICACION
            Sql = "SELECT rclasifauto_clasif.numnotac, rclasifauto_clasif.codvarie, rclasifauto_clasif.codcalid,"
            Sql = Sql & " rcalidad.nomcalid, rclasifauto_clasif.kiloscal,  rclasifauto_clasif.ordinal "
            Sql = Sql & " from rclasifauto_clasif left join rcalidad on rclasifauto_clasif.codcalid = rcalidad.codcalid "
            Sql = Sql & " and rclasifauto_clasif.codvarie = rcalidad.codvarie "
            
            If enlaza Then
                Sql = Sql & ObtenerWhereCab(True)
            Else
                Sql = Sql & " WHERE rclasifauto_clasif.numnotac = -1"
            End If
'            SQL = SQL & " and rclasifauto_clasif.codcalid = rcalidad.codcalid "
'            SQL = SQL & " and rclasifauto_clasif.codvarie = rcalidad.codvarie "
            Sql = Sql & " ORDER BY rclasifauto_clasif.codvarie, rclasifauto_clasif.codcalid"
               
        Case 1 ' ENTRADAS DE CABECERA
            Sql = "select rclasifauto.numnotac, rclasifauto.codvarie, variedades.nomvarie, "
            Sql = Sql & "rclasifauto.codsocio, rsocios.nomsocio, rclasifauto.codcampo, rclasifauto.fechacla ,rclasifauto.situacion,"
            Sql = Sql & "CASE rclasifauto.situacion WHEN 0 THEN ""SIN ERROR"" WHEN 1 THEN ""NO EXISTE CALIDAD"" "
            Sql = Sql & " WHEN 2 THEN ""NO EXISTE SOCIO"" WHEN 3 THEN ""NO EXISTE VARIEDAD"" "
            Sql = Sql & " WHEN 4 THEN ""NO EXISTE NRO.CAMPO""  WHEN 5 THEN ""FECHA INCORRECTA"" END, "
            Sql = Sql & " rclasifauto.kilosnet, rclasifauto.kilospod, rclasifauto.kilosdes, rclasifauto.kilospeq, rclasifauto.porcdest, rclasifauto.ordinal, rclasifauto.observac "
            Sql = Sql & " from (rclasifauto left join variedades on rclasifauto.codvarie = variedades.codvarie) "
            Sql = Sql & " left join rsocios on rclasifauto.codsocio = rsocios.codsocio"
            
            If enlaza Then
                Sql = Sql & " WHERE 1=1 "
                If CadB <> "" Then
                    Sql = Sql & " and " & CadB
                End If
            Else
                Sql = Sql & " WHERE rclasifauto.codvarie = -1"
            End If
            Sql = Sql & " ORDER BY rclasifauto.codvarie, rclasifauto.codsocio, rclasifauto.codcampo"
    
        Case 2 'PLAGAS
            Sql = "SELECT rclasifauto_plagas.numnotac, rclasifauto_plagas.codvarie, rclasifauto_plagas.codcampo,"
            Sql = Sql & " rclasifauto_plagas.codsocio, rclasifauto_plagas.fechacla, rclasifauto_plagas.codplaga,"
            Sql = Sql & " rincidencia.nomincid, rincidencia.tipincid, CASE tipincid WHEN 0 THEN ""Leve"" WHEN 1 THEN ""Grave"" WHEN 2 THEN ""Muy grave"" END, rclasifauto_plagas.ordinal "
            Sql = Sql & " from rclasifauto_plagas left join rincidencia on rclasifauto_plagas.codplaga = rincidencia.codincid "
            
            If enlaza Then
                Sql = Sql & Replace(ObtenerWhereCab(True), "rclasifauto_clasif", "rclasifauto_plagas")
            Else
                Sql = Sql & " WHERE rclasifauto_plagas.numnotac = -1"
            End If
'            SQL = SQL & " and rclasifauto_clasif.codcalid = rcalidad.codcalid "
'            SQL = SQL & " and rclasifauto_clasif.codvarie = rcalidad.codvarie "
            Sql = Sql & " ORDER BY rclasifauto_plagas.codvarie, rclasifauto_plagas.codsocio, rclasifauto_plagas.codcampo "
    
    
    
    End Select
    
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
        '   Com la clau principal es única, en posar el sql apuntant
        '   al valor retornat sobre la clau ppal es suficient
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " " & Ordenacion
        ' **********************************
        PonerCadenaBusqueda
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub frmC_Selec(vFecha As Date)
    Text1(indice).Text = Format(vFecha, "dd/mm/yyyy") 'fecha clasificacion
End Sub

Private Sub frmCal_DatoSeleccionado(CadenaSeleccion As String)
'Calidad
    txtAux(indice).Text = RecuperaValor(CadenaSeleccion, 1) 'codcalid
    txtAux2(indice).Text = RecuperaValor(CadenaSeleccion, 2) 'descripcion
End Sub

Private Sub frmCam_DatoSeleccionado(CadenaSeleccion As String)
'Campos
    Text1(indice).Text = RecuperaValor(CadenaSeleccion, 1) 'codcampo
End Sub

Private Sub frmInc_DatoSeleccionado(CadenaSeleccion As String)
'Plagas
Dim Tipo As String

    txtAux1(5).Text = RecuperaValor(CadenaSeleccion, 1) 'codplaga
    txtAux2(5).Text = RecuperaValor(CadenaSeleccion, 2) 'descripcion
    Tipo = RecuperaValor(CadenaSeleccion, 3) ' tipo de incidencia
    Select Case Tipo
        Case 0
            txtAux2(1).Text = "Leve"
        Case 1
            txtAux2(1).Text = "Grave"
        Case 2
            txtAux2(1).Text = "Muy grave"
    End Select
End Sub

Private Sub frmSoc_DatoSeleccionado(CadenaSeleccion As String)
'Socios
    Text1(indice).Text = RecuperaValor(CadenaSeleccion, 1) 'codsocio
    Text2(indice).Text = RecuperaValor(CadenaSeleccion, 2) 'descripcion
End Sub

Private Sub frmTar_DatoSeleccionado(CadenaSeleccion As String)
'Tarifas
    Text1(indice).Text = RecuperaValor(CadenaSeleccion, 1) 'codtarifa
    Text2(indice).Text = RecuperaValor(CadenaSeleccion, 2) 'descripcion
End Sub

Private Sub frmVar_DatoSeleccionado(CadenaSeleccion As String)
'Variedades
    Text1(3).Text = RecuperaValor(CadenaSeleccion, 1) 'codvariedad
    Text2(3).Text = RecuperaValor(CadenaSeleccion, 2) 'descripcion
End Sub

Private Sub frmZ_Actualizar(vCampo As String)
     Text1(indice).Text = vCampo
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
        indice = 2
        frmZ.pTitulo = "Observaciones de la Clasificación"
        frmZ.pValor = Text1(indice).Text
        frmZ.pModo = Modo
    
        frmZ.Show vbModal
        Set frmZ = Nothing
            
        PonerFoco Text1(indice)
    End If
End Sub

Private Sub mnActualizar_Click()
    BotonActualizar
End Sub

Private Sub mnBuscar_Click()
Dim i As Integer
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

Private Sub mnGeneracion_Click()
    BotonGeneracion
End Sub

Private Sub mnModificar_Click()
    ' *** repasar el nom de l'adodc, l'index del Field i el camp que te la PK ***
    ' ### [Monica] 26/09/2006 dejamos modificar y eliminar el codigo 0
    'El registre de codi 0 no es pot Modificar ni Eliminar
    'If EsCodigoCero(CStr(adoaux(1).Recordset.Fields(0).Value), FormatoCampo(Text1(0))) Then Exit Sub
    ' ***************************************************************************
    
    If BLOQUEADesdeFormulario2(Me, Adoaux(1), 2, "FrameAux1") Then
        BotonModificar
    End If
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
        Case 11 ' Generacion
            mnGeneracion_Click
        
        Case 12 'Actualizar entradas
            mnActualizar_Click
        Case 13    'Eixir
            mnSalir_Click
            
'        Case btnPrimero To btnPrimero + 3 'Fleches Desplaçament
'            Desplazamiento (Button.Index - btnPrimero)
    End Select
End Sub

Private Sub BotonBuscar()
Dim i As Integer

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
Dim Sql As String
Dim b As Boolean

    Sql = "select count(*) from rclasifauto where situacion <> 0"
    
    If TotalRegistros(Sql) <> 0 Then
        MsgBox "Hay entradas con error. Revise.", vbExclamation
    Else
        b = False
        Sql = "select count(*) from rclasifauto where codcampo = 9999 and codsocio = 999"
        If TotalRegistros(Sql) <> 0 Then
            MsgBox "Exiten clasificaciones conjuntas. Revise.", vbExclamation
            Exit Sub
        Else
            If vParamAplic.Cooperativa = 2 Then
                Sql = "select count(*) from rclasifauto where porcdest = 0 or porcdest is null"
                If TotalRegistros(Sql) <> 0 Then
                    MsgBox "Existen clasificaciones sin porcentaje de destrio. Revise.", vbExclamation
                Else
                    b = ActualizarEntradasPicassent
                End If
            Else
                b = ActualizarEntradasPicassent
            End If
        End If
        If b Then
            MsgBox "Proceso realizado correctamente", vbExclamation
            BotonVerTodos
        End If
    End If
        
End Sub


Private Sub BotonGeneracion()
Dim Sql As String
Dim b As Boolean
Dim Cad As String

    If Adoaux(1).Recordset!Codsocio <> 999 And Adoaux(1).Recordset!codcampo <> 9999 Then
        MsgBox "La clasificacióon no es conjunta.", vbExclamation
        Exit Sub
    End If

'    CadTag = "codvarie|fechacla|codsocio|codcampo|kilosnet|observac|situacio|"
    Cad = Me.Adoaux(1).Recordset!codvarie & "|" & Me.Adoaux(1).Recordset!fechacla & "|999|9999|" & Me.Adoaux(1).Recordset!KilosNet & "|"
    Cad = Cad & Me.Adoaux(1).Recordset!Observac & "|" & Me.Adoaux(1).Recordset!Situacion & "|"

    frmListado.OpcionListado = 30
    frmListado.CadTag = Cad
    frmListado.Show vbModal

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
    Dim Cad As String
        
    'Cridem al form
    ' **************** arreglar-ho per a vore lo que es desije ****************
    ' NOTA: el total d'amples de ParaGrid, ha de sumar 100
    Cad = ""
    Cad = Cad & ParaGrid(Text1(0), 20, "Código")
    Cad = Cad & ParaGrid(Text1(1), 50, "Confección")
'    cad = cad & ParaGrid(text1(2), 60, "Descripción")
    Cad = Cad & "Variedad|nomvarie|T||30·"
    If Cad <> "" Then
        
        Screen.MousePointer = vbHourglass
        Set frmB = New frmBuscaGrid
        frmB.vCampos = Cad
        Cad = NombreTabla & " inner join variedades on forfaits.codvarie = variedades.codvarie "
        frmB.vtabla = Cad 'NombreTabla
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
Dim Cad As String
Dim Aux As String
Dim i As Integer
Dim J As Integer

    If Adoaux(1).Recordset.EOF Then
        MsgBox "Ningún registro devuelto.", vbExclamation
        Exit Sub
    End If
    
    Cad = ""
    i = 0
    Do
        J = i + 1
        i = InStr(J, DatosADevolverBusqueda, "|")
        If i > 0 Then
            Aux = Mid(DatosADevolverBusqueda, J, i - J)
            J = Val(Aux)
            Cad = Cad & Text1(J).Text & "|"
        End If
    Loop Until i = 0
    RaiseEvent DatoSeleccionado(Cad)
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
        CargaGrid 2, False
    Else
        CargaGrid 0, True
        CargaGrid 2, True
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
    Dim i As Integer
    
    Screen.MousePointer = vbHourglass
    
    If DataGridAux(1).Bookmark < DataGridAux(1).FirstRow Or DataGridAux(1).Bookmark > (DataGridAux(1).FirstRow + DataGridAux(1).VisibleRows - 1) Then
        i = DataGridAux(1).Bookmark - DataGridAux(1).FirstRow
        DataGridAux(1).Scroll 0, i
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
    Text1(1).Text = DataGridAux(1).Columns(6).Text
    Text1(6).Text = DataGridAux(1).Columns(12).Text
    Text1(7).Text = DataGridAux(1).Columns(13).Text
    Text1(8).Text = DataGridAux(1).Columns(14).Text
    Text1(9).Text = DataGridAux(1).Columns(15).Text
    Text2(3).Text = DataGridAux(1).Columns(2).Text
    Text2(4).Text = DataGridAux(1).Columns(4).Text
    
    ' ***** canviar-ho pel nom del camp del combo *********
    i = Adoaux(1).Recordset!Situacion
    ' *****************************************************
    PosicionarCombo Me.Combo1(0), i
'    For j = 0 To Combo1.ListCount - 1
'        If Combo1.ItemData(j) = i Then
'            Combo1.ListIndex = j
'            Exit For
'        End If
'    Next j

    LLamaLineas 1, 4, anc 'Pone el form en Modo=4, Modificar
   
    'Como es modificar
    PonerFoco Text1(6)
    Screen.MousePointer = vbDefault
End Sub

Private Sub BotonEliminar()
Dim Cad As String

    On Error GoTo EEliminar

    'Ciertas comprobaciones
    If Adoaux(1).Recordset.EOF Then Exit Sub

    ' *** repasar el nom de l'adodc, l'index del Field i el camp que te la PK ***
    ' ### [Monica] 26/09/2006 dejamos modificar y eliminar el codigo 0
    'El registre de codi 0 no es pot Modificar ni Eliminar
    'If EsCodigoCero(CStr(adoaux(1).Recordset.Fields(0).Value), FormatoCampo(text1(0))) Then Exit Sub
    ' ***************************************************************************

    ' *************** canviar la pregunta ****************
    Cad = "¿Seguro que desea eliminar la Clasificación?"
    Cad = Cad & vbCrLf & "Variedad: " & Adoaux(1).Recordset!codvarie
    Cad = Cad & vbCrLf & "Socio   : " & Adoaux(1).Recordset!Codsocio
    Cad = Cad & vbCrLf & "Campo   : " & Adoaux(1).Recordset!codcampo
    Cad = Cad & vbCrLf & "Fecha   : " & Adoaux(1).Recordset!fechacla
    
    
    If MsgBox(Cad, vbQuestion + vbYesNo) = vbYes Then
        On Error GoTo EEliminar
        Screen.MousePointer = vbHourglass
        NumRegElim = Adoaux(1).Recordset.AbsolutePosition
        If Not Eliminar Then
            Screen.MousePointer = vbDefault
            Exit Sub
        ElseIf SituarDataTrasEliminar(Adoaux(1), NumRegElim, True) Then
'            PonerCampos
            CargaGrid 1, True, CadB
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
Dim i As Integer
Dim CodPobla As String, desPobla As String
Dim CPostal As String, desProvi As String, desPais As String

    If Adoaux(1).Recordset.EOF Then Exit Sub
    
    PonerCamposForma2 Me, Adoaux(1), 2, "FrameAux1" 'opcio=1: posa el format o els camps de la capçalera
    
    ' *** si n'hi han llínies en datagrids ***
    'For i = 0 To DataGridAux.Count - 1
    CargaGrid i, True
    If Not Adoaux(i).Recordset.EOF Then _
        PonerCamposForma2 Me, Adoaux(i), 2, "FrameAux" & i

    
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
Dim i As Integer
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
                        LLamaLineas NumTabMto, 2 'ocultar txtAux
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
'                    PonerModo 4
                    If Not Adoaux(NumTabMto).Recordset.EOF Then
                        ' *** l'Index de Fields es el que canvie de la PK de llínies ***
                        V = Adoaux(NumTabMto).Recordset.Fields(2) 'el 2 es el nº de llinia
                        Adoaux(NumTabMto).Recordset.Find (Adoaux(NumTabMto).Recordset.Fields(2).Name & " =" & V)
                        ' ***************************************************************
                    End If
            End Select
            
'            PosicionarData
            
            ' *** si n'hi han llínies en grids i camps fora d'estos ***
            If Not Adoaux(NumTabMto).Recordset.EOF Then
                DataGridAux_RowColChange NumTabMto, 1, 1
            Else
                LimpiarCamposFrame NumTabMto
            End If
    End Select
End Sub

Private Function DatosOk() As Boolean
Dim b As Boolean
Dim Sql As String
Dim Rs As ADODB.Recordset

'Dim Datos As String

    On Error GoTo EDatosOK

    DatosOk = False
    
    b = CompForm2(Me, 1)
    If Not b Then Exit Function
    
    ' ************************************************************************************
    DatosOk = b
    
EDatosOK:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function

Private Sub PosicionarData()
Dim Cad As String, Indicador As String

    ' *** canviar-ho per tota la PK de la capçalera, no llevar els () ***
    Cad = "(numnotac=" & DBSet(Text1(0).Text, "N") & ")"
    
    ' *** gastar SituarData o SituarDataMULTI depenent de si la PK es simple o composta ***
    'If SituarDataMULTI(adoaux(1), cad, Indicador) Then
    If SituarData(Adoaux(1), Cad, Indicador) Then
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
    vWhere = " WHERE numnotac=" & Adoaux(1).Recordset!numnotac & _
             " and codvarie=" & Adoaux(1).Recordset!codvarie & _
             " and codsocio=" & Adoaux(1).Recordset!Codsocio & _
             " and codcampo=" & Adoaux(1).Recordset!codcampo
     
     
    '[Monica]06/03/2014: la fecha que me llega a'0000-00-00 00:00:00' el ado la lee como un nulo
     If IsNull(Adoaux(1).Recordset!fechacla) Then
        vWhere = vWhere & " and (fechacla is null or fechacla = '0000-00-00')"
     Else
        vWhere = vWhere & " and fechacla=" & DBSet(Adoaux(1).Recordset!fechacla, "F")
     End If
     
     vWhere = vWhere & _
             " and ordinal=" & Adoaux(1).Recordset!Ordinal

    ' ***** elimina les llínies ****
    conn.Execute "DELETE FROM rclasifauto_clasif " & vWhere
        
    conn.Execute "DELETE FROM rclasifauto_plagas " & vWhere
        
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
                
        Case 1 'Fecha
            PonerFormatoFecha Text1(Index)
        
        Case 5 'campo
            PonerFormatoEntero Text1(Index)
        
        Case 6 ' kilos peq
            PonerFormatoEntero Text1(Index)
        
        Case 7 ' porcentaje de destrio
            PonerFormatoDecimal Text1(7), 4
    
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

Private Sub KEYFecha(KeyAscii As Integer, indice As Integer)
    KeyAscii = 0
    imgFec_Click (indice)
End Sub

Private Sub KEYpress(KeyAscii As Integer)
    If KeyAscii = 13 Then 'ENTER
        KeyAscii = 0
        SendKeys "{tab}"
    ElseIf KeyAscii = 27 Then 'ESC
        If (Modo = 0 Or Modo = 2) Then Unload Me
    End If
End Sub

Private Sub KEYBusqueda(KeyAscii As Integer, indice As Integer)
    KeyAscii = 0
    imgBuscar_Click (indice)
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
Dim Sql As String
Dim vWhere As String
Dim Eliminar As Boolean

    On Error GoTo Error2

    ModoLineas = 3 'Posem Modo Eliminar Llínia
    
    If Modo = 4 Then 'Modificar Capçalera
        cmdAceptar_Click
        If ModoLineas = 0 Then Exit Sub
    End If
       
    NumTabMto = Index
'    PonerModo 5, Index

    If Adoaux(Index).Recordset.EOF Then Exit Sub
    If Not SepuedeBorrar(Index) Then Exit Sub
    NumTabMto = Index
    Eliminar = False
   
    vWhere = ObtenerWhereCab(True)
    
    ' ***** independentment de si tenen datagrid o no,
    ' canviar els noms, els formats i el DELETE *****
    Select Case Index
        Case 0 'calidades
            Sql = "¿Seguro que desea eliminar la Calidad?"
            Sql = Sql & vbCrLf & "Calidad: " & Adoaux(Index).Recordset!codcalid
            If MsgBox(Sql, vbQuestion + vbYesNo) = vbYes Then
                Eliminar = True
                Sql = "DELETE FROM rclasifauto_clasif "
                Sql = Sql & vWhere & " AND codvarie= " & Adoaux(Index).Recordset!codvarie
                Sql = Sql & " and codcalid= " & Adoaux(Index).Recordset!codcalid
                Sql = Sql & " and codsocio=" & Adoaux(Index).Recordset!Codsocio
                Sql = Sql & " and codcampo=" & Adoaux(Index).Recordset!codcampo
                
            End If
        Case 2 'plagas
            Sql = "¿Seguro que desea eliminar la Plaga?"
            Sql = Sql & vbCrLf & "Plaga: " & Adoaux(Index).Recordset!codplaga
            If MsgBox(Sql, vbQuestion + vbYesNo) = vbYes Then
                Eliminar = True
                Sql = "DELETE FROM rclasifauto_plagas "
                Sql = Sql & Replace(vWhere, "rclasifauto_clasif", "rclasifauto_plagas") '& " AND codvarie= " & DBSet(AdoAux(Index).Recordset!CodVarie, "N")
'                Sql = Sql & " and numnotac= " & DBSet(AdoAux(Index).Recordset!numnotac, "N")
'                Sql = Sql & " and codsocio=" & DBSet(AdoAux(Index).Recordset!CodSocio, "N")
'                Sql = Sql & " and codcampo=" & DBSet(AdoAux(Index).Recordset!CodCampo, "N")
'                Sql = Sql & " and fechacla=" & DBSet(AdoAux(Index).Recordset!fechacla, "F")
                Sql = Sql & " and codplaga=" & DBSet(Adoaux(Index).Recordset!codplaga, "N")
                
                
            End If
        
    End Select

    If Eliminar Then
        NumRegElim = Adoaux(Index).Recordset.AbsolutePosition
        TerminaBloquear
        conn.Execute Sql
        ' *** si n'hi han tabs sense datagrid, posar l'If ***
        CargaGrid Index, True
        If Not SituarDataTrasEliminar(Adoaux(Index), NumRegElim, True) Then
'            PonerCampos
            
        End If
        If BLOQUEADesdeFormulario2(Me, Adoaux(1), 2, "FrameAux1") Then
            CargaGrid Index, True
'            BotonModificar
        End If
        ' *** si n'hi han tabs ***
'        SituarTab (NumTabMto + 1)
    End If
    
    ModoLineas = 0
'    PosicionarData
    
    Exit Sub
Error2:
    Screen.MousePointer = vbDefault
    MuestraError Err.Number, "Eliminando linea", Err.Description
End Sub

Private Sub BotonAnyadirLinea(Index As Integer)
Dim NumF As String
Dim vWhere As String, vtabla As String
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

    ' *** posar el nom del les distintes taules de llínies ***
    Select Case Index
        Case 2: vtabla = "rclasifauto_plagas"
    End Select
    
    vWhere = Replace(ObtenerWhereCab(False), "rclasifauto_clasif", "rclasifauto_plagas")
    
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
            
            LLamaLineas Index, 5, anc
        
            Select Case Index
                ' *** valor per defecte a l'insertar i formateig de tots els camps ***
                Case 0 'calidades
                    txtAux(0).Text = Text1(0).Text 'numnotac
                    txtAux(1).Text = Text1(3).Text 'codvarie
                    txtAux(2).Text = ""
                    txtAux2(2).Text = ""
                    txtAux(3).Text = ""
                    txtAux(4).Text = Text1(8).Text
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
                    For i = 9 To 9
                        BloquearTxt txtAux(i), False
                    Next i
                    BloquearbtnBuscar Me, Modo, ModoLineas, "FrameAux1"
                    PonerFoco txtAux(9)
            End Select
            
        Case 2 ' *** pose els index dels tabs de llínies que tenen datagrid ***
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
            
            LLamaLineas Index, 5, anc
        
            txtAux1(0).Text = Adoaux(1).Recordset!numnotac 'numnotac
            txtAux1(1).Text = Adoaux(1).Recordset!codvarie 'Text1(3).Text 'codvarie
            txtAux1(2).Text = Adoaux(1).Recordset!Codsocio 'Text1(4).Text 'socio
            txtAux1(3).Text = Adoaux(1).Recordset!codcampo 'Text1(5).Text 'campo
            txtAux1(4).Text = Adoaux(1).Recordset!fechacla 'Text1(1).Text 'fecha
            txtAux1(6).Text = Adoaux(1).Recordset!Ordinal 'Text1(1).Text 'ordinal
            txtAux1(5).Text = ""
            txtAux2(5).Text = ""
            txtAux2(1).Text = ""
            
            BloquearTxt txtAux1(5), False
            BloquearbtnBuscar Me, Modo, ModoLineas, "FrameAux2"
            PonerFoco txtAux1(5)
            
            
            
            
    End Select
End Sub

Private Sub BotonModificarLinea(Index As Integer)
    Dim anc As Single
    Dim i As Integer
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
Dim b As Boolean

    ' *** si n'hi han tabs sense datagrid posar el If ***
    DeseleccionaGrid DataGridAux(Index)
    PonerModo xModo
       
    Select Case Index
        Case 0 'muestras
            b = (xModo = 1 Or xModo = 2) 'Insertar o Modificar Llínies
            For jj = 2 To 3
                txtAux(jj).visible = b
                txtAux(jj).Top = alto
            Next jj
            For jj = 2 To 2
                txtAux2(jj).visible = b
                txtAux2(jj).Top = alto
            Next jj
            btnBuscar(2).visible = b
            btnBuscar(2).Top = alto
       Case 1 ' entradas
            b = (xModo = 1 Or xModo = 4)
            Text1(0).visible = b
            Text1(0).Top = alto
            Text1(3).visible = b
            Text1(3).Top = alto
            Text1(4).visible = b
            Text1(4).Top = alto
            Text1(5).visible = b
            Text1(5).Top = alto
            Text1(6).visible = b
            Text1(6).Top = alto
            Text1(7).visible = b
            Text1(7).Top = alto
            Text1(1).visible = b
            Text1(1).Top = alto

            Text1(8).visible = False
            Text1(9).visible = False

'            Text1(6).visible = b
'            Text1(6).Top = alto
'            Text1(7).visible = b
'            Text1(7).Top = alto
'            Text1(8).visible = b
'            Text1(8).Top = alto
            Text2(3).visible = b
            Text2(3).Top = alto
            Text2(4).visible = b
            Text2(4).Top = alto
'            btnBuscar(1).visible = (Modo = 1) 'b
'            btnBuscar(1).Top = alto
'            btnBuscar(3).visible = (Modo = 1) 'b
'            btnBuscar(3).Top = alto
'            btnBuscar(4).visible = (modo = 1)b
'            btnBuscar(4).Top = alto
            Combo1(0).visible = b
            Combo1(0).Top = alto
        
        Case 2 'plagas
            b = (xModo = 5) 'Insertar o Modificar Llínies
            For jj = 5 To 5
                txtAux1(jj).visible = b
                txtAux1(jj).Top = alto
            Next jj
            For jj = 5 To 5
                txtAux2(jj).visible = b
                txtAux2(jj).Top = alto
            Next jj
            For jj = 1 To 1
                txtAux2(jj).visible = b
                txtAux2(jj).Top = alto
            Next jj
            btnBuscar(5).visible = b
            btnBuscar(5).Top = alto
            
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
                    cmdAceptar.SetFocus
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
            
            cmdAceptar.SetFocus
    End Select
End Sub

Private Sub txtAux_GotFocus(Index As Integer)
   If Not txtAux(Index).MultiLine Then ConseguirFocoLin txtAux(Index)
End Sub

Private Sub txtAux1_GotFocus(Index As Integer)
   If Not txtAux1(Index).MultiLine Then ConseguirFocoLin txtAux1(Index)
End Sub

Private Sub txtAux1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
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
        Case 5 ' codigo de plaga
            If PonerFormatoEntero(txtAux1(Index)) Then
                txtAux2(Index).Text = DevuelveValor("select nomincid from rincidencia where codincid = " & DBSet(txtAux1(Index), "N"))
                txtAux2(1).Text = DevuelveValor("select CASE tipincid WHEN 0 THEN ""Leve"" WHEN 1 THEN ""Grave"" WHEN 2 THEN ""Muy grave"" END  from rincidencia where codincid = " & txtAux1(5).Text)
                If txtAux2(Index).Text = "" Then
                    cadMen = "No existe el Código de Incidencia/Plaga: " & txtAux1(Index).Text & vbCrLf
                    cadMen = cadMen & "¿Desea crearla?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmInc = New frmManInciden
                        frmInc.DatosADevolverBusqueda = "0|1|2|"
                        frmInc.NuevoCodigo = txtAux1(Index).Text
                        txtAux1(Index).Text = ""
                        TerminaBloquear
                        frmInc.Show vbModal
                        Set frmInc = Nothing
                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Adoaux(1), 1
                    Else
                        txtAux1(Index).Text = ""
                    End If
                    PonerFoco txtAux1(Index)
                Else
                    cmdAceptar.SetFocus
                End If
            Else
                txtAux2(Index).Text = ""
            End If
            
    End Select
End Sub


Private Function DatosOkLlin(nomframe As String) As Boolean
Dim Rs As ADODB.Recordset
Dim Sql As String
Dim b As Boolean
Dim Cant As Integer
Dim Mens As String
Dim vFact As Byte, vDocum As Byte

    On Error GoTo EDatosOKLlin

    Mens = ""
    DatosOkLlin = False
        
    b = CompForm2(Me, 2, nomframe) 'Comprovar formato datos ok
    If Not b Then Exit Function
    
    ' ******************************************************************************
    DatosOkLlin = b
    
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
    indice = Index + 3
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
Dim i As Byte

    Select Case Index
        Case 1 ' entradas
            PonerContRegIndicador
            CargaGrid 0, True
            CargaGrid 2, True
            CalcularTotales
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
Dim i As Byte

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
Dim b As Boolean
Dim i As Byte
Dim tots As String

    On Error GoTo ECarga

    tots = MontaSQLCarga(Index, enlaza, CadB)

    CargaGridGnral Me.DataGridAux(Index), Me.Adoaux(Index), tots, PrimeraVez
    
    Select Case Index
        Case 0 'clasificacion
            'si es visible|control|tipo campo|nombre campo|ancho control|
            tots = "N||||0|;" 'numnotac
            tots = tots & "N|txtAux(1)|T|Variedad|800|;" '"N|btnBuscar(0)|B|||;N|txtAux2(0)|T|Nombre|2000|;"
            tots = tots & "S|txtAux(2)|T|Calidad|1000|;S|btnBuscar(2)|B|||;S|txtAux2(2)|T|Nombre|2200|;"
            tots = tots & "S|txtAux(3)|T|Muestra|1400|;N|txtAux(4)|T|Ordinal|1400|;"
            
            arregla tots, DataGridAux(Index), Me
            
'            DataGridAux(0).Columns(3).Alignment = dbgLeft
'            DataGridAux(0).Columns(5).NumberFormat = "###,##0"
'            DataGridAux(0).Columns(5).Alignment = dbgRight
        
            b = (Modo = 4) And ((ModoLineas = 1) Or (ModoLineas = 2))
            
            CalcularTotales
    
        Case 1 'entradas
            'si es visible|control|tipo campo|nombre campo|ancho control|
            tots = "N||||0|;" 'numnotac
            tots = tots & "S|Text1(3)|T|Codigo|700|;S|btnBuscar(1)|B|||;S|Text2(3)|T|Variedad|1000|;"
            tots = tots & "S|Text1(4)|T|Socio|650|;S|btnBuscar(3)|B|||;S|Text2(4)|T|Nombre|3100|;"
            tots = tots & "S|Text1(5)|T|Campo|800|;S|btnBuscar(4)|B|||;S|Text1(1)|T|Fecha|1100|;N||||0|;S|Combo1(0)|C|Situación|2100|;"
            tots = tots & "N||||0|;"
            tots = tots & "N||||0|;"
            tots = tots & "N||||0|;"
            tots = tots & "S|Text1(6)|T|Kil.Man|800|;S|Text1(7)|T|%Destrio|800|;N|Text1(8)|T|Ordinal|800|;N|Text1(9)|T|Observacion|800|;"
            
            arregla tots, DataGridAux(Index), Me
            
            DataGridAux(1).Columns(1).Alignment = dbgLeft
            DataGridAux(1).Columns(3).Alignment = dbgLeft
            DataGridAux(1).Columns(5).Alignment = dbgLeft
        
            b = (Modo = 4) And ((ModoLineas = 1) Or (ModoLineas = 2))
    
    
        Case 2 'plagas
            'si es visible|control|tipo campo|nombre campo|ancho control|
            tots = "N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;" 'numnotac
            tots = tots & "S|txtAux1(5)|T|Plaga|800|;S|btnBuscar(5)|B|||;S|txtAux2(5)|T|Nombre|3100|;"
            tots = tots & "N||||0|;S|txtAux2(1)|T|Tipo|1000|;N|txtAux1(6)|T|Ordinal|800|;"
            
            arregla tots, DataGridAux(Index), Me
            
            DataGridAux(2).Columns(6).Alignment = dbgLeft
'            DataGridAux(0).Columns(5).NumberFormat = "###,##0"
'            DataGridAux(0).Columns(5).Alignment = dbgRight
        
'            b = (Modo = 4) And ((ModoLineas = 1) Or (ModoLineas = 2))
            
    
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
Dim b As Boolean

    On Error Resume Next

    ' *** posa els noms del frames, tant si son de grid com si no ***
    Select Case NumTabMto
        Case 0: nomframe = "FrameAux0" 'envases
        Case 2: nomframe = "FrameAux2" 'plagas
    End Select
    
    If DatosOkLlin(nomframe) Then
        TerminaBloquear
        If InsertarDesdeForm2(Me, 2, nomframe) Then
            b = BLOQUEADesdeFormulario2(Me, Adoaux(1), 2, "FrameAux1")
'            b = BloqueaRegistro("rclasifauto", "codvarie = " & AdoAux(1).Recordset!CodVarie)
            Select Case NumTabMto
                Case 0, 2  ' *** els index de les llinies en grid (en o sense tab) ***
                     CargaGrid NumTabMto, True
                    If b Then BotonAnyadirLinea NumTabMto
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
        Case 2: nomframe = "FrameAux1" 'costes
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
    vWhere = vWhere & " rclasifauto_clasif.numnotac=" & Me.Adoaux(1).Recordset!numnotac & _
                      " and rclasifauto_clasif.codvarie=" & Me.Adoaux(1).Recordset!codvarie & _
                      " and rclasifauto_clasif.codsocio=" & Me.Adoaux(1).Recordset!Codsocio & _
                      " and rclasifauto_clasif.codcampo=" & Me.Adoaux(1).Recordset!codcampo & _
                      " and rclasifauto_clasif.fechacla=" & DBSet(Me.Adoaux(1).Recordset!fechacla, "F") & _
                      " and rclasifauto_clasif.ordinal=" & Me.Adoaux(1).Recordset!Ordinal
                      
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

Private Sub CargaCombo()
Dim i As Integer

    ' *** neteje els combos, els pose valor i seleccione el valor per defecte ***
    For i = 0 To Combo1.Count - 1
        Combo1(i).Clear
    Next i
    'situacion:
    ' 0 = sin error
    ' 1 = No existe calidad
    ' 2 = No existe socio
    ' 3 = No exite variedad
    ' 4 = No existe nro campo
    
    Combo1(0).AddItem "SIN ERROR"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 0
    Combo1(0).AddItem "NO EXISTE CALIDAD"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 1
    Combo1(0).AddItem "NO EXISTE SOCIO"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 2
    Combo1(0).AddItem "NO EXISTE VARIEDAD"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 3
    Combo1(0).AddItem "NO EXISTE NRO CAMPO"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 4
    Combo1(0).AddItem "FECHA INCORRECTA"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 5

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
Dim i As Integer
Dim Sql As String
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
    
    Sql = "select * from rclasifauto order by numnotac"
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
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
            
                Sql = "select count(*) from rclasifica_clasif where numnotac = " & DBSet(Rs!numnotac, "N")
                Sql = Sql & " and codvarie = " & DBSet(Rs!codvarie, "N")
                Sql = Sql & " and codcalid = " & DBSet(Rs2!codcalid, "N")
                
                If TotalRegistros(Sql) = 0 Then
                    Sql = "insert into rclasifica_clasif (numnotac, codvarie, codcalid, muestra, kilosnet) "
                    Sql = Sql & " values (" & DBSet(Rs!numnotac, "N") & "," & DBSet(Rs!codvarie, "N")
                    Sql = Sql & "," & DBSet(Rs2!codcalid, "N") & "," & DBSet(Rs!KilosCal, "N")
                    Sql = Sql & "," & DBSet(Kilos, "N") & ")"
                    
                    conn.Execute Sql
                Else
                    Sql = "update rclasifica_clasif set muestra = " & DBSet(Rs2!KilosCal, "N") & ","
                    Sql = Sql & " kilosnet = " & DBSet(Kilos, "N")
                    Sql = Sql & " where numnotac = " & DBSet(Rs!numnotac, "N")
                    Sql = Sql & " and codvarie = " & DBSet(Rs!codvarie, "N")
                    Sql = Sql & " and codcalid = " & DBSet(Rs2!codcalid, "N")
                
                    conn.Execute Sql
                End If
                
                Rs2.MoveNext
            Wend
            
            Set Rs2 = Nothing
            
            ' borramos las lineas de clasificacion que no tienen calidad
            Sql = "delete from rclasifica_clasif where numnotac = " & DBSet(Rs!numnotac, "N")
            Sql = Sql & " and muestra is null "
            
            conn.Execute Sql
            
            
            ' si la diferencia es positiva se suma a la ultima calidad
            If KilosEntrada - KilosTot > 0 Then
                Sql = "update rclasifica_clasif set kilosnet = kilosnet + " & DBSet(KilosEntrada - KilosTot, "N")
                Sql = Sql & " where numnotac = " & DBSet(Rs!numnotac, "N")
                Sql = Sql & " and codvarie = " & DBSet(Rs!codvarie, "N")
                Sql = Sql & " and codcalid = " & DBSet(UltCalidad, "N")
                
                conn.Execute Sql
            Else
            ' si es negativa a la primera que no deje el importe negqativo
                Sql = "select min(codcalid) from rclasifica_clasif "
                Sql = Sql & " where numnotac = " & DBSet(Rs!numnotac, "N")
                Sql = Sql & " and codvarie = " & DBSet(Rs!codvarie, "N")
                Sql = Sql & " and kiloscal >= " & DBSet(KilosEntrada - KilosTot, "N")
                
                PrimCalidad = DevuelveValor(Sql)
                
                Sql = "update rclasifica_clasif set kilosnet = kilosnet + " & DBSet(KilosEntrada - KilosTot, "N")
                Sql = Sql & " where numnotac = " & DBSet(Rs!numnotac, "N")
                Sql = Sql & " and codvarie = " & DBSet(Rs!codvarie, "N")
                Sql = Sql & " and codcalid = " & DBSet(PrimCalidad, "N")
                
                conn.Execute Sql
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
                
                PrimCalidad = DevuelveValor(Sql)
                
                Sql3 = "update rclasifica_clasif set kilosnet = kilosnet + " & DBSet(KilosEntrada - TotalKilos, "N")
                Sql3 = Sql3 & " where numnotac = " & DBSet(Rs!numnotac, "N")
                Sql3 = Sql3 & " and codvarie = " & DBSet(Rs!codvarie, "N")
                Sql3 = Sql3 & " and codcalid = " & DBSet(PrimCalidad, "N")
                
                conn.Execute Sql
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
        
        Sql = "delete from rclasifauto_clasif where numnotac = " & DBSet(Rs!numnotac, "N")
        Sql = Sql & " and codvarie = " & DBSet(Rs!codvarie, "N")
        conn.Execute Sql

        Sql = "delete from rclasifauto where numnotac = " & DBSet(Rs!numnotac, "N")
        Sql = Sql & " and codvarie = " & DBSet(Rs!codvarie, "N")
        conn.Execute Sql
    
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
Dim i As Integer
Dim Sql As String
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

Dim b As Boolean
Dim cadErr As String
Dim EntClasif As String

    On Error GoTo eActualizarEntradasCatadau

    conn.BeginTrans
    
    ActualizarEntradasCatadau = False
    
    Sql = "select * from rclasifauto order by numnotac"
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    b = True
    
    EntClasif = ""
    While Not Rs.EOF And b
    
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
                
                    Sql = "select count(*) from rclasifica_clasif where numnotac = " & DBSet(Rs!numnotac, "N")
                    Sql = Sql & " and codvarie = " & DBSet(Rs!codvarie, "N")
                    Sql = Sql & " and codcalid = " & DBSet(Rs2!codcalid, "N")
                    
                    If TotalRegistros(Sql) = 0 Then
                        Sql = "insert into rclasifica_clasif (numnotac, codvarie, codcalid, muestra, kilosnet) "
                        Sql = Sql & " values (" & DBSet(Rs!numnotac, "N") & "," & DBSet(Rs!codvarie, "N")
                        Sql = Sql & "," & DBSet(Rs2!codcalid, "N") & "," & DBSet(Rs2!KilosCal, "N")
                        Sql = Sql & "," & DBSet(Kilos, "N") & ")"
                        
                        conn.Execute Sql
                    Else
                        Sql = "update rclasifica_clasif set muestra = " & DBSet(Rs2!KilosCal, "N") & ","
                        Sql = Sql & " kilosnet = " & DBSet(Kilos, "N")
                        Sql = Sql & " where numnotac = " & DBSet(Rs!numnotac, "N")
                        Sql = Sql & " and codvarie = " & DBSet(Rs!codvarie, "N")
                        Sql = Sql & " and codcalid = " & DBSet(Rs2!codcalid, "N")
                    
                        conn.Execute Sql
                    End If
                    
                    Rs2.MoveNext
                Wend
                
                Set Rs2 = Nothing
                
                ' si la diferencia es positiva se suma a la ultima calidad
                If KilosNet - KilosTot > 0 Then
                    Sql = "update rclasifica_clasif set kilosnet = kilosnet + " & DBSet(KilosNet - KilosTot, "N")
                    Sql = Sql & " where numnotac = " & DBSet(Rs!numnotac, "N")
                    Sql = Sql & " and codvarie = " & DBSet(Rs!codvarie, "N")
                    Sql = Sql & " and codcalid = " & DBSet(UltCalidad, "N")
                    
                    conn.Execute Sql
                Else
                ' si es negativa a la primera
                    Sql = "select min(codcalid) from rclasifica_clasif "
                    Sql = Sql & " where numnotac = " & DBSet(Rs!numnotac, "N")
                    Sql = Sql & " and codvarie = " & DBSet(Rs!codvarie, "N")
                    Sql = Sql & " and kilosnet >= " & DBSet((KilosNet - KilosTot) * (-1), "N")
                    
                    PrimCalidad = DevuelveValor(Sql)
                    
                    Sql = "update rclasifica_clasif set kilosnet = kilosnet + " & DBSet(KilosNet - KilosTot, "N")
                    Sql = Sql & " where numnotac = " & DBSet(Rs!numnotac, "N")
                    Sql = Sql & " and codvarie = " & DBSet(Rs!codvarie, "N")
                    Sql = Sql & " and codcalid = " & DBSet(PrimCalidad, "N")
                    
                    conn.Execute Sql
                End If
            End If
        
            Sql = "delete from rclasifica_clasif where numnotac = " & DBSet(Rs!numnotac, "N")
            Sql = Sql & " and codvarie = " & DBSet(Rs!codvarie, "N") & " and kilosnet is null "
            conn.Execute Sql
            
            Sql = "delete from rclasifauto_clasif where numnotac = " & DBSet(Rs!numnotac, "N")
            Sql = Sql & " and codvarie = " & DBSet(Rs!codvarie, "N")
            conn.Execute Sql
            
            Sql = "delete from rclasifauto where numnotac = " & DBSet(Rs!numnotac, "N")
            Sql = Sql & " and codvarie = " & DBSet(Rs!codvarie, "N")
            conn.Execute Sql
            
            '++ 20-05-2009: calculamos los gastos de recoleccion para la entrada clasificadaç
            Sql = "select * from rclasifica where numnotac = " & DBSet(Rs!numnotac, "N")
            
            Set RsGastos = New ADODB.Recordset
            RsGastos.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            
            If Not RsGastos.EOF Then
                cadErr = "Actualizando Gastos"
                b = ActualizarGastos(RsGastos, cadErr)
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

    If b Then
        ActualizarEntradasCatadau = True
        conn.CommitTrans
        Exit Function
    End If

eActualizarEntradasCatadau:
    If Err.Number <> 0 Or Not b Then
        conn.RollbackTrans
        MuestraError Err.Number, "Actualizar entradas", Err.Description & cadErr
    End If
End Function



Private Function ActualizarEntradasAlzira() As Boolean
Dim Rs As ADODB.Recordset
Dim Rs2 As ADODB.Recordset
Dim RsGastos As ADODB.Recordset
Dim i As Integer
Dim Sql As String
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

Dim b As Boolean
Dim cadErr As String


    On Error GoTo eActualizarEntradasAlzira

    conn.BeginTrans
    
    ActualizarEntradasAlzira = False
    
    Sql = "select * from rclasifauto order by numnotac"
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    b = True
    
    While Not Rs.EOF And b
    
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
            
                Sql = "select count(*) from rclasifica_clasif where numnotac = " & DBSet(Rs!numnotac, "N")
                Sql = Sql & " and codvarie = " & DBSet(Rs!codvarie, "N")
                Sql = Sql & " and codcalid = " & DBSet(Rs2!codcalid, "N")
                
                If TotalRegistros(Sql) = 0 Then
                    Sql = "insert into rclasifica_clasif (numnotac, codvarie, codcalid, muestra, kilosnet) "
                    Sql = Sql & " values (" & DBSet(Rs!numnotac, "N") & "," & DBSet(Rs!codvarie, "N")
                    Sql = Sql & "," & DBSet(Rs2!codcalid, "N") & "," & DBSet(Rs2!KilosCal, "N")
                    Sql = Sql & "," & DBSet(Kilos, "N") & ")"
                    
                    conn.Execute Sql
                Else
                    Sql = "update rclasifica_clasif set muestra = " & DBSet(Rs2!KilosCal, "N") & ","
                    Sql = Sql & " kilosnet = " & DBSet(Kilos, "N")
                    Sql = Sql & " where numnotac = " & DBSet(Rs!numnotac, "N")
                    Sql = Sql & " and codvarie = " & DBSet(Rs!codvarie, "N")
                    Sql = Sql & " and codcalid = " & DBSet(Rs2!codcalid, "N")
                
                    conn.Execute Sql
                End If
                
                Rs2.MoveNext
            Wend
            
            Set Rs2 = Nothing
            
            ' si la diferencia es positiva se suma a la ultima calidad
            If KilosNet - KilosTot > 0 Then
                Sql = "update rclasifica_clasif set kilosnet = kilosnet + " & DBSet(KilosNet - KilosTot, "N")
                Sql = Sql & " where numnotac = " & DBSet(Rs!numnotac, "N")
                Sql = Sql & " and codvarie = " & DBSet(Rs!codvarie, "N")
                Sql = Sql & " and codcalid = " & DBSet(UltCalidad, "N")
                
                conn.Execute Sql
            Else
            ' si es negativa a la primera
                Sql = "select min(codcalid) from rclasifica_clasif "
                Sql = Sql & " where numnotac = " & DBSet(Rs!numnotac, "N")
                Sql = Sql & " and codvarie = " & DBSet(Rs!codvarie, "N")
                Sql = Sql & " and kilosnet >= " & DBSet((KilosNet - KilosTot) * (-1), "N")
                
                PrimCalidad = DevuelveValor(Sql)
                
                Sql = "update rclasifica_clasif set kilosnet = kilosnet + " & DBSet(KilosNet - KilosTot, "N")
                Sql = Sql & " where numnotac = " & DBSet(Rs!numnotac, "N")
                Sql = Sql & " and codvarie = " & DBSet(Rs!codvarie, "N")
                Sql = Sql & " and codcalid = " & DBSet(PrimCalidad, "N")
                
                conn.Execute Sql
            End If
        End If
    
        Sql = "delete from rclasifica_clasif where numnotac = " & DBSet(Rs!numnotac, "N")
        Sql = Sql & " and codvarie = " & DBSet(Rs!codvarie, "N") & " and kilosnet is null "
        conn.Execute Sql
        
        Sql = "delete from rclasifauto_clasif where numnotac = " & DBSet(Rs!numnotac, "N")
        Sql = Sql & " and codvarie = " & DBSet(Rs!codvarie, "N")
        conn.Execute Sql
        
        Sql = "delete from rclasifauto where numnotac = " & DBSet(Rs!numnotac, "N")
        Sql = Sql & " and codvarie = " & DBSet(Rs!codvarie, "N")
        conn.Execute Sql
        
        '++ 20-05-2009: calculamos los gastos de recoleccion para la entrada clasificadaç
        Sql = "select * from rclasifica where numnotac = " & DBSet(Rs!numnotac, "N")
        
        Set RsGastos = New ADODB.Recordset
        RsGastos.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        If Not RsGastos.EOF Then
            cadErr = "Actualizando Gastos"
            b = ActualizarGastos(RsGastos, cadErr)
        End If
        
        Set RsGastos = Nothing
        '++
    
        Rs.MoveNext
    Wend

    Set Rs = Nothing

    If b Then
        ActualizarEntradasAlzira = True
        conn.CommitTrans
        Exit Function
    End If

eActualizarEntradasAlzira:
    If Err.Number <> 0 Or Not b Then
        conn.RollbackTrans
        MuestraError Err.Number, "Actualizar entradas", Err.Description & cadErr
    End If
End Function


Private Function ActualizarEntradasCastelduc() As Boolean
Dim Rs As ADODB.Recordset
Dim Rs2 As ADODB.Recordset
Dim RsGastos As ADODB.Recordset
Dim i As Integer
Dim Sql As String
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

Dim b As Boolean
Dim cadErr As String

Dim EntClasif As String

    On Error GoTo eActualizarEntradasCastelduc

    conn.BeginTrans
    
    ActualizarEntradasCastelduc = False
    
    Sql = "select * from rclasifauto order by numnotac"
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    b = True
    EntClasif = ""
    While Not Rs.EOF And b
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
                While Not Rs2.EOF
                
                    '[Monica] 04/06/2010
                    ' comprobamos si es la calidad de destrio a la que le ponemos el total de kilos
                    Sql = "select count(*) from rcalidad where codvarie = " & DBSet(Rs2!codvarie, "N")
                    Sql = Sql & " and codcalid = " & DBSet(Rs2!codcalid, "N")
                    Sql = Sql & " and tipcalid = 1 "
                    
                    If TotalRegistros(Sql) > 0 Then
                        Kilos = DBLet(Rs2!KilosCal, "N")
                        KilosTot = KilosTot + Kilos
                    
                    Else
                        UltCalidad = Rs2!codcalid
                    
                        Kilos = Round2(KilosNet * DBLet(Rs2!KilosCal, "N") / KilMuestra, 0)
                        KilosTot = KilosTot + Kilos
                    End If
                    '[Monica] 04/06/2010
                    
                
                    Sql = "select count(*) from rclasifica_clasif where numnotac = " & DBSet(Rs!numnotac, "N")
                    Sql = Sql & " and codvarie = " & DBSet(Rs!codvarie, "N")
                    Sql = Sql & " and codcalid = " & DBSet(Rs2!codcalid, "N")
                    
                    If TotalRegistros(Sql) = 0 Then
                        Sql = "insert into rclasifica_clasif (numnotac, codvarie, codcalid, muestra, kilosnet) "
                        Sql = Sql & " values (" & DBSet(Rs!numnotac, "N") & "," & DBSet(Rs!codvarie, "N")
                        Sql = Sql & "," & DBSet(Rs2!codcalid, "N") & "," & DBSet(Rs2!KilosCal, "N")
                        Sql = Sql & "," & DBSet(Kilos, "N") & ")"
                        
                        conn.Execute Sql
                    Else
                        Sql = "update rclasifica_clasif set muestra = " & DBSet(Rs2!KilosCal, "N") & ","
                        Sql = Sql & " kilosnet = " & DBSet(Kilos, "N")
                        Sql = Sql & " where numnotac = " & DBSet(Rs!numnotac, "N")
                        Sql = Sql & " and codvarie = " & DBSet(Rs!codvarie, "N")
                        Sql = Sql & " and codcalid = " & DBSet(Rs2!codcalid, "N")
                    
                        conn.Execute Sql
                    End If
                    
                    Rs2.MoveNext
                Wend
                
                Set Rs2 = Nothing
                
                ' si la diferencia es positiva se suma a la ultima calidad
                If KilosNet - KilosTot > 0 Then
                    Sql = "update rclasifica_clasif set kilosnet = kilosnet + " & DBSet(KilosNet - KilosTot, "N")
                    Sql = Sql & " where numnotac = " & DBSet(Rs!numnotac, "N")
                    Sql = Sql & " and codvarie = " & DBSet(Rs!codvarie, "N")
                    Sql = Sql & " and codcalid = " & DBSet(UltCalidad, "N")
                    
                    conn.Execute Sql
                Else
                ' si es negativa a la primera
                    Sql = "select min(codcalid) from rclasifica_clasif "
                    Sql = Sql & " where numnotac = " & DBSet(Rs!numnotac, "N")
                    Sql = Sql & " and codvarie = " & DBSet(Rs!codvarie, "N")
                    Sql = Sql & " and kilosnet >= " & DBSet((KilosNet - KilosTot) * (-1), "N")
                    
                    PrimCalidad = DevuelveValor(Sql)
                    
                    Sql = "update rclasifica_clasif set kilosnet = kilosnet + " & DBSet(KilosNet - KilosTot, "N")
                    Sql = Sql & " where numnotac = " & DBSet(Rs!numnotac, "N")
                    Sql = Sql & " and codvarie = " & DBSet(Rs!codvarie, "N")
                    Sql = Sql & " and codcalid = " & DBSet(PrimCalidad, "N")
                    
                    conn.Execute Sql
                End If
            End If
        
            Sql = "delete from rclasifica_clasif where numnotac = " & DBSet(Rs!numnotac, "N")
            Sql = Sql & " and codvarie = " & DBSet(Rs!codvarie, "N") & " and kilosnet is null "
            conn.Execute Sql
            
            Sql = "delete from rclasifauto_clasif where numnotac = " & DBSet(Rs!numnotac, "N")
            Sql = Sql & " and codvarie = " & DBSet(Rs!codvarie, "N")
            conn.Execute Sql
            
            Sql = "delete from rclasifauto where numnotac = " & DBSet(Rs!numnotac, "N")
            Sql = Sql & " and codvarie = " & DBSet(Rs!codvarie, "N")
            conn.Execute Sql
            
            '++ 20-05-2009: calculamos los gastos de recoleccion para la entrada clasificadaç
            Sql = "select * from rclasifica where numnotac = " & DBSet(Rs!numnotac, "N")
            
            Set RsGastos = New ADODB.Recordset
            RsGastos.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            
            If Not RsGastos.EOF Then
                cadErr = "Actualizando Gastos"
                b = ActualizarGastos(RsGastos, cadErr)
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

    If b Then
        ActualizarEntradasCastelduc = True
        conn.CommitTrans
        Exit Function
    End If

eActualizarEntradasCastelduc:
    If Err.Number <> 0 Or Not b Then
        conn.RollbackTrans
        MuestraError Err.Number, "Actualizar entradas", Err.Description & cadErr
    End If
End Function




Private Function EntradaClasificada(Nota As Long) As Boolean
Dim Sql As String

    EntradaClasificada = False
    
    Sql = "select sum(kilosnet) from rclasifica_clasif where numnotac = " & DBSet(Nota, "N")
    
    EntradaClasificada = (DevuelveValor(Sql) <> 0)

End Function


Private Sub CalcularTotales()
Dim Importe  As Long
Dim Rs As ADODB.Recordset
Dim Sql As String
    
    If Me.Adoaux(1).Recordset.EOF Then Exit Sub

    Sql = "select sum(kiloscal) from rclasifauto_clasif where numnotac = " & Me.Adoaux(1).Recordset!numnotac
    Sql = Sql & " and codvarie = " & Me.Adoaux(1).Recordset!codvarie
    Sql = Sql & " and codsocio = " & Me.Adoaux(1).Recordset!Codsocio
    Sql = Sql & " and codcampo = " & Me.Adoaux(1).Recordset!codcampo
    Sql = Sql & " and fechacla = " & DBSet(Me.Adoaux(1).Recordset!fechacla, "F")

    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText

    Importe = 0
    If Not Rs.EOF Then
        If Rs.Fields(0).Value <> 0 Then Importe = DBLet(Rs.Fields(0).Value, "N") 'Solo es para saber que hay registros que mostrar
    End If
    Rs.Close
    Set Rs = Nothing

    txtAux2(3).Text = Format(Importe, "##,###,##0")

End Sub



Private Function ActualizarEntradasPicassent() As Boolean
Dim Rs As ADODB.Recordset
Dim Rs2 As ADODB.Recordset
Dim rs3 As ADODB.Recordset
Dim i As Integer
Dim Sql As String
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
Dim TKilosDes As Long

Dim CalManual As String
Dim campo As String
Dim PorcDest As Currency

    On Error GoTo eActualizarEntradasPicassent

    conn.BeginTrans
    
    ActualizarEntradasPicassent = False
    
    ' actualizamos primero la clasificacion automatica, introduciendo los kilos manuales (en tipo pequeño) y destrio
    Sql = "select * from rclasifauto order by codsocio, codcampo, codvarie, fechacla, ordinal "
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not Rs.EOF
        KilosPeq = DBLet(Rs!KilosPeq, "N")
        
        If KilosPeq <> 0 Then
            ' calidad de kilos manuales (pequeño)
            Sql2 = "select codcalid from rcalidad where codvarie = " & DBSet(Rs!codvarie, "N")
            Sql2 = Sql2 & " and tipcalid = 4 "
            CalManual = DevuelveValor(Sql2)
        
            If CalManual = 0 Then
                ' no existe la calidad de pequeño damos un error
                MsgBox "No existe calidad de kilos manuales(pequeño) para la variedad " & DBLet(Rs!codvarie, "N") & ". Revise.", vbExclamation
                conn.RollbackTrans
                Exit Function
            End If
        
            Sql2 = "select count(*) from rclasifauto_clasif "
            Sql2 = Sql2 & " where codsocio = " & DBSet(Rs!Codsocio, "N")
            Sql2 = Sql2 & " and codvarie = " & DBSet(Rs!codvarie, "N")
            Sql2 = Sql2 & " and fechacla = " & DBSet(Rs!fechacla, "F")
            Sql2 = Sql2 & " and ordinal = " & DBSet(Rs!Ordinal, "N")
            Sql2 = Sql2 & " and codcampo = " & DBSet(Rs!codcampo, "N")
            Sql2 = Sql2 & " and numnotac = " & DBSet(Rs!numnotac, "N")
            Sql2 = Sql2 & " and codcalid = " & DBSet(CalManual, "N")
            
            If TotalRegistros(Sql2) = 0 Then
                Sql2 = "insert into rclasifauto_clasif (numnotac,codvarie,codcalid,kiloscal,codcampo,codsocio,fechacla,ordinal) values "
                Sql2 = Sql2 & "(" & DBSet(Rs!numnotac, "N") & "," & DBSet(Rs!codvarie, "N") & "," & DBSet(CalManual, "N") & ","
                Sql2 = Sql2 & DBSet(KilosPeq, "N") & "," & DBSet(Rs!codcampo, "N") & "," & DBSet(Rs!Codsocio, "N") & "," & DBSet(Rs!fechacla, "F") & ","
                Sql2 = Sql2 & DBSet(Rs!Ordinal, "N") & ")"
            Else
                Sql2 = "update rclasifauto_clasif set kiloscal = kiloscal + " & DBSet(KilosPeq, "N")
                Sql2 = Sql2 & " where codsocio = " & DBSet(Rs!Codsocio, "N")
                Sql2 = Sql2 & " and codvarie = " & DBSet(Rs!codvarie, "N")
                Sql2 = Sql2 & " and fechacla = " & DBSet(Rs!fechacla, "F")
                Sql2 = Sql2 & " and ordinal = " & DBSet(Rs!Ordinal, "N")
                Sql2 = Sql2 & " and codcampo = " & DBSet(Rs!codcampo, "N")
                Sql2 = Sql2 & " and numnotac = " & DBSet(Rs!numnotac, "N")
                Sql2 = Sql2 & " and codcalid = " & DBSet(CalManual, "N")
            End If
            
            conn.Execute Sql2
        End If
        
        ' calidad de destrio
        Sql2 = "select codcalid from rcalidad where codvarie = " & DBSet(Rs!codvarie, "N")
        Sql2 = Sql2 & " and tipcalid = 1 "
        CalDestrio = DevuelveValor(Sql2)
        
        If CalDestrio = 0 Then
            ' no existe la calidad de destrio damos un error
            MsgBox "No existe calidad de destrio para la variedad " & DBLet(Rs!codvarie, "N") & ". Revise.", vbExclamation
            conn.RollbackTrans
            Exit Function
        End If
    
        PorcDest = DBLet(Rs!PorcDest, "N")
    
        Sql = "select * from rclasifauto_clasif "
        Sql = Sql & " where codsocio = " & DBSet(Rs!Codsocio, "N")
        Sql = Sql & " and codvarie = " & DBSet(Rs!codvarie, "N")
        Sql = Sql & " and fechacla = " & DBSet(Rs!fechacla, "F")
        Sql = Sql & " and ordinal = " & DBSet(Rs!Ordinal, "N")
        Sql = Sql & " and codcampo = " & DBSet(Rs!codcampo, "N")
        Sql = Sql & " and numnotac = " & DBSet(Rs!numnotac, "N")
        
        Set Rs2 = New ADODB.Recordset
        Rs2.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
        TKilosDes = 0
        While Not Rs2.EOF
            KilosDes = Round2(Rs2!KilosCal * PorcDest / 100, 0)
        
            TKilosDes = TKilosDes + KilosDes
            
            Sql = "update rclasifauto_clasif set kiloscal = kiloscal - " & DBSet(KilosDes, "N")
            Sql = Sql & " where codsocio = " & DBSet(Rs!Codsocio, "N")
            Sql = Sql & " and codvarie = " & DBSet(Rs!codvarie, "N")
            Sql = Sql & " and fechacla = " & DBSet(Rs!fechacla, "F")
            Sql = Sql & " and ordinal = " & DBSet(Rs!Ordinal, "N")
            Sql = Sql & " and codcampo = " & DBSet(Rs!codcampo, "N")
            Sql = Sql & " and numnotac = " & DBSet(Rs!numnotac, "N")
            Sql = Sql & " and codcalid = " & DBSet(Rs2!codcalid, "N")
        
            conn.Execute Sql
        
            Rs2.MoveNext
        Wend
        
        Set Rs2 = Nothing
        
        Sql = "select count(*) from rclasifauto_clasif  "
        Sql = Sql & " where codsocio = " & DBSet(Rs!Codsocio, "N")
        Sql = Sql & " and codvarie = " & DBSet(Rs!codvarie, "N")
        Sql = Sql & " and fechacla = " & DBSet(Rs!fechacla, "F")
        Sql = Sql & " and ordinal = " & DBSet(Rs!Ordinal, "N")
        Sql = Sql & " and codcampo = " & DBSet(Rs!codcampo, "N")
        Sql = Sql & " and numnotac = " & DBSet(Rs!numnotac, "N")
        Sql = Sql & " and codcalid = " & DBSet(CalDestrio, "N")
        
        If TotalRegistros(Sql) = 0 Then
            Sql2 = "insert into rclasifauto_clasif (numnotac,codvarie,codcalid,kiloscal,codcampo,codsocio,fechacla,ordinal) values "
            Sql2 = Sql2 & "(" & DBSet(Rs!numnotac, "N") & "," & DBSet(Rs!codvarie, "N") & "," & DBSet(CalDestrio, "N") & ","
            Sql2 = Sql2 & DBSet(TKilosDes, "N") & "," & DBSet(Rs!codcampo, "N") & "," & DBSet(Rs!Codsocio, "N") & "," & DBSet(Rs!fechacla, "F") & ","
            Sql2 = Sql2 & DBSet(Rs!Ordinal, "N") & ")"
        Else
            Sql2 = "update rclasifauto_clasif set kiloscal = kiloscal + " & DBSet(TKilosDes, "N")
            Sql2 = Sql2 & " where codsocio = " & DBSet(Rs!Codsocio, "N")
            Sql2 = Sql2 & " and codvarie = " & DBSet(Rs!codvarie, "N")
            Sql2 = Sql2 & " and fechacla = " & DBSet(Rs!fechacla, "F")
            Sql2 = Sql2 & " and ordinal = " & DBSet(Rs!Ordinal, "N")
            Sql2 = Sql2 & " and codcampo = " & DBSet(Rs!codcampo, "N")
            Sql2 = Sql2 & " and numnotac = " & DBSet(Rs!numnotac, "N")
            Sql2 = Sql2 & " and codcalid = " & DBSet(CalDestrio, "N")
        End If
        
        conn.Execute Sql2
        
        Rs.MoveNext
        
    Wend
    
    Set Rs = Nothing
    
        
'***********************
'*********************** REPARTIMOS LOS KILOS SEGUN LO MUESTREADO
    Sql = "select  codsocio, codcampo, codvarie, fechacla,  sum(porcdest) porcdest, sum(kilospeq) kilospeq from rclasifauto  "
    Sql = Sql & " group by 1,2,3,4"
    Sql = Sql & " order by 1,2,3,4"

    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
    While Not Rs.EOF
        ' obtenemos el codcampo de rclasifica pq en rclasifauto en picassent llevamos el antiguo nro de campo ej:1001
        Sql = "select codcampo from rcampos where codsocio = " & DBSet(Rs!Codsocio, "N")
        Sql = Sql & " and codvarie = " & DBSet(Rs!codvarie, "N")
        Sql = Sql & " and nrocampo = " & DBSet(Rs!codcampo, "N")
        
        campo = DevuelveValor(Sql)
    
        Sql2 = "select rclasifica.numnotac, rclasifica.codvarie, sum(rclasifica_clasif.kilosnet) "
        Sql2 = Sql2 & " from rclasifica left join rclasifica_clasif on rclasifica.numnotac = rclasifica_clasif.numnotac "
        Sql2 = Sql2 & " where rclasifica.codsocio = " & DBSet(Rs!Codsocio, "N")
        Sql2 = Sql2 & " and rclasifica.codvarie = " & DBSet(Rs!codvarie, "N")
        Sql2 = Sql2 & " and rclasifica.fechaent = " & DBSet(Rs!fechacla, "F")
        Sql2 = Sql2 & " and rclasifica.codcampo = " & DBSet(campo, "N")
        Sql2 = Sql2 & " group by 1,2 "
        Sql2 = Sql2 & " having sum(rclasifica_clasif.kilosnet) = 0 or sum(rclasifica_clasif.kilosnet) is null"
        Sql2 = Sql2 & " order by 1,2 "
    
        Set Rs2 = New ADODB.Recordset
        Rs2.Open Sql2, conn, adOpenDynamic, adLockPessimistic, adCmdText
    
        If Rs2.EOF Then ' entrada no existe
            Sql2 = "update rclasifauto set situacion = 5 "
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
                    
                        Sql = "select count(*) from rclasifica_clasif where numnotac = " & DBSet(Rs2!numnotac, "N")
                        Sql = Sql & " and codvarie = " & DBSet(Rs!codvarie, "N")
                        Sql = Sql & " and codcalid = " & DBSet(rs3!codcalid, "N")
                        
                        If TotalRegistros(Sql) = 0 Then
                            Sql = "insert into rclasifica_clasif (numnotac, codvarie, codcalid, muestra, kilosnet) "
                            Sql = Sql & " values (" & DBSet(Rs2!numnotac, "N") & "," & DBSet(Rs!codvarie, "N")
                            Sql = Sql & "," & DBSet(rs3!codcalid, "N") & "," & DBSet(rs3!KilosCal, "N")
                            Sql = Sql & "," & DBSet(Kilos, "N") & ")"
                            
                            conn.Execute Sql
                        Else
                            Sql = "update rclasifica_clasif set muestra = " & DBSet(rs3!KilosCal, "N") & ","
                            Sql = Sql & " kilosnet = " & DBSet(Kilos, "N")
                            Sql = Sql & " where numnotac = " & DBSet(Rs2!numnotac, "N")
                            Sql = Sql & " and codvarie = " & DBSet(Rs!codvarie, "N")
                            Sql = Sql & " and codcalid = " & DBSet(rs3!codcalid, "N")
                        
                            conn.Execute Sql
                        End If
                        
                        rs3.MoveNext
                    Wend
                
                    Set rs3 = Nothing
                    
                    ' borramos las lineas de clasificacion que no tienen calidad
                    Sql = "delete from rclasifica_clasif where numnotac = " & DBSet(Rs2!numnotac, "N")
                    Sql = Sql & " and muestra is null "
                    
                    conn.Execute Sql
                    
                    ' si la diferencia es positiva se suma a la ultima calidad
                    If KilosEntrada - KilosTot > 0 Then
                        Sql = "update rclasifica_clasif set kilosnet = kilosnet + " & DBSet(KilosEntrada - KilosTot, "N")
                        Sql = Sql & " where numnotac = " & DBSet(Rs2!numnotac, "N")
                        Sql = Sql & " and codvarie = " & DBSet(Rs!codvarie, "N")
                        Sql = Sql & " and codcalid = " & DBSet(UltCalidad, "N")
                        
                        conn.Execute Sql
                    Else
                    ' si es negativa a la primera que no deje el importe negqativo
                        Sql = "select min(codcalid) from rclasifica_clasif "
                        Sql = Sql & " where numnotac = " & DBSet(Rs2!numnotac, "N")
                        Sql = Sql & " and codvarie = " & DBSet(Rs!codvarie, "N")
                        Sql = Sql & " and kilosnet >= " & DBSet(KilosEntrada - KilosTot, "N")
                        
                        PrimCalidad = DevuelveValor(Sql)
                        
                        Sql = "update rclasifica_clasif set kilosnet = kilosnet + " & DBSet(KilosEntrada - KilosTot, "N")
                        Sql = Sql & " where numnotac = " & DBSet(Rs2!numnotac, "N")
                        Sql = Sql & " and codvarie = " & DBSet(Rs!codvarie, "N")
                        Sql = Sql & " and codcalid = " & DBSet(PrimCalidad, "N")
                        
                        conn.Execute Sql
                    End If
                
                End If
                
                
                ' insertamos las incidencias en la clasificacion
                Sql2 = "insert into rclasifica_incidencia (numnotac, codincid) values "
                                
                Sql3 = " select distinct " & DBSet(Rs2!numnotac, "N") & ", codplaga from rclasifauto_plagas "
                Sql3 = Sql3 & " where codvarie = " & DBSet(Rs!codvarie, "N")
                Sql3 = Sql3 & " and codsocio = " & DBSet(Rs!Codsocio, "N")
                Sql3 = Sql3 & " and codcampo = " & DBSet(Rs!codcampo, "N")
                Sql3 = Sql3 & " and fechacla = " & DBSet(Rs!fechacla, "F")
                
                Set rs3 = New ADODB.Recordset
                rs3.Open Sql3, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                Sql = ""
                While Not rs3.EOF
                    Sql3 = "select count(*) from rclasifica_incidencia "
                    Sql3 = Sql3 & " where numnotac = " & DBSet(Rs2!numnotac, "N")
                    Sql3 = Sql3 & " and codincid = " & DBSet(rs3!codplaga, "N")
                    If TotalRegistros(Sql3) = 0 Then
                        Sql = Sql & "(" & DBSet(Rs2!numnotac, "N") & "," & DBSet(rs3!codplaga, "N") & "),"
                    End If
                
                    rs3.MoveNext
                Wend
                Set rs3 = Nothing
                If Sql <> "" Then conn.Execute Sql2 & Mid(Sql, 1, Len(Sql) - 1)
                
                Rs2.MoveNext
            Wend
        
            Set Rs2 = Nothing
        
            Sql = "delete from rclasifauto_clasif where codcampo = " & DBSet(Rs!codcampo, "N")
            Sql = Sql & " and codvarie = " & DBSet(Rs!codvarie, "N")
            Sql = Sql & " and codsocio = " & DBSet(Rs!Codsocio, "N")
            Sql = Sql & " and fechacla = " & DBSet(Rs!fechacla, "F")
            conn.Execute Sql
    
            Sql = "delete from rclasifauto_plagas where codcampo = " & DBSet(Rs!codcampo, "N")
            Sql = Sql & " and codvarie = " & DBSet(Rs!codvarie, "N")
            Sql = Sql & " and codsocio = " & DBSet(Rs!Codsocio, "N")
            Sql = Sql & " and fechacla = " & DBSet(Rs!fechacla, "F")
            conn.Execute Sql
    
            Sql = "delete from rclasifauto where codcampo = " & DBSet(Rs!codcampo, "N")
            Sql = Sql & " and codvarie = " & DBSet(Rs!codvarie, "N")
            Sql = Sql & " and codsocio = " & DBSet(Rs!Codsocio, "N")
            Sql = Sql & " and fechacla = " & DBSet(Rs!fechacla, "F")
            conn.Execute Sql
            
        End If
        Rs.MoveNext
    Wend
        
    Set Rs = Nothing
    
    ActualizarEntradasPicassent = True
    conn.CommitTrans
    Exit Function

eActualizarEntradasPicassent:
    If Err.Number <> 0 Then
        conn.RollbackTrans
        MuestraError Err.Number, "Actualizar entradas", Err.Description
    End If
End Function

