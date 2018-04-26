VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmManLinFactSocios 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Variedades de Facturas Socios"
   ClientHeight    =   6435
   ClientLeft      =   45
   ClientTop       =   30
   ClientWidth     =   8835
   Icon            =   "frmManLinFactSocios.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6435
   ScaleWidth      =   8835
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   2475
      Index           =   0
      Left            =   120
      TabIndex        =   20
      Top             =   495
      Width           =   8625
      Begin VB.CheckBox Check1 
         Caption         =   "Descontada"
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
         Left            =   6615
         TabIndex        =   39
         Tag             =   "Descontada|N|N|0|1|rfactsoc_variedad|descontado|0||"
         Top             =   540
         Width           =   1545
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
         Index           =   7
         Left            =   2610
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   36
         Text            =   "Text2"
         Top             =   1380
         Width           =   5505
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
         Left            =   1590
         MaxLength       =   8
         TabIndex        =   4
         Tag             =   "Campo|N|N|0|99999999|rfactsoc_variedad|codcampo|00000000|S|"
         Text            =   "Text1"
         Top             =   1380
         Width           =   960
      End
      Begin VB.ComboBox Combo1 
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
         Left            =   150
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   480
         Width           =   3330
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
         Left            =   1590
         MaxLength       =   6
         TabIndex        =   3
         Tag             =   "Variedad|N|N|0|999999|rfactsoc_variedad|codvarie|000000|S|"
         Text            =   "Text1"
         Top             =   975
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
         Index           =   2
         Left            =   2610
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   30
         Text            =   "Text2"
         Top             =   975
         Width           =   5505
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
         Left            =   3675
         MaxLength       =   8
         TabIndex        =   23
         Tag             =   "Kilos Netos|N|S|||rfactsoc_variedad|kilosnet|###,##0||"
         Top             =   1980
         Width           =   1215
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
         Left            =   5055
         MaxLength       =   8
         TabIndex        =   22
         Tag             =   "precio Medio|N|S|||rfactsoc_variedad|preciomed|#0.0000||"
         Top             =   1980
         Width           =   1215
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
         Left            =   6570
         MaxLength       =   10
         TabIndex        =   21
         Tag             =   "Importe|N|S|||rfactsoc_variedad|imporvar|###,##0.00||"
         Top             =   1980
         Width           =   1530
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
         Left            =   3720
         MaxLength       =   7
         TabIndex        =   1
         Tag             =   "Nº Factura|N|S|||rfacsoc_variedad|numfactu|0000000|S|"
         Text            =   "Text1"
         Top             =   480
         Width           =   1100
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
         Index           =   1
         Left            =   5040
         MaxLength       =   10
         TabIndex        =   2
         Tag             =   "Fecha Factura|F|N|||rfactsoc_variedad|fecfactu|dd/mm/yyyy|S|"
         Top             =   480
         Width           =   1350
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   6
         Left            =   150
         MaxLength       =   6
         TabIndex        =   35
         Tag             =   "Tipo Movimiento|T|N|||rfactsoc_variedad|codtipom||S|"
         Text            =   "Text1"
         Top             =   510
         Width           =   840
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
         Index           =   1
         Left            =   210
         TabIndex        =   37
         Top             =   1380
         Width           =   675
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   1
         Left            =   1230
         ToolTipText     =   "Buscar Variedad"
         Top             =   1410
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   1230
         ToolTipText     =   "Buscar Variedad"
         Top             =   1005
         Width           =   240
      End
      Begin VB.Label Label1 
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
         Index           =   0
         Left            =   210
         TabIndex        =   31
         Top             =   975
         Width           =   1020
      End
      Begin VB.Label Label1 
         Caption         =   "Kilos Netos"
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
         Index           =   6
         Left            =   3675
         TabIndex        =   29
         Top             =   1740
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Precio Medio"
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
         Left            =   5055
         TabIndex        =   28
         Top             =   1740
         Width           =   1365
      End
      Begin VB.Label Label1 
         Caption         =   "Importe Variedad"
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
         Index           =   10
         Left            =   6570
         TabIndex        =   27
         Top             =   1740
         Width           =   1545
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo Factura"
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
         Index           =   9
         Left            =   150
         TabIndex        =   26
         Top             =   240
         Width           =   2475
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha Fac"
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
         Left            =   5070
         TabIndex        =   25
         Top             =   240
         Width           =   780
      End
      Begin VB.Label Label1 
         Caption         =   "Nº Factura"
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
         Left            =   3720
         TabIndex        =   24
         Top             =   240
         Width           =   1080
      End
   End
   Begin VB.Frame FrameAux0 
      Caption         =   "Calidades"
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
      Height          =   2595
      Left            =   135
      TabIndex        =   16
      Top             =   3030
      Width           =   8595
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
         Index           =   8
         Left            =   2700
         MaxLength       =   8
         TabIndex        =   38
         Tag             =   "Campo|N|N|0|99999999|rfactsoc_calidad|codcampo|00000000|S|"
         Text            =   "campo"
         Top             =   1800
         Visible         =   0   'False
         Width           =   585
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
         Height          =   290
         Index           =   7
         Left            =   6600
         MaxLength       =   10
         TabIndex        =   9
         Tag             =   "Importe|N|N|||rfactsoc_calidad|imporcal|###,##0.00||"
         Text            =   "Impor"
         Top             =   1800
         Visible         =   0   'False
         Width           =   585
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
         Height          =   290
         Index           =   6
         Left            =   6150
         MaxLength       =   7
         TabIndex        =   34
         Tag             =   "precio Medio|N|N|||rfactsoc_calidad|precio|#0.0000||"
         Text            =   "Pre.med"
         Top             =   1800
         Visible         =   0   'False
         Width           =   375
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
         Height          =   290
         Index           =   5
         Left            =   5550
         MaxLength       =   7
         TabIndex        =   8
         Tag             =   "Kilos Netos|N|N|||rfactsoc_calidad|kilosnet|###,##0||"
         Text            =   "Neto"
         Top             =   1800
         Visible         =   0   'False
         Width           =   570
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
         Index           =   3
         Left            =   2190
         MaxLength       =   6
         TabIndex        =   33
         Tag             =   "Variedad|N|N|0|999999|rfactsoc_calidad|codvarie|000000|S|"
         Text            =   "var"
         Top             =   1800
         Visible         =   0   'False
         Width           =   435
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
         Left            =   1410
         MaxLength       =   10
         TabIndex        =   32
         Tag             =   "Fecha Factura|F|N|||rfactsoc_calidad|fecfactu|dd/mm/yyyy|S|"
         Text            =   "fecfactu"
         Top             =   1800
         Visible         =   0   'False
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
         Height          =   290
         Index           =   1
         Left            =   870
         MaxLength       =   9
         TabIndex        =   11
         Tag             =   "Num.Factura|N|N|||rfactsoc_calidad|numfactu|0000000|S|"
         Text            =   "fact"
         Top             =   1800
         Visible         =   0   'False
         Width           =   510
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
         Left            =   3750
         MaskColor       =   &H00000000&
         TabIndex        =   12
         ToolTipText     =   "Buscar calidad"
         Top             =   1800
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
         Height          =   285
         Index           =   4
         Left            =   3960
         TabIndex        =   19
         Top             =   1800
         Visible         =   0   'False
         Width           =   1515
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
         Height          =   290
         Index           =   4
         Left            =   3360
         MaxLength       =   4
         TabIndex        =   7
         Tag             =   "Calidad|N|N|0|999999|rfactsoc_calidad|codcalid|00|S|"
         Text            =   "cal"
         Top             =   1800
         Visible         =   0   'False
         Width           =   360
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
         Index           =   0
         Left            =   225
         MaxLength       =   16
         TabIndex        =   10
         Tag             =   "Tipo Movimiento|T|N|||rfactsoc_calidad|codtipom||S|"
         Text            =   "tipo"
         Top             =   1800
         Visible         =   0   'False
         Width           =   585
      End
      Begin MSComctlLib.Toolbar ToolAux 
         Height          =   390
         Index           =   0
         Left            =   135
         TabIndex        =   17
         Top             =   225
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
         Index           =   0
         Left            =   3735
         Top             =   720
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
         Bindings        =   "frmManLinFactSocios.frx":000C
         Height          =   1695
         Index           =   0
         Left            =   150
         TabIndex        =   18
         Top             =   630
         Width           =   8275
         _ExtentX        =   14605
         _ExtentY        =   2990
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
      Left            =   120
      TabIndex        =   13
      Top             =   5730
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
         TabIndex        =   14
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
      Left            =   7560
      TabIndex        =   6
      Top             =   5820
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
      Left            =   6390
      TabIndex        =   5
      Top             =   5820
      Width           =   1065
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   1980
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
      Left            =   7560
      TabIndex        =   15
      Top             =   5820
      Visible         =   0   'False
      Width           =   1065
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
      Begin VB.Menu mnBarra2 
         Caption         =   "-"
      End
      Begin VB.Menu mnExpandirOperaciones 
         Caption         =   "Expandir &Operaciones"
         Shortcut        =   ^O
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
Attribute VB_Name = "frmManLinFactSocios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' +-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+
' +-+- Autor: MONICA                   -+-+
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
Public tipoMov As String
Public Factura As Long
Public fecfactu As Date
Public Variedad As Long
Public campo As Long

Public ModoExt As Byte

' *** declarar els formularis als que vaig a cridar ***
Private WithEvents frmB As frmBuscaGrid
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmC As frmCal 'calendario fecha
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmZ As frmZoom  'Zoom para campos Text
Attribute frmZ.VB_VarHelpID = -1

Private WithEvents frmCal As frmManCalidades 'calidades
Attribute frmCal.VB_VarHelpID = -1
Private WithEvents frmVar As frmComVar 'variedades
Attribute frmVar.VB_VarHelpID = -1
Private WithEvents frmCam As frmManCampos 'campos
Attribute frmCam.VB_VarHelpID = -1

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

Dim btnPrimero As Byte 'Variable que indica el nº del Botó PrimerRegistro en la Toolbar1
'Dim CadAncho() As Boolean  'array, per a quan cridem al form de llínies
Dim Indice As Byte 'Index del text1 on es poses els datos retornats des d'atres Formularis de Mtos
Dim CadB As String

Dim KilosAnt As Currency
Dim CajasAnt As Currency
Dim ForfaitAnt As String

Private Sub btnBuscar_Click(Index As Integer)
    TerminaBloquear
    '++monica
'    BloqueaRegistro "palets", "numpalet = " & Text1(0).Text
    
    Select Case Index
        Case 0 'calidades
            Set frmCal = New frmManCalidades
            frmCal.DatosADevolverBusqueda = "2|3|"
            frmCal.CodigoActual = txtAux(4).Text
'            frmInc.ParamVariedad = txtAux(4).Text
            frmCal.Show vbModal
            Set frmCal = Nothing
            PonerFoco txtAux(4)
    End Select
    If Modo = 4 Then BloqueaRegistro "rfactsoc", "codtipom = " & Combo1(0).ListIndex & " and numfactu = " & Text1(0).Text & " and fecfactu = " & DBSet(Text1(1).Text, "F")

    'BLOQUEADesdeFormulario2 Me, Data1, 1
End Sub


Private Sub cmdAceptar_Click()
Dim B As Boolean
Dim V As Integer
Dim Forfait As String
Dim cadena As String

    Screen.MousePointer = vbHourglass
    
    On Error GoTo Error1
    
    Select Case Modo
        Case 1  'BÚSQUEDA
            HacerBusqueda
            
        Case 3 'INSERTAR
            If DatosOK Then
                '------------------------------------------------------------------------------
                '  LOG de acciones
                
                cadena = Trim(Text1(6).Text) & " " & Text1(0).Text & " " & Text1(1).Text
                
                Set LOG = New cLOG
                LOG.Insertar 12, vUsu, "Inserta Lineas: " & cadena & vbCrLf
                Set LOG = Nothing
                '-----------------------------------------------------------------------------
            
            
                If InsertarDesdeForm2(Me, 1) Then
'                    text2(9).Text = PonerNombreCuenta(text1(9), Modo, text1(0).Text)
        
'                    Data1.RecordSource = "Select * from " & NombreTabla & _
'                                        " where numpalet = " & DBSet(text1(0).Text, "N") & _
'                                        " and numlinea = " & DBSet(text1(1).Text, "N") & " " & Ordenacion
'                    PosicionarData

                    TerminaBloquear
                    BloqueaRegistro "rfactsoc", "codtipom = '" & Trim(Text1(6).Text) & "' and numfactu = " & Text1(0).Text & " and fecfactu = " & DBSet(Text1(1).Text, "F")
                    
                    
                    CadenaConsulta = "Select * from " & NombreTabla & ObtenerWhereCP(True) & Ordenacion
                    PonerCadenaBusqueda
                    'Ponerse en Modo Insertar Lineas
                    BotonAnyadirLinea 0

                End If
            Else
                ModoLineas = 0
            End If
            
        Case 4  'MODIFICAR
            If DatosOK Then
                Modificar
                
                TerminaBloquear
                '++monica
                BloqueaRegistro "rfactsoc_variedad", "codtipom = '" & tipoMov & "' and numfactu = " & Factura & " and fecfactu = " & DBSet(fecfactu, "F") & " and codvarie = " & DBSet(Text1(2).Text, "N") & " and codcampo = " & DBSet(Text1(7).Text, "N")
                
                PonerModo 2
'                PosicionarData
            Else
                ModoLineas = 0
            End If
        ' *** si n'hi han llínies ***
        Case 5 'LLÍNIES
            Select Case ModoLineas
                Case 1 'afegir llínia
                    If InsertarLinea Then
                        CadenaConsulta = "Select * from " & NombreTabla & ObtenerWhereCP(True) & Ordenacion
                        PonerCadenaBusqueda
                        B = BLOQUEADesdeFormulario2(Me, Data1, 1)
                        CargaGrid 0, True
                        If B Then BotonAnyadirLinea NumTabMto
            
                        
                    End If
                Case 2 'modificar llínies
                    If ModificarLinea Then
                        ModoLineas = 0
                        
                        V = Adoaux(NumTabMto).Recordset.Fields(1) 'el 2 es el nº de llinia
                        
                        CadenaConsulta = "Select * from " & NombreTabla & ObtenerWhereCP(True) & Ordenacion
                        PonerCadenaBusqueda
                        B = BLOQUEADesdeFormulario2(Me, Data1, 1)
                        
                        CargaGrid NumTabMto, True
                        
                        PonerFocoGrid Me.DataGridAux(NumTabMto)
                        Adoaux(NumTabMto).Recordset.Find (Adoaux(NumTabMto).Recordset.Fields(1).Name & " =" & V)
                        
                        LLamaLineas NumTabMto, 0
                        
                        TerminaBloquear
                        '++monica
                        BloqueaRegistro "rfactsoc", "codtipom = '" & Text1(6).Text & "' and numfactu = " & Text1(0).Text & " and fecfactu = " & DBSet(Text1(1).Text, "F")
                        PosicionarData
                    Else
                        PonerFoco txtAux(1)
                    End If
            End Select
'--monica: la actualizacion de costes se hace en insertarlinea y modificarlinea
'            ActualizarCostes Data1.Recordset.Fields(0), Data1.Recordset.Fields(1), True

            'nuevo calculamos los totales de lineas
            CalcularTotales
        ' **************************
    End Select
    Screen.MousePointer = vbDefault
 
Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
    If PrimeraVez Then
        PrimeraVez = False
    
        PonerCampos
        ModoLineas = 0
           
        PosicionarCombo2 Combo1(0), tipoMov
        
        Modo = ModoExt
        Select Case Modo
            Case 0
                DatosADevolverBusqueda = "ZZ"
                PonerModo Modo
                CargaGrid 0, True
            Case 3
                mnNuevo_Click
            Case 4
                mnModificar_Click
        End Select
        
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim cad As String

    cad = ""
    If Combo1(0).ListIndex <> -1 And Text1(0).Text <> "" And Text1(1).Text <> "" And Text1(2).Text <> "" Then
        cad = Combo1(0).ListIndex & "|" & Text1(0).Text & "|" & Text1(1).Text & "|" & Text1(2).Text
    End If
    RaiseEvent DatoSeleccionado(cad)

'    CheckValueGuardar Me.Name, Me.chkVistaPrevia(0).Value
    Screen.MousePointer = vbDefault
    
    TerminaBloquear
    
End Sub

Private Sub Form_Load()
Dim I As Integer
    'Icono del formulario
    Me.Icon = frmPpal.Icon

    PrimeraVez = True
    
    ' ICONETS DE LA BARRA
    btnPrimero = 16 'index del botó "primero"
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
'        .Buttons(11).Image = 19   'Expandir Añadir, Borrar y Modificar
'        'el 10 i el 11 son separadors
'        .Buttons(12).Image = 10  'Imprimir
'        .Buttons(13).Image = 11  'Eixir
'        'el 13 i el 14 son separadors
'        .Buttons(btnPrimero).Image = 6  'Primer
'        .Buttons(btnPrimero + 1).Image = 7 'Anterior
'        .Buttons(btnPrimero + 2).Image = 8 'Següent
'        .Buttons(btnPrimero + 3).Image = 9 'Últim
'    End With
    
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
    For I = 0 To Me.imgBuscar.Count - 1
        Me.imgBuscar(I).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next I
    
    'carga IMAGES de mail
'    For i = 0 To Me.imgMail.Count - 1
'        Me.imgMail(i).Picture = frmPpal.imgListImages16.ListImages(2).Picture
'    Next i
    
    CargaCombo
    
    LimpiarCampos   'Neteja els camps TextBox
    ' ******* si n'hi han llínies *******
    DataGridAux(0).ClearFields
    
    '*** canviar el nom de la taula i l'ordenació de la capçalera ***
    NombreTabla = "rfactsoc_variedad"
    Ordenacion = " ORDER BY codtipom, numfactu, fecfactu"
    
    'Mirem com està guardat el valor del check
'    chkVistaPrevia(0).Value = CheckValueLeer(Name)
    
    Data1.ConnectionString = conn
    '***** cambiar el nombre de la PK de la cabecera *************
    Data1.RecordSource = "Select * from " & NombreTabla & " where codtipom='" & tipoMov & "' and numfactu = " & Factura & " and fecfactu = " & DBSet(fecfactu, "F") & " and codvarie = " & Variedad & " and codcampo = " & campo
    Data1.Refresh
    
'    If DatosADevolverBusqueda = "" Then
'        PonerModo 0
'    Else
'        PonerModo 1 'búsqueda
'        ' *** posar de groc els camps visibles de la clau primaria de la capçalera ***
'        Text1(0).BackColor = vblightblue 'codforfait
'        ' ****************************************************************************
'    End If
End Sub

Private Sub LimpiarCampos()
    On Error Resume Next
    
    limpiar Me   'Mètode general: Neteja els controls TextBox
    lblIndicador.Caption = ""
    

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
'    If DatosADevolverBusqueda <> "" Then
'        cmdRegresar.visible = (Modo = 2)
'    Else
'        cmdRegresar.visible = False
'    End If
    
    
    '=======================================
    B = (Modo = 2)
    'Posar Fleches de desplasament visibles
    NumReg = 1
    If Not Data1.Recordset.EOF Then
        If Data1.Recordset.RecordCount > 1 Then NumReg = 2 'Només es per a saber que n'hi ha + d'1 registre
    End If
'    DesplazamientoVisible Me.Toolbar1, btnPrimero, B, NumReg
    '---------------------------------------------
    
    B = Modo <> 0 And Modo <> 2
    cmdCancelar.visible = B
    cmdAceptar.visible = B
    cmdRegresar.visible = Not B

    'Bloqueja els camps Text1 si no estem modificant/Insertant Datos
    'Si estem en Insertar a més neteja els camps Text1
    BloquearText1 Me, Modo
    
    'Descontado en liquidaciones siempre va a estar bloqueado
    Check1(0).Enabled = False
    
    '*** si n'hi han combos a la capçalera ***
    '**************************
    
    ' ***** bloquejar tots els controls visibles de la clau primaria de la capçalera ***
    If Modo = 4 Then
        BloquearCombo Me, Modo
        For I = 0 To 2
            BloquearTxt Text1(I), True 'si estic en  modificar, bloqueja la clau primaria
        Next I
    End If
    ' **********************************************************************************
    imgBuscar(0).Enabled = (Modo = 3) Or (Modo = 4 And vParamAplic.Cooperativa = 12)
    
    ' kilos, precio e importe
    BloquearTxt Text1(3), True
    BloquearTxt Text1(4), True
    BloquearTxt Text1(5), True
    
    ' **** si n'hi han imagens de buscar en la capçalera *****
'    BloquearImgBuscar Me, Modo, ModoLineas
    BloquearImgZoom Me, Modo, ModoLineas
    ' ********************************************************
'    imgBuscar(0).visible = (Modo = 3)
'    imgBuscar(0).Enabled = (Modo = 3)
    
    ' el precio medio de lineas está siempre bloqueado
    BloquearTxt txtAux(6), True
    
'    chkVistaPrevia(0).Enabled = (Modo <= 2)
    
    PonerLongCampos

    If (Modo < 2) Or (Modo = 3) Then
        CargaGrid 0, False
    End If
    
    B = (Modo = 4) Or (Modo = 2)
    DataGridAux(0).Enabled = B
      
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
'    PonerOpcionesMenuGeneral Me
'    PonerOpcionesMenuGeneralNew Me
End Sub

Private Sub PonerModoOpcionesMenu(Modo)
'Actives unes Opcions de Menú i Toolbar según el modo en que estem
Dim B As Boolean, bAux As Boolean
Dim I As Byte
    
    'Barra de CAPÇALERA
    '------------------------------------------
    'b = (Modo = 2 Or Modo = 0 Or Modo = 1)
'    B = (Modo = 2 Or Modo = 0)
'    'Buscar
'    Toolbar1.Buttons(3).Enabled = B
'    Me.mnBuscar.Enabled = B
'    'Vore Tots
'    Toolbar1.Buttons(4).Enabled = B
'    Me.mnVerTodos.Enabled = B
'
'    'Insertar
'    Toolbar1.Buttons(7).Enabled = B And Not DeConsulta
'    Me.mnNuevo.Enabled = B And Not DeConsulta
'
'    B = (Modo = 2 And Data1.Recordset.RecordCount > 0) And Not DeConsulta
'    'Modificar
'    Toolbar1.Buttons(8).Enabled = B
'    Me.mnModificar.Enabled = B
'    'eliminar
'    Toolbar1.Buttons(9).Enabled = B
'    Me.mnEliminar.Enabled = B
'
'    'Expandir operaciones
'    Toolbar1.Buttons(11).Enabled = True And Not DeConsulta
'    'Imprimir
'    'Toolbar1.Buttons(12).Enabled = (b Or Modo = 0)
'    Toolbar1.Buttons(12).Enabled = True And Not DeConsulta
'
    ' *** si n'hi han llínies que tenen grids (en o sense tab) ***
'++monica: si insertamos lo he quitado
'    b = (Modo = 3 Or Modo = 4 Or Modo = 2) And Not DeConsulta
    B = (Modo = 4 Or Modo = 2) And Not DeConsulta
    For I = 0 To ToolAux.Count - 1
        ToolAux(I).Buttons(1).Enabled = B
        If B Then bAux = (B And Me.Adoaux(I).Recordset.RecordCount > 0)
        ToolAux(I).Buttons(2).Enabled = bAux
        ToolAux(I).Buttons(3).Enabled = bAux
    Next I
    
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
Dim tabla As String
    
    ' ********* si n'hi han tabs, dona igual si en datagrid o no ***********
    Select Case Index
               
        Case 0 'CALIDADES
            Sql = "SELECT rfactsoc_calidad.codtipom, rfactsoc_calidad.numfactu, rfactsoc_calidad.fecfactu, "
            Sql = Sql & " rfactsoc_calidad.codvarie, rfactsoc_calidad.codcampo, rfactsoc_calidad.codcalid, "
            Sql = Sql & " rcalidad.nomcalid, rfactsoc_calidad.kilosnet, rfactsoc_calidad.precio, rfactsoc_calidad.imporcal "
            Sql = Sql & " FROM rfactsoc_calidad, rcalidad "
            Sql = Sql & " where rfactsoc_calidad.codvarie = rcalidad.codvarie "
            Sql = Sql & " and rfactsoc_calidad.codcalid = rcalidad.codcalid and "
            If enlaza Then
                Sql = Sql & ObtenerWhereCab(False)
            Else
                Sql = Sql & " rfactsoc_calidad.numfactu = -1"
            End If
            Sql = Sql & " ORDER BY rfactsoc_calidad.codtipom, rfactsoc_calidad.numfactu "
               
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

Private Sub frmCal_DatoSeleccionado(CadenaSeleccion As String)
    txtAux(4).Text = RecuperaValor(CadenaSeleccion, 1) 'codcalidad
    txtAux2(4).Text = RecuperaValor(CadenaSeleccion, 2) 'descripcion
End Sub

Private Sub frmCam_DatoSeleccionado(CadenaSeleccion As String)
    Text1(7).Text = RecuperaValor(CadenaSeleccion, 1) 'codvariedad
    Text2(7).Text = RecuperaValor(CadenaSeleccion, 2) 'descripcion
End Sub

Private Sub frmVar_DatoSeleccionado(CadenaSeleccion As String)
    Text1(2).Text = RecuperaValor(CadenaSeleccion, 1) 'codvariedad
    Text2(2).Text = RecuperaValor(CadenaSeleccion, 2) 'descripcion
End Sub

Private Sub frmZ_Actualizar(vCampo As String)
     Text1(Indice).Text = vCampo
End Sub

Private Sub imgZoom_Click(Index As Integer)
    
    Set frmZ = New frmZoom

    If Index = 0 Then
        Indice = 15
        frmZ.pTitulo = "Observaciones de la Nota de Entrada de Albarán"
        frmZ.pValor = Text1(Indice).Text
        frmZ.pModo = Modo
    
        frmZ.Show vbModal
        Set frmZ = Nothing
            
        PonerFoco Text1(Indice)
    End If

End Sub


Private Sub imgBuscar_Click(Index As Integer)
    TerminaBloquear
    
    Select Case Index
        Case 0 'variedades
            Set frmVar = New frmComVar
            frmVar.DatosADevolverBusqueda = "0|1|"
            frmVar.CodigoActual = Text1(2).Text
            frmVar.Show vbModal
            Set frmVar = Nothing
            PonerFoco Text1(2)
            
        Case 1 'campo
            Set frmCam = New frmManCampos
            frmCam.DatosADevolverBusqueda = "0|1|"
'            frmCam.CodigoActual = Text1(7).Text
            frmCam.Show vbModal
            Set frmCam = Nothing
            PonerFoco Text1(7)
        
            
    End Select
    If Modo = 4 Then BloqueaRegistro "rfactsoc", "codtipom = '" & Text1(6).Text & "' and numfactu = " & Text1(0).Text & " and fecfactu = " & DBSet(Text1(1).Text, "F")


End Sub

Private Sub mnBuscar_Click()
    BotonBuscar
End Sub

Private Sub mnEliminar_Click()
    BotonEliminar
End Sub

Private Sub mnImprimir_Click()
    Screen.MousePointer = vbHourglass
'    frmListConfeccion.Show vbModal
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnModificar_Click()
    ' *** repasar el nom de l'adodc, l'index del Field i el camp que te la PK ***
    ' ### [Monica] 26/09/2006 dejamos modificar y eliminar el codigo 0
    'El registre de codi 0 no es pot Modificar ni Eliminar
    'If EsCodigoCero(CStr(Data1.Recordset.Fields(0).Value), FormatoCampo(Text1(0))) Then Exit Sub
    ' ***************************************************************************
'--monica
'    If BLOQUEADesdeFormulario2(Me, Data1, 1) Then
     BotonModificar
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
        Case 12 'Imprimir
            mnImprimir_Click
        Case 13    'Eixir
            mnSalir_Click
            
        Case btnPrimero To btnPrimero + 3 'Fleches Desplaçament
            Desplazamiento (Button.Index - btnPrimero)
    End Select
End Sub

Private Sub BotonBuscar()
Dim I As Integer
' ***** Si la clau primaria de la capçalera no es Text1(0), canviar-ho en <=== *****
'    If Modo <> 1 Then
'        LimpiarCampos
'        PonerModo 1
'        PonerFoco Text1(0) ' <===
'        Text1(0).BackColor = vblightblue ' <===
'        ' *** si n'hi han combos a la capçalera ***
'    Else
        HacerBusqueda
        If Data1.Recordset.EOF Then
            Text1(kCampo).Text = ""
            Text1(kCampo).BackColor = vbLightBlue
            PonerFoco Text1(kCampo)
        End If
'    End If
' ******************************************************************************
End Sub

Private Sub HacerBusqueda()

'    CadB = ObtenerBusqueda2(Me, 1)
'
'    If chkVistaPrevia(0) = 1 Then
'        MandaBusquedaPrevia CadB
'    ElseIf CadB <> "" Then
'        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " " & Ordenacion
'        PonerCadenaBusqueda
'    Else
'        ' *** foco al 1r camp visible de la capçalera que siga clau primaria ***
'        PonerFoco Text1(0)
'        ' **********************************************************************
'    End If
End Sub

Private Sub MandaBusquedaPrevia(CadB As String)
    Dim cad As String
        
    'Cridem al form
    ' **************** arreglar-ho per a vore lo que es desije ****************
    ' NOTA: el total d'amples de ParaGrid, ha de sumar 100
    cad = ""
    cad = cad & ParaGrid(Text1(0), 20, "Código")
    cad = cad & ParaGrid(Text1(1), 20, "Confección")
    cad = cad & ParaGrid(Text1(2), 60, "Descripción")
    If cad <> "" Then
        Screen.MousePointer = vbHourglass
        Set frmB = New frmBuscaGrid
        frmB.vCampos = cad
        frmB.vtabla = NombreTabla
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
'    LimpiarCampos 'Neteja els Text1
'    CadB = ""
'
'    If chkVistaPrevia(0).Value = 1 Then
'        MandaBusquedaPrevia ""
'    Else
'        CadenaConsulta = "Select * from " & NombreTabla & Ordenacion
'        PonerCadenaBusqueda
'    End If
End Sub

Private Sub BotonAnyadir()

    LimpiarCampos 'Huida els TextBox
    
    
    PonerModo 3
    
    ' ****** Valors per defecte a l'afegir, repasar si n'hi ha
    ' codEmpre i quins camps tenen la PK de la capçalera *******
'    text1(0).Text = SugerirCodigoSiguienteStr("forfaits", "codforfait")
'    FormateaCampo text1(0)

    PosicionarCombo2 Combo1(0), tipoMov
    
    Text1(6).Text = tipoMov
    Text1(0).Text = Factura
    Text1(1).Text = fecfactu
    
    Text1(0).BackColor = &H80000013
    Text1(1).BackColor = &H80000013
    Combo1(0).BackColor = &H80000013
    
    Text1(0).Locked = True
    Text1(1).Locked = True
    Combo1(0).Locked = True
    
    Text1(3).Text = "0"
    Text1(4).Text = "0"
    Text1(5).Text = "0"
    
    PonerFoco Text1(2) '*** 1r camp visible que siga PK ***
    
    ' *** si n'hi han camps de descripció a la capçalera ***
    'PosarDescripcions

End Sub

Private Sub BotonModificar()

    PonerModo 4
    
    PosicionarCombo2 Combo1(0), tipoMov
    Text1(6).Text = tipoMov
    Text1(0).Text = Factura
    Text1(1).Text = fecfactu
'    Text1(2).Text = Variedad
    
    Text1(0).BackColor = &H80000013
    Text1(1).BackColor = &H80000013
    Text1(2).BackColor = &H80000013
    Combo1(0).BackColor = &H80000013

    ' *** bloquejar els camps visibles de la clau primaria de la capçalera ***
    BloquearTxt Text1(0), True
    BloquearTxt Text1(1), True
    
    If vParamAplic.Cooperativa = 12 Then
        BloquearTxt Text1(2), False
    Else
        BloquearTxt Text1(2), True
    End If
    Combo1(0).Enabled = False
    
    ' *** foco al 1r camp visible que NO siga clau primaria ***
'    cmdAceptar.SetFocus
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
    cad = "¿Seguro que desea eliminar la Nota de Entrada?"
    cad = cad & vbCrLf & "Código: " & Data1.Recordset.Fields(0)
    cad = cad & vbCrLf & "Nombre: " & Data1.Recordset.Fields(1)
    
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
    
    Text2(2).Text = DevuelveDesdeBDNew(cAgro, "variedades", "nomvarie", "codvarie", Text1(2).Text, "N")
    Text2(7).Text = PartidaCampo(Text1(7).Text)
        
    ' *** si n'hi han llínies en datagrids ***
    CargaGrid I, True
    If Not Adoaux(I).Recordset.EOF Then _
        PonerCamposForma2 Me, Adoaux(I), 2, "FrameAux" & I

    
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
                '++monica
                BloqueaRegistro "rfactsoc", "codtipom = '" & Text1(6).Text & "' and numfactu = " & Text1(0).Text & " and fecfactu = " & DBSet(Text1(1).Text, "F")
                
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
Dim Sql As String
'Dim Datos As String

    On Error GoTo EDatosOK

    DatosOK = False
 
    B = CompForm2(Me, 1)
    If Not B Then Exit Function
    
    ' *** canviar els arguments de la funcio, el mensage i repasar si n'hi ha codEmpre ***
    If (Modo = 3) Then 'insertar
'        'comprobar si existe ya el cod. del campo clave primaria
        Sql = "select count(*) from rfactsoc_variedad where codtipom = " & DBSet(tipoMov, "T")
        Sql = Sql & " and numfactu = " & Factura
        Sql = Sql & " and fecfactu = " & DBSet(fecfactu, "F")
        Sql = Sql & " and codvarie = " & DBSet(Text1(2).Text, "N")
        Sql = Sql & " and codcampo = " & DBSet(Text1(7).Text, "N")
        
        If TotalRegistros(Sql) > 0 Then
            MsgBox "Ya existe la Variedad/Campo para esta factura. Reintroduzca.", vbExclamation
            PonerFoco Text1(2)
            B = False
        End If
    End If
    '[Monica]28/11/2013: solo si es campo a 0 no compruebo que exista
    If B And Modo = 3 And DBSet(Text1(7).Text, "N") <> 0 Then
        ' comprobamos que el campo sea de la variedad introducida
        Sql = "select count(*) from rcampos where codcampo = " & DBSet(Text1(7).Text, "N")
        Sql = Sql & " and codvarie = " & DBSet(Text1(2).Text, "N")
        
        If TotalRegistros(Sql) = 0 Then
            MsgBox "El campo introducido no es de la variedad. Revise.", vbExclamation
            PonerFoco Text1(7)
            B = False
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
    cad = "(codtipom = " & DBSet(Text1(6).Text, "T") & " and "
    cad = cad & "numfactu=" & DBSet(Text1(0).Text, "N")
    cad = cad & " and fecfactu = " & DBSet(Text1(1).Text, "F")
    cad = cad & " and codvarie = " & DBSet(Text1(2).Text, "N") & ")"
    
    ' *** gastar SituarData o SituarDataMULTI depenent de si la PK es simple o composta ***
    If SituarDataMULTI(Data1, cad, Indicador) Then
    'If SituarData(Data1, cad, Indicador) Then
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
    vWhere = " WHERE codforfait=" & DBSet(Data1.Recordset!codforfait, "T")
        
    ' ***** elimina les llínies ****
    conn.Execute "DELETE FROM forfaits_envases " & vWhere
        
    conn.Execute "DELETE FROM forfaits_costes " & vWhere
        
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
Dim Variedad As String

    If Not PerderFocoGnral(Text1(Index), Modo) Then Exit Sub
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    
    
    ' ***************** configurar els LostFocus dels camps de la capçalera *****************
    Select Case Index
        Case 2 ' variedad
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PonerNombreDeCod(Text1(Index), "variedades", "nomvarie")
                If Text2(Index).Text = "" Then
                    cadMen = "No existe la Variedad: " & Text1(Index).Text & vbCrLf
                    cadMen = cadMen & "¿Desea crearla?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmVar = New frmComVar
                        frmVar.DatosADevolverBusqueda = "0|1|"
                        Text1(Index).Text = ""
                        TerminaBloquear
                        frmVar.Show vbModal
                        Set frmVar = Nothing
                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
                    Else
                        Text1(Index).Text = ""
                    End If
                    PonerFoco Text1(Index)
                End If
            End If
        
        Case 7 ' campo
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index).Text = PartidaCampo(Text1(Index).Text)
                If Text2(Index).Text = "" Then
                    cadMen = "No existe el Campo: " & Text1(Index).Text & vbCrLf
                    cadMen = cadMen & "¿Desea crearlo?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmCam = New frmManCampos
                        frmCam.DatosADevolverBusqueda = "0|1|"
                        Text1(Index).Text = ""
                        TerminaBloquear
                        frmCam.Show vbModal
                        Set frmCam = Nothing
                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
                    Else
                        Text1(Index).Text = ""
                    End If
                    PonerFoco Text1(Index)
                End If
            End If
        
        
            cmdAceptar.SetFocus
            
    End Select
        ' ***************************************************************************
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = teclaBuscar Then
        If Modo = 1 Or Modo = 3 Or Modo = 4 Then
            Select Case Index
                Case 2: KEYBusqueda KeyAscii, 0 'VARIEDAD
'                Case 3: KEYBusqueda KeyAscii, 1 'VARIEDAD COMERCIAL
'                Case 4: KEYBusqueda KeyAscii, 2 'MARCA
'                Case 5: KEYBusqueda KeyAscii, 3 'FORFAIT
'                Case 13: KEYBusqueda KeyAscii, 4 'INCIDENCIA
            End Select
        End If
    Else
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
'    imgBuscar_Click (indice)
End Sub



'************* LLINIES: ****************************
Private Sub ToolAux_ButtonClick(Index As Integer, ByVal Button As MSComctlLib.Button)
'-- pon el bloqueo aqui

    If BLOQUEADesdeFormulario2(Me, Data1, 1) Then
    
'    'guardamos los kilos, cajas y forfaits
'    KilosAnt = DBLet(Data1.Recordset!PesoNeto, "N")
'    CajasAnt = DBLet(Data1.Recordset!NumCajas, "N")
'    ForfaitAnt = DBLet(Data1.Recordset!codforfait, "T")
    
    Select Case Button.Index
        Case 1
            BotonAnyadirLinea Index
        Case 2
'            BotonModificarLinea Index
        Case 3
            BotonEliminarLinea Index
        Case Else
    End Select
    End If
End Sub

Private Sub BotonEliminarLinea(Index As Integer)
Dim Sql As String
Dim vWhere As String
Dim Eliminar As Boolean
Dim bol As Boolean
Dim MenError As String
Dim cadena As String

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
        Case 0 'calidad
            Sql = "¿Seguro que desea eliminar la Calidad?"
            Sql = Sql & vbCrLf & "Calidad: " & Adoaux(Index).Recordset!codcalid
            If MsgBox(Sql, vbQuestion + vbYesNo) = vbYes Then
                Eliminar = True
                Sql = "DELETE FROM rfactsoc_calidad "
                Sql = Sql & vWhere & " AND codcalid= " & Adoaux(Index).Recordset!codcalid
            End If
            
    End Select

    If Eliminar Then
        NumRegElim = Adoaux(Index).Recordset.AbsolutePosition
        TerminaBloquear
        BloqueaRegistro "rfactsoc", "codtipom = '" & Trim(Text1(6).Text) & "' and numfactu = " & Text1(0).Text & " and fecfactu = " & DBSet(Text1(1).Text, "F")
        
        '------------------------------------------------------------------------------
        '  LOG de acciones
        
        cadena = Trim(Text1(6).Text) & " " & Text1(0).Text & " " & Text1(1).Text & " " & Adoaux(Index).Recordset!codcalid
        
        Set LOG = New cLOG
        LOG.Insertar 12, vUsu, "Eliminar Linea Calidad : " & cadena & vbCrLf
        Set LOG = Nothing
        '-----------------------------------------------------------------------------
        
        
        
        conn.Execute Sql
        
        
        
        
        
        CalcularTotales
'        ActualizarVariedades Text1(2).Text
    End If
    
    ModoLineas = 0
    PosicionarData
    
Error2:
    If Err.Number <> 0 Then
        Screen.MousePointer = vbDefault
        MuestraError Err.Number, "Eliminando linea" & MenError, Err.Description
    Else
        
        CadenaConsulta = "Select * from " & NombreTabla & ObtenerWhereCP(True) & Ordenacion
        PonerCadenaBusqueda
'--monica:02102008
'        ' *** si n'hi han tabs sense datagrid, posar l'If ***
'        CargaGrid Index, True
'        If Not SituarDataTrasEliminar(AdoAux(Index), NumRegElim, True) Then
''            PonerCampos
'
'        End If
'        CalcularTotales
'--monica
        If BLOQUEADesdeFormulario2(Me, Data1, 1) Then
'--monica:02102008
'            BotonModificar
'--monica
        End If
        ' *** si n'hi han tabs ***
'        SituarTab (NumTabMto + 1)
    
    End If
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
    BloquearTxt Text1(1), True
    

    ' *** posar el nom del les distintes taules de llínies ***
    Select Case Index
        Case 0: vtabla = "rfactsoc_calidad"
    End Select
    
    vWhere = ObtenerWhereCab(False)
    
    Select Case Index
        Case 0, 1 ' *** pose els index dels tabs de llínies que tenen datagrid ***
            ' *** canviar la clau primaria de les llínies,
            'pasar a "" si no volem que mos sugerixca res a l'afegir ***
            
'            If Index = 1 Then NumF = SugerirCodigoSiguienteStr(vTabla, "codcoste", vWhere)

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
                    txtAux(0).Text = tipoMov 'tipo de factura
                    txtAux(1).Text = Factura 'numero de factura
                    txtAux(2).Text = fecfactu ' fecha de factura
                    txtAux(3).Text = Text1(2).Text  ' codigo de variedad
                    txtAux(8).Text = Text1(7).Text  ' codigo de campo
                    txtAux(4).Text = ""
                    txtAux2(4).Text = ""
                    txtAux(5).Text = ""
                    txtAux(6).Text = ""
                    txtAux(7).Text = ""
                    BloquearbtnBuscar Me, Modo, ModoLineas, "FrameAux0"
                    PonerFoco txtAux(4)
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
        Case 0 ' incidencias
        
            txtAux(0).Text = DataGridAux(Index).Columns(0).Text
            txtAux(1).Text = DataGridAux(Index).Columns(1).Text
            txtAux(2).Text = DataGridAux(Index).Columns(2).Text
            txtAux2(2).Text = DataGridAux(Index).Columns(3).Text
            For I = 0 To 1
                BloquearTxt txtAux(I), True
            Next I
            BloquearbtnBuscar Me, Modo, ModoLineas, "FrameAux0"
            
    End Select
    
    LLamaLineas Index, ModoLineas, anc
   
    ' *** foco al 1r camp visible de les llinies en grids que no siga PK (en o sense tab) ***
    Select Case Index
        Case 0 'incidencias
            PonerFoco txtAux(2)
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
        Case 0 'calidades
            txtAux(4).visible = B 'codcalid
            txtAux(4).Top = alto
            txtAux2(4).visible = B
            txtAux2(4).Top = alto
            btnBuscar(0).visible = B
            btnBuscar(0).Top = alto
            txtAux(5).visible = B 'kilosnet
            txtAux(5).Top = alto
            txtAux(6).visible = B 'preciomed
            txtAux(6).Top = alto
            txtAux(7).visible = B 'imporcal
            txtAux(7).Top = alto
            
    End Select
End Sub

Private Sub txtAux_LostFocus(Index As Integer)
Dim cadMen As String
Dim Nuevo As Boolean
Dim Forfait As String
Dim Sql As String
Dim KilosUni As Currency

    If Not PerderFocoGnral(txtAux(Index), Modo) Then Exit Sub
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    
    ' ******* configurar el LostFocus dels camps de llínies (dins i fora grid) ********
    Select Case Index
        Case 4 ' codigo de calidad
            If PonerFormatoEntero(txtAux(Index)) Then
                txtAux2(Index).Text = PonerNombreDeCod(txtAux(Index), "rcalidad", "nomcalid", "codcalid", "N", , "codvarie", txtAux(3).Text, "N")
                If txtAux2(Index).Text = "" Then
                    cadMen = "No existe la Calidad: " & txtAux(Index).Text & vbCrLf
                    cadMen = cadMen & "¿Desea crearla?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmCal = New frmManCalidades
                        frmCal.DatosADevolverBusqueda = "2|3|"
                        frmCal.NuevoCodigo = txtAux(Index).Text
                        txtAux(Index).Text = ""
                        TerminaBloquear
                        frmCal.Show vbModal
                        Set frmCal = Nothing
                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
                    Else
                        txtAux(Index).Text = ""
                    End If
                    PonerFoco txtAux(Index)
                End If
            Else
                txtAux2(Index).Text = ""
            End If
        
        Case 5  ' kilos netos
            If txtAux(Index).Text <> "" Then PonerFormatoEntero txtAux(Index)
            ' calculamos el precio medio
            If txtAux(5).Text <> "" And txtAux(7).Text <> "" Then
                If Val(txtAux(5).Text) <> 0 Then
                    txtAux(6).Text = Round2(CCur(ImporteSinFormato(txtAux(7).Text)) / CCur(ImporteSinFormato(txtAux(5).Text)), 4)
                End If
            End If
            
        Case 6  ' precio medio
            If txtAux(Index).Text <> "" Then
                PonerFormatoDecimal txtAux(Index), 9
            End If
            
        Case 7  ' importe calidad
            If txtAux(Index).Text <> "" Then
                If PonerFormatoDecimal(txtAux(Index), 3) Then
                    ' calculamos el precio medio
                    If txtAux(5).Text <> "" And txtAux(7).Text <> "" Then
                        If Val(txtAux(5).Text) <> 0 Then
                            txtAux(6).Text = Round2(CCur(ImporteSinFormato(txtAux(7).Text)) / CCur(ImporteSinFormato(txtAux(5).Text)), 4)
                        Else
                            txtAux(6).Text = "0"
                        
                        End If
                    Else
                            txtAux(6).Text = "0"
                    End If
                    cmdAceptar.SetFocus
                End If
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
                    Case 1: 'articulo
                        KeyAscii = 0
                        btnBuscar_Click (0)
                    Case 9: 'coste
                        KeyAscii = 0
                        btnBuscar_Click (1)
                End Select
            End If
        Else
            KEYpress KeyAscii
        End If
    End If
End Sub

Private Function DatosOkLlin(nomframe As String) As Boolean
Dim Rs As ADODB.Recordset
Dim Sql As String
Dim B As Boolean
Dim Cant As Integer
Dim Mens As String
Dim vFact As Byte, vDocum As Byte

    On Error GoTo EDatosOKLlin

    Mens = ""
    DatosOkLlin = False
        
    B = CompForm2(Me, 2, nomframe) 'Comprovar formato datos ok
    If Not B Then Exit Function
    
    ' comprobamos que no exista ya la calidad para la variedad
    Sql = "select count(*) from rfactsoc_calidad where codtipom = " & DBSet(tipoMov, "T")
    Sql = Sql & " and numfactu = " & Factura & " and fecfactu = " & DBSet(fecfactu, "F")
    Sql = Sql & " and codvarie = " & Text1(2).Text
    Sql = Sql & " and codcampo = " & Text1(7).Text
    Sql = Sql & " and codcalid = " & txtAux(4).Text
    
    If TotalRegistros(Sql) > 0 Then
        MsgBox "Código de calidad ya existe para la variedad en la factura. Reintroduzca.", vbExclamation
        B = False
        PonerFoco txtAux(4)
    End If
    
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

'Private Sub imgBuscar_Click(Index As Integer)
'    TerminaBloquear
'    '++monica
''    BloqueaRegistro "palets", "numpalet = " & Text1(0).Text
'
'     indice = Index + 2
'     Select Case Index
'        Case 0, 1 'variedad y variedad comercial
'            indice = Index + 2
'            Set frmVar = New frmManVariedad
'            frmVar.DatosADevolverBusqueda = "0|1|"
'            frmVar.CodigoActual = Text1(indice).Text
'            frmVar.Show vbModal
'            Set frmVar = Nothing
'            PonerFoco Text1(indice)
'        Case 2 'Marca
'            Set frmMar = New frmManMarcas
'            frmMar.DatosADevolverBusqueda = "0|1|"
'            frmMar.CodigoActual = Text1(4).Text
'            frmMar.Show vbModal
'            Set frmMar = Nothing
'            PonerFoco Text1(4)
'        Case 3 'forfait
'            Set frmFor = New frmManForfaits
'            frmFor.DatosADevolverBusqueda = "0|1|"
'            frmFor.CodigoActual = Text1(5).Text
'            frmFor.Show vbModal
'            Set frmFor = Nothing
'            PonerFoco Text1(5)
'        Case 4 'incidencia
'            indice = 13
'            Set frmIncid = New frmManInciden
'            frmIncid.DatosADevolverBusqueda = "0|1|"
'            frmIncid.CodigoActual = Text1(13).Text
'            frmIncid.Show vbModal
'            Set frmIncid = Nothing
'            PonerFoco Text1(13)
'    End Select
'
'    If Modo = 4 Then BloqueaRegistro "albaran", "numalbar = " & Text1(0).Text
'                'BLOQUEADesdeFormulario2 Me, Data1, 1
'End Sub
'

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
        Case 0 'calidades
            'si es visible|control|tipo campo|nombre campo|ancho control|
            tots = "N||||0|;N||||0|;N||||0|;N||||0|;N||||0|;" 'codtipom, numfactu, fecfactu, codvarie, codcampo
            tots = tots & "S|txtAux(4)|T|Calidad|900|;S|btnBuscar(0)|B|||;"
            tots = tots & "S|txtAux2(4)|T|Denominación|2400|;"
            tots = tots & "S|txtAux(5)|T|Peso Neto|1455|;"
            tots = tots & "S|txtAux(6)|T|Pr.Medio|1455|;"
            tots = tots & "S|txtAux(7)|T|Importe|1455|;"
            
            arregla tots, DataGridAux(Index), Me, 350
        
            DataGridAux(0).Columns(7).Alignment = dbgRight
            DataGridAux(0).Columns(8).Alignment = dbgRight
            DataGridAux(0).Columns(9).Alignment = dbgRight
        
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

Private Function InsertarLinea() As Boolean
'Inserta registre en les taules de Llínies
Dim nomframe As String
Dim bol As Boolean
Dim MenError As String
Dim PesoNeto As String
Dim NumCajas As String
Dim cadena As String

    On Error GoTo EInsertarLinea

    ' *** posa els noms del frames, tant si son de grid com si no ***
    Select Case NumTabMto
        Case 0: nomframe = "FrameAux0" 'incidencias
    End Select
    
    If DatosOkLlin(nomframe) Then
        TerminaBloquear
        '++monica
        BloqueaRegistro "rfactsoc", "codtipom = '" & Trim(Text1(6).Text) & "' and numfactu = " & Text1(0).Text & " and fecfactu = " & DBSet(Text1(1).Text, "F")
        InsertarDesdeForm2 Me, 2, nomframe
        CalcularTotales
    
        '------------------------------------------------------------------------------
        '  LOG de acciones
        
        cadena = Trim(Text1(6).Text) & " " & Text1(0).Text & " " & Text1(1).Text & " " & txtAux(4).Text
        
        Set LOG = New cLOG
        LOG.Insertar 12, vUsu, "Inserta Linea Calidad : " & cadena & vbCrLf
        Set LOG = Nothing
        '-----------------------------------------------------------------------------
    
    Else
        InsertarLinea = False
        Exit Function
    End If

EInsertarLinea:
        If Err.Number <> 0 Then
            MenError = "Insertando Linea." & vbCrLf & "----------------------------" & vbCrLf & MenError
            MuestraError Err.Number, MenError, Err.Description
            InsertarLinea = False
        Else
            InsertarLinea = True
        End If
End Function

Private Function ModificarLinea() As Boolean
'Modifica registre en les taules de Llínies
Dim nomframe As String
Dim V As Integer
Dim bol As Boolean
Dim MenError As String
Dim PesoNeto As String
Dim NumCajas As String
Dim cadena As String

    On Error GoTo eModificarLinea

    ' *** posa els noms del frames, tant si son de grid com si no ***
    nomframe = "FrameAux0" 'calibres
    
    ModificarLinea = False
    If DatosOkLlin(nomframe) Then
        TerminaBloquear
        
        '------------------------------------------------------------------------------
        '  LOG de acciones
        
        cadena = Trim(Text1(6).Text) & " " & Text1(0).Text & " " & Text1(1).Text & " " & txtAux(4).Text
        
        Set LOG = New cLOG
        LOG.Insertar 12, vUsu, "Modifica Linea Calidad : " & cadena & vbCrLf
        Set LOG = Nothing
        '-----------------------------------------------------------------------------
        
        bol = ModificaDesdeFormulario2(Me, 2, nomframe)

        CalcularTotales
'        ActualizarVariedades txtAux(3).Text


'            ModoLineas = 0
'
'            V = AdoAux(NumTabMto).Recordset.Fields(1) 'el 2 es el nº de llinia
'
'            CargaGrid NumTabMto, True
'
'            ' *** si n'hi han tabs ***
''            SituarTab (NumTabMto + 1)
'
'            ' *** si n'hi han tabs que no tenen datagrid, posar el if ***
'            PonerFocoGrid Me.DataGridAux(NumTabMto)
'            AdoAux(NumTabMto).Recordset.Find (AdoAux(NumTabMto).Recordset.Fields(1).Name & " =" & V)
'
'            LLamaLineas NumTabMto, 0
'            ModificarLinea = True
'        End If
        
        '++monica
'        BloqueaRegistro "pedidos", "numpedid = " & Text1(0).Text
        
    End If
eModificarLinea:
    If Err.Number <> 0 Or Not bol Then
        MenError = "Modificando Linea." & vbCrLf & "----------------------------" & vbCrLf & MenError
        MuestraError Err.Number, MenError, Err.Description
        ModificarLinea = False
    Else
        ModificarLinea = True
    End If
End Function

Private Function ObtenerWhereCab(conW As Boolean) As String
Dim vWhere As String
    
    vWhere = ""
    If conW Then vWhere = " WHERE "
    ' *** canviar-ho per la clau primaria de la capçalera ***
    vWhere = vWhere & " codtipom = '" & tipoMov & "' and numfactu = " & Factura
    vWhere = vWhere & " and fecfactu = " & DBSet(fecfactu, "F")
    vWhere = vWhere & " and rfactsoc_calidad.codvarie = " & DBSet(Text1(2).Text, "N")
    vWhere = vWhere & " and rfactsoc_calidad.codcampo = " & DBSet(Text1(7).Text, "N")
    
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

'Private Sub VisualizaPrecio()
'    Select Case vParamAplic.TipoPrecio
'        Case 0
'            txtAux2(0).Text = DevuelveDesdeBDNew(cAgro, "sartic", "preciomp", "codartic", txtAux(1), "T")
'        Case 1
'            txtAux2(0).Text = DevuelveDesdeBDNew(cAgro, "sartic", "preciouc", "codartic", txtAux(1), "T")
'    End Select
'End Sub

Private Sub CalcularTotales()
Dim Rs As ADODB.Recordset
Dim Sql As String
Dim KNetoTotal As String
Dim ImporteTotal As String
Dim PreTotal As String
Dim Valor As Currency
Dim ModoAnt As Integer

    On Error Resume Next

    'total importes de envases para ese forfait
    Sql = "select sum(kilosnet), sum(imporcal) "
    Sql = Sql & " from rfactsoc_calidad where codtipom = '" & Trim(tipoMov) & "'"
    Sql = Sql & " and numfactu = " & Factura
    Sql = Sql & " and fecfactu = " & DBSet(fecfactu, "F")
    Sql = Sql & " and codvarie = " & DBSet(Text1(2).Text, "N")
    Sql = Sql & " and codcampo = " & DBSet(Text1(7).Text, "N")
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    KNetoTotal = 0
    ImporteTotal = 0
    PreTotal = 0
    If Not Rs.EOF Then
        If Rs.Fields(0).Value <> 0 Then KNetoTotal = Rs.Fields(0).Value
        If Rs.Fields(1).Value <> 0 Then ImporteTotal = Rs.Fields(1).Value
        If KNetoTotal <> 0 Then
            PreTotal = Round2(ImporteTotal / KNetoTotal, 4)
        End If
    End If
    Rs.Close
    Set Rs = Nothing
    
    Text1(3).Text = Format(KNetoTotal, "###,##0")
    Text1(4).Text = Format(PreTotal, "#0.0000")
    Text1(5).Text = Format(ImporteTotal, "###,##0.00")
    
 
    ModoAnt = Modo
    BotonModificar
    cmdAceptar_Click
    
    Modo = ModoAnt
    PonerModo Modo
    
    If Err.Number <> 0 Then
        Err.Clear
    End If

End Sub

Private Function ObtenerWhereCP(conW As Boolean) As String
Dim Sql As String
On Error Resume Next
    
    Sql = ""
    If conW Then Sql = " WHERE "
    Sql = Sql & NombreTabla & ".codtipom= " & DBSet(tipoMov, "T")
    Sql = Sql & " and " & NombreTabla & ".numfactu = " & Factura
    Sql = Sql & " and " & NombreTabla & ".fecfactu = " & DBSet(fecfactu, "F")
    Sql = Sql & " and " & NombreTabla & ".codvarie = " & Val(Text1(2).Text)
    Sql = Sql & " and " & NombreTabla & ".codcampo = " & Val(Text1(7).Text)
    
    ObtenerWhereCP = Sql
End Function



Private Function ActualizarVariedades(Variedad As String) As Boolean
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim Sql1 As String
Dim PrecMed As Currency

    On Error GoTo eActualizarVariedades

    ActualizarVariedades = False

    Sql1 = "select sum(kilosnet), sum(imporcal) from rfactsoc_calidad where codtipom = " & DBSet(tipoMov, "T")
    Sql1 = Sql1 & " and numfactu = " & Factura
    Sql1 = Sql1 & " and fecfactu = " & DBSet(fecfactu, "F")
    Sql1 = Sql1 & " and codvarie = " & Text1(2).Text
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql1, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not Rs.EOF Then
        PrecMed = 0
        If DBLet(Rs.Fields(0).Value, "N") <> 0 Then
            PrecMed = Round2(DBLet(Rs.Fields(1).Value, "N") / DBLet(Rs.Fields(0).Value, "N"), 4)
        End If
        
        Sql = "update rfactsoc_variedad set kilosnet = " & DBSet(Rs.Fields(0).Value, "N") & ","
        Sql = Sql & " imporvar = " & DBSet(Rs.Fields(1).Value, "N") & ","
        Sql = Sql & " preciomed = " & DBSet(PrecMed, "N")
        Sql = Sql & " where codtipom = " & DBSet(tipoMov, "T")
        Sql = Sql & " and numfactu = " & DBSet(Factura, "N")
        Sql = Sql & " and fecfactu = " & DBSet(fecfactu, "F")
        Sql = Sql & " and codvarie = " & Variedad

        conn.Execute Sql
    End If
    
    Rs.Close
    Set Rs = Nothing

eActualizarVariedades:
    If Err.Number = 0 Then ActualizarVariedades = True
    
End Function




Private Function Modificar() As Boolean
'Modifica registre en les taules de Llínies
Dim nomframe As String
Dim V As Integer
Dim bol As Boolean
Dim MenError As String
Dim Forfait As String
Dim Sql As String

    On Error GoTo eModificar

    TerminaBloquear
    
    '[Monica]07/10/2013: solo dejo modificar la variedad a montifrut
    '                    cambio la siguiente instruccion por la de abajo
    
'    ModificaDesdeFormulario2 Me, 1

    Sql = "update rfactsoc_variedad set codvarie = " & DBSet(Text1(2).Text, "N")
    Sql = Sql & ", kilosnet = " & DBSet(Text1(3).Text, "N")
    Sql = Sql & ", preciomed = " & DBSet(Text1(4).Text, "N")
    Sql = Sql & ", imporvar = " & DBSet(Text1(5).Text, "N")
    Sql = Sql & " where numfactu = " & DBSet(Factura, "N")
    Sql = Sql & " and codvarie = " & DBSet(Text1(2).Text, "N") '[Monica]21/04/2015: antes variedad
    Sql = Sql & " and codtipom = " & DBSet(tipoMov, "T")
    Sql = Sql & " and fecfactu = " & DBSet(fecfactu, "F")
    Sql = Sql & " and codcampo = " & DBSet(Text1(7).Text, "N")

    conn.Execute Sql


    '------------------------------------------------------------------------------
    '  LOG de acciones
    Dim cadena As String
    If CLng(Me.Data1.Recordset!Codvarie) <> CLng(Text1(2).Text) Then
        cadena = Trim(Text1(6).Text) & " " & Text1(0).Text & " " & Text1(1).Text & " de " & Me.Data1.Recordset!Codvarie & " a " & CInt(Text1(2).Text)
        
        Set LOG = New cLOG
        LOG.Insertar 12, vUsu, "Modificar Variedad : " & cadena & vbCrLf
        Set LOG = Nothing
    End If
    '-----------------------------------------------------------------------------
        



eModificar:
    If Err.Number <> 0 Then
        MenError = "Modificando Registro Nota de Entrada." & vbCrLf & "----------------------------" & vbCrLf & MenError
        MuestraError Err.Number, MenError, Err.Description
        Modificar = False
    Else
        Modificar = True
    End If
End Function

Private Sub CargaCombo()
Dim Rs As ADODB.Recordset
Dim Sql As String
Dim I As Byte
    
    ' *** neteje els combos, els pose valor i seleccione el valor per defecte ***
    For I = 0 To Combo1.Count - 1
        Combo1(I).Clear
    Next I
    
    'tipo de factura
    Sql = "select codtipom, nomtipom from usuarios.stipom where tipodocu > 0 "

    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    I = 1
    While Not Rs.EOF
'        Sql = Replace(Rs.Fields(1).Value, "Factura", "Fac.")
        Sql = Rs.Fields(1).Value
        Sql = Rs.Fields(0).Value & " - " & Sql
        Combo1(0).AddItem Sql 'campo del codigo
        Combo1(0).ItemData(Combo1(0).NewIndex) = I
        I = I + 1
        Rs.MoveNext
    Wend
    
End Sub

