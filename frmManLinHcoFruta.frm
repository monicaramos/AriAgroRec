VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmManLinHcoFruta 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Entradas de Histórico de Fruta Clasificada"
   ClientHeight    =   8580
   ClientLeft      =   45
   ClientTop       =   30
   ClientWidth     =   8685
   Icon            =   "frmManLinHcoFruta.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8580
   ScaleWidth      =   8685
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameAux0 
      Caption         =   "Incidencias"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00972E0B&
      Height          =   2505
      Left            =   135
      TabIndex        =   34
      Top             =   5415
      Width           =   8445
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
         Left            =   1035
         MaxLength       =   9
         TabIndex        =   25
         Tag             =   "Num.Nota|N|N|||rhisfruta_incidencia|numnotac|00000000|S|"
         Text            =   "nota"
         Top             =   1800
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
         Height          =   350
         Index           =   0
         Left            =   2295
         MaskColor       =   &H00000000&
         TabIndex        =   26
         ToolTipText     =   "Buscar incidencia"
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
         Index           =   2
         Left            =   2520
         TabIndex        =   41
         Top             =   1800
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
         Height          =   290
         Index           =   2
         Left            =   1710
         MaxLength       =   4
         TabIndex        =   23
         Tag             =   "Incidencia|N|N|||rhisfruta_incidencia|codincid|0000||"
         Text            =   "inci"
         Top             =   1800
         Visible         =   0   'False
         Width           =   540
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
         TabIndex        =   24
         Tag             =   "Número Albaran|N|N|||rhisfruta_incidencia|numalbar|000000|S|"
         Text            =   "numalbar"
         Top             =   1800
         Visible         =   0   'False
         Width           =   735
      End
      Begin MSComctlLib.Toolbar ToolAux 
         Height          =   390
         Index           =   0
         Left            =   135
         TabIndex        =   35
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
         Bindings        =   "frmManLinHcoFruta.frx":000C
         Height          =   1695
         Index           =   0
         Left            =   135
         TabIndex        =   36
         Top             =   630
         Width           =   7495
         _ExtentX        =   13229
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
   Begin VB.Frame Frame2 
      Height          =   4980
      Index           =   0
      Left            =   120
      TabIndex        =   29
      Top             =   420
      Width           =   8460
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
         Left            =   180
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Tag             =   "Tipo Recolección|N|N|0|1|rhisfruta_entradas|tiporecol||N|"
         Top             =   1680
         Width           =   1680
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
         Left            =   4035
         MaxLength       =   8
         TabIndex        =   11
         Tag             =   "Nro.Trabajadores|N|S|||rhisfruta_entradas|numtraba|##0||"
         Top             =   1680
         Width           =   1665
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
         Left            =   2010
         MaxLength       =   8
         TabIndex        =   10
         Tag             =   "Horas Trabajadas|N|S|||rhisfruta_entradas|horastra|#,##0.00||"
         Top             =   1680
         Width           =   1665
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
         Index           =   17
         Left            =   3165
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   55
         Text            =   "Text2"
         Top             =   3600
         Width           =   4995
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
         Left            =   2010
         MaxLength       =   2
         TabIndex        =   18
         Tag             =   "Tarifa|N|S|||rhisfruta_entradas|codtarif|00||"
         Text            =   "Text1"
         Top             =   3600
         Width           =   1110
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
         Index           =   16
         Left            =   4725
         MaxLength       =   7
         TabIndex        =   7
         Tag             =   "Peso Trans.|N|N|0|999999|rhisfruta_entradas|kilostra|###,##0||"
         Top             =   1050
         Width           =   1515
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
         Left            =   2010
         MaxLength       =   10
         TabIndex        =   17
         Tag             =   "Transportista|T|S|||rhisfruta_entradas|codtrans|||"
         Text            =   "Text1"
         Top             =   3210
         Width           =   1110
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
         Index           =   15
         Left            =   3165
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   52
         Text            =   "Text2"
         Top             =   3210
         Width           =   4995
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
         Index           =   14
         Left            =   3165
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   50
         Text            =   "Text2"
         Top             =   2820
         Width           =   4995
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
         Left            =   2010
         MaxLength       =   4
         TabIndex        =   16
         Tag             =   "Capataz|N|S|0|9999|rhisfruta_entradas|codcapat|0000||"
         Text            =   "Text1"
         Top             =   2820
         Width           =   1110
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
         Left            =   6240
         MaxLength       =   12
         TabIndex        =   8
         Tag             =   "Pr.Estimado|N|S|||rhisfruta_entradas|prestimado|###,##0.0000||"
         Top             =   1050
         Width           =   1515
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
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
         Left            =   4725
         MaxLength       =   8
         TabIndex        =   3
         Top             =   390
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
         Index           =   10
         Left            =   6090
         MaxLength       =   8
         TabIndex        =   15
         Tag             =   "Imp.Penalización|N|S|||rhisfruta_entradas|imppenal|#,##0.00||"
         Top             =   2355
         Width           =   1665
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
         Height          =   540
         Index           =   11
         Left            =   180
         MaxLength       =   200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   19
         Tag             =   "Observaciones|T|S|||rhisfruta_entradas|observac|||"
         Top             =   4290
         Width           =   7965
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
         Left            =   4035
         MaxLength       =   8
         TabIndex        =   14
         Tag             =   "Imp.Recolec|N|S|||rhisfruta_entradas|imprecol|#,##0.00||"
         Top             =   2355
         Width           =   1665
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
         Left            =   180
         MaxLength       =   8
         TabIndex        =   12
         Tag             =   "Imp.Transporte|N|S|||rhisfruta_entradas|imptrans|#,##0.00||"
         Top             =   2355
         Width           =   1665
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
         Left            =   2025
         MaxLength       =   8
         TabIndex        =   13
         Tag             =   "Imp.Acarreo|N|S|||rhisfruta_entradas|impacarr|#,##0.00||"
         Top             =   2355
         Width           =   1665
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
         Left            =   180
         MaxLength       =   7
         TabIndex        =   4
         Tag             =   "Numero Cajas|N|N|0|999999|rhisfruta_entradas|numcajon|###,##0||"
         Top             =   1050
         Width           =   1515
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   3
         Left            =   4725
         MaxLength       =   20
         TabIndex        =   22
         Tag             =   "Hora Ent|FH|N|||rhisfruta_entradas|horaentr|yyyy-mm-dd hh:mm:ss||"
         Top             =   405
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
         Index           =   6
         Left            =   3210
         MaxLength       =   7
         TabIndex        =   6
         Tag             =   "Peso Neto|N|N|0|999999|rhisfruta_entradas|kilosnet|###,##0||"
         Top             =   1050
         Width           =   1515
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
         Left            =   1695
         MaxLength       =   7
         TabIndex        =   5
         Tag             =   "Peso Bruto|N|N|0|999999|rhisfruta_entradas|kilosbru|###,##0||"
         Top             =   1050
         Width           =   1515
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
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
         Left            =   3210
         MaxLength       =   10
         TabIndex        =   2
         Tag             =   "Fecha Entrada|F|N|||rhisfruta_entradas|fechaent|dd/mm/yyyy||"
         Top             =   405
         Width           =   1350
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
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
         Left            =   1695
         MaxLength       =   8
         TabIndex        =   1
         Tag             =   "Nota|N|N|||rhisfruta_entradas|numnotac|00000000|S|"
         Text            =   "12346578"
         Top             =   405
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
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
         Left            =   180
         MaxLength       =   7
         TabIndex        =   0
         Tag             =   "Número albaran|N|N|||rhisfruta_entradas|numalbar|0000000|S|"
         Text            =   "1234567"
         Top             =   405
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo Recolección"
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
         Left            =   180
         TabIndex        =   59
         Top             =   1440
         Width           =   1665
      End
      Begin VB.Label Label1 
         Caption         =   "Nro.Trabajadores"
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
         Left            =   4035
         TabIndex        =   58
         Top             =   1440
         Width           =   1725
      End
      Begin VB.Label Label1 
         Caption         =   "Horas Trabajad."
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
         Left            =   2010
         TabIndex        =   57
         Top             =   1440
         Width           =   1275
      End
      Begin VB.Label Label1 
         Caption         =   "Tarifa"
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
         Index           =   14
         Left            =   180
         TabIndex        =   56
         Top             =   3645
         Width           =   765
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   2
         Left            =   1695
         ToolTipText     =   "Buscar Tarifa"
         Top             =   3630
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Peso Trans."
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
         Left            =   4740
         TabIndex        =   54
         Top             =   825
         Width           =   1485
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   1
         Left            =   1695
         ToolTipText     =   "Buscar Tranportista"
         Top             =   3240
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Transportista"
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
         Left            =   180
         TabIndex        =   53
         Top             =   3255
         Width           =   1380
      End
      Begin VB.Label Label1 
         Caption         =   "Capataz"
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
         Index           =   11
         Left            =   180
         TabIndex        =   51
         Top             =   2865
         Width           =   975
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   1695
         ToolTipText     =   "Buscar Capataz"
         Top             =   2850
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Precio Estimado"
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
         Left            =   6270
         TabIndex        =   49
         Top             =   825
         Width           =   1590
      End
      Begin VB.Label Label1 
         Caption         =   "Imp.Penalización"
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
         Left            =   6090
         TabIndex        =   48
         Top             =   2100
         Width           =   1230
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
         TabIndex        =   47
         Top             =   4020
         Width           =   1515
      End
      Begin VB.Image imgZoom 
         Height          =   240
         Index           =   0
         Left            =   1695
         ToolTipText     =   "Zoom descripción"
         Top             =   4020
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Imp.Recolección"
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
         Left            =   4035
         TabIndex        =   46
         Top             =   2100
         Width           =   1680
      End
      Begin VB.Label Label1 
         Caption         =   "Imp.Transporte"
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
         Left            =   180
         TabIndex        =   45
         Top             =   2100
         Width           =   1665
      End
      Begin VB.Label Label1 
         Caption         =   "Imp.Acarreo"
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
         Left            =   2025
         TabIndex        =   44
         Top             =   2100
         Width           =   1395
      End
      Begin VB.Label Label1 
         Caption         =   "Nro.Cajas"
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
         Left            =   180
         TabIndex        =   43
         Top             =   825
         Width           =   960
      End
      Begin VB.Label Label1 
         Caption         =   "Hora Entrada"
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
         Left            =   4725
         TabIndex        =   42
         Top             =   180
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Peso Neto"
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
         Left            =   3240
         TabIndex        =   40
         Top             =   825
         Width           =   1515
      End
      Begin VB.Label Label1 
         Caption         =   "Peso Bruto"
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
         Left            =   1710
         TabIndex        =   39
         Top             =   825
         Width           =   1470
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha Entrada"
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
         Left            =   3210
         TabIndex        =   38
         Top             =   180
         Width           =   1455
      End
      Begin VB.Label Label6 
         Caption         =   "Nota"
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
         Left            =   1695
         TabIndex        =   37
         Top             =   180
         Width           =   600
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
         Index           =   0
         Left            =   180
         TabIndex        =   30
         Top             =   165
         Width           =   1230
      End
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Index           =   0
      Left            =   120
      TabIndex        =   27
      Top             =   7950
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
         TabIndex        =   28
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
      Left            =   7545
      TabIndex        =   21
      Top             =   8085
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
      Left            =   6420
      TabIndex        =   20
      Top             =   8085
      Width           =   1035
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   1980
      Top             =   5475
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
      TabIndex        =   32
      Top             =   0
      Visible         =   0   'False
      Width           =   8685
      _ExtentX        =   15319
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
            Enabled         =   0   'False
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir"
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
      Enabled         =   0   'False
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   195
         Index           =   0
         Left            =   6525
         TabIndex        =   33
         Top             =   90
         Width           =   1215
      End
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
      Left            =   7545
      TabIndex        =   31
      Top             =   8100
      Visible         =   0   'False
      Width           =   1035
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
Attribute VB_Name = "frmManLinHcoFruta"
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
Public Albaran As Long
Public Nota As Long

Public ModoExt As Byte

' *** declarar els formularis als que vaig a cridar ***
Private WithEvents frmB As frmBuscaGrid
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmC As frmCal 'calendario fecha
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmZ As frmZoom  'Zoom para campos Text
Attribute frmZ.VB_VarHelpID = -1


Private WithEvents frmInc As frmManInciden 'incidencias
Attribute frmInc.VB_VarHelpID = -1
Private WithEvents frmCap As frmManCapataz 'capataces
Attribute frmCap.VB_VarHelpID = -1
Private WithEvents frmTrans As frmManTranspor 'transportistas
Attribute frmTrans.VB_VarHelpID = -1
Private WithEvents frmTar As frmManTarTra 'tarifas de transporte
Attribute frmTar.VB_VarHelpID = -1


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
Dim CodTarifAnt As String

Dim CodTipoMov As String

Private Sub btnBuscar_Click(Index As Integer)
    TerminaBloquear
    '++monica
'    BloqueaRegistro "palets", "numpalet = " & Text1(0).Text
    
    Select Case Index
        Case 0 'incidencias
            Set frmInc = New frmManInciden
            frmInc.DatosADevolverBusqueda = "0|1|"
            frmInc.CodigoActual = txtAux(1).Text
'            frmInc.ParamVariedad = txtAux(4).Text
            frmInc.Show vbModal
            Set frmInc = Nothing
            PonerFoco txtAux(1)
    End Select
    If Modo = 4 Then BloqueaRegistro "albaran", "numalbar = " & Text1(0).Text
    'BLOQUEADesdeFormulario2 Me, Data1, 1
End Sub


Private Sub cmdAceptar_Click()
Dim B As Boolean
Dim V As Integer
Dim Forfait As String
Dim vTipoMov As CTiposMov

    Screen.MousePointer = vbHourglass
    
    On Error GoTo Error1
    
    Select Case Modo
        Case 1  'BÚSQUEDA
            HacerBusqueda
            
        Case 3 'INSERTAR
            
            '[Monica]10/09/2012:Mogente
            If Not vParamAplic.NroNotaManual Then
                Set vTipoMov = New CTiposMov
                If vTipoMov.Leer(CodTipoMov) Then
                    Text1(1).Text = vTipoMov.ConseguirContador(CodTipoMov)
                End If
                Text1(1).Text = Format(Text1(1).Text, "0000000")
            End If
            
            If DatosOK Then
            
                If InsertarDesdeForm2(Me, 1) Then
'                    text2(9).Text = PonerNombreCuenta(text1(9), Modo, text1(0).Text)
        
'                    Data1.RecordSource = "Select * from " & NombreTabla & _
'                                        " where numpalet = " & DBSet(text1(0).Text, "N") & _
'                                        " and numlinea = " & DBSet(text1(1).Text, "N") & " " & Ordenacion
'                    PosicionarData

                    '[Monica]10/09/2012:Mogente
                    If Not vParamAplic.NroNotaManual Then
                        vTipoMov.IncrementarContador (CodTipoMov)
                    End If
                    
                    '[Monica]10/09/2012:Mogente
                    ActualizarClasificacionHco Text1(0).Text, Text1(6).Text

                    TerminaBloquear
                    BloqueaRegistro "rhisfruta", "numalbar = " & Text1(0).Text
                    
                    CadenaConsulta = "Select * from " & NombreTabla & ObtenerWhereCP(True) & Ordenacion
                    PonerCadenaBusqueda
                    'Ponerse en Modo Insertar Lineas
                    
                    '[Monica]10/09/2012:Mogente
                    If vParamAplic.Cooperativa <> 3 Then
                        BotonAnyadirLinea 0
                    End If

                End If
            Else
                ModoLineas = 0
            End If
            
        Case 4  'MODIFICAR
            If DatosOK Then
                Modificar
                TerminaBloquear
                '++monica
                BloqueaRegistro "rhisfruta", "numalbar = " & Text1(0).Text
                
                PosicionarData
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
                        BloqueaRegistro "rhisfruta", "numalbar = " & Text1(0).Text
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


' *** si n'hi han combos a la capçalera ***
Private Sub Combo1_GotFocus(Index As Integer)
    If Modo = 1 Then Combo1(Index).BackColor = vbLightBlue
  
End Sub

Private Sub Combo1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Combo1_LostFocus(Index As Integer)
Dim B As Boolean

    If Combo1(Index).BackColor = vbLightBlue Then Combo1(Index).BackColor = vbWhite
    
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
    If PrimeraVez Then
        PrimeraVez = False
    
        PonerCampos
        ModoLineas = 0
           
        CalcularTotales
        
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
    If Text1(0).Text <> "" And Text1(1).Text <> "" Then
        cad = Text1(0).Text & "|" & Text1(1).Text & "|"
    End If
    RaiseEvent DatoSeleccionado(cad)

    CheckValueGuardar Me.Name, Me.chkVistaPrevia(0).Value
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
        .Buttons(11).Image = 19   'Expandir Añadir, Borrar y Modificar
        'el 10 i el 11 son separadors
        .Buttons(12).Image = 10  'Imprimir
        .Buttons(13).Image = 11  'Eixir
        'el 13 i el 14 son separadors
        .Buttons(btnPrimero).Image = 6  'Primer
        .Buttons(btnPrimero + 1).Image = 7 'Anterior
        .Buttons(btnPrimero + 2).Image = 8 'Següent
        .Buttons(btnPrimero + 3).Image = 9 'Últim
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
    NombreTabla = "rhisfruta_entradas"
    Ordenacion = " ORDER BY numalbar"
    
    'Mirem com està guardat el valor del check
    chkVistaPrevia(0).Value = CheckValueLeer(Name)
    
    Data1.ConnectionString = conn
    '***** cambiar el nombre de la PK de la cabecera *************
    Data1.RecordSource = "Select * from " & NombreTabla & " where numalbar=" & Albaran & " and numnotac = " & Nota
    Data1.Refresh
    
    CodTipoMov = "NOC"
    
'    If DatosADevolverBusqueda = "" Then
'        PonerModo 0
'    Else
'        PonerModo 1 'búsqueda
'        ' *** posar de groc els camps visibles de la clau primaria de la capçalera ***
'        Text1(0).BackColor = vbLightBlue 'codforfait
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
    
    Text1(5).Enabled = True
    
    
    '=======================================
    B = (Modo = 2)
    'Posar Fleches de desplasament visibles
    NumReg = 1
    If Not Data1.Recordset.EOF Then
        If Data1.Recordset.RecordCount > 1 Then NumReg = 2 'Només es per a saber que n'hi ha + d'1 registre
    End If
    DesplazamientoVisible Me.Toolbar1, btnPrimero, B, NumReg
    '---------------------------------------------
    
    B = Modo <> 0 And Modo <> 2
    cmdCancelar.visible = B
    cmdAceptar.visible = B
    cmdRegresar.visible = Not B

    'Bloqueja els camps Text1 si no estem modificant/Insertant Datos
    'Si estem en Insertar a més neteja els camps Text1
    BloquearText1 Me, Modo
    
    BloquearImgBuscar Me, Modo
    
    Text1(12).Locked = Not (Modo = 3 Or Modo = 4 Or Modo = 1)  '((Not b) And (Modo <> 1))
    If (Modo = 3 Or Modo = 4 Or Modo = 1) Then
        Text1(12).BackColor = vbWhite
    Else
        Text1(12).BackColor = &H80000018 'groc
    End If

    
    '*** si n'hi han combos a la capçalera ***
    BloquearCombo Me, Modo
    '**************************
    
    ' ***** bloquejar tots els controls visibles de la clau primaria de la capçalera ***
    If Modo = 4 Then
        BloquearTxt Text1(0), True 'si estic en  modificar, bloqueja la clau primaria
        BloquearTxt Text1(1), True 'si estic en  modificar, bloqueja la clau primaria
    End If
    ' **********************************************************************************
    
'    ' numero de cajas, peso bruto y peso neto siempre bloqueados
'    BloquearTxt Text1(7), True
'    BloquearTxt Text1(8), True
'    BloquearTxt Text1(10), True
'    BloquearTxt Text1(16), True
    
    ' **** si n'hi han imagens de buscar en la capçalera *****
'    BloquearImgBuscar Me, Modo, ModoLineas
    BloquearImgZoom Me, Modo, ModoLineas
    ' ********************************************************
'    imgBuscar(0).visible = (Modo = 3)
'    imgBuscar(0).Enabled = (Modo = 3)
    
        
    chkVistaPrevia(0).Enabled = (Modo <= 2)
    
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
    Toolbar1.Buttons(3).Enabled = B
    Me.mnBuscar.Enabled = B
    'Vore Tots
    Toolbar1.Buttons(4).Enabled = B
    Me.mnVerTodos.Enabled = B
    
    'Insertar
    Toolbar1.Buttons(7).Enabled = B And Not DeConsulta
    Me.mnNuevo.Enabled = B And Not DeConsulta
    
    B = (Modo = 2 And Data1.Recordset.RecordCount > 0) And Not DeConsulta
    'Modificar
    Toolbar1.Buttons(8).Enabled = B
    Me.mnModificar.Enabled = B
    'eliminar
    Toolbar1.Buttons(9).Enabled = B
    Me.mnEliminar.Enabled = B
    
    'Expandir operaciones
    Toolbar1.Buttons(11).Enabled = True And Not DeConsulta
    'Imprimir
    'Toolbar1.Buttons(12).Enabled = (b Or Modo = 0)
    Toolbar1.Buttons(12).Enabled = True And Not DeConsulta
       
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
               
        Case 0 'INCIDENCIAS
            Sql = "SELECT rhisfruta_incidencia.numalbar, rhisfruta_incidencia.numnotac, rhisfruta_incidencia.codincid, "
            Sql = Sql & "rincidencia.nomincid "
            Sql = Sql & " FROM rhisfruta_incidencia, rincidencia "
            If enlaza Then
                Sql = Sql & ObtenerWhereCab(True)
            Else
                Sql = Sql & " WHERE rhisfruta_incidencia.numalbar = '-1'"
            End If
            Sql = Sql & " and rhisfruta_incidencia.codincid = rincidencia.codincid "
            Sql = Sql & " ORDER BY rhisfruta_incidencia.codincid "
               
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

Private Sub frmCap_DatoSeleccionado(CadenaSeleccion As String)
    Text1(Indice).Text = RecuperaValor(CadenaSeleccion, 1) 'codcapataz
    PonerFormatoEntero Text1(Indice)
    Text2(Indice).Text = RecuperaValor(CadenaSeleccion, 2) 'descripcion
End Sub

Private Sub frmInc_DatoSeleccionado(CadenaSeleccion As String)
    txtAux(2).Text = RecuperaValor(CadenaSeleccion, 1) 'codincid
    txtAux2(2).Text = RecuperaValor(CadenaSeleccion, 2) 'descripcion
End Sub

Private Sub frmTar_DatoSeleccionado(CadenaSeleccion As String)
    Text1(Indice).Text = Format(RecuperaValor(CadenaSeleccion, 1), "00") 'codigo de tarifa
    Text2(Indice).Text = RecuperaValor(CadenaSeleccion, 2) 'nombre
End Sub

Private Sub frmTrans_DatoSeleccionado(CadenaSeleccion As String)
    Text1(Indice).Text = RecuperaValor(CadenaSeleccion, 1) 'codtrasnportista
    Text2(Indice).Text = RecuperaValor(CadenaSeleccion, 2) 'nombre
End Sub

Private Sub frmZ_Actualizar(vCampo As String)
     Text1(Indice).Text = vCampo
End Sub

Private Sub imgBuscar_Click(Index As Integer)
            
     Select Case Index
        Case 0
            Indice = 14
            PonerFoco Text1(Indice)
            Set frmCap = New frmManCapataz
            frmCap.DatosADevolverBusqueda = "0|1|"
            frmCap.Show vbModal
            Set frmCap = Nothing
            PonerFoco Text1(Indice)
        Case 1
            Indice = 15
            PonerFoco Text1(Indice)
            Set frmTrans = New frmManTranspor
            frmTrans.DatosADevolverBusqueda = "0|1|"
            frmTrans.Show vbModal
            Set frmTrans = Nothing
            PonerFoco Text1(Indice)
            
        Case 2 ' Codigo de tarifa
            Indice = 17
            PonerFoco Text1(Indice)
            Set frmTar = New frmManTarTra
            frmTar.DatosADevolverBusqueda = "0|1|"
            frmTar.Show vbModal
            Set frmTar = Nothing
            PonerFoco Text1(Indice)
        
     End Select

End Sub

Private Sub imgZoom_Click(Index As Integer)
    
    Set frmZ = New frmZoom

    If Index = 0 Then
        Indice = 11
        frmZ.pTitulo = "Observaciones de la Nota de Entrada de Albarán"
        frmZ.pValor = Text1(Indice).Text
        frmZ.pModo = Modo
    
        frmZ.Show vbModal
        Set frmZ = Nothing
            
        PonerFoco Text1(Indice)
    End If

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
'        Text1(0).BackColor = vbLightBlue ' <===
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
    LimpiarCampos 'Neteja els Text1
    CadB = ""
    
    If chkVistaPrevia(0).Value = 1 Then
        MandaBusquedaPrevia ""
    Else
        CadenaConsulta = "Select * from " & NombreTabla & Ordenacion
        PonerCadenaBusqueda
    End If
End Sub

Private Sub BotonAnyadir()

    LimpiarCampos 'Huida els TextBox
    
    
    PonerModo 3
    
    ' ****** Valors per defecte a l'afegir, repasar si n'hi ha
    ' codEmpre i quins camps tenen la PK de la capçalera *******
'    text1(0).Text = SugerirCodigoSiguienteStr("forfaits", "codforfait")
'    FormateaCampo text1(0)
    
    Text1(0).Text = Albaran
'    Text1(1).Text = SugerirCodigoSiguienteStr("albaran_variedad", "numlinea", "numalbar = " & Text1(0).Text)
    Text1(1).Text = ""
    Text1(0).BackColor = &HFFFFC0 '&H80000013
    Text1(1).BackColor = &HFFFFC0 '&H80000013
    Text1(0).Locked = True
'    Text1(1).Locked = True
    
    PonerFoco Text1(1) '*** 1r camp visible que siga PK ***
    
    ' *** si n'hi han camps de descripció a la capçalera ***
    'PosarDescripcions


    '[Monica]10/09/2012: entradas para Mogente más rapidas sin pasar por bascula
    Text1(2).Text = DevuelveDesdeBDNew(cAgro, "rhisfruta", "fecalbar", "numalbar", CStr(Albaran), "N")
    Text1(12).Text = Time
    
    If vParamAplic.NroNotaManual Then
        'claveprimaria
        BloquearTxt Text1(1), False
        PonerFoco Text1(1) '*** 1r camp visible que siga PK ***
        
    Else
        'Campos Nº entrada bloqueado y en azul
        BloquearTxt Text1(1), True, True
    
        PonerFoco Text1(2) '*** 1r camp visible que siga PK ***
    End If


End Sub

Private Sub BotonModificar()

    PonerModo 4
    
    Text1(0).Text = Albaran
    Text1(1).Text = Nota
    
    Text1(0).BackColor = &H80000013
    Text1(1).BackColor = &H80000013

    ' *** bloquejar els camps visibles de la clau primaria de la capçalera ***
    BloquearTxt Text1(0), True
    BloquearTxt Text1(1), True
    
    CodTarifAnt = Text1(17).Text
    
    ' *** foco al 1r camp visible que NO siga clau primaria ***
    PonerFoco Text1(2)
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
    
    Text1(12).Text = Mid(Text1(3).Text, 12, 8)
    
    Text2(14) = PonerNombreDeCod(Text1(14), "rcapataz", "nomcapat")
    Text2(15) = PonerNombreDeCod(Text1(15), "rtransporte", "nomtrans")
    Text2(17) = PonerNombreDeCod(Text1(17), "rtarifatra", "nomtarif")
    
    
    ' *** si n'hi han llínies en datagrids ***
    'For i = 0 To DataGridAux.Count - 1
    For I = 0 To 0
            CargaGrid I, True
            If Not Adoaux(I).Recordset.EOF Then _
                PonerCamposForma2 Me, Adoaux(I), 2, "FrameAux" & I
    Next I

    
'    CalcularTotales
    
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
                BloqueaRegistro "albaran", "numalbar= " & Text1(0).Text
                
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
 
    Text1(3).Text = Format(Text1(2).Text, "dd/mm/yyyy") & " " & Format(Text1(12).Text, "HH:MM:SS")
    
    B = CompForm2(Me, 1)
    If Not B Then Exit Function
    
    ' *** canviar els arguments de la funcio, el mensage i repasar si n'hi ha codEmpre ***
    If (Modo = 3) Then 'insertar
    
        'comprobar si existe ya el cod. del campo clave primaria
        Sql = ""
        Sql = DevuelveDesdeBDNew(cAgro, "rhisfruta_entradas", "numalbar", "numalbar", Text1(0).Text, "N", , "numnotac", Text1(1).Text, "N")
        If Sql <> "" Then
            MsgBox "Ya existe el numero de nota para este albarán", vbExclamation
            B = False
        End If
    End If
    
    ' ************************************************************************************
    
    '[Monica]29/11/2017: comprobamos recolectado por y transportado por
    '                    de momento solo para picassent, deberia generalizarlo
    If B Then
        If vParamAplic.Cooperativa = 2 Then
            If ComprobarCero(Text1(14).Text) = 0 And EntradaRecolectadaporCooperativa(CStr(Albaran)) Then
                MsgBox "Si la entrada está recolectada por la cooperativa, debe introducir capataz. Revise.", vbExclamation
                B = False
                PonerFoco Text1(14)
            End If
            If ComprobarCero(Text1(14).Text) <> 0 And Not EntradaRecolectadaporCooperativa(CStr(Albaran)) Then
                MsgBox "Si la entrada está recolectada por el socio, no debe introducir capataz. Revise.", vbExclamation
                B = False
                PonerFoco Text1(14)
            End If
        End If
    End If
    If B Then
        If vParamAplic.Cooperativa = 2 Then
            If ComprobarCero(Text1(15).Text) = 0 And EntradaTransportadaporCooperativa(CStr(Albaran)) Then
                MsgBox "Si la entrada está transportada por la cooperativa, debe introducir transportista. Revise.", vbExclamation
                B = False
                PonerFoco Text1(15)
            End If
            If ComprobarCero(Text1(15).Text) <> 0 And Not EntradaTransportadaporCooperativa(CStr(Albaran)) Then
                MsgBox "Si la entrada está transportada por el socio, no debe introducir transportista. Revise.", vbExclamation
                B = False
                PonerFoco Text1(15)
            End If
        End If
    End If
    
    
    DatosOK = B
    
EDatosOK:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function


Private Function EntradaRecolectadaporCooperativa(Albaran As String) As Boolean
Dim Sql As String

    Sql = "select recolect from rhisfruta where numalbar = " & DBSet(Albaran, "N")
    EntradaRecolectadaporCooperativa = (DevuelveValor(Sql) = 0)
    
End Function

Private Function EntradaTransportadaporCooperativa(Albaran As String) As Boolean
Dim Sql As String

    Sql = "select transportadopor from rhisfruta where numalbar = " & DBSet(Albaran, "N")
    EntradaTransportadaporCooperativa = (DevuelveValor(Sql) = 0)

End Function




Private Sub PosicionarData()
Dim cad As String, Indicador As String

    ' *** canviar-ho per tota la PK de la capçalera, no llevar els () ***
    cad = "(numalbar=" & DBSet(Text1(0).Text, "N") & ")"
    cad = cad & " and (numnotac = " & DBSet(Text1(1).Text, "N") & ")"
    
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
        Case 0 'codigo de forfait
            Text1(Index).Text = UCase(Text1(Index).Text)
            
        Case 2 ' fecha de entradas
            If Text1(Index).Text <> "" Then PonerFormatoFecha Text1(2), True
        
        Case 4, 5, 6, 16 'cajas, peso bruto y peso neto, peso trans
            PonerFormatoEntero Text1(Index)
            
        Case 7, 8, 9, 10
            If Text1(Index).Text <> "" Then PonerFormatoDecimal Text1(Index), 10

        Case 12
            If Modo = 1 Then Exit Sub
            PonerFormatoHora Text1(Index)

        Case 13 ' precio estimado para Valsur
            PonerFormatoDecimal Text1(Index), 11 'decimal(10,4)

        Case 19 ' nro de trabajadores
            PonerFormatoEntero Text1(Index)
        
        Case 18 ' horas trabajadas
            If Modo = 1 Then Exit Sub
            PonerFormatoDecimal Text1(Index), 3
        
        Case 11
'            cmdAceptar.SetFocus
            
        Case 14 ' codigo de capataz
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index) = PonerNombreDeCod(Text1(Index), "rcapataz", "nomcapat")
                If Text2(Index).Text = "" Then
                    cadMen = "No existe el Capataz. Reintroduzca. " & vbCrLf
                    MsgBox cadMen, vbExclamation
                    PonerFoco Text1(Index)
                End If
            Else
                Text2(Index).Text = ""
            End If
        
        Case 15 'Transportista
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index) = PonerNombreDeCod(Text1(Index), "rtransporte", "nomtrans")
                If Text2(Index).Text = "" Then
                    cadMen = "No existe el Transportista. Reintroduzca. " & vbCrLf
                    MsgBox cadMen, vbExclamation
                    PonerFoco Text1(Index)
                End If
            Else
                Text2(Index).Text = ""
            End If
            
        Case 17 'codigo de tarifa
            If PonerFormatoEntero(Text1(Index)) Then
                Text2(Index) = PonerNombreDeCod(Text1(Index), "rtarifatra", "nomtarif")
                If Text2(Index).Text = "" Then
                    cadMen = "No existe la Tarifa de Transporte. Reintroduzca." & vbCrLf
                    MsgBox cadMen, vbExclamation
                    PonerFoco Text1(Index)
                Else
                    '[Monica]02/04/2012: solo en el caso de picassent si me cambian la tarifa cambio el importe de acarreo
                    If CInt(ComprobarCero(CodTarifAnt)) <> CInt(ComprobarCero(Text1(17).Text)) Then
                        Dim Precio As Currency
                        Dim Transporte As Currency
                        Dim TipoEntr As String
                        
                        If vParamAplic.Cooperativa = 2 Or vParamAplic.Cooperativa = 16 Then
                            Precio = DevuelveValor("select preciokg from rtarifatra where codtarif = " & DBSet(Text1(Index).Text, "N"))
                            Transporte = Round2(Text1(16).Text * Precio, 2)
                            ' dentro del importe de acarreo
                            Text1(8).Text = Format(Transporte, "#,##0.00")
                            TipoEntr = DevuelveDesdeBDNew(cAgro, "rhisfruta", "tipoentr", "numalbar", Text1(0).Text, "N")
                            If vParamAplic.TipoPortesTRA And CInt(TipoEntr) <> 1 Then
                                MsgBox "Se ha modificado el codigo de tarifa y los gastos de acarreo. " & vbCrLf & "Revisar los gastos de socio.", vbExclamation
                            End If
                        End If
                    End If
                End If
            Else
                Text2(Index).Text = ""
            End If
        
            
    End Select
        ' ***************************************************************************
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index <> 11 Then
        If KeyAscii = teclaBuscar Then
            If Modo = 1 Or Modo = 3 Or Modo = 4 Then
                Select Case Index
                    Case 14: KEYBusqueda KeyAscii, 0 'Capataz
                    Case 15: KEYBusqueda KeyAscii, 1 'Transportista
                    Case 17: KEYBusqueda KeyAscii, 2 'Tarifa
                End Select
            End If
        Else
            KEYpress KeyAscii
        End If
    Else
        If Text1(Index) = "" And KeyAscii = teclaBuscar Then
            imgZoom_Click (0)
        Else
            KEYpress KeyAscii
        End If
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
            BotonModificarLinea Index
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
        Case 0 'calibres
            Sql = "¿Seguro que desea eliminar la Incidencia?"
            Sql = Sql & vbCrLf & "Incidencia: " & Adoaux(Index).Recordset!codincid
            If MsgBox(Sql, vbQuestion + vbYesNo) = vbYes Then
                Eliminar = True
                Sql = "DELETE FROM rhisfruta_incidencia "
                Sql = Sql & vWhere & " AND codincid= " & Adoaux(Index).Recordset!codincid
            End If
            
    End Select

    If Eliminar Then
        NumRegElim = Adoaux(Index).Recordset.AbsolutePosition
        TerminaBloquear
        BloqueaRegistro "rhisfruta", "numalbar = " & Text1(0).Text
        
        conn.Execute Sql
        
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
        Case 0: vtabla = "rhisfruta_incidencia"
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
                anc = anc + 240
            Else
                anc = anc + DataGridAux(Index).RowTop(DataGridAux(Index).Row) + 5
            End If
            
            LLamaLineas Index, ModoLineas, anc
        
            Select Case Index
                ' *** valor per defecte a l'insertar i formateig de tots els camps ***
                Case 0 'incidencias
                    txtAux(0).Text = Text1(0).Text 'numalbar
                    txtAux(1).Text = Text1(1).Text 'numnotac
                    txtAux(2).Text = ""
                    txtAux2(2).Text = ""
                    BloquearbtnBuscar Me, Modo, ModoLineas, "FrameAux0"
                    PonerFoco txtAux(2)
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
                anc = anc + 240
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
        Case 0 'incidencias
            txtAux(2).visible = B 'codincie
            txtAux(2).Top = alto
            txtAux2(2).visible = B
            txtAux2(2).Top = alto
            btnBuscar(0).visible = B
            btnBuscar(0).Top = alto
            
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
        Case 2 ' codigo de incidencia
            If PonerFormatoEntero(txtAux(Index)) Then
                txtAux2(2).Text = DevuelveDesdeBDNew(cAgro, "rincidencia", "nomincid", "codincid", txtAux(2).Text, "N")
                If txtAux2(2).Text = "" Then
                    cadMen = "No existe la Incidencia: " & txtAux(Index).Text & vbCrLf
                    cadMen = cadMen & "¿Desea crearla?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmInc = New frmManInciden
                        frmInc.DatosADevolverBusqueda = "0|1|"
                        frmInc.NuevoCodigo = txtAux(Index).Text
                        txtAux(Index).Text = ""
                        TerminaBloquear
                        '++monica
                        
                        frmInc.Show vbModal
                        Set frmInc = Nothing
'                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
                        BloqueaRegistro "rhisfruta_incidencia", "numalbar = " & Text1(0).Text
                    Else
                        txtAux(Index).Text = ""
                    End If
                    PonerFoco txtAux(Index)
                End If
                cmdAceptar.SetFocus
            Else
                txtAux2(2).Text = ""
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
                    Case 2: 'incidencia
                        KeyAscii = 0
                        btnBuscar_Click (0)
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
        Case 0 'incidencias
            'si es visible|control|tipo campo|nombre campo|ancho control|
            tots = "N||||0|;N||||0|;" 'numalbar,numnotac,codincid,nomincid
            tots = tots & "S|txtAux(2)|T|Código|1100|;S|btnBuscar(0)|B|||;"
            tots = tots & "S|txtAux2(2)|T|Denominación|5800|;"
            
            arregla tots, DataGridAux(Index), Me, 350
        
            DataGridAux(0).Columns(2).Alignment = dbgLeft
            DataGridAux(0).Columns(3).Alignment = dbgLeft
        
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

    On Error GoTo EInsertarLinea

    ' *** posa els noms del frames, tant si son de grid com si no ***
    Select Case NumTabMto
        Case 0: nomframe = "FrameAux0" 'incidencias
    End Select
    
    If DatosOkLlin(nomframe) Then
        TerminaBloquear
        '++monica
        BloqueaRegistro "rhisfruta_entradas", "numalbar = " & Text1(0).Text
        InsertarDesdeForm2 Me, 2, nomframe
         
'
'            b = BLOQUEADesdeFormulario2(Me, Data1, 1)
'            Select Case NumTabMto
'                Case 0, 1 ' *** els index de les llinies en grid (en o sense tab) ***
'                    CargaGrid NumTabMto, True
'                    If b Then BotonAnyadirLinea NumTabMto
'            End Select
'
'            SituarTab (NumTabMto + 1)
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
    
    On Error GoTo eModificarLinea

    ' *** posa els noms del frames, tant si son de grid com si no ***
    nomframe = "FrameAux0" 'calibres
    
    ModificarLinea = False
    If DatosOkLlin(nomframe) Then
        TerminaBloquear
        
        
        bol = ModificaDesdeFormulario2(Me, 2, nomframe)
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
    vWhere = vWhere & " numalbar=" & Me.Data1.Recordset!numalbar & " and numnotac = " & Me.Data1.Recordset!NumNotac
    
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
Dim TotalEnvases As String
Dim TotalCostes As String
Dim Valor As Currency

    On Error Resume Next

    'total importes de envases para ese forfait
    Sql = "select sum(numcajas) "
    Sql = Sql & " from albaran_calibre where numalbar = " & DBSet(Text1(0).Text, "N")
    Sql = Sql & " and numlinea = " & DBSet(Text1(1).Text, "N")
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    TotalEnvases = 0
    If Not Rs.EOF Then
        If Rs.Fields(0).Value > 0 Then TotalEnvases = Rs.Fields(0).Value
    End If
    Rs.Close
    Set Rs = Nothing
    
'    Text3(0).Text = Format(TotalEnvases, "###,##0")
    If Err.Number <> 0 Then
        Err.Clear
    End If

End Sub

Private Function ObtenerWhereCP(conW As Boolean) As String
Dim Sql As String
On Error Resume Next
    
    Sql = ""
    If conW Then Sql = " WHERE "
    Sql = Sql & NombreTabla & ".numalbar= " & DBSet(Text1(0).Text, "N")
    Sql = Sql & " and " & NombreTabla & ".numnotac=" & Val(Text1(1).Text)
    ObtenerWhereCP = Sql
End Function



Private Function ActualizarVariedades(Albaran As String, Linea As String) As Boolean
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim Sql1 As String

    On Error GoTo eActualizarVariedades

    ActualizarVariedades = False

    Sql1 = "select sum(pesobrut), sum(pesoneto), sum(numcajas), sum(unidades) from albaran_calibre where numalbar = " & DBSet(Albaran, "N")
    Sql1 = Sql1 & " and numlinea = " & DBSet(Linea, "N")
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql1, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not Rs.EOF Then
        If DBLet(Rs.Fields(0).Value, "N") = 0 Then
            Sql = "update albaran_variedad set pesobrut = null "
            Sql = Sql & " where numalbar = " & DBSet(Albaran, "N")
            Sql = Sql & " and numlinea = " & DBSet(Linea, "N")
    
            conn.Execute Sql
        End If
        If DBLet(Rs.Fields(1).Value, "N") = 0 Then
            Sql = "update albaran_variedad set pesoneto = null "
            Sql = Sql & " where numalbar = " & DBSet(Albaran, "N")
            Sql = Sql & " and numlinea = " & DBSet(Linea, "N")
    
            conn.Execute Sql
        End If
        If DBLet(Rs.Fields(2).Value, "N") = 0 Then
            Sql = "update albaran_variedad set numcajas = null "
            Sql = Sql & " where numalbar = " & DBSet(Albaran, "N")
            Sql = Sql & " and numlinea = " & DBSet(Linea, "N")
    
            conn.Execute Sql
        End If
        If DBLet(Rs.Fields(3).Value, "N") = 0 Then
            Sql = "update albaran_variedad set unidades = null "
            Sql = Sql & " where numalbar = " & DBSet(Albaran, "N")
            Sql = Sql & " and numlinea = " & DBSet(Linea, "N")
    
            conn.Execute Sql
        End If
        
        If DBLet(Rs.Fields(0).Value, "N") <> 0 Then
            Sql = "update albaran_variedad set pesobrut = " & DBSet(Rs.Fields(0).Value, "N")
            Sql = Sql & " where numalbar = " & DBSet(Albaran, "N")
            Sql = Sql & " and numlinea = " & DBSet(Linea, "N")
    
            conn.Execute Sql
        End If
        
        If DBLet(Rs.Fields(1).Value, "N") <> 0 Then
            Sql = "update albaran_variedad set pesoneto = " & DBSet(Rs.Fields(1).Value, "N")
            Sql = Sql & " where numalbar = " & DBSet(Albaran, "N")
            Sql = Sql & " and numlinea = " & DBSet(Linea, "N")
    
            conn.Execute Sql
        End If
        If DBLet(Rs.Fields(2).Value, "N") <> 0 Then
            Sql = "update albaran_variedad set numcajas = " & DBSet(Rs.Fields(2).Value, "N")
            Sql = Sql & " where numalbar = " & DBSet(Albaran, "N")
            Sql = Sql & " and numlinea = " & DBSet(Linea, "N")
    
            conn.Execute Sql
        End If
        If DBLet(Rs.Fields(3).Value, "N") <> 0 Then
            Sql = "update albaran_variedad set unidades = " & DBSet(Rs.Fields(3).Value, "N")
            Sql = Sql & " where numalbar = " & DBSet(Albaran, "N")
            Sql = Sql & " and numlinea = " & DBSet(Linea, "N")
    
            conn.Execute Sql
        End If
    
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
    
    On Error GoTo eModificar

    TerminaBloquear
    
    ModificaDesdeFormulario2 Me, 1

    '[Monica]10/09/2012: para optimizar la entrada en Mogente
    ActualizarClasificacionHco Text1(0).Text, Text1(6).Text

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
Dim Ini As Integer
Dim Fin As Integer
Dim I As Integer

    ' *** neteje els combos, els pose valor i seleccione el valor per defecte ***
    Combo1(0).Clear
    
    'tipo de recoleccion
    Combo1(0).AddItem "Horas"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 0
    Combo1(0).AddItem "Destajo"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 1
    
End Sub


