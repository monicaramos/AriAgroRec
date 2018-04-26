VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmADVTratamientos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tratamientos ADV"
   ClientHeight    =   8190
   ClientLeft      =   45
   ClientTop       =   30
   ClientWidth     =   9705
   Icon            =   "frmADVTratamientos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8190
   ScaleWidth      =   9705
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "Cuenta Principal|N|N|0|1|cltebanc|ctaprpal|0||"
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
      Left            =   8505
      TabIndex        =   6
      Top             =   7530
      Width           =   1035
   End
   Begin VB.Frame FrameBotonGnral 
      Height          =   705
      Left            =   180
      TabIndex        =   28
      Top             =   135
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
   Begin VB.Frame FrameDesplazamiento 
      Height          =   705
      Left            =   3855
      TabIndex        =   26
      Top             =   135
      Width           =   2415
      Begin MSComctlLib.Toolbar ToolbarDes 
         Height          =   330
         Left            =   240
         TabIndex        =   27
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
      Left            =   6615
      TabIndex        =   25
      Top             =   360
      Width           =   1605
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
      Left            =   8490
      TabIndex        =   12
      Top             =   7530
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Frame Frame2 
      Height          =   2370
      Index           =   0
      Left            =   180
      TabIndex        =   9
      Top             =   855
      Width           =   9330
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
         Height          =   735
         Index           =   4
         Left            =   210
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Tag             =   "Observaciones|T|S|||advtrata|observac|||"
         Top             =   1500
         Width           =   8880
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
         Index           =   3
         Left            =   5100
         MaxLength       =   10
         TabIndex        =   3
         Tag             =   "Fecha Fin|F|S|||advtrata|fechafin|dd/mm/yyyy||"
         Text            =   "1234567890"
         Top             =   840
         Width           =   1335
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
         Index           =   2
         Left            =   2070
         MaxLength       =   10
         TabIndex        =   2
         Tag             =   "Fecha Inicio|F|S|||advtrata|fechaini|dd/mm/yyyy||"
         Text            =   "1234567890"
         Top             =   840
         Width           =   1335
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
         Index           =   0
         Left            =   225
         MaxLength       =   4
         TabIndex        =   0
         Tag             =   "Código|T|N|||advtrata|codtrata||S|"
         Top             =   450
         Width           =   1095
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
         Left            =   1755
         MaxLength       =   30
         TabIndex        =   1
         Tag             =   "Descripción|T|N|||advtrata|nomtrata|||"
         Top             =   450
         Width           =   7350
      End
      Begin VB.Image imgZoom 
         Height          =   240
         Index           =   0
         Left            =   1785
         Tag             =   "-1"
         ToolTipText     =   "Zoom descripción"
         Top             =   1200
         Width           =   240
      End
      Begin VB.Label Label8 
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
         Left            =   210
         TabIndex        =   15
         Top             =   1200
         Width           =   1515
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   1
         Left            =   4800
         Picture         =   "frmADVTratamientos.frx":000C
         ToolTipText     =   "Buscar fecha"
         Top             =   840
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   0
         Left            =   1770
         Picture         =   "frmADVTratamientos.frx":0097
         ToolTipText     =   "Buscar fecha"
         Top             =   870
         Width           =   240
      End
      Begin VB.Label Label22 
         Caption         =   "Fecha Fin"
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
         Left            =   3690
         TabIndex        =   14
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label18 
         Caption         =   "Fecha Inicio"
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
         Left            =   210
         TabIndex        =   13
         Top             =   870
         Width           =   1290
      End
      Begin VB.Label Label4 
         Caption         =   "Código"
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
         Left            =   270
         TabIndex        =   11
         Top             =   225
         Width           =   690
      End
      Begin VB.Label Label1 
         Caption         =   "Nombre"
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
         Height          =   255
         Index           =   0
         Left            =   1755
         TabIndex        =   10
         Top             =   210
         Width           =   1005
      End
   End
   Begin VB.Frame FrameAux1 
      BorderStyle     =   0  'None
      Height          =   4110
      Left            =   225
      TabIndex        =   16
      Top             =   3255
      Width           =   9290
      Begin VB.Frame FrameToolAux 
         Height          =   645
         Left            =   45
         TabIndex        =   31
         Top             =   45
         Width           =   1455
         Begin MSComctlLib.Toolbar ToolAux 
            Height          =   330
            Index           =   1
            Left            =   135
            TabIndex        =   32
            Top             =   180
            Width           =   1155
            _ExtentX        =   2037
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
      Begin VB.TextBox txtaux1 
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
         Height          =   310
         Index           =   4
         Left            =   7335
         MaxLength       =   10
         TabIndex        =   23
         Tag             =   "Cantidad|N|S|||advtrata_lineas|cantidad|##,##0.000||"
         Text            =   "Cantida"
         Top             =   2925
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.TextBox txtaux1 
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
         Height          =   310
         Index           =   3
         Left            =   6615
         MaxLength       =   12
         TabIndex        =   22
         Tag             =   "Dosis Habitual|N|S|||advtrata_lineas|dosishab|###,##0.000||"
         Text            =   "dosis H"
         Top             =   2925
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.TextBox txtaux1 
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
         Height          =   310
         Index           =   2
         Left            =   945
         MaxLength       =   3
         TabIndex        =   20
         Tag             =   "Linea Tratamiento|N|N|0|999|advtrata_lineas|numlinea|000|S|"
         Text            =   "lin"
         Top             =   2925
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.TextBox txtaux1 
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
         Height          =   310
         Index           =   0
         Left            =   225
         MaxLength       =   4
         TabIndex        =   19
         Tag             =   "Código Tratamiento|T|N|||advtrata_lineas|codtrata||S|"
         Text            =   "cod"
         Top             =   2925
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.TextBox txtaux1 
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
         Height          =   310
         Index           =   1
         Left            =   1665
         MaxLength       =   16
         TabIndex        =   21
         Tag             =   "Articulo|T|N|||advtrata_lineas|codartic||N|"
         Text            =   "articulo"
         Top             =   2925
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CommandButton cmdAux 
         Appearance      =   0  'Flat
         Caption         =   "+"
         Height          =   310
         Index           =   1
         Left            =   3030
         TabIndex        =   18
         ToolTipText     =   "Buscar artículo ADV"
         Top             =   2910
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
         Height          =   310
         Index           =   1
         Left            =   3270
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   17
         Text            =   "Nombre articulo"
         Top             =   2910
         Visible         =   0   'False
         Width           =   3285
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
         Bindings        =   "frmADVTratamientos.frx":0122
         Height          =   3195
         Index           =   1
         Left            =   45
         TabIndex        =   24
         Top             =   720
         Width           =   9160
         _ExtentX        =   16166
         _ExtentY        =   5636
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
      Left            =   195
      TabIndex        =   7
      Top             =   7440
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
         TabIndex        =   8
         Top             =   180
         Width           =   2655
      End
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
      Left            =   7290
      TabIndex        =   5
      Top             =   7530
      Width           =   1035
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   3690
      Top             =   7215
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
      Left            =   9060
      TabIndex        =   30
      Top             =   300
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
Attribute VB_Name = "frmADVTratamientos"
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

' *** declarar els formularis als que vaig a cridar ***
Private WithEvents frmB As frmBuscaGrid
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmC As frmCal 'calendario fecha
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmC1 As frmCal 'calendario fecha
Attribute frmC1.VB_VarHelpID = -1
Private WithEvents frmZ As frmZoom  'Zoom para campos Text
Attribute frmZ.VB_VarHelpID = -1

Private WithEvents frmVar As frmComVar 'variedades
Attribute frmVar.VB_VarHelpID = -1
Private WithEvents frmArtic As frmADVArticulos 'articulos
Attribute frmArtic.VB_VarHelpID = -1
Private WithEvents frmTtoPrev As frmADVTratamientosPrev ' tratamientos vista previa
Attribute frmTtoPrev.VB_VarHelpID = -1

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
Dim Indice As Byte 'Index del text1 on es poses els datos retornats des d'atres Formularis de Mtos
Dim CadB As String

Dim vSeccion As CSeccion

Dim B As Boolean

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
            If DatosOK Then
                If InsertarDesdeForm2(Me, 1) Then
                    ' *** canviar o llevar el WHERE, repasar codEmpre ****
                    Data1.RecordSource = "Select * from " & NombreTabla & Ordenacion
                    'Data1.RecordSource = "Select * from " & NombreTabla & " where codempre = " & codEmpre & Ordenacion
                    ' ***************************************************************
                    PosicionarData
                    PonerCampos
                    BotonAnyadirLinea 1
                End If
            Else
                ModoLineas = 0
            End If
            
        Case 4  'MODIFICAR
            If DatosOK Then
                If ModificaDesdeFormulario2(Me, 1) Then
                    TerminaBloquear
                    PosicionarData
                    CargaGrid 1, True
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
            Set frmArtic = New frmADVArticulos
            frmArtic.DatosADevolverBusqueda = "0|1|"
            frmArtic.CodigoActual = txtAux1(1).Text
            frmArtic.Show vbModal
            Set frmArtic = Nothing
            PonerFoco txtAux1(1)

    End Select
    If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
    If PrimeraVez Then PrimeraVez = False
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Form_Unload(Cancel As Integer)
    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
Dim I As Integer

    'Icono del formulario
    Me.Icon = frmPpal.Icon


    PrimeraVez = True
    
    ' ICONETS DE LA BARRA
    btnPrimero = 17 'index del botó "primero"
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
    
    
'    'cargar IMAGES de busqueda
'    For I = 0 To Me.imgBuscar.Count - 1
'        Me.imgBuscar(I).Picture = frmPpal.imgListImages16.ListImages(1).Picture
'    Next I
    
    'IMAGES para zoom
    For I = 0 To Me.imgZoom.Count - 1
        Me.imgZoom(I).Picture = frmPpal.imgListImages16.ListImages(3).Picture
    Next I
    
    NumTabMto = 1
'    ' *** si n'hi han tabs, per a que per defecte sempre es pose al 1r***
'    Me.SSTab1.Tab = 0
'    ' *******************************************************************
    
    LimpiarCampos   'Neteja els camps TextBox
'    ' ******* si n'hi han llínies *******
'    DataGridAux(0).ClearFields
    DataGridAux(1).ClearFields
    
    '*** canviar el nom de la taula i l'ordenació de la capçalera ***
    NombreTabla = "advtrata"
    Ordenacion = " ORDER BY codtrata"
    '************************************************
    
    'Mirem com està guardat el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    
    'ASignamos un SQL al DATA1
    Data1.ConnectionString = conn
    '***** canviar el nom de la PK de la capçalera; repasar codEmpre *************
    Data1.RecordSource = "Select * from " & NombreTabla & " where codtrata='-1'"
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
         
    
    If DatosADevolverBusqueda = "" Then
        PonerModo 0
    Else
        PonerModo 1 'búsqueda
        ' *** posar de groc els camps visibles de la clau primaria de la capçalera ***
        Text1(0).BackColor = vbLightBlue 'codclien
        ' ****************************************************************************
    End If
End Sub


Private Sub LimpiarCampos()
Dim I As Integer

    On Error Resume Next
    
    limpiar Me   'Mètode general: Neteja els controls TextBox
    lblIndicador.Caption = ""
'    Me.chkAbonos(0).Value = 0
    
    
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
Dim I As Integer, NumReg As Byte
Dim B As Boolean

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
       
    'Bloqueja els camps Text1 si no estem modificant/Insertant Datos
    'Si estem en Insertar a més neteja els camps Text1
    BloquearText1 Me, Modo
    BloquearCombo Me, Modo
'    BloquearChk Me.chkAbonos(0), (Modo = 0 Or Modo = 2 Or Modo = 5)
'    BloquearChk Me.chkAbonos(1), (Modo = 0 Or Modo = 2 Or Modo = 5)
    
    ' ***** bloquejar tots els controls visibles de la clau primaria de la capçalera ***
'    If Modo = 4 Then _
'        BloquearTxt Text1(0), True 'si estic en  modificar, bloqueja la clau primaria
    ' **********************************************************************************
    
    ' **** si n'hi han imagens de buscar en la capçalera *****
    BloquearImgBuscar Me, Modo, ModoLineas
    For I = 0 To imgFec.Count - 1
        BloquearImgFec Me, I, Modo
    Next I
'    BloquearImgFec Me, 25, Modo, ModoLineas
    BloquearImgZoom Me, Modo, ModoLineas
    ' ********************************************************
    ' *** si n'hi han llínies i imagens de buscar que no estiguen als grids ******
    'Llínies Departaments
    B = (Modo = 5 And (ModoLineas = 1 Or ModoLineas = 2) And (NumTabMto = 0))
'    BloquearImage imgBuscar(3), Not b
'    BloquearImage imgBuscar(4), Not b
'    BloquearImage imgBuscar(7), Not b
'    imgBuscar(3).Enabled = b
'    imgBuscar(3).visible = b
    ' ****************************************************************************
            
    chkVistaPrevia.Enabled = (Modo <= 2)
    
    PonerLongCampos

    If (Modo < 2) Or (Modo = 3) Then
'        CargaGrid 0, False
        CargaGrid 1, False
    End If
    
    B = (Modo = 4) Or (Modo = 2)
    

'    DataGridAux(0).Enabled = b
    DataGridAux(1).Enabled = B
'    ' ****** si n'hi han combos a la capçalera ***********************
'    If (Modo = 0) Or (Modo = 2) Or (Modo = 4) Or (Modo = 5) Then
'        Combo1(0).Enabled = False
'        Combo1(0).BackColor = &H80000018 'groc
'    ElseIf (Modo = 1) Or (Modo = 3) Then
'        Combo1(0).Enabled = True
'        Combo1(0).BackColor = &H80000005 'blanc
'    End If
'    ' ****************************************************************
    
    ' *** si n'hi han llínies i algún tab que no te datagrid ***
'    BloquearFrameAux Me, "FrameAux3", Modo, NumTabMto
'    BloquearFrameAux2 Me, "FrameAux3", (Modo <> 5) Or (Modo = 5 And indFrame <> 3) 'frame datos viaje indiv.
    ' ***************************
        
    'lineas de tratamiento
    B = (Modo = 5) And (NumTabMto = 1) 'And (ModoLineas <> 3)
    For I = 1 To txtAux1.Count - 1
        BloquearTxt txtAux1(I), Not B
    Next I
    B = (Modo = 5) And (NumTabMto = 1) And ModoLineas = 2
    BloquearTxt txtAux1(1), B
    BloquearTxt txtAux1(2), B Or ModoLineas = 1
    BloquearBtn cmdAux(1), B
    
    
     '-----------------------------
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
    Toolbar1.Buttons(1).Enabled = B
    Me.mnNuevo.Enabled = B
    
    B = (Modo = 2 And Data1.Recordset.RecordCount > 0) 'And Not DeConsulta
    'Modificar
    Toolbar1.Buttons(2).Enabled = B
    Me.mnModificar.Enabled = B
    'eliminar
    Toolbar1.Buttons(3).Enabled = B
    Me.mnEliminar.Enabled = B
    'Imprimir
    Toolbar1.Buttons(8).Enabled = B
       
    ' *** si n'hi han llínies que tenen grids (en o sense tab) ***
    B = (Modo = 3 Or Modo = 4 Or Modo = 2)
    For I = 1 To ToolAux.Count
        ToolAux(I).Buttons(1).Enabled = B
        If B Then bAux = (B And Me.Adoaux(I).Recordset.RecordCount > 0)
        ToolAux(I).Buttons(2).Enabled = bAux
        ToolAux(I).Buttons(3).Enabled = bAux
    Next I
    ' ****************************************
    
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
Dim Sql As String
Dim tabla As String
    
    ' ********* si n'hi han tabs, dona igual si en datagrid o no ***********
    Select Case Index
       Case 1 ' lineas de tratamiento
            tabla = "advtrata_lineas"
            Sql = "SELECT advtrata_lineas.codtrata, advtrata_lineas.numlinea, advtrata_lineas.codartic, advartic.nomartic, "
            Sql = Sql & " advtrata_lineas.dosishab, advtrata_lineas.cantidad "
            Sql = Sql & " FROM " & tabla & " INNER JOIN advartic ON advtrata_lineas.codartic = advartic.codartic "
            If enlaza Then
                Sql = Sql & ObtenerWhereCab(True)
            Else
                Sql = Sql & " WHERE advtrata_lineas.codtrata = '-1'"
            End If
            Sql = Sql & " ORDER BY " & tabla & ".codtrata "
            
            
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
Dim Indice As Byte
    Indice = CByte(Me.cmdAux(0).Tag + 2)
    txtAux1(Indice).Text = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub frmC1_Selec(vFecha As Date)
Dim Indice As Byte
    Indice = CByte(Me.imgFec(0).Tag)
    Text1(Indice).Text = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub frmArtic_DatoSeleccionado(CadenaSeleccion As String)
    txtAux1(1).Text = RecuperaValor(CadenaSeleccion, 1) 'codigo articulo
    txtAux2(1).Text = RecuperaValor(CadenaSeleccion, 2) 'nombre articulo
End Sub

Private Sub frmTtoPrev_DatoSeleccionado(CadenaSeleccion As String)
Dim CadB As String
    
    If CadenaSeleccion <> "" Then
        CadB = "codtrata = " & DBSet(RecuperaValor(CadenaSeleccion, 1), "N")
        
        'Se muestran en el mismo form
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " " & Ordenacion
        PonerCadenaBusqueda
        Screen.MousePointer = vbDefault
    End If
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
                Indice = Index + 2
       End Select
       
       Me.imgFec(0).Tag = Indice
       
       PonerFormatoFecha Text1(Indice)
       If Text1(Indice).Text <> "" Then frmC1.NovaData = CDate(Text1(Indice).Text)
    
       Screen.MousePointer = vbDefault
       frmC1.Show vbModal
       Set frmC1 = Nothing
       PonerFoco Text1(Indice)
    
End Sub


Private Sub imgZoom_Click(Index As Integer)
    
    Set frmZ = New frmZoom

    Select Case Index
        Case 0
            Indice = 4
            frmZ.pTitulo = "Observaciones del Tratamiento"
            frmZ.pValor = Text1(Indice).Text
            frmZ.pModo = Modo
            frmZ.Show vbModal
            Set frmZ = Nothing
            PonerFoco Text1(Indice)
    End Select
            
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
    ' *** repasar el nom de l'adodc, l'index del Field i el camp que te la PK ***
    'El registre de codi 0 no es pot Modificar ni Eliminar
'    If EsCodigoCero(CStr(Data1.Recordset.Fields(0).Value), FormatoCampo(Text1(0))) Then Exit Sub
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
        Case 1  'Nou
            mnNuevo_Click
        Case 2  'Modificar
            mnModificar_Click
        Case 3  'Borrar
            mnEliminar_Click
        Case 5  'Búscar
           mnBuscar_Click
        Case 6  'Tots
            mnVerTodos_Click
        Case 8 'Imprimir
'            AbrirListado (10)
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
        Text1(0).BackColor = vbLightBlue ' <===
        ' *** si n'hi han combos a la capçalera ***
    Else
        HacerBusqueda
        If Data1.Recordset.EOF Then
            Text1(kCampo).Text = ""
            Text1(kCampo).BackColor = vbLightBlue
            PonerFoco Text1(kCampo)
        End If
    End If
' ******************************************************************************
End Sub

Private Sub HacerBusqueda()

    CadB = ObtenerBusqueda2(Me, BuscaChekc, 1)
    
    If chkVistaPrevia = 1 Then
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
'    Dim cad As String
'    Dim NombreTabla1 As String
'
'    'Cridem al form
'    ' **************** arreglar-ho per a vore lo que es desije ****************
'    ' NOTA: el total d'amples de ParaGrid, ha de sumar 100
'    cad = ""
'    cad = cad & "Código|advtrata.codtrata|T||10·"
'    cad = cad & "Descripción|advtrata.nomtrata|T||60·"
'    cad = cad & "Fecha Inicio|advtrata.fechaini|F||15·"
'    cad = cad & "Fecha Fin|advtrata.fechafin|F||15·"
'
'
''    NombreTabla1 = "(rprecios inner join variedades on rprecios.codvarie = variedades.codvarie)"
'    NombreTabla1 = "advtrata"
'    If cad <> "" Then
'        Screen.MousePointer = vbHourglass
'        Set frmB = New frmBuscaGrid
'        frmB.vCampos = cad
'        frmB.vtabla = NombreTabla1
'        frmB.vSQL = CadB
'        HaDevueltoDatos = False
'        frmB.vDevuelve = "0|" '*** els camps que volen que torne ***
'        frmB.vTitulo = "Tratamientos" ' ***** repasa açò: títol de BuscaGrid *****
'        frmB.vSelElem = 0
'
'        frmB.Show vbModal
'        Set frmB = Nothing
'        'Si ha posat valors i tenim que es formulari de búsqueda llavors
'        'tindrem que tancar el form llançant l'event
'        If HaDevueltoDatos Then
'            If (Not Data1.Recordset.EOF) And DatosADevolverBusqueda <> "" Then _
'                cmdRegresar_Click
'        Else   'de ha retornat datos, es a decir NO ha retornat datos
'            PonerFoco Text1(kCampo)
'        End If
'    End If

    Set frmTtoPrev = New frmADVTratamientosPrev
    frmTtoPrev.cWhere = CadB
    frmTtoPrev.DatosADevolverBusqueda = "0|1|2|"
    frmTtoPrev.Show vbModal
    
    Set frmTtoPrev = Nothing

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
    
    PonerModo 0
    
    If chkVistaPrevia.Value = 1 Then
        MandaBusquedaPrevia ""
    Else
        ' *** canviar o llevar, si cal, el WHERE; repasar codEmpre ***
        CadenaConsulta = "Select * from " & NombreTabla & Ordenacion
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
    
    ' *** foco al 1r camp visible que NO siga clau primaria ***
    PonerFoco Text1(1)
    ' *********************************************************
End Sub


Private Sub BotonEliminar()
Dim cad As String

    On Error GoTo EEliminar

    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub

    ' *** repasar el nom de l'adodc, l'index del Field i el camp que te la PK ***
    'El registre de codi 0 no es pot Modificar ni Eliminar
    If EsCodigoCero(CStr(Data1.Recordset.Fields(0).Value), FormatoCampo(Text1(0))) Then Exit Sub
    ' ***************************************************************************

    ' *************** canviar la pregunta ****************
    cad = "¿Seguro que desea eliminar el Tratamiento?"
    cad = cad & vbCrLf & "Código: " & Format(Data1.Recordset.Fields(0), FormatoCampo(Text1(0)))
    cad = cad & vbCrLf & "Nombre: " & Data1.Recordset.Fields(1)
    ' **************************************************************************
    
    'borrem
    If MsgBox(cad, vbQuestion + vbYesNo) = vbYes Then
        On Error GoTo EEliminar
        Screen.MousePointer = vbHourglass
        NumRegElim = Data1.Recordset.AbsolutePosition
        If Not Eliminar Then
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
    If Err.Number > 0 Then MuestraError Err.Number, "Eliminar Tratamiento", Err.Description
End Sub


Private Sub PonerCampos()
Dim I As Integer
Dim CodPobla As String, desPobla As String
Dim CPostal As String, desProvi As String, desPais As String

    If Data1.Recordset.EOF Then Exit Sub
    PonerCamposForma2 Me, Data1, 1 'opcio=1: posa el format o els camps de la capçalera
    
    ' *** si n'hi han llínies en datagrids ***
    For I = 1 To DataGridAux.Count ' - 1
        If I <> 3 Then
            CargaGrid I, True
            If Not Adoaux(I).Recordset.EOF Then _
                PonerCamposForma2 Me, Adoaux(I), 2, "FrameAux" & I
        End If
    Next I
    ' *******************************************

    ' *** si n'hi han llínies sense datagrid ***
    ' ***************************************
    
    ' ************* configurar els camps de les descripcions de la capçalera *************
'    txtAux2(22).Text = PonerNombreDeCod(txtAux(22), "poblacio", "despobla", "codpobla", "N")

    PosarDescripcions


    ' ********************************************************************************
    
    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
    
    PonerModoOpcionesMenu (Modo)
    PonerOpcionesMenu   'Activar opcions de menú según nivell
                        'de permisos de l'usuari
    
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
                                For I = 0 To txtAux1.Count - 1
                                    txtAux1(I).Text = ""
                                    BloquearTxt txtAux1(I), True
                                Next I
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

                    If Not Adoaux(NumTabMto).Recordset.EOF Then
                        Adoaux(NumTabMto).Recordset.MoveFirst
                    End If

                Case 2 'modificar llínies
                    ModoLineas = 0

                    ' *** si n'hi han tabs ***
                    SituarTab (NumTabMto)
                    'SSTab1.Tab = 1
                    'SSTab2.Tab = NumTabMto
                    ' ***********************

                    PonerModo 4
                    If Not Adoaux(NumTabMto).Recordset.EOF Then
                        ' *** l'Index de Fields es el que canvie de la PK de llínies ***
                        V = Adoaux(NumTabMto).Recordset.Fields(1) 'el 2 es el nº de llinia
                        Adoaux(NumTabMto).Recordset.Find (Adoaux(NumTabMto).Recordset.Fields(1).Name & " =" & V)
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
            

            PosicionarData
            
            TerminaBloquear

            ' *** si n'hi han llínies en grids i camps fora d'estos ***
            If Not Adoaux(NumTabMto).Recordset.EOF Then
                DataGridAux_RowColChange NumTabMto, 1, 1
            Else
                LimpiarCamposFrame NumTabMto
            End If
            ' *********************************************************
    End Select
End Sub


Private Function DatosOK() As Boolean
Dim B As Boolean
Dim Sql As String
Dim cad As String
'Dim Datos As String

    On Error GoTo EDatosOK

    DatosOK = False
    B = CompForm2(Me, 1)
    If Not B Then Exit Function
    
    ' *** canviar els arguments de la funcio, el mensage i repasar si n'hi ha codEmpre ***
    If (Modo = 3) Then 'insertar
        'comprobar si existe ya el cod. del campo clave primaria
        Sql = DevuelveDesdeBDNew(cAgro, "advtrata", "codtrata", "codtrata", Text1(0).Text, "T")
        If Sql <> "" Then
            MsgBox "Ya existe el codigo de tratamiento. Revise.", vbExclamation
            B = False
        End If
    End If
    
' --monica: de momento quitamos que no se puedan solapar
'    'miramos si hay otros campos con la misma ubicacion
'    If b And (Modo = 3 Or Modo = 4) Then
'        b = ComprobacionRangoFechas(Text1(0).Text, CStr(Combo1(0).ListIndex), Text1(1).Text, Text1(2).Text, Text1(3).Text)
'
'        If b = False Then
'            MsgBox "El rango de fechas se solapa con otro registro del mismo tipo de esta variedad. Revise.", vbExclamation
'        End If
'    End If
    
    
    ' ************************************************************************************
    DatosOK = B
    
EDatosOK:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function


Private Sub PosicionarData()
Dim cad As String, Indicador As String

    ' *** canviar-ho per tota la PK de la capçalera, no llevar els () ***
    cad = "(codtrata='" & Trim(Text1(0).Text) & "') "
    ' ***************************************
    
    ' *** gastar SituarData o SituarDataMULTI depenent de si la PK es simple o composta ***
    'If SituarDataMULTI(Data1, cad, Indicador) Then
    If SituarDataMULTI(Data1, cad, Indicador) Then
        If ModoLineas <> 1 Then PonerModo 2
        lblIndicador.Caption = Indicador
    Else
       LimpiarCampos
       PonerModo 0
    End If
    ' ***********************************************************************************
End Sub


Private Function Eliminar() As Boolean
Dim vWhere As String

    On Error GoTo FinEliminar

    conn.BeginTrans
    ' ***** canviar el nom de la PK de la capçalera, repasar codEmpre *******
    vWhere = " WHERE codtrata=" & DBSet(Data1.Recordset!codtrata, "T")
        ' ***********************************************************************
        
    ' ***** elimina les llínies ****
    conn.Execute "DELETE FROM advtrata_lineas " & vWhere

'    ' *******************************
'    'Eliminar la CAPÇALERA
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
    Select Case Index
        Case 0 'codigo
            If Text1(Index).Text <> "" Then Text1(Index).Text = UCase(Text1(Index).Text)
            
        Case 2, 3 ' fechas de inicio y fin
            PonerFormatoFecha Text1(Index), True
            If Text1(2).Text <> "" And Text1(3).Text <> "" Then
                If CDate(Text1(2).Text) > CDate(Text1(3).Text) Then
                    MsgBox "La Fecha Inicio debe ser inferior a la Fecha Fin. Revise", vbExclamation
                    Text1(Index).Text = ""
                    PonerFoco Text1(Index)
                End If
            End If
        
                

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
        If Index <> 4 Then KEYpress KeyAscii
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

' **** si n'hi han camps de descripció a la capçalera ****
Private Sub PosarDescripcions()
Dim NomEmple As String

    On Error GoTo EPosarDescripcions


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
        Case 1 'linea de tratamiento
            Sql = "¿Seguro que desea eliminar la línea de tratamiento?"
            Sql = Sql & vbCrLf & "Código: " & Adoaux(Index).Recordset!codtrata & " - " & Adoaux(Index).Recordset!numlinea
            If MsgBox(Sql, vbQuestion + vbYesNo) = vbYes Then
                Eliminar = True
                Sql = "DELETE FROM advtrata_lineas "
                Sql = Sql & vWhere & " and numlinea = " & DBLet(Adoaux(Index).Recordset!numlinea, "N")
            End If
        
    End Select

    If Eliminar Then
        NumRegElim = Adoaux(Index).Recordset.AbsolutePosition
        TerminaBloquear
        conn.Execute Sql
        ' *** si n'hi han tabs sense datagrid, posar l'If ***
        If Index <> 3 Then _
            CargaGrid Index, True
        ' ***************************************************
        If Not SituarDataTrasEliminar(Adoaux(Index), NumRegElim, True) Then
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
    ' **************************************************

    ' *** posar el nom del les distintes taules de llínies ***
    Select Case Index
        Case 1: vtabla = "advtrata_lineas"
    End Select
    ' ********************************************************

    vWhere = ObtenerWhereCab(False)

    Select Case Index
         Case 1   'clasificacion
            ' *** canviar la clau primaria de les llínies,
            'pasar a "" si no volem que mos sugerixca res a l'afegir ***
            NumF = ""
            NumF = SugerirCodigoSiguienteStr(vtabla, "numlinea", vWhere)
            ' ***************************************************************

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
                Case 1 'lineas de tratamiento
                    For I = 0 To txtAux1.Count - 1
                        txtAux1(I).Text = ""
                    Next I
                    txtAux1(0).Text = Text1(0).Text 'codigo tratamiento
                    txtAux1(2).Text = Format(NumF, "000") 'linea contador
                    
                    txtAux1(1).Text = "" 'articulo
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
    ' *********************************

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
        Case 1 'articulos
            txtAux1(0).Text = DataGridAux(Index).Columns(0).Text
            txtAux1(1).Text = DataGridAux(Index).Columns(2).Text
            txtAux1(2).Text = DataGridAux(Index).Columns(1).Text
            
            txtAux2(1).Text = DataGridAux(Index).Columns(3).Text ' nombre articulo
            txtAux1(3).Text = DataGridAux(Index).Columns(4).Text 'dosis habitual
            txtAux1(4).Text = DataGridAux(Index).Columns(5).Text 'cantidad
            
    End Select

    LLamaLineas Index, ModoLineas, anc
    
    BloquearCantidad
    
    ' *** foco al 1r camp visible de les llinies en grids que no siga PK (en o sense tab) ***
    If txtAux1(3).Enabled = False Then
        txtAux1(3).Text = ""
        PonerFoco txtAux1(4)
    Else
        txtAux1(4).Text = ""
        PonerFoco txtAux1(3)
    End If
    
    ' ***************************************************************************************
End Sub


Private Sub LLamaLineas(Index As Integer, xModo As Byte, Optional alto As Single)
Dim jj As Integer
Dim B As Boolean

    ' *** si n'hi han tabs sense datagrid posar el If ***
    If Index <> 3 Then DeseleccionaGrid DataGridAux(Index)
    ' ***************************************************

    B = (xModo = 1 Or xModo = 2) 'Insertar o Modificar Llínies
    Select Case Index
        Case 1 'articulos de adv
            For jj = 1 To txtAux1.Count - 1
                txtAux1(jj).visible = B
                txtAux1(jj).Top = alto
            Next jj
            
            txtAux2(1).visible = B
            txtAux2(1).Top = alto

            For jj = 1 To cmdAux.Count
                cmdAux(jj).visible = B
                cmdAux(jj).Top = txtAux1(3).Top
                cmdAux(jj).Height = txtAux1(3).Height
            Next jj
    End Select
End Sub



Private Function DatosOkLlin(nomframe As String) As Boolean
Dim Rs As ADODB.Recordset
Dim Sql As String
Dim B As Boolean
Dim Cant As Integer
Dim Mens As String
Dim vFact As Byte, vDocum As Byte
Dim vArtADV As CArticuloADV


    DatosOkLlin = True
    
    On Error GoTo EDatosOKLlin

    Mens = ""
    DatosOkLlin = False

    B = CompForm2(Me, 2, nomframe) 'Comprovar formato datos ok
    If Not B Then Exit Function
    
    If B And (Modo = 5 And ModoLineas = 1) Then  'insertar
        'comprobar si existe ya el cod. de la calidad para ese campo
        Sql = ""
        Sql = DevuelveDesdeBDNew(cAgro, "advtrata_lineas", "codartic", "codtrata", txtAux1(0).Text, "T", , "codartic", txtAux1(1).Text, "T")
        If Sql <> "" Then
            MsgBox "Ya existe el artículo en el Tratamiento. Revise.", vbExclamation
            PonerFoco txtAux1(1)
            B = False
        End If
    End If
    
    If B And Modo = 5 Then ' tanto si insertamos como si modificamos en lineas
        Set vArtADV = New CArticuloADV
        If vArtADV.LeerDatos(txtAux1(1).Text) Then
            If vArtADV.TipoProd = 0 Then
                If txtAux1(3).Text = "" Then
                    MsgBox "Los artículos de tipo producto deben de llevar dosis.", vbExclamation
                    PonerFoco txtAux1(3)
                    B = False
                End If
            Else
                If txtAux1(4).Text = "" Then
                    MsgBox "Los artículos de tipo Trabajo o Varios deben de llevar cantidad. Revise.", vbExclamation
                    PonerFoco txtAux1(4)
                    B = False
                End If
            End If
        End If
        Set vArtADV = Nothing
    End If
    
    DatosOkLlin = B

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
Dim B As Boolean
Dim I As Byte
Dim tots As String

    On Error GoTo ECarga

    tots = MontaSQLCarga(Index, enlaza)

    B = DataGridAux(Index).Enabled
    DataGridAux(Index).Enabled = False
    
    Adoaux(Index).ConnectionString = conn
    Adoaux(Index).RecordSource = MontaSQLCarga(Index, enlaza)
    Adoaux(Index).CursorType = adOpenDynamic
    Adoaux(Index).LockType = adLockPessimistic
    DataGridAux(Index).ScrollBars = dbgNone
    Adoaux(Index).Refresh
    Set DataGridAux(Index).DataSource = Adoaux(Index)
    
    DataGridAux(Index).AllowRowSizing = False
    DataGridAux(Index).RowHeight = 290
    If PrimeraVez Then
        DataGridAux(Index).ClearFields
        DataGridAux(Index).ReBind
        DataGridAux(Index).Refresh
    End If

    For I = 0 To DataGridAux(Index).Columns.Count - 1
        DataGridAux(Index).Columns(I).AllowSizing = False
    Next I
    
    CargaGridGnral Me.DataGridAux(Index), Me.Adoaux(Index), tots, PrimeraVez
    
    
    'DataGridAux(Index).Enabled = b
'    PrimeraVez = False
    
    Select Case Index
        Case 1 'lineas de tratamiento
            'si es visible|control|tipo campo|nombre campo|ancho control|
            tots = "N||||0|;S|txtaux1(2)|T|Lin|600|;S|txtaux1(1)|T|Código.|2100|;S|cmdAux(1)|B|||;" 'codsocio,codsecci
            tots = tots & "S|txtAux2(1)|T|Artículo|3750|;"
            tots = tots & "S|txtaux1(3)|T|Dosis Hab.|1100|;"
            tots = tots & "S|txtaux1(4)|T|Cantidad|1100|;"
            
            arregla tots, DataGridAux(Index), Me, 350
        
            DataGridAux(Index).Columns(3).Alignment = dbgLeft
            DataGridAux(Index).Columns(4).Alignment = dbgRight
            DataGridAux(Index).Columns(5).Alignment = dbgRight
            
            
            B = (Modo = 4) And ((ModoLineas = 1) Or (ModoLineas = 2))
'            BloquearTxt txtAux(14), Not b
'            BloquearTxt txtAux(15), Not b

            If (enlaza = True) And (Not Adoaux(Index).Recordset.EOF) Then 'per a que pose els valors de les arees de text la primera volta
'                txtAux2(4).Text = PonerNombreCuenta(txtaux1(4), Modo)
'                txtAux2(5).Text = PonerNombreCuenta(txtaux1(5), Modo)
'                txtAux2(0).Text = PonerNombreDeCod(txtaux1(6), "tiposiva", "nombriva", "codigiva", "N", cConta)
'                If VisualizaClasificacion Then
'                    PonerClasificacionGrafica
'
''                    SumaTotalPorcentajes
'                End If
            Else
                For I = 0 To 4
                    txtAux1(I).Text = ""
                Next I
                txtAux2(1).Text = ""
'                Me.MSChart1.visible = False
            End If
    End Select
    DataGridAux(Index).ScrollBars = dbgAutomatic
      
    ' **** si n'hi han llínies en grids i camps fora d'estos ****
    If Not Adoaux(Index).Recordset.EOF Then
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
Dim B As Boolean

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
            B = BLOQUEADesdeFormulario2(Me, Data1, 1)
            
            '++monica: en caso de estar insertando seccion y que no existan las
            'cuentas contables hacemos esto para que las inserte en contabilidad.
'            If NumTabMto = 1 Then
'               txtAux2(4).Text = PonerNombreCuenta(txtaux1(4), 3, Text1(0))
'               txtAux2(5).Text = PonerNombreCuenta(txtaux1(5), 3, Text1(0))
'            End If
            
            Select Case NumTabMto
                Case 0, 1 ' *** els index de les llinies en grid (en o sense tab) ***
                    CargaGrid NumTabMto, True
                    If B Then BotonAnyadirLinea NumTabMto
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
Dim cad As String
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
                V = Adoaux(NumTabMto).Recordset.Fields(1) 'el 2 es el nº de llinia
                CargaGrid NumTabMto, True
            End If

            ' *** si n'hi han tabs ***
            SituarTab (NumTabMto)

            ' *** si n'hi han tabs que no tenen datagrid, posar el if ***
            If NumTabMto <> 3 Then
                DataGridAux(NumTabMto).SetFocus
                Adoaux(NumTabMto).Recordset.Find (Adoaux(NumTabMto).Recordset.Fields(1).Name & " =" & V)
            End If
            ' ***********************************************************

            LLamaLineas NumTabMto, 0
            
        End If
    End If
        
End Sub




Private Function ObtenerWhereCab(conW As Boolean) As String
Dim vWhere As String
    
    vWhere = ""
    If conW Then vWhere = " WHERE "
    ' *** canviar-ho per la clau primaria de la capçalera ***
    vWhere = vWhere & " advtrata_lineas.codtrata=" & DBSet(Text1(0).Text, "T")
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

Private Sub printNou()
    With frmImprimir2
        .cadTabla2 = "advtrata"
        .Informe2 = "rADVTratamientos.rpt"
        If CadB <> "" Then
            .cadRegSelec = SQL2SF(CadB)
        Else
            .cadRegSelec = ""
        End If
        .cadRegActua = POS2SF(Data1, Me)
        '[Monica]13/07/2012: falla si hay un solo registro seleccionado y apretamos registros buscados
        If Data1.Recordset.RecordCount = 1 Then .cadRegSelec = .cadRegActua
        .cadTodosReg = ""
        '.OtrosParametros2 = "pEmpresa='" & vEmpresa.NomEmpre & "'|pOrden={clientes.ape_raso}|"
        .OtrosParametros2 = "pEmpresa='" & vEmpresa.nomempre & "'|"
        .NumeroParametros2 = 1
        .MostrarTree2 = False
        .InfConta2 = False
        .ConSubInforme2 = False
        .SubInformeConta = ""
        .Show vbModal
    End With
End Sub

Private Sub ToolbarDes_ButtonClick(ByVal Button As MSComctlLib.Button)
    Desplazamiento (Button.Index)
End Sub

'' ### [DavidV] 26/04/2006: Activar/desactivar la rueda del ratón.
'Private Sub DataGridAux_GotFocus(Index As Integer)
'  WheelHook DataGridAux(Index)
'End Sub
'Private Sub DataGridAux_LostFocus(Index As Integer)
'  WheelUnHook
'End Sub

Private Sub txtAux1_LostFocus(Index As Integer)
Dim cadMen As String
Dim Nuevo As Boolean

    If Not PerderFocoGnral(txtAux1(Index), Modo) Then Exit Sub

    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub


    ' ******* configurar el LostFocus dels camps de llínies (dins i fora grid) ********
    Select Case Index
        Case 1 ' articulo adv
            If txtAux1(Index).Text <> "" Then
                txtAux2(Index).Text = PonerNombreDeCod(txtAux1(Index), "advartic", "nomartic", "codartic", "T")
                If txtAux2(Index).Text = "" Then
                    cadMen = "No existe el Artículo de ADV: " & txtAux1(Index).Text & vbCrLf
                    cadMen = cadMen & "¿Desea crearlo?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmArtic = New frmADVArticulos
                        frmArtic.DatosADevolverBusqueda = "0|1|"
                        frmArtic.NuevoCodigo = txtAux1(Index).Text
                        txtAux1(Index).Text = ""
                        TerminaBloquear
                        frmArtic.Show vbModal
                        Set frmArtic = Nothing
                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
                    Else
                        txtAux1(Index).Text = ""
                    End If
                    PonerFoco txtAux1(Index)
                End If
            Else
                txtAux2(Index).Text = ""
            End If
            
            BloquearCantidad

        Case 3 'dosis habitual
            '[Monica]22/07/2013: antes era ponerformato decimal 12.
            If PonerFormatoDecimal(txtAux1(Index), 5) Then
                If txtAux1(4).Enabled = False Then cmdAceptar.SetFocus
            End If
            
            
        Case 4 'cantidad
            If PonerFormatoDecimal(txtAux1(Index), 2) Then cmdAceptar.SetFocus

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
                    Case 1: KEYBusqueda KeyAscii, 1 'articulo
                End Select
            End If
        Else
            KEYpress KeyAscii
        End If
    End If
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


Private Sub BloquearCantidad()
Dim vArticADV As CArticuloADV

    If txtAux1(1).Text <> "" Then
        Set vArticADV = New CArticuloADV
        If vArticADV.LeerDatos(txtAux1(1).Text) Then
            BloquearTxt txtAux1(4), (vArticADV.TipoProd = 0)
            txtAux1(4).Enabled = Not (vArticADV.TipoProd = 0)
        
            BloquearTxt txtAux1(3), (vArticADV.TipoProd <> 0)
            txtAux1(3).Enabled = Not (vArticADV.TipoProd <> 0)
        End If
        Set vArticADV = Nothing
    End If

End Sub
