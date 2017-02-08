VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmManOrdRecogida 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ordenes de Recogida"
   ClientHeight    =   6675
   ClientLeft      =   45
   ClientTop       =   30
   ClientWidth     =   9780
   Icon            =   "frmManOrdRecogida.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6675
   ScaleWidth      =   9780
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Left            =   7080
      TabIndex        =   32
      Top             =   210
      Width           =   1605
   End
   Begin VB.Frame FrameDesplazamiento 
      Height          =   705
      Left            =   3780
      TabIndex        =   30
      Top             =   90
      Width           =   2415
      Begin MSComctlLib.Toolbar ToolbarDes 
         Height          =   330
         Left            =   240
         TabIndex        =   31
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
   Begin VB.Frame FrameBotonGnral 
      Height          =   705
      Left            =   120
      TabIndex        =   27
      Top             =   90
      Width           =   3585
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   240
         TabIndex        =   28
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
   Begin VB.Frame Frame2 
      Height          =   1815
      Index           =   0
      Left            =   135
      TabIndex        =   8
      Top             =   825
      Width           =   9450
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
         Left            =   1350
         MaxLength       =   6
         TabIndex        =   3
         Tag             =   "Variedad|N|N|1|999999|rordrecogida|codvarie|000000||"
         Top             =   750
         Width           =   1230
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
         Left            =   2580
         MaxLength       =   30
         TabIndex        =   25
         Top             =   750
         Width           =   6405
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Capazo Desinfectado"
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
         Index           =   3
         Left            =   6600
         TabIndex        =   24
         Tag             =   "Capazo Desinfectado|N|N|0|1|rordrecogida|capazodesinf|0||"
         Top             =   1290
         Width           =   2625
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Capazo Limpio"
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
         Index           =   2
         Left            =   4680
         TabIndex        =   23
         Tag             =   "Capazo Limpio|N|N|0|1|rordrecogida|capazolimp|0||"
         Top             =   1290
         Width           =   2055
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Tijeras Desinfectadas"
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
         Index           =   1
         Left            =   2070
         TabIndex        =   22
         Tag             =   "Tijeras Desinfectadas|N|N|0|1|rordrecogida|tijeradesinf|0||"
         Top             =   1290
         Width           =   2505
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Tijeras Limpias"
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
         Left            =   150
         TabIndex        =   21
         Tag             =   "Tijeras Limpias|N|N|0|1|rordrecogida|tijeralimp|0||"
         Top             =   1290
         Width           =   1815
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
         Left            =   4470
         MaxLength       =   10
         TabIndex        =   1
         Tag             =   "Fecha Impresión|F|N|||rordrecogida|fecimpre|dd/mm/yyyy||"
         Text            =   "1234567890"
         Top             =   270
         Width           =   1320
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
         Left            =   7860
         MaxLength       =   10
         TabIndex        =   2
         Tag             =   "Nro.Campo|N|N|||rordrecogida|nrocampo|00000000||"
         Top             =   270
         Width           =   1110
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
         Left            =   1320
         MaxLength       =   9
         TabIndex        =   0
         Tag             =   "Nro.Nota|N|S|||rordrecogida|nroorden|000000000|S|"
         Text            =   "123456789"
         Top             =   270
         Width           =   1230
      End
      Begin VB.Label Label5 
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
         Left            =   150
         TabIndex        =   26
         Top             =   750
         Width           =   855
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   2
         Left            =   1050
         ToolTipText     =   "Buscar Variedad"
         Top             =   750
         Width           =   240
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
         Left            =   3480
         TabIndex        =   20
         Top             =   270
         Width           =   690
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   0
         Left            =   4185
         Picture         =   "frmManOrdRecogida.frx":000C
         ToolTipText     =   "Buscar fecha"
         Top             =   270
         Width           =   240
      End
      Begin VB.Label Label26 
         Caption         =   "NroCampo"
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
         Left            =   6690
         TabIndex        =   19
         Top             =   270
         Width           =   990
      End
      Begin VB.Label Label1 
         Caption         =   "Nro.Orden"
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
         Left            =   150
         TabIndex        =   9
         Top             =   270
         Width           =   990
      End
   End
   Begin VB.Frame FrameAux1 
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
      Height          =   3360
      Left            =   120
      TabIndex        =   11
      Top             =   2670
      Width           =   9460
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
         Left            =   6540
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Tag             =   "Nivel|N|N|0|2|rordrecogida_incid|nivel||N|"
         Top             =   2610
         Width           =   1380
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
         Left            =   1755
         MaskColor       =   &H00000000&
         TabIndex        =   18
         ToolTipText     =   "Buscar Incidencia"
         Top             =   2610
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
         Index           =   9
         Left            =   1980
         TabIndex        =   17
         Text            =   "nombre"
         Top             =   2610
         Visible         =   0   'False
         Width           =   4410
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
         Index           =   9
         Left            =   1080
         MaxLength       =   4
         TabIndex        =   13
         Tag             =   "Incidencia|N|N|||rordrecogida_incid|idplaga|0000|S|"
         Text            =   "inci"
         Top             =   2610
         Visible         =   0   'False
         Width           =   600
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
         Index           =   8
         Left            =   360
         MaxLength       =   16
         TabIndex        =   12
         Tag             =   "Nro.Orden|N|N|||rordrecogida_incid|nroorden|0000000|S|"
         Text            =   "codfor"
         Top             =   2610
         Visible         =   0   'False
         Width           =   555
      End
      Begin MSComctlLib.Toolbar ToolAux 
         Height          =   390
         Index           =   1
         Left            =   90
         TabIndex        =   15
         Top             =   270
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
      Begin MSDataGridLib.DataGrid DataGridAux 
         Bindings        =   "frmManOrdRecogida.frx":0097
         Height          =   2610
         Index           =   1
         Left            =   90
         TabIndex        =   16
         Top             =   630
         Width           =   9070
         _ExtentX        =   16007
         _ExtentY        =   4604
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
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Index           =   0
      Left            =   150
      TabIndex        =   6
      Top             =   6060
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
         TabIndex        =   7
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
      Left            =   8535
      TabIndex        =   5
      Top             =   6090
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
      Left            =   7440
      TabIndex        =   4
      Top             =   6075
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
      Left            =   8550
      TabIndex        =   10
      Top             =   6090
      Visible         =   0   'False
      Width           =   1035
   End
   Begin MSAdodcLib.Adodc Data1 
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
      Left            =   9090
      TabIndex        =   29
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
Attribute VB_Name = "frmManOrdRecogida"
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

Private Const IdPrograma = 4002
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
'Private WithEvents frmExp As frmExpImpExcel ' Exportacion o importacion a pagina excel

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
Private WithEvents frmInc As frmBasico 'incidencias
Attribute frmInc.VB_VarHelpID = -1
Private WithEvents frmCal As frmManCalidades 'calidades
Attribute frmCal.VB_VarHelpID = -1
Private WithEvents frmCam As frmManCampos 'campos
Attribute frmCam.VB_VarHelpID = -1

Private WithEvents frmOrdPrev As frmManOrdRecogidaPrev ' previsualizacion de ordenes de recogida
Attribute frmOrdPrev.VB_VarHelpID = -1


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

Dim CodTipoMov As String
Dim NotaExistente As Boolean

Dim btnPrimero As Byte 'Variable que indica el nº del Botó PrimerRegistro en la Toolbar1
'Dim CadAncho() As Boolean  'array, per a quan cridem al form de llínies
Dim Indice As Byte 'Index del text1 on es poses els datos retornats des d'atres Formularis de Mtos
Dim CadB As String

Dim VarieAnt As String
Dim MuestraAnt As String


'GENERALES PARA PASARLE A CRYSTAL REPORT
Private cadFormula As String 'Cadena con la FormulaSelection para Crystal Report
Private CadParam As String 'Cadena con los parametros para Crystal Report
Private numParam As Byte 'Numero de parametros que se pasan a Crystal Report
Private cadSelect As String 'Cadena para comprobar si hay datos antes de abrir Informe
Private cadTitulo As String 'Titulo para la ventana frmImprimir
Private cadNombreRPT As String 'Nombre del informe
Private cadSelect1 As String 'Cadena para comprobar si hay datos antes de abrir Informe


Dim ConSubInforme As Boolean

Private BuscaChekc As String


Private Sub btnBuscar_Click(Index As Integer)
    TerminaBloquear
    Select Case Index
'        Case 0 'Variedades
'            Set frmVar = New frmComVar
'            frmVar.DatosADevolverBusqueda = "0|1|"
'            frmVar.CodigoActual = txtAux(1).Text
'            frmVar.Show vbModal
'            Set frmVar = Nothing
'            PonerFoco txtAux(1)
        Case 1 'Incidencia
            Indice = 9
            
            Set frmInc = New frmBasico
            
            AyudaIncidenciasOrdenesRecogida frmInc, txtAux(Indice)
            
            Set frmInc = Nothing
            PonerFoco txtAux(9)
    End Select
    If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Me.Adoaux, 1
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
    Screen.MousePointer = vbHourglass
    
    On Error GoTo Error1
    
    Select Case Modo
        Case 1  'BÚSQUEDA
            HacerBusqueda
            
        Case 3 'INSERTAR
            If DatosOK Then
                NotaExistente = False
                InsertarCabecera
            Else
                ModoLineas = 0
            End If
            
        Case 4  'MODIFICAR
            If DatosOK Then
'                If Gastos Then CalcularGastos
                If ModificaCabecera Then
                    espera 0.2
                    TerminaBloquear
                    PosicionarData
                    PonerCampos
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
                        PosicionarData
                        PonerCampos
                    End If
            End Select
            'nuevo calculamos los totales de lineas

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

Private Sub Combo1_LostFocus(Index As Integer)

    If Index = 0 Then CmdAceptar.SetFocus

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
    btnPrimero = 16 'index del botó "primero"
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
    For I = 1 To 1
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
    Me.imgBuscar(2).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    
    CodTipoMov = "ORR"
    
    ' *** si n'hi han combos (capçalera o llínies) ***
    CargaCombo
    
    LimpiarCampos   'Neteja els camps TextBox
    ' ******* si n'hi han llínies *******
    DataGridAux(1).ClearFields
    
    '*** canviar el nom de la taula i l'ordenació de la capçalera ***
    NombreTabla = "rordrecogida"
    Ordenacion = " ORDER BY nroorden"
    
    'Mirem com està guardat el valor del check
    chkVistaPrevia(0).Value = CheckValueLeer(Name)
    
    Data1.ConnectionString = conn
    '***** cambiar el nombre de la PK de la cabecera *************
    Data1.RecordSource = "Select * from " & NombreTabla & " where nroorden=-1"
    Data1.Refresh
       
    CargaGrid 1, False
       
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
Dim I As Integer

    On Error Resume Next
    
    limpiar Me   'Mètode general: Neteja els controls TextBox
    lblIndicador.Caption = ""
    
    For I = 0 To 3
        Me.Check1(I).Value = 0
    Next I
    
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
    
    '=======================================
    B = (Modo = 2)
    'Posar Fleches de desplasament visibles
    NumReg = 1
    If Not Data1.Recordset.EOF Then
        If Data1.Recordset.RecordCount > 1 Then NumReg = 2 'Només es per a saber que n'hi ha + d'1 registre
    End If
'    DesplazamientoVisible Me.Toolbar1, btnPrimero, b, NumReg
    DesplazamientoVisible B And Data1.Recordset.RecordCount > 1
    '---------------------------------------------
    
    
    B = Modo <> 0 And Modo <> 2
    cmdCancelar.visible = B
    CmdAceptar.visible = B
       
    'Bloqueja els camps Text1 si no estem modificant/Insertant Datos
    'Si estem en Insertar a més neteja els camps Text1
    BloquearText1 Me, Modo
    
    B = (Modo <> 1) And (Modo <> 3)
    'Campos Nº entrada bloqueado y en azul
    BloquearTxt Text1(0), B, True
    
    
    '*** si n'hi han combos a la capçalera ***
'    BloquearCombo Me, Modo
    '**************************
    
    ' **** si n'hi han imagens de buscar en la capçalera *****
    BloquearImgBuscar Me, Modo, ModoLineas
    BloquearImgZoom Me, Modo, ModoLineas
    BloquearImgFec Me, 0, Modo, ModoLineas
    
    imgFec(0).Enabled = (Modo = 1 Or Modo = 4)
    imgFec(0).visible = (Modo = 1 Or Modo = 4)
    BloquearTxt Text1(2), Not (Modo = 3 Or Modo = 4 Or Modo = 1)
    chkVistaPrevia(0).Enabled = (Modo <= 2)
    
    
    imgBuscar(2).Enabled = (Modo = 1 Or Modo = 4)
    imgBuscar(2).visible = (Modo = 1 Or Modo = 4)
    BloquearTxt Text1(3), Not (Modo = 3 Or Modo = 4 Or Modo = 1)
    
    
    PonerLongCampos

    Combo1(0).visible = (Modo = 5)
    Combo1(0).Enabled = (Modo = 5)
    
    If (Modo < 2) Or (Modo = 3) Then
        CargaGrid 1, False
    End If
    
    B = (Modo = 4) Or (Modo = 2)
    DataGridAux(1).Enabled = B
      
    For I = 0 To 3
        Check1(I).Enabled = (Modo = 1 Or Modo = 3 Or Modo = 4)
    Next I
    
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
    Toolbar1.Buttons(1).Enabled = False 'b And Not DeConsulta
    Me.mnNuevo.Enabled = False 'b And Not DeConsulta
    
    B = (Modo = 2 And Data1.Recordset.RecordCount > 0) And Not DeConsulta
    'Modificar
    Toolbar1.Buttons(2).Enabled = B
    Me.mnModificar.Enabled = B
    'eliminar
    Toolbar1.Buttons(3).Enabled = B
    Me.mnEliminar.Enabled = B
    
    
    'Imprimir
    'Toolbar1.Buttons(12).Enabled = (b Or Modo = 0)
    Toolbar1.Buttons(8).Enabled = B And Not DeConsulta
       
    ' *** si n'hi han llínies que tenen grids (en o sense tab) ***
'++monica: si insertamos lo he quitado
'    b = (Modo = 3 Or Modo = 4 Or Modo = 2) And Not DeConsulta
    B = (Modo = 2) And Not DeConsulta
    For I = 1 To 1
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
            SQL = "SELECT rordrecogida_incid.nroorden, rordrecogida_incid.idplaga, rplagasaux.nomplaga, rordrecogida_incid.nivel, "
            SQL = SQL & " CASE rordrecogida_incid.nivel WHEN 0 THEN ""<10"" WHEN 1 THEN ""10 a 20"" WHEN 2 THEN "">20"" END "
            SQL = SQL & " FROM rordrecogida_incid, rplagasaux "
            If enlaza Then
                SQL = SQL & ObtenerWhereCab(True)
            Else
                SQL = SQL & " WHERE rordrecogida_incid.nroorden = -1"
            End If
            SQL = SQL & " and rordrecogida_incid.idplaga = rplagasaux.idplaga"
            SQL = SQL & " ORDER BY rordrecogida_incid.idplaga"
            
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



Private Sub frmInc_DatoSeleccionado(CadenaSeleccion As String)
'Incidencias
    txtAux(Indice).Text = RecuperaValor(CadenaSeleccion, 1) 'codincid
    txtAux2(Indice).Text = RecuperaValor(CadenaSeleccion, 2) 'descripcion
End Sub

Private Sub frmMens_DatoSeleccionado(CadenaSeleccion As String)
    Text1(5).Text = RecuperaValor(CadenaSeleccion, 1)
End Sub


Private Sub frmOrdPrev_DatoSeleccionado(CadenaSeleccion As String)
Dim CadB As String
    
    If CadenaSeleccion <> "" Then
        CadB = "nroorden = " & DBSet(RecuperaValor(CadenaSeleccion, 1), "N")
        
        'Se muestran en el mismo form
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " " & Ordenacion
        PonerCadenaBusqueda
        Screen.MousePointer = vbDefault
    End If
End Sub


Private Sub frmVar_DatoSeleccionado(CadenaSeleccion As String)
    Text1(3).Text = RecuperaValor(CadenaSeleccion, 1) 'codvariedad
    FormateaCampo Text1(3)
    Text2(3).Text = RecuperaValor(CadenaSeleccion, 2) 'nombre variedad
End Sub

Private Sub frmZ_Actualizar(vCampo As String)
     Text1(20).Text = vCampo
End Sub

Private Sub imgAyuda_Click(Index As Integer)
Dim vCadena As String
    Select Case Index
        Case 0
           CargaCadenaAyuda vCadena, Index
    End Select
    MsgBox vCadena, vbInformation, "Descripción de Ayuda"
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
        Indice = 20
        frmZ.pTitulo = "Observaciones de la Clasificación"
        frmZ.pValor = Text1(Indice).Text
        frmZ.pModo = Modo
    
        frmZ.Show vbModal
        Set frmZ = Nothing
            
        PonerFoco Text1(Indice)
    End If
End Sub

Private Sub mnBuscar_Click()
Dim I As Integer
    BotonBuscar
End Sub

Private Sub mnEliminar_Click()
    BotonEliminar
End Sub

Private Sub mnImprimir_Click()
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String
    
    Screen.MousePointer = vbHourglass
    
    'Añadir el parametro de Empresa
    CadParam = "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    CadParam = CadParam & "pFecha=""" & Text1(1).Text & """|"
    numParam = numParam + 1

    cadNombreRPT = "rOrdenRecol.rpt"
    
    indRPT = 96 ' Ordenes de recoleccion
    If Not PonerParamRPT(indRPT, CadParam, numParam, nomDocu) Then Exit Sub
      
    'Nombre fichero .rpt a Imprimir
    cadNombreRPT = nomDocu
    ConSubInforme = False
    
    '[Monica]20/10/2014:añadida la condicion de que el campo no esté dado de baja
    cadSelect = " codcampo in (select codcampo from rcampos where nrocampo = " & DBSet(Text1(2).Text, "N") & " and codvarie = " & DBSet(Text1(3).Text, "N") & " and fecbajas is null ) "

    If InsertarTemporal2("rcampos", cadSelect) Then
        cadTitulo = "Reimpresión Orden de Recolección"
        cadFormula = "{tmpinformes.codusu} = " & vUsu.Codigo
        
        With frmImprimir
            .FormulaSeleccion = cadFormula
            .OtrosParametros = CadParam
            .NumeroParametros = numParam
            .SoloImprimir = False
            .EnvioEMail = False
            .Titulo = cadTitulo
            .NombreRPT = cadNombreRPT
            .ConSubInforme = True
            .Opcion = 0
            '[Monica]11/09/2013: ahora el nro de copias es 1
            .NroCopias = 1
            .Show vbModal
        End With
    End If
    
    
    Screen.MousePointer = vbDefault
End Sub


Private Function InsertarTemporal2(cTabla As String, cSelect As String) As Boolean
Dim SQL As String
Dim Sql2 As String
Dim Rs As ADODB.Recordset
Dim NumF As Long

    On Error GoTo eInsertarTemporal

    
    InsertarTemporal2 = False


    SQL = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute SQL
    
    SQL = "insert into tmpinformes (codusu, importe1, importe2, importe3) "
    SQL = SQL & " select " & vUsu.Codigo & ",nrocampo, codcampo, " & DBSet(Text1(0).Text, "N") & "  from rcampos where " & cSelect ' copia 1
    SQL = SQL & " order by 1,2,3 "
    
    conn.Execute SQL
    
    
    InsertarTemporal2 = True
    
    Exit Function
eInsertarTemporal:
    MuestraError Err.Number, "Insertar en Temporal", Err.Description
End Function




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
        Case 8 'Imprimir
            mnImprimir_Click
            
        Case btnPrimero To btnPrimero + 3 'Fleches Desplaçament
            Desplazamiento (Button.Index - btnPrimero)
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

    CadB = ObtenerBusqueda2(Me, BuscaChekc, 1)
    
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
    
    Set frmOrdPrev = New frmManOrdRecogidaPrev
    frmOrdPrev.cWhere = CadB
    frmOrdPrev.DatosADevolverBusqueda = "0|1|2|"
    frmOrdPrev.Show vbModal
    
    Set frmOrdPrev = Nothing

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
        MandaBusquedaPrevia ""
    Else
        CadenaConsulta = "Select * from " & NombreTabla & Ordenacion
        PonerCadenaBusqueda
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

    PonerModo 4

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
    cad = "¿Seguro que desea eliminar la Orden de Recogida?"
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
    Text2(3).Text = PonerNombreDeCod(Text1(3), "variedades", "nomvarie")
    
    ' *** si n'hi han llínies en datagrids ***
    'For i = 0 To DataGridAux.Count - 1
    For I = 1 To 1
        CargaGrid I, True
        If Not Adoaux(I).Recordset.EOF Then _
            PonerCamposForma2 Me, Adoaux(I), 2, "FrameAux" & I
    Next I

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
                        V = Adoaux(NumTabMto).Recordset.Fields(3) 'el 2 es el nº de llinia
                        Adoaux(NumTabMto).Recordset.Find (Adoaux(NumTabMto).Recordset.Fields(3).Name & " =" & V)
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
    
    ' *** canviar els arguments de la funcio, el mensage i repasar si n'hi ha codEmpre ***
'    If (Modo = 3) Then 'insertar
'        'comprobar si existe ya el cod. del campo clave primaria
'        If ExisteCP(Text1(0)) Then b = False
'    End If
    
    If Modo = 3 Or Modo = 4 Then
'        If b Then
'            If Not EsCampoSocioVariedad(Text1(5).Text, Text1(4).Text, Text1(3).Text) Then
'                MsgBox "El campo no es del socio o de la variedad introducida. Revise.", vbExclamation
'                b = False
'            End If
'        End If
    End If
    
    ' ************************************************************************************
    DatosOK = B
    
EDatosOK:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function

Private Sub PosicionarData()
Dim cad As String, Indicador As String

    ' *** canviar-ho per tota la PK de la capçalera, no llevar els () ***
    cad = "(nroorden=" & DBSet(Text1(0).Text, "N") & ")"
    
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
Dim vTipoMov As CTiposMov

    On Error GoTo FinEliminar

    conn.BeginTrans
    ' ***** canviar el nom de la PK de la capçalera, repasar codEmpre *******
    vWhere = " WHERE nroorden=" & Data1.Recordset!nroorden
    
    ' ***** elimina les llínies ****
    conn.Execute "DELETE FROM rordrecogida_incid " & vWhere
        
    'Eliminar la CAPÇALERA
    conn.Execute "Delete from " & NombreTabla & vWhere
       
    'Decrementar contador si borramos el ult. palet
    Set vTipoMov = New CTiposMov
    vTipoMov.DevolverContador "ORR", Val(Text1(0).Text)
    Set vTipoMov = Nothing
       
       
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
                Text2(3).Text = PonerNombreDeCod(Text1(3), "variedades", "nomvarie")
            End If
            
        Case 2 'orden
            PonerFormatoEntero Text1(Index)
       
        Case 1 ' formato fecha
            '[Monica]28/08/2013: comprobamos que la fecha esté en la campaña
            PonerFormatoFecha Text1(Index), True
       
    End Select
    
    
    ' ***************************************************************************
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = teclaBuscar Then
        If Modo = 1 Or Modo = 3 Or Modo = 4 Then
            Select Case Index
                Case 1: KEYFecha KeyAscii, 0 'fecha
                Case 3: KEYBusqueda KeyAscii, 2 'variedad
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
            SQL = "¿Seguro que desea eliminar la Incidencia de la Orden?"
            SQL = SQL & vbCrLf & "Nro.Orden: " & Adoaux(Index).Recordset!nroorden
            SQL = SQL & vbCrLf & "Incidencia: " & Adoaux(Index).Recordset!nomplaga
            If MsgBox(SQL, vbQuestion + vbYesNo) = vbYes Then
                Eliminar = True
                SQL = "DELETE FROM rordrecogida_incid "
                SQL = SQL & vWhere & " AND idplaga= " & Adoaux(Index).Recordset!idplaga
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
        If BLOQUEADesdeFormulario2(Me, Data1, 1) Then
        End If
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
        Case 1: vtabla = "rordrecogida_incid"
    End Select
    
    vWhere = ObtenerWhereCab(False)
    
    Select Case Index
        Case 1 ' *** pose els index dels tabs de llínies que tenen datagrid ***
            ' *** canviar la clau primaria de les llínies,
            'pasar a "" si no volem que mos sugerixca res a l'afegir ***
            
'            If Index = 0 Then NumF = SugerirCodigoSiguienteStr(vTabla, "numlinea", vWhere)

            AnyadirLinea DataGridAux(Index), Adoaux(Index)
    
            anc = DataGridAux(Index).Top
            If DataGridAux(Index).Row < 0 Then
                anc = anc + 240
            Else
                anc = anc + DataGridAux(Index).RowTop(DataGridAux(Index).Row) + 5
            End If
            
            LLamaLineas Index, ModoLineas, anc
            
            Combo1(0).ListIndex = 0

            Select Case Index
                ' *** valor per defecte a l'insertar i formateig de tots els camps ***
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
                anc = anc + 240
            Else
                anc = anc + DataGridAux(Index).RowTop(DataGridAux(Index).Row) + 5
            End If
            
            LLamaLineas Index, ModoLineas, anc
            
            
    End Select
    
    Select Case Index
        ' *** valor per defecte al modificar dels camps del grid ***
        Case 1 'incidencias
            For J = 8 To 9
                txtAux(J).Text = DataGridAux(Index).Columns(J - 8).Text
            Next J
            txtAux2(9).Text = DataGridAux(Index).Columns(2).Text
            For I = 9 To 9
                BloquearTxt txtAux(I), True
            Next I
            BloquearbtnBuscar Me, Modo, ModoLineas, "FrameAux1"
            
    End Select
    
    PosicionarCombo Combo1(0), DataGridAux(Index).Columns(3).Text  ' nivel de incidencia    LLamaLineas Index, ModoLineas, anc
   
    ' *** foco al 1r camp visible de les llinies en grids que no siga PK (en o sense tab) ***
    Select Case Index
        Case 1 'incidencias
            PonerFocoCmb Combo1(0)
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
            For jj = 9 To 9
                txtAux(jj).visible = B
                txtAux(jj).Top = alto
            Next jj
            txtAux2(9).visible = B
            txtAux2(9).Top = alto
            btnBuscar(1).visible = B
            btnBuscar(1).Enabled = (xModo = 1)
            btnBuscar(1).Top = alto
            
            Combo1(0).Top = alto
            Combo1(0).visible = B
            Combo1(0).Enabled = B
            
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
    Combo1(0).AddItem "<10"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 0
    Combo1(0).AddItem "10-20"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 1
    Combo1(0).AddItem ">20"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 2
End Sub

Private Sub ToolbarAyuda_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1
            LanzaVisorMimeDocumento Me.hWnd, DireccionAyuda & IdPrograma & ".html"
    End Select
End Sub

Private Sub ToolbarDes_ButtonClick(ByVal Button As MSComctlLib.Button)
    Desplazamiento (Button.Index)
End Sub

Private Sub txtaux_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = teclaBuscar Then
        Select Case Index
            Case 9:  KEYImage KeyAscii, 1
        End Select
    Else
        KEYpress KeyAscii
    End If
End Sub

Private Sub KEYImage(KeyAscii As Integer, Indice As Integer)
    KeyAscii = 0
    btnBuscar_Click (Indice)
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
        
        Case 9 ' codigo de incidencia
            If PonerFormatoEntero(txtAux(Index)) Then
                txtAux2(Index).Text = PonerNombreDeCod(txtAux(Index), "rplagasaux", "nomplaga")
                If txtAux2(Index).Text = "" Then
                    cadMen = "No existe el Código de Incidencia. Revise."
                    MsgBox cadMen, vbInformation
                    PonerFoco txtAux(Index)
                End If
            Else
                txtAux2(Index).Text = ""
            End If
        
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
    
    '[Monica]23/09/2014: si me insertan el concepto de no hay incidencias no puede haber ningun otro concepto
    If B Then
        If ComprobarCero(txtAux(9).Text) = 8 Then
            SQL = "select count(*) from rordrecogida_incid where nroorden = " & DBSet(Text1(0).Text, "N") & " and idplaga < 8"
            If TotalRegistros(SQL) <> 0 Then
                MsgBox "No puede haber una incidencia junto con Revisión realizada sin incidencias. Revise.", vbExclamation
                B = False
            End If
        Else
            SQL = "select count(*) from rordrecogida_incid where nroorden = " & DBSet(Text1(0).Text, "N") & " and idplaga = 8"
            If TotalRegistros(SQL) <> 0 Then
                MsgBox "No puede haber una incidencia junto con Revisión realizada sin incidencias. Revise.", vbExclamation
                B = False
            End If
        End If
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

Private Sub imgBuscar_Click(Index As Integer)
    TerminaBloquear
    Indice = Index + 3
     Select Case Index
        Case 2 'variedades
            Set frmVar = New frmComVar
            frmVar.DatosADevolverBusqueda = "0|1|"
            frmVar.CodigoActual = Text1(3).Text
            frmVar.Show vbModal
            Set frmVar = Nothing
            PonerFoco Text1(3)
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
            tots = "N||||0|;S|txtAux(9)|T|Incidencia|1500|;S|btnBuscar(1)|B||195|;"
            tots = tots & "S|txtAux2(9)|T|Denominación|5450|;N||||0|;S|Combo1(0)|C|Nivel|1520|;"

            arregla tots, DataGridAux(Index), Me, 350
            
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
        Case 1: nomframe = "FrameAux1" 'incidencias
    End Select
    
    If DatosOkLlin(nomframe) Then
        TerminaBloquear
        If InsertarDesdeForm2(Me, 2, nomframe) Then
            Select Case NumTabMto
                Case 0 ' *** els index de les llinies en grid (en o sense tab) ***
                    CargaGrid NumTabMto, True
                    
                    B = BLOQUEADesdeFormulario2(Me, Data1, 1)
                    If B Then BotonAnyadirLinea NumTabMto
                Case 1
                    B = BLOQUEADesdeFormulario2(Me, Data1, 1)
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
        Case 1: nomframe = "FrameAux1" 'incidencias
    End Select
    ModificarLinea = False
    If DatosOkLlin(nomframe) Then
        TerminaBloquear
        If ModificaDesdeFormulario2(Me, 2, nomframe) Then
            ModoLineas = 0
            Select Case NumTabMto
                Case 1
                    V = Adoaux(NumTabMto).Recordset.Fields(1) 'el 2 es el nº de llinia
            End Select
            CargaGrid NumTabMto, True
            
            ' *** si n'hi han tabs ***
'            SituarTab (NumTabMto + 1)

            ' *** si n'hi han tabs que no tenen datagrid, posar el if ***
            PonerFocoGrid Me.DataGridAux(NumTabMto)
            Adoaux(NumTabMto).Recordset.Find (Adoaux(NumTabMto).Recordset.Fields(1).Name & " =" & V)
            
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
    vWhere = vWhere & " nroorden=" & Me.Data1.Recordset!nroorden
    
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


Private Sub InsertarCabecera()
Dim vTipoMov As CTiposMov 'Clase Tipo Movimiento
Dim SQL As String
Dim actualiza As Boolean
    Set vTipoMov = New CTiposMov
    If vTipoMov.Leer(CodTipoMov) Then
        If Text1(0).Text = "" Then
            Text1(0).Text = vTipoMov.ConseguirContador(CodTipoMov)
            actualiza = True
        Else
            actualiza = False
        End If
        
        SQL = CadenaInsertarDesdeForm(Me)
        If InsertarOferta(SQL, vTipoMov, actualiza) Then
            CadenaConsulta = "Select * from " & NombreTabla & ObtenerWhereCP(True) & Ordenacion
            PonerCadenaBusqueda
            PonerModo 2
        
            If Not NotaExistente Then
                Data1.RecordSource = "Select * from " & NombreTabla & Ordenacion
                PosicionarData
                BotonModificar
                cmdAceptar_Click
            End If
        
        End If
    End If
    Text1(0).Text = Format(Text1(0).Text, "0000000")
End Sub


Private Function InsertarOferta(vSQL As String, vTipoMov As CTiposMov, ActualizarContador As Boolean) As Boolean
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
    'Comprobar si mientras tanto se incremento el contador de Factura
    'para ello vemos si existe una oferta con ese contador y si existe la incrementamos
    Do
        If ExisteNota(Text1(0).Text) Then
            devuelve = Text1(0).Text
        Else
            devuelve = ""
        End If
        
'        devuelve = DevuelveDesdeBDNew(cAgro, NombreTabla, "numnotac", "numnotac", Text1(0).Text, "N")
        If devuelve <> "" Then
            'Ya existe el contador incrementarlo
            Existe = True
            If ActualizarContador Then
                vTipoMov.IncrementarContador (CodTipoMov)
                Text1(0).Text = vTipoMov.ConseguirContador(CodTipoMov)
            Else
                MsgBox "Número de Nota ya existe. Reintroduzca.", vbExclamation
                PonerFoco Text1(0)
                NotaExistente = True
                InsertarOferta = False
                Exit Function
            End If
            cambiaSQL = True
        Else
            Existe = False
        End If
    Loop Until Not Existe
    If cambiaSQL Then vSQL = CadenaInsertarDesdeForm(Me)
    
    
    'Aqui empieza transaccion
    conn.BeginTrans
    MenError = "Error al insertar en la tabla Cabecera de Factura (" & NombreTabla & ")."
    conn.Execute vSQL, , adCmdText
    
    cadMen = ""
    Sql3 = "select * from rordrecogida where numnotac = " & DBSet(Text1(0).Text, "N")
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql3, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not Rs.EOF Then
        cadMen = "Insertando Clasificacion: " & cadMen
'26-05-2009: Santi no quiere que se calcule el transporte quiere meterlo él
'           cuando se da de alta.
'        If bol Then
'            bol = ActualizarTransporte(Rs, cadMen)
'            cadMen = "Actualizando Gastos de Transporte" & cadMen
'        End If
'26-05-2009
    End If
    
    Set Rs = Nothing
    
    If bol Then
        If ActualizarContador Then
            MenError = "Error al actualizar el contador de la Factura."
            vTipoMov.IncrementarContador (CodTipoMov)
        End If
    End If
    
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
    
    SQL = " numnotac= " & Text1(0).Text
    If conWhere Then SQL = " WHERE " & SQL
    ObtenerWhereCP = SQL
    
    If Err.Number <> 0 Then MuestraError Err.Number, "Obteniendo cadena WHERE.", Err.Description
End Function

Private Function ModificaCabecera() As Boolean
Dim B As Boolean
Dim MenError As String
Dim SQL As String

    On Error GoTo EModificarCab

    conn.BeginTrans
    
    B = True
    
        
    If B Then B = ModificaDesdeFormulario1(Me, 1) 'ModificaDesdeFormulario2(Me, 2, "Frame2")

EModificarCab:
    If Err.Number <> 0 Or Not B Then
        MenError = "Modificando Entrada." & vbCrLf & "----------------------------" & vbCrLf & MenError
        If Err.Number <> 0 Then
            MuestraError Err.Number, MenError, Err.Description
        Else
            MsgBox MenError, vbExclamation
        End If
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


