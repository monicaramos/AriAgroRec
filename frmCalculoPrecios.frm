VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmCalculoPrecios 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cálculo Precios"
   ClientHeight    =   8310
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8715
   Icon            =   "frmCalculoPrecios.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8310
   ScaleWidth      =   8715
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "Cuenta Principal|N|N|0|1|cltebanc|ctaprpal|0||"
   Begin VB.Frame FrameAux1 
      BorderStyle     =   0  'None
      Height          =   3930
      Left            =   210
      TabIndex        =   21
      Top             =   3330
      Width           =   8340
      Begin VB.TextBox txtaux1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   290
         Index           =   6
         Left            =   6180
         MaxLength       =   7
         TabIndex        =   35
         Tag             =   "Kilos|N|N|||tmppreciosaux|kilosnet|##,###,##0||"
         Text            =   "Kilos"
         Top             =   2940
         Visible         =   0   'False
         Width           =   585
      End
      Begin VB.TextBox txtaux1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   290
         Index           =   5
         Left            =   7410
         MaxLength       =   7
         TabIndex        =   30
         Tag             =   "Precio|N|N|||tmppreciosaux|precio|#0.0000||"
         Text            =   "Precio"
         Top             =   2925
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.TextBox txtaux1 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   290
         Index           =   4
         Left            =   180
         MaxLength       =   10
         TabIndex        =   26
         Tag             =   "Usuario|N|N|||tmppreciosaux|codusu|000000|S|"
         Text            =   "usu"
         Top             =   2910
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.TextBox txtaux1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   290
         Index           =   3
         Left            =   6810
         MaxLength       =   7
         TabIndex        =   29
         Tag             =   "Porcentaje|N|N|||tmppreciosaux|porcentaje|##0.00||"
         Text            =   "Porcen"
         Top             =   2940
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.TextBox txtaux1 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   290
         Index           =   2
         Left            =   1740
         MaxLength       =   6
         TabIndex        =   25
         Tag             =   "Tipo Factura|N|N|0|1|tmppreciosaux|tipofact||S|"
         Text            =   "tipo"
         Top             =   2940
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.TextBox txtaux1 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   290
         Index           =   0
         Left            =   1050
         MaxLength       =   6
         TabIndex        =   24
         Tag             =   "Código Variedad|N|N|1|999999|tmppreciosaux|codvarie|000000|S|"
         Text            =   "var"
         Top             =   2910
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.TextBox txtaux1 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   290
         Index           =   1
         Left            =   2385
         MaxLength       =   2
         TabIndex        =   27
         Tag             =   "Calidad|N|N|||tmppreciosaux|codcalid|00|S|"
         Text            =   "cal"
         Top             =   2925
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.CommandButton cmdAux 
         Appearance      =   0  'Flat
         Caption         =   "+"
         Height          =   315
         Index           =   1
         Left            =   3060
         TabIndex        =   23
         ToolTipText     =   "Buscar calidad"
         Top             =   2925
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.TextBox txtAux2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   1
         Left            =   3240
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   22
         Text            =   "Nombre calidad"
         Top             =   2925
         Visible         =   0   'False
         Width           =   3285
      End
      Begin MSComctlLib.Toolbar ToolAux 
         Height          =   390
         Index           =   1
         Left            =   45
         TabIndex        =   28
         Top             =   0
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   688
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   3
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Object.Visible         =   0   'False
               Object.ToolTipText     =   "Nuevo"
               Object.Tag             =   "2"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Modificar"
               Object.Tag             =   "2"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Object.Visible         =   0   'False
               Object.ToolTipText     =   "Eliminar"
               Object.Tag             =   "2"
            EndProperty
         EndProperty
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
         Bindings        =   "frmCalculoPrecios.frx":000C
         Height          =   3225
         Index           =   1
         Left            =   45
         TabIndex        =   31
         Top             =   450
         Width           =   8160
         _ExtentX        =   14393
         _ExtentY        =   5689
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
   Begin VB.Frame Frame2 
      Height          =   2670
      Index           =   0
      Left            =   180
      TabIndex        =   12
      Top             =   570
      Width           =   8370
      Begin VB.CheckBox Check1 
         Caption         =   "Cálculo por porcentajes"
         Height          =   255
         Left            =   6060
         TabIndex        =   3
         Top             =   1170
         Width           =   2025
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   1
         Left            =   6060
         MaxLength       =   15
         TabIndex        =   7
         Top             =   2100
         Width           =   1545
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   1
         Left            =   4350
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Tag             =   "Tipo Factura|N|N|0|1|rprecios|tipofact||S|"
         Top             =   2070
         Width           =   1305
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   4
         Left            =   1440
         MaxLength       =   30
         TabIndex        =   4
         Tag             =   "Texto|T|N|||rprecios|textoper|||"
         Text            =   "123456789012345678901234567890"
         Top             =   1620
         Width           =   4185
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   3
         Left            =   4080
         MaxLength       =   10
         TabIndex        =   2
         Tag             =   "Fecha Fin|F|S|||rprecios|fechafin|dd/mm/yyyy||"
         Text            =   "1234567890"
         Top             =   1170
         Width           =   1110
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   2
         Left            =   1440
         MaxLength       =   10
         TabIndex        =   1
         Tag             =   "Fecha Inicio|F|N|||rprecios|fechaini|dd/mm/yyyy||"
         Text            =   "1234567890"
         Top             =   1170
         Width           =   1200
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   0
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Tag             =   "Tipo Factura|N|N|0|1|rprecios|tipofact||S|"
         Top             =   2070
         Width           =   1305
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   1470
         MaxLength       =   6
         TabIndex        =   0
         Tag             =   "Código Variedad|N|N|1|999999|rprecios|codvarie|000000|S|"
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   2640
         MaxLength       =   40
         TabIndex        =   16
         Top             =   480
         Width           =   4950
      End
      Begin VB.Label Label3 
         Caption         =   "Tipo Precio"
         ForeColor       =   &H00972E0B&
         Height          =   255
         Left            =   3210
         TabIndex        =   34
         Top             =   2100
         Width           =   1035
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha"
         ForeColor       =   &H00972E0B&
         Height          =   255
         Left            =   240
         TabIndex        =   33
         Top             =   840
         Width           =   705
      End
      Begin VB.Label Label1 
         Caption         =   "Importe"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   255
         Index           =   1
         Left            =   6030
         TabIndex        =   32
         Top             =   1860
         Width           =   1215
      End
      Begin VB.Label Label8 
         Caption         =   "Texto"
         ForeColor       =   &H00972E0B&
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   1650
         Width           =   705
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   1
         Left            =   3780
         Picture         =   "frmCalculoPrecios.frx":0024
         ToolTipText     =   "Buscar fecha"
         Top             =   1170
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   0
         Left            =   1140
         Picture         =   "frmCalculoPrecios.frx":00AF
         ToolTipText     =   "Buscar fecha"
         Top             =   1170
         Width           =   240
      End
      Begin VB.Label Label22 
         Caption         =   "Hasta"
         Height          =   255
         Left            =   3090
         TabIndex        =   19
         Top             =   1200
         Width           =   585
      End
      Begin VB.Label Label18 
         Caption         =   "Desde"
         Height          =   255
         Left            =   450
         TabIndex        =   18
         Top             =   1200
         Width           =   930
      End
      Begin VB.Label Label9 
         Caption         =   "Recolectado"
         ForeColor       =   &H00972E0B&
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   2100
         Width           =   1035
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   1170
         ToolTipText     =   "Buscar Variedad"
         Top             =   480
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Variedad"
         ForeColor       =   &H00972E0B&
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   450
         Width           =   690
      End
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Index           =   0
      Left            =   195
      TabIndex        =   10
      Top             =   7350
      Width           =   2865
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
         TabIndex        =   11
         Top             =   180
         Width           =   2655
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   7335
      TabIndex        =   9
      Top             =   7440
      Width           =   1035
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   6120
      TabIndex        =   8
      Top             =   7440
      Width           =   1035
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   3690
      Top             =   7485
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
      TabIndex        =   15
      Top             =   0
      Width           =   8715
      _ExtentX        =   15372
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   20
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Pedir Datos"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Generar Precios"
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
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Modificar"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Eliminar"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Generación Masiva"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Imprimir"
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Primero"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Anterior"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Siguiente"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Último"
            ImageIndex      =   9
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   7320
      TabIndex        =   14
      Top             =   7440
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Menu mnOpciones 
      Caption         =   "&Opciones"
      Begin VB.Menu mnPedirDatos 
         Caption         =   "&Pedir Datos"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnGenerarPrecios 
         Caption         =   "&Generar Precios"
         Shortcut        =   ^G
      End
      Begin VB.Menu mnBarra1 
         Caption         =   "-"
      End
      Begin VB.Menu mnSalir 
         Caption         =   "&Salir"
         Shortcut        =   ^S
      End
   End
End
Attribute VB_Name = "frmCalculoPrecios"
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

Public Tipo As Byte ' 0 = variedades que no son del grupo 5 ni 6
                    ' 1 = variedades del grupo 5 (almazara)
                    ' 2 = variedades del grupo 6 (bodega)

Private CadB1 As String
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
Private WithEvents frmCalid As frmManCalidades 'calidades
Attribute frmCalid.VB_VarHelpID = -1

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
Dim indice As Byte 'Index del text1 on es poses els datos retornats des d'atres Formularis de Mtos
Dim CadB As String

Dim vSeccion As CSeccion
Dim b As Boolean

Private BuscaChekc As String

'GENERALES PARA PASARLE A CRYSTAL REPORT
Private cadFormula As String 'Cadena con la FormulaSelection para Crystal Report
Private CadParam As String 'Cadena con los parametros para Crystal Report
Private numParam As Byte 'Numero de parametros que se pasan a Crystal Report
Private cadSelect As String 'Cadena para comprobar si hay datos antes de abrir Informe
Private cadTitulo As String 'Titulo para la ventana frmImprimir
Private cadNombreRPT As String 'Nombre del informe

Dim VarieAnt As String

Private Sub Check1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub cmdAceptar_Click()

    Screen.MousePointer = vbHourglass
    
    On Error GoTo Error1
    
    Select Case Modo
            
        Case 3 'INSERTAR
            If DatosOk Then
                CargarCalidades
                PonerModo 2
            Else
'                PonerModo 0
            End If
            
        ' *** si n'hi han llínies ***
        Case 5 'LLÍNIES
            Select Case ModoLineas
                Case 2 'modificar llínies
                    ModificarLinea

                    
                    PonerModo 2
                    
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
            Set frmCalid = New frmManCalidades
            frmCalid.DatosADevolverBusqueda = "0|1|2|3|"
            frmCalid.CodigoActual = txtAux1(1).Text
            frmCalid.ParamVariedad = txtAux1(0).Text
            frmCalid.Show vbModal
            Set frmCalid = Nothing
            PonerFoco txtAux1(1)

    End Select
    If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
End Sub

' *** si n'hi han combos a la capçalera ***
Private Sub Combo1_GotFocus(Index As Integer)
    If Modo = 1 Then Combo1(Index).BackColor = vbYellow
End Sub

Private Sub Combo1_LostFocus(Index As Integer)
    If Combo1(Index).BackColor = vbYellow Then Combo1(Index).BackColor = vbWhite
End Sub

Private Sub Combo1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
    If PrimeraVez Then PrimeraVez = False
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Screen.MousePointer = vbDefault
    If Modo = 4 Or Modo = 5 Then TerminaBloquear
End Sub

Private Sub Form_Load()
Dim i As Integer

    PrimeraVez = True
    
    ' ICONETS DE LA BARRA
    btnPrimero = 17 'index del botó "primero"
    With Me.Toolbar1
        .HotImageList = frmPpal.imgListComun_OM
        .DisabledImageList = frmPpal.imgListComun_BN
        .ImageList = frmPpal.imgListComun
        'l'1 i el 2 son separadors
        .Buttons(3).Image = 1   'Buscar
        .Buttons(4).Image = 13   'Generar precios
        'el 5 i el 6 son separadors
        .Buttons(7).Image = 3   'Insertar
        .Buttons(8).Image = 4   'Modificar
        .Buttons(9).Image = 5   'Borrar
        'el 10 i el 11 son separadors
        .Buttons(11).Image = 17  ' generacion masiva
        
        .Buttons(13).Image = 10  'Imprimir
        .Buttons(14).Image = 11  'Eixir
        'el 13 i el 14 son separadors
        .Buttons(btnPrimero).Image = 6  'Primer
        .Buttons(btnPrimero + 1).Image = 7 'Anterior
        .Buttons(btnPrimero + 2).Image = 8 'Següent
        .Buttons(btnPrimero + 3).Image = 9 'Últim
    End With
    
    ' ******* si n'hi han llínies *******
    'ICONETS DE LES BARRES ALS TABS DE LLÍNIA
    For i = 1 To ToolAux.Count
        With Me.ToolAux(i)
            .HotImageList = frmPpal.imgListComun_OM16
            .DisabledImageList = frmPpal.imgListComun_BN16
            .ImageList = frmPpal.imgListComun16
            .Buttons(1).Image = 3   'Insertar
            .Buttons(2).Image = 4   'Modificar
            .Buttons(3).Image = 5   'Borrar
        End With
    Next i
    ' ***********************************
    
    
    'cargar IMAGES de busqueda
    For i = 0 To Me.imgBuscar.Count - 1
        Me.imgBuscar(i).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next i
    
    LimpiarCampos   'Neteja els camps TextBox
'    ' ******* si n'hi han llínies *******
    DataGridAux(1).ClearFields
    
    '*** canviar el nom de la taula i l'ordenació de la capçalera ***
    NombreTabla = "rprecios"
    Ordenacion = " ORDER BY codvarie"
    '************************************************
    

    'ASignamos un SQL al DATA1
    Data1.ConnectionString = conn
    '***** canviar el nom de la PK de la capçalera; repasar codEmpre *************
    Data1.RecordSource = "Select * from " & NombreTabla & " where codvarie=-1"

    Data1.Refresh
       
    
    ModoLineas = 0
       
         
    ' *** si n'hi han combos (capçalera o llínies) ***
    CargaCombo
    ' ************************************************
    
    If DatosADevolverBusqueda = "" Then
        PonerModo 0
    Else
        PonerModo 1 'búsqueda
        ' *** posar de groc els camps visibles de la clau primaria de la capçalera ***
        Text1(0).BackColor = vbYellow 'codclien
        ' ****************************************************************************
    End If
End Sub


Private Sub LimpiarCampos()
Dim i As Integer

    On Error Resume Next
    
    limpiar Me   'Mètode general: Neteja els controls TextBox
    lblIndicador.Caption = ""
'    Me.chkAbonos(0).Value = 0
    
    For i = 0 To Combo1.Count - 1
        Combo1(i).ListIndex = -1
'        Me.chkAbonos(I).Value = 0
    Next i
    
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
Dim i As Integer, NumReg As Byte
Dim b As Boolean

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
    b = (Modo = 2)
    'Posar Fleches de desplasament visibles
    NumReg = 1
    If Not Data1.Recordset.EOF Then
        If Data1.Recordset.RecordCount > 1 Then NumReg = 2 'Només es per a saber que n'hi ha + d'1 registre
    End If
    DesplazamientoVisible Me.Toolbar1, btnPrimero, b, NumReg
    
    '---------------------------------------------
    b = Modo <> 0 And Modo <> 2
    cmdCancelar.visible = b
    cmdAceptar.visible = b
       
    'Bloqueja els camps Text1 si no estem modificant/Insertant Datos
    'Si estem en Insertar a més neteja els camps Text1
    BloquearText1 Me, Modo
    BloquearCombo Me, Modo
    
    BloquearChk Check1, (Modo = 0 Or Modo = 5)
    
    Text1(1).Locked = Not b  '((Not b) And (Modo <> 1))
    If b Then
          Text1(1).BackColor = vbWhite
    Else
          Text1(1).BackColor = &H80000018 'groc
    End If
    
    
    
    ' ***** bloquejar tots els controls visibles de la clau primaria de la capçalera ***
'    If Modo = 4 Then _
'        BloquearTxt Text1(0), True 'si estic en  modificar, bloqueja la clau primaria
    ' **********************************************************************************
    
    ' **** si n'hi han imagens de buscar en la capçalera *****
    BloquearImgBuscar Me, Modo, ModoLineas
    For i = 0 To imgFec.Count - 1
        BloquearImgFec Me, i, Modo
    Next i
'    BloquearImgFec Me, 25, Modo, ModoLineas
    BloquearImgZoom Me, Modo, ModoLineas
    ' ********************************************************
    ' *** si n'hi han llínies i imagens de buscar que no estiguen als grids ******
    'Llínies Departaments
    b = (Modo = 5 And (ModoLineas = 1 Or ModoLineas = 2) And (NumTabMto = 0))
'    BloquearImage imgBuscar(3), Not b
'    BloquearImage imgBuscar(4), Not b
'    BloquearImage imgBuscar(7), Not b
'    imgBuscar(3).Enabled = b
'    imgBuscar(3).visible = b
    ' ****************************************************************************
            
    
    PonerLongCampos

    If (Modo < 2) Or (Modo = 3) Then
'        CargaGrid 0, False
        CargaGrid 1, False
    End If
    
    b = (Modo = 4) Or (Modo = 2)
    

    DataGridAux(1).Enabled = b
    
    b = (Modo = 5) And (NumTabMto = 1) 'And (ModoLineas <> 3)
    For i = 1 To txtAux1.Count - 1
        BloquearTxt txtAux1(i), Not b
    Next i
    b = (Modo = 5) And (NumTabMto = 1) And ModoLineas = 2
    BloquearTxt txtAux1(1), b
    BloquearBtn cmdAux(1), b
     '-----------------------------
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
End Sub

Private Sub PonerModoOpcionesMenu(Modo)
'Actives unes Opcions de Menú i Toolbar según el modo en que estem
Dim b As Boolean, bAux As Boolean
Dim i As Byte
    
    'Barra de CAPÇALERA
    '------------------------------------------
    'b = (Modo = 2 Or Modo = 0 Or Modo = 1)
    b = (Modo = 2 Or Modo = 0)
    'Pedir datos
    Toolbar1.Buttons(3).Enabled = b
    Me.mnPedirDatos.Enabled = b
    'Generar Precios
    Toolbar1.Buttons(4).Enabled = b
    Me.mnGenerarPrecios.Enabled = b
    
    ' *** si n'hi han llínies que tenen grids (en o sense tab) ***
    b = (Modo = 3 Or Modo = 4 Or Modo = 2)
    For i = 1 To ToolAux.Count
        ToolAux(i).Buttons(1).Enabled = b
        If b Then bAux = (b And Me.Adoaux(i).Recordset.RecordCount > 0)
        ToolAux(i).Buttons(2).Enabled = bAux
        ToolAux(i).Buttons(3).Enabled = bAux
    Next i
    ' ****************************************
    
'    ' *** si n'hi han tabs que no tenen grids ***
'    i = 3
'    If AdoAux(i).Recordset.EOF Then
'        ToolAux(i).Buttons(1).Enabled = b
'        ToolAux(i).Buttons(2).Enabled = False
'        ToolAux(i).Buttons(3).Enabled = False
'    Else
'        ToolAux(i).Buttons(1).Enabled = False
'        ToolAux(i).Buttons(2).Enabled = b
'    End If
    ' *******************************************
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
Dim Tabla As String
    
    ' ********* si n'hi han tabs, dona igual si en datagrid o no ***********
    Select Case Index
       Case 1 ' calidades
            Tabla = "tmppreciosaux"
            Sql = "SELECT tmppreciosaux.codusu, tmppreciosaux.codvarie, tmppreciosaux.tipofact, tmppreciosaux.codcalid, rcalidad.nomcalid, tmppreciosaux.kilosnet, "
            Sql = Sql & "tmppreciosaux.porcentaje, tmppreciosaux.precio"
            Sql = Sql & " FROM " & Tabla & " INNER JOIN rcalidad ON tmppreciosaux.codvarie = rcalidad.codvarie "
            Sql = Sql & " and tmppreciosaux.codcalid = rcalidad.codcalid  "
            If enlaza Then
                Sql = Sql & ObtenerWhereCab(True)
            Else
                Sql = Sql & " WHERE tmppreciosaux.codusu = -1"
            End If
            Sql = Sql & " ORDER BY " & Tabla & ".codcalid "
            
            
    End Select
    ' ********************************************************************************
    
    MontaSQLCarga = Sql
End Function


Private Sub frmC_Selec(vFecha As Date)
Dim indice As Byte
    indice = CByte(Me.cmdAux(0).Tag + 2)
    txtAux1(indice).Text = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub frmC1_Selec(vFecha As Date)
Dim indice As Byte
    indice = CByte(Me.imgFec(0).Tag)
    Text1(indice).Text = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub frmCalid_DatoSeleccionado(CadenaSeleccion As String)
    txtAux1(0).Text = RecuperaValor(CadenaSeleccion, 1) 'codigo variedad
    txtAux1(1).Text = RecuperaValor(CadenaSeleccion, 3) 'codigo calidad
    txtAux2(1).Text = RecuperaValor(CadenaSeleccion, 4) 'nombre calidad
End Sub


Private Sub frmVar_DatoSeleccionado(CadenaSeleccion As String)
    Text1(0).Text = RecuperaValor(CadenaSeleccion, 1) 'codvariedad
    FormateaCampo Text1(0)
    Text2(0).Text = RecuperaValor(CadenaSeleccion, 2) 'nombre variedad
End Sub

Private Sub frmZ_Actualizar(vCampo As String)
     Text1(indice).Text = vCampo
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
                indice = Index + 2
       End Select
       
       Me.imgFec(0).Tag = indice
       
       PonerFormatoFecha Text1(indice)
       If Text1(indice).Text <> "" Then frmC1.NovaData = CDate(Text1(indice).Text)
    
       Screen.MousePointer = vbDefault
       frmC1.Show vbModal
       Set frmC1 = Nothing
       PonerFoco Text1(indice)
    
End Sub

Private Sub imgZoom_Click(Index As Integer)
    
    Set frmZ = New frmZoom

    Select Case Index
        Case 0
            indice = 21
            frmZ.pTitulo = "Observaciones del Campo"
            frmZ.pValor = Text1(indice).Text
            frmZ.pModo = Modo
            frmZ.Show vbModal
            Set frmZ = Nothing
            PonerFoco Text1(indice)
    End Select
            
End Sub


Private Sub mnGenerarPrecios_Click()
    BotonGenerarPrecios
End Sub

Private Sub mnPedirDatos_Click()
    BotonPedirDatos
End Sub

Private Sub mnSalir_Click()
    Unload Me
End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.Index
        Case 3  'Búscar
           mnPedirDatos_Click
        Case 4  'Generar Precios
            mnGenerarPrecios_Click
        Case 14    'Eixir
            mnSalir_Click
            
        Case btnPrimero To btnPrimero + 3 'Fleches Desplaçament
            Desplazamiento (Button.Index - btnPrimero)
    End Select
End Sub

Private Sub BotonPedirDatos()
Dim Nombre As String

    'Vaciamos todos los Text
    LimpiarCampos
    'Vaciamos el ListView
'    InicializarListView
    CargaGrid 1, False
    
    PonerModo 3
    
    'fecha recepcion
'    Text1(4).Text = "PRUEBA DE FUNCIONAMIENTO"
    
    PonerFoco Text1(0)
End Sub

Private Function HayPrecios() As Boolean
Dim Sql As String

    HayPrecios = False
    
    Sql = "select count(*) from tmppreciosaux where codusu = " & vUsu.Codigo & " and codvarie = " & DBSet(Text1(0).Text, "N")
    Sql = Sql & " and precio <> 0 "
    
    HayPrecios = (TotalRegistros(Sql) <> 0)

End Function


Private Sub BotonGenerarPrecios()
Dim vFactu As CFacturaTer
Dim Cad As String
Dim Sql As String
Dim Sql1 As String
Dim Sql2 As String
Dim cadMen As String

Dim Contador As Long
Dim Rs As ADODB.Recordset


    Screen.MousePointer = vbHourglass
    On Error GoTo Error1
    
    
    ' comprobamos que los datos de cabecera no estan introducidos en la tabla de precios
    ' y si lo están preguntamos si hay que updatearlos
    If Not HayPrecios Then
        MsgBox "No hay precios calculados. Debe modificar previamente.", vbExclamation
        Exit Sub
    Else
        If Me.Check1.Value Then
            Sql = "select sum(porcentaje) from tmppreciosaux where codusu = " & vUsu.Codigo
            If DevuelveValor(Sql) <> 100 Then
                cadMen = "La suma de porcentajes es diferente al 100%."
                cadMen = cadMen & vbCrLf & vbCrLf & " ¿ Desea continuar ?"
                If MsgBox(cadMen, vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
            End If
        End If
    End If

    conn.BeginTrans
    
    
    Sql = "select max(contador) "
    Sql = Sql & " from rprecios where codvarie = " & DBSet(Text1(0).Text, "N")
    Sql = Sql & " and fechaini = " & DBSet(Text1(2).Text, "F")
    Sql = Sql & " and fechafin = " & DBSet(Text1(3).Text, "F")
    Sql = Sql & " and tipofact = " & Combo1(1).ListIndex
    
    Contador = DevuelveValor(Sql)
    If Contador <> 0 Then
        Sql = "select count(*) from rprecios_calidad where codvarie = " & DBSet(Text1(0).Text, "N")
        Sql = Sql & " and tipofact = " & Combo1(1).ListIndex
        Sql = Sql & " and contador = " & DBSet(Contador, "N")
        Select Case Combo1(0).ListIndex
            Case 0
                Sql = Sql & " and precoop <> 0"
            Case 1
                Sql = Sql & " and presocio <> 0"
        End Select
        
        If TotalRegistros(Sql) = 0 Then
            ' no existen registros para ese tipo de precio : actualizamos o insertamos registros de calidad dependiendo
            Sql = "select codcalid, precio from tmppreciosaux where codusu = " & vUsu.Codigo
            Sql = Sql & " and codvarie = " & DBSet(Text1(0).Text, "N")
            Sql = Sql & " and precio <> 0 "
            Sql = Sql & " order by 1"
            
            Set Rs = New ADODB.Recordset
            Rs.Open Sql, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
                    
            While Not Rs.EOF
                Sql2 = "select count(*) from rprecios_calidad where codvarie = " & DBSet(Text1(0).Text, "N")
                Sql2 = Sql2 & " and tipofact = " & Combo1(1).ListIndex
                Sql2 = Sql2 & " and contador = " & DBSet(Contador, "N")
                Sql2 = Sql2 & " and codcalid = " & DBSet(Rs!codcalid, "N")
                
                If TotalRegistros(Sql2) = 0 Then
                    'insertamos
                    Sql2 = "insert into rprecios_calidad (codvarie,tipofact,contador,codcalid,precoop,presocio)"
                    Sql2 = Sql2 & " values (" & DBSet(Text1(0).Text, "N") & ","
                    Sql2 = Sql2 & Combo1(1).ListIndex & ","
                    Sql2 = Sql2 & DBSet(Contador, "N") & ","
                    Sql2 = Sql2 & DBSet(Rs!codcalid, "N") & ","
                    Select Case Combo1(0).ListIndex
                        Case 0 'cooperativa
                            Sql2 = Sql2 & DBSet(Rs!Precio, "N") & ",0)" ' & ValorNulo & ")"
                        Case 1 'socio
                            Sql2 = Sql2 & "0," & DBSet(Rs!Precio, "N") & ")"
                    End Select
                    
                    conn.Execute Sql2
                Else
                    'modificamos
                    Sql2 = "update rprecios_calidad set "
                    Select Case Combo1(0).ListIndex
                        Case 0
                            Sql2 = Sql2 & "precoop = " & DBSet(Rs!Precio, "N")
                        Case 1
                            Sql2 = Sql2 & "presocio = " & DBSet(Rs!Precio, "N")
                    End Select
                    Sql2 = Sql2 & " where codvarie = " & DBSet(Text1(0).Text, "N")
                    Sql2 = Sql2 & " and tipofact = " & Combo1(1).ListIndex
                    Sql2 = Sql2 & " and contador = " & DBSet(Contador, "N")
                    Sql2 = Sql2 & " and codcalid = " & DBSet(Rs!codcalid, "N")
                    
                    conn.Execute Sql2
                
                End If
                Rs.MoveNext
            Wend
            
            Set Rs = Nothing
            
        Else
            ' existen registros del tipo que tenemos : preguntamos si quieren updatearlos
            cadMen = "Existen precios de "
            Select Case Combo1(1).ListIndex
                Case 0
                    cadMen = cadMen & "anticipos "
                Case 1
                    cadMen = cadMen & "liquidacion "
            End Select
            cadMen = cadMen & "para ese rango de fechas." & vbCrLf & vbCrLf
            cadMen = cadMen & " ¿ Desea crear un contador nuevo ? "
            
            Select Case MsgBox(cadMen, vbQuestion + vbYesNoCancel)
                Case vbYes
                    ' creamos un regitro nuevo
                    Contador = Contador + 1
                    
                    Sql2 = "insert into rprecios (codvarie,tipofact,contador,fechaini,fechafin,textoper,precioindustria) values ("
                    Sql2 = Sql2 & DBSet(Text1(0).Text, "N") & "," & Combo1(1).ListIndex & ","
                    Sql2 = Sql2 & DBSet(Contador, "N") & "," & DBSet(Text1(2).Text, "F") & "," & DBSet(Text1(3).Text, "F") & ","
                    Sql2 = Sql2 & DBSet(Text1(4).Text, "T") & "," & ValorNulo & ")"
                    
                    conn.Execute Sql2
                    
                    ' creamos las lineas de precios
                    Sql2 = "insert into rprecios_calidad (codvarie,tipofact,contador,codcalid,precoop,presocio) "
                    Sql2 = Sql2 & " select " & DBSet(Text1(0).Text, "N") & "," & Combo1(1).ListIndex & ","
                    Sql2 = Sql2 & DBSet(Contador, "N") & ",codcalid,"
                    Select Case Combo1(0).ListIndex
                        Case 0
                            Sql2 = Sql2 & "precio, 0" ' & ValorNulo
                        Case 1
                            Sql2 = Sql2 & "0, precio "
                    End Select
                    Sql2 = Sql2 & " from tmppreciosaux where codusu = " & vUsu.Codigo
                    Sql2 = Sql2 & " and codvarie = " & DBSet(Text1(0).Text, "N")
                    Sql2 = Sql2 & " and precio <> 0"
                    
                    conn.Execute Sql2
                    
                Case vbNo
                    ' actualizamos  los existentes o creamos
                    
                    ' primero ponemos rprecios_calidad con precio 0
                    Sql2 = "update rprecios_calidad set "
                    Select Case Combo1(0).ListIndex
                        Case 0
                            Sql2 = Sql2 & " precoop = 0 "
                        Case 1
                            Sql2 = Sql2 & " presocio = 0 "
                    End Select
                    Sql2 = Sql2 & " where codvarie = " & DBSet(Text1(0).Text, "N")
                    Sql2 = Sql2 & " and tipofact = " & Combo1(1).ListIndex
                    Sql2 = Sql2 & " and contador = " & DBSet(Contador, "N")
                    
                    conn.Execute Sql2
                    
                    ' ahora actualizamos los existentes o creamos
                    Sql2 = "select * from tmppreciosaux where codusu = " & vUsu.Codigo & " and codvarie = " & DBSet(Text1(0).Text, "N")
                    Sql2 = Sql2 & " and precio <> 0 "
                    Sql2 = Sql2 & " order by codcalid "
                    
                    Set Rs = New ADODB.Recordset
                    Rs.Open Sql2, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
                    
                    While Not Rs.EOF
                        Sql = "select count(*) from rprecios_calidad where codvarie = " & DBSet(Text1(0).Text, "N")
                        Sql = Sql & " and tipofact = " & Combo1(1).ListIndex
                        Sql = Sql & " and contador = " & DBSet(Contador, "N")
                        Sql = Sql & " and codcalid = " & DBSet(Rs!codcalid, "N")
                        
                        If TotalRegistros(Sql) <> 0 Then
                            Sql = "update rprecios_calidad set "
                            Select Case Combo1(0).ListIndex
                                Case 0
                                    Sql = Sql & " precoop = " & DBSet(Rs!Precio, "N")
                                Case 1
                                    Sql = Sql & " presocio= " & DBSet(Rs!Precio, "N")
                            End Select
                            Sql = Sql & " where codvarie = " & DBSet(Text1(0).Text, "N")
                            Sql = Sql & " and tipofact = " & Combo1(1).ListIndex
                            Sql = Sql & " and contador = " & DBSet(Contador, "N")
                            Sql = Sql & " and codcalid = " & DBSet(Rs!codcalid, "N")
                            
                        Else
                            Sql = "insert into rprecios_calidad (codvarie,tipofact,contador,codcalid,precoop,presocio) values ("
                            Sql = Sql & DBSet(Text1(0).Text, "N") & "," & Combo1(1).ListIndex & "," & DBSet(Contador, "N") & ","
                            Sql = Sql & DBSet(Rs!codcalid, "N") & ","
                            Select Case Combo1(0).ListIndex
                                Case 0
                                    Sql = Sql & DBSet(Rs!Precio, "N") & ",0)" ' & ValorNulo & ")"
                                Case 1
                                    Sql = Sql & "0," & DBSet(Rs!Precio, "N") & ")"
                            End Select
                        End If
                        conn.Execute Sql

                        Rs.MoveNext
                    Wend
                    Set Rs = Nothing
                    
                Case vbCancel
                    ' no hacemos nada
                    
            End Select
            
        End If
    Else
        ' creamos un regitro nuevo
        Sql = "select max(contador) "
        Sql = Sql & " from rprecios where codvarie = " & DBSet(Text1(0).Text, "N")
        Sql = Sql & " and tipofact = " & Combo1(1).ListIndex

        Contador = DevuelveValor(Sql)
        Contador = Contador + 1
        
        Sql2 = "insert into rprecios (codvarie,tipofact,contador,fechaini,fechafin,textoper,precioindustria) values ("
        Sql2 = Sql2 & DBSet(Text1(0).Text, "N") & "," & Combo1(1).ListIndex & ","
        Sql2 = Sql2 & DBSet(Contador, "N") & "," & DBSet(Text1(2).Text, "F") & "," & DBSet(Text1(3).Text, "F") & ","
        Sql2 = Sql2 & DBSet(Text1(4).Text, "T") & "," & ValorNulo & ")"
        
        conn.Execute Sql2
        
        ' creamos las lineas de precios
        Sql2 = "insert into rprecios_calidad (codvarie,tipofact,contador,codcalid,precoop,presocio) "
        Sql2 = Sql2 & " select " & DBSet(Text1(0).Text, "N") & "," & Combo1(1).ListIndex & ","
        Sql2 = Sql2 & DBSet(Contador, "N") & ",codcalid,"
        Select Case Combo1(0).ListIndex
            Case 0
                Sql2 = Sql2 & "precio, 0" '& ValorNulo
            Case 1
                Sql2 = Sql2 & "0, precio "
        End Select
        Sql2 = Sql2 & " from tmppreciosaux where codusu = " & vUsu.Codigo
        Sql2 = Sql2 & " and codvarie = " & DBSet(Text1(0).Text, "N")
        Sql2 = Sql2 & " and precio <> 0"
        
        conn.Execute Sql2
    
    End If

    conn.CommitTrans

    Screen.MousePointer = vbDefault

    MsgBox "Proceso realizado correctamente.", vbExclamation

    Exit Sub
    
Error1:
    conn.RollbackTrans
    
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub




Private Sub cmdRegresar_Click()
Dim Cad As String
Dim Aux As String
Dim i As Integer
Dim J As Integer

    If Data1.Recordset.EOF Then
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



Private Sub PonerCampos()
Dim i As Integer
Dim CodPobla As String, desPobla As String
Dim CPostal As String, desProvi As String, desPais As String

    If Data1.Recordset.EOF Then Exit Sub
    PonerCamposForma2 Me, Data1, 1 'opcio=1: posa el format o els camps de la capçalera
    
    ' *** si n'hi han llínies en datagrids ***
    For i = 1 To DataGridAux.Count ' - 1
        If i <> 3 Then
            CargaGrid i, True
            If Not Adoaux(i).Recordset.EOF Then _
                PonerCamposForma2 Me, Adoaux(i), 2, "FrameAux" & i
        End If
    Next i
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
Dim i As Integer
Dim V

    Select Case Modo
        Case 3 ' Insertar
                LimpiarCampos
                PonerModo 0
                CargaGrid 1, False
                ' *** foco al primer camp visible de la capçalera ***
                PonerFoco Text1(0)
                ' ***************************************************

        Case 5 'LLÍNIES
            Select Case ModoLineas
                Case 2 'modificar llínies
                    ModoLineas = 0

                    ' *** si n'hi han tabs ***
                    SituarTab (NumTabMto)
                    'SSTab1.Tab = 1
                    'SSTab2.Tab = NumTabMto
                    ' ***********************

                    PonerModo 2
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

            
'            SumaTotalPorcentajes
'
'            PosicionarData
'
'            TerminaBloquear
'
'            ' *** si n'hi han llínies en grids i camps fora d'estos ***
'            If Not AdoAux(1).Recordset.EOF Then
'                DataGridAux_RowColChange 1, 1, 1
'            Else
'                LimpiarCamposFrame 1
'            End If
'            ' *********************************************************
    End Select
End Sub


Private Function DatosOk() As Boolean
Dim b As Boolean
Dim Sql As String
Dim Cad As String
'Dim Datos As String

    On Error GoTo EDatosOK

    DatosOk = False
    b = CompForm2(Me, 1)
    If Not b Then Exit Function
    
    ' *** canviar els arguments de la funcio, el mensage i repasar si n'hi ha codEmpre ***
    If (Modo = 3) Then 'pedir datos
        If Text1(1).Text = "" Then
            MsgBox "El importe debe de tener un valor. Reintroduzca.", vbExclamation
            PonerFoco Text1(1)
            b = False
        End If
    End If
    
    ' ************************************************************************************
    DatosOk = b
    
EDatosOK:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function


Private Sub PosicionarData()
Dim Cad As String, Indicador As String

    ' *** canviar-ho per tota la PK de la capçalera, no llevar els () ***
    Cad = "(codvarie=" & Text1(0).Text & " and codusu = " & vUsu.Codigo & " and codcalid = " & txtAux1(1).Text & " )"
    ' ***************************************
    
    ' *** gastar SituarData o SituarDataMULTI depenent de si la PK es simple o composta ***
    'If SituarDataMULTI(Data1, cad, Indicador) Then
    If SituarDataMULTI(Data1, Cad, Indicador) Then
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
    vWhere = " WHERE codvarie=" & Data1.Recordset!CodVarie
    vWhere = vWhere & " and tipofact = " & Data1.Recordset!TipoFact
    vWhere = vWhere & " and contador = " & Data1.Recordset!Contador
        ' ***********************************************************************
        
    ' ***** elimina les llínies ****
    conn.Execute "DELETE FROM rprecios_calidad " & vWhere

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
    
    If Index = 5 Then
        If Tipo = 0 Then
            Text1(Index).Enabled = (Combo1(0).ListIndex = 2)
        Else
            Text1(Index).Enabled = True
        End If
    End If
    
End Sub

Private Sub Text1_LostFocus(Index As Integer)
Dim cadMen As String
Dim Nuevo As Boolean

    If Not PerderFocoGnral(Text1(Index), Modo) Then Exit Sub

    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    Select Case Index
        Case 0 'Variedad
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
'                    Select Case Tipo
'                        Case 0  'cualquier tipo
                            If EsVariedadGrupo5(Text1(Index)) Then
                                MsgBox "Esta variedad es del grupo de almazara. Reintroduzca.", vbExclamation
                                PonerFoco Text1(Index)
                            Else
                                If EsVariedadGrupo6(Text1(Index)) Then
                                    MsgBox "Esta variedad es del grupo de bodega. Reintroduzca.", vbExclamation
                                    PonerFoco Text1(Index)
                                End If
                            End If
'                        Case 1  'almazara
'                            If Not EsVariedadGrupo5(Text1(Index)) Then
'                                MsgBox "Esta variedad no es del grupo de almazara. Reintroduzca.", vbExclamation
'                                PonerFoco Text1(Index)
'                            End If
'                        Case 2  'bodega
'                            If Not EsVariedadGrupo6(Text1(Index)) Then
'                                MsgBox "Esta variedad no es del grupo de bodega. Reintroduzca.", vbExclamation
'                                PonerFoco Text1(Index)
'                            End If
'                    End Select
                End If
            Else
                Text2(Index).Text = ""
            End If
        
        Case 2, 3 ' fechas de inicio y fin
            If Index = 2 And Text1(2).Text = "" Then Text1(2).Text = Format(vParam.FecIniCam, "dd/mm/yyyy")
            If Index = 3 And Text1(3).Text = "" Then Text1(3).Text = Format(vParam.FecFinCam, "dd/mm/yyyy")
                    
            If PonerFormatoFecha(Text1(Index), True) Then
                If Text1(2).Text <> "" And Text1(3).Text <> "" Then
                    If CDate(Text1(2).Text) > CDate(Text1(3).Text) Then
                        MsgBox "La Fecha Inicio debe ser inferior a la Fecha Fin. Revise", vbExclamation
                        Text1(Index).Text = ""
                        PonerFoco Text1(Index)
                    End If
                End If
            End If
'        Case 1 'contador
'            PonerFormatoEntero Text1(Index)
                
        Case 4 'texto
            If Text1(Index).Text <> "" Then Text1(Index).Text = UCase(Text1(Index).Text)
            
        Case 1 'importe
            PonerFormatoDecimal Text1(Index), 3
            
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
        If Index <> 21 Then KEYpress KeyAscii
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


Private Sub KEYBusqueda(KeyAscii As Integer, indice As Integer)
    KeyAscii = 0
    imgBuscar_Click (indice)
End Sub

' **** si n'hi han camps de descripció a la capçalera ****
Private Sub PosarDescripcions()
Dim NomEmple As String

    On Error GoTo EPosarDescripcions

    Text2(0).Text = PonerNombreDeCod(Text1(0), "variedades", "nomvarie", "codvarie", "N")
    
'    PonerDatosPartida
'
'    If Text1(5).Text <> "" Then
'        Text2(5).Text = Round2(ImporteSinFormato(Text1(5).Text) / cFaneca, 4)
'        PonerFormatoDecimal Text2(5), 7
'    End If
'
'    If Text1(6).Text <> "" Then
'        Text2(6).Text = Round2(ImporteSinFormato(Text1(6).Text) / cFaneca, 4)
'        PonerFormatoDecimal Text2(6), 7
'    End If
'
'    If Text1(7).Text <> "" Then
'        Text2(7).Text = Round2(ImporteSinFormato(Text1(7).Text) / cFaneca, 4)
'        PonerFormatoDecimal Text2(7), 7
'    End If
    
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
        Case 1 'calidad
            Sql = "¿Seguro que desea eliminar la calidad?"
            Sql = Sql & vbCrLf & "Calidad: " & Adoaux(Index).Recordset!codcalid & " - " & Adoaux(Index).Recordset!nomcalid
            If MsgBox(Sql, vbQuestion + vbYesNo) = vbYes Then
                Eliminar = True
                Sql = "DELETE FROM rprecios_calidad "
                Sql = Sql & vWhere & " and codcalid = " & DBLet(Adoaux(Index).Recordset!codcalid, "N")
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
        If BLOQUEADesdeFormulario2(Me, Data1, 1) Then
        End If
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
    ' **************************************************

    ' *** posar el nom del les distintes taules de llínies ***
    Select Case Index
        Case 1: vtabla = "rprecios_calidad"
    End Select
    ' ********************************************************

    vWhere = ObtenerWhereCab(False)

    Select Case Index
         Case 1   'clasificacion
            ' *** canviar la clau primaria de les llínies,
            'pasar a "" si no volem que mos sugerixca res a l'afegir ***
'            If Index <> 4 Then ' *** els index als que no volem sugerir-li un codi ***
'                NumF = SugerirCodigoSiguienteStr(vTabla, "codsecci", vWhere)
'            Else
'                NumF = ""
'            End If
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
                Case 1 'calidades
                    For i = 0 To txtAux1.Count - 1
                        txtAux1(i).Text = ""
                    Next i
                    txtAux1(0).Text = Text1(0).Text 'codvariedad
                    txtAux1(2).Text = Combo1(0).ListIndex  'tipo de factura
                    txtAux1(4).Text = Text1(1).Text 'contador
                    
                    txtAux1(1).Text = "" 'calidad
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
    ' *********************************

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
        Case 1 'calidades
         
            txtAux1(4).Text = DataGridAux(Index).Columns(0).Text 'codusu
            txtAux1(0).Text = DataGridAux(Index).Columns(1).Text 'codvarie
            txtAux1(2).Text = DataGridAux(Index).Columns(2).Text 'tipo
            
            txtAux1(1).Text = DataGridAux(Index).Columns(3).Text 'calidad
            txtAux2(1).Text = DataGridAux(Index).Columns(4).Text ' nombre calidad
            txtAux1(6).Text = DataGridAux(Index).Columns(5).Text 'kilos
            txtAux1(3).Text = DataGridAux(Index).Columns(6).Text 'porcentaje
            txtAux1(5).Text = DataGridAux(Index).Columns(7).Text 'precio
            
            
    End Select

    LLamaLineas Index, ModoLineas, anc

    ' *** foco al 1r camp visible de les llinies en grids que no siga PK (en o sense tab) ***
    Select Case Index
        Case 1 'calidades
            PonerFoco txtAux1(3)
    End Select
    ' ***************************************************************************************
End Sub


Private Sub LLamaLineas(Index As Integer, xModo As Byte, Optional alto As Single)
Dim jj As Integer
Dim b As Boolean

    ' *** si n'hi han tabs sense datagrid posar el If ***
    If Index <> 3 Then DeseleccionaGrid DataGridAux(Index)
    ' ***************************************************

    b = (xModo = 1 Or xModo = 2) 'Insertar o Modificar Llínies
    Select Case Index
        Case 1 'calidad
            For jj = 1 To txtAux1.Count - 1
                If jj = 3 Then
                    txtAux1(jj).visible = b
                    txtAux1(jj).Top = alto
                End If
            Next jj
            
    End Select
End Sub



Private Function DatosOkLlin(nomframe As String) As Boolean
Dim Rs As ADODB.Recordset
Dim Sql As String
Dim b As Boolean
Dim Cant As Integer
Dim Mens As String
Dim vFact As Byte, vDocum As Byte

    DatosOkLlin = True
    
    On Error GoTo EDatosOKLlin

    Mens = ""
    DatosOkLlin = False

    b = CompForm2(Me, 2, nomframe) 'Comprovar formato datos ok
    If Not b Then Exit Function
    
    If b And (Modo = 5 And ModoLineas = 1) Then  'insertar
        'comprobar si existe ya el cod. de la calidad para ese campo
        Sql = ""
'        SQL = DevuelveDesdeBDNew(cAgro, "rprecios_calidad", "codcalid", "codvarie", txtaux1(0).Text, "N", , "tipofact", txtaux1(2).Text, "N", "codcalid", txtaux1(1).Text, "N")
        If Sql <> "" Then
            MsgBox "Ya existe la calidad. Revise.", vbExclamation
            PonerFoco txtAux1(1)
            b = False
        End If
    End If
    
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
    DatosOkLlin = b

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

' *** si n'hi han formularis de buscar codi a les llínies ***
Private Sub imgBuscar_Click(Index As Integer)
    TerminaBloquear
    
    Select Case Index
        Case 0 'situacion
            Set frmVar = New frmComVar
'            frmVar.DeConsulta = True
            frmVar.DatosADevolverBusqueda = "0|1|"
'            frmVar.CodigoActual = Text1(2).Text
            frmVar.Show vbModal
            Set frmVar = Nothing
            PonerFoco Text1(2)
        
    End Select
    
    If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
End Sub


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
Dim b As Boolean
Dim i As Byte
Dim tots As String

    On Error GoTo ECarga

    tots = MontaSQLCarga(Index, enlaza)

    b = DataGridAux(Index).Enabled
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

    For i = 0 To DataGridAux(Index).Columns.Count - 1
        DataGridAux(Index).Columns(i).AllowSizing = False
    Next i
    
    CargaGridGnral Me.DataGridAux(Index), Me.Adoaux(Index), tots, PrimeraVez
    
    
    'DataGridAux(Index).Enabled = b
'    PrimeraVez = False
    
    Select Case Index
        Case 1 'clasificacion segun la calidad
            'si es visible|control|tipo campo|nombre campo|ancho control|
            ' codusu, codvarie, tipofact, codcalid, rcalidad.nomcalid, kilosnet, porcentaje, precio "
            tots = "N||||0|;N||||0|;N||||0|;S|txtaux1(1)|T|Cód.|800|;" 'S|cmdAux(1)|B|||;" 'codsocio,codsecci
            tots = tots & "S|txtAux2(1)|T|Nombre|2870|;"
            tots = tots & "S|txtaux1(6)|T|Kilos|1300|;"
            tots = tots & "S|txtaux1(3)|T|Porcentaje|1300|;"
            tots = tots & "S|txtaux1(5)|T|Precio|1300|;"
            arregla tots, DataGridAux(Index), Me
        
            DataGridAux(Index).Columns(4).Alignment = dbgLeft
            DataGridAux(Index).Columns(5).Alignment = dbgRight
            DataGridAux(Index).Columns(6).Alignment = dbgRight
            
            b = (Modo = 4) And ((ModoLineas = 1) Or (ModoLineas = 2))

            If (enlaza = True) And (Not Adoaux(Index).Recordset.EOF) Then 'per a que pose els valors de les arees de text la primera volta
            Else
                For i = 0 To 4
                    txtAux1(i).Text = ""
                Next i
                txtAux2(1).Text = ""
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
Dim b As Boolean

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
            b = BLOQUEADesdeFormulario2(Me, Data1, 1)
            
            '++monica: en caso de estar insertando seccion y que no existan las
            'cuentas contables hacemos esto para que las inserte en contabilidad.
'            If NumTabMto = 1 Then
'               txtAux2(4).Text = PonerNombreCuenta(txtaux1(4), 3, Text1(0))
'               txtAux2(5).Text = PonerNombreCuenta(txtaux1(5), 3, Text1(0))
'            End If
            
            Select Case NumTabMto
                Case 0, 1 ' *** els index de les llinies en grid (en o sense tab) ***
                    CargaGrid NumTabMto, True
                    If b Then BotonAnyadirLinea NumTabMto
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
Dim Cad As String
    On Error Resume Next

    ' *** posa els noms del frames, tant si son de grid com si no ***
    Select Case NumTabMto
        Case 1: nomframe = "FrameAux1" 'porcentajes
    End Select
    ' **************************************************************

    If DatosOkLlin(nomframe) Then
        TerminaBloquear
        If ModificaDesdeFormulario2(Me, 2, nomframe) Then
            ModoLineas = 0
                    
            If Check1.Value = 0 Then
                CalculodePrecios
            Else
                CalculodePreciosPorcentaje
            End If

            V = Adoaux(NumTabMto).Recordset.Fields(3) 'el 2 es el nº de llinia
            CargaGrid NumTabMto, True


            ' *** si n'hi han tabs que no tenen datagrid, posar el if ***
            If NumTabMto <> 3 Then
                DataGridAux(NumTabMto).SetFocus
                Adoaux(NumTabMto).Recordset.Find (Adoaux(NumTabMto).Recordset.Fields(3).Name & " =" & V)
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
    vWhere = vWhere & " tmppreciosaux.codvarie=" & Val(Text1(0).Text)
    vWhere = vWhere & " and tmppreciosaux.codusu = " & Val(vUsu.Codigo)
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


' ********* si n'hi han combos a la capçalera ************
Private Sub CargaCombo()
Dim Ini As Integer
Dim Fin As Integer
Dim i As Integer

    ' *** neteje els combos, els pose valor i seleccione el valor per defecte ***
    For i = 0 To Combo1.Count - 1
        Combo1(i).Clear
    Next i
    
    'tipo de precios
    Combo1(0).AddItem "Cooperativa"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 0
    Combo1(0).AddItem "Socio"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 1
    
    Combo1(1).Clear
    Combo1(1).AddItem "Anticipo"
    Combo1(1).ItemData(Combo1(1).NewIndex) = 0
    Combo1(1).AddItem "Liquidacion"
    Combo1(1).ItemData(Combo1(1).NewIndex) = 1
    
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
        Case 1 ' calidad
            If PonerFormatoEntero(txtAux1(Index)) Then
                txtAux2(Index).Text = PonerNombreDeCod(txtAux1(Index), "rcalidad", "nomcalid", "codcalid", "N", , "codvarie", txtAux1(0).Text, "N")
                If txtAux2(Index).Text = "" Then
                    cadMen = "No existe la Calidad: " & txtAux1(Index).Text & vbCrLf
                    cadMen = cadMen & "¿Desea crearla?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmCalid = New frmManCalidades
                        frmCalid.DatosADevolverBusqueda = "0|1|"
                        frmCalid.NuevoCodigo = txtAux1(Index).Text
                        txtAux1(Index).Text = ""
                        TerminaBloquear
                        frmCalid.Show vbModal
                        Set frmCalid = Nothing
                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Data1, 1
                    Else
                        txtAux1(Index).Text = ""
                    End If
                    PonerFoco txtAux1(Index)
                End If
            Else
                txtAux2(Index).Text = ""
            End If

        Case 3 'porcentaje
            If PonerFormatoDecimal(txtAux1(Index), 4) Then
                If ModoLineas = 1 Then cmdAceptar.SetFocus
            End If

    End Select

    ' ******************************************************************************
End Sub

Private Sub txtAux1_GotFocus(Index As Integer)
   If Not txtAux1(Index).MultiLine Then ConseguirFocoLin txtAux1(Index)
End Sub

Private Sub txtAux1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'    If Not txtAux1(Index).MultiLine Then KEYdown KeyCode
    
On Error GoTo EKeyD
    ' si no estamos en muestra salimos
    If Index <> 3 Then Exit Sub
    
    Select Case KeyCode
        Case 38 'Desplazamieto Fecha Hacia Arriba
'050509
'            cmdAceptar_Click
            ModificarLinea
            
'            If Me.DataGridAux(0).Bookmark > 0 Then
'                DataGridAux(0).Bookmark = DataGridAux(0).Bookmark - 1
'            End If
            PasarAntReg
        Case 40 'Desplazamiento Flecha Hacia Abajo
            'ModificarExistencia
'050509
'            cmdAceptar_Click
            ModificarLinea
            
            PasarSigReg
    End Select
EKeyD:
    If Err.Number <> 0 Then Err.Clear
    
End Sub

Private Sub txtAux1_KeyPress(Index As Integer, KeyAscii As Integer)
    If Not txtAux1(Index).MultiLine Then
        If KeyAscii = teclaBuscar Then
            If Modo = 5 And (ModoLineas = 1 Or ModoLineas = 2) Then
                Select Case Index
                    Case 1: KEYBusqueda KeyAscii, 1 'calidad
                End Select
            End If
        Else
            If Index = 3 Then ' estoy introduciendo la muestra
               If KeyAscii = 13 Then 'ENTER
                    PonerFormatoDecimal txtAux1(Index), 3
                    If ModoLineas = 2 Then
                        '050509 cmdAceptar_Click 'ModificarExistencia
                        ModificarLinea

                        PasarSigReg
                    End If
                    If ModoLineas = 1 Then
                        cmdAceptar.SetFocus
                    End If
                    
                    '050509
'                    If ModoLineas = 1 Then
'                        cmdAceptar.SetFocus
'                    End If
               ElseIf KeyAscii = 27 Then
                    cmdCancelar_Click 'ESC
               End If
            Else
                KEYpress KeyAscii
            End If
        End If
    End If
End Sub


Private Function CargarCalidades() As Boolean
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim i As Integer
Dim arrData()
Dim TotalPorc As Currency
       
    On Error GoTo eCargarCalidades
       
    CargarCalidades = False
       
    Sql = "delete from tmppreciosaux where codusu = " & vUsu.Codigo
     
    conn.Execute Sql
     
    Sql = "insert into tmppreciosaux (codusu, codvarie, tipofact, codcalid, kilosnet, porcentaje, precio) "
    Sql = Sql & " select " & vUsu.Codigo & "," & DBSet(Text1(0).Text, "N") & "," & DBSet(Combo1(1).ListIndex, "N")
    Sql = Sql & ",rhisfruta_clasif.codcalid, sum(rhisfruta_clasif.kilosnet) kilosnet, 0, 0  from rhisfruta_clasif "
    Sql = Sql & " where rhisfruta_clasif.codvarie = " & DBSet(Text1(0), "N")
    Sql = Sql & " and rhisfruta_clasif.numalbar in (select numalbar from rhisfruta, rsocios where fecalbar >= "
    Sql = Sql & DBSet(Text1(2).Text, "F") & " and fecalbar <= " & DBSet(Text1(3).Text, "F")
    Sql = Sql & " and recolect = " & Combo1(0).ListIndex
    Sql = Sql & " and tipoentr <> 1 and tipoentr <> 3 " ' entradas que no sean venta campo ni industria
    Sql = Sql & " and rsocios.tipoprod <> 1 "  ' que el socio no sea tercero
    Sql = Sql & " and rhisfruta.codsocio = rsocios.codsocio) "
    Sql = Sql & " group by 1, 2, 3, 4 "
    Sql = Sql & " having kilosnet <> 0 "
    Sql = Sql & " order by 1, 2, 3, 4 "
     
    conn.Execute Sql
     
    CargaGrid 1, True

    CargarCalidades = True

    Exit Function

eCargarCalidades:
    MuestraError Err.Number, "Cargando kilos de calidades", Err.Description
End Function


Private Sub CalculodePrecios()
Dim Sql As String
Dim i As Currency
Dim Rs As ADODB.Recordset
Dim vCalcul As Currency
Dim vCalcul1 As Currency
Dim PrecioLin As Currency
Dim Sql2 As String

    On Error GoTo eCalculodePrecios

    Sql = "select * from tmppreciosaux where codusu = " & vUsu.Codigo
    Sql = Sql & " and porcentaje <>0 and porcentaje is not null "
    Sql = Sql & " order by codcalid "

    vCalcul = DevuelveValor("select sum(kilosnet * porcentaje / 100) from tmppreciosaux where codusu = " & vUsu.Codigo)
    vCalcul1 = Round2(ImporteSinFormato(Text1(1).Text) / vCalcul, 4)
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText

    While Not Rs.EOF
        
        PrecioLin = vCalcul1 + Round2(vCalcul1 * (Rs!Porcentaje - 100) / 100, 4)
        
        Sql2 = "update tmppreciosaux set precio = " & DBSet(PrecioLin, "N")
        Sql2 = Sql2 & " where codusu = " & vUsu.Codigo
        Sql2 = Sql2 & " and codcalid = " & DBSet(Rs!codcalid, "N")
        
        conn.Execute Sql2
        
        Rs.MoveNext
    Wend

    Set Rs = Nothing
    
    Exit Sub
    
eCalculodePrecios:
    MuestraError Err.Number, "Calculo de Precios", Err.Description
End Sub




Private Sub CalculodePreciosPorcentaje()
Dim Sql As String
Dim i As Currency
Dim Rs As ADODB.Recordset
Dim vCalcul As Currency
Dim vCalcul1 As Currency
Dim PrecioLin As Currency
Dim Sql2 As String
Dim vNetoTotal As Long
Dim ImporteLin As Currency
Dim ImporteTotal As Currency

    On Error GoTo eCalculodePreciosPorcentaje

    Sql = "select * from tmppreciosaux where codusu = " & vUsu.Codigo
    Sql = Sql & " and porcentaje <>0 and porcentaje is not null "
    Sql = Sql & " order by codcalid "

    vNetoTotal = DevuelveValor("select sum(kilosnet) from tmppreciosaux where codusu = " & vUsu.Codigo)
    ImporteTotal = ImporteSinFormato(Text1(1).Text)
'
'    vCalcul = DevuelveValor("select sum(kilosnet * porcentaje / 100) from tmppreciosaux where codusu = " & vUsu.Codigo)
'    vCalcul1 = Round2(ImporteSinFormato(Text1(1).Text) / vCalcul, 4)
'
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockOptimistic, adCmdText

    While Not Rs.EOF
        ImporteLin = Round2(ImporteTotal * Rs!Porcentaje / 100, 2)
        
        PrecioLin = Round2(ImporteLin / Rs!KilosNet, 4) 'vCalcul1 + Round2(vCalcul1 * (RS!Porcentaje - 100) / 100, 4)
        
        Sql2 = "update tmppreciosaux set precio = " & DBSet(PrecioLin, "N")
        Sql2 = Sql2 & " where codusu = " & vUsu.Codigo
        Sql2 = Sql2 & " and codcalid = " & DBSet(Rs!codcalid, "N")
        
        conn.Execute Sql2
        
        Rs.MoveNext
    Wend

    Set Rs = Nothing
    
    Exit Sub
    
eCalculodePreciosPorcentaje:
    MuestraError Err.Number, "Calculo de Precios", Err.Description
End Sub


Private Sub PasarSigReg()
'Nos situamos en el siguiente registro
    If Me.DataGridAux(1).Bookmark < Me.Adoaux(1).Recordset.RecordCount Then
'        DataGridAux(0).Row = DataGridAux(0).Row + 1
        DataGridAux(1).Bookmark = DataGridAux(1).Bookmark + 1
        BotonModificarLinea 1
    ElseIf DataGridAux(1).Bookmark = Adoaux(1).Recordset.RecordCount Then
'        PonerFocoBtn Me.cmdAceptar
        BotonModificarLinea 1
    End If
End Sub


Private Sub PasarAntReg()
'Nos situamos en el siguiente registro
    If Me.DataGridAux(1).Bookmark > 1 Then
'        DataGridAux(0).Row = DataGridAux(0).Row + 1
        DataGridAux(1).Bookmark = DataGridAux(1).Bookmark - 1
        BotonModificarLinea 1
    ElseIf DataGridAux(1).Bookmark = 1 Then
'        PonerFocoBtn Me.cmdAceptar
        BotonModificarLinea 1
    End If
End Sub


