VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmPOZLecturas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Introducción de Lecturas "
   ClientHeight    =   6240
   ClientLeft      =   45
   ClientTop       =   30
   ClientWidth     =   15345
   Icon            =   "frmPOZLecturas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6240
   ScaleWidth      =   15345
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameBotonGnral 
      Height          =   705
      Left            =   90
      TabIndex        =   25
      Top             =   90
      Width           =   2190
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   210
         TabIndex        =   26
         Top             =   180
         Width           =   1830
         _ExtentX        =   3228
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   10
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
               Enabled         =   0   'False
               Object.Visible         =   0   'False
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
      Left            =   2445
      TabIndex        =   23
      Top             =   90
      Width           =   1335
      Begin MSComctlLib.Toolbar Toolbar5 
         Height          =   330
         Left            =   210
         TabIndex        =   24
         Top             =   180
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   2
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Cargar Lecturas"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Actualizar Contadores"
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame FrameFiltro 
      Height          =   705
      Left            =   11430
      TabIndex        =   20
      Top             =   90
      Width           =   2415
      Begin VB.ComboBox cboFiltro 
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
         ItemData        =   "frmPOZLecturas.frx":000C
         Left            =   90
         List            =   "frmPOZLecturas.frx":0019
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   210
         Width           =   2235
      End
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
      Height          =   330
      Index           =   7
      Left            =   9660
      MaxLength       =   250
      TabIndex        =   18
      Tag             =   "Consumo|N|N|||rpozos|consumo|0000000|N|"
      Top             =   4410
      Width           =   1095
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
      Height          =   330
      Index           =   6
      Left            =   8490
      MaxLength       =   250
      TabIndex        =   17
      Tag             =   "Numero Orden|N|N|||rpozos|nroorden|0000000|N|"
      Top             =   4410
      Width           =   1095
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
      Left            =   2010
      MaskColor       =   &H00000000&
      TabIndex        =   16
      ToolTipText     =   "Buscar socio"
      Top             =   4380
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
      Left            =   3990
      MaskColor       =   &H00000000&
      TabIndex        =   15
      ToolTipText     =   "Buscar partida"
      Top             =   4410
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
      Index           =   2
      Left            =   7500
      MaskColor       =   &H00000000&
      TabIndex        =   14
      ToolTipText     =   "Buscar fecha"
      Top             =   4410
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.TextBox txtAux 
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
      Height          =   330
      Index           =   3
      Left            =   5700
      MaxLength       =   7
      TabIndex        =   3
      Tag             =   "Lectura Anterior|N|S|||rpozos|lect_ant|0000000||"
      Text            =   "1234567"
      Top             =   4410
      Width           =   765
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
      Height          =   330
      Index           =   4
      Left            =   6480
      MaxLength       =   10
      TabIndex        =   4
      Tag             =   "Fecha Lectura Actual|F|S|||rpozos|fech_act|dd/mm/yyyy||"
      Text            =   "1234567890"
      Top             =   4410
      Width           =   1065
   End
   Begin VB.TextBox txtAux 
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
      Height          =   330
      Index           =   5
      Left            =   7680
      MaxLength       =   7
      TabIndex        =   5
      Tag             =   "Contador Actual|N|S|||rpozos|lect_act|0000000||"
      Text            =   "1234567"
      Top             =   4410
      Width           =   795
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
      Left            =   1260
      MaxLength       =   6
      TabIndex        =   1
      Tag             =   "Socio|N|N|1|999999|rpozos|codsocio|000000||"
      Top             =   4410
      Width           =   705
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
      Height          =   330
      Index           =   2
      Left            =   2220
      MaxLength       =   30
      TabIndex        =   13
      Top             =   4410
      Width           =   1125
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
      Index           =   2
      Left            =   3390
      MaxLength       =   4
      TabIndex        =   2
      Tag             =   "Partida|N|N|1|9999|rpozos|codparti|0000||"
      Top             =   4410
      Width           =   585
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
      Height          =   330
      Index           =   3
      Left            =   4200
      MaxLength       =   40
      TabIndex        =   12
      Top             =   4410
      Width           =   1485
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
      Height          =   330
      Index           =   0
      Left            =   120
      MaxLength       =   250
      TabIndex        =   0
      Tag             =   "Numero Hidrante|T|N|||rpozos|hidrante||S|"
      Top             =   4410
      Width           =   1095
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
      Left            =   13050
      TabIndex        =   6
      Tag             =   "   "
      Top             =   5640
      Visible         =   0   'False
      Width           =   1035
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
      Left            =   14190
      TabIndex        =   7
      Top             =   5625
      Visible         =   0   'False
      Width           =   1095
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmPOZLecturas.frx":0050
      Height          =   4545
      Left            =   120
      TabIndex        =   10
      Top             =   900
      Width           =   15135
      _ExtentX        =   26696
      _ExtentY        =   8017
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
      Left            =   14175
      TabIndex        =   11
      Top             =   5625
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Height          =   600
      Index           =   1
      Left            =   120
      TabIndex        =   8
      Top             =   5505
      Width           =   3285
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
         Height          =   255
         Left            =   45
         TabIndex        =   9
         Top             =   240
         Width           =   3105
      End
   End
   Begin MSAdodcLib.Adodc adodc1 
      Height          =   375
      Left            =   2205
      Top             =   0
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
      Left            =   14760
      TabIndex        =   27
      Top             =   225
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
      Left            =   12330
      TabIndex        =   22
      Top             =   225
      Visible         =   0   'False
      Width           =   1605
   End
   Begin VB.Label Label1 
      Caption         =   "Procesando Registro:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   3465
      TabIndex        =   19
      Top             =   5685
      Visible         =   0   'False
      Width           =   3375
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
         Caption         =   "&Ver todos"
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
         Enabled         =   0   'False
         HelpContextID   =   2
         Shortcut        =   ^E
         Visible         =   0   'False
      End
      Begin VB.Menu mnBarra2 
         Caption         =   "-"
      End
      Begin VB.Menu mnCargaLecturas 
         Caption         =   "&Cargar Lecturas"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnActualizar 
         Caption         =   "&Actualizar Contadores"
      End
      Begin VB.Menu mnBarra3 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnSalir 
         Caption         =   "&Salir"
         Shortcut        =   ^S
      End
   End
   Begin VB.Menu mnFiltro 
      Caption         =   "&Filtro"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu mnFil_Filtro 
         Caption         =   "Con fecha actual"
         Index           =   1
      End
      Begin VB.Menu mnFil_Filtro 
         Caption         =   "Sin fecha actual"
         Index           =   2
      End
      Begin VB.Menu mnFil_Filtro 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnFil_Filtro 
         Caption         =   "Sin filtro"
         Checked         =   -1  'True
         Index           =   4
      End
   End
End
Attribute VB_Name = "frmPOZLecturas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' +-+-+-+-+-+-+-+-+-+-+-+-
' +-+- Autor: MANOLO  +-+-
' +-+-+-+-+-+-+-+-+-+-+-+-

' **************** PER A QUE FUNCIONE EN UN ATRE MANTENIMENT ********************
' 0. Posar-li l'atribut Datasource a "adodc1" del Datagrid1. Canviar el Caption
'    del formulari
' 1. Canviar els TAGs i els Maxlength de TextAux(0) i TextAux(1)
' 2. En PonerModo(vModo) repasar els indexs del botons, per si es canvien
' 3. En la funció BotonAnyadir() canviar la taula i el camp per a SugerirCodigoSiguienteStr
' 4. En la funció BotonBuscar() canviar el nom de la clau primaria
' 5. En la funció BotonEliminar() canviar la pregunta, les descripcions de la
'    variable SQL i el contingut del DELETE
' 6. En la funció PonerLongCampos() posar els camps als que volem canviar el MaxLength quan busquem
' 7. En Form_Load() repasar la barra d'iconos (per si es vol canviar algún) i
'    canviar la consulta per a vore tots els registres
' 8. En Toolbar1_ButtonClick repasar els indexs de cada botó per a que corresponguen
' 9. En la funció CargaGrid canviar l'ORDER BY (normalment per la clau primaria);
'    canviar ademés els noms dels camps, el format i si fa falta la cantitat;
'    repasar els index dels botons modificar i eliminar.
'    NOTA: si en Form_load ya li he posat clausula WHERE, canviar el `WHERE` de
'    `SQL = CadenaConsulta & " WHERE " & vSQL` per un `AND`
' 10. En txtAux_LostFocus canviar el mensage i el format del camp
' 11. En la funció DatosOk() canviar els arguments de DevuelveDesdeBD i el mensage
'    en cas d'error
' 12. En la funció SepuedeBorrar() canviar les comprovacions per a vore si es pot
'    borrar el registre
' *******************************SI N'HI HA COMBO*******************************
' 0. Comprovar que en el SQL de Form_Load() es faça referència a la taula del Combo
' 1. Pegar el Combo1 al  costat dels TextAux. Canviar-li el TAG
' 2. En BotonModificar() canviar el camp del Combo
' 3. En CargaCombo() canviar la consulta i els noms del camps, o posar els valor
'    a ma si no es llig de cap base de datos els valors del Combo

Option Explicit

Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)

Private WithEvents frmPar As frmManPartidas 'partidas
Attribute frmPar.VB_VarHelpID = -1
Private WithEvents frmSoc As frmManSocios 'socios
Attribute frmSoc.VB_VarHelpID = -1
Private WithEvents frmC As frmCal 'calendario fecha
Attribute frmC.VB_VarHelpID = -1


'codi per al registe que s'afegix al cridar des d'atre formulari.
'Obrir en modo Insertar i tornar datos del registre insertat
Public NuevoCodigo As String
Public CodigoActual As String
Public DeConsulta As Boolean

Private CadenaConsulta As String
Private cadB As String

Dim Ordenacion As String

Dim Modo As Byte
'----------- MODOS --------------------------------
'   0.-  Formulari net sense cap camp ple
'   1.-  Preparant per a fer la búsqueda
'   2.-  Ja tenim registres i els anem a recorrer
'        ,podem editar-los Edició del camp
'   3.-  Inserció de nou registre
'   4.-  Modificar
'--------------------------------------------------
Dim PrimeraVez As Boolean
Dim indice As Byte 'Index del text1 on es poses els datos retornats des d'atres Formularis de Mtos
Dim i As Integer

Dim FechaAnt As String
Dim OK As Boolean
Dim CadB1 As String
Dim Filtro As Integer
Dim Sql As String


Dim CadB2 As String

Private Sub PonerModo(vModo)
Dim b As Boolean

    Modo = vModo
    
    b = (Modo = 2)
    If b Then
        PonerContRegIndicador lblIndicador, adodc1, cadB
    Else
        PonerIndicador lblIndicador, Modo
    End If
    
    For i = 0 To txtAux.Count - 1
        txtAux(i).visible = (Modo = 1)
        txtAux(i).Enabled = (Modo = 1)
        txtAux(i).BackColor = vbWhite
    Next i
    
    For i = 4 To 5
        txtAux(i).visible = (Modo = 1 Or Modo = 4)
        txtAux(i).Enabled = (Modo = 1 Or Modo = 4)
    Next i
    
    For i = 0 To Me.btnBuscar.Count - 2
        btnBuscar(i).visible = (Modo = 1)
        btnBuscar(i).Enabled = (Modo = 1)
    Next i
    
    btnBuscar(2).visible = (Modo <> 2)
    btnBuscar(2).Enabled = (Modo <> 2)
    
    Text2(2).visible = (Modo = 1)
    Text2(3).visible = (Modo = 1)
    
    cmdAceptar.visible = Not b
    cmdCancelar.visible = Not b
    DataGrid1.Enabled = b
    
    'Si es regresar
    If DatosADevolverBusqueda <> "" Then cmdRegresar.visible = b
    
    PonerLongCampos
    PonerModoOpcionesMenu 'Activar/Desact botones de menu segun Modo
    PonerOpcionesMenu  'En funcion del usuario
    
    'Si estamos modo Modificar bloquear clave primaria
    BloquearTxt txtAux(0), (Modo = 4)
End Sub


Private Sub PonerModoOpcionesMenu()
'Activa/Desactiva botones del la toobar y del menu, segun el modo en que estemos
Dim b As Boolean

    b = (Modo = 2)
    'Busqueda
    Toolbar1.Buttons(5).Enabled = b
    Me.mnBuscar.Enabled = b
    'Ver Todos
    Toolbar1.Buttons(6).Enabled = b
    Me.mnVerTodos.Enabled = b
    
    'Insertar
    Toolbar1.Buttons(1).Enabled = b And Not DeConsulta
    Me.mnNuevo.Enabled = b And Not DeConsulta
    
    b = (b And adodc1.Recordset.RecordCount > 0) And Not DeConsulta
    'Modificar
    Toolbar1.Buttons(2).Enabled = b
    Me.mnModificar.Enabled = b
    'Eliminar
    Toolbar1.Buttons(3).Enabled = b
    Me.mnEliminar.Enabled = b
    
    FrameBotonGnral2.visible = (vParamAplic.Cooperativa = 10)
    FrameBotonGnral2.Enabled = (vParamAplic.Cooperativa = 10)
    
    'cargarLecturas y actualizar lecturas
    Toolbar5.Buttons(1).Enabled = b And (vParamAplic.Cooperativa = 10) ' solo escalona
    Me.mnCargaLecturas.Enabled = b And (vParamAplic.Cooperativa = 10)
    Toolbar5.Buttons(2).visible = b And (vParamAplic.Cooperativa = 10) ' solo escalona
    Me.mnActualizar.visible = b And (vParamAplic.Cooperativa = 10)
    
End Sub

Private Sub BotonAnyadir()
    Dim NumF As String
    Dim anc As Single
    
    CargaGrid 'primer de tot carregue tot el grid
    cadB = ""
    '******************** canviar taula i camp **************************
    If (DatosADevolverBusqueda <> "") And NuevoCodigo <> "" Then
        NumF = NuevoCodigo
    Else
        NumF = SugerirCodigoSiguienteStr("rpozos", "hidrante")
    End If
    '********************************************************************
    'Situamos el grid al final
    AnyadirLinea DataGrid1, adodc1
         
    anc = DataGrid1.Top
    If DataGrid1.Row < 0 Then
        anc = anc + 206
    Else
        anc = anc + DataGrid1.RowTop(DataGrid1.Row) + 5
    End If
    For i = 0 To txtAux.Count - 1
        txtAux(i).Text = ""
    Next i

    LLamaLineas anc, 3 'Pone el form en Modo=3, Insertar
       
    'Ponemos el foco
    PonerFoco txtAux(0)
End Sub

Private Sub BotonVerTodos()
Dim Sql2 As String
Dim Sql As String

    cadB = AnyadeCadenaFiltro(False)
    
                ' inicio
                conn.Execute "delete from tmpinformes where codusu = " & vUsu.Codigo
            
                Sql2 = "insert into tmpinformes  (codusu, nombre1) select " & vUsu.Codigo & ", hidrante from rpozos where (1=1) "
                If cadB <> "" Then Sql2 = Sql2 & " and " & cadB
                conn.Execute Sql2
                
                ' fin
    
    
    
    CargaGrid "" 'CadB
    PonerModo 2
End Sub

Private Sub BotonBuscar()
    ' ***************** canviar per la clau primaria ********
    CargaGrid "rpozos.hidrante is null"
    '*******************************************************************************
    'Buscar
    For i = 0 To txtAux.Count - 1
        txtAux(i).Text = ""
    Next i
    
    Text2(2).Text = ""
    Text2(3).Text = ""
    
    LLamaLineas DataGrid1.Top + 240, 1 'Pone el form en Modo=1, Buscar
    PonerFoco txtAux(0)
End Sub

Private Sub BotonModificar()
    Dim anc As Single
    Dim i As Integer
    
    Screen.MousePointer = vbHourglass
    
    If DataGrid1.Bookmark < DataGrid1.FirstRow Or DataGrid1.Bookmark > (DataGrid1.FirstRow + DataGrid1.VisibleRows - 1) Then
        i = DataGrid1.Bookmark - DataGrid1.FirstRow
        DataGrid1.Scroll 0, i
        DataGrid1.Refresh
    End If
    
    If DataGrid1.Row < 0 Then
        anc = 320
    Else
        anc = DataGrid1.RowTop(DataGrid1.Row) + DataGrid1.Top 'DataGrid1.RowTop(DataGrid1.Row) + 670 '545
    End If

    'Llamamos al form
    txtAux(0).Text = DataGrid1.Columns(0).Text
    txtAux(1).Text = DataGrid1.Columns(2).Text 'codsocio
    Text2(2).Text = DataGrid1.Columns(3).Text
    txtAux(2).Text = DataGrid1.Columns(4).Text
    Text2(3).Text = DataGrid1.Columns(5).Text
    txtAux(3).Text = DataGrid1.Columns(6).Text
    
    If FechaAnt = "" Then
        txtAux(4).Text = DataGrid1.Columns(7).Text
    Else
        txtAux(4).Text = FechaAnt
    End If
    
    txtAux(5).Text = DataGrid1.Columns(8).Text
    txtAux(6).Text = DataGrid1.Columns(1).Text
    txtAux(7).Text = DataGrid1.Columns(9).Text
    
    ' ***** canviar-ho pel nom del camp del combo *********
'    SelComboBool DataGrid1.Columns(2).Text, Combo1(0)
    ' *****************************************************

    'PosicionarCombo Me.Combo1(0), i
    'PosicionarCombo Me.Combo1(1), i

    LLamaLineas anc, 4 'Pone el form en Modo=4, Modificar
   
    'Como es modificar
    PonerFoco txtAux(4)
    Screen.MousePointer = vbDefault
End Sub

Private Sub LLamaLineas(alto As Single, xModo As Byte)
    DeseleccionaGrid Me.DataGrid1
    PonerModo xModo
    
    'Fijamos el ancho
    For i = 0 To txtAux.Count - 1
        txtAux(i).Top = alto
    Next i
    Text2(2).Top = alto
    Text2(3).Top = alto
    For i = 0 To Me.btnBuscar.Count - 1
        btnBuscar(i).Top = alto
    Next i
    ' ### [Monica] 12/09/2006
    
End Sub


Private Sub BotonCargarLecturas()
Dim Sql As String
Dim temp As Boolean

    On Error GoTo Error2
    'Ciertas comprobaciones
    If adodc1.Recordset.EOF Then Exit Sub

    If Dir(App.Path & "\Escalona", vbDirectory) = "" Then
   
        MsgBox "El directorio de carga de lecturas no existe. Revise.", vbExclamation
    
    Else
        If Dir(App.Path & "\Escalona\escalona.z") = "" Then
            MsgBox "El proceso de carga debe de estar realizandose. Espere.", vbExclamation
        Else
            Sql = "Se va a proceder a realizar la carga de la tabla intermedia. " & vbCrLf & vbCrLf & "¿ Desea continuar ?"
            If MsgBox(Sql, vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
            
                '------------------------------------------------------------------------------
                '  LOG de acciones
                Set LOG = New cLOG
                LOG.Insertar 7, vUsu, "Lectura de contadores de Pozos: " & vbCrLf & vUsu.Codigo & vbCrLf & Now
                Set LOG = Nothing
                '-----------------------------------------------------------------------------
                     
                ' Primero eliminamos todos los registros rpozos_lectura que no tengan la fecha de proceso
                Sql = "delete from rpozos_lectura where fecproceso is null"
                conn.Execute Sql
                    
                ' eliminamos el registro chivato
                Kill App.Path & "\Escalona\escalona.z"
                    
                Shell App.Path & "\Escalona\escalonaconsola ariadna ariadna000 1 v"
            End If
        End If
    End If

    Exit Sub
    
    
Error2:
    Screen.MousePointer = vbDefault
    MuestraError Err.Number, "Cargar Lecturas", Err.Description
End Sub


Private Sub BotonActualizar()
Dim Sql As String
Dim temp As Boolean

    On Error GoTo Error2
    'Ciertas comprobaciones
    If adodc1.Recordset.EOF Then Exit Sub

    If Dir(App.Path & "\Escalona", vbDirectory) = "" Then
   
        MsgBox "El directorio de carga de lecturas no existe. Revise.", vbExclamation
    
    Else
        If Dir(App.Path & "\Escalona\escalona.z", vbDirectory) = "" Then
            Sql = "No se puede realizar una actualización sin que haya realizado la carga."
            MsgBox Sql, vbInformation
        Else
            Sql = "select count(*) from rpozos_lectura where fecproceso is null"
            If TotalRegistros(Sql) = 0 Then
                MsgBox "No hay cargas pendientes de procesar.", vbExclamation
                Exit Sub
            End If
        
            Sql = "Se va a proceder a realizar la actualización de los contadores. " & vbCrLf & vbCrLf
            '[Monica]17/05/2013: indicamos que tipo de lectura se va a actualizar
            ' leemos la lectura de la base de datos
            If vParamAplic.TipoLecturaPoz Then
                Sql = Sql & "Se va a utilizar la LECTURA de la BASE DE DATOS." & vbCrLf & vbCrLf
            Else
                Sql = Sql & "Se va a utilizar la LECTURA del CONTADOR." & vbCrLf & vbCrLf
            End If
            Sql = Sql & "¿ Desea continuar ?"
            If MsgBox(Sql, vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
            
                '------------------------------------------------------------------------------
                '  LOG de acciones
                Set LOG = New cLOG
                LOG.Insertar 7, vUsu, "Actualizacion de contadores de Pozos: " & vbCrLf & vUsu.Codigo & vbCrLf & Now
                Set LOG = Nothing
                '-----------------------------------------------------------------------------
                     
                If ActualizarContadores Then
                    MsgBox "Proceso realizado correctamente.", vbExclamation
                End If
                     
            End If
        End If
    End If
    Exit Sub
    
    
Error2:
    Screen.MousePointer = vbDefault
    MuestraError Err.Number, "Cargar Lecturas", Err.Description
End Sub

Private Function ActualizarContadores() As Boolean
Dim Sql As String, Sql2 As String, Sql3 As String
Dim Rs As ADODB.Recordset, Rs2 As ADODB.Recordset
Dim b As Boolean
Dim Hidrante As String
Dim Inicio As Long
Dim Fin As Long
Dim Limite As Long
Dim Consumo As Long
Dim NroDig As Long

    On Error GoTo eActualizarContadores

    ActualizarContadores = False
    
    conn.BeginTrans
    
    Screen.MousePointer = vbHourglass
    Label1.visible = True
    
    Sql = "select * from rpozos_lectura where fecproceso is null order by hidrante"
    
    b = True
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not Rs.EOF And b
        
        ' SCCHHHTT (Rafa)--> SSHHTT (nuestro)
        '[Monica]13/09/2012: De Rafa me viene el contador con longitud 5 en lugar de con 8 HHHTT --> SSHHTT
        'Hidrante = Right("00" & Mid(DBLet(Rs!Contador), 1, 1), 2) & Mid(DBLet(Rs!Contador), 5, 2) & Mid(DBLet(Rs!Contador), 7, 2)
        Hidrante = Right("00" & Mid(DBLet(Rs!Contador), 1, 1), 2) & Mid(DBLet(Rs!Contador), 2, 4)
        
        Label1.Caption = "Procesando contador: " & Hidrante
        DoEvents
        
        
        Sql2 = "select lect_ant, fech_ant, digcontrol from rpozos where hidrante = " & DBSet(Hidrante, "T")
        Set Rs2 = New ADODB.Recordset
        Rs2.Open Sql2, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        If Not Rs2.EOF Then
            Inicio = 0
            Fin = 0
            
            NroDig = DBLet(Rs2!Digcontrol)
            Limite = 10 ^ NroDig
                 
            Inicio = CLng(ComprobarCero(DBLet(Rs2!lect_ant)))
            
            ' leemos la lectura de la base de datos, la lectura directa del contador puede fallar por comunicacion
            If vParamAplic.TipoLecturaPoz Then
                Fin = CLng(Round2(DBLet(Rs!lectura_bd) / 1000, 0))
            Else
                Fin = CLng(Round2(DBLet(Rs!lectura_equipo) / 1000, 0))
            End If
            
            If Fin >= Inicio Then
              Consumo = Fin - Inicio
            Else
              If MsgBox("¿ Es un reinicio de contador ? ", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
                  Consumo = (Limite - Inicio) + Fin
              Else
                  Consumo = Fin - Inicio
              End If
            End If
            
            If Consumo > Limite - 1 Or Consumo < 0 Then
               MsgBox "Error en la lectura del contador " & Trim(Hidrante) & " . Revise", vbExclamation
               b = False
            Else
                FechaAnt = DBLet(Rs2!fech_ant)
                If FechaAnt = "" Then FechaAnt = "1900-01-01"
                If CDate(DBLet(Rs!fecha_hora)) < FechaAnt Then
                    MsgBox "La fecha de lectura actual es inferior a la de última lectura del contador " & Trim(Hidrante) & " . Revise.", vbExclamation
                    b = False
                End If
            End If
        
            If b Then
                Sql3 = "update rpozos set lect_act = " & DBSet(Fin, "N") & ", fech_act = date(" & DBSet(Rs!fecha_hora, "F") & "), consumo = " & DBSet(Consumo, "N")
                Sql3 = Sql3 & " where hidrante = " & DBSet(Hidrante, "T")
                
                conn.Execute Sql3
                
            End If
            
        End If
        
        ' lo haya o no encontrado el contador lo actualiza en la tabla intermedia
        Sql3 = "update rpozos_lectura set fecproceso =  date(" & DBSet(Rs!fecha_hora, "F") & ") where contador = " & DBSet(Rs!Contador, "T")
        Sql3 = Sql3 & " and id = " & DBSet(Rs!Id, "N")
        conn.Execute Sql3

        Set Rs2 = Nothing
    
        Rs.MoveNext
    Wend
    Set Rs = Nothing
    
    conn.CommitTrans
    ActualizarContadores = True
    Screen.MousePointer = vbDefault
    Label1.visible = False
    DoEvents
    Exit Function

eActualizarContadores:
    Screen.MousePointer = vbDefault
    Label1.visible = False
    DoEvents
    conn.RollbackTrans
    MuestraError Err.Number, "Actualizar contadores", Err.Description
End Function




Private Sub BotonEliminar()
Dim Sql As String
Dim temp As Boolean

    On Error GoTo Error2
    'Ciertas comprobaciones
    If adodc1.Recordset.EOF Then Exit Sub
'    If Not SepuedeBorrar Then Exit Sub
        
    ' ### [Monica] 26/09/2006 dejamos modificar y eliminar el codigo 0
    ' *** repasar el nom de l'adodc, l'index del Field i el camp que te la PK ***
    'El registre de codi 0 no es pot Modificar ni Eliminar
    'If EsCodigoCero(CStr(adodc1.Recordset.Fields(0).Value), FormatoCampo(txtAux(0))) Then Exit Sub
    ' ***************************************************************************
    
    '*************** canviar els noms i el DELETE **********************************
    Sql = "¿Seguro que desea eliminar el Hidrante?"
    Sql = Sql & vbCrLf & "Código: " & adodc1.Recordset.Fields(0)
    
    If MsgBox(Sql, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        NumRegElim = adodc1.Recordset.AbsolutePosition
        
        Sql = "Delete from rpozos where hidrante='" & adodc1.Recordset!Hidrante & "'"
        conn.Execute Sql
        CargaGrid cadB
        
        temp = SituarDataTrasEliminar(adodc1, NumRegElim, True)
        PonerModoOpcionesMenu
        adodc1.Recordset.Cancel
    End If
    Exit Sub
    
Error2:
    Screen.MousePointer = vbDefault
    MuestraError Err.Number, "Eliminando registro", Err.Description
End Sub

Private Sub PonerLongCampos()
'Modificar el MaxLength del campo en funcion de si es modo de búsqueda o no
'para los campos que permitan introducir criterios más largos del tamaño del campo
'Si maxlength=6 para codprove, en modo busqueda Maxlenth=13 'el doble + 1
    
    'para los txtAux
    PonerLongCamposGnral Me, Modo, 3
End Sub

Private Sub btnBuscar_Click(Index As Integer)
    Select Case Index
        Case 0 'socios
            Set frmSoc = New frmManSocios
'            frmSoc.DeConsulta = True
            frmSoc.DatosADevolverBusqueda = "0|1|"
'            frmSoc.CodigoActual = Text1(1).Text
            frmSoc.Show vbModal
            Set frmSoc = Nothing
            PonerFoco txtAux(1)
        Case 1 'partida
            Set frmPar = New frmManPartidas
            frmPar.DeConsulta = True
            frmPar.DatosADevolverBusqueda = "0|1|"
            frmPar.CodigoActual = txtAux(2).Text
            frmPar.Show vbModal
            Set frmPar = Nothing
            PonerFoco txtAux(2)
            
        Case 2 ' fecha
            Dim esq As Long
            Dim dalt As Long
            Dim menu As Long
            Dim obj As Object
        
            Set frmC = New frmCal
            
            indice = Index
            
            esq = btnBuscar(Index).Left
            dalt = btnBuscar(Index).Top
                
            Set obj = btnBuscar(Index).Container
              
              While btnBuscar(Index).Parent.Name <> obj.Name
                    esq = esq + obj.Left
                    dalt = dalt + obj.Top
                    Set obj = obj.Container
              Wend
            
            menu = Me.Height - Me.ScaleHeight 'ací tinc el heigth del menú i de la toolbar
        
            frmC.Left = esq + btnBuscar(Index).Parent.Left + 30
            frmC.Top = dalt + btnBuscar(Index).Parent.Top + btnBuscar(Index).Height + menu - 40
        
            btnBuscar(Index).Tag = Index '<===
            ' *** repasar si el camp es txtAux o Text1 ***
            If txtAux(4).Text <> "" Then frmC.NovaData = txtAux(4).Text
            
            ' ********************************************
        
            frmC.Show vbModal
            Set frmC = Nothing
            ' *** repasar si el camp es txtAux o Text1 ***
            PonerFoco txtAux(4) '<===
            ' ********************************************
            
    End Select
    
End Sub

Private Sub cboFiltro_Change()
    Filtro = Me.cboFiltro.ListIndex
    AbrirFicheroFiltro False
End Sub

Private Sub cboFiltro_Click()
    Filtro = Me.cboFiltro.ListIndex
    AbrirFicheroFiltro False
End Sub

Private Sub cboFiltro_Validate(Cancel As Boolean)
    Filtro = Me.cboFiltro.ListIndex
    AbrirFicheroFiltro False
End Sub

Private Sub cmdAceptar_Click()
    Dim i As String
    Dim NReg As Long
    Dim Sql As String
    Dim Sql2 As String
    
    Select Case Modo
        Case 1 'BUSQUEDA
            cadB = ObtenerBusqueda(Me)
            If cadB <> "" Then
            
                ' inicio
                conn.Execute "delete from tmpinformes where codusu = " & vUsu.Codigo
                Sql2 = "insert into tmpinformes  (codusu, nombre1) select " & vUsu.Codigo & ", hidrante from rpozos where " & cadB & AnyadeCadenaFiltro(True)
                conn.Execute Sql2
                ' fin
                
                CargaGrid "" ' CadB & AnyadeCadenaFiltro(True)
                PonerModo 2
'                lblIndicador.Caption = "BUSQUEDA: " & PonerContRegistros(Me.adodc1)
                PonerFocoGrid Me.DataGrid1
            End If
            
        Case 3 'INSERTAR
            If DatosOk Then
                If InsertarDesdeForm(Me) Then
                    CargaGrid
                    If (DatosADevolverBusqueda <> "") And NuevoCodigo <> "" Then
                        cmdCancelar_Click
'                        If Not adodc1.Recordset.EOF Then adodc1.Recordset.MoveLast
                        If Not adodc1.Recordset.EOF Then
                            adodc1.Recordset.Find (adodc1.Recordset.Fields(0).Name & " =" & NuevoCodigo)
                        End If
                        cmdRegresar_Click
                    Else
                        BotonAnyadir
                    End If
                    cadB = ""
                End If
            End If
            
        Case 4 'MODIFICAR
            
            OK = False
            If DatosOk Then
                If ModificaDesdeFormulario2(Me, 0) Then
                    OK = True
                
                
                    FechaAnt = txtAux(4).Text
                    TerminaBloquear
                    i = adodc1.Recordset.Fields(0)
                    PonerModo 2
                    CargaGrid "" 'CadB
                    adodc1.Recordset.Find (adodc1.Recordset.Fields(0).Name & " ='" & i & "'")
                    PonerFocoGrid Me.DataGrid1
                    
                    
                End If
            End If
    End Select
End Sub

Private Sub cmdCancelar_Click()
    On Error Resume Next
    
    Select Case Modo
        Case 1 'búsqueda
            CargaGrid cadB
        Case 3 'insertar
            DataGrid1.AllowAddNew = False
            'CargaGrid
            If Not adodc1.Recordset.EOF Then adodc1.Recordset.MoveFirst
        Case 4 'modificar
            TerminaBloquear
    End Select
    
    PonerModo 2
    
    PonerFocoGrid Me.DataGrid1
    If Err.Number <> 0 Then Err.Clear
End Sub

Private Sub cmdRegresar_Click()
Dim Cad As String
Dim i As Integer
Dim J As Integer
Dim Aux As String

    If adodc1.Recordset.EOF Then
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
            Cad = Cad & adodc1.Recordset.Fields(J) & "|"
        End If
    Loop Until i = 0
    RaiseEvent DatoSeleccionado(Cad)
    Unload Me
End Sub

Private Sub DataGrid1_HeadClick(ByVal ColIndex As Integer)
Dim Cad As String

If adodc1.Recordset Is Nothing Then Exit Sub
If adodc1.Recordset.EOF Then Exit Sub

Me.Refresh
DoEvents

Screen.MousePointer = vbHourglass

Ordenacion = "ORDER BY " & DataGrid1.Columns(0).DataField

'ColIndexAnt = ColIndex
cadB = AnyadeCadenaFiltro(False)
CargaGrid cadB

Screen.MousePointer = vbDefault
End Sub

Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If Modo = 2 Then PonerContRegIndicador lblIndicador, adodc1, cadB
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault

    If PrimeraVez Then
        PrimeraVez = False
        If (DatosADevolverBusqueda <> "") And NuevoCodigo <> "" Then
            BotonAnyadir
        Else
            PonerModo 2
             If Me.CodigoActual <> "" Then
                SituarData Me.adodc1, "hidrante='" & CodigoActual & "'", "", True
            End If
        End If
    End If
End Sub

Private Sub Form_Load()
Dim Sql2 As String

    'Icono del formulario
    Me.Icon = frmPpal.Icon

    PrimeraVez = True


    CargaFiltros


    LeerFiltro True
    PonerFiltro Filtro

'    With Me.Toolbar1
'        .HotImageList = frmPpal.imgListComun_OM
'        .DisabledImageList = frmPpal.imgListComun_BN
'        .ImageList = frmPpal.imgListComun
'        'el 1 es separadors
'        .Buttons(2).Image = 1   'Buscar
'        .Buttons(3).Image = 2   'Todos
'        'el 4 i el 5 son separadors
'        .Buttons(6).Image = 3   'Insertar
'        .Buttons(7).Image = 4   'Modificar
'        .Buttons(8).Image = 5   'Borrar
'        'el 9 i el 10 son separadors
'        .Buttons(11).Image = 34  'cargar de consola para escalona
'        .Buttons(12).Image = 35  'actualizar contadores
'        .Buttons(13).Image = 11  'Salir
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
        .Buttons(1).Image = 34  'cargar de consola para escalona
        .Buttons(2).Image = 35  'actualizar contadores
    End With



    '## A mano
'    chkVistaPrevia.Value = CheckValueLeer(Name)
    
    cadB = AnyadeCadenaFiltro(False)
    
    ' inicio
    conn.Execute "delete from tmpinformes where codusu = " & vUsu.Codigo
    Sql2 = "insert into tmpinformes  (codusu, nombre1) select " & vUsu.Codigo & ", hidrante from rpozos where 1=1 "
    If cadB <> "" Then Sql2 = Sql2 & " and " & cadB
    conn.Execute Sql2
    ' fin
    
    
    '****************** canviar la consulta *********************************+
    CadenaConsulta = "SELECT rpozos.hidrante, rpozos.nroorden, rpozos.codsocio, rsocios.nomsocio,"
    CadenaConsulta = CadenaConsulta & "rpozos.codparti, rpartida.nomparti, rpozos.lect_ant, "
    CadenaConsulta = CadenaConsulta & "rpozos.fech_act, rpozos.lect_act, rpozos.consumo "
    CadenaConsulta = CadenaConsulta & " FROM ((rpozos INNER JOIN rsocios ON rpozos.codsocio = rsocios.codsocio) "
    CadenaConsulta = CadenaConsulta & " INNER JOIN rpartida ON rpozos.codparti = rpartida.codparti)"
    CadenaConsulta = CadenaConsulta & " INNER JOIN tmpinformes ON tmpinformes.nombre1 = rpozos.hidrante and codusu =" & vUsu.Codigo
    CadenaConsulta = CadenaConsulta & " WHERE rpozos.fechabaja is null "
    '************************************************************************
    
    Ordenacion = " ORDER BY nroorden "
    
    
    cadB = ""
    CargaGrid
    
    FechaAnt = ""
    
'    If (DatosADevolverBusqueda <> "") And NuevoCodigo <> "" Then
'        BotonAnyadir
'    Else
'        PonerModo 2
'    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
'    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
    Screen.MousePointer = vbDefault
    
    LeerFiltro False
    
    If Modo = 4 Then TerminaBloquear
End Sub

Private Sub frmC_Selec(vFecha As Date)
    txtAux(4).Text = Format(vFecha, "dd/mm/yyyy") '<===
End Sub

Private Sub frmPar_DatoSeleccionado(CadenaSeleccion As String)
    txtAux(2).Text = RecuperaValor(CadenaSeleccion, 1) 'codigo de partida
    FormateaCampo txtAux(2)
    Text2(3).Text = RecuperaValor(CadenaSeleccion, 2) 'nombre de partida
End Sub

Private Sub frmSoc_DatoSeleccionado(CadenaSeleccion As String)
    txtAux(1).Text = RecuperaValor(CadenaSeleccion, 1) 'codsocio
    FormateaCampo txtAux(1)
    Text2(2).Text = RecuperaValor(CadenaSeleccion, 2) 'nomsocio
End Sub

Private Sub mnActualizar_Click()
    BotonActualizar
End Sub

Private Sub mnBuscar_Click()
    BotonBuscar
End Sub

Private Sub mnCargaLecturas_Click()
    BotonCargarLecturas
End Sub

Private Sub mnEliminar_Click()
    BotonEliminar
End Sub

Private Sub mnFil_Filtro_Click(Index As Integer)
    For i = 1 To mnFil_Filtro.Count
        If i <> 3 Then mnFil_Filtro(i).Checked = False
    Next i
    mnFil_Filtro(Index).Checked = True
    If Index = 4 Then
        Filtro = 0
    Else
        Filtro = Index
    End If
    AbrirFicheroFiltro False
End Sub

Private Sub mnImprimir_Click()
    printNou
End Sub

Private Sub mnModificar_Click()
    'Comprobaciones
    '--------------
    If adodc1.Recordset.EOF Then Exit Sub
    
    If adodc1.Recordset.RecordCount < 1 Then Exit Sub
    
    ' ### [Monica] 26/09/2006 dejamos modificar y eliminar el codigo 0
    ' *** repasar el nom de l'adodc, l'index del Field i el camp que te la PK ***
    'El registre de codi 0 no es pot Modificar ni Eliminar
    'If EsCodigoCero(CStr(adodc1.Recordset.Fields(0).Value), FormatoCampo(txtAux(0))) Then Exit Sub
    
    
    'Preparamos para modificar
    '-------------------------
    If BLOQUEADesdeFormulario2(Me, adodc1, 1) Then BotonModificar
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
        Case 5
            mnBuscar_Click
        Case 6
            mnVerTodos_Click
        Case 1
            mnNuevo_Click
        Case 2
            mnModificar_Click
        Case 3
            mnEliminar_Click
    End Select
End Sub

Private Sub Toolbar5_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1
            mnCargaLecturas_Click
        Case 2
            mnActualizar_Click
    End Select
End Sub
Private Sub CargaGrid(Optional vSQL As String)
    Dim Sql As String
    Dim tots As String
    Dim Sql2 As String
    
'    adodc1.ConnectionString = Conn
    If vSQL <> "" Then
        Sql = CadenaConsulta & " AND " & vSQL
    Else
        Sql = CadenaConsulta
    End If
    '********************* canviar el ORDER BY *********************++
'        Sql = Sql & " ORDER BY rpozos.nroorden"
    Sql = Sql & " " & Ordenacion
    '**************************************************************++
    
    
    CargaGridGnral Me.DataGrid1, Me.adodc1, Sql, PrimeraVez
    
    ' *******************canviar els noms i si fa falta la cantitat********************
    tots = "S|txtAux(0)|T|Hidrante|1000|;S|txtAux(6)|T|Orden|1000|;S|txtAux(1)|T|Socio|1000|;S|btnBuscar(0)|B||195|;S|Text2(2)|T|Nombre|3000|;"
    tots = tots & "S|txtAux(2)|T|Código|800|;S|btnBuscar(1)|B||195|;S|Text2(3)|T|Partida|2500|;"
    tots = tots & "S|txtAux(3)|T|Lect.Anterior|1450|;S|txtAux(4)|T|F.Lect.Actual|1450|;S|btnBuscar(2)|B||195|;"
    tots = tots & "S|txtAux(5)|T|Lect.Actual|1250|;S|txtAux(7)|T|Consumo|1100|;"
    'N||||0|;
    arregla tots, DataGrid1, Me, 350
    
    DataGrid1.ScrollBars = dbgAutomatic
    DataGrid1.Columns(0).Alignment = dbgLeft
    DataGrid1.Columns(2).Alignment = dbgLeft
    DataGrid1.Columns(4).Alignment = dbgLeft
    DataGrid1.Columns(6).Alignment = dbgCenter
    DataGrid1.Columns(7).Alignment = dbgLeft
    DataGrid1.Columns(8).Alignment = dbgCenter
    
'   DataGrid1.Columns(2).Alignment = dbgRight
End Sub

Private Sub txtAux_GotFocus(Index As Integer)
    ConseguirFoco txtAux(Index), Modo
End Sub


Private Sub txtAux_LostFocus(Index As Integer)
    If Not PerderFocoGnral(txtAux(Index), Modo) Then Exit Sub
    
    Select Case Index
        Case 1, 2 ' <1> = socio <2> = partida
            PonerFormatoEntero txtAux(Index)
        
        Case 3, 5 ' lectura anterior / lectura actual
            PonerFormatoEntero txtAux(Index)
             
        Case 4 ' fecha de lectura actual
            '[Monica]28/08/2013: no comprobamos que la fecha esté en la campaña
            PonerFormatoFecha txtAux(Index)
    End Select
    
End Sub

Private Function DatosOk() As Boolean
'Dim Datos As String
Dim b As Boolean
Dim Sql As String
Dim Mens As String
Dim FechaAnt As Date
Dim NroDig As Integer
Dim Inicio As Long
Dim Fin As Long
Dim Consumo As Long
Dim Limite As Long

    b = CompForm(Me)
    If Not b Then Exit Function
    
    If Modo = 3 Then   'Estamos insertando
         If ExisteCP(txtAux(0)) Then b = False
    End If
    
    If b And Modo = 4 Then
        If txtAux(5).Text <> "" Then
             Inicio = 0
             Fin = 0
             NroDig = DevuelveValor("select digcontrol from rpozos where hidrante = " & DBSet(txtAux(0).Text, "T"))
             Limite = 10 ^ NroDig
             
             If txtAux(3).Text <> "" Then Inicio = CLng(txtAux(3).Text)
             If txtAux(5).Text <> "" Then Fin = CLng(txtAux(5).Text)
             
             
'             If Fin = 0 Then
'                 Consumo = 0
'             Else
                 If Fin >= Inicio Then
                    Consumo = Fin - Inicio
                 Else
                    If MsgBox("¿ Es un reinicio de contador ? ", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
                        Consumo = (Limite - Inicio) + Fin
                    Else
                        Consumo = Fin - Inicio
                    End If
                 End If
'             End If
             
             If Consumo > Limite - 1 Or Consumo < 0 Then
                MsgBox "Error en la lectura. Revise", vbExclamation
                PonerFoco txtAux(5)
                b = False
             Else
                
                If txtAux(4).Text = "" Then
                    MsgBox "La fecha de lectura debe tener un valor. Revise.", vbExclamation
                    PonerFoco txtAux(4)
                    b = False
                Else
                    FechaAnt = DevuelveValor("select fech_ant from rpozos where hidrante = " & DBSet(txtAux(0).Text, "T"))
                    If CDate(txtAux(4).Text) < FechaAnt Then
                        MsgBox "La fecha de lectura actual es inferior a la de última lectura. Revise.", vbExclamation
                        PonerFoco txtAux(4)
                        b = False
                    End If
                End If
             End If
        Else
            If txtAux(5).Text = "" And txtAux(4).Text = "" Then
                txtAux(7).Text = ""
                b = True
            Else
                b = False
            End If
        End If
    
    End If
    
    If b Then txtAux(7).Text = Consumo
    
    
    DatosOk = b
End Function

Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
    PonerOpcionesMenuGeneralNew Me
End Sub


Private Sub printNou()
    With frmImprimir2
        .cadTabla2 = "rcapataz"
        .Informe2 = "rManCapataz.rpt"
        If cadB <> "" Then
            '.cadRegSelec = Replace(SQL2SF(CadB), "clientes", "clientes_1")
            .cadRegSelec = SQL2SF(cadB)
        Else
            .cadRegSelec = ""
        End If
        ' *** repasar el nom de l'adodc ***
        '.cadRegActua = Replace(POS2SF(Data1, Me), "clientes", "clientes_1")
        .cadRegActua = POS2SF(adodc1, Me)
        '[Monica]13/07/2012: falla si hay un solo registro seleccionado y apretamos registros buscados
        If adodc1.Recordset.RecordCount = 1 Then .cadRegSelec = .cadRegActua
        
        ' *** repasar codEmpre ***
        .cadTodosReg = ""
        '.cadTodosReg = "{itinerar.codempre} = " & codEmpre
        ' *** repasar si li pose ordre o no ****
        .OtrosParametros2 = "pEmpresa='" & vEmpresa.nomempre & "'|pOrden={rcapataz.codcapat}|"
        '.OtrosParametros2 = "pEmpresa='" & vEmpresa.NomEmpre & "'|"
        ' *** posar el nº de paràmetres que he posat en OtrosParametros2 ***
        '.NumeroParametros2 = 1
        .NumeroParametros2 = 2
        ' ******************************************************************
        .MostrarTree2 = False
        .InfConta2 = False
        .ConSubInforme2 = False
        .SubInformeConta = ""
        .Show vbModal
    End With
End Sub

'' ### [DavidV] 26/04/2006: Activar/desactivar la rueda del ratón.
'Private Sub DataGrid1_GotFocus()
'  WheelHook DataGrid1
'End Sub
'Private Sub DataGrid1_Lostfocus()
'  WheelUnHook
'End Sub

Private Sub txtAux_KeyPress(Index As Integer, KeyAscii As Integer)

    If Index = 5 Then ' estoy introduciendo la lectura
       If KeyAscii = 13 Then 'ENTER
            PonerFormatoEntero txtAux(Index)
            If Modo = 4 Then
                '050509 cmdAceptar_Click 'ModificarExistencia
                cmdAceptar_Click
                'ModificarLinea

                If OK Then PasarSigReg
                    
            End If
            If Modo = 1 Or Modo = 3 Then
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

End Sub


Private Sub txtAux_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo EKeyD
    
'Alvançar/Retrocedir els camps en les fleches de desplaçament del teclat.
'    KEYdown KeyCode
    
    
    ' si no estamos en muestra salimos
    If Index <> 5 Then Exit Sub
    
    Select Case KeyCode
        Case 38 'Desplazamieto Fecha Hacia Arriba
'050509
            cmdAceptar_Click
'            ModificarLinea
            
'            If Me.DataGridAux(0).Bookmark > 0 Then
'                DataGridAux(0).Bookmark = DataGridAux(0).Bookmark - 1
'            End If
            If OK Then PasarAntReg
        Case 40 'Desplazamiento Flecha Hacia Abajo
            'ModificarExistencia
'050509
            cmdAceptar_Click
'            ModificarLinea
            
            If OK Then PasarSigReg
    End Select
EKeyD:
    If Err.Number <> 0 Then Err.Clear
End Sub

Private Sub PasarSigReg()
'Nos situamos en el siguiente registro
    If Me.DataGrid1.Bookmark < Me.adodc1.Recordset.RecordCount Then
'        DataGridAux(0).Row = DataGridAux(0).Row + 1
        DataGrid1.Bookmark = DataGrid1.Bookmark + 1
        BotonModificar
        PonerFoco txtAux(5)
    ElseIf DataGrid1.Bookmark = adodc1.Recordset.RecordCount Then
'        PonerFocoBtn Me.cmdAceptar
        BotonModificar
        PonerFoco txtAux(5)
    End If
End Sub


Private Sub PasarAntReg()
'Nos situamos en el siguiente registro
    If Me.DataGrid1.Bookmark > 1 Then
'        DataGridAux(0).Row = DataGridAux(0).Row + 1
        DataGrid1.Bookmark = DataGrid1.Bookmark - 1
        BotonModificar
        PonerFoco txtAux(5)
    ElseIf DataGrid1.Bookmark = 1 Then
'        PonerFocoBtn Me.cmdAceptar
        BotonModificar
        PonerFoco txtAux(5)
    End If
End Sub




Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me

End Sub


Private Sub LeerFiltro(Leer As Boolean)
    Sql = App.Path & "\filtro.dat"
    If Leer Then
        Filtro = 0
        If Dir(Sql) <> "" Then
            AbrirFicheroFiltro True
            If IsNumeric(Sql) Then Filtro = CByte(Sql)
        End If
    Else
        AbrirFicheroFiltro False
    End If
End Sub


Private Sub AbrirFicheroFiltro(Leer As Boolean)
On Error GoTo EAbrir
    i = FreeFile
    If Leer Then
        Open Sql For Input As #i
        Sql = "0"
        Line Input #i, Sql
    Else
        Open Sql For Output As #i
        Print #i, Filtro
    End If
    Close #i
    Exit Sub
EAbrir:
    Err.Clear
End Sub


Private Sub PonerFiltro(NumFilt As Integer)
    Filtro = NumFilt
'    Me.mnFil_Filtro(1).Checked = (NumFilt = 1)
'    Me.mnFil_Filtro(2).Checked = (NumFilt = 2)
'    Me.mnFil_Filtro(4).Checked = (NumFilt = 0)
    PosicionarCombo Me.cboFiltro, NumFilt
End Sub


Private Function AnyadeCadenaFiltro(Anyade As Boolean) As String
Dim Aux As String

    Aux = ""
    If Filtro <> 0 Then ' si hay filtro
        If Anyade Then Aux = " and "
        If Filtro = 1 Then
            Aux = Aux & " not fech_act is null "
        Else
            Aux = Aux & " fech_act is null "
        End If
        
    End If  'filtro=0
    AnyadeCadenaFiltro = Aux
End Function


Private Sub CargaFiltros()
Dim Aux As String

    cboFiltro.Clear
    
    cboFiltro.AddItem "Sin Filtro "
    cboFiltro.ItemData(cboFiltro.NewIndex) = 0
    cboFiltro.AddItem "Con fecha actual "
    cboFiltro.ItemData(cboFiltro.NewIndex) = 1
    cboFiltro.AddItem "Sin fecha actual "
    cboFiltro.ItemData(cboFiltro.NewIndex) = 2

End Sub
    
