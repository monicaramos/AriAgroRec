VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmContrRecFactPre 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Asignacion de Precios"
   ClientHeight    =   8265
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8820
   Icon            =   "frmContrRecFactPre.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8265
   ScaleWidth      =   8820
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtAux2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   2550
      TabIndex        =   17
      Top             =   2010
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame FrameAux1 
      Caption         =   "Albaranes"
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
      Height          =   4440
      Left            =   90
      TabIndex        =   12
      Top             =   2880
      Width           =   8520
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   45
         Left            =   6780
         MaxLength       =   10
         TabIndex        =   22
         Tag             =   "Kilos|N|S|||rhisfruta_clasif|kilosnet|##,###,##0||"
         Text            =   "Kilos"
         Top             =   1290
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   44
         Left            =   7500
         MaxLength       =   9
         TabIndex        =   21
         Tag             =   "Precio|N|S|||rhisfruta_clasif|precio|#,##0.0000||"
         Text            =   "Precio"
         Top             =   1290
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtAux2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   1440
         TabIndex        =   20
         Text            =   "fecha"
         Top             =   1260
         Visible         =   0   'False
         Width           =   420
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   43
         Left            =   4590
         MaxLength       =   6
         TabIndex        =   19
         Tag             =   "Calidad|N|N|||rhisfruta_clasif|codcalid|00|S|"
         Text            =   "Calida"
         Top             =   1290
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtAux2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Index           =   43
         Left            =   5400
         TabIndex        =   18
         Text            =   "nombre"
         Top             =   1290
         Visible         =   0   'False
         Width           =   2040
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   41
         Left            =   540
         MaxLength       =   7
         TabIndex        =   15
         Tag             =   "Albaran|N|N|||rhisfruta_clasif|numalbar|0000000|S|"
         Text            =   "albaran"
         Top             =   1260
         Visible         =   0   'False
         Width           =   780
      End
      Begin VB.TextBox txtAux2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Index           =   42
         Left            =   2610
         TabIndex        =   14
         Text            =   "nombre"
         Top             =   1290
         Visible         =   0   'False
         Width           =   2040
      End
      Begin VB.TextBox txtAux 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   42
         Left            =   1620
         MaxLength       =   6
         TabIndex        =   13
         Tag             =   "Variedad|N|N|||rhisfruta_clasif|codvarie|000000|S|"
         Text            =   "Varied"
         Top             =   1260
         Visible         =   0   'False
         Width           =   735
      End
      Begin MSComctlLib.Toolbar ToolAux 
         Height          =   390
         Index           =   1
         Left            =   120
         TabIndex        =   16
         Top             =   240
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
         Left            =   1350
         Top             =   660
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
         Bindings        =   "frmContrRecFactPre.frx":000C
         Height          =   3195
         Index           =   1
         Left            =   120
         TabIndex        =   23
         Top             =   720
         Width           =   8160
         _ExtentX        =   14393
         _ExtentY        =   5636
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
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   2
      Left            =   7110
      MaxLength       =   9
      TabIndex        =   2
      Tag             =   "Precio|N|S|||rhisfruta|precio|#,##0.0000||"
      Top             =   2010
      Width           =   1125
   End
   Begin VB.TextBox txtAux2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   285
      Index           =   2
      Left            =   5220
      TabIndex        =   11
      Top             =   2010
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   1
      Left            =   3720
      MaxLength       =   4
      TabIndex        =   1
      Tag             =   "Calidad|N|N|||rhisfruta|codcalid|00||"
      Top             =   2010
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   6315
      TabIndex        =   3
      Top             =   7725
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   7440
      TabIndex        =   4
      Top             =   7725
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   0
      Left            =   1260
      MaxLength       =   6
      TabIndex        =   0
      Tag             =   "Variedad|N|N|||rhisfruta|codvarie|000000|S|"
      Top             =   2010
      Visible         =   0   'False
      Width           =   1395
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmContrRecFactPre.frx":0024
      Height          =   2190
      Left            =   60
      TabIndex        =   7
      Top             =   570
      Width           =   8520
      _ExtentX        =   15028
      _ExtentY        =   3863
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
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   7440
      TabIndex        =   10
      Top             =   7740
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Height          =   675
      Index           =   1
      Left            =   120
      TabIndex        =   5
      Top             =   7410
      Width           =   2385
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
         Height          =   255
         Left            =   40
         TabIndex        =   6
         Top             =   240
         Width           =   2295
      End
   End
   Begin MSAdodcLib.Adodc adodc1 
      Height          =   375
      Left            =   2790
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
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   8820
      _ExtentX        =   15558
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   12
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Buscar"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Ver todos"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Nuevo"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Modificar"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Eliminar"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Imprimir"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   195
         Left            =   5040
         TabIndex        =   9
         Top             =   90
         Visible         =   0   'False
         Width           =   1215
      End
   End
   Begin VB.Menu mnOpciones 
      Caption         =   "&Opciones"
      Begin VB.Menu mnBuscar 
         Caption         =   "&Buscar"
         Enabled         =   0   'False
         Shortcut        =   ^F
         Visible         =   0   'False
      End
      Begin VB.Menu mnVerTodos 
         Caption         =   "&Ver todos"
         Enabled         =   0   'False
         Shortcut        =   ^B
         Visible         =   0   'False
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
      Begin VB.Menu mnImprimir 
         Caption         =   "&Imprimir"
         Enabled         =   0   'False
         Shortcut        =   ^I
         Visible         =   0   'False
      End
      Begin VB.Menu mnSalir 
         Caption         =   "&Salir"
         Shortcut        =   ^S
      End
   End
End
Attribute VB_Name = "frmContrRecFactPre"
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

'codi per al registe que s'afegix al cridar des d'atre formulari.
'Obrir en modo Insertar i tornar datos del registre insertat
Public NuevoCodigo As String
Public CodigoActual As String
Public DeConsulta As Boolean
Public Precios As String
Public Albaranes As String
Public Precio As String


Private CadenaConsulta As String
Private CadB As String

Private WithEvents frmBas As frmBasico ' Ayuda de Areas
Attribute frmBas.VB_VarHelpID = -1
Private WithEvents frmCtas As frmCtasConta 'cuentas contables
Attribute frmCtas.VB_VarHelpID = -1


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
Dim I As Integer
Dim indCodigo As Integer

Dim NumTabMto As Integer 'Indica quin nº de Tab està en modo Mantenimient

Dim ModoLineas As Byte
'1.- Afegir,  2.- Modificar,  3.- Borrar,  0.-Passar control a Llínies

Private Sub PonerModo(vModo)
Dim b As Boolean

    Modo = vModo

    b = (Modo = 2) Or (Modo = 5)
    If b Then
        PonerContRegIndicador
    Else
        PonerIndicador lblIndicador, Modo, ModoLineas
    End If
    
    For I = 0 To 2 'txtAux.Count - 1
        txtaux(I).visible = False
    Next I
    
    txtaux(2).visible = (Modo = 4)
    txtaux(2).Enabled = (Modo = 4)
    
    txtaux(44).visible = (Modo = 5)
    txtaux(44).Enabled = (Modo = 5)
    
    cmdAceptar.visible = Modo = 4 Or Modo = 5
    cmdCancelar.visible = Modo = 4 Or Modo = 5
    
    cmdAceptar.Enabled = Modo = 4 Or Modo = 5
    cmdCancelar.Enabled = Modo = 4 Or Modo = 5
    DataGrid1.Enabled = b
    
    'Si es regresar
    If DatosADevolverBusqueda <> "" Then cmdRegresar.visible = b
    
    PonerLongCampos
    PonerModoOpcionesMenu 'Activar/Desact botones de menu segun Modo
    PonerOpcionesMenu  'En funcion del usuario
    
    
    'Si estamos modo Modificar bloquear clave primaria
    BloquearTxt txtaux(0), (Modo = 4)
End Sub


Private Sub PonerModoOpcionesMenu()
'Activa/Desactiva botones del la toobar y del menu, segun el modo en que estemos
Dim b As Boolean
Dim bAux As Boolean

    b = (Modo = 2)
    'Busqueda
    Toolbar1.Buttons(2).Enabled = b
    Me.mnBuscar.Enabled = b
    'Ver Todos
    Toolbar1.Buttons(3).Enabled = b
    Me.mnVerTodos.Enabled = b
    
    'Insertar
    Toolbar1.Buttons(6).Enabled = b And Not DeConsulta
    Me.mnNuevo.Enabled = b And Not DeConsulta
    
    b = (b And adodc1.Recordset.RecordCount > 0) And Not DeConsulta
    'Modificar
    Toolbar1.Buttons(7).Enabled = b
    Me.mnModificar.Enabled = b
    'Eliminar
    Toolbar1.Buttons(8).Enabled = b
    Me.mnEliminar.Enabled = b
    'Imprimir
    Toolbar1.Buttons(11).Enabled = b
    Me.mnImprimir.Enabled = b
    
    b = (Modo = 4 Or Modo = 2) And Not DeConsulta
    For I = 1 To 1 'ToolAux.Count - 1
        ToolAux(I).Buttons(1).Enabled = b
        If b Then bAux = (b And Me.AdoAux(I).Recordset.RecordCount > 0)
'        ToolAux(i).Buttons(2).Enabled = bAux
        ToolAux(I).Buttons(3).Enabled = bAux
    Next I
    
    
    
End Sub

Private Sub BotonAnyadir()
    Dim NumF As String
    Dim anc As Single
    
    CargaGrid 'primer de tot carregue tot el grid
    CadB = ""
    '******************** canviar taula i camp **************************
    If (DatosADevolverBusqueda <> "") And NuevoCodigo <> "" Then
        NumF = NuevoCodigo
    Else
        NumF = SugerirCodigoSiguienteStr("rhisfruta", "codcoste")
    End If
    
    PonerModo 3
    
    '********************************************************************
    'Situamos el grid al final
    AnyadirLinea DataGrid1, adodc1
         
    anc = DataGrid1.Top
    If DataGrid1.Row < 0 Then
        anc = anc + 206
    Else
        anc = anc + DataGrid1.RowTop(DataGrid1.Row) + 5
    End If
    txtaux(0).Text = NumF
    FormateaCampo txtaux(0)
    For I = 1 To 5
        txtaux(I).Text = ""
    Next I
    txtAux2(2).Text = ""

    LLamaLineas anc, 3 'Pone el form en Modo=3, Insertar
       
    'Ponemos el foco
    PonerFoco txtaux(0)
End Sub

Private Sub BotonVerTodos()
    CadB = ""
    CargaGrid ""
    PonerModo 2
End Sub

Private Sub BotonBuscar()
    ' ***************** canviar per la clau primaria ********
    CargaGrid "rhisfruta.codcoste = -1"
    CargaGridAux 1, False

    '*******************************************************************************
    'Buscar
    For I = 0 To 8 'txtAux.Count - 1
        txtaux(I).Text = ""
    Next I
    txtAux2(2).Text = ""
    txtAux2(4).Text = ""
    
    LLamaLineas DataGrid1.Top + 206, 1 'Pone el form en Modo=1, Buscar
    PonerFoco txtaux(0)
End Sub

Private Sub BotonModificar()
    Dim anc As Single
    Dim I As Integer
    
    Screen.MousePointer = vbHourglass
    
    If DataGrid1.Bookmark < DataGrid1.FirstRow Or DataGrid1.Bookmark > (DataGrid1.FirstRow + DataGrid1.VisibleRows - 1) Then
        I = DataGrid1.Bookmark - DataGrid1.FirstRow
        DataGrid1.Scroll 0, I
        DataGrid1.Refresh
    End If
    
    If DataGrid1.Row < 0 Then
        anc = 320
    Else
        anc = DataGrid1.RowTop(DataGrid1.Row) + 545
    End If

    'Llamamos al form
    txtaux(0).Text = DataGrid1.Columns(0).Text
    txtAux2(0).Text = DataGrid1.Columns(1).Text
    txtaux(1).Text = DataGrid1.Columns(2).Text
    txtAux2(2).Text = DataGrid1.Columns(3).Text
    txtaux(2).Text = DataGrid1.Columns(4).Text
    
    
    LLamaLineas anc, 4 'Pone el form en Modo=4, Modificar
   
    'Como es modificar
    PonerFoco txtaux(2)
    Screen.MousePointer = vbDefault
End Sub

Private Sub LLamaLineas(alto As Single, xModo As Byte)
    DeseleccionaGrid Me.DataGrid1
    PonerModo xModo
    

    'Fijamos el ancho
    For I = 0 To 2
        txtaux(I).Top = alto
    Next I

    ' ### [Monica] 12/09/2006
    txtAux2(0).Top = alto
    txtAux2(2).Top = alto

End Sub

Private Sub BotonEliminar()
Dim sql As String
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
    sql = "¿Seguro que desea eliminar el Concepto de Coste?"
    sql = sql & vbCrLf & "Código: " & adodc1.Recordset.Fields(0)
    sql = sql & vbCrLf & "Descripción: " & adodc1.Recordset.Fields(1)
    
    If MsgBox(sql, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        NumRegElim = adodc1.Recordset.AbsolutePosition
        
        sql = "Delete from rhisfruta_clasif where codcoste=" & adodc1.Recordset!codCoste
        conn.Execute sql
        
        sql = "Delete from rhisfruta_cta where codcoste=" & adodc1.Recordset!codCoste
        conn.Execute sql
        
        sql = "Delete from rhisfruta where codcoste=" & adodc1.Recordset!codCoste
        conn.Execute sql
        CargaGrid CadB
        
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
 TerminaBloquear
    
    Select Case Index
        Case 0 'area
            
            indice = Index + 2
            Set frmBas = New frmBasico
            frmBas.DatosADevolverBusqueda = "0|1|"
            frmBas.DeConsulta = True
            frmBas.CodigoActual = txtaux(indice).Text
            frmBas.CadenaTots = "S|txtAux(0)|T|Código|800|;S|txtAux(1)|T|Descripción|3930|;"
            frmBas.CadenaConsulta = "SELECT ccareas.codarea, ccareas.nomarea "
            frmBas.CadenaConsulta = frmBas.CadenaConsulta & " FROM ccareas "
            frmBas.CadenaConsulta = frmBas.CadenaConsulta & " WHERE (1=1) "
            frmBas.Tag1 = "Código|N|N|0|9999|ccareas|codarea|0000|S|"
            frmBas.Tag2 = "Descripción|T|N|||ccareas|nomarea|||"
            frmBas.Maxlen1 = 4
            frmBas.Maxlen2 = 50
            frmBas.Tabla = "ccareas"
            frmBas.CampoCP = "codarea"
            frmBas.Report = "rManCCAreas.rpt"
            frmBas.Caption = "Áreas"
            
            frmBas.Show vbModal
            Set frmBas = Nothing
            PonerFoco txtaux(indice)
    
            
    
    
    End Select
    
    If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Me.adodc1, 1
End Sub


Private Sub cmdAceptar_Click()
    Dim I As Long
    Dim J As Long

    Select Case Modo
        Case 1 'BUSQUEDA
            CadB = ObtenerBusqueda2(Me, , 1)
            If CadB <> "" Then
                CargaGrid CadB
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
                        PosicionarData
                        Me.DataGrid1.AllowAddNew = False
                        BotonAnyadirLinea 1
                    End If
                    CadB = ""
                End If
            End If
            
        Case 4 'MODIFICAR
            If DatosOk Then
                If ModificaRegistro Then
                    TerminaBloquear
                    
                    I = adodc1.Recordset.Fields(0)
                    PonerModo 2
                    CargaGrid CadB
                    
                    adodc1.Recordset.Find (adodc1.Recordset.Fields(0).Name & " =" & I)
                    PonerFocoGrid Me.DataGrid1
                End If
            End If
            
        Case 5 'LLÍNIES
            Select Case ModoLineas
                Case 1 'afegir llínia
                    InsertarLinea
                Case 2 'modificar llínies
                    If ModificarLinea Then
'                        PosicionarData
'                        cmdCancelar_Click
                        PonerModo 2
                    Else
                        PonerFoco txtaux(5)
                    End If
            End Select
            'nuevo calculamos los totales de lineas
            
    End Select
End Sub

Private Sub cmdCancelar_Click()
Dim V

    On Error Resume Next
    
    Select Case Modo
        Case 1 'búsqueda
            CargaGrid CadB
        Case 3 'insertar
            DataGrid1.AllowAddNew = False
            If Not adodc1.Recordset.EOF Then adodc1.Recordset.MoveFirst
            PonerModo 2
            DataGrid1_RowColChange 1, 0
        Case 4 'modificar
            TerminaBloquear
        Case 5 'LLÍNIES
            Select Case ModoLineas
                Case 1 'afegir llínia
                
                    ModoLineas = 0
                    ' *** les llínies que tenen datagrid (en o sense tab) ***
                    DataGridAux(NumTabMto).AllowAddNew = False
                    ' **** repasar si es diu Data1 l'adodc de la capçalera ***
                    'If BLOQUEADesdeFormulario2(Me, Data1, 1) Then BotonModificar 'Modificar
                    LLamaLineas2 NumTabMto, ModoLineas 'ocultar txtAux
                    DataGridAux(NumTabMto).Enabled = True
                    DataGridAux(NumTabMto).SetFocus

                    ' *** si n'hi han camps de descripció dins del grid, els neteje ***
                    'txtAux2(2).text = ""
                    If Not AdoAux(NumTabMto).Recordset.EOF Then
                        AdoAux(NumTabMto).Recordset.MoveFirst
                    End If

                Case 2 'modificar llínies
                    ModoLineas = 0
                    
                    ' *** si n'hi han tabs ***
'                    SituarTab (NumTabMto + 1)
                    LLamaLineas 1, ModoLineas 'ocultar txtAux
                    PonerModo 4
                    If Not AdoAux(1).Recordset.EOF Then
                        ' *** l'Index de Fields es el que canvie de la PK de llínies ***
                        V = AdoAux(1).Recordset.Fields(0) 'el 2 es el nº de llinia
                        AdoAux(1).Recordset.Find (AdoAux(1).Recordset.Fields(0).Name & " =" & V)
                        ' ***************************************************************
                    End If
            End Select
            
            TerminaBloquear
    End Select
    
    PonerModo 2
    
    PonerFocoGrid Me.DataGrid1
    If Err.Number <> 0 Then Err.Clear
End Sub

Private Sub cmdRegresar_Click()
Dim cad As String
Dim I As Integer
Dim J As Integer
Dim Aux As String

    If adodc1.Recordset.EOF Then
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
            cad = cad & adodc1.Recordset.Fields(J) & "|"
        End If
    Loop Until I = 0
    RaiseEvent DatoSeleccionado(cad)
    Unload Me
End Sub



Private Sub DataGrid1_DblClick()
    If cmdRegresar.visible Then cmdRegresar_Click
End Sub

Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    PonerContRegIndicador
'    CargaForaGrid
 
    If Modo = 3 Then
        CargaGridAux 1, False
    Else
        CargaGridAux 1, True
    End If
    
    
    PonerModoOpcionesMenu
    PonerOpcionesMenu

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
                SituarData Me.adodc1, "codcoste=" & CodigoActual, "", True
            End If
        End If
    End If
End Sub

Private Sub Form_Load()
    PrimeraVez = True

    With Me.Toolbar1
        .HotImageList = frmPpal.imgListComun_OM
        .DisabledImageList = frmPpal.imgListComun_BN
        .ImageList = frmPpal.imgListComun
        'el 1 es separadors
        .Buttons(2).Image = 1   'Buscar
        .Buttons(3).Image = 2   'Todos
        'el 4 i el 5 son separadors
        .Buttons(6).Image = 3   'Insertar
        .Buttons(7).Image = 4   'Modificar
        .Buttons(8).Image = 5   'Borrar
        'el 9 i el 10 son separadors
        .Buttons(11).Image = 10  'imprimir
        .Buttons(12).Image = 11  'Salir
    End With

    '## A mano
'    chkVistaPrevia.Value = CheckValueLeer(Name)
    
    
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
'    LimpiarCampos   'Neteja els camps TextBox
    ' ******* si n'hi han llínies *******
    DataGridAux(1).ClearFields
    
    LimpiarPrecios
    
    '****************** canviar la consulta *********************************+
    CadenaConsulta = "SELECT distinct rhisfruta_clasif.codvarie, variedades.nomvarie, rhisfruta_clasif.codcalid, rcalidad.nomcalid, rhisfruta_clasif.precio "
    CadenaConsulta = CadenaConsulta & " FROM rhisfruta_clasif, variedades, rcalidad "
    CadenaConsulta = CadenaConsulta & " WHERE rhisfruta_clasif.codvarie = variedades.codvarie "
    CadenaConsulta = CadenaConsulta & " and rhisfruta_clasif.codcalid = rcalidad.codcalid and rhisfruta_clasif.codvarie = rcalidad.codvarie"
    CadenaConsulta = CadenaConsulta & " and rhisfruta_clasif.numalbar in ( " & Albaranes & ")"
    '************************************************************************
    
    CadB = ""
    CargaGrid CadB
    
    ' ****** Si n'hi han camps fora del grid ******
'    CargaForaGrid
    ' *********************************************
    CargaGridAux 1, True

'    If (DatosADevolverBusqueda <> "") And NuevoCodigo <> "" Then
'        BotonAnyadir
'    Else
'        PonerModo 2
'    End If

    ModoLineas = 0

End Sub

Private Sub LimpiarPrecios()
Dim sql As String

    sql = "update rhisfruta set prestimado = " & DBSet(Precio, "N")
    sql = sql & " where numalbar in (" & Albaranes & ")"

    conn.Execute sql
    
    sql = "update rhisfruta_clasif set precio = " & DBSet(Precio, "N")
    sql = sql & " where numalbar in (" & Albaranes & ")"
    
    conn.Execute sql
    
End Sub


Private Sub Form_Unload(Cancel As Integer)
'    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
    If Modo = 4 Then TerminaBloquear
    Screen.MousePointer = vbDefault
End Sub

Private Sub frmBas_DatoSeleccionado(CadenaSeleccion As String)
'Ayuda de Areas de Costes
    txtaux(indice).Text = Format(RecuperaValor(CadenaSeleccion, 1), "0000") 'codarea
    txtAux2(indice).Text = RecuperaValor(CadenaSeleccion, 2) 'nombre de area
End Sub


Private Sub frmCtas_DatoSeleccionado(CadenaSeleccion As String)
'Cuentas contables de la Contabilidad
    txtaux(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1) 'codmacta
    txtAux2(4).Text = RecuperaValor(CadenaSeleccion, 2) 'des macta
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
        Case 2
                mnBuscar_Click
        Case 3
                mnVerTodos_Click
        Case 6
                mnNuevo_Click
        Case 7
                mnModificar_Click
        Case 8
                mnEliminar_Click
        Case 11
                'MsgBox "Imprimir...under construction"
                mnImprimir_Click
        Case 12
                mnSalir_Click
    End Select
End Sub

Private Sub CargaGrid(Optional vSQL As String)
    Dim sql As String
    Dim tots As String
    
'    adodc1.ConnectionString = Conn
    If vSQL <> "" Then
        sql = CadenaConsulta & " AND " & vSQL
    Else
        sql = CadenaConsulta
    End If
    '********************* canviar el ORDER BY *********************++
    sql = sql & " ORDER BY rhisfruta_clasif.codvarie, rhisfruta_clasif.codcalid "
    '**************************************************************++
    
    CargaGridGnral Me.DataGrid1, Me.adodc1, sql, PrimeraVez
    
    ' *******************canviar els noms i si fa falta la cantitat********************
    tots = "S|txtAux(0)|T|Código|800|;S|txtAux2(0)|T|Variedad|2575|;"
    tots = tots & "S|txtAux(1)|T|Código|800|;S|txtAux2(1)|T|Calidad|2275|;S|txtAux(2)|T|Precio|1500|;"
    
    arregla tots, DataGrid1, Me
    
    DataGrid1.ScrollBars = dbgAutomatic
    DataGrid1.Columns(0).Alignment = dbgRight
    DataGrid1.Columns(2).Alignment = dbgLeft
'    DataGrid1.Columns(4).Alignment = dbgLeft
    
End Sub

Private Sub txtAux_GotFocus(Index As Integer)
    ConseguirFocoLin txtaux(Index)
End Sub


Private Sub txtAux_LostFocus(Index As Integer)
    If Not PerderFocoGnral(txtaux(Index), Modo) Then Exit Sub
    
    Select Case Index
        Case 2, 44 ' precios
            If PonerFormatoDecimal(txtaux(Index), 7) Then cmdAceptar.SetFocus
    End Select
    
End Sub

Private Function DatosOk() As Boolean
'Dim Datos As String
Dim b As Boolean
Dim sql As String
Dim Mens As String
Dim TipoCoste As Byte


    b = CompForm(Me)
    If Not b Then Exit Function
    
    If Modo = 3 Then   'Estamos insertando
         If ExisteCP(txtaux(0)) Then b = False
    End If
    
    
    DatosOk = b
End Function


Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
    PonerOpcionesMenuGeneralNew Me
End Sub


Private Sub PonerContRegIndicador()
'si estamos en modo ver registros muestra el numero de registro en el que estamos
'situados del total de registros mostrados: 1 de 24
Dim cadReg As String

    If (Modo = 2 Or Modo = 0) Then
        cadReg = PonerContRegistros(Me.adodc1)
        If CadB = "" Then
            lblIndicador.Caption = cadReg
        Else
            lblIndicador.Caption = "BUSQUEDA: " & cadReg
        End If
    End If
End Sub

Private Sub printNou()
    With frmImprimir2
        .cadTabla2 = "rhisfruta"
        .Informe2 = "rManCCConceptos.rpt"
        If CadB <> "" Then
            '.cadRegSelec = Replace(SQL2SF(CadB), "clientes", "clientes_1")
            .cadRegSelec = SQL2SF(CadB)
        Else
            .cadRegSelec = ""
        End If
        ' *** repasar el nom de l'adodc ***
        '.cadRegActua = Replace(POS2SF(Data1, Me), "clientes", "clientes_1")
        .cadRegActua = POS2SF(adodc1, Me)
        ' *** repasar codEmpre ***
        .cadTodosReg = ""
        '.cadTodosReg = "{itinerar.codempre} = " & codEmpre
        ' *** repasar si li pose ordre o no ****
        .OtrosParametros2 = "pEmpresa='" & vEmpresa.nomempre & "'|pOrden={rhisfruta.codcoste}|"
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

'Private Sub txtAux_KeyPress(Index As Integer, KeyAscii As Integer)
'    KEYpress KeyAscii
'End Sub
Private Sub txtAux_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = teclaBuscar Then
        If Modo = 1 Or Modo = 3 Or Modo = 4 Then
            Select Case Index
                Case 2: KEYBusqueda KeyAscii, 0 'cuenta contable
            End Select
        End If
    Else
        KEYpress KeyAscii
    End If
End Sub

Private Sub txtAux_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Alvançar/Retrocedir els camps en les fleches de desplaçament del teclat.
    KEYdown KeyCode
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
    btnBuscar_Click (indice)
End Sub



'****************************************************
Private Sub CargaGridAux(Index As Integer, enlaza As Boolean)
Dim b As Boolean
Dim I As Byte
Dim tots As String

    On Error GoTo ECarga

    tots = MontaSQLCarga(Index, enlaza)

    CargaGridGnral Me.DataGridAux(Index), Me.AdoAux(Index), tots, PrimeraVez
    
    Select Case Index
        Case 1 'lineas
            
'            sql = "SELECT rhisfruta_clasif.numalbar, rhisfruta.fecalbar, rhisfruta_clasif.codvarie, variedades.nomvarie, rhisfruta_clasif.codcalid, rcalidad.nomcalid, rhisfruta_clasif.precio"
            
            
            'si es visible|control|tipo campo|nombre campo|ancho control|
            tots = "S|txtAux(41)|T|Albaran|820|;S|txtAux2(1)|T|Fecha|1000|;S|txtAux(42)|T|Codigo|800|;"
            tots = tots & "S|txtAux2(42)|T|Nombre|1220|;S|txtAux(43)|T|Codigo|700|;S|txtAux2(43)|T|Calidad|1220|;S|txtAux(45)|T|Kilos|800|;S|txtAux(44)|T|Precio|1020|;"
            
            arregla tots, DataGridAux(Index), Me
            
            DataGridAux(1).Columns(2).Alignment = dbgLeft
            
            b = (Modo = 4) And ((ModoLineas = 1) Or (ModoLineas = 2))
            
    End Select
    
    DataGridAux(Index).ScrollBars = dbgAutomatic
      
ECarga:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando datos grid: " & DataGridAux(Index).Tag, Err.Description
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
Dim sql As String
Dim Tabla As String
    
    ' ********* si n'hi han tabs, dona igual si en datagrid o no ***********
    Select Case Index
        Case 1 ' LINEAS
            sql = "SELECT rhisfruta_clasif.numalbar, rhisfruta.fecalbar, rhisfruta_clasif.codvarie, variedades.nomvarie, rhisfruta_clasif.codcalid, rcalidad.nomcalid, rhisfruta_clasif.kilosnet, rhisfruta_clasif.precio"
            sql = sql & " FROM rhisfruta_clasif, rhisfruta, variedades, rcalidad "
            sql = sql & " where rhisfruta.numalbar = rhisfruta_clasif.numalbar "
            sql = sql & " and rhisfruta_clasif.codvarie = variedades.codvarie "
            sql = sql & " and rhisfruta_clasif.codvarie = rcalidad.codvarie "
            sql = sql & " and rhisfruta_clasif.codcalid = rcalidad.codcalid "
            sql = sql & " and rhisfruta_clasif.numalbar in (" & Albaranes & ")"
            
            sql = sql & " ORDER BY rhisfruta_clasif.numalbar, rhisfruta_clasif.codvarie, rhisfruta_clasif.codcalid "
    
    End Select
    
    MontaSQLCarga = sql
End Function

Private Function ObtenerWhereCab(conW As Boolean) As String
Dim vWhere As String
    
    vWhere = ""
    If conW Then vWhere = " WHERE "
    ' *** canviar-ho per la clau primaria de la capçalera ***
    vWhere = vWhere & " codcoste=" & Me.adodc1.Recordset!codCoste
    
    ObtenerWhereCab = vWhere
End Function




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
Dim sql As String
Dim vWhere As String
Dim eliminar As Boolean

    On Error GoTo Error2

    ModoLineas = 3 'Posem Modo Eliminar Llínia
    
    If Modo = 4 Then 'Modificar Capçalera
        cmdAceptar_Click
        If ModoLineas = 0 Then Exit Sub
    End If
       
'    PonerModo 5

    If AdoAux(Index).Recordset.EOF Then Exit Sub
    eliminar = False
   
    vWhere = ObtenerWhereCab(True)
    
    ' ***** independentment de si tenen datagrid o no,
    ' canviar els noms, els formats i el DELETE *****
    Select Case Index
        Case 1 'lineas de confeccion
            sql = "¿ Seguro que desea eliminar la linea ?"
            sql = sql & vbCrLf & "Linea: " & AdoAux(Index).Recordset!codlinea '& " " & Adoaux(Index).Recordset!NomTraba
            If MsgBox(sql, vbQuestion + vbYesNo) = vbYes Then
                eliminar = True
                sql = "DELETE FROM rhisfruta_clasif "
                sql = sql & vWhere & " AND numlinea= " & AdoAux(Index).Recordset!numlinea
            End If
            
        Case 2 'lineas de cuentas contables
            sql = "¿ Seguro que desea eliminar la cuenta asociada ?"
            sql = sql & vbCrLf & "Cuenta Contable: " & AdoAux(Index).Recordset!ctacontable '& " " & Adoaux(Index).Recordset!NomTraba
            If MsgBox(sql, vbQuestion + vbYesNo) = vbYes Then
                eliminar = True
                sql = "DELETE FROM rhisfruta_cta "
                sql = sql & vWhere & " AND numlinea= " & AdoAux(Index).Recordset!numlinea
            End If
    End Select

    If eliminar Then
        NumRegElim = AdoAux(Index).Recordset.AbsolutePosition
        TerminaBloquear
        conn.Execute sql
        
        ' *** si n'hi han tabs sense datagrid, posar l'If ***
        CargaGridAux Index, True
        If Not SituarDataTrasEliminar(AdoAux(Index), NumRegElim, True) Then
            
        End If
    End If
    
    ModoLineas = 0
    
    Exit Sub
Error2:
    Screen.MousePointer = vbDefault
    MuestraError Err.Number, "Eliminando linea", Err.Description
End Sub

Private Sub BotonAnyadirLinea(Index As Integer)
Dim NumF As String
Dim vWhere As String, vTabla As String
Dim anc As Single
Dim I As Integer
    
    
    NumTabMto = Index
    
    ModoLineas = 1 'Posem Modo Afegir Llínia
    
    If (Modo = 3) Or (Modo = 4) Then 'Insertar o Modificar Capçalera
        cmdAceptar_Click
        If ModoLineas = 0 Then Exit Sub
    End If
       
    PonerModo 5
    
    ' *** bloquejar la clau primaria de la capçalera ***
'    BloquearTxt text1(0), True

    ' *** posar el nom del les distintes taules de llínies ***
    Select Case Index
        Case 1: vTabla = "rhisfruta_clasif"
        Case 2: vTabla = "rhisfruta_cta"
    End Select
    
    vWhere = ObtenerWhereCab(False)
    
    Select Case Index
        Case 1, 2 ' *** pose els index dels tabs de llínies que tenen datagrid ***
            ' *** canviar la clau primaria de les llínies,
            'pasar a "" si no volem que mos sugerixca res a l'afegir ***
            
            NumF = SugerirCodigoSiguienteStr(vTabla, "numlinea", vWhere)

            AnyadirLinea DataGridAux(Index), AdoAux(Index)
    
            anc = DataGridAux(Index).Top
            If DataGridAux(Index).Row < 0 Then
                anc = anc + 210
            Else
                anc = anc + DataGridAux(Index).RowTop(DataGridAux(Index).Row) + 5
            End If
            
            LLamaLineas2 Index, ModoLineas, anc
        
            Select Case Index
                ' *** valor per defecte a l'insertar i formateig de tots els camps ***
                Case 1 'lineas de confeccion
                    txtaux(40).Text = Me.adodc1.Recordset!codCoste  'codorden
                    txtaux(41).Text = NumF 'numlinea
                    txtaux(42).Text = ""
                    
                    BloquearbtnBuscar Me, Modo, ModoLineas, "FrameAux1"
                    
                    PonerFoco txtaux(42)
            End Select
            
    End Select
End Sub

Private Sub BotonModificarLinea(Index As Integer)
    Dim anc As Single
    Dim I As Integer
    Dim J As Integer
    
    If AdoAux(Index).Recordset.EOF Then Exit Sub
    If AdoAux(Index).Recordset.RecordCount < 1 Then Exit Sub
    
    ModoLineas = 2 'Modificar llínia
       
    If Modo = 4 Then 'Modificar Capçalera
        cmdAceptar_Click
        If ModoLineas = 0 Then Exit Sub
    End If
       
    PonerModo 5
    ' *** bloqueje la clau primaria de la capçalera ***
  
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
                anc = anc + DataGridAux(Index).RowTop(DataGridAux(Index).Row) + 5
            End If
    End Select
    
    Select Case Index
        ' *** valor per defecte al modificar dels camps del grid ***
        Case 1 ' precios de calidades de albaranes
            
            txtaux(41).Text = DataGridAux(Index).Columns(0).Text
            txtAux2(1).Text = DataGridAux(Index).Columns(1).Text
            txtaux(42).Text = DataGridAux(Index).Columns(2).Text
            txtAux2(42).Text = DataGridAux(Index).Columns(3).Text
            txtaux(43).Text = DataGridAux(Index).Columns(4).Text
            txtAux2(43).Text = DataGridAux(Index).Columns(5).Text
            txtaux(45).Text = DataGridAux(Index).Columns(6).Text
            txtaux(44).Text = DataGridAux(Index).Columns(7).Text
            
    End Select
    
    LLamaLineas2 Index, ModoLineas, anc
   
    ' *** foco al 1r camp visible de les llinies en grids que no siga PK (en o sense tab) ***
    Select Case Index
        Case 1 'lineas de conceptos
            PonerFoco txtaux(44)
    End Select
    ' ***************************************************************************************
End Sub

Private Sub LLamaLineas2(Index As Integer, xModo As Byte, Optional alto As Single)
Dim jj As Integer
Dim b As Boolean

    ' *** si n'hi han tabs sense datagrid posar el If ***
    DeseleccionaGrid DataGridAux(Index)
       
    b = (xModo = 1 Or xModo = 2) 'Insertar o Modificar Llínies
    Select Case Index
        Case 1 'lineas de confeccion
             For jj = 44 To 44
                txtaux(jj).visible = b
                txtaux(jj).Top = alto
            Next jj
    End Select
End Sub

Private Sub InsertarLinea()
'Inserta registre en les taules de Llínies
Dim nomframe As String
Dim b As Boolean

    On Error Resume Next

    ' *** posa els noms del frames, tant si son de grid com si no ***
    
    Select Case NumTabMto
        Case 1
            nomframe = "FrameAux1" 'lineas de conceptos
    End Select
    
    
    If DatosOkLlin(nomframe) Then
        
        TerminaBloquear
        
        If InsertarDesdeForm2(Me, 2, nomframe) Then
            b = BLOQUEADesdeFormulario2(Me, Me.adodc1, 1)
            ' *** els index de les llinies en grid (en o sense tab) ***
                CargaGridAux NumTabMto, True
                If b Then BotonAnyadirLinea NumTabMto
        End If
    End If
    
End Sub

Private Function ModificarLinea() As Boolean
'Modifica registre en les taules de Llínies
Dim nomframe As String
Dim V As Integer
    
    On Error Resume Next

    ' *** posa els noms del frames, tant si son de grid com si no ***
    nomframe = "FrameAux1" 'lineas de confeccion
    
    ModificarLinea = False
    If DatosOkLlin(nomframe) Then
        TerminaBloquear
        If ModificaDesdeFormulario2(Me, 2, nomframe) Then
            ModoLineas = 0
            
'            V = AdoAux(1).Recordset.Fields(0) 'el 2 es el nº de llinia
            CargaGridAux 1, True
            
            ' *** si n'hi han tabs ***

            ' *** si n'hi han tabs que no tenen datagrid, posar el if ***
            PonerFocoGrid Me.DataGridAux(1)
            
'            AdoAux(1).Recordset.Find (AdoAux(1).Recordset.Fields(0).Name & " =" & V)
            
'            LLamaLineas 0, 0
'            PosicionarData

            ModificarLinea = True
        End If
    End If
End Function


Private Function ModificaRegistro() As Boolean
'Modifica registre en les taules de Llínies
Dim nomframe As String
Dim V As Integer
Dim sql As String

    
    On Error GoTo eModificaRegistro


    ModificaRegistro = False
    conn.BeginTrans

    ' *** posa els noms del frames, tant si son de grid com si no ***
    nomframe = "FrameAux1" 'lineas de confeccion
    
    
    sql = "update rhisfruta set prestimado = " & DBSet(txtaux(2).Text, "N")
    sql = sql & " where codvarie = " & DBSet(txtaux(0).Text, "N")
    sql = sql & " and numalbar in (" & Albaranes & ")"
    
    conn.Execute sql
    
    If MsgBox("¿ Desea actualizar los precios de todas las calidades de los albaranes ?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
        sql = "update rhisfruta_clasif set precio = " & DBSet(txtaux(2).Text, "N")
        sql = sql & " where codvarie = " & DBSet(txtaux(0).Text, "N")
        sql = sql & " and codcalid = " & DBSet(txtaux(1).Text, "N")
        sql = sql & " and numalbar in (" & Albaranes & ")"
        
        conn.Execute sql
    End If
    
    CargaGridAux 1, True
    
    ' *** si n'hi han tabs que no tenen datagrid, posar el if ***
    PonerFocoGrid Me.DataGridAux(0)
'            AdoAux(0).Recordset.Find (AdoAux(0).Recordset.Fields(1).Name & " =" & V)
            
    LLamaLineas 0, 0
    
    ModificaRegistro = True
    conn.CommitTrans
    Exit Function
    
eModificaRegistro:
    conn.RollbackTrans
End Function




Private Sub PosicionarData()
Dim cad As String, Indicador As String

    ' *** canviar-ho per tota la PK de la capçalera, no llevar els () ***
    cad = "rhisfruta_clasif.codvarie=" & DBSet(txtaux(42).Text, "N") & " and rhisfruta_clasif.codcalid = " & DBSet(txtaux(43).Text, "N") & " and rhisfruta_clasif.numalbar = " & DBSet(txtaux(41).Text, "N")
    
    
    ' *** gastar SituarData o SituarDataMULTI depenent de si la PK es simple o composta ***
    'If SituarDataMULTI(Data1, cad, Indicador) Then
    If SituarDataMULTI(Me.AdoAux(1), cad, Indicador) Then
        If ModoLineas <> 1 Then PonerModo 2
        lblIndicador.Caption = Indicador
    Else
       LimpiarCampos
       PonerModo 0
    End If
End Sub


Private Function DatosOkLlin(nomframe As String) As Boolean
Dim Rs As ADODB.Recordset
Dim sql As String
Dim b As Boolean
Dim cant As Integer
Dim Mens As String
Dim vFact As Byte, vDocum As Byte

    On Error GoTo EDatosOKLlin

    Mens = ""
    DatosOkLlin = False
        
    b = CompForm2(Me, 2, nomframe) 'Comprovar formato datos ok
    If Not b Then Exit Function
    
    If b And nomframe = "FrameAux1" Then
        'comprobamos que no exista ese codigo de linea
        If ModoLineas = 1 Then
            sql = "select count(*) from rhisfruta_clasif where codcoste = " & Me.adodc1.Recordset!codCoste
            sql = sql & " and codlinea = " & DBSet(txtaux(42).Text, "N")
            If TotalRegistros(sql) <> 0 Then
                MsgBox "Código de linea ya existe. Revise.", vbExclamation
                b = False
            End If
        End If
    End If
    
    If b And nomframe = "FramAux2" Then
        'comprobamos que no exista ese codigo de linea
        If ModoLineas = 1 Then
            sql = "select count(*) from rhisfruta_cta where codcoste = " & Me.adodc1.Recordset!codCoste
            sql = sql & " and ctacontable = " & DBSet(txtaux(52).Text, "T")
            If TotalRegistros(sql) <> 0 Then
                MsgBox "Cuenta contable ya existe. Revise.", vbExclamation
                b = False
            End If
        End If
    
    End If
    
    ' ******************************************************************************
    DatosOkLlin = b
    
EDatosOKLlin:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function


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


