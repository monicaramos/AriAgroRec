VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmManPreClasifica 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Entradas para Clasificación"
   ClientHeight    =   9855
   ClientLeft      =   45
   ClientTop       =   30
   ClientWidth     =   13770
   Icon            =   "frmManPreClasifica.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9855
   ScaleWidth      =   13770
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Left            =   7200
      TabIndex        =   24
      Top             =   210
      Visible         =   0   'False
      Width           =   1605
   End
   Begin VB.Frame FrameBotonGnral2 
      Height          =   705
      Left            =   3780
      TabIndex        =   22
      Top             =   30
      Width           =   1575
      Begin MSComctlLib.Toolbar Toolbar2 
         Height          =   330
         Left            =   210
         TabIndex        =   23
         Top             =   180
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   2
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Actualizar Entradas"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Importar Fichero Excel"
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame FrameBotonGnral 
      Height          =   705
      Left            =   120
      TabIndex        =   20
      Top             =   30
      Width           =   3585
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   210
         TabIndex        =   21
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
   Begin VB.CheckBox chkAux 
      BackColor       =   &H80000005&
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
      Left            =   10620
      TabIndex        =   19
      Tag             =   "Procesado|N|N|||rclasifica_imp|procesado|||"
      Top             =   4980
      Visible         =   0   'False
      Width           =   255
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
      Index           =   5
      Left            =   2340
      MaxLength       =   6
      TabIndex        =   2
      Tag             =   "Código Variedad|N|N|0|999999|rclasifica_imp|codvarie|000000||"
      Top             =   4920
      Width           =   465
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
      Index           =   3
      Left            =   2880
      MaskColor       =   &H00000000&
      TabIndex        =   18
      ToolTipText     =   "Buscar variedad"
      Top             =   4920
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
      Index           =   5
      Left            =   3090
      TabIndex        =   17
      Top             =   4950
      Visible         =   0   'False
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
      Index           =   6
      Left            =   8730
      MaxLength       =   7
      TabIndex        =   5
      Tag             =   "Cajas|N|N|||rclasifica_imp|numcajon|###,##0||"
      Top             =   4950
      Width           =   885
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
      Index           =   2
      Left            =   8430
      MaskColor       =   &H00000000&
      TabIndex        =   16
      ToolTipText     =   "Buscar campos"
      Top             =   4950
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
      Height          =   350
      Index           =   1
      Left            =   2085
      MaskColor       =   &H00000000&
      TabIndex        =   15
      ToolTipText     =   "Buscar fecha"
      Top             =   4905
      Visible         =   0   'False
      Width           =   195
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
      Index           =   4
      Left            =   9660
      MaxLength       =   8
      TabIndex        =   6
      Tag             =   "Kilos|N|S|||rclasifica_imp|kilosnet|###,##0||"
      Top             =   4950
      Width           =   900
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
      Index           =   3
      Left            =   -330
      MaxLength       =   8
      TabIndex        =   0
      Tag             =   "Num.Nota|N|N|||rclasifica_imp|numnotac|######0|S|"
      Top             =   4920
      Width           =   990
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
      Index           =   0
      Left            =   5355
      TabIndex        =   14
      Top             =   4935
      Visible         =   0   'False
      Width           =   1815
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
      Left            =   5145
      MaskColor       =   &H00000000&
      TabIndex        =   13
      ToolTipText     =   "Buscar socio"
      Top             =   4920
      Visible         =   0   'False
      Width           =   195
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
      Left            =   7230
      MaxLength       =   16
      TabIndex        =   4
      Tag             =   "Codigo Campo|T|N|||rclasifica_imp|codcampo||N|"
      Top             =   4935
      Width           =   1155
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
      Left            =   11355
      TabIndex        =   7
      Top             =   9150
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
      Left            =   12525
      TabIndex        =   8
      Top             =   9150
      Visible         =   0   'False
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
      Index           =   1
      Left            =   765
      MaxLength       =   10
      TabIndex        =   1
      Tag             =   "Fecha Entrada|F|N|||rclasifica_imp|fechaent|dd/mm/yyyy||"
      Top             =   4905
      Width           =   1320
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
      Index           =   0
      Left            =   4290
      MaxLength       =   6
      TabIndex        =   3
      Tag             =   "Código Socio|N|N|0|999999|rclasifica_imp|codsocio|000000||"
      Top             =   4935
      Width           =   800
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmManPreClasifica.frx":000C
      Height          =   8145
      Left            =   120
      TabIndex        =   11
      Top             =   795
      Width           =   13480
      _ExtentX        =   23786
      _ExtentY        =   14367
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
      Left            =   12525
      TabIndex        =   12
      Top             =   9150
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Index           =   1
      Left            =   120
      TabIndex        =   9
      Top             =   9045
      Width           =   2415
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
         Left            =   40
         TabIndex        =   10
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
   Begin MSComctlLib.Toolbar ToolbarAyuda 
      Height          =   330
      Left            =   13170
      TabIndex        =   25
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
         Caption         =   "&Ver todos"
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
      Begin VB.Menu mnActualizar 
         Caption         =   "&Actualizar Entradas"
      End
      Begin VB.Menu mnImprimir 
         Caption         =   "&Imprimir"
         Shortcut        =   ^I
      End
      Begin VB.Menu mnImportacion 
         Caption         =   "Importar E&xcel"
         Shortcut        =   ^X
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
Attribute VB_Name = "frmManPreClasifica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' +-+-+-+-+-+-+-+-+-+-+-+-
' +-+- Autor: MONICA  +-+-
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

Private Const IdPrograma = 4015



Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)

'codi per al registe que s'afegix al cridar des d'atre formulari.
'Obrir en modo Insertar i tornar datos del registre insertat
Public NuevoCodigo As String
Public CodigoActual As String
Public DeConsulta As Boolean

Private CadenaConsulta As String
Private CadB As String

Private WithEvents frmSoc As frmManSocios 'mantenimiento de socios
Attribute frmSoc.VB_VarHelpID = -1
Private WithEvents frmCam As frmManCampos 'mantenimiento de campos
Attribute frmCam.VB_VarHelpID = -1
Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmVar As frmComVar 'variedades
Attribute frmVar.VB_VarHelpID = -1
Private WithEvents frmMens As frmMensajes 'entradas erroneas
Attribute frmMens.VB_VarHelpID = -1

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
Dim Indice As Byte 'Index del text1 on es poses els datos retornats des d'atres Formularis de Mtos
Dim i As Integer

Dim indCodigo As Integer

' utilizado para buscar por checks
Private BuscaChekc As String


Private Sub PonerModo(vModo)
Dim B As Boolean

    Modo = vModo
    
    B = (Modo = 2)
    If B Then
        PonerContRegIndicador lblIndicador, adodc1, CadB
    Else
        PonerIndicador lblIndicador, Modo
    End If
    
    For i = 0 To txtAux.Count - 1
        txtAux(i).visible = Not B
        txtAux(i).BackColor = vbWhite
    Next i
    
    txtAux2(0).visible = Not B
    txtAux2(5).visible = Not B
    
    For i = 0 To btnBuscar.Count - 1
        btnBuscar(i).visible = Not B
    Next i

    CmdAceptar.visible = Not B
    CmdCancelar.visible = Not B
    DataGrid1.Enabled = B
    chkAux(0).visible = Not B
    
    'Si es regresar
    If DatosADevolverBusqueda <> "" Then cmdRegresar.visible = B
    
    PonerLongCampos
    PonerModoOpcionesMenu 'Activar/Desact botones de menu segun Modo
    PonerOpcionesMenu  'En funcion del usuario
    
    'Si estamos modo Modificar bloquear clave primaria
    BloquearTxt txtAux(3), (Modo = 4)
'    txtAux(5).visible = (Modo = 1)
    
    BloquearChk Me.chkAux(0), (Modo = 4) Or (Modo = 3)
    Me.chkAux(0).visible = (Modo = 1)

End Sub


Private Sub PonerModoOpcionesMenu()
'Activa/Desactiva botones del la toobar y del menu, segun el modo en que estemos
Dim B As Boolean

    B = (Modo = 2)
    'Busqueda
    Toolbar1.Buttons(5).Enabled = B
    Me.mnBuscar.Enabled = B
    'Ver Todos
    Toolbar1.Buttons(6).Enabled = B
    Me.mnVerTodos.Enabled = B
    
    'Insertar
    Toolbar1.Buttons(1).Enabled = B And Not DeConsulta
    Me.mnNuevo.Enabled = B And Not DeConsulta
    
    B = (B And adodc1.Recordset.RecordCount > 0) And Not DeConsulta
    'Modificar
    Toolbar1.Buttons(2).Enabled = B
    Me.mnModificar.Enabled = B
    'Eliminar
    Toolbar1.Buttons(3).Enabled = B
    Me.mnEliminar.Enabled = B
    'Imprimir
    Toolbar1.Buttons(8).Enabled = B
    Me.mnImprimir.Enabled = B
    
    'Actualizar
    Toolbar2.Buttons(1).Enabled = B
    Me.mnActualizar.Enabled = B

    'Importacion
    Toolbar2.Buttons(2).Enabled = UCase(Dir(App.Path & "\controlclas.cfg")) = UCase("controlclas.cfg")
    Me.mnImportacion.Enabled = UCase(Dir(App.Path & "\controlclas.cfg")) = UCase("controlclas.cfg")

    
End Sub

Private Sub BotonAnyadir()
    Dim NumF As String
    Dim anc As Single
    
    CargaGrid CadB, True 'primer de tot carregue tot el grid
'    CadB = ""
   
'    '******************** canviar taula i camp **************************
'    If (DatosADevolverBusqueda <> "") And NuevoCodigo <> "" Then
'        NumF = NuevoCodigo
'    Else
'        NumF = SugerirCodigoSiguienteStr("productos", "codprodu")
'    End If
'    '********************************************************************
    'Situamos el grid al final
    AnyadirLinea DataGrid1, adodc1
         
    anc = DataGrid1.Top
    If DataGrid1.Row < 0 Then
        anc = anc + 230
    Else
        anc = anc + DataGrid1.RowTop(DataGrid1.Row) + 5
    End If
    For i = 0 To txtAux.Count - 1
        txtAux(i).Text = ""
    Next i
    txtAux(1).Text = Format(Now, "dd/mm/yyyy")
    txtAux2(0).Text = ""
    txtAux2(5).Text = ""
    chkAux(0).Value = 0

    LLamaLineas anc, 3 'Pone el form en Modo=3, Insertar
       
    'Ponemos el foco
    PonerFoco txtAux(3)
End Sub

Private Sub BotonVerTodos()
    CadB = ""
    CargaGrid ""
    PonerModo 2
End Sub

Private Sub BotonBuscar()
    ' ***************** canviar per la clau primaria ********
    CargaGrid "rclasifica_imp.numnotac = -1"
    '*******************************************************************************
    'Buscar
    For i = 0 To txtAux.Count - 1
        txtAux(i).Text = ""
    Next i
    Me.txtAux2(0).Text = ""
    Me.txtAux2(5).Text = ""
    chkAux(0).Value = 0
    
    LLamaLineas DataGrid1.Top + 240, 1 'Pone el form en Modo=1, Buscar
    PonerFoco txtAux(3)
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
        anc = DataGrid1.RowTop(DataGrid1.Row) + 495 '545
    End If

    'Llamamos al form
    txtAux(3).Text = DataGrid1.Columns(0).Text
    txtAux(1).Text = DataGrid1.Columns(1).Text
    txtAux(5).Text = DataGrid1.Columns(2).Text
    txtAux2(5).Text = DataGrid1.Columns(3).Text
    txtAux(0).Text = DataGrid1.Columns(4).Text
    txtAux2(0).Text = DataGrid1.Columns(5).Text
    txtAux(2).Text = DataGrid1.Columns(6).Text
    txtAux(6).Text = DataGrid1.Columns(7).Text
    txtAux(4).Text = DataGrid1.Columns(8).Text
    
    Me.chkAux(0).Value = Me.adodc1.Recordset!procesado
    

    LLamaLineas anc, 4 'Pone el form en Modo=4, Modificar
   
    'Como es modificar
    PonerFoco txtAux(1)
    Screen.MousePointer = vbDefault
End Sub

Private Sub LLamaLineas(alto As Single, xModo As Byte)
    DeseleccionaGrid Me.DataGrid1
    PonerModo xModo
    
    'Fijamos el ancho
    For i = 0 To txtAux.Count - 1
        txtAux(i).Top = alto
    Next i
    
    ' ### [Monica] 12/09/2006
    txtAux2(0).Top = alto
    txtAux2(5).Top = alto
    For i = 0 To btnBuscar.Count - 1
        btnBuscar(i).Top = alto - 5
    Next i
    Me.chkAux(0).Top = alto

End Sub

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
    Sql = "¿Seguro que desea eliminar el Registro?"
    Sql = Sql & vbCrLf & "Num.Nota: " & adodc1.Recordset.Fields(0)
    
    If MsgBox(Sql, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        NumRegElim = adodc1.Recordset.AbsolutePosition
        Sql = "Delete from rclasifica_imp where numnotac=" & adodc1.Recordset!NumNotac
        
        conn.Execute Sql
        CargaGrid CadB
'        If CadB <> "" Then
'            CargaGrid CadB
'            lblIndicador.Caption = "BUSQUEDA: " & PonerContRegistros(Me.adodc1)
'        Else
'            CargaGrid ""
'            lblIndicador.Caption = ""
'        End If
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
        Case 0 'socios
            
            Indice = Index
            Set frmSoc = New frmManSocios
            frmSoc.DatosADevolverBusqueda = "0|1|"
            frmSoc.Show vbModal
            Set frmSoc = Nothing
            PonerFoco txtAux(Indice)
    
       Case 1 ' Fecha
            Dim esq As Long
            Dim dalt As Long
            Dim menu As Long
            Dim obj As Object
        
            Set frmC = New frmCal
            
            Indice = Index
            
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
            If Index = 1 Then
                If txtAux(1).Text <> "" Then frmC.NovaData = txtAux(1).Text
            Else
                If txtAux(5).Text <> "" Then frmC.NovaData = txtAux(5).Text
            End If
            
            ' ********************************************
        
            frmC.Show vbModal
            Set frmC = Nothing
            ' *** repasar si el camp es txtAux o Text1 ***
            PonerFoco txtAux(1) '<===
            ' ********************************************
     
        Case 2 'codigo de campo
            Set frmCam = New frmManCampos
            frmCam.DatosADevolverBusqueda = "0|"
'            frmCam.CodigoActual = Text1(5).Text
            frmCam.Show vbModal
            Set frmCam = Nothing
            PonerFoco txtAux(2)
            
        Case 3
            AbrirFrmVariedades Index + 2
            
    End Select
    
    If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Me.adodc1, 1
End Sub


Private Sub AbrirFrmVariedades(Indice As Integer)
    indCodigo = Indice
    Set frmVar = New frmComVar
    frmVar.DatosADevolverBusqueda = "0|1|"
    frmVar.CodigoActual = txtAux(indCodigo)
    frmVar.Show vbModal
    Set frmVar = Nothing
End Sub


Private Sub chkAux_Click(Index As Integer)
    If Modo = 1 Then
        'Buscqueda
        If InStr(1, BuscaChekc, "chkAux(" & Index & ")") = 0 Then BuscaChekc = BuscaChekc & "chkAux(" & Index & ")|"
    End If
End Sub

Private Sub chkAux_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub cmdAceptar_Click()
    Dim i As Long

    Select Case Modo
        Case 1 'BUSQUEDA
            CadB = ObtenerBusqueda(Me)
            If CadB <> "" Then
                CargaGrid CadB
                PonerModo 2
'                lblIndicador.Caption = "BUSQUEDA: " & PonerContRegistros(Me.adodc1)
                PonerFocoGrid Me.DataGrid1
            End If
            
        Case 3 'INSERTAR
            If DatosOK Then
                If InsertarDesdeForm(Me) Then
                    CargaGrid CadB
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
                    CadB = ""
                End If
            End If
            
        Case 4 'MODIFICAR
            If DatosOK Then
                If ModificaDesdeFormulario(Me) Then
                    TerminaBloquear
                    i = adodc1.Recordset.Fields(0)
                    PonerModo 2
                    CargaGrid CadB
'                    If CadB <> "" Then
'                        CargaGrid CadB
'                        lblIndicador.Caption = "BUSQUEDA: " & PonerContRegistros(Me.adodc1)
'                    Else
'                        CargaGrid
'                        lblIndicador.Caption = ""
'                    End If
                    adodc1.Recordset.Find (adodc1.Recordset.Fields(0).Name & " =" & i)
                    PonerFocoGrid Me.DataGrid1
                End If
            End If
    End Select
End Sub

Private Sub cmdCancelar_Click()
    On Error Resume Next
    
    Select Case Modo
        Case 1 'búsqueda
            CargaGrid CadB
        Case 3 'insertar
            DataGrid1.AllowAddNew = False
            CargaGrid
            If Not adodc1.Recordset.EOF Then adodc1.Recordset.MoveFirst
        Case 4 'modificar
            TerminaBloquear
    End Select
    
    PonerModo 2
    
'    If CadB <> "" Then
'        lblIndicador.Caption = "BUSQUEDA: " & PonerContRegistros(Me.adodc1)
''    Else
''        lblIndicador.Caption = ""
'    End If
    
    PonerFocoGrid Me.DataGrid1
    If Err.Number <> 0 Then Err.Clear
End Sub

Private Sub cmdRegresar_Click()
Dim cad As String
Dim i As Integer
Dim J As Integer
Dim Aux As String

    If adodc1.Recordset.EOF Then
        MsgBox "Ningún registro devuelto.", vbExclamation
        Exit Sub
    End If
    cad = ""
    i = 0
    Do
        J = i + 1
        i = InStr(J, DatosADevolverBusqueda, "|")
        If i > 0 Then
            Aux = Mid(DatosADevolverBusqueda, J, i - J)
            J = Val(Aux)
            cad = cad & adodc1.Recordset.Fields(J) & "|"
        End If
    Loop Until i = 0
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
    If Modo = 2 Then PonerContRegIndicador lblIndicador, adodc1, CadB
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
                SituarData Me.adodc1, "codprodu=" & CodigoActual, "", True
            End If
        End If
    End If
End Sub

Private Sub Form_Load()
    PrimeraVez = True
    
    'Icono del formulario
    Me.Icon = frmPpal.Icon
    

    With Me.Toolbar1
        .HotImageList = frmPpal.imgListComun_OM
        .DisabledImageList = frmPpal.imgListComun_BN
        .ImageList = frmPpal.imgListComun
        'el 1 es separadors
        .Buttons(5).Image = 1   'Buscar
        .Buttons(6).Image = 2   'Todos
        'el 4 i el 5 son separadors
        .Buttons(1).Image = 3   'Insertar
        .Buttons(2).Image = 4   'Modificar
        .Buttons(3).Image = 5   'Borrar
        
        .Buttons(8).Image = 10  'Imprimir
    End With

    With Me.Toolbar2
        .HotImageList = frmPpal.imgListComun_OM
        .DisabledImageList = frmPpal.imgListComun_BN
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 33 'actualizar entradas
        .Buttons(2).Image = 35  'exportacion a excel
    End With

    ' La Ayuda
    With Me.ToolbarAyuda
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 12
    End With


    '## A mano
'    chkVistaPrevia.Value = CheckValueLeer(Name)
    
    '****************** canviar la consulta *********************************+
    CadenaConsulta = "SELECT rclasifica_imp.numnotac, rclasifica_imp.fechaent, rclasifica_imp.codvarie, variedades.nomvarie, rclasifica_imp.codsocio, rsocios.nomsocio,"
    CadenaConsulta = CadenaConsulta & "rclasifica_imp.codcampo, rclasifica_imp.numcajon, rclasifica_imp.kilosnet, rclasifica_imp.procesado,  IF(procesado=1,'*','') as procesa "
    CadenaConsulta = CadenaConsulta & " FROM (rclasifica_imp left join rsocios on rclasifica_imp.codsocio = rsocios.codsocio) left join variedades on rclasifica_imp.codvarie = variedades.codvarie "
    CadenaConsulta = CadenaConsulta & " where (1=1) "
    '************************************************************************
    
    CadB = ""
    CargaGrid "rclasifica_imp.numnotac is null"
    
'    If (DatosADevolverBusqueda <> "") And NuevoCodigo <> "" Then
'        BotonAnyadir
'    Else
'        PonerModo 2
'    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
'    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
    If Modo = 4 Then TerminaBloquear
    Screen.MousePointer = vbDefault
End Sub

Private Sub frmCam_DatoSeleccionado(CadenaSeleccion As String)
    txtAux(2).Text = RecuperaValor(CadenaSeleccion, 1) 'codigo campo
End Sub

Private Sub frmC_Selec(vFecha As Date)
    ' *** repasar si el camp es txtAux o Text1 ***
    If Indice = 1 Then
        txtAux(1).Text = Format(vFecha, "dd/mm/yyyy") '<===
    End If
    ' ********************************************
End Sub

Private Sub frmSoc_DatoSeleccionado(CadenaSeleccion As String)
    txtAux(0).Text = RecuperaValor(CadenaSeleccion, 1) 'codigo socio
    txtAux2(0).Text = RecuperaValor(CadenaSeleccion, 2) 'nombre socio
End Sub

Private Sub frmVar_DatoSeleccionado(CadenaSeleccion As String)
    txtAux(5).Text = RecuperaValor(CadenaSeleccion, 1) 'codigo variedad
    txtAux2(5).Text = RecuperaValor(CadenaSeleccion, 2) 'nombre variedad
End Sub

Private Sub mnBuscar_Click()
    BotonBuscar
End Sub

Private Sub mnActualizar_Click()
    BotonActualizarEntradas
End Sub

Private Sub mnEliminar_Click()
    BotonEliminar
End Sub

Private Sub mnImportacion_Click()
    
'    If Not CargarCondicion Then Exit Sub
    
    Shell App.Path & "\clasificacion.exe /A|" & vUsu.CadenaConexion & "|", vbNormalFocus
    
  
    
End Sub

Private Sub mnImprimir_Click()
Dim Sql As String

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
        Case 1
                mnNuevo_Click
        Case 2
                mnModificar_Click
        Case 3
                mnEliminar_Click
        Case 5
                mnBuscar_Click
        Case 6
                mnVerTodos_Click
    End Select
End Sub

Private Sub CargaGrid(Optional vSQL As String, Optional Ascendente As Boolean)
    Dim Sql As String
    Dim tots As String
    
'    adodc1.ConnectionString = Conn
    If vSQL <> "" Then
        Sql = CadenaConsulta & " AND " & vSQL
    Else
        Sql = CadenaConsulta
    End If
    Sql = Sql & " ORDER BY  rclasifica_imp.numnotac "
    
    CargaGridGnral Me.DataGrid1, Me.adodc1, Sql, PrimeraVez
    
    ' *******************canviar els noms i si fa falta la cantitat********************
    tots = "S|txtAux(3)|T|Nota|1200|;S|txtAux(1)|T|Fecha|1400|;S|btnBuscar(1)|B||195|;"
    tots = tots & "S|txtAux(5)|T|Código|1000|;S|btnBuscar(3)|B||195|;S|txtAux2(5)|T|Variedad|1800|;"
    tots = tots & "S|txtAux(0)|T|Código|1000|;S|btnBuscar(0)|B||195|;S|txtAux2(0)|T|Nombre Socio|2500|;"
    tots = tots & "S|txtAux(2)|T|Campo|1400|;S|btnBuscar(2)|B||195|;"
    tots = tots & "S|txtAux(6)|T|Cajones|1100|;"
    tots = tots & "S|txtAux(4)|T|Kilos|1100|;N||||0|;S|chkAux(0)|CB|IA|360|;"
    
    
    
    arregla tots, DataGrid1, Me, 350
    
    
    DataGrid1.ScrollBars = dbgAutomatic
    DataGrid1.Columns(0).Alignment = dbgLeft
    DataGrid1.Columns(3).Alignment = dbgLeft
'    DataGrid1.Columns(10).Alignment = dbgCenter
'    DataGrid1.Columns(12).Alignment = dbgCenter
End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1
                mnActualizar_Click
        Case 2
                mnImportacion_Click
    End Select

End Sub

Private Sub ToolbarAyuda_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1
            LanzaVisorMimeDocumento Me.hWnd, DireccionAyuda & IdPrograma & ".html"
    End Select
End Sub

Private Sub txtaux_GotFocus(Index As Integer)
    ConseguirFoco txtAux(Index), Modo
End Sub


Private Sub txtAux_LostFocus(Index As Integer)
Dim cadMen As String
Dim Sql As String


    If Not PerderFocoGnral(txtAux(Index), Modo) Then Exit Sub
    
    Select Case Index
        Case 0 'codigo de socio
            If PonerFormatoEntero(txtAux(Index)) Then
                If Modo = 1 Then Exit Sub
                txtAux2(Index).Text = PonerNombreDeCod(txtAux(Index), "rsocios", "nomsocio")
                If txtAux2(Index).Text = "" Then
                    cadMen = "No existe el Socio: " & txtAux(Index).Text & vbCrLf
                    cadMen = cadMen & "¿Desea crearlo?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmSoc = New frmManSocios
                        frmSoc.DatosADevolverBusqueda = "0|1|"
                        txtAux(Index).Text = ""
                        TerminaBloquear
                        frmSoc.Show vbModal
                        Set frmSoc = Nothing
                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, adodc1, 1
                    Else
                        txtAux(Index).Text = ""
                    End If
                    PonerFoco txtAux(Index)
                End If
            End If
        
        Case 1 'fecha
            '[Monica]28/08/2013: comprobamos que la fecha esté en la campaña
            PonerFormatoFecha txtAux(Index), True
            
        Case 5 'codigo de variedad
            If PonerFormatoEntero(txtAux(Index)) Then
                If Modo = 1 Then Exit Sub
                txtAux2(Index).Text = PonerNombreDeCod(txtAux(Index), "variedades", "nomvarie")
                If txtAux2(Index).Text = "" Then
                    cadMen = "No existe la Variedad: " & txtAux(Index).Text & vbCrLf
                    cadMen = cadMen & "¿Desea crearla?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmVar = New frmComVar
                        frmVar.DatosADevolverBusqueda = "0|1|"
                        txtAux(Index).Text = ""
                        TerminaBloquear
                        frmVar.Show vbModal
                        Set frmVar = Nothing
                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, adodc1, 1
                    Else
                        txtAux(Index).Text = ""
                    End If
                    PonerFoco txtAux(Index)
                End If
                
            End If
    
        Case 2 'Campo
            If PonerFormatoEntero(txtAux(Index)) Then
                If Modo = 1 Then Exit Sub
                Sql = ""
                Sql = DevuelveDesdeBDNew(cAgro, "rcampos", "codcampo", "codcampo", txtAux(Index).Text, "N")
                If Sql = "" Then
                    cadMen = "No existe el Campo: " & txtAux(Index).Text & vbCrLf
                    cadMen = cadMen & "¿Desea crearlo?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmCam = New frmManCampos
                        frmCam.DatosADevolverBusqueda = "0|1|"
'                        frmCamp.NuevoCodigo = Text1(Index).Text
                        txtAux(Index).Text = ""
                        TerminaBloquear
                        frmCam.Show vbModal
                        Set frmCam = Nothing
                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Me.adodc1, 1
                    Else
                        txtAux(Index).Text = ""
                    End If
                    PonerFoco txtAux(Index)
                Else
                    If Not EstaCampoDeAlta(txtAux(Index).Text) Then
                        MsgBox "El campo no está dado de alta. Reintroduzca.", vbExclamation
                        txtAux(Index).Text = ""
                        PonerFoco txtAux(Index)
                    End If
                End If
            End If

        Case 6 ' cajas
            If Modo = 1 Then Exit Sub
            
            PonerFormatoEntero txtAux(6)
            
        Case 4 ' kilos
            If Modo = 1 Then Exit Sub
            
            PonerFormatoEntero txtAux(4)
        
        Case 3 ' numero de nota de entrada
            If Modo = 1 Then Exit Sub
            
            PonerFormatoEntero txtAux(3)
        
    End Select
    
End Sub

Private Function DatosOK() As Boolean
'Dim Datos As String
Dim B As Boolean
Dim Sql As String
Dim Mens As String


    B = CompForm(Me)
    If Not B Then Exit Function
    
    If Modo = 3 Then   'Estamos insertando
        Sql = "select count(*) from rclasifica_imp where numnotac = " & DBSet(txtAux(3).Text, "N")
        
        If TotalRegistros(Sql) <> 0 Then
            MsgBox "Número de nota ya existe.", vbExclamation
            PonerFoco txtAux(3)
            B = False
        End If
    End If
    
    DatosOK = B
End Function


Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
    PonerOpcionesMenuGeneralNew Me
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
Private Sub txtaux_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = teclaBuscar Then
        If Modo = 1 Or Modo = 3 Or Modo = 4 Then
            Select Case Index
                Case 1: KEYBusqueda KeyAscii, 1 'fecha
                Case 5: KEYBusqueda KeyAscii, 3 'variedad
                Case 0: KEYBusqueda KeyAscii, 0 'socio
                Case 2: KEYBusqueda KeyAscii, 2 'campo
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
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me

End Sub

Private Sub KEYBusqueda(KeyAscii As Integer, Indice As Integer)
    KeyAscii = 0
    btnBuscar_Click (Indice)
End Sub

Private Sub BotonActualizarEntradas()
Dim Sql As String

    Sql = "select * from rclasifica_imp "
    
    If TotalRegistrosConsulta(Sql) = 0 Then
        MsgBox "No hay entradas para actualizar.", vbExclamation
        Exit Sub
    Else
        If Not HayEntradasErroneas Then
        
            If ActualizarEntradas Then
                MsgBox "Proceso realizado correctamente.", vbExclamation
                CargaGrid
            End If
        
        End If
    End If
End Sub

Private Function HayEntradasErroneas() As Boolean
Dim Sql As String
Dim RS1 As ADODB.Recordset
Dim YaExiste As Boolean
Dim Rs As ADODB.Recordset
Dim Sql2 As String

    On Error GoTo eHayEntradasErroneas

    HayEntradasErroneas = True

    Sql = "delete from tmpexcel where codusu = " & vUsu.Codigo
    conn.Execute Sql
    
    Sql = "select * from rclasifica_imp order by numnotac"
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not Rs.EOF
        'Comprobamos que el numero de nota no exista ya en rclasifica  or rhisfruta_entradas
        YaExiste = False
        Sql = "select count(*) from rclasifica where numnotac = " & DBSet(Rs!NumNotac, "N")
        If TotalRegistros(Sql) <> 0 Then
            YaExiste = True
        Else
            Sql = "select count(*) from rhisfruta_entradas where numnotac = " & DBSet(Rs!NumNotac, "N")
            If TotalRegistros(Sql) <> 0 Then
                YaExiste = True
            End If
        End If
        
        If YaExiste Then
            Sql2 = "insert into tmpexcel (codusu,numalbar,fecalbar,codvarie,codsocio,codcampo,tipoentr) values ("
            Sql2 = Sql2 & vUsu.Codigo & "," & DBSet(Rs!NumNotac, "N") & "," & DBSet(Rs!FechaEnt, "F") & "," & DBSet(Rs!Codvarie, "N") & ","
            Sql2 = Sql2 & DBSet(Rs!Codsocio, "N") & "," & DBSet(Rs!codcampo, "N") & ",0)"
            
            conn.Execute Sql2
        End If
        
        'Comprobamos que exista la variedad
        Sql = "select count(*) from variedades where codvarie = " & DBSet(Rs!Codvarie, "N")
        If TotalRegistros(Sql) = 0 Then
            Sql2 = "insert into tmpexcel (codusu,numalbar,fecalbar,codvarie,codsocio,codcampo,tipoentr) values ("
            Sql2 = Sql2 & vUsu.Codigo & "," & DBSet(Rs!NumNotac, "N") & "," & DBSet(Rs!FechaEnt, "F") & "," & DBSet(Rs!Codvarie, "N") & ","
            Sql2 = Sql2 & DBSet(Rs!Codsocio, "N") & "," & DBSet(Rs!codcampo, "N") & ",1)"
            
            conn.Execute Sql2
        End If
        
        'Comprobamos que exita el socio
        Sql = "select count(*) from rsocios where codsocio = " & DBSet(Rs!Codsocio, "N")
        If TotalRegistros(Sql) = 0 Then
            Sql2 = "insert into tmpexcel (codusu,numalbar,fecalbar,codvarie,codsocio,codcampo,tipoentr) values ("
            Sql2 = Sql2 & vUsu.Codigo & "," & DBSet(Rs!NumNotac, "N") & "," & DBSet(Rs!FechaEnt, "F") & "," & DBSet(Rs!Codvarie, "N") & ","
            Sql2 = Sql2 & DBSet(Rs!Codsocio, "N") & "," & DBSet(Rs!codcampo, "N") & ",2)"
            
            conn.Execute Sql2
        End If
        
        'Comprobamos que exista el campo para el socio / variedad
        Sql = "select count(*) from rcampos where codcampo = " & DBSet(Rs!codcampo, "N")
        Sql = Sql & " and codsocio = " & DBSet(Rs!Codsocio, "N") & " and codvarie = " & DBSet(Rs!Codvarie, "N")
        If TotalRegistros(Sql) = 0 Then
            Sql2 = "insert into tmpexcel (codusu,numalbar,fecalbar,codvarie,codsocio,codcampo,tipoentr) values ("
            Sql2 = Sql2 & vUsu.Codigo & "," & DBSet(Rs!NumNotac, "N") & "," & DBSet(Rs!FechaEnt, "F") & "," & DBSet(Rs!Codvarie, "N") & ","
            Sql2 = Sql2 & DBSet(Rs!Codsocio, "N") & "," & DBSet(Rs!codcampo, "N") & ",3)"
            
            conn.Execute Sql2
        End If
    
        Rs.MoveNext
    Wend
    
    Set Rs = Nothing
    
    Sql = "select count(*) from tmpexcel where codusu = " & vUsu.Codigo
    If TotalRegistros(Sql) <> 0 Then
        Set frmMens = frmMensajes
        frmMens.OpcionMensaje = 41
        frmMens.Show vbModal
        Set frmMens = Nothing
            
        HayEntradasErroneas = True
    Else
    
        HayEntradasErroneas = False
    End If
    
    Exit Function
    
eHayEntradasErroneas:
    MuestraError Err.Number, "Hay entradas erróneas", Err.Description
End Function

Private Function ActualizarEntradas() As Boolean
Dim Sql As String

    On Error GoTo eActualizarEntradas

    ActualizarEntradas = False

    conn.BeginTrans


    Sql = "insert into rclasifica (numnotac,fechaent,horaentr,codvarie,codsocio,codcampo,tipoentr,recolect,kilosbru,numcajon,kilosnet,tiporecol,impreso,transportadopor,kilostra) "
    Sql = Sql & " select numnotac,fechaent,concat(fechaent , ' 12:00:00'), codvarie, codsocio, codcampo, 0,0,kilosnet,numcajon,kilosnet,0,0,0,kilosnet from rclasifica_imp "
    
    conn.Execute Sql
    
    '[Monica]16/01/2014: insertamos las calidades de la variedad a cero
    Sql = "insert into rclasifica_clasif (numnotac, codvarie, codcalid, muestra, kilosnet) "
    Sql = Sql & " select cc.numnotac, cc.codvarie, dd.codcalid, 0, 0  "
    Sql = Sql & " from rclasifica_imp cc inner join rcalidad dd on cc.codvarie = dd.codvarie "
    
    conn.Execute Sql
    
    conn.Execute "delete from rclasifica_imp"
    
    ActualizarEntradas = True
    conn.CommitTrans
    Exit Function
    
eActualizarEntradas:
    conn.RollbackTrans
    MuestraError Err.Number, "Actualizar Entradas", Err.Description
End Function



Private Function Kilos(Cajas As String, Forfait As String) As String
Dim KilosCajon As String

    Kilos = ""

    ' si no hay cajas ni variedad, no podemos calcular los kilos
    If Cajas = "" Or Forfait = "" Then Exit Function
    
    KilosCajon = 0
    KilosCajon = DevuelveValor("select kiloscaj from forfaits where codforfait = " & DBSet(Forfait, "T"))
    
    Kilos = CStr(Round2(KilosCajon * CCur(ImporteSinFormato(Cajas)), 0))
    
    
End Function

Private Function Importe(Kilos As String, Forfait As String) As String
Dim PrecKilo As String
    
    Importe = ""
    
    If Kilos = "" Or Forfait = "" Then Exit Function
    
    PrecKilo = DevuelveValor("select preciokilonom from forfaits where codforfait = " & DBSet(Forfait, "T"))
    
    Importe = Round2(CCur(Kilos) * CCur(ImporteSinFormato(PrecKilo)), 2)
    
End Function


Private Function Horas(Importe As String, Trabajador As String) As String
Dim PrecHora As String
Dim Sql As String

    Sql = "select impsalar from straba inner join salarios on straba.codcateg = salarios.codcateg "
    Sql = Sql & " where straba.codtraba = " & DBSet(Trabajador, "N")
    
    PrecHora = DevuelveValor(Sql)
    
    Horas = 0
    
    If CCur(PrecHora) <> 0 Then
        Horas = Round2(CCur(Importe) / PrecHora, 2)
    End If
End Function


Private Function CargarCondicion() As Boolean
Dim Sql As String
Dim NFic As Integer

    On Error GoTo eCargarCondicion

    CargarCondicion = False

    If Dir(App.Path & "\condicionsql.txt", vbArchive) <> "" Then Kill (App.Path & "\condicionsql.txt")
        
    If Me.adodc1.Recordset.RecordCount = 0 Then
        CargarCondicion = True
        Exit Function
    End If
        
    NFic = FreeFile
    
    Open App.Path & "\condicionsql.txt" For Output As #NFic
    
    Print #NFic, adodc1.RecordSource
        
    Close #NFic
    
    CargarCondicion = True
    
    Exit Function
    
eCargarCondicion:
    MuestraError Err.Number, "Cargando condición", Err.Description
End Function

Private Sub printNou()
    With frmImprimir2
        .cadTabla2 = "rclasifica_imp"
        .Informe2 = "rInfEntradasImp.rpt"
        If CadB <> "" Then
            '.cadRegSelec = Replace(SQL2SF(CadB), "clientes", "clientes_1")
            .cadRegSelec = SQL2SF(CadB)
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
        .OtrosParametros2 = "pEmpresa='" & vEmpresa.nomempre & "'|pOrden={rclasifica_imp.numnotac}|"
        '.OtrosParametros2 = "pEmpresa='" & vEmpresa.NomEmpre & "'|"
        ' *** posar el nº de paràmetres que he posat en OtrosParametros2 ***
        '.NumeroParametros2 = 1
        .NumeroParametros2 = 2
        ' ******************************************************************
        .MostrarTree2 = False
        .InfConta2 = False
        .ConSubInforme2 = True
        .SubInformeConta = ""
        .Show vbModal
    End With
End Sub


