VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmManCalidades 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Calidades"
   ClientHeight    =   10065
   ClientLeft      =   45
   ClientTop       =   30
   ClientWidth     =   17265
   Icon            =   "frmManCalidades.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10065
   ScaleWidth      =   17265
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Height          =   255
      Index           =   1
      Left            =   12150
      TabIndex        =   11
      Tag             =   "Se aplica bonificaciones|N|N|0|1|rcalidad|seaplicabonif|||"
      Top             =   4980
      Visible         =   0   'False
      Width           =   255
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
      Left            =   11220
      MaxLength       =   6
      TabIndex        =   10
      Tag             =   "Eu.Rec.Coop|N|S|||rcalidad|eurreccoop|0.0000||"
      Text            =   "E.Rec."
      Top             =   4950
      Width           =   900
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
      Left            =   10230
      MaxLength       =   6
      TabIndex        =   9
      Tag             =   "Eu.Rec.Soc|N|S|||rcalidad|eurrecsoc|0.0000||"
      Text            =   "E.Rec.S"
      Top             =   4950
      Width           =   900
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
      Index           =   5
      Left            =   8940
      MaxLength       =   30
      TabIndex        =   7
      Tag             =   "Nom.Calibrador 2|T|S|||rcalidad|nomcalibrador2|||"
      Text            =   "nom.calibrador 2"
      Top             =   4950
      Width           =   900
   End
   Begin VB.CheckBox chkAux 
      BackColor       =   &H80000005&
      Height          =   255
      Index           =   0
      Left            =   9930
      TabIndex        =   8
      Tag             =   "Hay gastos|N|N|0|1|rcalidad|gastosrec|||"
      Top             =   4950
      Visible         =   0   'False
      Width           =   255
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
      Left            =   7965
      MaxLength       =   30
      TabIndex        =   6
      Tag             =   "Nom.Calibrador 1|T|N|||rcalidad|nomcalibrador1|||"
      Text            =   "nom.calibrador 1"
      Top             =   4950
      Width           =   900
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
      Index           =   1
      ItemData        =   "frmManCalidades.frx":000C
      Left            =   7245
      List            =   "frmManCalidades.frx":000E
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Tag             =   "Tipo Calidad 1|N|N|0|2|rcalidad|tipcalid1|||"
      Top             =   4950
      Width           =   735
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
      ItemData        =   "frmManCalidades.frx":0010
      Left            =   6525
      List            =   "frmManCalidades.frx":0012
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Tag             =   "Tipo Calidad|N|N|0|4|rcalidad|tipcalid|||"
      Top             =   4950
      Width           =   735
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
      Index           =   3
      Left            =   5580
      MaxLength       =   3
      TabIndex        =   3
      Tag             =   "Calidad Abrev.|T|N|||rcalidad|nomcalab|||"
      Text            =   "abr"
      Top             =   4950
      Width           =   900
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
      Index           =   2
      Left            =   1170
      TabIndex        =   19
      Top             =   4950
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton btnBuscar 
      Appearance      =   0  'Flat
      Caption         =   "+"
      Height          =   350
      Index           =   0
      Left            =   900
      MaskColor       =   &H00000000&
      TabIndex        =   18
      ToolTipText     =   "Buscar Variedad"
      Top             =   4950
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
      Height          =   330
      Index           =   1
      Left            =   3060
      MaxLength       =   2
      TabIndex        =   1
      Tag             =   "Codigo Calidad|N|S|1|20|rcalidad|codcalid|00|S|"
      Text            =   "Ca"
      Top             =   4950
      Width           =   1425
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
      Left            =   14835
      TabIndex        =   12
      Top             =   9360
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
      Left            =   15990
      TabIndex        =   13
      Top             =   9360
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
      Index           =   2
      Left            =   4590
      MaxLength       =   12
      TabIndex        =   2
      Tag             =   "Nombre Calidad|T|N|||rcalidad|nomcalid|||"
      Text            =   "nomcal"
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
      Index           =   0
      Left            =   60
      MaxLength       =   6
      TabIndex        =   0
      Tag             =   "Código Variedad|N|N|0|999999|rcalidad|codvarie|000000|S|"
      Text            =   "Var"
      Top             =   4950
      Width           =   825
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmManCalidades.frx":0014
      Height          =   8145
      Left            =   135
      TabIndex        =   16
      Top             =   900
      Width           =   16920
      _ExtentX        =   29845
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
      Left            =   15990
      TabIndex        =   17
      Top             =   9360
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Index           =   1
      Left            =   150
      TabIndex        =   14
      Top             =   9285
      Width           =   2385
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
         TabIndex        =   15
         Top             =   210
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
      Left            =   16590
      TabIndex        =   22
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
      Begin VB.Menu mnImprimir 
         Caption         =   "&Imprimir"
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
Attribute VB_Name = "frmManCalidades"
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

Private Const IdPrograma = 2013

Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)

Public ParamVariedad As String

'codi per al registe que s'afegix al cridar des d'atre formulari.
'Obrir en modo Insertar i tornar datos del registre insertat
Public NuevoCodigo As String
Public CodigoActual As String
Public DeConsulta As Boolean

Private CadenaConsulta As String
Private CadB As String

Private WithEvents frmVar As frmComVar 'variedades
Attribute frmVar.VB_VarHelpID = -1
' utilizado para buscar por checks
Private BuscaChekc As String

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

Private Sub PonerModo(vModo)
Dim B As Boolean

    Modo = vModo
    BuscaChekc = ""
    
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
    
    txtAux2(2).visible = Not B
    btnBuscar(0).visible = Not B
    Combo1(0).visible = Not B
    Combo1(1).visible = Not B
    chkAux(0).visible = Not B
    chkAux(1).visible = Not B

    CmdAceptar.visible = Not B
    CmdCancelar.visible = Not B
    DataGrid1.Enabled = B
    
    'Si es regresar
    If DatosADevolverBusqueda <> "" Then cmdRegresar.visible = B
    
    PonerLongCampos
    PonerModoOpcionesMenu 'Activar/Desact botones de menu segun Modo
    PonerOpcionesMenu  'En funcion del usuario
    
    'Si estamos modo Modificar bloquear clave primaria
    BloquearTxt txtAux(0), (Modo = 4)
    BloquearTxt txtAux(1), (Modo = 4)
    BloquearBtn Me.btnBuscar(0), (Modo = 4)
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
    
End Sub

Private Sub BotonAnyadir()
    Dim NumF As String
    Dim anc As Single
    
'    CargaGrid 'primer de tot carregue tot el grid
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
        anc = anc + 240
    Else
        anc = anc + DataGrid1.RowTop(DataGrid1.Row) + 5
    End If
    txtAux(0).Text = NumF
    FormateaCampo txtAux(0)
    For i = 1 To txtAux.Count - 1
        txtAux(i).Text = ""
    Next i
    txtAux2(2).Text = ""
    Combo1(0).ListIndex = -1
    Combo1(1).ListIndex = -1
    chkAux(0).Value = 0
    chkAux(1).Value = 0

    LLamaLineas anc, 3 'Pone el form en Modo=3, Insertar
       
    'Ponemos el foco
    PonerFoco txtAux(0)
End Sub

Private Sub BotonVerTodos()
    CadB = ""
    CargaGrid ""
    PonerModo 2
End Sub

Private Sub BotonBuscar()
    ' ***************** canviar per la clau primaria ********
    CargaGrid "rcalidad.codvarie = -1"
    '*******************************************************************************
    'Buscar
    For i = 0 To txtAux.Count - 1
        txtAux(i).Text = ""
    Next i
    txtAux2(2).Text = ""
    Combo1(0).ListIndex = -1
    Combo1(1).ListIndex = -1
    chkAux(0).Value = 0
    chkAux(1).Value = 0
    
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
        anc = DataGrid1.RowTop(DataGrid1.Row) + DataGrid1.Top ' 545
    End If

    'Llamamos al form
    txtAux(0).Text = DataGrid1.Columns(0).Text
    txtAux2(2).Text = DataGrid1.Columns(1).Text
    txtAux(1).Text = DataGrid1.Columns(2).Text
    txtAux(2).Text = DataGrid1.Columns(3).Text
    txtAux(3).Text = DataGrid1.Columns(4).Text
    txtAux(4).Text = DataGrid1.Columns(9).Text
    txtAux(5).Text = DataGrid1.Columns(10).Text
    ' ***** canviar-ho pel nom del camp del combo *********
    i = adodc1.Recordset!tipcalid
    ' *****************************************************
    PosicionarCombo Me.Combo1(0), i
    
    i = adodc1.Recordset!tipcalid1
    ' *****************************************************
    PosicionarCombo Me.Combo1(1), i
    
    Me.chkAux(0).Value = Me.adodc1.Recordset!gastosrec
    
    '[Monica]11/12/2013
    txtAux(6).Text = DataGrid1.Columns(13).Text
    txtAux(7).Text = DataGrid1.Columns(14).Text
    
    Me.chkAux(1).Value = Me.adodc1.Recordset!seaplicabonif
    
    LLamaLineas anc, 4 'Pone el form en Modo=4, Modificar
   
    'Como es modificar
    PonerFoco txtAux(2)
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
    txtAux2(2).Top = alto
    btnBuscar(0).Top = alto - 10
    Combo1(0).Top = alto - 20
    Combo1(1).Top = alto - 20
    Me.chkAux(0).Top = alto
    Me.chkAux(1).Top = alto
    
End Sub


Private Sub BotonEliminar()
Dim Sql As String
Dim temp As Boolean

    On Error GoTo EEliminar
    'Ciertas comprobaciones
    If adodc1.Recordset.EOF Then Exit Sub
    If Not SepuedeBorrar Then Exit Sub
        
    ' ### [Monica] 26/09/2006 dejamos modificar y eliminar el codigo 0
    ' *** repasar el nom de l'adodc, l'index del Field i el camp que te la PK ***
    'El registre de codi 0 no es pot Modificar ni Eliminar
    'If EsCodigoCero(CStr(adodc1.Recordset.Fields(0).Value), FormatoCampo(txtAux(0))) Then Exit Sub
    ' ***************************************************************************
    
    '*************** canviar els noms i el DELETE **********************************
    Sql = "¿Seguro que desea eliminar la Calidad?"
    Sql = Sql & vbCrLf & "Variedad: " & adodc1.Recordset.Fields(0) & " " & adodc1.Recordset.Fields(1)
    Sql = Sql & vbCrLf & "Código: " & adodc1.Recordset.Fields(2)
    Sql = Sql & vbCrLf & "Descripción: " & adodc1.Recordset.Fields(3)
    
    If MsgBox(Sql, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        NumRegElim = adodc1.Recordset.AbsolutePosition
        Sql = "Delete from rcalidad where codvarie=" & adodc1.Recordset!Codvarie
        Sql = Sql & " and codcalid = " & adodc1.Recordset!codcalid
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
    
EEliminar:
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
        Case 0 'variedades de comercial
            
            Indice = Index
            Set frmVar = New frmComVar
            frmVar.DatosADevolverBusqueda = "0|1|"
            frmVar.CodigoActual = txtAux(Indice).Text
            frmVar.Show vbModal
            Set frmVar = Nothing
            PonerFoco txtAux(Indice)
    
    End Select
    
    If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Me.adodc1, 1
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
    Dim i As Integer
    Dim cad As String

    Select Case Modo
        Case 1 'BUSQUEDA
            CadB = ObtenerBusqueda(Me, BuscaChekc)
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
                            SituarDataMULTI adodc1, "codvarie = " & txtAux(0) & " and codcalid = " & txtAux(1), "" ' Find (adodc1.Recordset.Fields(2).Name & " =" & NuevoCodigo)
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
                    PosicionarData
                    PonerFocoGrid Me.DataGrid1
                End If
            End If
    End Select
End Sub

Private Sub PosicionarData()
Dim cad As String, Indicador As String

    ' *** canviar-ho per tota la PK de la capçalera, no llevar els () ***
    cad = "codvarie = " & adodc1.Recordset.Fields(0) & " and codcalid = " & adodc1.Recordset.Fields(2)
    ' ***************************************
    CargaGrid CadB
    ' *** gastar SituarData o SituarDataMULTI depenent de si la PK es simple o composta ***
    'If SituarDataMULTI(Data1, cad, Indicador) Then
    If SituarDataMULTI(adodc1, cad, Indicador, True) Then
        PonerModo 2
        lblIndicador.Caption = Indicador
    Else
       LimpiarCampos
       PonerModo 0
    End If
   
    ' ***********************************************************************************
End Sub

Private Sub LimpiarCampos()
Dim i As Integer

    On Error Resume Next
    
    limpiar Me   'Mètode general: Neteja els controls TextBox
    lblIndicador.Caption = ""
    Me.chkAux(0).Value = 0
    Me.chkAux(1).Value = 0
    
    For i = 0 To Combo1.Count - 1
        Combo1(i).ListIndex = -1
    Next i
    
    ' *** si n'hi han combos a la capçalera ***
    ' *****************************************

    If Err.Number <> 0 Then Err.Clear
End Sub




Private Sub cmdCancelar_Click()
    On Error Resume Next
    
    Select Case Modo
        Case 1 'búsqueda
            CargaGrid CadB
        Case 3 'insertar
            DataGrid1.AllowAddNew = False
            'CargaGrid
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

Private Sub Combo1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Combo1_GotFocus(Index As Integer)
    If Modo = 1 Then Combo1(Index).BackColor = vbLightBlue
End Sub

Private Sub Combo1_LostFocus(Index As Integer)
    If Combo1(Index).BackColor = vbLightBlue Then Combo1(Index).BackColor = vbWhite
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
                SituarData Me.adodc1, "codvarie=" & CodigoActual, "", True
            End If
        End If
    End If
End Sub

Private Sub Form_Load()
    'Icono del formulario
    Me.Icon = frmPpal.Icon
    PrimeraVez = True

    With Me.Toolbar1
        .HotImageList = frmPpal.imgListComun_OM
        .DisabledImageList = frmPpal.imgListComun_BN
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 3   'Insertar
        .Buttons(2).Image = 4   'Modificar
        .Buttons(3).Image = 5   'Borrar
        'el 1 es separadors
        .Buttons(5).Image = 1   'Buscar
        .Buttons(6).Image = 2   'Todos
        'el 9 i el 10 son separadors
        .Buttons(8).Image = 10  'imprimir
    End With

    ' La Ayuda
    With Me.ToolbarAyuda
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 12
    End With



    '## A mano
'    chkVistaPrevia.Value = CheckValueLeer(Name)
    
    CargaCombo
    
    '****************** canviar la consulta *********************************+
    CadenaConsulta = "SELECT rcalidad.codvarie, variedades.nomvarie, rcalidad.codcalid,"
    CadenaConsulta = CadenaConsulta & "rcalidad.nomcalid, rcalidad.nomcalab, rcalidad.tipcalid, "
    CadenaConsulta = CadenaConsulta & " CASE rcalidad.tipcalid WHEN 0 THEN ""Normal"" WHEN 1 THEN ""Destrio"" WHEN 2 THEN ""Venta Campo"" WHEN 3 THEN ""Mermas"" WHEN 4 THEN ""Pequeño"" WHEN 5 THEN ""Pixat"" WHEN 6 THEN ""Destr.Comer."" END, "
    CadenaConsulta = CadenaConsulta & "rcalidad.tipcalid1, "
    CadenaConsulta = CadenaConsulta & " CASE rcalidad.tipcalid1 WHEN 0 THEN ""Comercial"" WHEN 1 THEN ""No Comercial"" WHEN 2 THEN ""Retirada"" END, "
    CadenaConsulta = CadenaConsulta & " nomcalibrador1, nomcalibrador2, gastosrec, IF(gastosrec=1,'*','') as dgastorec, "
    CadenaConsulta = CadenaConsulta & " rcalidad.eurrecsoc, rcalidad.eurreccoop, "
    '[Monica]21/01/2016: se aplica bonificaciones para anticipos y liquidaciones (PICASSENT)
    CadenaConsulta = CadenaConsulta & " seaplicabonif, IF(seaplicabonif=1,'*','') as dseaplicabonif"
    CadenaConsulta = CadenaConsulta & " FROM variedades, rcalidad"
    CadenaConsulta = CadenaConsulta & " WHERE variedades.codvarie = rcalidad.codvarie "
    '************************************************************************
    
'    Me.DataGrid1.Height = Me.DataGrid1.Height + 60

    CadB = ""
    CargaGrid
    
'    If (DatosADevolverBusqueda <> "") And NuevoCodigo <> "" Then
'        BotonAnyadir
'    Else
'        PonerModo 2
'    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
'    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
    Screen.MousePointer = vbDefault
    If Modo = 4 Then TerminaBloquear
End Sub


Private Sub frmVar_DatoSeleccionado(CadenaSeleccion As String)
'Variedad comercial
    txtAux(Indice).Text = RecuperaValor(CadenaSeleccion, 1) 'codvarie
    txtAux2(Indice + 2).Text = RecuperaValor(CadenaSeleccion, 2) 'nombre variedad
End Sub

Private Sub mnBuscar_Click()
    BotonBuscar
End Sub

Private Sub mnEliminar_Click()
    BotonEliminar
End Sub

Private Sub mnImprimir_Click()
    AbrirListado (12)
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
        Case 8
                'MsgBox "Imprimir...under construction"
                mnImprimir_Click
    End Select
End Sub

Private Sub CargaGrid(Optional vSQL As String)
    Dim Sql As String
    Dim tots As String
    
'    adodc1.ConnectionString = Conn
    If vSQL <> "" Then
        Sql = CadenaConsulta & " AND " & vSQL
    Else
        Sql = CadenaConsulta
    End If
    
    If ParamVariedad <> "" Then Sql = Sql & " and rcalidad.codvarie = " & ParamVariedad
    
    
    '********************* canviar el ORDER BY *********************++
    Sql = Sql & " ORDER BY rcalidad.codvarie, rcalidad.codcalid"
    '**************************************************************++
    
    CargaGridGnral Me.DataGrid1, Me.adodc1, Sql, PrimeraVez
    
    ' *******************canviar els noms i si fa falta la cantitat********************
    tots = "S|txtAux(0)|T|Variedad|1000|;S|btnBuscar(0)|B|||;S|txtAux2(2)|T|Denominación|2700|;"
    tots = tots & "S|txtAux(1)|T|Calidad|850|;"
    tots = tots & "S|txtAux(2)|T|Descripción|2500|;"
    tots = tots & "S|txtAux(3)|T|Abrev|700|;N||||0|;S|Combo1(0)|C|Tipo|1220|;N||||0|;S|Combo1(1)|C|Tipo|1220|;"
    tots = tots & "S|txtAux(4)|T|Calibrador 1|1500|;S|txtAux(5)|T|Calibrador 2|1500|;N||||0|;S|chkAux(0)|CB|GR|360|;"
    tots = tots & "S|txtAux(6)|T|Eu.R.Soc|1200|;S|txtAux(7)|T|Eu.R.Coop|1200|;"
    '[Monica]21/01/2016:
    tots = tots & "N||||0|;S|chkAux(1)|CB|Bo|360|;"
    
    arregla tots, DataGrid1, Me, 350
    
    DataGrid1.ScrollBars = dbgAutomatic
    DataGrid1.Columns(0).Alignment = dbgRight
    DataGrid1.Columns(11).Alignment = dbgCenter
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
    If Not PerderFocoGnral(txtAux(Index), Modo) Then Exit Sub
    
    Select Case Index
        Case 0 'codigo de variedad
            If txtAux(Index).Text = "" Then Exit Sub
            txtAux2(2).Text = PonerNombreDeCod(txtAux(Index), "variedades", "nomvarie", "codvarie", "N")
        
        Case 1 'codigo de calibre
            PonerFormatoEntero txtAux(Index)
            
        Case 2, 3 'nombre de calibre y de calibre abreviado
            txtAux(Index).Text = UCase(txtAux(Index).Text)
            
        Case 6, 7 'precios
            PonerFormatoDecimal txtAux(Index), 6
    End Select
    
End Sub

Private Sub txtaux_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = teclaBuscar Then
        If Modo = 1 Or Modo = 3 Or Modo = 4 Then
            Select Case Index
                Case 0: KEYBusqueda KeyAscii, 0 'variedad
            End Select
        End If
    Else
        KEYpress KeyAscii
    End If
End Sub

Private Sub KEYBusqueda(KeyAscii As Integer, Indice As Integer)
    KeyAscii = 0
    btnBuscar_Click (Indice)
End Sub

Private Function DatosOK() As Boolean
'Dim Datos As String
Dim B As Boolean
Dim Sql As String
Dim Mens As String

    B = CompForm(Me)
    If Not B Then Exit Function
    
    If Modo = 3 Then   'Estamos insertando
        Sql = DevuelveDesdeBDNew(cAgro, "rcalidad", "codcalid", "codvarie", txtAux(0).Text, "N", , "codcalid", txtAux(1).Text, "N")
        If Sql <> "" Then
            MsgBox "Código de calidad existente para esta variedad. Reintroduzca.", vbExclamation
            PonerFoco txtAux(0)
            B = False
        End If
    End If
    
    If B And (Modo = 3 Or Modo = 4) Then
        Sql = "select count(*) from rcalidad where codvarie = " & DBSet(txtAux(0).Text, "N")
        Sql = Sql & " and codcalid <> " & DBSet(txtAux(1).Text, "N")
        Sql = Sql & " and nomcalibrador1 = " & DBSet(txtAux(4).Text, "T")
        '[Monica]23/01/2014: preguntamos si queremos continuar
        If TotalRegistros(Sql) <> 0 Then
            If MsgBox("El nombre que utiliza el Calibrador 1 está asignado a otra calidad. " & vbCrLf & "¿ Desea continuar ?" & vbCrLf, vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then
                PonerFoco txtAux(4)
                B = False
            End If
        End If
    End If
    
    If B And (Modo = 3 Or Modo = 4) Then
        '[Monica]23/01/2014: añado la condicion de que el nombre del calibrador 2 no sea blanco
        If txtAux(5).Text <> "" Then
            Sql = "select count(*) from rcalidad where codvarie = " & DBSet(txtAux(0).Text, "N")
            Sql = Sql & " and codcalid <> " & DBSet(txtAux(1).Text, "N")
            Sql = Sql & " and nomcalibrador2 = " & DBSet(txtAux(5).Text, "T")
        
            If TotalRegistros(Sql) <> 0 Then
                MsgBox "El nombre que utiliza el Calibrador 2 de esta calidad está asignada a otra. Revise.", vbExclamation
                PonerFoco txtAux(5)
                B = False
            End If
        End If
    End If
    
    ' solo puede haber una calidad de destrio en una variedad
    If B And (Modo = 3 Or Modo = 4) Then
        Sql = "select count(*) from rcalidad where codvarie = " & DBSet(txtAux(0).Text, "N")
        Sql = Sql & " and codcalid <> " & DBSet(txtAux(1).Text, "N")
        Sql = Sql & " and tipcalid = 1"
    
        If TotalRegistros(Sql) <> 0 And Combo1(0).ListIndex = 1 Then
            MsgBox "Sólo puede haber una calidad de destrio para la variedad. Revise.", vbExclamation
            Combo1(0).SetFocus
            B = False
        End If
    End If
    
    ' solo puede haber una calidad de venta campo en una variedad
    If B And (Modo = 3 Or Modo = 4) Then
        Sql = "select count(*) from rcalidad where codvarie = " & DBSet(txtAux(0).Text, "N")
        Sql = Sql & " and codcalid <> " & DBSet(txtAux(1).Text, "N")
        Sql = Sql & " and tipcalid = 2"
    
        If TotalRegistros(Sql) <> 0 And Combo1(0).ListIndex = 2 Then
            MsgBox "Sólo puede haber una calidad de venta campo para la variedad. Revise.", vbExclamation
            Combo1(0).SetFocus
            B = False
        End If
    End If
    
    ' solo puede haber una calidad de mermas en una variedad
    If B And (Modo = 3 Or Modo = 4) Then
        Sql = "select count(*) from rcalidad where codvarie = " & DBSet(txtAux(0).Text, "N")
        Sql = Sql & " and codcalid <> " & DBSet(txtAux(1).Text, "N")
        Sql = Sql & " and tipcalid = 3"
    
        If TotalRegistros(Sql) <> 0 And Combo1(0).ListIndex = 3 Then
            MsgBox "Sólo puede haber una calidad de mermas para la variedad. Revise.", vbExclamation
            Combo1(0).SetFocus
            B = False
        End If
    End If
    
    ' solo puede haber una calidad de pequeño en una variedad
    If B And (Modo = 3 Or Modo = 4) Then
        Sql = "select count(*) from rcalidad where codvarie = " & DBSet(txtAux(0).Text, "N")
        Sql = Sql & " and codcalid <> " & DBSet(txtAux(1).Text, "N")
        Sql = Sql & " and tipcalid = 4"
    
        If TotalRegistros(Sql) <> 0 And Combo1(0).ListIndex = 4 Then
            MsgBox "Sólo puede haber una calidad de pequeños para la variedad. Revise.", vbExclamation
            Combo1(0).SetFocus
            B = False
        End If
    End If
    
    
    
    
    DatosOK = B
End Function


Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
    PonerOpcionesMenuGeneralNew Me
End Sub


'
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


Private Sub txtAux_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Alvançar/Retrocedir els camps en les fleches de desplaçament del teclat.
    KEYdown KeyCode
End Sub

Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me

End Sub


Private Sub CargaCombo()

    On Error GoTo ErrCarga
    
    'Tipo de Calidad
    Combo1(0).Clear
    
    Combo1(0).AddItem "Normal"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 0
    Combo1(0).AddItem "Destrio"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 1
    Combo1(0).AddItem "Venta Campo"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 2
    Combo1(0).AddItem "Merma"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 3
    Combo1(0).AddItem "Pequeño"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 4
    Combo1(0).AddItem "Pixat"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 5
    Combo1(0).AddItem "Destrio Comercial"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 6
    
    
    ' Tipo de calidad 2
    Combo1(1).Clear
    Combo1(1).AddItem "Comercial"
    Combo1(1).ItemData(Combo1(1).NewIndex) = 0
    Combo1(1).AddItem "No Comercial"
    Combo1(1).ItemData(Combo1(1).NewIndex) = 1
    Combo1(1).AddItem "Retirada"
    Combo1(1).ItemData(Combo1(1).NewIndex) = 2
    
ErrCarga:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargar datos combo.", Err.Description
End Sub


Private Function SepuedeBorrar() As Boolean
Dim Sql As String
Dim cad As String

    SepuedeBorrar = False
    Sql = DevuelveDesdeBDNew(cAgro, "rcalidad_calibrador", "codcalid", "codvarie", adodc1.Recordset!Codvarie, "N", , "codcalid", adodc1.Recordset!codcalid, "N")
    If Sql <> "" Then
        cad = "No se puede eliminar la fila, "
        MsgBox cad & "tiene Calibradores asociados.", vbExclamation
        Exit Function
    End If
    SepuedeBorrar = True
End Function

