VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmManPortesPobla 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Portes por Población"
   ClientHeight    =   6120
   ClientLeft      =   45
   ClientTop       =   30
   ClientWidth     =   10320
   Icon            =   "frmManPortesPobla.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6120
   ScaleWidth      =   10320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameBotonGnral 
      Height          =   705
      Left            =   90
      TabIndex        =   13
      Top             =   30
      Width           =   3585
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   210
         TabIndex        =   14
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
      Left            =   6240
      MaxLength       =   6
      TabIndex        =   2
      Tag             =   "Importe por Kg|N|N|0.0000|9.9999|rportespobla|imptrans|0.0000||"
      Top             =   4950
      Width           =   720
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
      Left            =   4035
      MaskColor       =   &H00000000&
      TabIndex        =   12
      ToolTipText     =   "Buscar Producto"
      Top             =   4950
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
      Index           =   0
      Left            =   1380
      TabIndex        =   11
      Top             =   4950
      Visible         =   0   'False
      Width           =   1815
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
      Left            =   3270
      MaxLength       =   6
      TabIndex        =   1
      Tag             =   "Producto|N|N|0|999|rportespobla|codprodu|000|S|"
      Top             =   4950
      Width           =   720
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
      Index           =   1
      Left            =   4305
      TabIndex        =   10
      Top             =   4950
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
      Height          =   330
      Index           =   0
      Left            =   1155
      MaskColor       =   &H00000000&
      TabIndex        =   9
      ToolTipText     =   "Buscar poblacion"
      Top             =   4950
      Visible         =   0   'False
      Width           =   195
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
      Left            =   7890
      TabIndex        =   3
      Top             =   5505
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
      Left            =   9045
      TabIndex        =   4
      Top             =   5505
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
      Index           =   0
      Left            =   165
      MaxLength       =   6
      TabIndex        =   0
      Tag             =   "Codigo Pueblo|T|N|||rportespobla|codpobla||S|"
      Top             =   4945
      Width           =   945
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmManPortesPobla.frx":000C
      Height          =   4545
      Left            =   120
      TabIndex        =   7
      Top             =   840
      Width           =   10015
      _ExtentX        =   17674
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
      Left            =   9060
      TabIndex        =   8
      Top             =   5505
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Index           =   1
      Left            =   120
      TabIndex        =   5
      Top             =   5460
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
         TabIndex        =   6
         Top             =   180
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
      Left            =   9660
      TabIndex        =   15
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
Attribute VB_Name = "frmManPortesPobla"
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

Private Const IdPrograma = 2022

Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)

Public ParamVariedad As String

'codi per al registe que s'afegix al cridar des d'atre formulari.
'Obrir en modo Insertar i tornar datos del registre insertat
Public NuevoCodigo As String
Public CodigoActual As String
Public DeConsulta As Boolean


Private WithEvents frmPob As frmManPueblos
Attribute frmPob.VB_VarHelpID = -1
Private WithEvents frmPro As frmBasico2 'Ayuda Productos de comercial)
Attribute frmPro.VB_VarHelpID = -1

Private CadenaConsulta As String
Private CadB As String


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
Dim I As Integer

Private Sub PonerModo(vModo)
Dim B As Boolean

    Modo = vModo
    
    B = (Modo = 2)
    If B Then
        PonerContRegIndicador lblIndicador, adodc1, CadB
    Else
        PonerIndicador lblIndicador, Modo
    End If
    
    For I = 0 To txtAux.Count - 1
        txtAux(I).visible = Not B
        txtAux(I).BackColor = vbWhite
    Next I
    
    txtAux2(0).visible = Not B
    btnBuscar(0).visible = Not B
    txtAux2(1).visible = Not B
    btnBuscar(1).visible = Not B
    
    cmdAceptar.visible = Not B
    cmdCancelar.visible = Not B
    DataGrid1.Enabled = B
    
    'Si es regresar
    If DatosADevolverBusqueda <> "" Then cmdRegresar.visible = B
    
    PonerLongCampos
    PonerModoOpcionesMenu 'Activar/Desact botones de menu segun Modo
    PonerOpcionesMenu  'En funcion del usuario
    
    'Si estamos modo Modificar bloquear clave primaria
    BloquearTxt txtAux(0), (Modo = 4)
    BloquearTxt txtAux(1), (Modo = 4)
    BloquearBtn Me.btnBuscar(0), Not (Modo = 1 Or Modo = 3)
    BloquearBtn Me.btnBuscar(1), Not (Modo = 1 Or Modo = 3)
    
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
    '******************** canviar taula i camp **************************
'    If (DatosADevolverBusqueda <> "") And NuevoCodigo <> "" Then
'        NumF = NuevoCodigo
'    Else
'        NumF = SugerirCodigoSiguienteStr("rpartida", "codparti")
'    End If
    '********************************************************************
    'Situamos el grid al final
    AnyadirLinea DataGrid1, adodc1
         
    anc = DataGrid1.Top
    If DataGrid1.Row < 0 Then
        anc = anc + 240
    Else
        anc = anc + DataGrid1.RowTop(DataGrid1.Row) + 5
    End If
'   txtAux(0).Text = NumF
    FormateaCampo txtAux(0)
    For I = 1 To txtAux.Count - 1
        txtAux(I).Text = ""
    Next I
    txtAux2(0).Text = ""
    txtAux2(1).Text = ""

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
    CargaGrid "rportespobla.codpobla = -1"
    '*******************************************************************************
    'Buscar
    For I = 0 To txtAux.Count - 1
        txtAux(I).Text = ""
    Next I
    LLamaLineas DataGrid1.Top + 240, 1 'Pone el form en Modo=1, Buscar
    PonerFoco txtAux(0)
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
        anc = DataGrid1.RowTop(DataGrid1.Row) + DataGrid1.Top '545
    End If

    'Llamamos al form
    txtAux(0).Text = DataGrid1.Columns(0).Text
    txtAux2(0).Text = DataGrid1.Columns(1).Text
    txtAux(1).Text = DataGrid1.Columns(2).Text
    txtAux2(1).Text = DataGrid1.Columns(3).Text
    txtAux(2).Text = DataGrid1.Columns(4).Text

    LLamaLineas anc, 4 'Pone el form en Modo=4, Modificar
   
    'Como es modificar
    PonerFoco txtAux(2)
    Screen.MousePointer = vbDefault
End Sub

Private Sub LLamaLineas(alto As Single, xModo As Byte)
    DeseleccionaGrid Me.DataGrid1
    PonerModo xModo
    
    'Fijamos el ancho
    For I = 0 To txtAux.Count - 1
        txtAux(I).Top = alto
    Next I
    
    ' ### [Monica] 12/09/2006
    txtAux2(0).Top = alto
    btnBuscar(0).Top = alto
    txtAux2(1).Top = alto
    btnBuscar(1).Top = alto
End Sub

Private Sub BotonEliminar()
Dim Sql As String
Dim temp As Boolean

    On Error GoTo Error2
    'Ciertas comprobaciones
    If adodc1.Recordset.EOF Then Exit Sub
'    If Not SepuedeBorrar Then Exit Sub
        
    '*************** canviar els noms i el DELETE **********************************
    Sql = "¿Seguro que desea eliminar el Registro?"
    Sql = Sql & vbCrLf & "Código: " & adodc1.Recordset.Fields(0)
    Sql = Sql & vbCrLf & "Descripción: " & adodc1.Recordset.Fields(1)
    
    If MsgBox(Sql, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        NumRegElim = adodc1.Recordset.AbsolutePosition
        Sql = "Delete from rportespobla where codpobla=" & adodc1.Recordset!CodPobla & " and codprodu=" & adodc1.Recordset!codprodu
        conn.Execute Sql
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
        Case 0 ' pueblos
            Indice = Index
            Set frmPob = New frmManPueblos
            frmPob.DatosADevolverBusqueda = "0|1|"
            frmPob.CodigoActual = txtAux(Indice).Text
            frmPob.Show vbModal
            Set frmPob = Nothing
            PonerFoco txtAux(Indice)
            
        Case 1 ' Productos
            Indice = Index
            Set frmPro = New frmBasico2
            AyudaProductosCom frmPro, txtAux(Indice).Text
'            frmPro.DatosADevolverBusqueda = "0|1|"
'            frmPro.CodigoActual = txtAux(indice).Text
'            frmPro.Show vbModal
            Set frmPro = Nothing
            PonerFoco txtAux(Indice)
    
    End Select
    
    If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Me.adodc1, 1
End Sub


Private Sub cmdAceptar_Click()
    Dim I As Integer

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
                             adodc1.Recordset.Find (adodc1.Recordset.Fields(2).Name & " =" & NuevoCodigo)
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
                    I = adodc1.Recordset.Fields(2)
                    PonerModo 2
                    CargaGrid CadB
                    adodc1.Recordset.Find (adodc1.Recordset.Fields(2).Name & " =" & I)
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
                SituarData Me.adodc1, "codpobla=" & CodigoActual, "", True
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
        'el 1 es separadors
        .Buttons(5).Image = 1   'Buscar
        .Buttons(6).Image = 2   'Todos
        'el 4 i el 5 son separadors
        .Buttons(1).Image = 3   'Insertar
        .Buttons(2).Image = 4   'Modificar
        .Buttons(3).Image = 5   'Borrar
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
    
    '****************** canviar la consulta *********************************+
    CadenaConsulta = "SELECT rportespobla.codpobla, rpueblos.despobla,  "
    CadenaConsulta = CadenaConsulta & "rportespobla.codprodu, productos.nomprodu, rportespobla.imptrans"
    CadenaConsulta = CadenaConsulta & " FROM rportespobla, rpueblos, productos"
    CadenaConsulta = CadenaConsulta & " WHERE rportespobla.codpobla = rpueblos.codpobla "
    CadenaConsulta = CadenaConsulta & " and rportespobla.codprodu = productos.codprodu "
    '************************************************************************
    
    CadB = ""
    CargaGrid
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
'    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
    Screen.MousePointer = vbDefault
    If Modo = 4 Then TerminaBloquear
End Sub

Private Sub frmPob_DatoSeleccionado(CadenaSeleccion As String)
'Pueblos
    txtAux(Indice).Text = RecuperaValor(CadenaSeleccion, 1) 'codpobla
    txtAux2(Indice).Text = RecuperaValor(CadenaSeleccion, 2) 'nombre poblacion
End Sub

Private Sub frmPro_DatoSeleccionado(CadenaSeleccion As String)
'productos
    txtAux(Indice).Text = RecuperaValor(CadenaSeleccion, 1) 'codprodu
    txtAux2(Indice).Text = RecuperaValor(CadenaSeleccion, 2) 'nombre producto
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
        Case 8
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
    
    '********************* canviar el ORDER BY *********************++
    Sql = Sql & " ORDER BY rportespobla.codpobla, rportespobla.codprodu"
    '**************************************************************++
    
    CargaGridGnral Me.DataGrid1, Me.adodc1, Sql, PrimeraVez
    
    ' *******************canviar els noms i si fa falta la cantitat********************
    tots = "S|txtAux(0)|T|C.Pueblo|1000|;S|btnBuscar(0)|B|||;"
    tots = tots & "S|txtAux2(0)|T|Población|3100|;S|txtAux(1)|T|Producto|1000|;"
    tots = tots & "S|btnBuscar(1)|B|||;S|txtAux2(1)|T|Denominación|3100|;S|txtAux(2)|T|Imp.Portes|1200|;"
    
    arregla tots, DataGrid1, Me, 350
    
    DataGrid1.ScrollBars = dbgAutomatic
    DataGrid1.Columns(0).Alignment = dbgLeft
    DataGrid1.Columns(4).Alignment = dbgRight
    
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

    If Not PerderFocoGnral(txtAux(Index), Modo) Then Exit Sub
    
    Select Case Index
        Case 0 'codigo de pueblo
            If txtAux(Index).Text = "" Then Exit Sub
            txtAux2(Index).Text = PonerNombreDeCod(txtAux(Index), "rpueblos", "despobla", "codpobla", "T")
            If txtAux2(Index).Text = "" Then
                cadMen = "No existe la Población: " & txtAux(Index).Text & vbCrLf
                cadMen = cadMen & "¿Desea crearla?" & vbCrLf
                If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                    Set frmPob = New frmManPueblos
                    frmPob.DatosADevolverBusqueda = "0|1|"
                    frmPob.NuevoCodigo = txtAux(Index).Text
                    txtAux(Index).Text = ""
                    TerminaBloquear
                    frmPob.Show vbModal
                    Set frmPob = Nothing
                    If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Me.adodc1, 1
                Else
                    txtAux(Index).Text = ""
                End If
                PonerFoco txtAux(Index)
            End If
            
        Case 1 'producto
            If PonerFormatoEntero(txtAux(Index)) Then
                txtAux2(Index).Text = PonerNombreDeCod(txtAux(Index), "productos", "nomprodu", "codprodu", "N")
                If txtAux2(Index).Text = "" And Modo = 3 Then
                    MsgBox "No existe el producto. Reintroduzca.", vbExclamation
                    PonerFoco txtAux(Index)
                End If
'                If txtAux2(Index).Text = "" Then
'                    cadMen = "No existe el Producto: " & txtAux(Index).Text & vbCrLf
'                    cadMen = cadMen & "¿Desea crearla?" & vbCrLf
'                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
'                        Set frmPro = New frmComProd
'                        frmPro.DatosADevolverBusqueda = "0|1|"
''                        frmVar.NuevoCodigo = txtAux(Index).Text
'                        txtAux(Index).Text = ""
'                        TerminaBloquear
'                        frmPro.Show vbModal
'                        Set frmPro = Nothing
'                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Me.adodc1, 1
'                    Else
'                        txtAux(Index).Text = ""
'                    End If
'                    PonerFoco txtAux(Index)
'                End If
            Else
                txtAux2(Index).Text = ""
            End If
            
        Case 2 ' importe por kilo
            If txtAux(Index).Text <> "" Then PonerFormatoDecimal txtAux(Index), 8
        
            
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
        Sql = ""
        Sql = DevuelveDesdeBDNew(cAgro, "rportespobla", "codpobla", "codpobla", txtAux(0).Text, "N", , "codprodu", txtAux(1).Text, "N")
        If Sql <> "" Then
            MsgBox "Ya existe la poblacion produto. Reintroduzca.", vbExclamation
            PonerFoco txtAux(0)
            B = False
        End If
    End If
    
    DatosOK = B
End Function


Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
    PonerOpcionesMenuGeneralNew Me
End Sub


Private Sub printNou()
    With frmImprimir2
        .cadTabla2 = "rportespobla"
        .Informe2 = "rManPortesPobla.rpt"
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
        .OtrosParametros2 = "pEmpresa='" & vEmpresa.nomempre & "'|pOrden={rportespobla.codpobla}|"
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
Private Sub txtaux_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = teclaBuscar Then
        If Modo = 1 Or Modo = 3 Or Modo = 4 Then
            Select Case Index
                Case 0: KEYBusqueda KeyAscii, 0 'poblacion
                Case 1: KEYBusqueda KeyAscii, 1 'producto
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



