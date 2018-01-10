VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmTrzManMovim 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reparto Albaranes"
   ClientHeight    =   5925
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15225
   Icon            =   "frmTrzManMovim.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5925
   ScaleWidth      =   15225
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkAux 
      BackColor       =   &H80000005&
      Height          =   255
      Index           =   0
      Left            =   12420
      TabIndex        =   24
      Tag             =   "Es Merma|N|N|0|1|trzmovim|esmerma|||"
      Top             =   4620
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   7
      Left            =   5460
      MaxLength       =   6
      TabIndex        =   4
      Tag             =   "Cliente|N|S|0|999999|trzmovim|codclien|000000||"
      Top             =   4560
      Width           =   540
   End
   Begin VB.CommandButton btnBuscar 
      Appearance      =   0  'Flat
      Caption         =   "+"
      Height          =   300
      Index           =   0
      Left            =   6060
      MaskColor       =   &H00000000&
      TabIndex        =   23
      ToolTipText     =   "Buscar cliente"
      Top             =   4530
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   645
      Left            =   8490
      TabIndex        =   20
      Top             =   5100
      Width           =   3885
      Begin VB.TextBox txtAux2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   270
         Index           =   4
         Left            =   1860
         TabIndex        =   21
         Top             =   240
         Width           =   1830
      End
      Begin VB.Label Label1 
         Caption         =   "KILOS TOTALES: "
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   270
         TabIndex        =   22
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.TextBox txtAux2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   6390
      TabIndex        =   19
      Top             =   4560
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   6
      Left            =   1020
      MaxLength       =   13
      TabIndex        =   1
      Tag             =   "Numero Traza|T|S|||trzmovim|nrotraza|||"
      Top             =   4500
      Width           =   945
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   5
      Left            =   11340
      MaxLength       =   8
      TabIndex        =   8
      Tag             =   "Kilos|N|N|||trzmovim|kilos|###,##00||"
      Top             =   4620
      Width           =   945
   End
   Begin VB.CommandButton btnBuscar 
      Appearance      =   0  'Flat
      Caption         =   "+"
      Height          =   300
      Index           =   2
      Left            =   9090
      MaskColor       =   &H00000000&
      TabIndex        =   18
      ToolTipText     =   "Buscar variedad"
      Top             =   4590
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.TextBox txtAux2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   285
      Index           =   2
      Left            =   9360
      TabIndex        =   17
      Top             =   4620
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton btnBuscar 
      Appearance      =   0  'Flat
      Caption         =   "+"
      Height          =   300
      Index           =   1
      Left            =   4080
      MaskColor       =   &H00000000&
      TabIndex        =   16
      ToolTipText     =   "Buscar fecha"
      Top             =   4560
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   3
      Left            =   4440
      MaxLength       =   13
      TabIndex        =   3
      Tag             =   "Numero Albaran|N|S|||trzmovim|numalbar|000000||"
      Top             =   4560
      Width           =   945
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   2
      Left            =   8490
      MaxLength       =   6
      TabIndex        =   7
      Tag             =   "Variedad|N|N|0|999999|trzmovim|codvarie|000000|S|"
      Top             =   4620
      Width           =   540
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   12870
      TabIndex        =   6
      Top             =   5160
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   14010
      TabIndex        =   9
      Top             =   5160
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   1
      Left            =   3060
      MaxLength       =   16
      TabIndex        =   2
      Tag             =   "Fecha|F|N|||trzmovim|fecha|dd/mm/yyyy||"
      Top             =   4530
      Width           =   900
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   0
      Left            =   150
      MaxLength       =   6
      TabIndex        =   0
      Tag             =   "Nro Palet|N|N|0|999999|trzmovim|numpalet|000000||"
      Top             =   4500
      Width           =   800
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmTrzManMovim.frx":000C
      Height          =   4410
      Left            =   120
      TabIndex        =   12
      Top             =   540
      Width           =   14950
      _ExtentX        =   26379
      _ExtentY        =   7779
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
      Left            =   14010
      TabIndex        =   15
      Top             =   5160
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Index           =   1
      Left            =   120
      TabIndex        =   10
      Top             =   5190
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
         TabIndex        =   11
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
      TabIndex        =   13
      Top             =   0
      Width           =   15225
      _ExtentX        =   26855
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
            Object.ToolTipText     =   "Buscar"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Ver todos"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Nuevo"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Modificar"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Eliminar"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Kilos Merma"
            Object.Tag             =   "2"
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
         TabIndex        =   14
         Top             =   90
         Visible         =   0   'False
         Width           =   1215
      End
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   4
      Left            =   210
      MaxLength       =   10
      TabIndex        =   5
      Tag             =   "Codigo|N|N|||trzmovim|codigo|00000000|S|"
      Top             =   4560
      Width           =   945
   End
   Begin VB.Menu mnOpciones 
      Caption         =   "&Opciones"
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
      Begin VB.Menu mnKilosMerma 
         Caption         =   "Kilos Mermas"
         HelpContextID   =   2
         Shortcut        =   ^K
      End
      Begin VB.Menu mnImprimir 
         Caption         =   "&Imprimir"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnBarra3 
         Caption         =   "-"
      End
      Begin VB.Menu mnSalir 
         Caption         =   "&Salir"
         Shortcut        =   ^S
      End
   End
   Begin VB.Menu mnFiltro 
      Caption         =   "Filtro"
      Begin VB.Menu mnFiltro1 
         Caption         =   "Pendiente de Asignar"
      End
      Begin VB.Menu mnFiltro2 
         Caption         =   "Todo"
      End
   End
End
Attribute VB_Name = "frmTrzManMovim"
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

Private CadenaConsulta As String
Private CadB As String

Private WithEvents frmVar As frmManVariedad 'variedades
Attribute frmVar.VB_VarHelpID = -1
Private WithEvents frmC As frmCal 'calendario fecha
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmCli As frmBasico2 'Basico
Attribute frmCli.VB_VarHelpID = -1

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
Dim I As Integer

Dim CadenaFiltro As String
Dim Filtro As Byte

Dim Ordenacion As String


Private Sub PonerModo(vModo)
Dim B As Boolean

    Modo = vModo
    
    BuscaChekc = ""
    
    B = (Modo = 2)
    If B Then
        PonerContRegIndicador
    Else
        PonerIndicador lblIndicador, Modo
    End If
    
    For I = 0 To txtAux.Count - 1
        txtAux(I).visible = Not B
    Next I
    txtAux2(2).visible = Not B
    txtAux2(0).visible = Not B
    btnBuscar(0).visible = Not B
    btnBuscar(1).visible = Not B
    btnBuscar(2).visible = Not B
    chkAux(0).visible = Not B

    cmdAceptar.visible = Not B
    cmdCancelar.visible = Not B
    DataGrid1.Enabled = B
    
    'Si es regresar
    If DatosADevolverBusqueda <> "" Then cmdRegresar.visible = B
    
    PonerLongCampos
    PonerModoOpcionesMenu 'Activar/Desact botones de menu segun Modo
    PonerOpcionesMenu  'En funcion del usuario
    
End Sub


Private Sub PonerModoOpcionesMenu()
'Activa/Desactiva botones del la toobar y del menu, segun el modo en que estemos
Dim B As Boolean

    B = (Modo = 2)
    'Busqueda
    Toolbar1.Buttons(2).Enabled = B
    Me.mnBuscar.Enabled = B
    'Ver Todos
    Toolbar1.Buttons(3).Enabled = B
    Me.mnVerTodos.Enabled = B
    
    'Insertar
    Toolbar1.Buttons(6).Enabled = B And Not DeConsulta
    Me.mnNuevo.Enabled = B And Not DeConsulta
    
    B = (B And adodc1.Recordset.RecordCount > 0) And Not DeConsulta
    'Modificar
    Toolbar1.Buttons(7).Enabled = B
    Me.mnModificar.Enabled = B
    'Eliminar
    Toolbar1.Buttons(8).Enabled = B
    Me.mnEliminar.Enabled = B
    
    'Boton de Merma
    Toolbar1.Buttons(10).Enabled = B
    Me.mnKilosMerma.Enabled = B
    
    'Imprimir
    Toolbar1.Buttons(11).Enabled = False
    Me.mnImprimir.Enabled = False
    
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
        NumF = SugerirCodigoSiguienteStr("trzmovim", "codigo")
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
    For I = 0 To txtAux.Count - 1
        txtAux(I).Text = ""
    Next I
    txtAux2(2).Text = ""
    txtAux2(0).Text = ""
    txtAux(4).Text = NumF
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
    CargaGrid "trzmovim.codigo = -1"
    '*******************************************************************************
    'Buscar
    For I = 0 To txtAux.Count - 1
        txtAux(I).Text = ""
    Next I
    Me.chkAux(0).Value = 0
    
    Me.txtAux2(0).Text = ""
    Me.txtAux2(2).Text = ""
    
    LLamaLineas DataGrid1.Top + 206, 1 'Pone el form en Modo=1, Buscar
    PonerFoco txtAux(0)
End Sub


Private Sub BotonMermar()
Dim SQL As String
Dim NumF As Long
Dim Result As String
Dim Totales As Long
Dim resto As Long

    On Error GoTo eBotonMermar


   
    
    If ComprobarCero(DBLet(Me.adodc1.Recordset!numalbar, "N")) <> 0 Then
        MsgBox "Este movimiento está asociado a un albarán.", vbExclamation
        Exit Sub
    End If
    
    If DBLet(Me.adodc1.Recordset!esmerma) = 1 Then
        MsgBox "Este movimiento es de merma. No se puede realizar esta operación.", vbExclamation
        Exit Sub
    End If
    
    Result = InputBox("Kilos merma:", "Merma")
    If ComprobarCero(Result) > 0 Then
        
        If CLng(DBLet(Me.adodc1.Recordset!Kilos)) < CLng(ComprobarCero(Result)) Then
            MsgBox "Valor de kilos de mermados superior a los iniciales.", vbExclamation
            Exit Sub
        End If
        
        Screen.MousePointer = vbHourglass
        
        conn.BeginTrans
    
        Totales = DBLet(Me.adodc1.Recordset!Kilos, "N")
        resto = Totales - ComprobarCero(Result)
        'merma
        SQL = "update trzmovim set kilos = " & DBSet(Result, "N") & ", esmerma = 1 where codigo = " & DBSet(Me.adodc1.Recordset!Codigo, "N")
        conn.Execute SQL
        
        'resto
        NumF = DevuelveValor("select max(coalesce(codigo,0)) from trzmovim")
        NumF = NumF + 1
        
        SQL = "insert into trzmovim (codigo,numpalet,fecha,codvarie,kilos) select " & NumF & ",numpalet,fecha,codvarie," & DBSet(resto, "N")
        SQL = SQL & " from trzmovim where codigo = " & DBSet(Me.adodc1.Recordset!Codigo, "N")
        
        conn.Execute SQL
        
        conn.CommitTrans
        
        CargaGrid CadB
        
        
    End If
    
    Screen.MousePointer = vbDefault
    Exit Sub
    
eBotonMermar:
    Screen.MousePointer = vbDefault
    conn.RollbackTrans
    MuestraError Err.Number, "Inserción de Merma", Err.Description
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
    txtAux(4).Text = DataGrid1.Columns(0).Text
    txtAux(0).Text = DataGrid1.Columns(1).Text
    txtAux(6).Text = DataGrid1.Columns(2).Text
    txtAux(1).Text = DataGrid1.Columns(3).Text
    txtAux(3).Text = DataGrid1.Columns(4).Text
    txtAux(7).Text = DataGrid1.Columns(5).Text
    txtAux2(0).Text = DataGrid1.Columns(6).Text
    txtAux(2).Text = DataGrid1.Columns(7).Text
    txtAux2(2).Text = DataGrid1.Columns(8).Text
    txtAux(5).Text = DataGrid1.Columns(9).Text
    
    ' ***** canviar-ho pel nom del camp del combo *********
'    SelComboBool DataGrid1.Columns(2).Text, Combo1(0)
    ' *****************************************************

    LLamaLineas anc, 4 'Pone el form en Modo=4, Modificar
   
    'Como es modificar
    PonerFoco txtAux(0)
    Screen.MousePointer = vbDefault
End Sub

Private Sub LLamaLineas(alto As Single, xModo As Byte)
    DeseleccionaGrid Me.DataGrid1
    PonerModo xModo
    
    'Fijamos el ancho
    For I = 0 To txtAux.Count - 1
        If I <> 4 Then txtAux(I).Top = alto
    Next I
    txtAux2(2).Top = alto
    txtAux2(0).Top = alto
    btnBuscar(0).Top = alto - 15
    btnBuscar(1).Top = alto - 15
    btnBuscar(2).Top = alto - 15
    Me.chkAux(0).Top = alto
End Sub

Private Sub BotonEliminar()
Dim SQL As String
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
    SQL = "¿Seguro que desea eliminar el Movimiento?"
    SQL = SQL & vbCrLf & "Codigo:    " & adodc1.Recordset.Fields(0)
    SQL = SQL & vbCrLf & "Palet:    " & adodc1.Recordset.Fields(1)
    SQL = SQL & vbCrLf & "Variedad:   " & adodc1.Recordset.Fields(7) & " - " & adodc1.Recordset.Fields(8)
    SQL = SQL & vbCrLf & "Fecha: " & adodc1.Recordset.Fields(3)
    
    If MsgBox(SQL, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        NumRegElim = adodc1.Recordset.AbsolutePosition
        SQL = "Delete from trzmovim where codigo= " & adodc1.Recordset.Fields(0)
        
        
        conn.Execute SQL
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
Dim esq As Long
Dim dalt As Long
Dim menu As Long
Dim obj As Object
 
 
 TerminaBloquear
    
    Select Case Index
        Case 1 ' fecha
        
            Set frmC = New frmCal
            
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
        
            btnBuscar(1).Tag = Index '<===
            ' *** repasar si el camp es txtAux o Text1 ***
            If txtAux(Index).Text <> "" Then frmC.NovaData = txtAux(Index).Text
            ' ********************************************
        
            frmC.Show vbModal
            Set frmC = Nothing
            ' *** repasar si el camp es txtAux o Text1 ***
            PonerFoco txtAux(1) '<===
            ' *********************
                
        
        
        
        Case 2 'variedades
            Indice = Index
            Set frmVar = New frmManVariedad
            frmVar.DatosADevolverBusqueda = "0|1|"
            frmVar.CodigoActual = txtAux(Indice).Text
            frmVar.Show vbModal
            Set frmVar = Nothing
            PonerFoco txtAux(Indice)
        
        Case 0 ' cliente
            Set frmCli = New frmBasico2
            AyudaClienteCom frmCli, txtAux(7)
            Set frmCli = Nothing
            PonerFoco txtAux(7)
    
    End Select
    
    If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Me.adodc1, 1
End Sub


Private Sub chkAux_Click(Index As Integer)
    If Modo = 1 Then
        'Buscqueda
        If InStr(1, BuscaChekc, "chkAux(" & Index & ")") = 0 Then BuscaChekc = BuscaChekc & "chkAux(" & Index & ")|"
    End If
End Sub

Private Sub cmdAceptar_Click()
    Dim I As Integer

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
                    CadB = ""
                End If
            End If
            
        Case 4 'MODIFICAR
            If DatosOK Then
                If ModificaDesdeFormulario(Me) Then
                    TerminaBloquear
                    I = adodc1.Recordset.Fields(0)
                    PonerModo 2
                    CargaGrid CadB
'                    If CadB <> "" Then
'                        CargaGrid CadB
'                        lblIndicador.Caption = "BUSQUEDA: " & PonerContRegistros(Me.adodc1)
'                    Else
'                        CargaGrid
'                        lblIndicador.Caption = ""
'                    End If
                    adodc1.Recordset.Find (adodc1.Recordset.Fields(0).Name & " =" & I)
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

Private Sub DataGrid1_HeadClick(ByVal ColIndex As Integer)
Dim cad As String

If adodc1.Recordset Is Nothing Then Exit Sub
If adodc1.Recordset.EOF Then Exit Sub

Me.Refresh
DoEvents
Screen.MousePointer = vbHourglass

Ordenacion = "ORDER BY " & DataGrid1.Columns(ColIndex).DataField

'ColIndexAnt = ColIndex
CargaGrid CadB

Screen.MousePointer = vbDefault
End Sub

Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    PonerContRegIndicador
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
        .Buttons(10).Image = 28   ' kilos de merma
        .Buttons(11).Image = 10  'imprimir
        .Buttons(12).Image = 11  'Salir
    End With

    '## A mano
'    chkVistaPrevia.Value = CheckValueLeer(Name)
    
    
    '****************** canviar la consulta *********************************+
    CadenaConsulta = "SELECT trzmovim.codigo, trzmovim.numpalet, trzmovim.nrotraza, trzmovim.fecha, trzmovim.numalbar, trzmovim.codclien, clientes.nomclien, trzmovim.codvarie, variedades.nomvarie, "
    CadenaConsulta = CadenaConsulta & "trzmovim.kilos, esmerma, IF(esmerma=1,'*','') as desmerma "
    CadenaConsulta = CadenaConsulta & " FROM  variedades, trzmovim left join clientes on trzmovim.codclien = clientes.codclien  "
    CadenaConsulta = CadenaConsulta & " WHERE trzmovim.codvarie = variedades.codvarie "
    '************************************************************************
    
    LeerFiltro True
    PonerFiltro Filtro
    
    
    Ordenacion = " ORDER BY trzmovim.codigo "
    
    
    CadB = ""
    CargaGrid "trzmovim.codigo is null"
    
'    If (DatosADevolverBusqueda <> "") And NuevoCodigo <> "" Then
'        BotonAnyadir
'    Else
'        PonerModo 2
'    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
'    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
    If Modo = 4 Then TerminaBloquear
    LeerFiltro False
    
    Screen.MousePointer = vbDefault
End Sub


Private Sub frmVar_DatoSeleccionado(CadenaSeleccion As String)
'Variedades
    txtAux(Indice).Text = RecuperaValor(CadenaSeleccion, 1) 'codvarie
    txtAux2(Indice).Text = RecuperaValor(CadenaSeleccion, 2) 'nombre variedad
End Sub
Private Sub frmC_Selec(vFecha As Date)
    ' *** repasar si el camp es txtAux o Text1 ***
    txtAux(1).Text = Format(vFecha, "dd/mm/yyyy")  '<===
    ' ********************************************
End Sub

Private Sub frmCli_DatoSeleccionado(CadenaSeleccion As String)
'Form Mantenimiento de cliente
    txtAux(7).Text = RecuperaValor(CadenaSeleccion, 1) 'Codigo de clientes
    txtAux2(0).Text = RecuperaValor(CadenaSeleccion, 2) 'Nombre
End Sub

Private Sub mnBuscar_Click()
    BotonBuscar
End Sub

Private Sub mnEliminar_Click()
    BotonEliminar
End Sub

Private Sub mnFiltro1_Click()
    PonerFiltro 1
End Sub

Private Sub mnFiltro2_Click()
    PonerFiltro 2
End Sub


Private Sub PonerFiltro(NumFilt As Byte)
    Filtro = NumFilt
    Me.mnFiltro1.Checked = (NumFilt = 1)
    Me.mnFiltro2.Checked = (NumFilt = 2)
'    Me.mnFiltro3.Checked = (NumFilt = 3)
End Sub

Private Sub mnImprimir_Click()
    printNou
End Sub

Private Sub mnKilosMerma_Click()
    If adodc1.Recordset.EOF Then Exit Sub
    
    If adodc1.Recordset.RecordCount < 1 Then Exit Sub
    
    
    
    'Preparamos para modificar
    '-------------------------
    If BLOQUEADesdeFormulario2(Me, adodc1, 1) Then BotonMermar
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
        Case 10
                'MsgBox "Imprimir...under construction"
                mnKilosMerma_Click
        Case 11
                'MsgBox "Imprimir...under construction"
                mnImprimir_Click
        Case 12
                mnSalir_Click
    End Select
End Sub

Private Sub CargaGrid(Optional vSQL As String)
    Dim SQL As String
    Dim tots As String
    
'    adodc1.ConnectionString = Conn
    CadenaFiltro = AnyadeCadenaFiltro()
'    adodc1.ConnectionString = Conn
    
    If vSQL <> "" Then
        SQL = CadenaConsulta & " and " & CadenaFiltro & " AND " & vSQL
    Else
        SQL = CadenaConsulta & " and " & CadenaFiltro & "  "
    End If

    '********************* canviar el ORDER BY *********************++
    'SQL = SQL & " ORDER BY trzmovim.codigo "
    SQL = SQL & " " & Ordenacion
    '**************************************************************++
    
    CargaGridGnral Me.DataGrid1, Me.adodc1, SQL, PrimeraVez
    
    ' *******************canviar els noms i si fa falta la cantitat********************
    tots = "N|txtAux(0)|T|Codigo|1200|;"
    tots = tots & "S|txtAux(0)|T|Palet|1050|;S|txtAux(6)|T|Nro.Traza|2000|;S|txtAux(1)|T|Fecha|1200|;S|btnBuscar(1)|B|||;S|txtAux(3)|T|Albarán|1000|;S|txtAux(7)|T|Codigo|1000|;S|btnBuscar(0)|B|||;S|txtAux2(0)|T|Cliente|2700|;"
    tots = tots & "S|txtAux(2)|T|Código|1000|;S|btnBuscar(2)|B|||;S|txtAux2(2)|T|Nombre Variedad|2500|;"
    tots = tots & "S|txtAux(5)|T|Kilos|1500|;N||||0|;S|chkAux(0)|CB|Me|360|;"
    
    arregla tots, DataGrid1, Me
    
    DataGrid1.ScrollBars = dbgAutomatic
    DataGrid1.Columns(0).Alignment = dbgLeft
    DataGrid1.Columns(1).Alignment = dbgRight
    DataGrid1.Columns(4).Alignment = dbgLeft
    DataGrid1.Columns(5).Alignment = dbgLeft
    DataGrid1.Columns(6).Alignment = dbgLeft
    DataGrid1.Columns(7).Alignment = dbgLeft
'   DataGrid1.Columns(2).Alignment = dbgRight

    CalcularTotales SQL

End Sub

Private Sub txtaux_GotFocus(Index As Integer)
    ConseguirFocoLin txtAux(Index)
End Sub


Private Sub txtAux_LostFocus(Index As Integer)
Dim cadMen As String

    If Not PerderFocoGnral(txtAux(Index), Modo) Then Exit Sub
    
    Select Case Index
        Case 0, 3, 4, 5
            PonerFormatoEntero txtAux(Index)
        
        Case 1 'fecha
            PonerFormatoFecha txtAux(Index)
            
        Case 2 'codigo de variedad
            If PonerFormatoEntero(txtAux(Index)) Then
            txtAux2(Index).Text = PonerNombreDeCod(txtAux(Index), "variedades", "nomvarie", "codvarie", "N")
                If txtAux2(Index).Text = "" Then
                    cadMen = "No existe la Variedad: " & txtAux(Index).Text & vbCrLf
                    cadMen = cadMen & "¿Desea crearla?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmVar = New frmManVariedad
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
            
    End Select
    
End Sub

Private Function DatosOK() As Boolean
'Dim Datos As String
Dim B As Boolean
Dim SQL As String
Dim Mens As String


    B = CompForm(Me)
    If Not B Then Exit Function
    
    If Modo = 3 Then   'Estamos insertando
         SQL = ""
         SQL = DevuelveDesdeBDNew(cAgro, "codigoean", "codclien", "codclien", txtAux(0).Text, "N", , "codforfait", txtAux(1), "T", "codvarie", txtAux(2), "N")
         If SQL <> "" Then
            MsgBox "Código Ean existente para el cliente, forfait y variedad. Revise.", vbExclamation
            B = False
         End If
    End If
    
    DatosOK = B
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
        .cadTabla2 = "codigoean"
        .Informe2 = "rManCodEAN.rpt"
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
        .OtrosParametros2 = "pEmpresa='" & vEmpresa.nomempre & "'|pOrden={codigoean.codclien}|"
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

' ### [DavidV] 26/04/2006: Activar/desactivar la rueda del ratón.
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
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me
End Sub

Private Sub KEYBusqueda(KeyAscii As Integer, Indice As Integer)
    KeyAscii = 0
    btnBuscar_Click (Indice)
End Sub


Private Sub CalcularTotales(cadena As String)
Dim Kilos  As Currency
Dim Compleme As Currency
Dim Penaliza As Currency

Dim Rs As ADODB.Recordset
Dim SQL As String

    On Error Resume Next
    
    SQL = "select sum(kilos) kilos from (" & cadena & ") aaaaa"
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    txtAux2(4).Text = ""
    
    If TotalRegistrosConsulta(cadena) = 0 Then Exit Sub
    
    If Not Rs.EOF Then
        If Rs.Fields(0).Value <> 0 Then Kilos = DBLet(Rs.Fields(0).Value, "N") 'Solo es para saber que hay registros que mostrar
    
        txtAux2(4).Text = Format(Kilos, "###,###,###,##0")
    End If
    Rs.Close
    Set Rs = Nothing
    
    DoEvents
    
End Sub

Private Sub LeerFiltro(Leer As Boolean)
Dim SQL As String

    SQL = App.Path & "\filtrotrz.dat"
    If Leer Then
        Filtro = 3
        If Dir(SQL) <> "" Then
            AbrirFicheroFiltro True, SQL
            If IsNumeric(Trim(SQL)) Then Filtro = CByte(SQL)
        End If
    Else
        AbrirFicheroFiltro False, SQL
    End If
End Sub


Private Sub AbrirFicheroFiltro(Leer As Boolean, Fichero As String)
Dim SQL As String
Dim I As Integer

On Error GoTo EAbrir
    I = FreeFile
    If Leer Then
        Open Fichero For Input As #I
        Fichero = "3"
        Line Input #I, Fichero
    Else
        Open Fichero For Output As #I
        Print #I, Filtro
    End If
    Close #I
    Exit Sub
EAbrir:
    Err.Clear
End Sub

Private Function AnyadeCadenaFiltro() As String
Dim Aux As String
'Filtro = 1: pendiente de asignar
'Filtro = 2: sin filtro (todo)
    Aux = ""
    If Filtro = 1 Then
        'pendiente de asignar
        Aux = " numalbar is null and esmerma = 0"
    Else
        Aux = "(1=1)"
    End If  'filtro=0
    AnyadeCadenaFiltro = Aux
End Function

