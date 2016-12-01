VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmManHcoFrutaMon 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Histórico de Fruta"
   ClientHeight    =   5925
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14385
   Icon            =   "frmManHcoFrutaMon.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5925
   ScaleWidth      =   14385
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Height          =   765
      Left            =   7770
      TabIndex        =   23
      Top             =   5070
      Width           =   3945
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Left            =   90
         TabIndex        =   26
         Top             =   390
         Width           =   1215
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1350
         TabIndex        =   25
         Top             =   390
         Width           =   1215
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Left            =   2610
         TabIndex        =   24
         Top             =   390
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Cajas"
         Height          =   195
         Left            =   90
         TabIndex        =   29
         Top             =   150
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Kilos Netos"
         Height          =   195
         Left            =   1350
         TabIndex        =   28
         Top             =   150
         Width           =   1185
      End
      Begin VB.Label Label4 
         Caption         =   "Arrobas"
         Height          =   195
         Left            =   2610
         TabIndex        =   27
         Top             =   150
         Width           =   1215
      End
   End
   Begin VB.TextBox txtAux2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   10590
      TabIndex        =   21
      Top             =   4950
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.CommandButton btnBuscar 
      Appearance      =   0  'Flat
      Caption         =   "+"
      Height          =   300
      Index           =   2
      Left            =   2310
      MaskColor       =   &H00000000&
      TabIndex        =   20
      ToolTipText     =   "Buscar Fecha"
      Top             =   4950
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.TextBox txtAux1 
      Height          =   435
      Left            =   2610
      MaxLength       =   250
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Top             =   5280
      Width           =   5010
   End
   Begin VB.CommandButton btnBuscar 
      Appearance      =   0  'Flat
      Caption         =   "+"
      Height          =   300
      Index           =   1
      Left            =   6510
      MaskColor       =   &H00000000&
      TabIndex        =   19
      ToolTipText     =   "Buscar Variedad"
      Top             =   4920
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.TextBox txtAux2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   285
      Index           =   3
      Left            =   6780
      TabIndex        =   18
      Top             =   4920
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   5
      Left            =   9660
      MaxLength       =   7
      TabIndex        =   6
      Tag             =   "Peso Neto|N|N|||rhisfruta|kilosnet|###,##0||"
      Text            =   "P.Neto"
      Top             =   4950
      Width           =   900
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   4
      Left            =   8700
      MaxLength       =   7
      TabIndex        =   5
      Tag             =   "Nro.Cajas|N|N|||rhisfruta|numcajon|###,##0||"
      Text            =   "Numcajo"
      Top             =   4950
      Width           =   900
   End
   Begin VB.ComboBox Combo1 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   0
      ItemData        =   "frmManHcoFrutaMon.frx":000C
      Left            =   12150
      List            =   "frmManHcoFrutaMon.frx":000E
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Tag             =   "Tipo entrada|N|N|0|2|rhisfruta|tipoentr|||"
      Top             =   4920
      Width           =   1185
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   3
      Left            =   5580
      MaxLength       =   6
      TabIndex        =   4
      Tag             =   "Variedad|N|N|0|999999|rhisfruta|codvarie|000000||"
      Text            =   "Varie"
      Top             =   4950
      Width           =   900
   End
   Begin VB.TextBox txtAux2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   285
      Index           =   2
      Left            =   3690
      TabIndex        =   17
      Top             =   4950
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton btnBuscar 
      Appearance      =   0  'Flat
      Caption         =   "+"
      Height          =   300
      Index           =   0
      Left            =   3420
      MaskColor       =   &H00000000&
      TabIndex        =   16
      ToolTipText     =   "Buscar Socio"
      Top             =   4950
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   1
      Left            =   900
      MaxLength       =   10
      TabIndex        =   2
      Tag             =   "Fecha Albaran|F|N|||rhisfruta|fecalbar|dd/mm/yyyy||"
      Text            =   "Fecha"
      Top             =   4920
      Width           =   1395
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   11985
      TabIndex        =   9
      Top             =   5265
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   13140
      TabIndex        =   10
      Top             =   5265
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   2
      Left            =   2580
      MaxLength       =   12
      TabIndex        =   3
      Tag             =   "Socio|N|N|0|999999|rhisfruta|codsocio|000000||"
      Text            =   "Socio"
      Top             =   4950
      Width           =   810
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   0
      Left            =   30
      MaxLength       =   7
      TabIndex        =   1
      Tag             =   "Nro Albaran|N|S|0|9999999|rhisfruta|numalbar|0000000|S|"
      Text            =   "Albaran"
      Top             =   4920
      Width           =   800
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmManHcoFrutaMon.frx":0010
      Height          =   4410
      Left            =   180
      TabIndex        =   13
      Top             =   540
      Width           =   14040
      _ExtentX        =   24765
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
      Left            =   13125
      TabIndex        =   15
      Top             =   5265
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Index           =   1
      Left            =   120
      TabIndex        =   11
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
         TabIndex        =   12
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
      TabIndex        =   14
      Top             =   0
      Width           =   14385
      _ExtentX        =   25374
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   13
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
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Imprimir"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Impresión Etiquetas"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   195
         Left            =   5040
         TabIndex        =   0
         Top             =   90
         Visible         =   0   'False
         Width           =   1215
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Nombre Socio"
      Height          =   225
      Index           =   29
      Left            =   2640
      TabIndex        =   22
      Top             =   5070
      Width           =   1050
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
      Begin VB.Menu mnImprimir 
         Caption         =   "&Imprimir"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnImpresionEtiquetas 
         Caption         =   "Impresión Eti&quetas"
         Shortcut        =   ^Q
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
Attribute VB_Name = "frmManHcoFrutaMon"
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
Private WithEvents frmSoc As frmManSocios 'socios
Attribute frmSoc.VB_VarHelpID = -1
Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1

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
Dim indice As Byte 'Index del text1 on es poses els datos retornats des d'atres Formularis de Mtos
Dim i As Long

Dim CodTipoMov As String

Public ImpresoraDefecto As String


Private Sub PonerModo(vModo)
Dim b As Boolean

    Modo = vModo
    BuscaChekc = ""
    
    b = (Modo = 2)
    If b Then
        PonerContRegIndicador lblIndicador, adodc1, CadB
    Else
        PonerIndicador lblIndicador, Modo
    End If
    
    For i = 0 To txtAux.Count - 1
        txtAux(i).visible = Not b
    Next i
    txtAux1.Enabled = (Modo = 1 Or Modo = 3 Or Modo = 4)
    txtAux2(0).visible = Not b
    txtAux2(2).visible = Not b
    txtAux2(3).visible = Not b
    btnBuscar(0).visible = Not b
    btnBuscar(1).visible = Not b
    btnBuscar(2).visible = Not b
    Combo1(0).visible = Not b

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
    'Imprimir etiquetas
    Toolbar1.Buttons(12).Enabled = b
    Me.mnImpresionEtiquetas.Enabled = b
    
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
'        NumF = SugerirCodigoSiguienteStr("rhisfruta", "numalbar")
'    End If
'    '********************************************************************
    'Situamos el grid al final
    AnyadirLinea DataGrid1, adodc1
         
    anc = DataGrid1.Top
    If DataGrid1.Row < 0 Then
        anc = anc + 206
    Else
        anc = anc + DataGrid1.RowTop(DataGrid1.Row) + 5
    End If
    txtAux(0).Text = ""
    FormateaCampo txtAux(0)
    For i = 1 To txtAux.Count - 1
        txtAux(i).Text = ""
    Next i
    
    txtAux1.Text = ""
    
    txtAux(1).Text = Format(Now, "dd/mm/yyyy")
    
    txtAux2(0).Text = ""
    txtAux2(2).Text = ""
    txtAux2(3).Text = ""
    
    Combo1(0).ListIndex = 0

    LLamaLineas anc, 3 'Pone el form en Modo=3, Insertar
       
    'Ponemos el foco
    PonerFoco txtAux(1)
End Sub

Private Sub BotonVerTodos()
    CadB = ""
    CargaGrid ""
    PonerModo 2
End Sub

Private Sub BotonBuscar()
    ' ***************** canviar per la clau primaria ********
    CargaGrid "rhisfruta.numalbar is null"
    '*******************************************************************************
    'Buscar
    For i = 0 To txtAux.Count - 1
        txtAux(i).Text = ""
    Next i
    txtAux1.Text = ""
    txtAux2(0).Text = ""
    txtAux2(2).Text = ""
    txtAux2(3).Text = ""
    Combo1(0).ListIndex = -1
    
    LLamaLineas DataGrid1.Top + 206, 1 'Pone el form en Modo=1, Buscar
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
        anc = DataGrid1.RowTop(DataGrid1.Row) + 545
    End If

    'Llamamos al form
    txtAux(0).Text = DataGrid1.Columns(0).Text
    txtAux(1).Text = DataGrid1.Columns(1).Text
    txtAux(2).Text = DataGrid1.Columns(2).Text
    txtAux2(2).Text = DataGrid1.Columns(3).Text
    txtAux(3).Text = DataGrid1.Columns(4).Text
    txtAux2(3).Text = DataGrid1.Columns(5).Text
    txtAux(4).Text = DataGrid1.Columns(6).Text
    txtAux(5).Text = DataGrid1.Columns(7).Text
    txtAux2(0).Text = DataGrid1.Columns(8).Text
'    txtAux(6).Text = DataGrid1.Columns(9).Text
    
    ' ***** canviar-ho pel nom del camp del combo *********
    i = adodc1.Recordset!TipoEntr
    ' *****************************************************
    PosicionarCombo Me.Combo1(0), i
    
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
    txtAux2(3).Top = alto
    txtAux2(0).Top = alto
    btnBuscar(0).Top = alto - 15
    btnBuscar(1).Top = alto - 15
    btnBuscar(2).Top = alto - 15
    Combo1(0).Top = alto - 15
    
End Sub


Private Sub BotonEliminar()
Dim Sql As String
Dim temp As Boolean
Dim vTipoMov As CTiposMov

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
    Sql = "¿Seguro que desea eliminar la Entrada?"
    Sql = Sql & vbCrLf & "Albarán: " & adodc1.Recordset.Fields(0)
    Sql = Sql & vbCrLf & "Fecha: " & adodc1.Recordset.Fields(1)
    
    If MsgBox(Sql, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        NumRegElim = adodc1.Recordset.AbsolutePosition
        
        Sql = "delete from rhisfruta_clasif where numalbar = " & adodc1.Recordset!numalbar
        conn.Execute Sql
        
        Sql = "delete from rhisfruta_entradas where numalbar = " & adodc1.Recordset!numalbar
        conn.Execute Sql
        
        Sql = "Delete from rhisfruta where numalbar=" & adodc1.Recordset!numalbar
        conn.Execute Sql
        
        conn.Execute "DELETE FROM trzpalets where numnotac = " & Trim(CStr(adodc1.Recordset!numalbar))

        
        
        'Devolvemos contador, si no estamos actualizando
        Set vTipoMov = New CTiposMov
        vTipoMov.DevolverContador CodTipoMov, adodc1.Recordset!numalbar
        Set vTipoMov = Nothing
        
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
        Case 0 'socios
            indice = 2
            PonerFoco txtAux(indice)
            Set frmSoc = New frmManSocios
            frmSoc.DatosADevolverBusqueda = "0|1|"
            frmSoc.Show vbModal
            Set frmSoc = Nothing
            PonerFoco txtAux(indice)
        
        Case 1 'variedades de comercial
            indice = 3
            Set frmVar = New frmComVar
            frmVar.DatosADevolverBusqueda = "0|1|"
            frmVar.CodigoActual = txtAux(indice).Text
            frmVar.Show vbModal
            Set frmVar = Nothing
            PonerFoco txtAux(indice)
    
        Case 2 ' fecha
            Dim esq As Long
            Dim dalt As Long
            Dim menu As Long
            Dim obj As Object
        
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
        
            btnBuscar(0).Tag = Index '<===
            ' *** repasar si el camp es txtAux o Text1 ***
            If txtAux(1).Text <> "" Then frmC.NovaData = txtAux(1).Text
        
            frmC.Show vbModal
            Set frmC = Nothing
            ' *** repasar si el camp es txtAux o Text1 ***
            PonerFoco txtAux(1) '<===
            ' ********************************************
        
    End Select
    
    If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Me.adodc1, 1
End Sub


Private Sub cmdAceptar_Click()
    Dim i As Long

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
            If DatosOk Then InsertarCabecera
                
        Case 4 'MODIFICAR
            If DatosOk Then
                If ModificaCabecera Then
                    TerminaBloquear
                    i = adodc1.Recordset.Fields(0)
                    PonerModo 2
                    CargaGrid CadB
                    
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

Private Sub Combo1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub DataGrid1_DblClick()
    If cmdRegresar.visible Then cmdRegresar_Click
End Sub

Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If Modo = 2 Then
        PonerContRegIndicador lblIndicador, adodc1, CadB
        If Not adodc1.Recordset.EOF Then
            txtAux1.Text = DevuelveDesdeBDNew(cAgro, "rhisfruta_entradas", "observac", "numalbar", adodc1.Recordset!numalbar, "N")
        End If
    End If
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
                SituarData Me.adodc1, "numalbar=" & CodigoActual, "", True
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
        .Buttons(2).Image = 1   'Buscar
        .Buttons(3).Image = 2   'Todos
        'el 4 i el 5 son separadors
        .Buttons(6).Image = 3   'Insertar
        .Buttons(7).Image = 4   'Modificar
        .Buttons(8).Image = 5   'Borrar
        'el 9 i el 10 son separadors
        .Buttons(11).Image = 10  'imprimir
        .Buttons(12).Image = 10  'imprimir etiquetas
        .Buttons(13).Image = 11  'Salir
    End With

    '## A mano
'    chkVistaPrevia.Value = CheckValueLeer(Name)
    
    CargaCombo
    
    '****************** canviar la consulta *********************************
    CadenaConsulta = "SELECT rhisfruta.numalbar, rhisfruta.fecalbar, rhisfruta.codsocio, rsocios.nomsocio,"
    CadenaConsulta = CadenaConsulta & "rhisfruta.codvarie, variedades.nomvarie, rhisfruta.numcajon,"
    CadenaConsulta = CadenaConsulta & "rhisfruta.kilosnet, round(rhisfruta.kilosnet / 13, 2) arrobas,  tipoentr, "
    CadenaConsulta = CadenaConsulta & " CASE rhisfruta.tipoentr WHEN 0 THEN ""Normal"" WHEN 1 THEN ""Venta Campo""  END "
    CadenaConsulta = CadenaConsulta & " FROM variedades, rhisfruta,  rsocios"
    CadenaConsulta = CadenaConsulta & " WHERE variedades.codvarie = rhisfruta.codvarie and "
    CadenaConsulta = CadenaConsulta & " rhisfruta.codsocio = rsocios.codsocio "
    '************************************************************************
    
    CadB = ""
    CargaGrid
    
    CodTipoMov = "ALF"

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

Private Sub frmC_Selec(vFecha As Date)
    ' *** repasar si el camp es txtAux o Text1 ***
    txtAux(1).Text = Format(vFecha, "dd/mm/yyyy") '<===
    ' ********************************************
End Sub

Private Sub frmSoc_DatoSeleccionado(CadenaSeleccion As String)
    txtAux(indice).Text = RecuperaValor(CadenaSeleccion, 1) 'codvarie
    txtAux2(indice).Text = RecuperaValor(CadenaSeleccion, 2) 'nombre variedad
End Sub

Private Sub frmVar_DatoSeleccionado(CadenaSeleccion As String)
'Variedad comercial
    txtAux(indice).Text = RecuperaValor(CadenaSeleccion, 1) 'codvarie
    txtAux2(indice).Text = RecuperaValor(CadenaSeleccion, 2) 'nombre variedad
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
    
    '[Monica]21/05/2013: si la entrada ya tiene las etiquetas creadas o ya ha sido volcada doy un aviso
    '                    y dejo continuar
    If EstaVolcado(adodc1.Recordset!numalbar) Then
        If MsgBox("Esta entrada ya ha sido volcada. ¿ Continuar ? ", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
    End If
    
    'Preparamos para modificar
    '-------------------------
    If BLOQUEADesdeFormulario2(Me, adodc1, 1) Then BotonModificar
End Sub

Private Function EstaVolcado(Albaran As String) As Boolean
Dim Sql As String
    
    Sql = "select count(*) from trzlineas_cargas where idpalet in (select idpalet from trzpalets where numnotac = " & DBSet(Albaran, "N") & ")"
    
    EstaVolcado = (TotalRegistros(Sql) <> 0)

End Function


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
                mnImpresionEtiquetas_Click
        Case 13
                mnSalir_Click
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
    
    If ParamVariedad <> "" Then Sql = Sql & " and rhisfruta.codvarie = " & ParamVariedad
    
    
    '********************* canviar el ORDER BY *********************++
    Sql = Sql & " ORDER BY rhisfruta.numalbar"
    '**************************************************************++
    
    CargaGridGnral Me.DataGrid1, Me.adodc1, Sql, PrimeraVez
    
    ' *******************canviar els noms i si fa falta la cantitat********************
    tots = "S|txtAux(0)|T|Albarán|800|;S|txtAux(1)|T|Fecha|1100|;S|btnBuscar(2)|B|||;"
    tots = tots & "S|txtAux(2)|T|Socio|800|;S|btnBuscar(0)|B|||;S|txtAux2(2)|T|Denominación|3550|;"
    tots = tots & "S|txtAux(3)|T|Código|800|;S|btnBuscar(1)|B|||;S|txtAux2(3)|T|Variedad|1600|;"
    tots = tots & "S|txtAux(4)|T|Cajas|1100|;S|txtAux(5)|T|Kilos|1100|;S|txtAux2(0)|T|Arrobas|1100|;"
    tots = tots & "N||||0|;S|Combo1(0)|C|Tipo Entrada|1520|;"
    
    arregla tots, DataGrid1, Me
    
    DataGrid1.ScrollBars = dbgAutomatic
    DataGrid1.Columns(0).Alignment = dbgRight
    DataGrid1.Columns(2).Alignment = dbgLeft
    DataGrid1.Columns(4).Alignment = dbgLeft
    
    DataGrid1.Columns(8).Alignment = dbgRight
    DataGrid1.Columns(9).Alignment = dbgRight
    DataGrid1.Columns(10).Alignment = dbgCenter
    
    If Not adodc1.Recordset.EOF Then
        txtAux1.Text = DevuelveDesdeBDNew(cAgro, "rhisfruta_entradas", "observac", "numalbar", adodc1.Recordset!numalbar, "N")
    End If
    
    CalcularTotales Sql
    
    
'    DataGrid1.Columns(11).Alignment = dbgCenter
End Sub

Private Sub txtAux_GotFocus(Index As Integer)
    ConseguirFocoLin txtAux(Index)
End Sub


Private Sub txtAux_LostFocus(Index As Integer)
Dim cadMen As String
Dim Kilos As Currency
Dim KilosCajon As Currency

    If Not PerderFocoGnral(txtAux(Index), Modo) Then Exit Sub
    
    Select Case Index
        Case 0 ' numero de albaran
            PonerFormatoEntero txtAux(0)
            
        Case 1 ' fecha
            PonerFormatoFecha txtAux(Index)
            
        Case 2 'codigo de socio
            If txtAux(Index).Text = "" Then Exit Sub
            txtAux2(Index).Text = PonerNombreDeCod(txtAux(Index), "rsocios", "nomsocio", "codsocio", "N")
            If Modo = 1 Then Exit Sub
            If txtAux2(Index).Text = "" Then
                MsgBox "No existe el Socio. Revise.", vbExclamation
                PonerFoco txtAux(Index)
            End If
    
        Case 3 'codigo de variedad
            If txtAux(Index).Text = "" Then Exit Sub
            txtAux2(Index).Text = PonerNombreDeCod(txtAux(Index), "variedades", "nomvarie", "codvarie", "N")
            If Modo = 1 Then Exit Sub
            If txtAux2(Index).Text = "" Then
                MsgBox "No existe la Variedad. Revise.", vbExclamation
                PonerFoco txtAux(Index)
            End If
            
        Case 4, 5 'caja y kilos
            PonerFormatoEntero txtAux(Index)
            If Index = 5 And txtAux(Index).Text <> "" Then
                txtAux2(0).Text = Round2(CCur(ImporteSinFormato(txtAux(Index).Text)) / 13, 2)
            End If
            
        Case 6 ' precio de contrato
            PonerFormatoDecimal txtAux(Index), 11
            
    End Select
    
    '[Monica]19/10/2015: si estamos modificando no se deben de cambiar los kilos (añadido modo <> 4)
    If (Index = 3 Or Index = 4) And Modo <> 1 And Modo <> 4 Then
        KilosCajon = 0
        If txtAux(3).Text <> "" Then
            KilosCajon = DevuelveValor("select kgscajon from variedades where codvarie = " & DBSet(txtAux(3).Text, "N"))
        End If
        Kilos = Round2(ComprobarCero(txtAux(4).Text) * KilosCajon, 0)
        txtAux(5).Text = Format(Kilos, "###,##0")
    End If
   
End Sub


Private Sub txtAux1_GotFocus()
    ConseguirFocoLin txtAux1
End Sub


Private Sub txtAux1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtAux1_LostFocus()
Dim cadMen As String

    If Not PerderFocoGnral(txtAux1, Modo) Then Exit Sub
    
    txtAux1.Text = UCase(txtAux1.Text)
    
End Sub




Private Function DatosOk() As Boolean
'Dim Datos As String
Dim b As Boolean
Dim Sql As String
Dim Mens As String

    b = CompForm(Me)
    If Not b Then Exit Function
    
    If Modo = 3 Then   'Estamos insertando
        Sql = DevuelveDesdeBDNew(cAgro, "rcalidad", "codcalid", "codvarie", txtAux(3).Text, "N")
        If Sql = "" Then
            MsgBox "No existe código de calidad para esta variedad. Reintroduzca.", vbExclamation
            PonerFoco txtAux(0)
            b = False
        End If
    End If
    
    
    DatosOk = b
End Function


Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
    PonerOpcionesMenuGeneralNew Me
End Sub


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
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me

End Sub

Private Sub KEYBusqueda(KeyAscii As Integer, indice As Integer)
    KeyAscii = 0
    btnBuscar_Click (indice)
End Sub


Private Sub CargaCombo()

    On Error GoTo ErrCarga
    
    'Tipo de Calidad
    Combo1(0).Clear
    
    Combo1(0).AddItem "Normal"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 0
    Combo1(0).AddItem "Venta Campo"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 1

ErrCarga:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargar datos combo.", Err.Description
End Sub


Private Function SepuedeBorrar() As Boolean
Dim Sql As String
Dim Cad As String

    SepuedeBorrar = False
    Sql = DevuelveDesdeBDNew(cAgro, "rfactsoc_albaran", "numfactu", "numalbar", adodc1.Recordset!numalbar, "N")
    If Sql <> "" Then
        Cad = "No se puede eliminar el albarán, "
        MsgBox Cad & "está facturado al socio.", vbExclamation
        Exit Function
        
    Else
        'miramos si el albaran ha sido volcado
        If EstaVolcado(CStr(adodc1.Recordset!numalbar)) Then
            Cad = "No se puede eliminar el albarán, "
            MsgBox Cad & "está volcado en traza.", vbExclamation
            Exit Function
        End If
    End If
    SepuedeBorrar = True
End Function


Private Sub InsertarCabecera()
Dim vTipoMov As CTiposMov 'Clase Tipo Movimiento
Dim Sql As String

    On Error GoTo EInsertarCab
    
    Set vTipoMov = New CTiposMov
    If vTipoMov.Leer(CodTipoMov) Then
        If txtAux(0).Text = "" Then txtAux(0).Text = vTipoMov.ConseguirContador(CodTipoMov)
        Sql = CadenaInsertarDesdeForm(Me)
        If Sql <> "" Then
            If InsertarOferta(Sql, vTipoMov) Then
                
                    i = txtAux(0).Text
                    CargaGrid CadB
                    If (DatosADevolverBusqueda <> "") And NuevoCodigo <> "" Then
                        cmdCancelar_Click
'                        If Not adodc1.Recordset.EOF Then adodc1.Recordset.MoveLast
                        If Not adodc1.Recordset.EOF Then
                            adodc1.Recordset.Find (adodc1.Recordset.Fields(0).Name & " =" & i)
                        End If
                        cmdRegresar_Click
                    Else
                        BotonAnyadir
                    End If
'[Monica]12/06/2013: quito esta instruccion pq quieren mantener la busqueda
'                    CadB = ""
                
'
'
'                CadenaConsulta = "Select * from rhisfruta where numalbar = " & DBSet(txtAux(0).Text, "N")
'                PonerCadenaBusqueda
'                PonerModo 2
            End If
        End If
        txtAux(0).Text = Format(txtAux(0).Text, "0000000")
    End If
    Set vTipoMov = Nothing
    
EInsertarCab:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub



Private Function InsertarOferta(vSQL As String, vTipoMov As CTiposMov) As Boolean
Dim MenError As String
Dim bol As Boolean, Existe As Boolean
Dim cambiaSQL As Boolean
Dim devuelve As String
Dim Sql2 As String

    On Error GoTo EInsertarOferta
    
    bol = True
    
    cambiaSQL = False
    'Comprobar si mientras tanto se incremento el contador de Albaranes
    'para ello vemos si existe una oferta con ese contador y si existe la incrementamos
    Do
        devuelve = DevuelveDesdeBDNew(cAgro, "rhisfruta", "numalbar", "numalbar", txtAux(0).Text, "N")
        If devuelve <> "" Then
            'Ya existe el contador incrementarlo
            Existe = True
            vTipoMov.IncrementarContador (CodTipoMov)
            txtAux(0).Text = vTipoMov.ConseguirContador(CodTipoMov)
            cambiaSQL = True
        Else
            Existe = False
            cambiaSQL = True
        End If
    Loop Until Not Existe
    If cambiaSQL Then vSQL = CadenaInsertarDesdeForm(Me)
    
    
    'Aqui empieza transaccion
    conn.BeginTrans
    MenError = "Error al insertar en la tabla de Albaranes (rhisfruta)."
    conn.Execute vSQL, , adCmdText
    
    Sql2 = "update rhisfruta set kilosbru = kilosnet, kilostra = kilosnet, codcampo = 0 "
    Sql2 = Sql2 & " where numalbar = " & DBSet(txtAux(0).Text, "N")
    conn.Execute Sql2
    
    
    '[Monica]12/04/2013: insertamos automaticamente en la tabla de lineas
    Sql2 = "insert into rhisfruta_entradas (numalbar,numnotac,fechaent,horaentr,kilosbru,numcajon,kilosnet,kilostra,observac) values ("
    Sql2 = Sql2 & DBSet(txtAux(0).Text, "N") & "," & DBSet(txtAux(0).Text, "N") & "," & DBSet(txtAux(1).Text, "F") & ","
    Sql2 = Sql2 & "'" & Format(txtAux(1).Text, "yyyy-mm-dd") & " " & Format(Now, "hh:mm:ss") & "'," & DBSet(txtAux(5).Text, "N") & ","
    Sql2 = Sql2 & DBSet(txtAux(4).Text, "N") & "," & DBSet(txtAux(5).Text, "N") & "," & DBSet(txtAux(5).Text, "N") & "," & DBSet(txtAux1.Text, "T") & ")"
    
    conn.Execute Sql2
    
    '[Monica]12/04/2013: insertamos automaticamente en la tabla de lineas
    Sql2 = "insert into rhisfruta_clasif (numalbar,codvarie,codcalid,kilosnet) values ("
    Sql2 = Sql2 & DBSet(txtAux(0).Text, "N") & "," & DBSet(txtAux(3).Text, "N") & ",1," & DBSet(txtAux(5).Text, "N") & ")"
    
    conn.Execute Sql2
    
   
    If txtAux(0).Text = vTipoMov.Contador + 1 Then
        MenError = "Error al actualizar el contador del Albarán."
        vTipoMov.IncrementarContador (CodTipoMov)
    End If

EInsertarOferta:
    If Err.Number <> 0 Then
        MenError = "Insertando Albarán." & vbCrLf & "----------------------------" & vbCrLf & MenError
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


Private Function ModificaCabecera() As Boolean
Dim b As Boolean
Dim MenError As String

    On Error GoTo EModificarCab

    conn.BeginTrans
    
    
    
    MenError = "Modificando Lineas clasificacion: "
    b = ModificandoClasificacion(adodc1.Recordset.Fields(0), "", MenError)

    If b Then b = ModificaDesdeFormulario(Me)
    
    
    '[Monica]08/02/2012: Si han modificado variedad socio o fecha en traza
    If CLng(adodc1.Recordset!codvarie) <> CLng(txtAux(3).Text) Or CLng(adodc1.Recordset!Codsocio) <> CLng(txtAux(2).Text) Or _
       DBLet(adodc1.Recordset!Fecalbar, "F") <> CDate(txtAux(1).Text) Then
         MenError = "No se han realizado los cambios en Trazabilidad. " & vbCrLf
         If Not ActualizarTraza2(txtAux(0).Text, txtAux(3).Text, txtAux(2).Text, txtAux(1).Text, MenError) Then
            MsgBox MenError, vbExclamation
         End If
    End If
    
    MenError = "Actualizando Paletización: "
    If b Then b = ActualizarPaletizacion(MenError)
    

                
EModificarCab:
    If Err.Number <> 0 Or Not b Then
        MenError = "Modificando Albarán." & vbCrLf & "----------------------------" & vbCrLf & MenError
        MuestraError Err.Number, MenError, Err.Description
        b = False
    End If
    If b Then
        ModificaCabecera = True
        conn.CommitTrans
    Else
        ModificaCabecera = False
        conn.RollbackTrans
    End If
End Function


Private Function ModificandoClasificacion(numalbar As String, Variedad As String, Mens As String) As Boolean
Dim Sql As String

    On Error GoTo eModificandoClasificacion

    ModificandoClasificacion = False

    Sql = "update rhisfruta_clasif set codvarie = " & DBSet(txtAux(3).Text, "N")
    Sql = Sql & " , kilosnet = " & DBSet(txtAux(5).Text, "N")
    Sql = Sql & " where numalbar = " & DBSet(numalbar, "N")
    
    conn.Execute Sql
    
    Sql = "update rhisfruta_entradas set fechaent = " & DBSet(txtAux(1).Text, "F")
    Sql = Sql & ",horaentr = '" & Format(txtAux(1).Text, "yyyy-mm-dd") & " " & Format(Now, "hh:mm:ss") & "'"
    Sql = Sql & ",kilosbru = " & DBSet(txtAux(5).Text, "N")
    Sql = Sql & ",numcajon = " & DBSet(txtAux(4).Text, "N")
    Sql = Sql & ",kilosnet = " & DBSet(txtAux(5).Text, "N")
    Sql = Sql & ",kilostra = " & DBSet(txtAux(5).Text, "N")
    Sql = Sql & ",observac = " & DBSet(txtAux1.Text, "T")
    Sql = Sql & " where numalbar = " & DBSet(numalbar, "N")
    
    conn.Execute Sql

    ModificandoClasificacion = True
    Exit Function
    
eModificandoClasificacion:
    Mens = Mens & vbCrLf & Err.Description
End Function

Private Sub printNou()
    With frmImprimir2
        .cadTabla2 = "rhisfruta"
        .Informe2 = "rManHcoFrutaMon.rpt"
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
        .OtrosParametros2 = "pEmpresa='" & vEmpresa.nomempre & "'|pOrden={rhisfruta.numalbar}|"
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

Private Sub CalcularTotales(cadena As String)
Dim Numcajon  As Currency
Dim KilosNet As Currency
Dim Arrobas As Currency

Dim Rs As ADODB.Recordset
Dim Sql As String

    On Error Resume Next
    
    Sql = "select sum(numcajon) numcajon , sum(kilosnet) kilosnet, sum(round(kilosnet/13,2)) arrrobas from (" & cadena & ") aaaaa"
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    Numcajon = 0
    KilosNet = 0
    Arrobas = 0
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    
    If TotalRegistrosConsulta(cadena) = 0 Then Exit Sub
    
    If Not Rs.EOF Then
        If Rs.Fields(0).Value <> 0 Then Numcajon = DBLet(Rs.Fields(0).Value, "N") 'Solo es para saber que hay registros que mostrar
        If Rs.Fields(1).Value <> 0 Then KilosNet = DBLet(Rs.Fields(1).Value, "N") 'Solo es para saber que hay registros que mostrar
        If Rs.Fields(2).Value <> 0 Then Arrobas = DBLet(Rs.Fields(2).Value, "N") 'Solo es para saber que hay registros que mostrar
    
        Text1.Text = Format(Numcajon, "###,###,##0")
        Text2.Text = Format(KilosNet, "###,###,##0")
        Text3.Text = Format(Arrobas, "###,###,##0.00")
    End If
    Rs.Close
    Set Rs = Nothing

    
    DoEvents
    
End Sub


Private Sub mnImpresionEtiquetas_Click()
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim cajas As Currency
Dim Cad As String
Dim nroPalets As Long
Dim crear As Byte
Dim Imprimir As Byte
Dim b As Boolean


    If vParamAplic.HayTraza = False Then Exit Sub
    
    crear = 1
    Imprimir = 1
    Sql = "select count(*) from trzpalets where numnotac = " & Trim(Me.adodc1.Recordset!numalbar)
    If TotalRegistros(Sql) <> 0 Then
        Cad = "La paletización para esta entrada ya está realizada." & vbCrLf
        Cad = Cad & vbCrLf & "            ¿ Desea imprimirla de nuevo ? "
        If MsgBox(Cad, vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
            crear = 0
            Imprimir = 0
        Else
            crear = 0
            Imprimir = 1
        End If
    End If
    
    b = True
    If crear = 1 Then
        nroPalets = InputBox("Nro de Palets:", "Número de Palets", 0)
        b = InsertarPalets(CStr(adodc1.Recordset!numalbar), nroPalets, CStr(adodc1.Recordset!Numcajon), CStr(adodc1.Recordset!KilosNet), adodc1.Recordset!Fecalbar, CStr(adodc1.Recordset!Codsocio), CStr(adodc1.Recordset!codvarie))
    End If
    
    If Imprimir = 1 Then
        If b Then ImprimirEtiquetas
    End If
    
End Sub




Private Function InsertarPalets(Albaran As String, Palets As Long, NumCajones As Long, Numkilos As Long, Fecha As Date, Socio As String, Variedad As String)
Dim nroPalets As Long
Dim Kilos As Long
Dim cajas As Long
Dim i As Long
Dim CRFID As String
Dim NroCRFID As String
Dim NumNota As Long
Dim NumF As Long
Dim Sql As String
Dim Hora As String
Dim KilosporPalet As Long
Dim RestoCajas As Long
Dim RestoKilos As Long
Dim TotKilos As Long


    On Error GoTo eInsertarPalets

    InsertarPalets = False
    
    If Palets = 0 Then
        nroPalets = Val(NumCajones) \ vParamAplic.CajasporPalet
        RestoCajas = Val(NumCajones) Mod vParamAplic.CajasporPalet
        
        KilosporPalet = (vParamAplic.CajasporPalet * Numkilos) \ Val(NumCajones)
        TotKilos = 0
    
        CRFID = Format(Fecha, "yyyymmdd") & Format(Albaran, "0000000")
        Hora = Mid(Format(Now, "dd/mm/yyyy hh:mm:ss"), 12, 8)
        
        For i = 1 To nroPalets
            NroCRFID = Format(Fecha, "yyyymmdd") & Format(Albaran, "0000000") & Format(i, "000")
            
            NumF = SugerirCodigoSiguienteStr("trzpalets", "idpalet")
            
            TotKilos = TotKilos + KilosporPalet
            
            Sql = "insert into trzpalets (idpalet,tipo,numcajones,numkilos,"
            Sql = Sql & "codsocio,codcampo,codvarie,fecha,hora,numnotac,CRFID) values ("
            Sql = Sql & DBSet(NumF, "N") & "," & DBSet(0, "N") & "," & DBSet(vParamAplic.CajasporPalet, "N") & ","
            Sql = Sql & DBSet(KilosporPalet, "N") & "," & DBSet(Socio, "N") & "," & DBSet(0, "N") & ","
            Sql = Sql & DBSet(Variedad, "N") & "," & DBSet(Fecha, "F") & "," & DBSet(Fecha & " " & Hora, "FH", "S") & ","
            Sql = Sql & DBSet(Albaran, "N") & "," & DBSet(NroCRFID, "T") & ")"
            
            conn.Execute Sql
        Next i
        
        If RestoCajas <> 0 Then ' insertamos el ultimo palet con el resto
            NroCRFID = Format(Fecha, "yyyymmdd") & Format(Albaran, "0000000") & Format(i, "000")
            
            NumF = SugerirCodigoSiguienteStr("trzpalets", "idpalet")
            
            RestoKilos = Numkilos - (KilosporPalet * nroPalets)
            
            TotKilos = TotKilos + RestoKilos
            
            Sql = "insert into trzpalets (idpalet,tipo,numcajones,numkilos,"
            Sql = Sql & "codsocio,codcampo,codvarie,fecha,hora,numnotac,CRFID) values ("
            Sql = Sql & DBSet(NumF, "N") & "," & DBSet(0, "N") & "," & DBSet(RestoCajas, "N") & ","
            Sql = Sql & DBSet(RestoKilos, "N") & "," & DBSet(Socio, "N") & "," & DBSet(0, "N") & ","
            Sql = Sql & DBSet(Variedad, "N") & "," & DBSet(Fecha, "F") & "," & DBSet(Fecha & " " & Hora, "FH", "S") & ","
            Sql = Sql & DBSet(Albaran, "N") & "," & DBSet(NroCRFID, "T") & ")"
            
            conn.Execute Sql
            
            nroPalets = nroPalets + 1
        End If
        
        RestoKilos = Numkilos - TotKilos
        
        If RestoKilos <> 0 Then ' actualizamos el ultimo registro si hay resto de kilos
            Sql = "update trzpalets set numkilos = numkilos + " & DBSet(RestoKilos, "N")
            Sql = Sql & " where idpalet = " & DBSet(NumF, "N")
            
            conn.Execute Sql
        End If
    
    End If
    
    If Palets > 0 Then
        nroPalets = Palets
        Kilos = Numkilos \ nroPalets
        cajas = Val(NumCajones) \ nroPalets
        
        CRFID = Format(Fecha, "yyyymmdd") & Format(Albaran, "0000000")
        Hora = Mid(Format(Now, "dd/mm/yyyy hh:mm:ss"), 12, 8)
        
        For i = 1 To nroPalets
            NroCRFID = Format(Fecha, "yyyymmdd") & Format(Albaran, "0000000") & Format(i, "000")
            
            NumF = SugerirCodigoSiguienteStr("trzpalets", "idpalet")
            
            ' el tipo de trzpalets va a ser siempre 0, pq se piden palets
            
            Sql = "insert into trzpalets (idpalet,tipo,numcajones,numkilos,"
            Sql = Sql & "codsocio,codcampo,codvarie,fecha,hora,numnotac,CRFID) values ("
            Sql = Sql & DBSet(NumF, "N") & "," & DBSet(0, "N") & "," & DBSet(cajas, "N") & ","
            Sql = Sql & DBSet(Kilos, "N") & "," & DBSet(Socio, "N") & "," & DBSet(0, "N") & ","
            Sql = Sql & DBSet(Variedad, "N") & "," & DBSet(Fecha, "F") & "," & DBSet(Fecha & " " & Hora, "FH", "S") & ","
            Sql = Sql & DBSet(Albaran, "N") & "," & DBSet(NroCRFID, "T") & ")"
            
            conn.Execute Sql
        Next i
        
        Sql = "update trzpalets set numcajones = numcajones + " & (CCur(NumCajones) - (cajas * nroPalets))
        Sql = Sql & ", numkilos = numkilos + " & CCur(Numkilos) - (Kilos * nroPalets)
        Sql = Sql & " where numnotac = " & DBSet(Albaran, "N")
        Sql = Sql & " and idpalet = " & DBSet(NumF, "N")
        
        conn.Execute Sql
    End If
    
    
    InsertarPalets = True
    Exit Function

eInsertarPalets:
    MuestraError Err.Number, "Insertar Palets", Err.Description
End Function

Private Sub ImprimirEtiquetas()

    If adodc1.Recordset.EOF Then Exit Sub
    
    Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
    Dim nomDocu As String 'Nombre de Informe rpt de crystal
    Dim ImprimeDirecto As Integer
     
    indRPT = 93 'Ticket de Entrada
     
    If Not PonerParamRPT(indRPT, "", 1, nomDocu, , ImprimeDirecto) Then Exit Sub
    'Nombre fichero .rpt a Imprimir
    ' he añadido estas dos lineas para que llame al rpt correspondiente
    
    frmImprimir.NombreRPT = nomDocu
    
    ActivaTicket
    
    With frmVisReport
        .FormulaSeleccion = "{trzpalets.numnotac}=" & adodc1.Recordset!numalbar
        .SoloImprimir = True
        .OtrosParametros = ""
        .NumeroParametros = 1
        .MostrarTree = False
        .Informe = App.Path & "\informes\" & nomDocu    ' "ValEntrada.rpt"
        .InfConta = False
        .ConSubInforme = False
        .SubInformeConta = ""
        .Opcion = 0
        .ExportarPDF = False
        .Show vbModal
    End With
    
    DesactivaTicket

End Sub


'***************************************
Private Sub ActivaTicket()
    ImpresoraDefecto = Printer.DeviceName
    XPDefaultPrinter vParamAplic.ImpresoraEntradas
End Sub

Private Sub DesactivaTicket()
    XPDefaultPrinter ImpresoraDefecto
End Sub


'---------------- Procesos para cambio de impresora por defecto ------------------
Private Sub XPDefaultPrinter(PrinterName As String)
    Dim Buffer As String
    Dim DeviceName As String
    Dim DriverName As String
    Dim PrinterPort As String
    Dim r As Long
    ' Get the printer information for the currently selected
    ' printer in the list. The information is taken from the
    ' WIN.INI file.
    Buffer = Space(1024)
    r = GetProfileString("PrinterPorts", PrinterName, "", _
        Buffer, Len(Buffer))

    ' Parse the driver name and port name out of the buffer
    GetDriverAndPort Buffer, DriverName, PrinterPort

       If DriverName <> "" And PrinterPort <> "" Then
           SetDefaultPrinter PrinterName, DriverName, PrinterPort
       End If
End Sub

Private Sub GetDriverAndPort(ByVal Buffer As String, DriverName As _
    String, PrinterPort As String)

    Dim iDriver As Integer
    Dim iPort As Integer
    DriverName = ""
    PrinterPort = ""

    ' The driver name is first in the string terminated by a comma
    iDriver = InStr(Buffer, ",")
    If iDriver > 0 Then

         ' Strip out the driver name
        DriverName = Left(Buffer, iDriver - 1)

        ' The port name is the second entry after the driver name
        ' separated by commas.
        iPort = InStr(iDriver + 1, Buffer, ",")

        If iPort > 0 Then
            ' Strip out the port name
            PrinterPort = Mid(Buffer, iDriver + 1, _
            iPort - iDriver - 1)
        End If
    End If
End Sub

Private Sub SetDefaultPrinter(ByVal PrinterName As String, _
    ByVal DriverName As String, ByVal PrinterPort As String)
    Dim DeviceLine As String
    Dim r As Long
    Dim L As Long
    DeviceLine = PrinterName & "," & DriverName & "," & PrinterPort
    ' Store the new printer information in the [WINDOWS] section of
    ' the WIN.INI file for the DEVICE= item
    r = WriteProfileString("windows", "Device", DeviceLine)
    ' Cause all applications to reload the INI file:
    L = SendMessage(HWND_BROADCAST, WM_WININICHANGE, 0, "windows")
End Sub
'------------------ Fin de los procesos relacionados con el cambio de impresora ----


Private Function ActualizarTraza2(Nota As String, Variedad As String, Socio As String, Fecha As String, MenError As String)
Dim Rs As ADODB.Recordset
Dim Sql As String
Dim Sql1 As String
Dim Sql2 As String
Dim IdPalet As Currency

    On Error GoTo eActualizarTraza2

    ActualizarTraza2 = True

    If vParamAplic.HayTraza = False Then Exit Function
    
    Sql = "select idpalet from trzpalets where numnotac = " & DBSet(Nota, "N")
    
    
    'Comprobamos si la fecha de abocamiento de alguno de sus palets es inferior a la de la entrada
    'para no permitir modificar la traza
    Sql2 = "select sum(resul) from (select if(date(fechahora)<" & DBSet(Fecha, "F") & ",1,0) as resul "
    Sql2 = Sql2 & " from trzlineas_cargas where idpalet in (" & Sql & ")) aaaaa "
    If CLng(DevuelveValor(Sql2)) > 0 Then
        MenError = MenError & vbCrLf & "No se permite una fecha de entrada superior a la de abocamiento de ninguno de sus palets. Revise."
        ActualizarTraza2 = False
        Exit Function
    End If
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenDynamic, adLockReadOnly, adCmdText
    
    While Not Rs.EOF
        
        Sql1 = "update trzpalets set codvarie = " & DBSet(Variedad, "N")
        Sql1 = Sql1 & ", codsocio = " & DBSet(Socio, "N")
        Sql1 = Sql1 & ", fecha = " & DBSet(Fecha, "F")
        '[Monica]13/12/2013: no funcionaba el date(hora)
        Sql1 = Sql1 & ", hora = concat('" & Format(Fecha, "yyyy-mm-dd") & "', ' ', time(hora))"
        Sql1 = Sql1 & " where idpalet = " & DBSet(Rs.Fields(0).Value, "N")
        
        conn.Execute Sql1
        
        Rs.MoveNext
    Wend
        
    Set Rs = Nothing
    
    Exit Function
    
eActualizarTraza2:
    ActualizarTraza2 = False
    MenError = MenError & vbCrLf & Err.Description
End Function


Private Function ActualizarPaletizacion(MenError As String) As Boolean
Dim Rs As ADODB.Recordset
Dim Sql As String
Dim Sql1 As String
Dim KilosTotal As Currency
Dim KilosNeto As Currency
Dim KilosLinea As Currency
Dim Numlineas As Currency
Dim IdPalet As Currency
Dim NumCajas As Long


    On Error GoTo eActualizarPaletizacion

    ActualizarPaletizacion = True

    If vParamAplic.HayTraza = False Then Exit Function
    
    Sql = "select numcajones, numkilos, idpalet from trzpalets where numnotac = " & Trim(adodc1.Recordset!numalbar)
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenDynamic, adLockReadOnly, adCmdText
    
    If Not Rs.EOF Then
        Rs.MoveFirst
        
        KilosNeto = ComprobarCero(txtAux(5).Text) 'DBLet(adodc1.Recordset!KilosNet, "N")

        NumCajas = ComprobarCero(txtAux(4).Text)
        
        KilosTotal = 0
        While Not Rs.EOF
            If NumCajas <> 0 Then ' estamos por palet
                KilosLinea = (KilosNeto * DBLet(Rs.Fields(0).Value, "N")) \ NumCajas
            Else ' estamos por palot
                KilosLinea = KilosNeto \ Numlineas
            End If
            
            Sql1 = "update trzpalets set numkilos = " & DBSet(KilosLinea, "N")
            Sql1 = Sql1 & " where idpalet = " & DBSet(Rs.Fields(2).Value, "N")
            
            conn.Execute Sql1
            
            KilosTotal = KilosTotal + KilosLinea
        
            IdPalet = DBLet(Rs.Fields(2).Value, "N")
            
            Rs.MoveNext
        Wend
        
        If KilosTotal <> KilosNeto Then ' en el ultimo registro metemos el restante
            Sql1 = "update trzpalets set numkilos = numkilos + " & DBSet(KilosNeto - KilosTotal, "N")
            Sql1 = Sql1 & " where idpalet = " & DBSet(IdPalet, "N")
            
            conn.Execute Sql1
        End If
    End If
    
    Set Rs = Nothing
    Exit Function
        
eActualizarPaletizacion:
    ActualizarPaletizacion = False
    MenError = MenError & vbCrLf & Err.Description
End Function


