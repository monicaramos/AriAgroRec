VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmAlmEnvRet 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Envases Retornables"
   ClientHeight    =   9915
   ClientLeft      =   45
   ClientTop       =   30
   ClientWidth     =   16305
   Icon            =   "frmAlmEnvRet.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9915
   ScaleWidth      =   16305
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameBotonGnral 
      Height          =   705
      Left            =   135
      TabIndex        =   20
      Top             =   135
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
      Left            =   2430
      MaxLength       =   12
      TabIndex        =   3
      Tag             =   "Cliente|N|N|||albaran_envase|codclien|000000|N|"
      Top             =   4950
      Width           =   495
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
      Index           =   2
      Left            =   2970
      MaskColor       =   &H00000000&
      TabIndex        =   19
      ToolTipText     =   "Buscar Cliente"
      Top             =   4950
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.TextBox Text2 
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
      Height          =   330
      Index           =   3
      Left            =   3195
      Locked          =   -1  'True
      MaxLength       =   40
      TabIndex        =   18
      Text            =   "Nombre cliente"
      Top             =   4950
      Width           =   2145
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
      Left            =   450
      MaxLength       =   6
      TabIndex        =   1
      Tag             =   "Num.Linea|N|N|||albaran_envase|numlinea|000000|S|"
      Top             =   4950
      Width           =   855
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
      Left            =   1350
      MaxLength       =   10
      TabIndex        =   2
      Tag             =   "Fecha Movimiento|F|N|||albaran_envase|fechamov|dd/mm/yyyy||"
      Top             =   4950
      Width           =   800
   End
   Begin VB.ComboBox cmbAux 
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
      Left            =   10260
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Tag             =   "Tipo Movimiento|N|N|||albaran_envase|tipomovi|0||"
      Top             =   4950
      Width           =   1215
   End
   Begin VB.TextBox Text2 
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
      Height          =   330
      Index           =   0
      Left            =   6300
      Locked          =   -1  'True
      MaxLength       =   40
      TabIndex        =   17
      Text            =   "Nombre articulo"
      Top             =   4950
      Width           =   1740
   End
   Begin VB.TextBox Text2 
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
      Height          =   330
      Index           =   2
      Left            =   8640
      Locked          =   -1  'True
      MaxLength       =   40
      TabIndex        =   15
      Text            =   "Nomtipart"
      Top             =   4950
      Width           =   1560
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
      Left            =   6075
      MaskColor       =   &H00000000&
      TabIndex        =   14
      ToolTipText     =   "Buscar Art�culo"
      Top             =   4950
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
      Index           =   3
      Left            =   11565
      MaxLength       =   7
      TabIndex        =   6
      Tag             =   "Cantidad|N|N|||albaran_envase|cantidad|###,##0||"
      Top             =   4950
      Width           =   1215
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
      Left            =   2160
      MaskColor       =   &H00000000&
      TabIndex        =   13
      ToolTipText     =   "Buscar Fecha"
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
      Left            =   13875
      TabIndex        =   7
      Top             =   9375
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
      Left            =   15030
      TabIndex        =   8
      Top             =   9375
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
      Left            =   5535
      MaxLength       =   12
      TabIndex        =   4
      Tag             =   "Art�culo|T|N|||albaran_envase|codartic||N|"
      Top             =   4950
      Width           =   495
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmAlmEnvRet.frx":000C
      Height          =   8145
      Left            =   180
      TabIndex        =   11
      Top             =   945
      Width           =   15915
      _ExtentX        =   28072
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
      Left            =   15030
      TabIndex        =   12
      Top             =   9360
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Index           =   1
      Left            =   180
      TabIndex        =   9
      Top             =   9180
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
         TabIndex        =   10
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
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   290
      Index           =   0
      Left            =   585
      MaxLength       =   6
      TabIndex        =   0
      Tag             =   "Num.Albaran|N|N|||albaran_envase|numalbar|000000|S|"
      Top             =   4950
      Width           =   660
   End
   Begin VB.TextBox Text2 
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
      Height          =   330
      Index           =   1
      Left            =   8055
      Locked          =   -1  'True
      MaxLength       =   40
      TabIndex        =   16
      Text            =   "TipArt"
      Top             =   4950
      Width           =   570
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
Attribute VB_Name = "frmAlmEnvRet"
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
' 3. En la funci� BotonAnyadir() canviar la taula i el camp per a SugerirCodigoSiguienteStr
' 4. En la funci� BotonBuscar() canviar el nom de la clau primaria
' 5. En la funci� BotonEliminar() canviar la pregunta, les descripcions de la
'    variable SQL i el contingut del DELETE
' 6. En la funci� PonerLongCampos() posar els camps als que volem canviar el MaxLength quan busquem
' 7. En Form_Load() repasar la barra d'iconos (per si es vol canviar alg�n) i
'    canviar la consulta per a vore tots els registres
' 8. En Toolbar1_ButtonClick repasar els indexs de cada bot� per a que corresponguen
' 9. En la funci� CargaGrid canviar l'ORDER BY (normalment per la clau primaria);
'    canviar adem�s els noms dels camps, el format i si fa falta la cantitat;
'    repasar els index dels botons modificar i eliminar.
'    NOTA: si en Form_load ya li he posat clausula WHERE, canviar el `WHERE` de
'    `SQL = CadenaConsulta & " WHERE " & vSQL` per un `AND`
' 10. En txtAux_LostFocus canviar el mensage i el format del camp
' 11. En la funci� DatosOk() canviar els arguments de DevuelveDesdeBD i el mensage
'    en cas d'error
' 12. En la funci� SepuedeBorrar() canviar les comprovacions per a vore si es pot
'    borrar el registre
' *******************************SI N'HI HA COMBO*******************************
' 0. Comprovar que en el SQL de Form_Load() es fa�a refer�ncia a la taula del Combo
' 1. Pegar el Combo1 al  costat dels TextAux. Canviar-li el TAG
' 2. En BotonModificar() canviar el camp del Combo
' 3. En CargaCombo() canviar la consulta i els noms del camps, o posar els valor
'    a ma si no es llig de cap base de datos els valors del Combo

Option Explicit

Public DatosADevolverBusqueda As String    'Tendra el n� de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)

Public ParamVariedad As String

'codi per al registe que s'afegix al cridar des d'atre formulari.
'Obrir en modo Insertar i tornar datos del registre insertat
Public NuevoCodigo As String
Public CodigoActual As String
Public DeConsulta As Boolean

Private CadenaConsulta As String
Private CadB As String

Private WithEvents frmVar As frmManVariedad  'Variedades
Attribute frmVar.VB_VarHelpID = -1
Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1

Dim Modo As Byte
'----------- MODOS --------------------------------
'   0.-  Formulari net sense cap camp ple
'   1.-  Preparant per a fer la b�squeda
'   2.-  Ja tenim registres i els anem a recorrer
'        ,podem editar-los Edici� del camp
'   3.-  Inserci� de nou registre
'   4.-  Modificar
'--------------------------------------------------
Dim PrimeraVez As Boolean
Dim indice As Byte 'Index del text1 on es poses els datos retornats des d'atres Formularis de Mtos
Dim i As Integer

Private Sub PonerModo(vModo)
Dim b As Boolean

    Modo = vModo
    
    b = (Modo = 2)
    If b Then
        PonerContRegIndicador
    Else
        PonerIndicador lblIndicador, Modo
    End If
    
    For i = 0 To txtAux.Count - 1
        txtAux(i).visible = Not b
    Next i
    For i = 0 To Text2.Count - 1
        Text2(i).visible = Not b
    Next i
    btnBuscar(0).visible = Not b
    btnBuscar(1).visible = Not b
    btnBuscar(2).visible = Not b
    BloquearBtn btnBuscar(0), Not b
    BloquearBtn btnBuscar(1), Not b
    BloquearBtn btnBuscar(2), Not b
    
    cmbAux(0).visible = Not b
    cmbAux(0).Enabled = Not b
'    For I = 0 To cmbAux.Count - 1
'        cmbAux(I).visible = False
'        cmbAux(I).Enabled = True
'    Next I
  
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
    txtAux(1).Enabled = (Modo = 1)
    
    BloquearBtn btnBuscar(0), Not (Modo <> 2)
    BloquearBtn btnBuscar(1), Not (Modo <> 2)
    BloquearBtn btnBuscar(2), Not (Modo <> 2)
    
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
    
    'Imprimir
    Toolbar1.Buttons(8).Enabled = True
    Me.mnImprimir.Enabled = True
    
End Sub

Private Sub BotonAnyadir()
    Dim NumF As String
    Dim anc As Single
    
'    CargaGrid 'primer de tot carregue tot el grid
'    CadB = ""
    '******************** canviar taula i camp **************************
    If (DatosADevolverBusqueda <> "") And NuevoCodigo <> "" Then
        NumF = NuevoCodigo
    Else
        NumF = SugerirCodigoSiguienteStr("albaran_envase", "numlinea", "numalbar = 0")
    End If
    '********************************************************************
    'Situamos el grid al final
    AnyadirLinea DataGrid1, adodc1
         
    anc = DataGrid1.Top
    If DataGrid1.Row < 0 Then
        anc = anc + 240
    Else
        anc = anc + DataGrid1.RowTop(DataGrid1.Row) + 5
    End If
    
    txtAux(0).Text = 0 ' numero de albaran: 0
    txtAux(1).Text = NumF
    
    FormateaCampo txtAux(1)
    For i = 2 To txtAux.Count - 1
        txtAux(i).Text = ""
    Next i
    For i = 0 To 3
        Text2(i).Text = ""
    Next i
    Me.cmbAux(0).ListIndex = -1
    LLamaLineas anc, 3 'Pone el form en Modo=3, Insertar
       
    'Ponemos el foco
    PonerFoco txtAux(4)
End Sub

Private Sub BotonVerTodos()
    CadB = ""
    CargaGrid ""
    PonerModo 2
End Sub

Private Sub BotonBuscar()
    ' ***************** canviar per la clau primaria ********
    CargaGrid "albaran_envase.numalbar = -1"
    '*******************************************************************************
    'Buscar
    For i = 0 To txtAux.Count - 1
        txtAux(i).Text = ""
    Next i
    For i = 0 To Text2.Count - 1
        Text2(i).Text = ""
    Next i
    Me.cmbAux(0).ListIndex = -1
    LLamaLineas DataGrid1.Top + 240, 1 'Pone el form en Modo=1, Buscar
    PonerFoco txtAux(1)
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
        anc = DataGrid1.Top + DataGrid1.RowTop(DataGrid1.Row) '+ 580 '545
    End If

    'Llamamos al form
    txtAux(0).Text = DataGrid1.Columns(0).Text
    txtAux(1).Text = DataGrid1.Columns(1).Text
    txtAux(4).Text = DataGrid1.Columns(2).Text
    txtAux(2).Text = DataGrid1.Columns(5).Text
    
    Text2(0).Text = DataGrid1.Columns(6).Text
    Text2(2).Text = DataGrid1.Columns(8).Text
    
    cmbAux(0).Text = DataGrid1.Columns(9).Text
    txtAux(3).Text = DataGrid1.Columns(10).Text
    
    txtAux(5).Text = DataGrid1.Columns(3).Text
    Text2(3).Text = DataGrid1.Columns(4).Text
       

    ' ***** canviar-ho pel nom del camp del combo *********
'    SelComboBool DataGrid1.Columns(2).Text, Combo1(0)
    ' *****************************************************

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
    
    ' ### [Monica] 12/09/2006
    For i = 0 To Text2.Count - 1
        Text2(i).Top = alto
    Next i
    btnBuscar(0).Top = alto
    btnBuscar(1).Top = alto
    btnBuscar(2).Top = alto
    Me.cmbAux(0).Top = alto
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
    Sql = "�Seguro que desea eliminar el Envase?"
    Sql = Sql & vbCrLf & "L�nea: " & adodc1.Recordset.Fields(1)
    Sql = Sql & vbCrLf & "Fecha " & adodc1.Recordset.Fields(2)
    Sql = Sql & vbCrLf & "Cliente: " & Format(adodc1.Recordset.Fields(3), "000000") & " - " & adodc1.Recordset.Fields(4)
    Sql = Sql & vbCrLf & "Articulo: " & adodc1.Recordset.Fields(5) & " - " & adodc1.Recordset.Fields(6)
    
    If MsgBox(Sql, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        NumRegElim = adodc1.Recordset.AbsolutePosition
        Sql = "Delete from albaran_envase where numalbar=" & adodc1.Recordset!NumAlbar
        Sql = Sql & " and numlinea = " & adodc1.Recordset!NumLinea
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
'Modificar el MaxLength del campo en funcion de si es modo de b�squeda o no
'para los campos que permitan introducir criterios m�s largos del tama�o del campo
'Si maxlength=6 para codprove, en modo busqueda Maxlenth=13 'el doble + 1
    
    'para los txtAux
    PonerLongCamposGnral Me, Modo, 3
End Sub
Private Sub frmC_Selec(vFecha As Date)
    ' *** repasar si el camp es txtAux o Text1 ***
    txtAux(4).Text = Format(vFecha, "dd/mm/yyyy") '<===
    ' ********************************************
End Sub

Private Sub btnBuscar_Click(Index As Integer)
 TerminaBloquear
    
    Select Case Index
        Case 0 'fecha
            indice = Index + 4
            
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
            
            menu = Me.Height - Me.ScaleHeight 'ac� tinc el heigth del men� i de la toolbar
        
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
            
        Case 1 'articulos
'            Set frmArt = New frmManArtic
'            frmArt.DatosADevolverBusqueda = "0|1|"
'            frmArt.CodigoActual = txtAux(2).Text
'            frmArt.Show vbModal
'            Set frmArt = Nothing
'            PonerFoco txtAux(2)
        
        Case 2 'clientes
'            Set frmCli = New frmClientes
'            frmCli.DatosADevolverBusqueda = "0|2|"
''            frmCli.CodigoActual = txtAux(5).Text
'            frmCli.Show vbModal
'            Set frmCli = Nothing
'            PonerFoco txtAux(5)
       
        
    End Select
    
    If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Me.adodc1, 1
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
            If DatosOk Then
                If InsertarDesdeForm(Me) Then
                    CargaGrid CadB
                    If (DatosADevolverBusqueda <> "") And NuevoCodigo <> "" Then
                        cmdCancelar_Click
'                        If Not adodc1.Recordset.EOF Then adodc1.Recordset.MoveLast
                        If Not adodc1.Recordset.EOF Then
                            SituarDataMULTI adodc1, "numalbar = " & txtAux(0) & " and numlinea = " & txtAux(1), ""
                        End If
                        cmdRegresar_Click
                    Else
                        BotonAnyadir
                    End If
                    CadB = ""
                End If
            End If
            
        Case 4 'MODIFICAR
            If DatosOk Then
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
        Case 1 'b�squeda
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
        MsgBox "Ning�n registro devuelto.", vbExclamation
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

Private Sub DataGrid1_DblClick()
    If cmdRegresar.visible Then cmdRegresar_Click
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
    
    'Icono del formulario
    Me.Icon = frmPpal.Icon
    
    PrimeraVez = True

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

    '## A mano
'    chkVistaPrevia.Value = CheckValueLeer(Name)
    
    
    '****************** canviar la consulta *********************************+
    CadenaConsulta = "SELECT albaran_envase.numalbar, numlinea, albaran_envase.fechamov, albaran_envase.codclien, clientes.nomclien,  albaran_envase.codartic, "
    CadenaConsulta = CadenaConsulta & " sartic.nomartic, sartic.codtipar, stipar.nomtipar, "
    CadenaConsulta = CadenaConsulta & " CASE albaran_envase.tipomovi WHEN 0 THEN ""Salida"" WHEN 1 THEN ""Entrada"" END, albaran_envase.cantidad "
    CadenaConsulta = CadenaConsulta & " FROM albaran_envase, sartic, stipar, clientes "
    CadenaConsulta = CadenaConsulta & " WHERE albaran_envase.codartic = sartic.codartic"
    CadenaConsulta = CadenaConsulta & " and albaran_envase.numalbar = 0"
    CadenaConsulta = CadenaConsulta & " and sartic.codtipar = stipar.codtipar"
    CadenaConsulta = CadenaConsulta & " and albaran_envase.codclien = clientes.codclien "
    '************************************************************************
    
    CadB = ""
    CargaGrid
    
    CargaCombo
    
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

Private Sub frmAlb_DatoSeleccionado(CadenaSeleccion As String)
'Alabaranes
    txtAux(0).Text = RecuperaValor(CadenaSeleccion, 1) 'numalbar

End Sub

Private Sub frmArt_DatoSeleccionado(CadenaSeleccion As String)
'Articulos
    txtAux(2).Text = RecuperaValor(CadenaSeleccion, 1) 'codartic
    Text2(0).Text = RecuperaValor(CadenaSeleccion, 2) 'descripcion
    If txtAux(2) <> "" Then
        Text2(1) = DevuelveDesdeBDNew(cAgro, "sartic", "codtipar", "codartic", txtAux(2), "N")
        Text2(2) = DevuelveDesdeBDNew(cAgro, "stipar", "nomtipar", "codtipar", Text2(1), "N")
    End If
End Sub

Private Sub frmCli_DatoSeleccionado(CadenaSeleccion As String)
'Clientes
    txtAux(5).Text = RecuperaValor(CadenaSeleccion, 1) 'codclien
    Text2(3).Text = RecuperaValor(CadenaSeleccion, 2) 'nomclien
End Sub

Private Sub mnBuscar_Click()
    BotonBuscar
End Sub

Private Sub mnEliminar_Click()
    BotonEliminar
End Sub

Private Sub mnImprimir_Click()
    AbrirListadoComer (14)

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
    
    
    '********************* canviar el ORDER BY *********************++
    Sql = Sql & " ORDER BY albaran_envase.numalbar, albaran_envase.numlinea"
    '**************************************************************++
    
    CargaGridGnral Me.DataGrid1, Me.adodc1, Sql, PrimeraVez
    
    ' *******************canviar els noms i si fa falta la cantitat********************
    
    tots = "N|txtAux(0)|T|Albar�n|800|;S|txtAux(1)|T|L�nea|850|;"
    tots = tots & "S|txtAux(4)|T|Fecha|1350|;S|btnBuscar(0)|B|||;S|txtAux(5)|T|Cliente|1100|;S|btnBuscar(2)|B|||;S|Text2(3)|T|Denominaci�n|3300|;"
    tots = tots & "S|txtAux(2)|T|Art�culo|1750|;S|btnBuscar(1)|B|||;S|Text2(0)|T|Denominaci�n|3000|;"
    tots = tots & "N|Text2(1)|T|Tipo|1000|;S|Text2(2)|T|Tipo Envase|1600|;S|cmbAux(0)|C|Tipo Mov.|1200|;"
    tots = tots & "S|txtAux(3)|T|Cantidad|1200|;"
    
    arregla tots, DataGrid1, Me, 350
    
    DataGrid1.ScrollBars = dbgAutomatic
    DataGrid1.Columns(0).Alignment = dbgLeft
    DataGrid1.Columns(1).Alignment = dbgLeft
    
    DataGrid1.Columns(3).Alignment = dbgLeft
    
    DataGrid1.Columns(7).Alignment = dbgLeft
    
'   DataGrid1.Columns(2).Alignment = dbgRight
End Sub

Private Sub txtAux_GotFocus(Index As Integer)
    ConseguirFocoLin txtAux(Index)
End Sub


Private Sub txtAux_LostFocus(Index As Integer)
Dim cadMen As String

    If Not PerderFocoGnral(txtAux(Index), Modo) Then Exit Sub
    
    Select Case Index
        Case 0 'numero de albaran
            If txtAux(Index).Text = "" Then Exit Sub
'            Text2(2).Text = PonerNombreDeCod(txtAux(Index), "variedades", "nomvarie", "codvarie", "N")
        
        Case 1 'numero de linea
            PonerFormatoEntero txtAux(Index)
            
        Case 5 'codigo de cliente
            If txtAux(Index) <> "" Then
                Text2(3).Text = PonerNombreDeCod(txtAux(Index), "clientes", "nomclien")
                If Text2(3).Text = "" Then
'                    cadMen = "No existe el Cliente: " & txtAux(Index).Text & vbCrLf
'                    cadMen = cadMen & "�Desea crearlo?" & vbCrLf
'                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
'                        Set frmCli = New frmClientes
'                        frmCli.DatosADevolverBusqueda = "0|1|"
''                        frmCli.NuevoCodigo = txtAux(Index).Text
'                        txtAux(Index).Text = ""
'                        frmCli.Show vbModal
'                        Set frmCli = Nothing
'                    Else
'                        txtAux(Index).Text = ""
'                        Text2(3).Text = ""
'                    End If
                    PonerFoco txtAux(Index)
                End If
            Else
                Text2(3).Text = ""
            End If
        
        Case 2 'codigo de articulo
            If txtAux(Index) <> "" Then
                Text2(0).Text = PonerNombreDeCod(txtAux(Index), "sartic", "nomartic")
                Text2(1) = DevuelveDesdeBDNew(cAgro, "sartic", "codtipar", "codartic", txtAux(2), "T")
                Text2(2) = ""
                If Text2(1) <> "" Then
                    If EsArticuloRetornable(Text2(1)) Then  ' CByte(Text2(1)) = 1 Or CByte(Text2(1)) = 2 Or CByte(Text2(1)) = 3 Then
                        Text2(2) = DevuelveDesdeBDNew(cAgro, "stipar", "nomtipar", "codtipar", Text2(1), "T")
                    Else
                        MsgBox "El Tipo de Envase ha de ser retornable. Reintroduzca.", vbExclamation
                        txtAux(Index).Text = ""
                        Text2(0).Text = ""
                        Text2(1).Text = ""
                        Text2(2).Text = ""
                        PonerFoco txtAux(Index)
                        Exit Sub
                    End If
                End If
                If Text2(0).Text = "" Then
'                    cadMen = "No existe el Envase: " & txtAux(Index).Text & vbCrLf
'                    cadMen = cadMen & "�Desea crearlo?" & vbCrLf
'                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
'                        Set frmArt = New frmManArtic
'                        frmArt.DatosADevolverBusqueda = "0|1|"
'                        frmArt.NuevoCodigo = txtAux(Index).Text
'                        txtAux(Index).Text = ""
'                        frmArt.Show vbModal
'                        Set frmArt = Nothing
'                    Else
'                        txtAux(Index).Text = ""
'                        Text2(0).Text = ""
'                        Text2(1).Text = ""
'                        Text2(2).Text = ""
'                    End If
                    PonerFoco txtAux(Index)
                End If
            Else
                Text2(0).Text = ""
                Text2(1).Text = ""
                Text2(2).Text = ""
            End If
            
        Case 3 ' cantidad
            If txtAux(Index) <> "" Then PonerFormatoEntero txtAux(Index)
 
        Case 4 ' fecha de movimiento
            PonerFormatoFecha txtAux(Index)
            
    End Select
    
End Sub

Private Function DatosOk() As Boolean
'Dim Datos As String
Dim b As Boolean
Dim Sql As String
Dim Mens As String

    b = CompForm(Me)
    If Not b Then Exit Function
    
    If Modo = 3 Then   'Estamos insertando
'        SQL = DevuelveDesdeBDNew(cAgro, "albaran", "numalbar", "numalbar", txtAux(0).Text, "N")
'        If SQL = "" Then
'            b = InsertarCabeceraAlbaran(txtAux(0).Text)
'        End If
        
        If b Then
            Sql = DevuelveDesdeBDNew(cAgro, "albaran_envase", "numalbar", "numalbar", txtAux(0).Text, "N", , "numlinea", txtAux(1).Text, "N")
            If Sql <> "" Then
                MsgBox "Esta l�nea existe para este albar�n. Reintroduzca.", vbExclamation
                PonerFoco txtAux(1)
                b = False
            End If
        End If
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


'' ### [DavidV] 26/04/2006: Activar/desactivar la rueda del rat�n.
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
'Alvan�ar/Retrocedir els camps en les fleches de despla�ament del teclat.
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


Private Sub cmbAux_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub CargaCombo()
Dim Rs As ADODB.Recordset
Dim Sql As String
Dim i As Byte
    
    cmbAux(0).Clear
    
    cmbAux(0).AddItem "Salida"
    cmbAux(0).ItemData(cmbAux(0).NewIndex) = 0
    
    cmbAux(0).AddItem "Entrada"
    cmbAux(0).ItemData(cmbAux(0).NewIndex) = 1
    
End Sub

Private Function InsertarCabeceraAlbaran(NumAlbar As String) As Boolean
Dim Sql As String
'    InsertarCabeceraAlbaran = False
'
'    On Error GoTo eInsertarCabeceraAlbaran
'    'forma de pago
'    SQL = DevuelveDesdeBDNew(cAgro, "forpago", "codforpa", "codforpa", 0, "N")
'    If SQL = "" Then
'        SQL = "insert into forpago(codforpa) values (0)"
'        Conn.Execute SQL
'    End If
'    ' cliente
'    SQL = DevuelveDesdeBDNew(cAgro, "clientes", "codclien", "codclien", 0, "N")
'    If SQL = "" Then
'        SQL = "insert into clientes (codclien) values (0)"
'        Conn.Execute SQL
'    End If
'    'banco propio
'    SQL = DevuelveDesdeBDNew(cAgro, "banpropi", "codbanpr", "codbanpr", 0, "N")
'    If SQL = "" Then
'        SQL = "insert into banpropi(codbanpr,digcontr,cuentaba) values (0,'0','0')"
'        Conn.Execute SQL
'    End If
'    'destinos
'    SQL = DevuelveDesdeBDNew(cAgro, "destinos", "codclien", "codclien", 0, "N", , "coddesti", 0, "N")
'    If SQL = "" Then
'        SQL = "insert into destinos(codclien,coddesti) values (0,0)"
'        Conn.Execute SQL
'    End If
'    'agencias
'    SQL = DevuelveDesdeBDNew(cAgro, "agencias", "codtrans", "codtrans", 0, "N")
'    If SQL = "" Then
'        SQL = "insert into agencias(codtrans) values (0)"
'        Conn.Execute SQL
'    End If
'    'tipo de mercado
'    SQL = DevuelveDesdeBDNew(cAgro, "tipomer", "codtimer", "codtimer", 0, "N")
'    If SQL = "" Then
'        SQL = "insert into tipomer(codtimer) values (0)"
'        Conn.Execute SQL
'    End If
'
'    SQL = "insert into albaran(numalbar,fechaalb,codclien,coddesti,codtrans,codtimer,pasaridoc,codalmac) "
'    SQL = SQL & " values ( " & DBSet(NumAlbar, "N") & "," & DBSet(txtAux(4).Text, "F") & ",0,0,0,0,0,1)"
'    Conn.Execute SQL
'
'    InsertarCabeceraAlbaran = True
'    Exit Function
'
'eInsertarCabeceraAlbaran:
'    If Err.Number <> 0 Then
'        InsertarCabeceraAlbaran = False
'    End If
End Function
