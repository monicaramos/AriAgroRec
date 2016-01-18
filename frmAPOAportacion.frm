VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmAPOAportacion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Aportaciones"
   ClientHeight    =   6075
   ClientLeft      =   195
   ClientTop       =   480
   ClientWidth     =   14415
   Icon            =   "frmAPOAportacion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6075
   ScaleWidth      =   14415
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtAux2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   285
      Index           =   2
      Left            =   12810
      TabIndex        =   22
      Top             =   5040
      Width           =   1215
   End
   Begin VB.TextBox txtAux2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   11640
      TabIndex        =   21
      Top             =   5040
      Width           =   1125
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   4
      Left            =   9120
      MaxLength       =   11
      TabIndex        =   5
      Tag             =   "Kilos|N|N|||raportacion|kilos|###,##0||"
      Top             =   4860
      Width           =   810
   End
   Begin VB.TextBox txtAux2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   285
      Index           =   6
      Left            =   5850
      TabIndex        =   19
      Top             =   4830
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   6
      Left            =   4800
      MaxLength       =   6
      TabIndex        =   2
      Tag             =   "Tipo Aportación|N|N|||raportacion|codaport|000|S|"
      Top             =   4830
      Width           =   885
   End
   Begin VB.CommandButton btnBuscar 
      Appearance      =   0  'Flat
      Caption         =   "+"
      Height          =   300
      Index           =   2
      Left            =   5670
      MaskColor       =   &H00000000&
      TabIndex        =   18
      ToolTipText     =   "Buscar tipo aportación"
      Top             =   4830
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.CommandButton btnBuscar 
      Appearance      =   0  'Flat
      Caption         =   "+"
      Height          =   300
      Index           =   1
      Left            =   1080
      MaskColor       =   &H00000000&
      TabIndex        =   17
      ToolTipText     =   "Buscar socio"
      Top             =   4800
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   5
      Left            =   9990
      MaxLength       =   8
      TabIndex        =   6
      Tag             =   "Importe|N|N|||raportacion|importe|###,##0.00||"
      Top             =   4860
      Width           =   1170
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   3
      Left            =   8280
      MaxLength       =   9
      TabIndex        =   4
      Tag             =   "Campaña|T|N|||raportacion|campanya|||"
      Top             =   4860
      Width           =   750
   End
   Begin VB.TextBox txtAux2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   1170
      TabIndex        =   16
      Top             =   4800
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton btnBuscar 
      Appearance      =   0  'Flat
      Caption         =   "+"
      Height          =   300
      Index           =   0
      Left            =   4530
      MaskColor       =   &H00000000&
      TabIndex        =   15
      ToolTipText     =   "Buscar fecha"
      Top             =   4800
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   2
      Left            =   7140
      MaxLength       =   30
      TabIndex        =   3
      Tag             =   "Descripción|T|N|||raportacion|descripcion|||"
      Top             =   4860
      Width           =   1065
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   11970
      TabIndex        =   7
      Top             =   5475
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   13140
      TabIndex        =   9
      Top             =   5490
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   1
      Left            =   3150
      MaxLength       =   10
      TabIndex        =   1
      Tag             =   "Fecha|F|N|||raportacion|fecaport|dd/mm/yyyy|S|"
      Top             =   4800
      Width           =   1320
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   0
      Left            =   240
      MaxLength       =   6
      TabIndex        =   0
      Tag             =   "Código|N|N|0|999999|raportacion|codsocio|000000|S|"
      Top             =   4800
      Width           =   800
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmAPOAportacion.frx":000C
      Height          =   4410
      Left            =   90
      TabIndex        =   11
      Top             =   540
      Width           =   14220
      _ExtentX        =   25083
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
      Left            =   13140
      TabIndex        =   14
      Top             =   5490
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Index           =   1
      Left            =   60
      TabIndex        =   8
      Top             =   5370
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
         Left            =   45
         TabIndex        =   10
         Top             =   210
         Width           =   2295
      End
   End
   Begin MSAdodcLib.Adodc adodc1 
      Height          =   375
      Left            =   5790
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
      TabIndex        =   12
      Top             =   0
      Width           =   14415
      _ExtentX        =   25426
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   18
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
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Impresión"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Informe Aportaciones"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Traspaso Aportaciones"
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Carga Kilos"
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Alta Socios"
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Baja Socios"
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   195
         Left            =   5040
         TabIndex        =   13
         Top             =   90
         Visible         =   0   'False
         Width           =   1215
      End
   End
   Begin VB.Label Label1 
      Caption         =   "TOTALES: "
      Height          =   225
      Left            =   10650
      TabIndex        =   20
      Top             =   5040
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
      Begin VB.Menu mnImprimir 
         Caption         =   "&Imprimir"
         Shortcut        =   ^I
      End
      Begin VB.Menu mnImpresion 
         Caption         =   "Informe Aportaciones"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnBarra3 
         Caption         =   "-"
      End
      Begin VB.Menu mnTraspaso 
         Caption         =   "&Traspaso Aportaciones"
         Shortcut        =   ^T
      End
      Begin VB.Menu mnCargaKilos 
         Caption         =   "&Carga de Kilos"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnAltaSocios 
         Caption         =   "Alta Socios"
         Shortcut        =   ^G
      End
      Begin VB.Menu mnBajaSocios 
         Caption         =   "Baja Socios"
         Shortcut        =   ^H
      End
      Begin VB.Menu mnBarra4 
         Caption         =   "-"
      End
      Begin VB.Menu mnSalir 
         Caption         =   "&Salir"
         Shortcut        =   ^S
      End
   End
End
Attribute VB_Name = "frmAPOAportacion"
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
Private WithEvents frmTra As frmManTraba 'mantenimiento de trabajadores
Attribute frmTra.VB_VarHelpID = -1
Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmApo As frmAPOTipos 'tipos de aportacion
Attribute frmApo.VB_VarHelpID = -1
Private WithEvents frmMens As frmMensajes 'sacamos la campaña
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
Dim indice As Byte 'Index del text1 on es poses els datos retornats des d'atres Formularis de Mtos
Dim indCodigo As Byte 'Index del text1 on es poses els datos retornats des d'atres Formularis de Mtos
Dim I As Integer
Dim SobreCampanya As String


' utilizado para buscar por checks
Private BuscaChekc As String


Private Sub PonerModo(vModo)
Dim b As Boolean

    Modo = vModo
    
    b = (Modo = 2)
    If b Then
        PonerContRegIndicador lblIndicador, adodc1, CadB
    Else
        PonerIndicador lblIndicador, Modo
    End If
    
    For I = 0 To txtAux.Count - 1
        txtAux(I).visible = Not b
    Next I
    
    txtAux2(0).visible = Not b
    txtAux2(6).visible = Not b
    
    For I = 0 To btnBuscar.Count - 1
        btnBuscar(I).visible = Not b
    Next I
    
    CmdAceptar.visible = Not b
    cmdCancelar.visible = Not b
    DataGrid1.Enabled = b
    
    'Si es regresar
    If DatosADevolverBusqueda <> "" Then cmdRegresar.visible = b
    
    PonerLongCampos
    PonerModoOpcionesMenu 'Activar/Desact botones de menu segun Modo
    PonerOpcionesMenu  'En funcion del usuario
    
    'Si estamos modo Modificar bloquear clave primaria
    BloquearTxt txtAux(0), (Modo = 4)
    BloquearTxt txtAux(1), (Modo = 4)
    BloquearTxt txtAux(6), (Modo = 4)
    BloquearBtn Me.btnBuscar(0), (Modo = 4)
    BloquearBtn Me.btnBuscar(1), (Modo = 4)
    BloquearBtn Me.btnBuscar(2), (Modo = 4)
    
    
    'El nro de parte unicamente lo podemos buscar
'    txtAux(8).Enabled = (Modo = 1)
'    txtAux(8).visible = (Modo = 1)
    
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
    
    'Traspaso de aportaciones
    Toolbar1.Buttons(13).Enabled = b
    Me.mnTraspaso.Enabled = b
    
    
    'Insertar
    Toolbar1.Buttons(6).Enabled = b And Not DeConsulta
    Me.mnNuevo.Enabled = b And Not DeConsulta
    
    ' Carga kilos, alta socios y baja socios --> solo para mogente
    Toolbar1.Buttons(14).Enabled = b And Not DeConsulta And vParamAplic.Cooperativa = 3
    Toolbar1.Buttons(14).visible = b And Not DeConsulta And vParamAplic.Cooperativa = 3
    Me.mnCargaKilos.Enabled = b And Not DeConsulta And vParamAplic.Cooperativa = 3
    Me.mnCargaKilos.visible = b And Not DeConsulta And vParamAplic.Cooperativa = 3
    
    Toolbar1.Buttons(15).Enabled = b And Not DeConsulta And vParamAplic.Cooperativa = 3
    Toolbar1.Buttons(15).visible = b And Not DeConsulta And vParamAplic.Cooperativa = 3
    Me.mnAltaSocios.Enabled = b And Not DeConsulta And vParamAplic.Cooperativa = 3
    Me.mnAltaSocios.visible = b And Not DeConsulta And vParamAplic.Cooperativa = 3
    
    Toolbar1.Buttons(16).Enabled = b And Not DeConsulta And vParamAplic.Cooperativa = 3
    Toolbar1.Buttons(16).visible = b And Not DeConsulta And vParamAplic.Cooperativa = 3
    Me.mnBajaSocios.Enabled = b And Not DeConsulta And vParamAplic.Cooperativa = 3
    Me.mnBajaSocios.visible = b And Not DeConsulta And vParamAplic.Cooperativa = 3
    
    
    
    b = (b And adodc1.Recordset.RecordCount > 0) And Not DeConsulta
    'Modificar
    Toolbar1.Buttons(7).Enabled = b
    Me.mnModificar.Enabled = b
    'Eliminar
    Toolbar1.Buttons(8).Enabled = b
    Me.mnEliminar.Enabled = b
    'Imprimir
    Toolbar1.Buttons(10).Enabled = b
    Me.mnImprimir.Enabled = b

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
        anc = anc + 206
    Else
        anc = anc + DataGrid1.RowTop(DataGrid1.Row) + 5
    End If
    For I = 0 To txtAux.Count - 1
        txtAux(I).Text = ""
    Next I
    txtAux(1).Text = Format(Now, "dd/mm/yyyy")
    txtAux2(0).Text = ""
    txtAux2(6).Text = ""
    
    txtAux(4).Text = 0
    txtAux(5).Text = 0

    LLamaLineas anc, 3 'Pone el form en Modo=3, Insertar
       
    'Ponemos el foco
    PonerFoco txtAux(0)
End Sub

Private Sub BotonTraspasoAportaciones()
    
    frmAPOTraspaso.Show vbModal

End Sub

Private Sub BotonVerTodos()
    CadB = ""
    CargaGrid ""
    PonerModo 2
End Sub

Private Sub BotonBuscar()
    ' ***************** canviar per la clau primaria ********
    CargaGrid "raportacion.codsocio = -1"
    '*******************************************************************************
    'Buscar
    For I = 0 To txtAux.Count - 1
        txtAux(I).Text = ""
    Next I
    Me.txtAux2(0).Text = ""
    Me.txtAux2(6).Text = ""
    
    
    LLamaLineas DataGrid1.Top + 206, 1 'Pone el form en Modo=1, Buscar
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
        anc = DataGrid1.RowTop(DataGrid1.Row) + 540 '565 '495 '545
    End If
    
    'Llamamos al form
    txtAux(0).Text = DataGrid1.Columns(0).Text ' codsocio
    txtAux2(0).Text = DataGrid1.Columns(1).Text ' nomsocio
    txtAux(1).Text = DataGrid1.Columns(2).Text ' fecha
    
    txtAux(6).Text = DataGrid1.Columns(3).Text 'codaport
    txtAux2(6).Text = DataGrid1.Columns(4).Text 'nomaport
    
    txtAux(2).Text = DataGrid1.Columns(5).Text 'descripcion
    txtAux(3).Text = DataGrid1.Columns(6).Text 'campaña
    txtAux(4).Text = DataGrid1.Columns(7).Text 'kilos
    txtAux(5).Text = DataGrid1.Columns(8).Text 'importe
    
    ' ***** canviar-ho pel nom del camp del combo *********
'    SelComboBool DataGrid1.Columns(2).Text, Combo1(0)
    ' *****************************************************
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
    txtAux2(6).Top = alto
    For I = 0 To btnBuscar.Count - 1
        btnBuscar(I).Top = alto - 15
    Next I
    
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
    SQL = "¿Seguro que desea eliminar el Registro?"
    SQL = SQL & vbCrLf & "Socio: " & adodc1.Recordset.Fields(0) & " " & adodc1.Recordset.Fields(1)
    SQL = SQL & vbCrLf & "Fecha: " & adodc1.Recordset.Fields(2)
    SQL = SQL & vbCrLf & "Tipo Aportación: " & adodc1.Recordset.Fields(3) & " " & adodc1.Recordset.Fields(4)
    
    
    If MsgBox(SQL, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        NumRegElim = adodc1.Recordset.AbsolutePosition
        SQL = "Delete from raportacion where codsocio=" & adodc1.Recordset!Codsocio
        SQL = SQL & " and fecaport = " & DBSet(adodc1.Recordset!fecaport, "F")
        SQL = SQL & " and codaport = " & DBLet(adodc1.Recordset!Codaport, "N")
        
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
 TerminaBloquear
    
    Select Case Index
        Case 1 'socio
            AbrirFrmSocio 0
    
        Case 2 'tipo de aportacion
            AbrirFrmAportacion 0
    
        Case 0 ' Fecha
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
        
            btnBuscar(Index).Tag = Index '<===
            ' *** repasar si el camp es txtAux o Text1 ***
            If Index = 0 Then
                indice = 1
                If txtAux(1).Text <> "" Then frmC.NovaData = txtAux(1).Text
            End If
            
            ' ********************************************
        
            frmC.Show vbModal
            Set frmC = Nothing
            ' *** repasar si el camp es txtAux o Text1 ***
            If Index = 0 Then
                PonerFoco txtAux(1) '<===
            End If
            ' ********************************************
     
    End Select
    
    If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Me.adodc1, 1
End Sub

Private Sub cmdAceptar_Click()
    Dim I As Integer

    Select Case Modo
        Case 1 'BUSQUEDA
            CadB = ObtenerBusqueda2(Me, BuscaChekc, 1) ' ObtenerBusqueda3(Me, False, BuscaChekc)
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
            If DatosOk Then
                If ModificaDesdeFormulario(Me) Then
'                If ModificaDesdeForm Then
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
                    adodc1.Recordset.Find (adodc1.Recordset.Fields(1).Name & " =" & I)
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
    If Modo = 2 Then PonerContRegIndicador Me.lblIndicador, adodc1, CadB
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
                SituarData Me.adodc1, "codsocio=" & CodigoActual, "", True
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
        .Buttons(10).Image = 10  'Imprimir
        .Buttons(11).Image = 26  'Informe de aportaciones
        
        .Buttons(13).Image = 32  'Traspaso Aportaciones
        .Buttons(14).Image = 17  'Carga de kilos
        .Buttons(15).Image = 19  'Alta de socios
        .Buttons(16).Image = 20  'Baja de socios
        .Buttons(18).Image = 11  'Salir
        
    End With

    '## A mano
'    chkVistaPrevia.Value = CheckValueLeer(Name)
    
    
    '****************** canviar la consulta *********************************+
    CadenaConsulta = "SELECT raportacion.codsocio, rsocios.nomsocio, raportacion.fecaport, raportacion.codaport, rtipoapor.nomaport, "
    CadenaConsulta = CadenaConsulta & " raportacion.descripcion, raportacion.campanya, raportacion.kilos, raportacion.importe   "
    CadenaConsulta = CadenaConsulta & " FROM  raportacion, rsocios, rtipoapor  "
    CadenaConsulta = CadenaConsulta & " WHERE raportacion.codsocio = rsocios.codsocio and  "
    CadenaConsulta = CadenaConsulta & " raportacion.codaport = rtipoapor.codaport "
    
    '[Monica]18/01/2016: si viene de socios cargamos la tabla
    If CodigoActual <> "" Then CadenaConsulta = CadenaConsulta & " and raportacion.codsocio = " & DBSet(CodigoActual, "N")
    '************************************************************************
    '[Monica]18/01/2016: si viene de socios cargamos la tabla
    If CodigoActual <> "" Then
        CadB = ""
        CargaGrid ""
    Else
        CadB = ""
        CargaGrid "raportacion.codsocio = -1"
    End If
    
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


Private Sub frmApo_DatoSeleccionado(CadenaSeleccion As String)
    txtAux(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1) 'codigo tipo de aportacion
    txtAux2(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2) 'nombre tipo de aportacion
End Sub

Private Sub frmC_Selec(vFecha As Date)
    ' *** repasar si el camp es txtAux o Text1 ***
    If indice = 1 Then
        txtAux(1).Text = Format(vFecha, "dd/mm/yyyy") '<===
    End If
    ' ********************************************
End Sub


Private Sub frmMens_DatoSeleccionado(CadenaSeleccion As String)
    If CadenaSeleccion <> "" Then
        SobreCampanya = CadenaSeleccion
    End If
End Sub

Private Function ProcesoCargaKilos(Campaña As String) As Boolean
Dim SQL As String
Dim vCampAnt As CCampAnt
Dim Sql2 As String
Dim Campanya As String
Dim CadValues As String
Dim b As Boolean
Dim Rs As ADODB.Recordset
Dim vFecha As Date


    On Error GoTo eCargaKilos

    ProcesoCargaKilos = False

    conn.BeginTrans


    SQL = "select codsocio, sum(kilosnet) kilos from (rhisfruta inner join variedades on rhisfruta.codvarie = variedades.codvarie) "
    SQL = SQL & " inner join productos on variedades.codprodu = productos.codprodu "
    SQL = SQL & " where productos.codgrupo = 5 group by 1 order by 1 "
    
'    Campanya = Mid(CStr(Year(vParam.FecIniCam) - 1), 3, 2) & "/" & Mid(CStr(Year(vParam.FecFinCam) - 1), 3, 2)
    
    CadValues = ""
    
    b = True
'    vFecha = DateAdd("yyyy", (-1), vParam.FecFinCam)
    
    ' Nos quedamos en la campaña actual
    If Campaña = vUsu.CadenaConexion Then
    
        Campanya = Mid(CStr(Year(vParam.FecIniCam)), 3, 2) & "/" & Mid(CStr(Year(vParam.FecFinCam)), 3, 2)
    
        vFecha = vParam.FecFinCam
    
    
        Set Rs = New ADODB.Recordset
        Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        While Not Rs.EOF And b
            Sql2 = "select * from raportacion where codaport = 1 and fecaport = " & DBSet(vFecha, "F") & " and codsocio = " & DBSet(Rs!Codsocio, "N")
            
            If TotalRegistrosConsulta(Sql2) <> 0 Then
                MsgBox "Existe una aportación para el socio " & Rs!Codsocio & ". Revise.", vbExclamation
                b = False
            Else
                CadValues = CadValues & "(" & DBSet(Rs!Codsocio, "N") & "," & DBSet(vFecha, "F") & ",1,"
                CadValues = CadValues & DBSet("CAMPAÑA " & Campanya, "T") & "," & DBSet(Campanya, "T") & "," & DBSet(Rs!Kilos, "N") & ",0),"
            End If
            
            Rs.MoveNext
        Wend
        
        Set Rs = Nothing
    Else

        Campanya = Mid(CStr(Year(vParam.FecIniCam) - 1), 3, 2) & "/" & Mid(CStr(Year(vParam.FecFinCam) - 1), 3, 2)

        vFecha = DateAdd("yyyy", (-1), vParam.FecFinCam)


        Set vCampAnt = New CCampAnt
                
        If vCampAnt.Leer = 0 Then
            If AbrirConexionCampAnterior(vCampAnt.BaseDatos) Then
                Set Rs = New ADODB.Recordset
                Rs.Open SQL, ConnCAnt, adOpenForwardOnly, adLockPessimistic, adCmdText
            
                While Not Rs.EOF And b
                    Sql2 = "select * from raportacion where codaport = 1 and fecaport = " & DBSet(vFecha, "F") & " and codsocio = " & DBSet(Rs!Codsocio, "N")
                    
                    If TotalRegistrosConsulta(Sql2) <> 0 Then
                        MsgBox "Existe una aportación para el socio " & Rs!Codsocio & ". Revise.", vbExclamation
                        b = False
                    Else
                        CadValues = CadValues & "(" & DBSet(Rs!Codsocio, "N") & "," & DBSet(vFecha, "F") & ",1,"
                        CadValues = CadValues & DBSet("CAMPAÑA " & Campanya, "T") & "," & DBSet(Campanya, "T") & "," & DBSet(Rs!Kilos, "N") & ",0),"
                    End If
                    
                    Rs.MoveNext
                Wend
                
                Set Rs = Nothing
            
            End If
        End If
    End If

    If b And CadValues <> "" Then
        CadValues = Mid(CadValues, 1, Len(CadValues) - 1)
        SQL = "insert into raportacion (codsocio,fecaport,codaport,descripcion,campanya,kilos,importe) values  " & CadValues
        
        conn.Execute SQL
    End If

eCargaKilos:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Proceso Carga de Datos", Err.Description
        b = False
    End If
    
    If Not b Then
        conn.RollbackTrans
    Else
        ProcesoCargaKilos = True
        conn.CommitTrans
    End If
End Function


Private Sub frmSoc_DatoSeleccionado(CadenaSeleccion As String)
    txtAux(indice).Text = RecuperaValor(CadenaSeleccion, 1) 'codigo socio
    txtAux2(indice).Text = RecuperaValor(CadenaSeleccion, 2) 'nombre socio
End Sub


Private Sub mnAltaSocios_Click()
    AbrirListadoAPOR 9
End Sub

Private Sub mnBajaSocios_Click()
    AbrirListadoAPOR 10
End Sub

Private Sub mnBuscar_Click()
    BotonBuscar
End Sub

Private Sub mnCargaKilos_Click()
    
    SobreCampanya = ""
    
    
    DesBloqueoManual ("CARAPO")
    If Not BloqueoManual("CARAPO", "1") Then
        MsgBox "No se puede realizar la Carga de Aportaciones. Hay otro usuario realizándola.", vbExclamation
        Screen.MousePointer = vbDefault
    Else
        
        
        Set frmMens = frmMensajes
        
        frmMens.OpcionMensaje = 43
        frmMens.Show vbModal
        
        Set frmMens = Nothing
        If SobreCampanya <> "" Then
            If ProcesoCargaKilos(SobreCampanya) Then
                MsgBox "Proceso realizado correctamente.", vbExclamation
                CargaGrid "fecaport = " & DBSet(DateAdd("yyyy", (-1), vParam.FecFinCam), "F") & " and raportacion.codaport = 1"
            End If
        End If
    End If
End Sub

Private Sub mnEliminar_Click()
    BotonEliminar
End Sub

Private Sub mnImpresion_Click()
    AbrirListadoAPOR (4)
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

Private Sub mnTraspaso_Click()
    BotonTraspasoAportaciones
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
                mnImprimir_Click
        Case 11
                mnImpresion_Click
        Case 13
                mnTraspaso_Click
                
        Case 14 ' carga de kilos
                mnCargaKilos_Click
        Case 15 ' alta de socios
                mnAltaSocios_Click
        Case 16 ' baja de socios
                mnBajaSocios_Click
        Case 18
                mnSalir_Click
    End Select
End Sub

Private Sub CargaGrid(Optional vSQL As String, Optional Ascendente As Boolean)
    Dim SQL As String
    Dim tots As String
    
'    adodc1.ConnectionString = Conn
    If vSQL <> "" Then
        SQL = CadenaConsulta & " AND " & vSQL
    Else
        SQL = CadenaConsulta
    End If
    SQL = SQL & " ORDER BY  raportacion.fecaport, raportacion.codsocio, raportacion.codaport "
    
    CargaGridGnral Me.DataGrid1, Me.adodc1, SQL, PrimeraVez
    
    ' *******************canviar els noms i si fa falta la cantitat********************
    tots = "S|txtAux(0)|T|Codigo|730|;S|btnBuscar(1)|B||195|;S|txtAux2(0)|T|Socio|2700|;"
    tots = tots & "S|txtAux(1)|T|Fecha|1200|;S|btnBuscar(0)|B||195|;"
    tots = tots & "S|txtAux(6)|T|Codigo|800|;S|btnBuscar(2)|B||195|;S|txtAux2(6)|T|Tipo Aport.|1800|;"
    tots = tots & "S|txtAux(2)|T|Descripcion|3000|;"
    tots = tots & "S|txtAux(3)|T|Campaña|1000|;"
    tots = tots & "S|txtAux(4)|T|Kilos|1200|;"
    tots = tots & "S|txtAux(5)|T|Importe|1200|;"
    
    arregla tots, DataGrid1, Me
    
    DataGrid1.ScrollBars = dbgAutomatic
    DataGrid1.Columns(0).Alignment = dbgLeft
    DataGrid1.Columns(3).Alignment = dbgLeft
    
    CalcularTotales SQL
    
'    DataGrid1.Columns(10).Alignment = dbgCenter
'    DataGrid1.Columns(12).Alignment = dbgCenter
End Sub

Private Sub txtAux_GotFocus(Index As Integer)
    ConseguirFocoLin txtAux(Index)
End Sub


Private Sub txtAux_LostFocus(Index As Integer)
Dim cadMen As String

    If Not PerderFocoGnral(txtAux(Index), Modo) Then Exit Sub
    
    Select Case Index
        Case 0 'codigo de socio
            If PonerFormatoEntero(txtAux(Index)) Then
                If Modo = 1 Then Exit Sub
                txtAux2(Index).Text = PonerNombreDeCod(txtAux(Index), "rsocios", "nomsocio")
                If txtAux2(Index).Text = "" Then
                    cadMen = "No existe el Socio: " & txtAux(Index).Text & vbCrLf
                    MsgBox cadMen, vbExclamation
                    PonerFoco txtAux(Index)
                End If
            Else
                txtAux2(Index).Text = ""
            End If
            
        
        Case 1 'fecha
            PonerFormatoFecha txtAux(Index)
            
        Case 5 'importe
            PonerFormatoDecimal txtAux(Index), 3
    
        Case 4 'kilos
            PonerFormatoEntero txtAux(Index)
    
        Case 6 'codigo de tipo de aportacion
            If PonerFormatoEntero(txtAux(Index)) Then
                If Modo = 1 Then Exit Sub
                txtAux2(Index).Text = PonerNombreDeCod(txtAux(Index), "rtipoapor", "nomaport")
                If txtAux2(Index).Text = "" Then
                    cadMen = "No existe el Tipo de Aportación " & txtAux(Index).Text & ". Reintroduzca." & vbCrLf
                    cadMen = cadMen & "¿Desea crearlo?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmApo = New frmAPOTipos
                        frmApo.DatosADevolverBusqueda = "0|1|"
                        txtAux(Index).Text = ""
                        TerminaBloquear
                        frmApo.Show vbModal
                        Set frmApo = Nothing
                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, adodc1, 1
                    Else
                        txtAux(Index).Text = ""
                    End If
                    PonerFoco txtAux(Index)
                End If
            End If
    
    End Select
    
End Sub

Private Function DatosOk() As Boolean
'Dim Datos As String
Dim b As Boolean
Dim SQL As String
Dim Mens As String


    b = CompForm(Me)
    If Not b Then Exit Function
    
    If Modo = 3 Then   'Estamos insertando
        SQL = "select count(*) from raportacion where codsocio = " & DBSet(txtAux(0).Text, "N")
        SQL = SQL & " and fecha = " & DBSet(txtAux(1).Text, "F")
        SQL = SQL & " and codaport = " & DBSet(txtAux(6).Text, "N")
        If TotalRegistros(SQL) <> 0 Then
            MsgBox "El socio existe para esta fecha, tipo de aportación. Reintroduzca.", vbExclamation
            PonerFoco txtAux(0)
            b = False
        End If
    End If
    
'    If b And (Modo = 3 Or Modo = 4) Then
'        If Not EntreFechas(vParam.FecIniCam, txtAux(1).Text, vParam.FecFinCam) Then
'            MsgBox "La fecha introducida no se encuentra dentro de campaña. Revise.", vbExclamation
'            b = False
'        End If
'    End If
    
    DatosOk = b
End Function


Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
    PonerOpcionesMenuGeneralNew Me
End Sub

Private Sub printNou()
    With frmImprimir2
        .cadTabla2 = "raportacion"
        .Informe2 = "rManAportacion.rpt"
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
        .OtrosParametros2 = "pEmpresa='" & vEmpresa.nomempre & "'|pOrden={raportacion.fecaport}|"
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
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me
End Sub

Private Sub KEYBusqueda(KeyAscii As Integer, indice As Integer)
    KeyAscii = 0
    btnBuscar_Click (indice)
End Sub

Private Sub AbrirFrmAportacion(indice As Integer)
    indCodigo = 6
    Set frmApo = New frmAPOTipos
    frmApo.DatosADevolverBusqueda = "0|1|"
    frmApo.CodigoActual = txtAux(indCodigo)
    frmApo.Show vbModal
    Set frmApo = Nothing
    
    PonerFoco txtAux(indCodigo)
    
End Sub


Private Sub AbrirFrmSocio(indice As Integer)
    indCodigo = 0
    Set frmSoc = New frmManSocios
    frmSoc.DatosADevolverBusqueda = "0|1|"
    frmSoc.Show vbModal
    Set frmSoc = Nothing
    
    PonerFoco txtAux(indCodigo)
End Sub

Private Function ModificaDesdeForm() As Boolean
Dim SQL As String

    On Error GoTo eModificaDesdeForm
    
    ModificaDesdeForm = False
    
    SQL = "update raportacion set "
    SQL = SQL & " importe = " & DBSet(ImporteSinFormato(txtAux(2).Text), "N")
    SQL = SQL & ", compleme = " & DBSet(ImporteSinFormato(txtAux(3).Text), "N")
    SQL = SQL & ", penaliza = " & DBSet(ImporteSinFormato(txtAux(4).Text), "N")
    SQL = SQL & " where codcapat = " & DBSet(txtAux(0).Text, "N")
    SQL = SQL & " and fechahora = " & DBSet(txtAux(1).Text, "F")
    SQL = SQL & " and codtraba = " & DBSet(txtAux(7).Text, "N")
    SQL = SQL & " and codvarie = " & DBSet(txtAux(6).Text, "N")
    
    conn.Execute SQL
    
    ModificaDesdeForm = True
    Exit Function
    
eModificaDesdeForm:
    MuestraError Err.Number, "Modificando registro", Err.Description
End Function

Private Sub CalcularTotales(Cadena As String)
Dim Importe  As Currency
Dim Kilos As Long
Dim Rs As ADODB.Recordset
Dim SQL As String

    On Error Resume Next
    
    SQL = "select sum(kilos) kilos , sum(importe) importe from (" & Cadena & ") aaaaa"
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    Importe = 0
    Kilos = 0
    txtAux2(1).Text = ""
    txtAux2(2).Text = ""
    
    If TotalRegistrosConsulta(Cadena) = 0 Then Exit Sub
    
    If Not Rs.EOF Then
        If Rs.Fields(0).Value <> 0 Then Kilos = DBLet(Rs.Fields(0).Value, "N") 'Solo es para saber que hay registros que mostrar
        If Rs.Fields(1).Value <> 0 Then Importe = DBLet(Rs.Fields(1).Value, "N") 'Solo es para saber que hay registros que mostrar
    
        txtAux2(1).Text = Format(Kilos, "###,###,##0")
        txtAux2(2).Text = Format(Importe, "###,###,##0.00")
    End If
    Rs.Close
    Set Rs = Nothing

    
    DoEvents
    
End Sub


