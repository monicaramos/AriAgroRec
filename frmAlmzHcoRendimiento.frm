VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmAlmzHcoRendimiento 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Hist�rico de Rendimientos"
   ClientHeight    =   5925
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9090
   Icon            =   "frmAlmzHcoRendimiento.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5925
   ScaleWidth      =   9090
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   5
      Left            =   2220
      MaxLength       =   8
      TabIndex        =   2
      Text            =   "Hora"
      Top             =   4920
      Width           =   480
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   6
      Left            =   5910
      MaxLength       =   10
      TabIndex        =   6
      Tag             =   "Rendimiento|N|N|||rrendim|rendimiento|#,##0.0000||"
      Text            =   "rendimient"
      Top             =   4920
      Width           =   900
   End
   Begin VB.CommandButton btnBuscar 
      Appearance      =   0  'Flat
      Caption         =   "+"
      Height          =   300
      Index           =   0
      Left            =   1980
      MaskColor       =   &H00000000&
      TabIndex        =   16
      ToolTipText     =   "Buscar fecha"
      Top             =   4920
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   1
      Left            =   1080
      MaxLength       =   10
      TabIndex        =   1
      Text            =   "fecha"
      Top             =   4920
      Width           =   900
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   4
      Left            =   4890
      MaxLength       =   10
      TabIndex        =   5
      Tag             =   "Humedad|N|N|||rrendim|humedad|###,##0.00||"
      Text            =   "humedad"
      Top             =   4920
      Width           =   900
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   3
      Left            =   3810
      MaxLength       =   10
      TabIndex        =   4
      Tag             =   "Acidez|N|N|||rrendim|acidez|###,##0.00||"
      Text            =   "acidez"
      Top             =   4920
      Width           =   900
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   6615
      TabIndex        =   7
      Top             =   5265
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   7770
      TabIndex        =   8
      Top             =   5265
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   2
      Left            =   2760
      MaxLength       =   7
      TabIndex        =   3
      Tag             =   "Albaran|N|N|0|9999999|rrendim|numalbar|0000000|S|"
      Text            =   "albaran"
      Top             =   4920
      Width           =   900
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   0
      Left            =   180
      MaxLength       =   7
      TabIndex        =   0
      Tag             =   "Muestra|N|N|0|9999999|rrendim|nromuestra|0000000|S|"
      Text            =   "Muestr"
      Top             =   4890
      Width           =   800
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmAlmzHcoRendimiento.frx":000C
      Height          =   4410
      Left            =   135
      TabIndex        =   12
      Top             =   540
      Width           =   8715
      _ExtentX        =   15372
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
      Left            =   7755
      TabIndex        =   15
      Top             =   5295
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
      Width           =   9090
      _ExtentX        =   16034
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
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
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
      Index           =   7
      Left            =   6990
      MaxLength       =   20
      TabIndex        =   9
      Tag             =   "Fecha Muestra|FH|N|||rrendim|fecha|yyyy-mm-dd hh:mm:ss|S|"
      Text            =   "Fecha"
      Top             =   4590
      Width           =   1395
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
Attribute VB_Name = "frmAlmzHcoRendimiento"
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
Private CadB1 As String

Private WithEvents frmVar As frmComVar 'variedades
Attribute frmVar.VB_VarHelpID = -1
Private WithEvents frmC1 As frmCal ' calendario fecha
Attribute frmC1.VB_VarHelpID = -1
' utilizado para buscar por checks
Private BuscaChekc As String

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
    
    btnBuscar(0).visible = Not b

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
    BloquearTxt txtAux(1), (Modo = 4)
    BloquearTxt txtAux(5), (Modo = 4)
    BloquearTxt txtAux(2), (Modo = 4)
    
    BloquearBtn Me.btnBuscar(0), (Modo = 4)
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
        anc = anc + 206
    Else
        anc = anc + DataGrid1.RowTop(DataGrid1.Row) + 5
    End If
    txtAux(0).Text = NumF
    FormateaCampo txtAux(0)
    For i = 1 To txtAux.Count - 1
        txtAux(i).Text = ""
    Next i

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
    CargaGrid "rrendim.nromuestra = -1"
    '*******************************************************************************
    'Buscar
    For i = 0 To txtAux.Count - 1
        txtAux(i).Text = ""
    Next i
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
    txtAux(5).Text = DataGrid1.Columns(2).Text
    txtAux(2).Text = DataGrid1.Columns(3).Text
    txtAux(3).Text = DataGrid1.Columns(4).Text
    txtAux(4).Text = DataGrid1.Columns(5).Text
    txtAux(6).Text = DataGrid1.Columns(6).Text
    txtAux(7).Text = DataGrid1.Columns(7).Text
    ' ***** canviar-ho pel nom del camp del combo *********
    
    LLamaLineas anc, 4 'Pone el form en Modo=4, Modificar
   
    'Como es modificar
    PonerFoco txtAux(3)
    Screen.MousePointer = vbDefault
End Sub

Private Sub LLamaLineas(alto As Single, xModo As Byte)
    DeseleccionaGrid Me.DataGrid1
    PonerModo xModo
    
    'Fijamos el ancho
    For i = 0 To txtAux.Count - 1
        txtAux(i).Top = alto
    Next i
    For i = 0 To 0
        btnBuscar(0).Top = alto
    Next i

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
    Sql = "�Seguro que desea eliminar el Registro?"
    Sql = Sql & vbCrLf & "Muestra: " & adodc1.Recordset.Fields(0)
    Sql = Sql & vbCrLf & "Fecha: " & adodc1.Recordset.Fields(2)
    Sql = Sql & vbCrLf & "Albar�n: " & adodc1.Recordset.Fields(3)
    
    If MsgBox(Sql, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        NumRegElim = adodc1.Recordset.AbsolutePosition
        Sql = "Delete from rrendim where nromuestra=" & adodc1.Recordset!NroMuestra
        Sql = Sql & " and fecha = " & DBSet(adodc1.Recordset!Fecha, "FH")
        Sql = Sql & " and numalbar = " & adodc1.Recordset!numalbar
        
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
'Modificar el MaxLength del campo en funcion de si es modo de b�squeda o no
'para los campos que permitan introducir criterios m�s largos del tama�o del campo
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
        Case 0 'fecha
            
            Set frmC1 = New frmCal
            
            esq = btnBuscar(Index).Left
            dalt = btnBuscar(Index).Top
                
            Set obj = btnBuscar(Index).Container
              
            While btnBuscar(Index).Parent.Name <> obj.Name
                   esq = esq + obj.Left
                   dalt = dalt + obj.Top
                   Set obj = obj.Container
            Wend
            
            menu = Me.Height - Me.ScaleHeight 'ac� tinc el heigth del men� i de la toolbar
        
            frmC1.Left = esq + btnBuscar(Index).Parent.Left + 30
            frmC1.Top = dalt + btnBuscar(Index).Parent.Top + btnBuscar(Index).Height + menu - 40
        
            btnBuscar(0).Tag = Index '<===
            Select Case Index
                Case 0
                    indice = 1
            End Select
            ' *** repasar si el camp es txtAux o Text1 ***
            If txtAux(indice).Text <> "" Then frmC1.NovaData = txtAux(indice).Text
            ' ********************************************
        
            frmC1.Show vbModal
            Set frmC1 = Nothing
            ' *** repasar si el camp es txtAux o Text1 ***
            PonerFoco txtAux(indice) '<===
            ' ********************************************
            
    End Select
            
    If Modo = 4 Then BLOQUEADesdeFormulario2 Me, Me.adodc1, 1
            
End Sub


Private Sub cmdAceptar_Click()
    Dim i As Integer

    Select Case Modo
        Case 1 'BUSQUEDA
            CadB = ObtenerBusqueda(Me, BuscaChekc)
            CadB1 = Replace(Replace(CadB, "date(rrendim.fecha)", "rrendim.fecha"), "time(rrendim.fecha)", "time[rrendim.fecha]")
            If CadB <> "" Then
                CargaGrid CadB
                PonerModo 2
'                lblIndicador.Caption = "BUSQUEDA: " & PonerContRegistros(Me.adodc1)
                PonerFocoGrid Me.DataGrid1
            End If
            
        Case 3 'INSERTAR
            If DatosOk Then
                txtAux(7).Text = Format(txtAux(1).Text, "yyyy-mm-dd") & " " & Format(txtAux(5).Text, "hh:mm:ss")
                txtAux(1).Tag = ""
                txtAux(5).Tag = ""
                
                If InsertarDesdeForm(Me) Then
                    CargaGrid CadB
                    If (DatosADevolverBusqueda <> "") And NuevoCodigo <> "" Then
                        cmdCancelar_Click
'                        If Not adodc1.Recordset.EOF Then adodc1.Recordset.MoveLast
                        If Not adodc1.Recordset.EOF Then
                            SituarDataMULTI adodc1, "nromuestra = " & txtAux(0) & " and fecha = " & DBSet(txtAux(7).Text, "FH"), "" ' Find (adodc1.Recordset.Fields(2).Name & " =" & NuevoCodigo)
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
                txtAux(7).Text = Format(txtAux(1).Text, "yyyy-mm-dd") & " " & Format(txtAux(5).Text, "hh:mm:ss")
                txtAux(1).Tag = ""
                txtAux(5).Tag = ""
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
                SituarData Me.adodc1, "nromuestra=" & CodigoActual, "", True
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
        .Buttons(12).Image = 11  'Salir
    End With

    '## A mano
'    chkVistaPrevia.Value = CheckValueLeer(Name)
    
    
    '****************** canviar la consulta *********************************+
    CadenaConsulta = "SELECT rrendim.nromuestra, date(fecha) F1, time(fecha) H1, numalbar, acidez, humedad, rendimiento, fecha"
    CadenaConsulta = CadenaConsulta & " FROM rrendim"
    CadenaConsulta = CadenaConsulta & " WHERE (1=1) "
    '************************************************************************
    
    
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



Private Sub frmC1_Selec(vFecha As Date)
    ' *** repasar si el camp es txtAux o Text1 ***
    txtAux(indice).Text = Format(vFecha, "dd/mm/yyyy") '<===
    ' ********************************************
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
    
    Dim Sql As String
    Dim tots As String
    
'    adodc1.ConnectionString = Conn
    If vSQL <> "" Then
        Sql = CadenaConsulta & " AND " & vSQL
    Else
        Sql = CadenaConsulta
    End If
    
    txtAux(1).Tag = "Fecha|FHF|N|||rrendim|fecha|dd/mm/yyyy||"
    txtAux(5).Tag = "Hora|FHH|N|||rrendim|fecha|hh:mm:ss||"
    
    '********************* canviar el ORDER BY *********************++
    Sql = Sql & " ORDER BY rrendim.nromuestra, rrendim.fecha, rrendim.numalbar "
    '**************************************************************++
    
    CargaGridGnral Me.DataGrid1, Me.adodc1, Sql, PrimeraVez
    
    ' *******************canviar els noms i si fa falta la cantitat********************
    tots = "S|txtAux(0)|T|Muestra|1100|;S|txtAux(1)|T|Fecha|1200|;S|btnBuscar(0)|B|||;S|txtAux(5)|T|Hora|1100|;"
    tots = tots & "S|txtAux(2)|T|Albar�n|1100|;"
    tots = tots & "S|txtAux(3)|T|Acidez|1200|;"
    tots = tots & "S|txtAux(4)|T|Humedad|1200|;"
    tots = tots & "S|txtAux(6)|T|Rendimiento|1200|;N||||0|;"
    
    arregla tots, DataGrid1, Me
    
    DataGrid1.ScrollBars = dbgAutomatic
    DataGrid1.Columns(0).Alignment = dbgLeft

End Sub

Private Sub txtAux_GotFocus(Index As Integer)
    ConseguirFocoLin txtAux(Index)
End Sub


Private Sub txtAux_LostFocus(Index As Integer)
    If Not PerderFocoGnral(txtAux(Index), Modo) Then Exit Sub
    
    Select Case Index
        Case 0 'nro de muestra
            PonerFormatoEntero txtAux(Index)
            
        Case 1 'fecha
            PonerFormatoFecha txtAux(Index)
            
        Case 5 ' hora
            PonerFormatoHora txtAux(Index)
            
        Case 2 ' albaran
            PonerFormatoEntero txtAux(Index)
            
        Case 3, 4 'acidez, humedad
            PonerFormatoDecimal txtAux(Index), 3
            
        Case 6 ' rendimiento
            If PonerFormatoDecimal(txtAux(Index), 7) Then cmdAceptar.SetFocus
            
    End Select
    
    
    If Index = 1 Or Index = 5 Then
        txtAux(7).Text = Trim(Format(txtAux(1).Text, "yyyy-mm-dd") & " " & Format(txtAux(5).Text, "hh:mm:ss"))
    End If
    
    
End Sub

Private Function DatosOk() As Boolean
'Dim Datos As String
Dim b As Boolean
Dim Sql As String
Dim Mens As String

    b = CompForm(Me)
    If Not b Then Exit Function
    
    If Modo = 3 Then   'Estamos insertando
        Sql = "select count(*) from rrendim where nromuestra = " & txtAux(0).Text & " and fecha = '" & txtAux(7).Text & "' and numalbar = " & txtAux(2).Text
        If TotalRegistros(Sql) > 0 Then
            MsgBox "C�digo ya existe. Revise.", vbExclamation
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


Private Sub printNou()
    With frmImprimir2
        .cadTabla2 = "rrendim"
        .Informe2 = "rManRendimiento.rpt"
        If CadB1 <> "" Then
            '.cadRegSelec = Replace(SQL2SF(CadB), "clientes", "clientes_1")
            .cadRegSelec = Replace(SQL2SF(CadB1), "{time[rrendim.fecha]}", "time({rrendim.fecha})")
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
        .OtrosParametros2 = "pEmpresa='" & vEmpresa.nomempre & "'|pOrden={rrendim.nromuestra}|"
        '.OtrosParametros2 = "pEmpresa='" & vEmpresa.NomEmpre & "'|"
        ' *** posar el n� de par�metres que he posat en OtrosParametros2 ***
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
'
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



Private Function SepuedeBorrar() As Boolean
Dim Sql As String
Dim Cad As String

    SepuedeBorrar = False
    SepuedeBorrar = True
End Function

