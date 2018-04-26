VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmTrzManCargas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Manejo de Cargas de Confección"
   ClientHeight    =   5925
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13020
   Icon            =   "frmTrzManCargas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5925
   ScaleWidth      =   13020
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   1
      Left            =   885
      MaxLength       =   8
      TabIndex        =   1
      Tag             =   "Id Palet|N|N|||trzlineas_cargas|idpalet|0000000|S|"
      Text            =   "Id"
      Top             =   2715
      Width           =   555
   End
   Begin VB.CommandButton btnBuscar 
      Appearance      =   0  'Flat
      Caption         =   "+"
      Height          =   290
      Index           =   0
      Left            =   3705
      MaskColor       =   &H00000000&
      TabIndex        =   19
      ToolTipText     =   "Buscar Cliente"
      Top             =   2715
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   5
      Left            =   3135
      MaxLength       =   6
      TabIndex        =   18
      Tag             =   "Socio|N|N|0|999999|trzlineas_cargas|codsocio|000000|N|"
      Text            =   "socio"
      Top             =   2715
      Width           =   555
   End
   Begin VB.TextBox txtAux2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   285
      Index           =   5
      Left            =   3945
      TabIndex        =   17
      Top             =   2715
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   0
      Left            =   285
      MaxLength       =   7
      TabIndex        =   0
      Tag             =   "Linea|N|N|||trzlineas_cargas|linea|0000000|S|"
      Text            =   "Linea"
      Top             =   2715
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   2
      Left            =   1440
      MaxLength       =   10
      TabIndex        =   2
      Tag             =   "Fecha|F|N|||trzlineas_cargas|fecha|dd/mm/yyyy||"
      Text            =   "Fecha"
      Top             =   2715
      Width           =   555
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   3
      Left            =   2145
      MaxLength       =   8
      TabIndex        =   3
      Tag             =   "Hora|FHH|N|||trzlineas_cargas|fechahora|hh:mm:ss||"
      Text            =   "hora"
      Top             =   2715
      Width           =   555
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   6
      Left            =   5865
      MaxLength       =   8
      TabIndex        =   16
      Tag             =   "Campo|N|N|0|99999999|trzlineas_cargas|codcampo|00000000||"
      Text            =   "Campo"
      Top             =   2715
      Width           =   795
   End
   Begin VB.TextBox txtAux2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   285
      Index           =   7
      Left            =   7545
      TabIndex        =   15
      Top             =   2715
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   7
      Left            =   6690
      MaxLength       =   6
      TabIndex        =   14
      Tag             =   "Variedad|N|N|0|999999|trzlineas_cargas|codvarie|000000|N|"
      Text            =   "Var"
      Top             =   2715
      Width           =   555
   End
   Begin VB.CommandButton btnBuscar 
      Appearance      =   0  'Flat
      Caption         =   "+"
      Height          =   290
      Index           =   1
      Left            =   7305
      MaskColor       =   &H00000000&
      TabIndex        =   13
      ToolTipText     =   "Buscar Articulo"
      Top             =   2715
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.TextBox txtAux 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   8
      Left            =   8745
      MaxLength       =   15
      TabIndex        =   12
      Tag             =   "Kilos|N|N|0|999999.99|trzlineas_cargas|numkilos|###,##0.00||"
      Text            =   "Kilos"
      Top             =   2715
      Width           =   780
   End
   Begin VB.CommandButton btnBuscar 
      Appearance      =   0  'Flat
      Caption         =   "+"
      Height          =   290
      Index           =   2
      Left            =   1965
      MaskColor       =   &H00000000&
      TabIndex        =   11
      ToolTipText     =   "Buscar Fecha"
      Top             =   2715
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   10305
      TabIndex        =   4
      Top             =   5445
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   11655
      TabIndex        =   5
      Top             =   5445
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   11655
      TabIndex        =   10
      Top             =   5445
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Index           =   1
      Left            =   90
      TabIndex        =   6
      Top             =   5310
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
         TabIndex        =   7
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
      Width           =   13020
      _ExtentX        =   22966
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   10
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmTrzManCargas.frx":000C
      Height          =   4705
      Left            =   135
      TabIndex        =   20
      Top             =   495
      Width           =   12785
      _ExtentX        =   22543
      _ExtentY        =   8308
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
      Begin VB.Menu mnSalir 
         Caption         =   "&Salir"
         Shortcut        =   ^S
      End
   End
End
Attribute VB_Name = "frmTrzManCargas"
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
' 6. En la funció posamaxlength() repasar el maxlength de TextAux(0)
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
' ********************************************************************************

Option Explicit

Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)

'codi per al registe que s'afegix al cridar des d'atre formulari.
'Obrir en modo Insertar i tornar datos del registre insertat
Public NuevoCodigo As String

'codigo que tiene el campo en el momento que se llama desde otro formulario
'nos situamos en ese valor
Public FechaCarga As String

Private WithEvents frmSoc As frmManSocios
Attribute frmSoc.VB_VarHelpID = -1
Private WithEvents frmVar As frmManVariedad 'Basico2
Attribute frmVar.VB_VarHelpID = -1
Private WithEvents frmB As frmBuscaGrid
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1

Private CadenaConsulta As String
Private CadB As String
Private PrimeraVez As Boolean
Private HaDevueltoDatos As Boolean

Dim ValorAnt As String
Dim SocioAnt As String

Dim tipoF As String
Dim Modo As Byte
Dim LineaAnt As Integer


'----------- MODOS --------------------------------
'   0.-  Formulari net sense cap camp ple
'   1.-  Preparant per a fer la búsqueda
'   2.-  Ja tenim registres i els anem a recorrer
'        ,podem editar-los Edició del camp
'   3.-  Inserció de nou registre
'   4.-  Modificar
'--------------------------------------------------

Private Sub PonerModo(vModo)
Dim B As Boolean
Dim I As Byte

    On Error Resume Next
    
    Modo = vModo
'    PonerIndicador lblIndicador, Modo
    B = (Modo = 4)
    For I = 1 To 1
        txtAux(I).Enabled = Not B Or (Modo = 1)
    Next I
    btnBuscar(2).visible = (B Or Modo = 1)
    
    B = (Modo = 2)
    If B Then
        PonerContRegIndicador lblIndicador, adodc1, CadB
    Else
        PonerIndicador lblIndicador, Modo
    End If
    
    For I = 0 To 8 'els txtAux del grid
        If I <> 0 And I <> 1 And I <> 2 And I <> 3 Then
            txtAux(I).visible = False ' Not b
        Else
            txtAux(I).visible = Not B
        End If
    Next I
    
    btnBuscar(0).visible = False
    btnBuscar(1).visible = False
    
    txtAux2(5).visible = False
    txtAux2(7).visible = False
    
    cmdAceptar.visible = Not B
    cmdCancelar.visible = Not B
    DataGrid1.Enabled = B
    
    'Si es regresar
    If DatosADevolverBusqueda <> "" Then cmdRegresar.visible = B
    
    If Modo = 3 Or Modo = 4 Or Modo = 1 Then I = 4 'Insertar/Modificar o busqueda
    BloquearImgBuscar Me, I
    
    PonerLongCampos
    PonerModoOpcionesMenu 'Activar/Desact botones de menu segun Modo
    PonerOpcionesMenu 'Activar/Desact botones de menu segun permisos del usuario
    
    'Si estamos modo Modificar bloquear clave primaria
'    BloquearTxt txtAux(0), (Modo = 4)
    BloquearTxt txtAux(1), (Modo = 4)
    For I = 5 To 8
        BloquearTxt txtAux(I), (Modo = 4)
    Next I
    
    PonerFocoGrid Me.DataGrid1
    
    If Err.Number <> 0 Then Err.Clear
End Sub

Private Sub PonerModoOpcionesMenu()
Dim B As Boolean
On Error Resume Next

    B = (Modo = 2)
    'Busqueda
    Toolbar1.Buttons(2).Enabled = B
    Me.mnBuscar.Enabled = B
    'Ver Todos
    Toolbar1.Buttons(3).Enabled = B
    Me.mnVerTodos.Enabled = B
    
    'Insertar
    Toolbar1.Buttons(6).Enabled = B 'And Not DeConsulta
    Me.mnNuevo.Enabled = B 'And Not DeConsulta
    
    B = (B And adodc1.Recordset.RecordCount > 0) 'And Not DeConsulta
    'Modificar
    Toolbar1.Buttons(7).Enabled = B
    Me.mnModificar.Enabled = B
    'Eliminar
    Toolbar1.Buttons(8).Enabled = B And (vParamAplic.Cooperativa = 9)
    Me.mnEliminar.Enabled = B And (vParamAplic.Cooperativa = 9)
    
    If Err.Number <> 0 Then Err.Clear
End Sub

Private Sub BotonAnyadir()
    Dim NumF As String
    Dim anc As Single
    Dim I As Integer
    
    AbrirListadoTraza 1
    
End Sub

Private Sub BotonVerTodos()
    CadB = ""
    CargaGrid ""
    PonerModo 2
End Sub

Private Sub BotonBuscar()
Dim I As Integer
    ' ***************** canviar per la clau primaria ********
    CargaGrid "trzlineas_cargas.idpalet = -1"
    '*******************************************************************************
    'Buscar
    For I = 0 To 3
        txtAux(I).Text = ""
    Next I
    txtAux2(5).Text = ""
    txtAux2(7).Text = ""

    'LLamaLineas DataGrid1.Top + 216, 1
    LLamaLineas 754.7638, 1
    PonerFoco txtAux(1)
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
    For I = 0 To 3
        txtAux(I).Text = DataGrid1.Columns(I).Text
    Next I
    
    For I = 5 To 5
        txtAux(I).Text = DataGrid1.Columns(4).Text
    Next I
    
    For I = 6 To 7
        txtAux(I).Text = DataGrid1.Columns(I).Text
    Next I
    
    txtAux2(5).Text = DataGrid1.Columns(4).Text
    txtAux2(7).Text = DataGrid1.Columns(8).Text
    
    txtAux(8).Text = DataGrid1.Columns(9).Text

    LLamaLineas anc, 4
   
    'Como es modificar
    
    '[Monica]29/10/2013: dejamos modificar la linea de volcado
    LineaAnt = txtAux(0).Text
    
    
    PonerFoco txtAux(0)
    Screen.MousePointer = vbDefault
End Sub

Private Sub LLamaLineas(alto As Single, xModo As Byte)
Dim I As Byte

    DeseleccionaGrid Me.DataGrid1
    PonerModo xModo
    
    'Fijamos el ancho
'    txtAux(0).Top = alto - 20
    For I = 0 To 3
        txtAux(I).Top = alto - 50
    Next I
    btnBuscar(2).Top = alto - 50
End Sub

Private Sub BotonEliminar()
Dim SQL As String
Dim temp As Boolean

    On Error GoTo Error2
    
    'Ciertas comprobaciones
    If adodc1.Recordset.EOF Then Exit Sub
    
    If Not SepuedeBorrar Then Exit Sub
    
    'El registre de codi 0 no es pot Modificar ni Eliminar
    ' ### [Monica] 27/10/2006 he quitado la linea de no eliminar el codigo 0
'    If EsCodigoCero(CStr(adodc1.Recordset.Fields(0).Value), FormatoCampo(txtAux(0))) Then Exit Sub
    
    '*************** canviar els noms i el DELETE **********************************
    SQL = "¿Seguro que desea eliminar el volcado?"
    SQL = SQL & vbCrLf & "IdPalet: " & adodc1.Recordset.Fields(1)
    
    If MsgBox(SQL, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        NumRegElim = adodc1.Recordset.AbsolutePosition
        SQL = "Delete from trzlineas_cargas where idpalet=" & adodc1.Recordset!IdPalet
        
        '  LOG de acciones
        Set LOG = New cLOG
        LOG.Insertar 16, vUsu, "Eliminar Volcado Traza: " & vbCrLf & "Id: " & adodc1.Recordset!IdPalet & vbCrLf
        Set LOG = Nothing
        
        
        conn.Execute SQL
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

Private Function SepuedeBorrar() As Boolean
Dim SQL As String
Dim temp As Boolean

    On Error GoTo eSepuedeBorrar

    SepuedeBorrar = False

    SQL = "select * from palets where idpalet = " & adodc1.Recordset.Fields(1)
    If TotalRegistrosConsulta(SQL) <> 0 Then
        MsgBox "Este palet ha sido traspasado a palets confeccionados. Revise.", vbExclamation
        Exit Function
    End If

    SepuedeBorrar = True
    Exit Function
    
eSepuedeBorrar:
    MuestraError Err.Number, "Comprobar para Borrar", Err.Description
End Function

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
        Case 0 'Socios
            Set frmSoc = New frmManSocios
            frmSoc.DatosADevolverBusqueda = "0|1|"
'            frmSoc.CodigoActual = txtAux(5).Text
            frmSoc.Show vbModal
            Set frmSoc = Nothing
            PonerFoco txtAux(5)
        Case 1 'Variedad
            Set frmVar = New frmBasico2
            
            AyudaVariedad frmVar, txtAux(7).Text
            
            Set frmVar = Nothing
            PonerFoco txtAux(7)
        Case 2 ' Fecha
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
            If txtAux(2).Text <> "" Then frmC.NovaData = txtAux(2).Text
            ' ********************************************
        
            frmC.Show vbModal
            Set frmC = Nothing
            ' *** repasar si el camp es txtAux o Text1 ***
            PonerFoco txtAux(2) '<===
            ' ********************************************
        
            
    End Select
    If Modo = 4 Then BLOQUEADesdeFormulario2 Me, adodc1, 1
End Sub

Private Sub cmdAceptar_Click()
Dim I As Long

Dim T1 As String
Dim T2 As String
Dim T3 As String
Dim T4 As String

Dim B As Boolean

    Select Case Modo
        Case 1 'BUSQUEDA
            For I = 5 To 8
                txtAux(I).Tag = ""
            Next I
        
            CadB = ObtenerBusqueda(Me)
            
            txtAux(5).Tag = "Socio|N|N|0|999999|trzlineas_cargas|codsocio|000000|N|"
            txtAux(6).Tag = "Campo|N|N|0|99999999|trzlineas_cargas|codcampo|00000000||"
            txtAux(7).Tag = "Variedad|N|N|0|999999|trzlineas_cargas|codvarie|000000|N|"
            txtAux(8).Tag = "Kilos|N|N|0|999999.99|trzlineas_cargas|numkilos|###,##0.00||"
            If CadB <> "" Then
                CargaGrid CadB
                PonerModo 2
'                lblIndicador.Caption = "BUSQUEDA: " & PonerContRegistros(Me.adodc1)
                PonerFocoGrid Me.DataGrid1
            End If
            
'        Case 3 'INSERTAR
'            If DatosOk Then
'                If InsertarDesdeForm2(Me, 1) Then
'                    CargaGrid
'                    If (DatosADevolverBusqueda <> "") And NuevoCodigo <> "" Then
'                        cmdCancelar_Click
'    '                    If Not adodc1.Recordset.EOF Then adodc1.Recordset.MoveLast
'                        If Not adodc1.Recordset.EOF Then
'                            adodc1.Recordset.Find (adodc1.Recordset.Fields(0).Name & " =" & NuevoCodigo)
'                        End If
'                        cmdRegresar_Click
'                    Else
'                        BotonAnyadir
'                    End If
'                    CadB = ""
'                End If
'            End If
            
        Case 4 'MODIFICAR
            If DatosOK Then
                txtAux(3).Tag = "Hora|FH|N|||trzlineas_cargas|fechahora|yyyy-mm-dd hh:mm:ss||"
                txtAux(3).MaxLength = 19
                txtAux(3).Text = Format(txtAux(2).Text & " " & txtAux(3).Text, "yyyy-mm-dd hh:mm:ss")
                
                '[Monica]29/10/2013: dejamos modificar la linea de volcado
                If CInt(LineaAnt) <> CInt(ComprobarCero(txtAux(0).Text)) Then
                    T1 = txtAux(5).Tag
                    T2 = txtAux(6).Tag
                    T3 = txtAux(7).Tag
                    T4 = txtAux(8).Tag
                    
                    txtAux(5).Tag = ""
                    txtAux(6).Tag = ""
                    txtAux(7).Tag = ""
                    txtAux(8).Tag = ""
                    
                    B = ModificaDesdeFormularioClaves(Me, "linea = " & DBSet(LineaAnt, "N") & " and idpalet = " & DBSet(txtAux(1).Text, "N"))
                
                    txtAux(5).Tag = T1
                    txtAux(6).Tag = T2
                    txtAux(7).Tag = T3
                    txtAux(8).Tag = T4
                
                
                Else
                                    
                    B = ModificaDesdeFormulario(Me)
                    
                End If
                If B Then
                    TerminaBloquear
                    I = adodc1.Recordset.Fields(1).Value
                    
                    txtAux(3).Tag = "Hora|FHH|N|||trzlineas_cargas|fechahora|hh:mm:ss||"
                    txtAux(3).MaxLength = 8
                    
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
                    
                    'si se ha modificado la empresa que estamos conectados
                    'refrescar los datos de la clase
    '                    If Val(vEmpresa.codEmpre) = Val(txtAux(0).Text) Then
    '                       If vEmpresa.LeerDatos(vEmpresa.codEmpre) = False Then
    '                            MsgBox "No se han podido cargar los datos de la empresa.", vbExclamation
    '                            AccionesCerrar
    '                            End
    '                       End If
    '                    End If
                    
                    
                    txtAux(3).Tag = "Hora|FHH|N|||trzlineas_cargas|fechahora|hh:mm:ss||"
                    txtAux(3).MaxLength = 8
                End If
            End If
    End Select
End Sub



Private Sub cmdCancelar_Click()
On Error Resume Next

    Select Case Modo
        Case 1 'BUSQUEDA
            CargaGrid CadB
        Case 3 'INSERTAR
            DataGrid1.AllowAddNew = False
            If Not adodc1.Recordset.EOF Then adodc1.Recordset.MoveFirst
        Case 4 'MODIFICAR
            TerminaBloquear
    End Select
    
    If Not adodc1.Recordset.EOF Then
'        CargaForaGrid
    Else
        LimpiarCampos
    End If
    
    PonerModo 2
'    If CadB <> "" Then
'        lblIndicador.Caption = "BUSQUEDA: " & PonerContRegistros(Me.adodc1)
'    Else
'        lblIndicador.Caption = ""
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

Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Dim I As Byte
    
    If Modo <> 4 Then 'Modificar
'        CargaForaGrid
    Else 'vamos a Insertar
        For I = 0 To txtAux.Count - 1
            txtAux(I).Text = ""
        Next I
    End If
    
'    If (Modo = 2 Or Modo = 0) Then
'        If CadB = "" Then
'            lblIndicador.Caption = PonerContRegistros(Me.adodc1)
'        Else
'            lblIndicador.Caption = "BUSQUEDA: " & PonerContRegistros(Me.adodc1)
'        End If
'    End If
    If Modo = 2 Then PonerContRegIndicador lblIndicador, adodc1, CadB
End Sub


Private Sub DataGrid1_DblClick()
    If cmdRegresar.visible Then cmdRegresar_Click
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
    
    If PrimeraVez Then
        PrimeraVez = False
        If (DatosADevolverBusqueda <> "") And NuevoCodigo <> "" Then
            BotonAnyadir
        Else
            PonerModo 2
'             If Me.CodigoActual <> "" Then
'                SituarData Me.adodc1, "codempre=" & CodigoActual, "", True
'            End If
        End If
    End If
End Sub

Private Sub Form_Load()
Dim I As Integer

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
        .Buttons(10).Image = 11  'Salir
    End With

    '[Monica]09/03/2017: el boton de eliminar solo lo tiene activo natural
    Toolbar1.Buttons(8).visible = (vParamAplic.Cooperativa = 9)
    Me.mnEliminar.visible = (vParamAplic.Cooperativa = 9)
    '## A mano
'    chkVistaPrevia.Value = CheckValueLeer(Name)


    '****************** canviar la consulta *********************************+
    CadenaConsulta = "Select trzlineas_cargas.linea, trzlineas_cargas.idpalet, trzlineas_cargas.fecha, "
    CadenaConsulta = CadenaConsulta & " trzlineas_cargas.fechahora, trzpalets.codsocio, rsocios.nomsocio, "
    CadenaConsulta = CadenaConsulta & " trzpalets.codcampo, trzpalets.codvarie, variedades.nomvarie, trzpalets.numkilos "
    CadenaConsulta = CadenaConsulta & " from trzlineas_cargas, trzpalets, rsocios, variedades "
    CadenaConsulta = CadenaConsulta & " where trzlineas_cargas.fecha = " & DBSet(FechaCarga, "F") & " and "
    CadenaConsulta = CadenaConsulta & " trzlineas_cargas.idpalet = trzpalets.idpalet and "
    CadenaConsulta = CadenaConsulta & " trzpalets.codsocio = rsocios.codsocio and  "
    CadenaConsulta = CadenaConsulta & " trzpalets.codvarie = variedades.codvarie "
    '************************************************************************
    
    CadB = ""
    CargaGrid ""
   
'    If (DatosADevolverBusqueda <> "") And NuevoCodigo <> "" Then
'        BotonAnyadir
'    Else
'        PonerModo 2
'    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
'    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
    DesBloqueoManual ("BORTUR")
    Screen.MousePointer = vbDefault
End Sub

Private Sub frmB_Selecionado(CadenaDevuelta As String)
   txtAux(6).Text = RecuperaValor(CadenaDevuelta, 1)
   txtAux(6).Text = Format(txtAux(6).Text, "00000000")
End Sub

Private Sub frmC_Selec(vFecha As Date)
    ' *** repasar si el camp es txtAux o Text1 ***
    txtAux(2).Text = Format(vFecha, "dd/mm/yyyy") '<===
    ' ********************************************
End Sub

Private Sub frmCli_DatoSeleccionado(CadenaSeleccion As String)
    txtAux(5).Text = RecuperaValor(CadenaSeleccion, 1)
    FormateaCampo txtAux(5)
    txtAux2(5).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmArt_DatoSeleccionado(CadenaSeleccion As String)
    txtAux(7).Text = RecuperaValor(CadenaSeleccion, 1)
    FormateaCampo txtAux(7)
    txtAux2(7).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmTra_DatoSeleccionado(CadenaSeleccion As String)
    txtAux(11).Text = RecuperaValor(CadenaSeleccion, 1) 'codtraba
    FormateaCampo txtAux(11)
    txtAux2(11).Text = RecuperaValor(CadenaSeleccion, 2) 'nomtraba
End Sub

Private Sub frmFPa_DatoSeleccionado(CadenaSeleccion As String)
    txtAux(12).Text = RecuperaValor(CadenaSeleccion, 1) 'cod fpa
    FormateaCampo txtAux(12)
    txtAux2(12).Text = RecuperaValor(CadenaSeleccion, 2) 'nomforpa
    
End Sub

Private Sub mnBuscar_Click()
    BotonBuscar
End Sub

Private Sub mnEliminar_Click()
    BotonEliminar
End Sub

Private Sub mnModificar_Click()
    'Comprobaciones
    '--------------
    If adodc1.Recordset.EOF Then Exit Sub
    If adodc1.Recordset.RecordCount < 1 Then Exit Sub
    
    'El registre de codi 0 no es pot Modificar ni Eliminar
    ' ### [Monica] 27/10/2006
    ' he quitado la linea de no poder eliminar ni modificar el registro 0
    'If EsCodigoCero(CStr(adodc1.Recordset.Fields(0).Value), FormatoCampo(txtAux(0))) Then Exit Sub
    
    
    'Preparar para modificar
    '-----------------------
    If BLOQUEADesdeFormulario2(Me, adodc1, 1) Then BotonModificar
End Sub

Private Sub mnNuevo_Click()
    BotonAnyadir
End Sub

Private Sub mnSalir_Click()
    Screen.MousePointer = vbDefault
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
        Case 10 'Salir
            mnSalir_Click
    End Select
End Sub

Private Sub CargaGrid(Optional vSQL As String)
    Dim SQL As String
    Dim tots As String
    
'    adodc1.ConnectionString = Conn
    If vSQL <> "" Then
        SQL = CadenaConsulta & " and " & vSQL
    Else
        SQL = CadenaConsulta
    End If
    '********************* canviar el ORDER BY *********************++
    SQL = SQL & " ORDER BY trzlineas_cargas.idpalet, trzlineas_cargas.fechahora"
    '**************************************************************++
    
    CargaGridGnral Me.DataGrid1, Me.adodc1, SQL, PrimeraVez
    
    tots = "S|txtAux(0)|T|Linea|800|;S|txtAux(1)|T|Id.Palet|800|;S|txtAux(2)|T|Fecha|1150|;S|btnBuscar(2)|B||195|;S|txtAux(3)|T|Hora|700|;"
    tots = tots & "S|txtAux(5)|T|Socio|800|;S|btnBuscar(0)|B||195|;S|txtAux2(5)|T|Nombre|2700|;"
    tots = tots & "S|txtAux(6)|T|Campo|800|;S|txtAux(7)|T|Variedad|800|;S|btnBuscar(1)|B||195|;"
    tots = tots & "S|txtAux2(7)|T|Denominacion|2580|;S|txtAux(8)|T|Kilos|1100|;"
    
    arregla tots, DataGrid1, Me
    DataGrid1.ScrollBars = dbgAutomatic
      
    If Not adodc1.Recordset.EOF Then
'        CargaForaGrid
    Else
        LimpiarCampos
    End If
    
    DataGrid1.Columns(0).Alignment = dbgRight
'    DataGrid1.Columns(2).Alignment = dbgRight
      
'   'Habilitamos modificar y eliminar
'   Toolbar1.Buttons(7).Enabled = Not adodc1.Recordset.EOF
'   Toolbar1.Buttons(8).Enabled = Not adodc1.Recordset.EOF
'   mnModificar.Enabled = Not adodc1.Recordset.EOF
'   mnEliminar.Enabled = Not adodc1.Recordset.EOF
   
End Sub

Private Sub txtaux_GotFocus(Index As Integer)
    ConseguirFoco txtAux(Index), Modo
End Sub

Private Sub txtAux_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub txtaux_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = teclaBuscar Then
        If Modo = 1 Or Modo = 3 Or Modo = 4 Then
            Select Case Index
                Case 5: 'cliente
                    KeyAscii = 0
                    btnBuscar_Click (0)
                Case 7: 'articulo
                    KeyAscii = 0
                    btnBuscar_Click (1)
                Case 2: 'fecha de albaran
                    KeyAscii = 0
                    btnBuscar_Click (2)
            End Select
        End If
    Else
        KEYpress KeyAscii
    End If
End Sub

Private Sub txtAux_LostFocus(Index As Integer)
Dim cadMen As String
Dim Famia As String

    If Modo = 1 Then Exit Sub 'Busquedas
    If Not PerderFocoGnral(txtAux(Index), Modo) Then Exit Sub
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    
    Select Case Index
        Case 0 'codclave
            PonerFormatoEntero txtAux(Index)
            
        Case 1, 13 'ALBARAN , MATRICULA
            txtAux(Index).Text = UCase(txtAux(Index).Text)
        
        Case 2 'FECHA
            PonerFormatoFecha txtAux(Index)
        
        Case 3 'Hora
            PonerFormatoHora txtAux(Index)
        
        Case 5 'cod cliente
            If PonerFormatoEntero(txtAux(Index)) Then
                txtAux2(5).Text = PonerNombreDeCod(txtAux(Index), "ssocio", "nomsocio", "codsocio", "N")
                If txtAux2(5).Text = "" Then
                    cadMen = "No existe el Cliente: " & txtAux(Index).Text & vbCrLf
                    cadMen = cadMen & "¿Desea crearlo?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmSoc = New frmManSocios
                        frmSoc.DatosADevolverBusqueda = "0|1|"
'                        frmSoc.NuevoCodigo = txtAux(Index).Text
                        txtAux(Index).Text = ""
                        TerminaBloquear
                        frmSoc.Show vbModal
                        Set frmSoc = Nothing
                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, adodc1, 1
                    Else
                        txtAux(Index).Text = ""
                    End If
                    PonerFoco txtAux(Index)
                Else
                End If
            Else
                txtAux2(5).Text = ""
            End If
            
        Case 7 'cod variedad
            If PonerFormatoEntero(txtAux(Index)) Then
                txtAux2(7).Text = PonerNombreDeCod(txtAux(Index), "variedades", "nomvarie", "codvarie", "N")
                If txtAux2(7).Text = "" Then
                    cadMen = "No existe la Variedad: " & txtAux(Index).Text & vbCrLf
                    cadMen = cadMen & "¿Desea crearlo?" & vbCrLf
                    If MsgBox(cadMen, vbQuestion + vbYesNo) = vbYes Then
                        Set frmVar = New frmManVariedad
                        frmVar.DatosADevolverBusqueda = "0|1|"
                        frmVar.NuevoCodigo = txtAux(Index).Text
                        txtAux(Index).Text = ""
                        TerminaBloquear
                        frmVar.Show vbModal
                        Set frmVar = Nothing
                        If Modo = 4 Then BLOQUEADesdeFormulario2 Me, adodc1, 1
                    Else
                        txtAux(Index).Text = ""
                    End If
                    PonerFoco txtAux(Index)
                Else
                End If
            Else
                txtAux2(7).Text = ""
            End If
            
        Case 8 'CANTIDAD
            If txtAux(Index).Text <> "" Then PonerFormatoDecimal txtAux(Index), 2
           
        Case 9 'PRECIO
' ### [Monica] 25/09/2006
' he quitado las dos lineas siguientes y he puesto ponerformatodecimal
'            cadMen = TransformaPuntosComas(txtAux(Index).Text)
'            txtAux(Index).Text = Format(cadMen, "##,##0.000")
            If txtAux(Index).Text <> "" Then PonerFormatoDecimal txtAux(Index), 2
            
        Case 10 'IMPORTE
            If txtAux(Index).Text <> "" Then PonerFormatoDecimal txtAux(Index), 3
            
        
    End Select
End Sub

Private Function DatosOK() As Boolean
Dim B As Boolean
Dim SQL As String
Dim Fpag As String

    B = CompForm(Me)
    If Not B Then Exit Function
    
    If Modo = 3 Then   'Estamos insertando
        'comprobar si ya existe el campo de clave primaria
        If ExisteCP(txtAux(0)) Then B = False
        
    End If
    
    DatosOK = B
End Function

Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
    PonerOpcionesMenuGeneralNew Me
End Sub

Private Sub LimpiarCampos()
Dim I As Integer

    On Error Resume Next

    For I = 11 To 13
        txtAux(I).Text = ""
    Next I
    txtAux2(11).Text = ""
    txtAux2(12).Text = ""
    
    If Err.Number <> 0 Then Err.Clear
End Sub

Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me

End Sub


'' ### [DavidV] 26/04/2006: Activar/desactivar la rueda del ratón.
'Private Sub DataGrid1_GotFocus()
'  WheelHook DataGrid1
'End Sub
'
'Private Sub DataGrid1_Lostfocus()
'  WheelUnHook
'End Sub

Private Sub CalcularSumaPantalla()
Dim Rs As ADODB.Recordset
Dim SQL As String

  If Not adodc1.Recordset.EOF And CadB = "" Then CadB = "codclave > 0"
  If CadB <> "" Then
     SQL = "select sum(cantidad), sum(importel) FROM scaalb "
     SQL = SQL & " WHERE " & CadB
     Set Rs = New ADODB.Recordset ' Crear objeto
     Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText ' abrir cursor
      If Not Rs.EOF Then
        SQL = "Cantidad: " & Format(Rs.Fields(0), "###,##0.000") & vbCrLf
        SQL = SQL & " Importe : " & Format(Rs.Fields(1), "####,##0.00")
        MsgBox "Totales Selección: " & vbCrLf & vbCrLf & SQL, vbInformation
      End If
     Rs.Close
     Set Rs = Nothing
    Else
        MsgBox "Haga primero una selección para ver Totales.", vbInformation
  End If
End Sub


