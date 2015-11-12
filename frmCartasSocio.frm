VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmCartasSocio 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cartas / SMS"
   ClientHeight    =   6660
   ClientLeft      =   45
   ClientTop       =   630
   ClientWidth     =   9855
   ClipControls    =   0   'False
   Icon            =   "frmCartasSocio.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6660
   ScaleWidth      =   9855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameRPT 
      BorderStyle     =   0  'None
      Height          =   435
      Left            =   3450
      TabIndex        =   25
      Top             =   6030
      Width           =   3495
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   8
         Left            =   1530
         MaxLength       =   100
         TabIndex        =   8
         Tag             =   "DocumRPT|T|S|||scartas|documrpt||N|"
         Text            =   "Text1"
         Top             =   60
         Width           =   1935
      End
      Begin VB.Image imgAyuda 
         Height          =   240
         Index           =   0
         Left            =   1260
         MousePointer    =   4  'Icon
         Tag             =   "-1"
         ToolTipText     =   "Ayuda"
         Top             =   90
         Width           =   240
      End
      Begin VB.Label Label22 
         Caption         =   "Documento RPT"
         Height          =   195
         Index           =   5
         Left            =   30
         TabIndex        =   26
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Index           =   7
      Left            =   480
      MaxLength       =   150
      MultiLine       =   -1  'True
      TabIndex        =   7
      Tag             =   "Texto SMS|T|S|||scartas|textoSMS||N|"
      Text            =   "frmCartasSocio.frx":000C
      Top             =   5460
      Width           =   8835
   End
   Begin VB.TextBox Text1 
      Height          =   700
      Index           =   5
      Left            =   480
      MaxLength       =   255
      MultiLine       =   -1  'True
      TabIndex        =   5
      Tag             =   "Párrafo 3|T|S|||scartas|parrafo3||N|"
      Text            =   "frmCartasSocio.frx":00A5
      Top             =   3750
      Width           =   8835
   End
   Begin VB.TextBox Text1 
      Height          =   700
      Index           =   3
      Left            =   480
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Tag             =   "Párrafo 1|T|S|||scartas|parrafo1||N|"
      Text            =   "frmCartasSocio.frx":00AB
      Top             =   1860
      Width           =   8835
   End
   Begin VB.TextBox Text1 
      Height          =   700
      Index           =   4
      Left            =   480
      MaxLength       =   255
      MultiLine       =   -1  'True
      TabIndex        =   4
      Tag             =   "Párrafo 2|T|S|||scartas|parrafo2||N|"
      Text            =   "frmCartasSocio.frx":01AD
      Top             =   2805
      Width           =   8835
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   2
      Left            =   2040
      MaxLength       =   80
      TabIndex        =   2
      Tag             =   "Saludos|T|S|||scartas|saludos||N|"
      Text            =   "Text1"
      Top             =   1260
      Width           =   7275
   End
   Begin VB.TextBox Text1 
      Height          =   435
      Index           =   6
      Left            =   480
      MaxLength       =   110
      MultiLine       =   -1  'True
      TabIndex        =   6
      Tag             =   "Despedida|T|S|||scartas|desped||N|"
      Text            =   "frmCartasSocio.frx":01B3
      Top             =   4710
      Width           =   8835
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   1
      Left            =   5040
      MaxLength       =   50
      TabIndex        =   1
      Tag             =   "Descripción|T|S|||scartas|descarta||N|"
      Text            =   "Text1"
      Top             =   750
      Width           =   4275
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   7170
      TabIndex        =   9
      Top             =   6120
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   8325
      TabIndex        =   10
      Top             =   6120
      Width           =   1035
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   8310
      TabIndex        =   11
      Top             =   6120
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Height          =   540
      Index           =   0
      Left            =   480
      TabIndex        =   16
      Top             =   6000
      Width           =   2655
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
         Height          =   240
         Left            =   240
         TabIndex        =   17
         Top             =   180
         Width           =   2115
      End
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   0
      Left            =   2070
      MaxLength       =   5
      TabIndex        =   0
      Tag             =   "Cod. Carta|N|N|0|999|scartas|codcarta|000|S|"
      Text            =   "Text1"
      Top             =   750
      Width           =   630
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   16
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Buscar"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Ver Todos"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Nuevo"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Modificar"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Eliminar"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Primero"
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Anterior"
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Siguiente"
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Último"
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   315
         Left            =   6720
         TabIndex        =   15
         Top             =   0
         Width           =   1215
      End
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   1800
      Top             =   6060
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   582
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
   Begin VB.Image imgZoom 
      Height          =   240
      Index           =   0
      Left            =   1260
      Tag             =   "-1"
      ToolTipText     =   "Zoom párrafo"
      Top             =   1620
      Width           =   240
   End
   Begin VB.Label Label22 
      Caption         =   "Cuerpo SMS"
      Height          =   195
      Index           =   4
      Left            =   480
      TabIndex        =   24
      Top             =   5220
      Width           =   930
   End
   Begin VB.Label Label22 
      Caption         =   "Párrafo 3"
      Height          =   195
      Index           =   3
      Left            =   480
      TabIndex        =   23
      Top             =   3540
      Width           =   795
   End
   Begin VB.Label Label22 
      Caption         =   "Saludos"
      Height          =   195
      Index           =   0
      Left            =   510
      TabIndex        =   22
      Top             =   1290
      Width           =   690
   End
   Begin VB.Label Label22 
      Caption         =   "Despedida"
      Height          =   195
      Index           =   8
      Left            =   480
      TabIndex        =   21
      Top             =   4500
      Width           =   930
   End
   Begin VB.Label Label22 
      Caption         =   "Descripción"
      Height          =   195
      Index           =   6
      Left            =   4050
      TabIndex        =   20
      Top             =   810
      Width           =   930
   End
   Begin VB.Label Label22 
      Caption         =   "Párrafo 2"
      Height          =   195
      Index           =   2
      Left            =   480
      TabIndex        =   19
      Top             =   2595
      Width           =   795
   End
   Begin VB.Label Label22 
      Caption         =   "Párrafo 1"
      Height          =   195
      Index           =   1
      Left            =   480
      TabIndex        =   18
      Top             =   1665
      Width           =   795
   End
   Begin VB.Label Label3 
      Caption         =   "Código Carta / SMS"
      Height          =   255
      Left            =   480
      TabIndex        =   14
      Top             =   780
      Width           =   1515
   End
   Begin VB.Label Label10 
      Caption         =   "Cargando datos ........."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   8220
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.Menu mnOpciones 
      Caption         =   "&Opciones"
      Begin VB.Menu mnBuscar 
         Caption         =   "&Buscar"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnVerTodos 
         Caption         =   "&Ver Todos"
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
Attribute VB_Name = "frmCartasSocio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public Event DatoSeleccionado(CadenaSeleccion As String)
Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
 
Public CodigoActual As String
 
 
 
Private WithEvents frmB As frmBuscaGrid 'Form para busquedas (frmBuscaGrid)
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmZ As frmZoom  'Zoom para campos Text
Attribute frmZ.VB_VarHelpID = -1


Dim NombreTabla As String
Dim Ordenacion As String
Private Modo As Byte
Private ModoAnterior As Byte
Dim kCampo As Integer
Dim indice As Integer
Dim PrimeraVez As Boolean

Dim btnPrimero As Byte
'Variable que indica el número del Boton  PrimerRegistro en la Toolbar1

Dim CadenaConsulta As String

Private HaDevueltoDatos As Boolean


Private Sub chkVistaPrevia_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub cmdAceptar_Click()
On Error GoTo Error1
    
    Screen.MousePointer = vbHourglass
    Select Case Modo
    Case 1 'BUSQUEDA
        HacerBusqueda
    Case 3 'INSERTAR
        If DatosOk Then
            If InsertarDesdeForm(Me) Then PosicionarData
        End If
    Case 4 'MODIFICAR
        If DatosOk Then
             If ModificaDesdeFormulario(Me) Then
                 TerminaBloquear
                 PosicionarData
             End If
         End If
    End Select
    
Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub cmdCancelar_Click()
On Error GoTo ECancelar

    Select Case Modo
        Case 1 'Buscar
            LimpiarCampos
            PonerModo 0
        Case 3 'Insertar
            If ModoAnterior = 0 Then
                LimpiarCampos
                PonerModo 0
            Else
                PonerModo 2
                PonerCampos
            End If
        Case 4  'Modificar
            TerminaBloquear
            PonerModo 2
            PonerCampos
            PonerFoco Text1(0)
    End Select
ECancelar:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Sub


Private Sub cmdRegresar_Click()
Dim cad As String

    If Data1.Recordset.EOF Then
        MsgBox "Ningún registro devuelto.", vbExclamation
        Exit Sub
    End If

    cad = Data1.Recordset.Fields(0) & "|"
    cad = cad & Data1.Recordset.Fields(1) & "|"
    RaiseEvent DatoSeleccionado(cad)
    Unload Me
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
    
    If PrimeraVez Then
        If CodigoActual <> "" Then
            Text1(0).Text = CodigoActual
            HacerBusqueda
        End If
    End If
    
End Sub


Private Sub Form_Load()
Dim I As Integer

'    'Icono del formulario
'    Me.Icon = frmPpal.Icon

    'ICONOS de La toolbar
    btnPrimero = 13 'Posicion del Boton Primero en la toolbar (+ 3 siguientes)
    With Toolbar1
        .ImageList = frmPpal.imgListComun
        .HotImageList = frmPpal.imgListComun_OM
        .DisabledImageList = frmPpal.imgListComun_BN
        'ASignamos botones
        .Buttons(1).Image = 1   'Buscar
        .Buttons(2).Image = 2 'Ver Todos
        .Buttons(5).Image = 3 'Añadir
        .Buttons(6).Image = 4 'Modificar
        .Buttons(7).Image = 5 'Eliminar
        .Buttons(10).Image = 11 'Salir
        .Buttons(btnPrimero).Image = 6 'Primero
        .Buttons(btnPrimero + 1).Image = 7 'Anterior
        .Buttons(btnPrimero + 2).Image = 8 'Siguiente
        .Buttons(btnPrimero + 3).Image = 9 'Ultimo
    End With
    
    LimpiarCampos   'Limpia los campos TextBox
    
    FrameRPT.Enabled = (vUsu.Nivel = 0)
    FrameRPT.visible = (vUsu.Nivel = 0)
    
    
    'IMAGES para zoom
    For I = 0 To Me.imgZoom.Count - 1
        Me.imgZoom(I).Picture = frmPpal.imgListImages16.ListImages(3).Picture
    Next I
    For I = 0 To imgAyuda.Count - 1
        imgAyuda(I).Picture = frmPpal.ImageListB.ListImages(10).Picture
    Next I
        
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
    
    NombreTabla = "scartas" 'Tabla Cartas Oferta
    Ordenacion = " ORDER BY codcarta"
    
'[Monica]19/06/2014: dejamos buscar todas las cartas
'    If CodigoActual <> "" Then
'        CadenaConsulta = "Select * from " & NombreTabla & " WHERE codcarta = " & CodigoActual
'    Else
        CadenaConsulta = "Select * from " & NombreTabla & " WHERE codcarta = -1" 'No recupera datos
'    End If
    Data1.ConnectionString = conn
    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    
    PonerModo 0
    Screen.MousePointer = vbDefault
    
    PrimeraVez = True
    
End Sub


Private Sub frmB_Selecionado(CadenaDevuelta As String)
'Formulario para Busqueda
Dim CadB As String
Dim Aux As String
      
    If CadenaDevuelta <> "" Then
        HaDevueltoDatos = True
        Screen.MousePointer = vbHourglass
        
        'Estamos en Cabecera
        'Recupera todo el registro de Tarifas de Precios
        'Sabemos que campos son los que nos devuelve
        'Creamos una cadena consulta y ponemos los datos
        CadB = ""
        Aux = ValorDevueltoFormGrid(Text1(0), CadenaDevuelta, 1)
        CadB = Aux
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " " & Ordenacion
        PonerCadenaBusqueda
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub frmZ_Actualizar(vCampo As String)
     Text1(indice).Text = vCampo
     cmdAceptar_Click
End Sub

Private Sub imgAyuda_Click(Index As Integer)
Dim vCadena As String
    Select Case Index
        Case 0
           ' "____________________________________________________________"
            vCadena = "Si no ponemos nada en el documento RPT, el report que se ejecutará" & vbCrLf & _
                      "es el de documentos (scryst) codigo 61. " & vbCrLf & vbCrLf & _
                      "En caso contrario se ejecutará el indicado en este campo." & vbCrLf & vbCrLf
            
            vCadena = vCadena & "" & vbCrLf & vbCrLf
                      
    End Select
    MsgBox vCadena, vbInformation, "Descripción de Ayuda"
    
End Sub


Private Sub imgZoom_Click(Index As Integer)
    
    Set frmZ = New frmZoom

    Select Case Index
        Case 0
            indice = 3
            frmZ.pTitulo = "Párrafo"
            frmZ.pValor = Text1(indice).Text
            frmZ.pModo = Modo
            frmZ.Show vbModal
            Set frmZ = Nothing
            PonerFoco Text1(indice)
    End Select
End Sub

Private Sub mnBuscar_Click()
    BotonBuscar
End Sub

Private Sub mnEliminar_Click()
    BotonEliminar
End Sub


Private Sub mnModificar_Click()
    If BLOQUEADesdeFormulario(Me) Then BotonModificar
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

Private Sub Text1_GotFocus(Index As Integer)
    kCampo = Index
    ConseguirFoco Text1(Index), Modo
End Sub


Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
    KEYdown KeyCode
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub Text1_LostFocus(Index As Integer)
  
    If Not PerderFocoGnral(Text1(Index), Modo) Then Exit Sub
      
    With Text1(Index)
        'Código de Carta
        If Index = 0 Then
            If PonerFormatoEntero(Text1(Index)) Then
                If Modo = 3 Then 'Insertar
                    If ExisteCP(Text1(Index)) Then PonerFoco Text1(Index)
                End If
            End If
        End If
    End With
End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1 'Busqueda
            mnBuscar_Click
        Case 2 'Ver Todos
            mnVerTodos_Click
        Case 5 'Nuevo
            mnNuevo_Click
        Case 6  'Modificar
            mnModificar_Click
        Case 7 'Eliminar
            mnEliminar_Click
        Case 10  'Salir
            mnSalir_Click
        Case btnPrimero To btnPrimero + 3 'Flechas de Desplazamiento
            Desplazamiento (Button.Index - btnPrimero)
    End Select
End Sub


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me
End Sub


Private Sub PonerModo(Kmodo As Byte)
Dim b As Boolean
Dim NumReg As Byte

    Modo = Kmodo
    PonerIndicador Me.lblIndicador, Modo
    
    '===========================================
    'Modo 2. Hay datos y estamos visualizandolos
    b = (Kmodo = 2)
    
    'Ponemos visible, si es formulario de busqueda, el boton regresar cuando hay datos
    If DatosADevolverBusqueda <> "" Then
        cmdRegresar.visible = b
        If Modo = 1 Then Me.lblIndicador.Caption = "BUSQUEDA"
    Else
        cmdRegresar.visible = False
    End If
    
    BloquearImgZoom Me, Modo, 0
    
    'Visualizar flechas de desplazamiento en la toolbar si modo=2
    NumReg = 1
    If Not Data1.Recordset.EOF Then
        If Data1.Recordset.RecordCount > 1 Then NumReg = 2 'Solo es para saber q hay + de 1 registro
    End If
    DesplazamientoVisible Me.Toolbar1, btnPrimero, b, NumReg
       
    'Bloquea los campos Text1 sino estamos modificando/Insertando Datos
    'Si estamos en Insertar además limpia los campos Text1
    BloquearText1 Me, Modo
    
         
    '==============================
    b = Modo <> 0 And Modo <> 2
    cmdCancelar.visible = b
    cmdAceptar.visible = b
    
    chkVistaPrevia.Enabled = (Modo <= 2)

    '===============================
    PonerModoOpcionesMenu 'Activa las Opciones de menu según Modo
    PonerOpcionesMenu   'Activar opciones de menu según nivel
                        'de permisos del usuario
End Sub


Private Sub PonerModoOpcionesMenu()
'Activas unas Opciones de Menu y Toolbar según el modo en que estemos
Dim b As Boolean

    b = (Modo = 2)
    'Modificar
    Toolbar1.Buttons(6).Enabled = b
    Me.mnModificar.Enabled = b
    'eliminar                       '[Monica]19/06/2014: dejamos buscar
    Toolbar1.Buttons(7).Enabled = b ' And CodigoActual = ""
    Me.mnEliminar.Enabled = b ' And CodigoActual = ""
    If Modo = 2 And Text1(0).Text <> "" Then
        Toolbar1.Buttons(7).Enabled = Toolbar1.Buttons(7).Enabled And CInt(Text1(0).Text) <> vParamAplic.CartaPOZ
        Me.mnEliminar.Enabled = Me.mnEliminar.Enabled And CInt(Text1(0).Text) <> vParamAplic.CartaPOZ
    End If

    b = (Modo >= 3)
    'Insertar                           '[Monica]19/06/2014: dejamos buscar
    Toolbar1.Buttons(5).Enabled = Not b 'And CodigoActual = ""
    Me.mnNuevo.Enabled = Not b 'And CodigoActual = ""
    'Buscar
    Toolbar1.Buttons(1).Enabled = Not b 'And CodigoActual = ""
    Me.mnBuscar.Enabled = Not b 'And CodigoActual = ""
    'Ver Todos
    Toolbar1.Buttons(2).Enabled = Not b 'And CodigoActual = ""
    Me.mnVerTodos.Enabled = Not b 'And CodigoActual = ""
End Sub


Private Sub LimpiarCampos()
    limpiar Me   'Metodo general: Limpia los controles TextBox
    'Aqui va el especifico de cada form es
    '### a mano
End Sub


Private Sub Desplazamiento(Index As Integer)
'Botones de Desplazamiento de la Toolbar
'Para desplazarse por los registros de control Data
    DesplazamientoData Data1, Index
    PonerCampos
End Sub


Private Sub BotonBuscar()
    If Modo <> 1 Then
        LimpiarCampos
        PonerModo 1
        'Si pasamos el control aqui lo ponemos en amarillo
        PonerFoco Text1(0)
        Text1(0).BackColor = vbYellow
    Else
        HacerBusqueda
        If Data1.Recordset.EOF Then
            Text1(kCampo).Text = ""
            Text1(kCampo).BackColor = vbYellow
            PonerFoco Text1(kCampo)
        End If
    End If
End Sub


Private Sub BotonVerTodos()
Dim cad As String

'Ver todos
    LimpiarCampos
    
    If chkVistaPrevia.Value = 1 Then
        cad = ""
'[Monica]19/06/2014: dejamos buscar
'        If CodigoActual <> "" Then cad = " codcarta = " & DBSet(CodigoActual, "N")
        MandaBusquedaPrevia cad
    Else
'[Monica]19/06/2014: dejamos buscar
'        If CodigoActual <> "" Then
'            CadenaConsulta = "Select * from " & NombreTabla & " where codcarta = " & DBSet(CodigoActual, "N") & Ordenacion
'        Else
            CadenaConsulta = "Select * from " & NombreTabla & Ordenacion
'        End If
        PonerCadenaBusqueda
    End If
End Sub


Private Sub BotonAnyadir()
    LimpiarCampos 'Vacía los TextBox
    ModoAnterior = Modo 'Para el botón Cancelar en Modo Insertar
    PonerModo 3
              
    Text1(0).Text = SugerirCodigoSiguienteStr(NombreTabla, "codcarta")
    FormateaCampo Text1(0)
    PonerFoco Text1(0)
End Sub


Private Sub BotonModificar()
    'Escondemos el navegador y ponemos Modo Modificar
    PonerModo 4
    PonerFoco Text1(1)
End Sub


Private Sub BotonEliminar()
Dim Sql As String

    'Ciertas comprobaciones
    If Data1.Recordset.EOF Then Exit Sub
       
    Sql = Sql & "¿Desea Eliminar la Carta? " & vbCrLf
    Sql = Sql & vbCrLf & "Código : " & Format(Text1(0).Text, "000")
    Sql = Sql & vbCrLf & "Descripción : " & Text1(1).Text
    
    If MsgBox(Sql, vbQuestion + vbYesNo) = vbYes Then
        'Hay que eliminar
        On Error GoTo Error2
        NumRegElim = Data1.Recordset.AbsolutePosition
        If Not Eliminar Then Exit Sub
        If SituarDataTrasEliminar(Data1, NumRegElim) Then
            PonerCampos
        Else
            LimpiarCampos
            PonerModo 0
        End If
    End If
    
Error2:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Eliminar Carta de Oferta.", Err.Description
        Data1.Recordset.CancelUpdate
    End If
End Sub


Private Function Eliminar() As Boolean
On Error GoTo FinEliminar
    
    If Data1.Recordset.EOF Then
        Eliminar = False
        Exit Function
    End If
    
    conn.Execute "Delete  from " & NombreTabla & ObtenerWhereCP
                      
FinEliminar:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Eliminar"
        Eliminar = False
    Else
        Eliminar = True
    End If
End Function


Private Function DatosOk() As Boolean
Dim b As Boolean
On Error Resume Next

    DatosOk = False
    b = CompForm(Me)
    If Not b Then Exit Function
    DatosOk = True
End Function


Private Sub MandaBusquedaPrevia(CadB As String)
'Carga el formulario frmBuscaGrid con los valores correspondientes
Dim cad As String
Dim Tabla As String
Dim Titulo As String

    'Llamamos a al form
    cad = ""
    'Estamos en Modo de Cabeceras
    'Registro de la tabla de cabeceras: scapla
    cad = cad & ParaGrid(Text1(0), 20, "Cod. Carta")
    cad = cad & ParaGrid(Text1(1), 80, "Descripción")
    
    Tabla = NombreTabla
    Titulo = "Cartas de Oferta"
           
    If cad <> "" Then
        Screen.MousePointer = vbHourglass
        Set frmB = New frmBuscaGrid
        frmB.vCampos = cad
        frmB.vTabla = Tabla
        frmB.vSQL = CadB
        HaDevueltoDatos = False
        '###A mano
        frmB.vDevuelve = "0|"
        frmB.vTitulo = Titulo
        frmB.vSelElem = 1
'        frmB.vConexionGrid = conAri 'Conexion a BD Ariges
'        frmB.vBuscaPrevia = chkVistaPrevia
        '#
        frmB.Show vbModal
        Set frmB = Nothing
        'Si ha puesto valores y tenemos que es formulario de busqueda entonces
        'tendremos que cerrar el form lanzando el evento
        If HaDevueltoDatos Then
''            If (Not Data1.Recordset.EOF) And DatosADevolverBusqueda <> "" Then _
''                cmdRegresar_Click
'        Else   'de ha devuelto datos, es decir NO ha devuelto datos
'            If Modo = 5 Then
'                PonerFoco txtAux(0)
'            Else
                PonerFoco Text1(kCampo)
'            End If
        End If
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub HacerBusqueda()
Dim CadB As String

    CadB = ObtenerBusqueda(Me, False)
    
'[Monica]19/06/2014: permito buscar todos
'    If CodigoActual <> "" Then
'        If CadB <> "" Then CadB = CadB & " and "
'        CadB = CadB & " codcarta = " & DBSet(CodigoActual, "N")
'    End If
    
    If chkVistaPrevia = 1 Then
        MandaBusquedaPrevia CadB
    ElseIf CadB <> "" Then
        'Se muestran en el mismo form
        CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " " & Ordenacion
        PonerCadenaBusqueda
    End If
End Sub


Private Sub PonerCadenaBusqueda()
Dim CadMen As String
    Screen.MousePointer = vbHourglass
    On Error GoTo EEPonerBusq

    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    If Data1.Recordset.RecordCount <= 0 Then
        CadMen = "No hay ningún registro en la tabla " & NombreTabla
        If Modo = 1 Then
            MsgBox CadMen & " para ese criterio de Búsqueda.", vbInformation
        Else
            MsgBox CadMen, vbInformation
        End If
        Screen.MousePointer = vbDefault
        PonerModo Modo
        PonerFoco Text1(0)
        Exit Sub
    Else
        PonerModo 2
        Data1.Recordset.MoveFirst
        PonerCampos
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
    
EEPonerBusq:
    MuestraError Err.Number, "PonerCadenaBusqueda"
    PonerModo 0
    Screen.MousePointer = vbDefault
End Sub


Private Sub PonerCampos()
On Error GoTo EPonerCampos

    If Data1.Recordset.EOF Then Exit Sub
    PonerCamposForma Me, Data1
    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
    
    PonerModoOpcionesMenu
EPonerCampos:
    If Err.Number <> 0 Then MuestraError Err.Number, "Poniendo Campos", Err.Description
End Sub


Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
End Sub


Private Sub PosicionarData()
Dim Indicador As String
Dim vWhere As String

    vWhere = Mid(ObtenerWhereCP, 7)
    If SituarData(Data1, vWhere, Indicador) Then
        PonerModo 2
        Indicador = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
        lblIndicador.Caption = Indicador
    Else
        PonerModo 0
    End If
End Sub


Private Function ObtenerWhereCP() As String
On Error Resume Next
    ObtenerWhereCP = " WHERE codcarta= " & Text1(0).Text
End Function

