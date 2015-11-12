VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmConfParamGral 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Datos de Empresa"
   ClientHeight    =   6375
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   6435
   Icon            =   "frmConfParamGral.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6375
   ScaleWidth      =   6435
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   4830
      Left            =   180
      TabIndex        =   21
      Top             =   765
      Width           =   6045
      _ExtentX        =   10663
      _ExtentY        =   8520
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      WordWrap        =   0   'False
      TabCaption(0)   =   "Datos empresa"
      TabPicture(0)   =   "frmConfParamGral.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "ImgMail(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "imgWeb"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1(10)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label1(9)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label1(8)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label1(7)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label1(6)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label1(5)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label1(4)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label1(3)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label1(2)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label1(1)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label1(11)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "text1(10)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "text1(9)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "text1(8)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "text1(7)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "text1(6)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "text1(5)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "text1(4)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "text1(3)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "text1(2)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "text1(1)"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "text1(13)"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).ControlCount=   24
      TabCaption(1)   =   "Datos Campaña"
      TabPicture(1)   =   "frmConfParamGral.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label2"
      Tab(1).Control(1)=   "imgFec(1)"
      Tab(1).Control(2)=   "Label21"
      Tab(1).Control(3)=   "imgFec(0)"
      Tab(1).Control(4)=   "text1(12)"
      Tab(1).Control(5)=   "text1(11)"
      Tab(1).ControlCount=   6
      Begin VB.TextBox text1 
         Height          =   315
         Index           =   13
         Left            =   1425
         MaxLength       =   40
         TabIndex        =   10
         Tag             =   "Persona Contacto|T|S|||empresas|percontacto|||"
         Top             =   4200
         Width           =   4125
      End
      Begin VB.TextBox text1 
         Enabled         =   0   'False
         Height          =   285
         Index           =   11
         Left            =   -73290
         MaxLength       =   10
         TabIndex        =   11
         Tag             =   "F.Inicio Campaña|F|N|||empresas|fechaini|dd/mm/yyyy||"
         Top             =   765
         Width           =   1200
      End
      Begin VB.TextBox text1 
         Enabled         =   0   'False
         Height          =   285
         Index           =   12
         Left            =   -73290
         MaxLength       =   10
         TabIndex        =   12
         Tag             =   "F.Fin Campaña|F|N|||empresas|fechafin|dd/mm/yyyy||"
         Top             =   1260
         Width           =   1200
      End
      Begin VB.TextBox text1 
         Height          =   315
         Index           =   1
         Left            =   1425
         MaxLength       =   40
         TabIndex        =   0
         Tag             =   "Nombre de la Empresa|T|N|||empresas|nomempre|||"
         Top             =   735
         Width           =   4125
      End
      Begin VB.TextBox text1 
         Height          =   315
         Index           =   2
         Left            =   1410
         MaxLength       =   40
         TabIndex        =   1
         Tag             =   "Domicilio de la Empresa|T|N|||empresas|domempre|||"
         Top             =   1170
         Width           =   4125
      End
      Begin VB.TextBox text1 
         Height          =   315
         Index           =   3
         Left            =   1410
         MaxLength       =   6
         TabIndex        =   2
         Tag             =   "CPostal|T|N|||empresas|codpobla|||"
         Top             =   1605
         Width           =   765
      End
      Begin VB.TextBox text1 
         Height          =   315
         Index           =   4
         Left            =   3210
         MaxLength       =   35
         TabIndex        =   3
         Tag             =   "Población|T|N|||empresas|pobempre|||"
         Top             =   1605
         Width           =   2325
      End
      Begin VB.TextBox text1 
         Height          =   315
         Index           =   5
         Left            =   1410
         MaxLength       =   35
         TabIndex        =   4
         Tag             =   "Provincia|T|N|||empresas|proempre|||"
         Top             =   2040
         Width           =   4110
      End
      Begin VB.TextBox text1 
         Height          =   315
         Index           =   6
         Left            =   1410
         MaxLength       =   9
         TabIndex        =   5
         Tag             =   "C.I.F.|T|N|||empresas|cifempre|||"
         Top             =   2475
         Width           =   1605
      End
      Begin VB.TextBox text1 
         Height          =   315
         Index           =   7
         Left            =   1425
         MaxLength       =   10
         TabIndex        =   6
         Tag             =   "Teléfono|T|S|||empresas|telempre|||"
         Top             =   2910
         Width           =   1725
      End
      Begin VB.TextBox text1 
         Height          =   315
         Index           =   8
         Left            =   3795
         MaxLength       =   10
         TabIndex        =   7
         Tag             =   "Fax|T|S|||empresas|faxempre|||"
         Top             =   2910
         Width           =   1725
      End
      Begin VB.TextBox text1 
         Height          =   315
         Index           =   9
         Left            =   1410
         MaxLength       =   100
         TabIndex        =   8
         Tag             =   "Web|T|S|||empresas|wwwempre|||"
         Top             =   3375
         Width           =   4125
      End
      Begin VB.TextBox text1 
         Height          =   315
         Index           =   10
         Left            =   1410
         MaxLength       =   100
         TabIndex        =   9
         Tag             =   "eMail|T|S|||empresas|maiempre|||"
         Top             =   3780
         Width           =   4125
      End
      Begin VB.Label Label1 
         Caption         =   "Contacto"
         Height          =   255
         Index           =   11
         Left            =   540
         TabIndex        =   34
         Top             =   4260
         Width           =   750
      End
      Begin VB.Image imgFec 
         Enabled         =   0   'False
         Height          =   240
         Index           =   0
         Left            =   -73560
         Picture         =   "frmConfParamGral.frx":0044
         ToolTipText     =   "Buscar fecha"
         Top             =   765
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label Label21 
         Caption         =   "Fecha Inicio"
         Height          =   255
         Left            =   -74640
         TabIndex        =   33
         Top             =   765
         Width           =   975
      End
      Begin VB.Image imgFec 
         Enabled         =   0   'False
         Height          =   240
         Index           =   1
         Left            =   -73560
         Picture         =   "frmConfParamGral.frx":00CF
         ToolTipText     =   "Buscar fecha"
         Top             =   1260
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha Fin"
         Height          =   255
         Left            =   -74640
         TabIndex        =   32
         Top             =   1260
         Width           =   1110
      End
      Begin VB.Label Label1 
         Caption         =   "Nombre"
         Height          =   255
         Index           =   1
         Left            =   525
         TabIndex        =   31
         Top             =   735
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Domicilio"
         Height          =   255
         Index           =   2
         Left            =   525
         TabIndex        =   30
         Top             =   1185
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "CPostal"
         Height          =   255
         Index           =   3
         Left            =   525
         TabIndex        =   29
         Top             =   1605
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Población"
         Height          =   255
         Index           =   4
         Left            =   2400
         TabIndex        =   28
         Top             =   1605
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Provincia"
         Height          =   255
         Index           =   5
         Left            =   525
         TabIndex        =   27
         Top             =   2070
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "C.I.F."
         Height          =   255
         Index           =   6
         Left            =   525
         TabIndex        =   26
         Top             =   2520
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Teléfono"
         Height          =   255
         Index           =   7
         Left            =   525
         TabIndex        =   25
         Top             =   2955
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Fax"
         Height          =   255
         Index           =   8
         Left            =   3450
         TabIndex        =   24
         Top             =   2955
         Width           =   255
      End
      Begin VB.Label Label1 
         Caption         =   "Web"
         Height          =   255
         Index           =   9
         Left            =   525
         TabIndex        =   23
         Top             =   3405
         Width           =   600
      End
      Begin VB.Label Label1 
         Caption         =   "E-mail"
         Height          =   255
         Index           =   10
         Left            =   525
         TabIndex        =   22
         Top             =   3855
         Width           =   510
      End
      Begin VB.Image imgWeb 
         Height          =   255
         Left            =   1140
         Picture         =   "frmConfParamGral.frx":015A
         Stretch         =   -1  'True
         Tag             =   "-1"
         ToolTipText     =   "Abrir web"
         Top             =   3420
         Width           =   255
      End
      Begin VB.Image ImgMail 
         Height          =   240
         Index           =   0
         Left            =   1140
         Tag             =   "-1"
         ToolTipText     =   "Enviar e-mail"
         Top             =   3855
         Width           =   240
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   4755
      TabIndex        =   14
      Top             =   5880
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Height          =   540
      Left            =   720
      TabIndex        =   19
      Top             =   5760
      Width           =   2355
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
         Left            =   240
         TabIndex        =   20
         Top             =   240
         Width           =   1920
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   3600
      TabIndex        =   13
      Top             =   5880
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   4755
      TabIndex        =   16
      Top             =   5880
      Width           =   1035
   End
   Begin VB.TextBox text1 
      Height          =   285
      Index           =   0
      Left            =   3360
      MaxLength       =   15
      TabIndex        =   15
      Tag             =   "Código Parámetros Generales|N|N|||empresas|codempre||S|"
      Text            =   "1"
      Top             =   1200
      Width           =   645
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Width           =   6435
      _ExtentX        =   11351
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Añadir"
            Object.Tag             =   "1"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Modificar"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   4440
      Top             =   1395
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
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
   Begin VB.Label Label1 
      Caption         =   "Código"
      Height          =   255
      Index           =   0
      Left            =   2460
      TabIndex        =   17
      Top             =   1200
      Width           =   495
   End
   Begin VB.Menu mnOpciones 
      Caption         =   "&Opciones"
      Begin VB.Menu mnAñadir 
         Caption         =   "&Añadir"
      End
      Begin VB.Menu mnModificar 
         Caption         =   "&Modificar"
         HelpContextID   =   1
         Shortcut        =   ^M
      End
      Begin VB.Menu mnBarra1 
         Caption         =   "-"
      End
      Begin VB.Menu mnSalir 
         Caption         =   "&Salir"
         Shortcut        =   ^S
      End
   End
End
Attribute VB_Name = "frmConfParamGral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private NombreTabla As String  'Nombre de la tabla o de la
Private Ordenacion As String
Private CadenaConsulta As String

Private VieneDeBuscar As Boolean
'Para cuando devuelve dos poblaciones con el mismo codigo Postal. Si viene de pulsar prismatico
'de busqueda poner el valor de poblacion seleccionado y no volver a recuperar de la Base de Datos

Private WithEvents frmC As frmCal 'calendario fechas
Attribute frmC.VB_VarHelpID = -1

Dim Modo As Byte
'Solo hay Modo=0 Visualizacion y Modo=4 para Modificar datos
Dim Encontrado As Boolean

Private Sub cmdAceptar_Click()
    If Modo = 3 Then
        If DatosOk Then
            'Cambiamos el path
            'CambiaPath True
            If InsertarDesdeForm(Me) Then
                PonerModo 0
'                ActualizaNombreEmpresa
                MsgBox "Debe salir de la aplicacion para que los cambios tengan efecto", vbExclamation
            End If

        End If
    End If
    
    
    If Modo = 4 Then
        If DatosOk Then
            'Modifica datos en la Tabla: sparam
            If Not ModificaDesdeFormulario(Me) Then Exit Sub
            
            'Actualizar campos de la clase
'            vEmpresa.nomempre = text1(1).Text
'            vEmpresa.ModificarDatos
    
            vParam.NombreEmpresa = text1(1).Text
            vParam.DomicilioEmpresa = text1(2).Text
            vParam.CPostal = text1(3).Text
            vParam.Poblacion = text1(4).Text
            vParam.Provincia = text1(5).Text
            vParam.CifEmpresa = text1(6).Text
            vParam.Telefono = text1(7).Text
            vParam.Fax = text1(8).Text
            vParam.WebEmpresa = text1(9).Text
            vParam.MailEmpresa = text1(10).Text
            vParam.FecIniCam = text1(11).Text
            vParam.FecFinCam = text1(12).Text
            vParam.PerContacto = text1(13).Text
            vParam.Modificar
            TerminaBloquear
            
            PonerModo 0
            PonerFocoBtn Me.cmdSalir
        End If
    End If
End Sub

Private Sub cmdCancelar_Click()
    TerminaBloquear
    If Data1.Recordset.EOF Then
        PonerModo 0
    Else
        PonerCampos
        PonerModo 2
    End If
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
    If Modo <> 4 Then PonerCadenaBusqueda 'Modo 4: MOdificar
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Form_Load()
Dim i As Integer

    'Icono del formulario
'    Me.Icon = frmPpal.Icon

    ' ICONITOS DE LA BARRA
    With Me.Toolbar1
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 3   'Anyadir
        .Buttons(2).Image = 4   'Modificar
        .Buttons(4).Image = 11  'Salir
    End With
    
    'carga IMAGES de mail
    For i = 0 To Me.ImgMail.Count - 1
        Me.ImgMail(i).Picture = frmPpal.imgListImages16.ListImages(2).Picture
    Next i
    
    VieneDeBuscar = False
    
    '## A mano
    NombreTabla = "empresas"
    Ordenacion = " ORDER BY codempre"
    'ASignamos un SQL al DATA1
    Data1.ConnectionString = conn
    CadenaConsulta = "Select * from " & NombreTabla & Ordenacion
    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    
    Encontrado = True
    If Data1.Recordset.EOF Then
        'No hay registro de datos de parametros
        'quitar###
        Encontrado = False
    End If
    Me.SSTab1.Tab = 0
    PonerModo 0
'    PonerCadenaBusqueda
End Sub

Private Sub PonerCadenaBusqueda()
On Error GoTo EEPonerBusq
    Screen.MousePointer = vbHourglass
    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    If Data1.Recordset.RecordCount <= 0 Then
        MsgBox "No hay ningún registro en la tabla " & NombreTabla, vbInformation
        Screen.MousePointer = vbDefault
        Exit Sub
    Else
        Data1.Recordset.MoveFirst
        PonerCampos
        'Bloquea los campos Text1 sino estamos modificando/Insertando Datos
        'Si estamos en Insertar además limpia los campos Text1
        BloquearText1 Me, Modo
    End If

    Screen.MousePointer = vbDefault
    Exit Sub
EEPonerBusq:
    MuestraError Err.Number, "PonerCadenaBusqueda"
    Screen.MousePointer = vbDefault
End Sub

Private Sub PonerCampos()
    If Data1.Recordset.EOF Then Exit Sub
    PonerCamposForma Me, Data1
    If Trim(text1(3).Text) = "0" Then text1(3).Text = ""
    If Trim(text1(6).Text) = "0" Then text1(6).Text = ""
End Sub

Private Sub frmC_Selec(vFecha As Date)
 'Fecha
    text1(CByte(imgFec(0).Tag)).Text = Format(vFecha, "dd/MM/yyyy")
End Sub

Private Sub imgFec_Click(Index As Integer)
'FEchas
    Dim esq, dalt As Long
    Dim obj As Object
    
    Set frmC = New frmCal

    esq = imgFec(Index).Left
    dalt = imgFec(Index).Top

    Set obj = imgFec(Index).Container

    While imgFec(Index).Parent.Name <> obj.Name
        esq = esq + obj.Left
        dalt = dalt + obj.Top
        Set obj = obj.Container
    Wend
       
    ' es desplega dalt i cap a la esquerra
    frmC.Left = esq + imgFec(Index).Parent.Left + 30
    frmC.Top = dalt + imgFec(Index).Parent.Top + imgFec(Index).Height + 420 + 30

    ' ***canviar l'index de imgFec pel 1r index de les imagens de buscar data***
    imgFec(0).Tag = Index + 11 'independentment de les dates que tinga, sempre pose l'index en la 27
    If text1(Index + 11).Text <> "" Then frmC.NovaData = text1(Index + 11).Text

    frmC.Show vbModal
    Set frmC = Nothing
    PonerFoco text1(CByte(imgFec(0).Tag) + 11)
    ' ***************************
End Sub

Private Sub imgMail_Click(Index As Integer)
'Abrir Outlook para enviar e-mail
Dim dirMail As String

'    If Modo = 0 Then Exit Sub
    Screen.MousePointer = vbHourglass
    
    Select Case Index
        Case 0: dirMail = text1(10).Text
    End Select

    If LanzaMailGnral(dirMail) Then espera 2
    Screen.MousePointer = vbDefault
End Sub

Private Sub imgWeb_Click()
'Abrimos el explorador de windows con la pagina Web del cliente

'    If Modo = 0 Then Exit Sub
    Screen.MousePointer = vbHourglass

    If LanzaHomeGnral(text1(9).Text) Then espera 2
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnModificar_Click()
    If BLOQUEADesdeFormulario(Me) Then BotonModificar
End Sub

Private Sub mnSalir_Click()
    Unload Me
End Sub

Private Sub Text1_GotFocus(Index As Integer)
    ConseguirFoco text1(Index), 3
End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
    KEYdown KeyCode
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = teclaBuscar Then
        Select Case Index
            Case 11: KEYFecha KeyAscii, 0 'fecha desde
            Case 12: KEYFecha KeyAscii, 1 'fecha hasta
        End Select
    Else
        KEYpress KeyAscii
    End If
End Sub

Private Sub KEYFecha(KeyAscii As Integer, indice As Integer)
    KeyAscii = 0
    imgFec_Click (indice)
End Sub

Private Sub Text1_LostFocus(Index As Integer)
    If Not PerderFocoGnral(text1(Index), Modo) Then Exit Sub
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub
    
    ' ***************** configurar els LostFocus dels camps de la capçalera *****************

    
    'Si queremos hacer algo ..
    Select Case Index
        Case 11, 12 'Fecha inicio de campaña y fecha fin de campaña
            PonerFormatoFecha text1(Index)
    End Select
    
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1 ' añadir
            mnAñadir_Click
        Case 2  'Modificar
            mnModificar_Click
        Case 4 'Salir
            mnSalir_Click
    End Select
End Sub

Private Sub BotonModificar()
    'Añadiremos el boton de aceptar y demas objetos para insertar
'    Me.lblIndicador.Caption = "MODIFICAR"
    PonerModo 4
    'Me.imgBuscar.Enabled = True
    PonerFoco text1(1)
End Sub

Private Function DatosOk() As Boolean
Dim b As Boolean
    DatosOk = False
    b = CompForm(Me)
    DatosOk = b
End Function
'
'Private Sub KEYpress(KeyAscii As Integer)
'Dim cerrar As Boolean
'
'    KEYpressGnral KeyAscii, Modo, cerrar
'    If cerrar Then Unload Me
'End Sub

Private Sub PonerBotonCabecera(b As Boolean)
    Me.cmdAceptar.visible = Not b
    Me.cmdCancelar.visible = Not b
    Me.cmdSalir.visible = b
    If b Then Me.lblIndicador.Caption = ""
End Sub

Private Sub PonerModo(vModo As Byte)
Dim b As Boolean

    Modo = vModo
    b = (Modo = 0)
    PonerIndicador Me.lblIndicador, Modo
'    If b Then Me.lblIndicador.Caption = ""
    
' ### [Monica] 13/11/2006
    b = (Modo >= 3) '-->Luego not b sera kmodo<3
    If cmdCancelar.visible Then
        cmdCancelar.Cancel = True
    Else
        cmdCancelar.Cancel = False
    End If
    PonerBotonCabecera Not b

    
    'Bloquea los campos Text1 sino estamos modificando/Insertando Datos
    'Si estamos en Insertar además limpia los campos Text1
    BloquearText1 Me, Modo
    
' ### [Monica] 13/11/2006
    'Poner Botones Aceptar/Cancelar si estamos Modificando datos
'    PonerBotonCabecera b
    
    'Solo si es root o administrador puede modificar el registro
    'cmdAceptar.Enabled = (vUsu.Nivel <= 1)
    
    'Modificar
    Toolbar1.Buttons(1).Enabled = b
    Me.mnModificar.Enabled = b

' ### [Monica] 13/11/2006
    PonerModoOpcionesMenu 'Activar opciones de menu según el Modo


    PonerOpcionesMenu   'Activar opciones de menu según nivel
                        'de permisos del usuario
End Sub

Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
End Sub



Private Sub KEYpress(KeyAscii As Integer)
    If KeyAscii = 13 Then 'ENTER
        KeyAscii = 0
        SendKeys "{tab}"
    ElseIf KeyAscii = 27 Then 'ESC
        If (Modo = 0 Or Modo = 2) Then Unload Me
    End If
End Sub

' ### [Monica] 13/11/2006
' añadida la opcion de añadir cuando no hay registro en la tabla

Private Sub BotonAnyadir()
    'LimpiarCampos
    PonerModo 3
    text1(0).Text = 1
    PonerFoco text1(1)
End Sub

Private Sub PonerModoOpcionesMenu()
Dim b As Boolean
    b = (Modo = 3) Or (Modo = 4)
    Me.Toolbar1.Buttons(1).Enabled = Not Encontrado And Not b  'Añadir
    Me.Toolbar1.Buttons(2).Enabled = Encontrado And Not b 'Modificar
    Me.mnAñadir.Enabled = Not Encontrado And Not b
    Me.mnModificar.Enabled = Encontrado And Not b
'    Me.Toolbar1.Buttons(2).Enabled = (Not b) 'Modificar
End Sub

Private Sub mnAñadir_Click()
    BotonAnyadir
End Sub

