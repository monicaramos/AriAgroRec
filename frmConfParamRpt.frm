VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmConfParamRpt 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tipos de Documentos"
   ClientHeight    =   5835
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11475
   Icon            =   "frmConfParamRpt.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5835
   ScaleWidth      =   11475
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameDesplazamiento 
      Height          =   705
      Left            =   3270
      TabIndex        =   27
      Top             =   30
      Width           =   2415
      Begin MSComctlLib.Toolbar ToolbarDes 
         Height          =   330
         Left            =   210
         TabIndex        =   28
         Top             =   210
         Width           =   1965
         _ExtentX        =   3466
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   4
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Primero"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Anterior"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Siguiente"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "�ltimo"
               ImageIndex      =   9
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame FrameBotonGnral 
      Height          =   705
      Left            =   60
      TabIndex        =   25
      Top             =   30
      Width           =   3075
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   240
         TabIndex        =   26
         Top             =   180
         Width           =   2655
         _ExtentX        =   4683
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
               Object.Tag             =   "0"
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Object.Visible         =   0   'False
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Object.Visible         =   0   'False
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
   Begin VB.CheckBox chkVistaPrevia 
      Caption         =   "Vista previa"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   8550
      TabIndex        =   24
      Top             =   180
      Width           =   1605
   End
   Begin VB.CheckBox chkImpDirecto 
      Caption         =   "Impresi�n Directa"
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
      Left            =   9030
      TabIndex        =   23
      Tag             =   "Impresi�n Directa|N|N|||scryst|imprimedirecto|||"
      Top             =   2160
      Width           =   2145
   End
   Begin VB.TextBox Text1 
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
      Index           =   10
      Left            =   1440
      MaxLength       =   100
      TabIndex        =   5
      Tag             =   "Fichero Aridoc rpt|T|N|||scryst|aridocrpt|||"
      Text            =   "Text1"
      Top             =   2205
      Width           =   7275
   End
   Begin VB.TextBox Text1 
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
      Index           =   6
      Left            =   300
      MaxLength       =   140
      TabIndex        =   7
      Tag             =   "Linea pie 2|T|S|||scryst|lineapi2|||"
      Text            =   "Text1"
      Top             =   3390
      Width           =   10755
   End
   Begin VB.TextBox Text1 
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
      Index           =   9
      Left            =   300
      MaxLength       =   140
      TabIndex        =   10
      Tag             =   "Linea pie 5|T|S|||scryst|lineapi5|||"
      Text            =   "Text1"
      Top             =   4560
      Width           =   10755
   End
   Begin VB.TextBox Text1 
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
      Index           =   8
      Left            =   300
      MaxLength       =   140
      TabIndex        =   9
      Tag             =   "Linea pie 4|T|S|||scryst|lineapi4|||"
      Text            =   "Text1"
      Top             =   4170
      Width           =   10755
   End
   Begin VB.TextBox Text1 
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
      Index           =   7
      Left            =   300
      MaxLength       =   140
      TabIndex        =   8
      Tag             =   "Linea pie 3|T|S|||scryst|lineapi3|||"
      Text            =   "Text1"
      Top             =   3780
      Width           =   10755
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
      Left            =   10005
      TabIndex        =   12
      Top             =   5280
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Height          =   540
      Left            =   240
      TabIndex        =   19
      Top             =   5160
      Width           =   3000
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
         Left            =   120
         TabIndex        =   20
         Top             =   210
         Width           =   2760
      End
   End
   Begin VB.TextBox Text1 
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
      Index           =   5
      Left            =   300
      MaxLength       =   140
      TabIndex        =   6
      Tag             =   "Linea pie 1|T|S|||scryst|lineapi1|||"
      Text            =   "Text1"
      Top             =   3000
      Width           =   10755
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
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
      Index           =   4
      Left            =   10410
      MaxLength       =   2
      TabIndex        =   3
      Tag             =   "Revisi�n ISO|N|S|0|99|scryst|codigrev|00||"
      Text            =   "Te"
      Top             =   1320
      Width           =   615
   End
   Begin VB.TextBox Text1 
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
      Index           =   3
      Left            =   7050
      MaxLength       =   10
      TabIndex        =   2
      Tag             =   "C�digo ISO|T|S|||scryst|codigiso|||"
      Text            =   "Text1"
      Top             =   1320
      Width           =   1665
   End
   Begin VB.TextBox Text1 
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
      Index           =   2
      Left            =   1455
      MaxLength       =   100
      TabIndex        =   4
      Tag             =   "Fichero rpt|T|N|||scryst|documrpt|||"
      Text            =   "Text1"
      Top             =   1770
      Width           =   7275
   End
   Begin VB.TextBox Text1 
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
      Left            =   1455
      MaxLength       =   30
      TabIndex        =   1
      Tag             =   "Descripci�n|T|N|||scryst|nomcryst|||"
      Text            =   "Text1"
      Top             =   1350
      Width           =   4275
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
      Left            =   8790
      TabIndex        =   11
      Top             =   5280
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
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
      Height          =   375
      Left            =   9990
      TabIndex        =   13
      Top             =   5280
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   1455
      MaxLength       =   4
      TabIndex        =   0
      Tag             =   "C�digo Documento|N|N|||scryst|codcryst||S|"
      Text            =   "Text"
      Top             =   900
      Width           =   915
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   8160
      Top             =   840
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
   Begin MSComctlLib.Toolbar ToolbarAyuda 
      Height          =   330
      Left            =   10680
      TabIndex        =   29
      Top             =   150
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
   Begin VB.Label Label1 
      Caption         =   "Aridoc rpt"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   315
      TabIndex        =   22
      Top             =   2205
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "L�neas para el  Pie del Informe"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   300
      TabIndex        =   21
      Top             =   2715
      Width           =   3135
   End
   Begin VB.Label Label1 
      Caption         =   "Revisi�n ISO"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   9030
      TabIndex        =   18
      Top             =   1365
      Width           =   1305
   End
   Begin VB.Label Label1 
      Caption         =   "C�digo ISO"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   5820
      TabIndex        =   17
      Top             =   1365
      Width           =   1155
   End
   Begin VB.Label Label1 
      Caption         =   "Fichero rpt"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   300
      TabIndex        =   16
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Descripci�n"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   300
      TabIndex        =   15
      Top             =   1350
      Width           =   1125
   End
   Begin VB.Label Label1 
      Caption         =   "C�digo"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   300
      TabIndex        =   14
      Top             =   900
      Width           =   795
   End
   Begin VB.Menu mnOpciones 
      Caption         =   "&Opciones"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu mnNuevo 
         Caption         =   "&Nuevo"
         Shortcut        =   ^N
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
Attribute VB_Name = "frmConfParamRpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ### [Monica] 06/09/2006
' procedimiento nuevo introducido de la gestion

Option Explicit

Private Const IdPrograma = 1004


Private NombreTabla As String  'Nombre de la tabla o de la
Private Ordenacion As String
Private CadenaConsulta As String

Private WithEvents frmB As frmBuscaGrid
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmTDoc As frmBasico2
Attribute frmTDoc.VB_VarHelpID = -1

Dim HaDevueltoDatos  As Boolean

Dim btnPrimero As Byte
'Variable que indica el n�mero del Boton  PrimerRegistro en la Toolbar1

Dim Modo As Byte
'0: Inicial
'2: Visualizacion
'3: A�adir
'4: Modificar

Dim CadB As String



Private Sub cmdAceptar_Click()
Dim vParamRpt As CParamRpt 'Clase Parametros para Reports
Dim cad As String, Indicador As String
Dim actualiza As Boolean

    Select Case Modo
        Case 1  'BUSQUEDA
            HacerBusqueda
        Case 3, 4
            If DatosOK Then
                'Modifica datos en la Tabla: scryst
        '        I = ModificaDesdeFormulario(Me)
                'Actualizar campos de la clase
                    Set vParamRpt = New CParamRpt
                    vParamRpt.Codigo = Text1(0).Text
                    vParamRpt.Descripcion = Text1(1).Text
                    vParamRpt.Documento = Text1(2).Text
                    vParamRpt.AridocRpt = Text1(10).Text
                    vParamRpt.CodigoISO = Text1(3).Text
                    If Trim(Text1(4).Text) <> "" Then
                        vParamRpt.CodigoRevision = CInt(Text1(4).Text)
                    Else
                        vParamRpt.CodigoRevision = -1
                    End If
                    vParamRpt.LineaPie1 = Text1(5).Text
                    vParamRpt.LineaPie2 = Text1(6).Text
                    vParamRpt.LineaPie3 = Text1(7).Text
                    vParamRpt.LineaPie4 = Text1(8).Text
                    vParamRpt.LineaPie5 = Text1(9).Text
                    
                    vParamRpt.ImprimeDirecto = chkImpDirecto.Value
                If Modo = 3 Then 'INSERTAR
                    actualiza = vParamRpt.Insertar
                ElseIf Modo = 4 Then 'MODIFICAR
                    actualiza = vParamRpt.Modificar(Text1(0).Text)
                    TerminaBloquear
                End If
                Set vParamRpt = Nothing
                If actualiza = 0 Then 'Inserta o Modifica
                    cad = "codcryst=" & Text1(0).Text
                    If SituarData(Data1, cad, Indicador) Then
                        PonerModo 2
                        Me.lblIndicador.Caption = Indicador
                    End If
                End If
'                PonerFocoBtn Me.cmdSalir
            End If
    End Select
End Sub


Private Sub cmdCancelar_Click()
    Select Case Modo
        Case 1, 3 'Insertar
            LimpiarCampos
            PonerModo 0
            'PonerFoco Text1(0)
        Case 4 'Modificar
            TerminaBloquear
            If Data1.Recordset.EOF Then
                PonerModo 0
                LimpiarCampos
            Else
                PonerCampos
                PonerModo 2
                Me.lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
            End If
    End Select
End Sub


Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
    PonerCadenaBusqueda
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub Form_Load()
    'Icono del formulario
    Me.Icon = frmPpal.Icon

    ' ICONITOS DE LA BARRA
    With Me.Toolbar1
        btnPrimero = 11
        .ImageList = frmPpal.imgListComun
        .Buttons(5).Image = 1
        .Buttons(6).Image = 2
        .Buttons(1).Image = 3   'Anyadir
        .Buttons(2).Image = 4   'Modificar
        .Buttons(3).Image = 5   'Eliminar
    End With
    
    ' desplazamiento
    With Me.ToolbarDes
        .HotImageList = frmPpal.imgListComun_OM
        .DisabledImageList = frmPpal.imgListComun_BN
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 6
        .Buttons(2).Image = 7
        .Buttons(3).Image = 8
        .Buttons(4).Image = 9
    End With
    
    ' La Ayuda
    With Me.ToolbarAyuda
        .ImageList = frmPpal.imgListComun
        .Buttons(1).Image = 12
    End With
    
    
    
    NombreTabla = "scryst"
    Ordenacion = " ORDER BY codcryst"
    'ASignamos un SQL al DATA1
    Data1.ConnectionString = conn
    CadenaConsulta = "Select * from " & NombreTabla & Ordenacion
    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    
    PonerModo 0
    PonerFoco Text1(0)

End Sub


Private Sub PonerCadenaBusqueda()
On Error GoTo EEPonerBusq
    Screen.MousePointer = vbHourglass

    Data1.RecordSource = CadenaConsulta
    Data1.Refresh
    If Data1.Recordset.RecordCount <= 0 Then
        LimpiarCampos
        MsgBox "No hay ning�n registro en la tabla " & NombreTabla, vbInformation
        Screen.MousePointer = vbDefault
        Me.Toolbar1.Buttons(2).Enabled = False
    Else
        PonerModo 2
        Data1.Recordset.MoveFirst
        PonerCampos
    End If
    'Bloquea los campos Text1 sino estamos modificando/Insertando Datos
    'Si estamos en Insertar adem�s limpia los campos Text1
    BloquearText1 Me, Modo
    Screen.MousePointer = vbDefault
    Exit Sub
EEPonerBusq:
    MuestraError Err.Number, "PonerCadenaBusqueda"
    Screen.MousePointer = vbDefault
End Sub



Private Sub frmTDoc_DatoSeleccionado(CadenaSeleccion As String)
Dim CadB As String
    CadB = "codcryst = " & RecuperaValor(CadenaSeleccion, 1)
    
    'Se muestran en el mismo form
    CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " " & Ordenacion
    PonerCadenaBusqueda
    Screen.MousePointer = vbDefault

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

Private Sub Text1_GotFocus(Index As Integer)
    ConseguirFoco Text1(Index), Modo
End Sub


Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Avanzar/Retroceder los campos con las flechas de desplazamiento del teclado.
    KEYdown KeyCode
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress (KeyAscii)
End Sub


Private Sub Text1_LostFocus(Index As Integer)

    If Not PerderFocoGnral(Text1(Index), Modo) Then Exit Sub


    If Index = 4 Then 'cod. ISO
        If Text1(Index).Text = "" Then Exit Sub
        If Not PonerFormatoEntero(Text1(Index)) Then
            Text1(Index).Text = ""
            PonerFoco Text1(Index)
        End If
    End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 5
            'BUSCAR
            BotonBuscar
        
        Case 6
            'Ver todos
            BotonVerTodos
    
        Case 1  'Anyadir
            mnNuevo_Click
        Case 2  'Modificar
            mnModificar_Click
    End Select
End Sub

Private Sub BotonBuscar()
    'Buscar
    If Modo <> 1 Then
        LimpiarCampos
        PonerModo 1
        '### A mano
        '-------------------------------------------------
        'Si pasamos el control aqui lo ponemos en amarillo
        PonerFoco Text1(0)
        Else
            HacerBusqueda
            If Data1.Recordset.EOF Then
                 '### A mano
                Text1(0).Text = ""
                PonerFoco Text1(0)
            End If
    End If
End Sub

Private Sub BotonAnyadir()
    LimpiarCampos
    PonerModo 3
    
    Text1(0).Text = SugerirCodigoSiguienteStr("scryst", "codcryst")
    PonerFoco Text1(0)
End Sub


Private Sub BotonModificar()
    PonerModo 4
    
    'Bloquear el c�digo que es clave primaria
    BloquearTxt Text1(0), True
    'Si no es root o administradar no Mofificar la descripcion del documento
    'If (vUsu.Nivel <> 0 And vUsu.Nivel <> 1) Then BloquearTxt Text1(1), True
    
    PonerFoco Text1(1)
End Sub

Private Sub BotonVerTodos()
    'Ver todos
    LimpiarCampos
    
    If chkVistaPrevia.Value = 1 Then
        MandaBusquedaPrevia ""
    Else
        CadenaConsulta = "Select * from " & NombreTabla & Ordenacion
        PonerCadenaBusqueda
    End If
End Sub

Private Sub HacerBusqueda()
Dim cad As String
Dim CadB As String
CadB = ObtenerBusqueda(Me)

If chkVistaPrevia = 1 Then
    MandaBusquedaPrevia CadB
    Else
        'Se muestran en el mismo form
        If CadB <> "" Then
            CadenaConsulta = "select * from " & NombreTabla & " WHERE " & CadB & " " & Ordenacion
            PonerCadenaBusqueda
        End If
End If
End Sub

Private Function DatosOK() As Boolean
Dim B As Boolean
    DatosOK = False
    B = CompForm(Me)
    DatosOK = B
End Function


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, Modo, cerrar
    If cerrar Then Unload Me
End Sub


Private Sub PonerBotonCabecera(B As Boolean)
    Me.cmdAceptar.visible = Not B
    Me.cmdCancelar.visible = Not B
'    Me.cmdSalir.visible = B
    If B Then Me.lblIndicador.Caption = ""
End Sub


Private Sub PonerCampos()
On Error GoTo EPonerCampos

    If Data1.Recordset.EOF Then Exit Sub
    PonerCamposForma Me, Data1
    
    '-- Esto permanece para saber donde estamos
    lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
    
EPonerCampos:
    If Err.Number <> 0 Then MuestraError Err.Number, "Poniendo Campos", Err.Description
End Sub

Private Sub DesplazamientoVisible(bol As Boolean)
    FrameDesplazamiento.visible = bol
    FrameDesplazamiento.Enabled = bol
End Sub

Private Sub Desplazamiento(Index As Integer)
Select Case Index
    Case 1
        Data1.Recordset.MoveFirst
    Case 2
        Data1.Recordset.MovePrevious
        If Data1.Recordset.BOF Then Data1.Recordset.MoveFirst
    Case 3
        Data1.Recordset.MoveNext
        If Data1.Recordset.EOF Then Data1.Recordset.MoveLast
    Case 4
        Data1.Recordset.MoveLast
End Select
PonerCampos
lblIndicador.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
End Sub


Private Sub PonerOpcionesMenu()
    PonerOpcionesMenuGeneral Me
End Sub


Private Sub LimpiarCampos()
    limpiar Me   'Metodo general: Limpia los controles TextBox
    lblIndicador.Caption = ""
    chkImpDirecto.Value = 0
    'Aqui va el especifico de cada form es
    '### a mano
End Sub


'----------------------------------------------------------------
'   En PONERMODO se habilitan, o no, los diverso campos del
'   formulario en funcion del modo en k vayamos a trabajar
Private Sub PonerModo(Kmodo As Byte)
Dim B As Boolean
Dim NumReg As Byte
   
    Modo = Kmodo
        
    '----------------------------------------------
    'Modo 2. Hay datos y estamos visualizandolos
   
    PonerIndicador lblIndicador, Modo
    
    'Visualizar flechas de desplazamiento en la toolbar si modo=2
    NumReg = 1
    If Not Data1.Recordset.EOF Then
        If Data1.Recordset.RecordCount > 1 Then NumReg = 2 'Solo es para saber q hay + de 1 registro
    End If
    B = (Kmodo = 2) Or (Kmodo = 0)
    DesplazamientoVisible B And Me.Data1.Recordset.RecordCount > 1 ' Me.Toolbar1, btnPrimero, b, NumReg
    
    
    '------------------------------------------------------
    'Modo insertar o modificar
    B = (Kmodo >= 3) Or Kmodo = 1 '-->Luego not b sera kmodo<3
    PonerBotonCabecera Not B
    If cmdCancelar.visible Then
        cmdCancelar.Cancel = True
    Else
        cmdCancelar.Cancel = False
    End If
   
       
    'Bloquea los campos Text1 sino estamos modificando/Insertando Datos
    'Si estamos en Insertar adem�s limpia los campos Text1
    BloquearText1 Me, Modo
    
    B = Not (Modo = 3 Or Modo = 4 Or Modo = 1)
    
    BloquearChk Me.chkImpDirecto, B
    
    PonerModoOpcionesMenu 'Activar opciones de menu seg�n el Modo
    PonerOpcionesMenu   'Activar opciones de menu seg�n nivel
                        'de permisos del usuario
End Sub


Private Sub PonerModoOpcionesMenu()
Dim B As Boolean
    B = (Modo = 3) Or (Modo = 4)
    Me.Toolbar1.Buttons(1).Enabled = Not B 'Insertar
    Me.mnNuevo.Enabled = Not B
    Me.Toolbar1.Buttons(2).Enabled = (Not B) 'Modificar
    Me.mnModificar.Enabled = Not B
End Sub


Private Sub MandaBusquedaPrevia(CadB As String)

    Set frmTDoc = New frmBasico2
    
    AyudaTiposDocumentos frmTDoc, , CadB
    
    Set frmTDoc = Nothing

End Sub


Private Sub ToolbarAyuda_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1
            LanzaVisorMimeDocumento Me.hWnd, DireccionAyuda & IdPrograma & ".html"
    End Select
End Sub

Private Sub ToolbarDes_ButtonClick(ByVal Button As MSComctlLib.Button)
    Desplazamiento (Button.Index)
End Sub
