VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmRepartoAlb 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informes"
   ClientHeight    =   7605
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   6690
   Icon            =   "frmlRepartoAlb.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7605
   ScaleWidth      =   6690
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameEntradasCampo 
      Height          =   7515
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6615
      Begin VB.TextBox txtcodigo 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   8
         Left            =   2700
         MaxLength       =   50
         PasswordChar    =   "*"
         TabIndex        =   31
         Tag             =   "admon"
         Top             =   1890
         Width           =   1545
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   855
         Left            =   570
         TabIndex        =   27
         Top             =   5490
         Width           =   4935
         Begin VB.TextBox txtcodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   3
            Left            =   3630
            MaxLength       =   10
            TabIndex        =   8
            Tag             =   "Nro.Albarán|N|S|||rhisfruta|numalbar|0000000|S|"
            Top             =   330
            Width           =   1095
         End
         Begin VB.TextBox txtcodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   2
            Left            =   1320
            MaxLength       =   10
            TabIndex        =   7
            Tag             =   "Nro.Albarán|N|S|||rhisfruta|numalbar|0000000|S|"
            Top             =   330
            Width           =   1095
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Albarán"
            ForeColor       =   &H00972E0B&
            Height          =   195
            Index           =   5
            Left            =   60
            TabIndex        =   30
            Top             =   90
            Width           =   540
         End
         Begin VB.Label Label2 
            Caption         =   "Desde"
            Height          =   195
            Index           =   4
            Left            =   390
            TabIndex        =   29
            Top             =   360
            Width           =   465
         End
         Begin VB.Label Label2 
            Caption         =   "Hasta"
            Height          =   195
            Index           =   3
            Left            =   2700
            TabIndex        =   28
            Top             =   360
            Width           =   420
         End
      End
      Begin VB.CommandButton Command4 
         Height          =   440
         Left            =   7860
         Picture         =   "frmlRepartoAlb.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   1440
         Visible         =   0   'False
         Width           =   380
      End
      Begin VB.CommandButton Command3 
         Height          =   440
         Left            =   7860
         Picture         =   "frmlRepartoAlb.frx":0316
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   2215
         Visible         =   0   'False
         Width           =   380
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   15
         Left            =   2730
         Locked          =   -1  'True
         TabIndex        =   14
         Text            =   "Text5"
         Top             =   4440
         Width           =   3375
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   14
         Left            =   2730
         Locked          =   -1  'True
         TabIndex        =   13
         Text            =   "Text5"
         Top             =   4080
         Width           =   3375
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   15
         Left            =   1905
         MaxLength       =   6
         TabIndex        =   4
         Top             =   4440
         Width           =   735
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   14
         Left            =   1905
         MaxLength       =   6
         TabIndex        =   3
         Top             =   4080
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   13
         Left            =   2730
         Locked          =   -1  'True
         TabIndex        =   12
         Text            =   "Text5"
         Top             =   3345
         Width           =   3375
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   12
         Left            =   2730
         Locked          =   -1  'True
         TabIndex        =   11
         Text            =   "Text5"
         Top             =   2985
         Width           =   3375
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   13
         Left            =   1905
         MaxLength       =   6
         TabIndex        =   2
         Top             =   3345
         Width           =   750
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   12
         Left            =   1905
         MaxLength       =   6
         TabIndex        =   1
         Top             =   2985
         Width           =   750
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   4080
         TabIndex        =   9
         Top             =   6855
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   5160
         TabIndex        =   10
         Top             =   6855
         Width           =   975
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   7
         Left            =   4245
         MaxLength       =   10
         TabIndex        =   6
         Top             =   5145
         Width           =   1095
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   6
         Left            =   1920
         MaxLength       =   10
         TabIndex        =   5
         Top             =   5130
         Width           =   1095
      End
      Begin MSComctlLib.ProgressBar Pb2 
         Height          =   255
         Left            =   570
         TabIndex        =   36
         Top             =   6360
         Width           =   5595
         _ExtentX        =   9869
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Es conveniente realizar una copia de seguridad previa. "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   360
         TabIndex        =   35
         Top             =   1200
         Width           =   5820
      End
      Begin VB.Label Label6 
         Caption         =   "Password:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   345
         Left            =   1665
         TabIndex        =   34
         Top             =   1890
         Width           =   2235
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Se recomienda realizarlo cuando no haya nadie trabajando "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   300
         Index           =   0
         Left            =   330
         TabIndex        =   33
         Top             =   1500
         Width           =   5955
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Caption         =   "Este proceso crea albaranes en el histórico. "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   435
         TabIndex        =   32
         Top             =   900
         Width           =   5820
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   0
         Left            =   1590
         Picture         =   "frmlRepartoAlb.frx":0620
         ToolTipText     =   "Buscar fecha"
         Top             =   5130
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   1
         Left            =   3930
         Picture         =   "frmlRepartoAlb.frx":06AB
         ToolTipText     =   "Buscar fecha"
         Top             =   5145
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   15
         Left            =   1590
         MouseIcon       =   "frmlRepartoAlb.frx":0736
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar variedad"
         Top             =   4470
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   14
         Left            =   1590
         MouseIcon       =   "frmlRepartoAlb.frx":0888
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar variedad"
         Top             =   4080
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   13
         Left            =   1590
         MouseIcon       =   "frmlRepartoAlb.frx":09DA
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   3375
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   12
         Left            =   1560
         MouseIcon       =   "frmlRepartoAlb.frx":0B2C
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   2970
         Width           =   240
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Socio"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   27
         Left            =   645
         TabIndex        =   26
         Top             =   2790
         Width           =   405
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   26
         Left            =   975
         TabIndex        =   25
         Top             =   4515
         Width           =   420
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   25
         Left            =   975
         TabIndex        =   24
         Top             =   4125
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "Reparto de Albaranes"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   345
         Left            =   450
         TabIndex        =   23
         Top             =   450
         Width           =   5805
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Variedad"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   24
         Left            =   645
         TabIndex        =   22
         Top             =   3885
         Width           =   630
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   23
         Left            =   930
         TabIndex        =   21
         Top             =   3390
         Width           =   420
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   22
         Left            =   930
         TabIndex        =   20
         Top             =   3030
         Width           =   465
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   21
         Left            =   3315
         TabIndex        =   19
         Top             =   5145
         Width           =   420
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   20
         Left            =   975
         TabIndex        =   18
         Top             =   5190
         Width           =   465
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   19
         Left            =   645
         TabIndex        =   17
         Top             =   4950
         Width           =   450
      End
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   6960
      Top             =   5220
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmRepartoAlb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' +-+-+-+-+-+-+-+-+-+-+-+
' +-+- Autor: MONICA +-+-
' +-+-+-+-+-+-+-+-+-+-+-+

'   CREACION DE ALBARANES SEGUN COOPROPIETARIOS

Option Explicit


Public CadTag As String 'Cadena con el Tag del campo que se va a poner en D/H en los listados
                        'Se necesita si el tipo de codigo es texto

Private Conexion As Byte
'1.- Conexión a BD Ariges  2.- Conexión a BD Conta

Private HaDevueltoDatos As Boolean


Private WithEvents frmPob As frmManPueblos 'Pueblos(procedencia)
Attribute frmPob.VB_VarHelpID = -1
Private WithEvents frmSoc As frmManSocios 'Socios
Attribute frmSoc.VB_VarHelpID = -1
Private WithEvents frmVar As frmComVar 'Variedades de comercial
Attribute frmVar.VB_VarHelpID = -1
Private WithEvents frmB As frmBuscaGrid  'Busquedas
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmC As frmCal 'calendario fechas
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmMens As frmMensajes 'Mensajes
Attribute frmMens.VB_VarHelpID = -1


'GENERALES PARA PASARLE A CRYSTAL REPORT
Private cadFormula As String 'Cadena con la FormulaSelection para Crystal Report
Private CadParam As String 'Cadena con los parametros para Crystal Report
Private numParam As Byte 'Numero de parametros que se pasan a Crystal Report
Private cadSelect As String 'Cadena para comprobar si hay datos antes de abrir Informe
Private cadTitulo As String 'Titulo para la ventana frmImprimir
Private cadNombreRPT As String 'Nombre del informe
Private cadSelect1 As String 'Cadena para comprobar si hay datos antes de abrir Informe


Dim indCodigo As Integer 'indice para txtCodigo
Dim indFrame As Single 'nº de frame en el que estamos
 
'Se inicializan para cada Informe (tabla de BD a la que hace referencia
Dim tabla As String
Dim Tabla1 As String
Dim Codigo As String 'Código para FormulaSelection de Crystal Report
Dim TipCod As String
Dim Orden1 As String 'Campo de Ordenacion (por codigo) para Cristal Report
Dim Orden2 As String 'Campo de Ordenacion (por nombre) para Cristal Report
Dim Tipo As String

Dim Indice As Integer

Dim PrimeraVez As Boolean
Dim Contabilizada As Byte
Dim ConSubInforme As Boolean

' Variables de reparto
Dim vKilosBru As Long
Dim vNumcajon As Long
Dim vKilosNet As Single
Dim vImpTrans As Single
Dim vImpAcarr As Single
Dim vImpRecol As Single
Dim vImppenal As Single
Dim vImpEntrada As Single


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 0, cerrar
    If cerrar Then Unload Me

End Sub


Private Sub cmdAceptar_Click()
Dim cOrden As String
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim numOp As Byte

Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String
    
Dim vSQL As String
Dim nTabla As String
Dim vWhere As String
Dim cTabla As String
Dim Sql As String


    InicializarVbles
    
    'Añadir el parametro de Empresa
    CadParam = CadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    '======== FORMULA  ====================================
    'D/H SOCIO
    cDesde = Trim(txtcodigo(12).Text)
    cHasta = Trim(txtcodigo(13).Text)
    nDesde = txtNombre(12).Text
    nHasta = txtNombre(13).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{" & tabla & ".codsocio}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHSocio=""") Then Exit Sub
    End If
    
    
'    'D/H CLASE
'    cDesde = Trim(txtcodigo(20).Text)
'    cHasta = Trim(txtcodigo(21).Text)
'    nDesde = txtNombre(20).Text
'    nHasta = txtNombre(21).Text
'    If Not (cDesde = "" And cHasta = "") Then
'        'Cadena para seleccion Desde y Hasta
'        Codigo = "{variedades.codclase}"
'        TipCod = "N"
'        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHClase=""") Then Exit Sub
'    End If
'
'
'    vSQL = ""
'    If txtcodigo(20).Text <> "" Then vSQL = vSQL & " and variedades.codclase >= " & DBSet(txtcodigo(20).Text, "N")
'    If txtcodigo(21).Text <> "" Then vSQL = vSQL & " and variedades.codclase <= " & DBSet(txtcodigo(21).Text, "N")
    
    
    'D/H VARIEDAD
    cDesde = Trim(txtcodigo(14).Text)
    cHasta = Trim(txtcodigo(15).Text)
    nDesde = txtNombre(14).Text
    nHasta = txtNombre(15).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{" & tabla & ".codvarie}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHVariedad=""") Then Exit Sub
    End If

    If txtcodigo(14).Text <> "" Then vSQL = vSQL & " and variedades.codvarie >= " & DBSet(txtcodigo(14).Text, "N")
    If txtcodigo(15).Text <> "" Then vSQL = vSQL & " and variedades.codvarie <= " & DBSet(txtcodigo(15).Text, "N")

    
    'D/H fecha
    cDesde = Trim(txtcodigo(6).Text)
    cHasta = Trim(txtcodigo(7).Text)
    nDesde = ""
    nHasta = ""
    If Not (cDesde = "" And cHasta = "") Then
                'Cadena para seleccion Desde y Hasta
        Codigo = "{" & tabla & ".fecalbar}"
        TipCod = "F"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHFecha=""") Then Exit Sub
    End If
        
        
    'D/H nro de albaran
    cDesde = Trim(txtcodigo(2).Text)
    cHasta = Trim(txtcodigo(3).Text)
    nDesde = ""
    nHasta = ""
    If Not (cDesde = "" And cHasta = "") Then
                'Cadena para seleccion Desde y Hasta
        Codigo = "{" & tabla & ".numalbar}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHAlbaran=""") Then Exit Sub
    End If
        
    If Not AnyadirAFormula(cadFormula, "{grupopro.codgrupo} not in [5,6]") Then Exit Sub
    If Not AnyadirAFormula(cadSelect, "{grupopro.codgrupo} not in (5,6)") Then Exit Sub
    
    nTabla = "(rhisfruta INNER JOIN variedades ON rhisfruta.codvarie = variedades.codvarie) "
    nTabla = "(" & nTabla & ") INNER JOIN productos ON variedades.codprodu = productos.codprodu "
    nTabla = "(" & nTabla & ") INNER JOIN grupopro ON productos.codgrupo = grupopro.codgrupo "
        
    Set frmMens = New frmMensajes

    frmMens.OpcionMensaje = 16
    frmMens.cadWHERE = vSQL
    frmMens.Show vbModal

    Set frmMens = Nothing
        
    'Comprobar si hay registros a Mostrar antes de abrir el Informe
    If HayRegParaInforme(nTabla, cadSelect) Then
        cTabla = QuitarCaracterACadena(nTabla, "{")
        cTabla = QuitarCaracterACadena(nTabla, "}")
        
        vWhere = ""
        If cadSelect <> "" Then
            vWhere = QuitarCaracterACadena(cadSelect, "{")
            vWhere = QuitarCaracterACadena(cadSelect, "}")
            vWhere = QuitarCaracterACadena(cadSelect, "_1")
        End If
        
        Sql = "numalbar in (select rhisfruta.numalbar from " & cTabla
        If vWhere <> "" Then
            Sql = Sql & " where " & vWhere & ")"
        Else
            Sql = Sql & ")"
        End If
        
        If Not BloqueaRegistro("rhisfruta", Sql) Then
            MsgBox "No se puede realizar el proceso. Hay albaranes bloqueados.", vbExclamation
            Screen.MousePointer = vbDefault
            Exit Sub
        Else
            If RepartoAlbaranes(nTabla, cadSelect) Then
                MsgBox "Proceso realizado correctamente.", vbExclamation
            Else
                MsgBox "El proceso no se ha realizado.", vbExclamation
            End If
            Me.Pb2.visible = False
        End If
    End If
    
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        
        PonerFoco txtcodigo(8)
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
Dim H As Integer, W As Integer
Dim List As Collection

    'Icono del formulario
    Me.Icon = frmPpal.Icon


    PrimeraVez = True
    limpiar Me
    
    
    ConSubInforme = False

    For H = 12 To 15
        Me.imgBuscar(H).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next H
    
    'Ocultar todos los Frames de Formulario
    FrameEntradasCampo.visible = False
    
    '###Descomentar
'    CommitConexion
    
    'LISTADOS DE MANTENIMIENTOS BASICOS
    '---------------------
    FrameEntradaBasculaVisible True, H, W
    indFrame = 1
    
    tabla = "rhisfruta"
    
    Me.Pb2.visible = False
    
    ActivarCLAVE
    
    'Esto se consigue poneinedo el cancel en el opcion k corresponda
'    Me.cmdCancel(indFrame).Cancel = True
    Me.Width = W + 70
    Me.Height = H + 350
End Sub


Private Sub frmC_Selec(vFecha As Date)
    ' *** repasar si el camp es txtAux o Text1 ***
    txtcodigo(Indice).Text = Format(vFecha, "dd/mm/yyyy") '<===
    ' ********************************************
End Sub


Private Sub frmMens_DatoSeleccionado(CadenaSeleccion As String)
Dim Sql As String
Dim Sql2 As String

    If CadenaSeleccion <> "" Then
        Sql = " {variedades.codvarie} in (" & CadenaSeleccion & ")"
        Sql2 = " {variedades.codvarie} in [" & CadenaSeleccion & "]"
    Else
        Sql = " {variedades.codvarie} = -1 "
    End If
    If Not AnyadirAFormula(cadSelect, Sql) Then Exit Sub
    If Not AnyadirAFormula(cadFormula, Sql2) Then Exit Sub

End Sub

Private Sub frmPob_DatoSeleccionado(CadenaSeleccion As String)
    txtcodigo(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1) ' codigo de poblacion
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2) ' nombre
End Sub

Private Sub frmSoc_DatoSeleccionado(CadenaSeleccion As String)
    txtcodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmVar_DatoSeleccionado(CadenaSeleccion As String)
    txtcodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub


Private Sub imgBuscar_Click(Index As Integer)
   Select Case Index
        Case 12, 13  'SOCIOS
            AbrirFrmSocios (Index)
        
        Case 14, 15  'VARIEDADES
            AbrirFrmVariedad (Index)
    
    End Select
    PonerFoco txtcodigo(indCodigo)
End Sub

Private Sub imgFec_Click(Index As Integer)
    Dim esq As Long
    Dim dalt As Long
    Dim menu As Long
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

    menu = Me.Height - Me.ScaleHeight 'ací tinc el heigth del menú i de la toolbar

    frmC.Left = esq + imgFec(Index).Parent.Left + 30
    frmC.Top = dalt + imgFec(Index).Parent.Top + imgFec(Index).Height + menu - 40

    Select Case Index
        Case 0, 1
            Indice = Index + 6
    End Select


    imgFec(0).Tag = Indice '<===
    ' *** repasar si el camp es txtAux o Text1 ***
    If txtcodigo(Indice).Text <> "" Then frmC.NovaData = txtcodigo(Indice).Text
    ' ********************************************

    frmC.Show vbModal
    Set frmC = Nothing
    ' *** repasar si el camp es txtAux o Text1 ***
    PonerFoco txtcodigo(Indice) '<===
    ' ********************************************
End Sub

Private Sub txtCodigo_GotFocus(Index As Integer)
    ConseguirFoco txtcodigo(Index), 3
End Sub

Private Sub txtCodigo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub txtCodigo_KeyPress(Index As Integer, KeyAscii As Integer)
'15/02/2007
'    KEYpress KeyAscii
'ahora
    If KeyAscii = teclaBuscar Then
        Select Case Index
            Case 0: KEYBusqueda KeyAscii, 0 'poblacion desde
            Case 1: KEYBusqueda KeyAscii, 1 'poblacion hasta
            
            Case 6: KEYFecha KeyAscii, 0 'fecha entrada
            Case 7: KEYFecha KeyAscii, 1 'fecha entrada
            
            Case 12: KEYBusqueda KeyAscii, 12 'socio desde
            Case 13: KEYBusqueda KeyAscii, 13 'socio hasta
            
            Case 14: KEYBusqueda KeyAscii, 14 'variedad desde
            Case 15: KEYBusqueda KeyAscii, 15 'variedad hasta
            
            Case 20: KEYBusqueda KeyAscii, 6 'clase desde
            Case 21: KEYBusqueda KeyAscii, 7 'clase hasta
        End Select
    Else
        KEYpress KeyAscii
    End If

End Sub

Private Sub KEYBusqueda(KeyAscii As Integer, Indice As Integer)
    KeyAscii = 0
    imgBuscar_Click (Indice)
End Sub

Private Sub KEYFecha(KeyAscii As Integer, Indice As Integer)
    KeyAscii = 0
    imgFec_Click (Indice)
End Sub


Private Sub txtCodigo_LostFocus(Index As Integer)
Dim cad As String, cadTipo As String 'tipo cliente

    'Quitar espacios en blanco por los lados
    txtcodigo(Index).Text = Trim(txtcodigo(Index).Text)
'    If txtCodigo(Index).Text = "" Then Exit Sub
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    Select Case Index
        Case 0, 1 'POBLACION
            txtNombre(Index).Text = PonerNombreDeCod(txtcodigo(Index), "rpueblos", "despobla", "codpobla", "T")
        
        Case 20, 21 'CLASES
            txtNombre(Index).Text = PonerNombreDeCod(txtcodigo(Index), "clases", "nomclase", "codclase", "N")
            If txtcodigo(Index).Text <> "" Then txtcodigo(Index).Text = Format(txtcodigo(Index).Text, "000")
    
        Case 12, 13  'SOCIOS
            txtNombre(Index).Text = PonerNombreDeCod(txtcodigo(Index), "rsocios", "nomsocio", "codsocio", "N")
            If txtcodigo(Index).Text <> "" Then txtcodigo(Index).Text = Format(txtcodigo(Index).Text, "000000")
            
        Case 6, 7  'FECHAS
            If txtcodigo(Index).Text <> "" Then PonerFormatoFecha txtcodigo(Index)
            
        Case 14, 15 'VARIEDADES
            txtNombre(Index).Text = PonerNombreDeCod(txtcodigo(Index), "variedades", "nomvarie", "codvarie", "N")
            If txtcodigo(Index).Text <> "" Then txtcodigo(Index).Text = Format(txtcodigo(Index).Text, "000000")
        
        Case 2, 3 ' nro de albaranes
            PonerFormatoEntero txtcodigo(Index)
        
        Case 8 ' password de REPARTO DE KILOS DE ALBARANES
            If txtcodigo(Index).Text = "" Then Exit Sub
            If Trim(txtcodigo(Index).Text) <> Trim(txtcodigo(Index).Tag) Then
                MsgBox "    ACCESO DENEGADO    ", vbExclamation
                txtcodigo(Index).Text = ""
                PonerFoco txtcodigo(Index)
            Else
                DesactivarCLAVE
                PonerFoco txtcodigo(12)
            End If

    End Select
End Sub

Private Sub FrameEntradaBasculaVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Frame para el listado de socios por seccion
    Me.FrameEntradasCampo.visible = visible
    If visible = True Then
        Me.FrameEntradasCampo.Top = -90
        Me.FrameEntradasCampo.Left = 0
        Me.FrameEntradasCampo.Height = 7515
        Me.FrameEntradasCampo.Width = 6615
        W = Me.FrameEntradasCampo.Width
        H = Me.FrameEntradasCampo.Height
    End If
End Sub

Private Sub InicializarVbles()
    cadFormula = ""
    cadSelect = ""
    cadSelect1 = ""
    CadParam = ""
    numParam = 0
End Sub

Private Function PonerDesdeHasta(codD As String, codH As String, nomD As String, nomH As String, param As String) As Boolean
'IN: codD,codH --> codigo Desde/Hasta
'    nomD,nomH --> Descripcion Desde/Hasta
'Añade a cadFormula y cadSelect la cadena de seleccion:
'       "(codigo>=codD AND codigo<=codH)"
' y añade a cadParam la cadena para mostrar en la cabecera informe:
'       "codigo: Desde codD-nomd Hasta: codH-nomH"
Dim devuelve As String
Dim devuelve2 As String

    PonerDesdeHasta = False
    devuelve = CadenaDesdeHasta(codD, codH, Codigo, TipCod)
    If devuelve = "Error" Then Exit Function
    If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Function
    If TipCod <> "F" Then 'Fecha
        If Not AnyadirAFormula(cadSelect, devuelve) Then Exit Function
    Else
        devuelve2 = CadenaDesdeHastaBD(codD, codH, Codigo, TipCod)
        If devuelve2 = "Error" Then Exit Function
        If Not AnyadirAFormula(cadSelect, devuelve2) Then Exit Function
    End If
    If devuelve <> "" Then
        If param <> "" Then
            'Parametro Desde/Hasta
            CadParam = CadParam & AnyadirParametroDH(param, codD, codH, nomD, nomH)
            numParam = numParam + 1
        End If
        PonerDesdeHasta = True
    End If
End Function


Private Sub AbrirFrmSocios(Indice As Integer)
    indCodigo = Indice
    Set frmSoc = New frmManSocios
    frmSoc.DatosADevolverBusqueda = "0|1|"
    frmSoc.Show vbModal
    Set frmSoc = Nothing
End Sub

Private Sub AbrirFrmVariedad(Indice As Integer)
    indCodigo = Indice
    Set frmVar = New frmComVar
    frmVar.DatosADevolverBusqueda = "0|1|"
    frmVar.Show vbModal
    Set frmVar = Nothing
End Sub



Private Function ConcatenarCampos(cTabla As String, cWhere As String) As String
Dim Rs As ADODB.Recordset
Dim Sql As String
Dim Sql1 As String

    ConcatenarCampos = ""

    cTabla = QuitarCaracterACadena(cTabla, "{")
    cTabla = QuitarCaracterACadena(cTabla, "}")
    Sql = "Select rcampos.codcampo FROM " & QuitarCaracterACadena(cTabla, "_1")
    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        Sql = Sql & " WHERE " & cWhere
    End If
    
    
    Sql = "select distinct rcampos.codcampo  from " & cTabla & " where " & cWhere
    Set Rs = New ADODB.Recordset
    
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    Sql1 = ""
    While Not Rs.EOF
        Sql1 = Sql1 & DBLet(Rs.Fields(0).Value, "N") & ","
        Rs.MoveNext
    Wend
    Set Rs = Nothing
    'quitamos el ultimo or
    ConcatenarCampos = Mid(Sql1, 1, Len(Sql1) - 1)
    
End Function


Private Function ActualizarRegistros(cTabla As String, cWhere As String) As Boolean
'Actualizar la marca de impreso
Dim Sql As String

    On Error GoTo eActualizarRegistros

    ActualizarRegistros = False
    cTabla = QuitarCaracterACadena(cTabla, "{")
    cTabla = QuitarCaracterACadena(cTabla, "}")
    Sql = "update " & QuitarCaracterACadena(cTabla, "_1") & " set impreso = 1 "
    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        Sql = Sql & " WHERE " & cWhere
    End If
    
    conn.Execute Sql
    
    ActualizarRegistros = True
    Exit Function
    
eActualizarRegistros:
    MuestraError Err.Number, "Actualizando registros", Err.Description
End Function


Private Function NombreCalidad(Var As String, Calid As String) As String
Dim Sql As String

    NombreCalidad = ""

    Sql = "select nomcalab from rcalidad where codvarie = " & DBSet(Var, "N")
    Sql = Sql & " and codcalid = " & DBSet(Calid, "N")
    
    NombreCalidad = DevuelveValor(Sql)
    
End Function

Private Sub ActivarCLAVE()
Dim I As Integer
    
    For I = 12 To 15
        txtcodigo(I).Enabled = False
        imgBuscar(I).Enabled = False
        imgBuscar(I).visible = False
    Next I
    For I = 2 To 3
        txtcodigo(I).Enabled = False
    Next I
    For I = 6 To 7
        txtcodigo(I).Enabled = False
    Next I
    For I = 0 To 1
        imgFec(I).Enabled = False
        imgFec(I).visible = False
    Next I
    txtcodigo(8).Enabled = True
    cmdAceptar.Enabled = False
    CmdCancel.Enabled = True
End Sub

Private Sub DesactivarCLAVE()
Dim I As Integer
    
    For I = 12 To 15
        txtcodigo(I).Enabled = True
        imgBuscar(I).Enabled = True
        imgBuscar(I).visible = True
    Next I
    For I = 2 To 3
        txtcodigo(I).Enabled = True
    Next I
    For I = 6 To 7
        txtcodigo(I).Enabled = True
    Next I
    For I = 0 To 1
        imgFec(I).Enabled = True
        imgFec(I).visible = True
    Next I
    txtcodigo(8).Enabled = True
    cmdAceptar.Enabled = True
    CmdCancel.Enabled = True
End Sub


Private Function RepartoAlbaranes(cTabla As String, cWhere As String) As Boolean
'Comprobar si hay registros a Mostrar antes de abrir el Informe
Dim Sql As String
Dim Sql2 As String
Dim Sql4 As String
Dim Rs As ADODB.Recordset
Dim Rs2 As ADODB.Recordset
Dim rs3 As ADODB.Recordset

Dim numalbar As Long
Dim vTipoMov As CTiposMov

Dim Albaranes As String

Dim tKilosBru As Long
Dim tNumcajon As Long
Dim tKilosNet As Single
Dim tImpTrans As Single
Dim tImpAcarr As Single
Dim tImpRecol As Single
Dim tImppenal As Single
Dim tImpEntrada As Single
Dim CodTipoMov As String
Dim B As Boolean
Dim devuelve As String
Dim Existe As Boolean
Dim Mens As String
Dim NroPropiedad As String
Dim NumReg As Long

    On Error GoTo eRepartoAlbaranes

    RepartoAlbaranes = False
    

    cTabla = QuitarCaracterACadena(cTabla, "{")
    cTabla = QuitarCaracterACadena(cTabla, "}")
    Sql = "Select rhisfruta.* FROM " & QuitarCaracterACadena(cTabla, "_1")
    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        Sql = Sql & " WHERE " & cWhere
    End If
    
    NumReg = TotalRegistrosConsulta(Sql)
    Pb2.visible = True
    Pb2.Max = NumReg
    
    CargarProgres Pb2, CInt(NumReg)
    conn.BeginTrans
    
    CodTipoMov = "ALF" 'albaranes de fruta
    
    B = True
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not Rs.EOF And B
        IncrementarProgres Pb2, 1
        
        If TieneCopropietarios(Rs!codcampo) Then
            Set vTipoMov = New CTiposMov
            If vTipoMov.Leer(CodTipoMov) Then

                tKilosBru = DBLet(Rs!KilosBru, "N")
                tNumcajon = DBLet(Rs!Numcajon, "N")
                tKilosNet = DBLet(Rs!KilosNet, "N")
                tImpTrans = DBLet(Rs!ImpTrans, "N")
                tImpAcarr = DBLet(Rs!impacarr, "N")
                tImpRecol = DBLet(Rs!imprecol, "N")
                tImppenal = DBLet(Rs!ImpPenal, "N")
                tImpEntrada = DBLet(Rs!ImpEntrada, "N")

                Albaranes = ""

                NroPropiedad = DevuelveValor("select codpropiedad from rcopropiedad where codcampo = " & DBSet(Rs!codcampo, "N"))
                
                Sql2 = "select * from rcopropiedad where codpropiedad = " & DBSet(NroPropiedad, "N")
                Sql2 = Sql2 & " order by numlinea "
                
                Set Rs2 = New ADODB.Recordset
                Rs2.Open Sql2, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                
                While Not Rs2.EOF And B
                    If DBLet(Rs2!codcampo, "N") <> DBLet(Rs!codcampo, "N") Then
                        numalbar = vTipoMov.ConseguirContador(CodTipoMov)
                        Do
                            devuelve = DevuelveDesdeBDNew(cAgro, "rhisfruta", "numalbar", "numalbar", CStr(numalbar), "N")
                            If devuelve <> "" Then
                                'Ya existe el contador incrementarlo
                                Existe = True
                                vTipoMov.IncrementarContador (CodTipoMov)
                                numalbar = vTipoMov.ConseguirContador(CodTipoMov)
                            Else
                                Existe = False
                            End If
                        Loop Until Not Existe
                        
                        Sql = "select codvarie, codsocio, codcampo from rcampos where codcampo = " & DBSet(Rs2!codcampo, "N")
                        
                        Set rs3 = New ADODB.Recordset
                        rs3.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                        
                        If Not rs3.EOF Then
                            If DBLet(rs3!codvarie, "N") <> DBLet(Rs!codvarie, "N") Then
                                B = False
                                Mens = "El campo " & Rs2!codcampo & " no es de la misma variedad que el campo " & Rs!codcampo
                            Else
                                vKilosBru = Round2(DBLet(Rs!KilosBru, "N") * DBLet(Rs2!Porcentaje, "N") / 100)
                                vNumcajon = Round2(DBLet(Rs!Numcajon, "N") * DBLet(Rs2!Porcentaje, "N") / 100)
                                vKilosNet = Round2(DBLet(Rs!KilosNet, "N") * DBLet(Rs2!Porcentaje, "N") / 100)
                                vImpTrans = Round2(DBLet(Rs!ImpTrans, "N") * DBLet(Rs2!Porcentaje, "N") / 100, 2)
                                vImpAcarr = Round2(DBLet(Rs!impacarr, "N") * DBLet(Rs2!Porcentaje, "N") / 100, 2)
                                vImpRecol = Round2(DBLet(Rs!imprecol, "N") * DBLet(Rs2!Porcentaje, "N") / 100, 2)
                                vImppenal = Round2(DBLet(Rs!ImpPenal, "N") * DBLet(Rs2!Porcentaje, "N") / 100, 2)
                                vImpEntrada = Round2(DBLet(Rs!ImpEntrada, "N") * DBLet(Rs2!Porcentaje, "N") / 100, 2)
                                
                                tKilosBru = tKilosBru - vKilosBru
                                tNumcajon = tNumcajon - vNumcajon
                                tKilosNet = tKilosNet - vKilosNet
                                tImpTrans = tImpTrans - vImpTrans
                                tImpAcarr = tImpAcarr - vImpAcarr
                                tImpRecol = tImpRecol - vImpRecol
                                tImppenal = tImppenal - vImppenal
                                tImpEntrada = tImpEntrada - vImpEntrada
                                
                                Sql4 = "insert into rhisfruta (numalbar,fecalbar,codvarie,codsocio,codcampo,tipoentr,recolect,"
                                Sql4 = Sql4 & "kilosbru,numcajon,kilosnet,imptrans,impacarr,imprecol,imppenal,impreso,impentrada,"
                                Sql4 = Sql4 & "cobradosn,prestimado,coddeposito,codpobla,transportadopor) values ("
                                Sql4 = Sql4 & DBSet(numalbar, "N") & ","
                                Sql4 = Sql4 & DBSet(Rs!Fecalbar, "F") & ","
                                Sql4 = Sql4 & DBSet(Rs!codvarie, "N") & ","
                                Sql4 = Sql4 & DBSet(rs3!Codsocio, "N") & ","
                                Sql4 = Sql4 & DBSet(Rs2!codcampo, "N") & ","
                                Sql4 = Sql4 & DBSet(Rs!TipoEntr, "N") & ","
                                Sql4 = Sql4 & DBSet(Rs!Recolect, "N") & ","
                                Sql4 = Sql4 & DBSet(vKilosBru, "N") & ","
                                Sql4 = Sql4 & DBSet(vNumcajon, "N") & ","
                                Sql4 = Sql4 & DBSet(vKilosNet, "N") & ","
                                Sql4 = Sql4 & DBSet(vImpTrans, "N", "S") & ","
                                Sql4 = Sql4 & DBSet(vImpAcarr, "N", "S") & ","
                                Sql4 = Sql4 & DBSet(vImpRecol, "N", "S") & ","
                                Sql4 = Sql4 & DBSet(vImppenal, "N", "S") & ","
                                Sql4 = Sql4 & DBSet(Rs!impreso, "N") & ","
                                Sql4 = Sql4 & DBSet(vImpEntrada, "N") & ","
                                Sql4 = Sql4 & DBSet(Rs!cobradosn, "N") & ","
                                Sql4 = Sql4 & DBSet(Rs!PrEstimado, "N", "S") & ","
                                Sql4 = Sql4 & DBSet(Rs!coddeposito, "N", "S") & ","
                                Sql4 = Sql4 & DBSet(Rs!CodPobla, "N", "S") & ","
                                Sql4 = Sql4 & DBSet(Rs!transportadopor, "N") & ")"
                                
                                conn.Execute Sql4
                                
                                Mens = "Reparto de Entradas."
                                If B Then B = RepartoEntradas(DBLet(Rs!numalbar, "N"), numalbar, DBLet(Rs2!Porcentaje, "N"), Mens)
                            
                                Mens = "Reparto de Clasificación."
                                If B Then B = RepartoClasificacion(DBLet(Rs!numalbar, "N"), numalbar, DBLet(Rs2!Porcentaje, "N"), Mens)
                                
                                Mens = "Reparto de Gastos."
                                If B Then B = RepartoGastos(DBLet(Rs!numalbar, "N"), numalbar, DBLet(Rs2!Porcentaje, "N"), Mens)
                                
                                Mens = "Grabar Incidencias."
                                If B Then B = GrabarIncidencias(DBLet(Rs!numalbar, "N"), numalbar, Mens)
                            
                                Albaranes = Albaranes & numalbar & ","
                            End If
                        End If
                        
                        Set rs3 = Nothing
                        
                    End If
                
                    Rs2.MoveNext
                Wend
                
                Set Rs2 = Nothing
                
                If B Then
                    ' ultimo registro la diferencia ( se updatean las tablas del registro de rhisfruta origen )
                    Sql4 = "update rhisfruta set kilosbru = " & DBSet(tKilosBru, "N") & ","
                    Sql4 = Sql4 & "numcajon = " & DBSet(tNumcajon, "N") & ","
                    Sql4 = Sql4 & "kilosnet = " & DBSet(tKilosNet, "N") & ","
                    Sql4 = Sql4 & "imptrans = " & DBSet(tImpTrans, "N") & ","
                    Sql4 = Sql4 & "impacarr = " & DBSet(tImpAcarr, "N") & ","
                    Sql4 = Sql4 & "imprecol = " & DBSet(tImpRecol, "N") & ","
                    Sql4 = Sql4 & "Imppenal = " & DBSet(tImppenal, "N") & ","
                    Sql4 = Sql4 & "Impentrada = " & DBSet(tImpEntrada, "N")
                    Sql4 = Sql4 & " where numalbar = " & DBSet(Rs!numalbar, "N")
                    
                    conn.Execute Sql4
                
                    Albaranes = "(" & Mid(Albaranes, 1, Len(Albaranes) - 1) & ")"
                
                    Mens = "Actualizar Entradas."
                    If B Then B = ActualizarEntradas(Rs!numalbar, Albaranes, Mens)
                
                    Mens = "Actualizar Clasificación."
                    If B Then B = ActualizarClasificacion(Rs!numalbar, Albaranes, Mens)
                    
                    Mens = "Actualizara Gastos."
                    If B Then B = ActualizarGastosAlbaranes(Rs!numalbar, Albaranes, Mens)
                
                
                    vTipoMov.IncrementarContador (CodTipoMov)
                ' fin de ultimo registro
                End If

            End If
        End If
        Rs.MoveNext
    
    Wend
    

    Set Rs = Nothing


eRepartoAlbaranes:
    If Err.Number <> 0 Or Not B Then
        conn.RollbackTrans
    
    Else
        conn.CommitTrans
        RepartoAlbaranes = True
    End If


End Function

Private Function TieneCopropietarios(campo As String) As Boolean
Dim NroCampo As String

    TieneCopropietarios = TotalRegistros("select count(*) from rcopropiedad where codcampo = " & DBSet(campo, "N")) > 0

End Function


Private Function RepartoEntradas(AlbAnt As Long, numalbar As Long, Porcentaje As Currency, Mens As String) As Boolean
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim Sql2 As String

Dim lKilosBru As Long
Dim lNumcajon As Long
Dim lKilosNet As Single
Dim lImpTrans As Single
Dim lImpAcarr As Single
Dim lImpRecol As Single
Dim lImppenal As Single
Dim lImpEntrada As Single
    
Dim tKilosBru As Long
Dim tNumcajon As Long
Dim tKilosNet As Single
Dim tImpTrans As Single
Dim tImpAcarr As Single
Dim tImpRecol As Single
Dim tImppenal As Single
Dim tImpEntrada As Single

Dim NumNota As Long

    On Error GoTo eRepartoEntradas

    RepartoEntradas = False


    tKilosBru = vKilosBru
    tNumcajon = vNumcajon
    tKilosNet = vKilosNet
    tImpTrans = vImpTrans
    tImpAcarr = vImpAcarr
    tImpRecol = vImpRecol
    tImppenal = vImppenal

    Sql = "select * from rhisfruta_entradas where numalbar = " & DBSet(AlbAnt, "N")
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not Rs.EOF
        lKilosBru = Round2(DBLet(Rs!KilosBru, "N") * Porcentaje / 100)
        lNumcajon = Round2(DBLet(Rs!Numcajon, "N") * Porcentaje / 100)
        lKilosNet = Round2(DBLet(Rs!KilosNet, "N") * Porcentaje / 100)
        lImpTrans = Round2(DBLet(Rs!ImpTrans, "N") * Porcentaje / 100, 2)
        lImpAcarr = Round2(DBLet(Rs!impacarr, "N") * Porcentaje / 100, 2)
        lImpRecol = Round2(DBLet(Rs!imprecol, "N") * Porcentaje / 100, 2)
        lImppenal = Round2(DBLet(Rs!ImpPenal, "N") * Porcentaje / 100, 2)
        
        tKilosBru = tKilosBru - lKilosBru
        tNumcajon = tNumcajon - lNumcajon
        tKilosNet = tKilosNet - lKilosNet
        tImpTrans = tImpTrans - lImpTrans
        tImpAcarr = tImpAcarr - lImpAcarr
        tImpRecol = tImpRecol - lImpRecol
        tImppenal = tImppenal - lImppenal
       
        NumNota = Rs!NumNotac
        
        Sql2 = "insert into rhisfruta_entradas (numalbar,numnotac,fechaent,horaentr,kilosbru,numcajon,kilosnet,observac,imptrans,"
        Sql2 = Sql2 & "impacarr,imprecol,imppenal,prestimado,codtrans,codtarif,codcapat) values ("
        Sql2 = Sql2 & DBSet(numalbar, "N") & ","
        Sql2 = Sql2 & DBSet(Rs!NumNotac, "N") & ","
        Sql2 = Sql2 & DBSet(Rs!FechaEnt, "F") & ","
        Sql2 = Sql2 & DBSet(Rs!horaentr, "FH") & ","
        Sql2 = Sql2 & DBSet(lKilosBru, "N") & ","
        Sql2 = Sql2 & DBSet(lNumcajon, "N") & ","
        Sql2 = Sql2 & DBSet(lKilosNet, "N") & ","
        Sql2 = Sql2 & DBSet(Rs!Observac, "T", "S") & ","
        Sql2 = Sql2 & DBSet(lImpTrans, "N", "S") & ","
        Sql2 = Sql2 & DBSet(lImpAcarr, "N", "S") & ","
        Sql2 = Sql2 & DBSet(lImpRecol, "N", "S") & ","
        Sql2 = Sql2 & DBSet(lImppenal, "N", "S") & ","
        Sql2 = Sql2 & DBSet(Rs!PrEstimado, "N", "S") & ","
        Sql2 = Sql2 & DBSet(Rs!codTrans, "T", "S") & ","
        Sql2 = Sql2 & DBSet(Rs!Codtarif, "N", "S") & ","
        Sql2 = Sql2 & DBSet(Rs!codcapat, "N", "S") & ")"
        
        conn.Execute Sql2
    
        Rs.MoveNext
    Wend
    
    Sql2 = "update rhisfruta_entradas set kilosbru = kilosbru + " & DBSet(tKilosBru, "N")
    Sql2 = Sql2 & ", numcajon = numcajon + " & DBSet(tNumcajon, "N")
    Sql2 = Sql2 & ", kilosnet = kilosnet + " & DBSet(tKilosNet, "N")
    Sql2 = Sql2 & ", imptrans = imptrans + " & DBSet(tImpTrans, "N")
    Sql2 = Sql2 & ", impacarr = impacarr + " & DBSet(tImpAcarr, "N")
    Sql2 = Sql2 & ", imprecol = imprecol + " & DBSet(tImpAcarr, "N")
    Sql2 = Sql2 & ", imppenal = imppenal + " & DBSet(tImppenal, "N")
    Sql2 = Sql2 & " where numalbar = " & DBSet(numalbar, "N") & " and numnotac = " & DBSet(NumNota, "N")
    
    conn.Execute Sql2
    
    Set Rs = Nothing
    
    RepartoEntradas = True
    Exit Function
    
    
eRepartoEntradas:
    Mens = Mens & vbCrLf & "Reparto Entradas. " & Err.Description
End Function




Private Function RepartoClasificacion(AlbAnt As Long, numalbar As Long, Porcentaje As Currency, Mens As String) As Boolean
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim Sql2 As String

Dim lKilosNet As Single
Dim tKilosNet As Single
Dim Calid As Long

    On Error GoTo eRepartoClasificacion

    RepartoClasificacion = False


    tKilosNet = vKilosNet
    Calid = 0
    
    Sql = "select * from rhisfruta_clasif where numalbar = " & DBSet(AlbAnt, "N")
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not Rs.EOF
        lKilosNet = Round2(DBLet(Rs!KilosNet, "N") * Porcentaje / 100)
        
        If lKilosNet <> 0 And Calid = 0 Then
            Calid = DBLet(Rs!codcalid, "N")
        End If
        
        tKilosNet = tKilosNet - lKilosNet
        
        Sql2 = "insert into rhisfruta_clasif (numalbar,codvarie,codcalid,kilosnet) values ("
        Sql2 = Sql2 & DBSet(numalbar, "N") & ","
        Sql2 = Sql2 & DBSet(Rs!codvarie, "N") & ","
        Sql2 = Sql2 & DBSet(Rs!codcalid, "N") & ","
        Sql2 = Sql2 & DBSet(lKilosNet, "N", "S") & ")"
        
        conn.Execute Sql2
    
        Rs.MoveNext
    Wend
    
    ' si no está correcto el totalkilonet el redondeo  va sobre la primera calidad con kilos
    Sql2 = "update rhisfruta_clasif set kilosnet = kilosnet + " & DBSet(tKilosNet, "N")
    Sql2 = Sql2 & " where numalbar = " & DBSet(numalbar, "N") & " and codcalid = " & DBSet(Calid, "N")
    
    conn.Execute Sql2
    
    Set Rs = Nothing
    
    RepartoClasificacion = True
    Exit Function
    
eRepartoClasificacion:
    Mens = Mens & vbCrLf & "Reparto Clasificación. " & Err.Description
End Function


Private Function RepartoGastos(AlbAnt As Long, numalbar As Long, Porcentaje As Currency, Mens As String) As Boolean
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim Sql2 As String

Dim lGastos As Single

    On Error GoTo eRepartogastos

    RepartoGastos = False

    Sql = "select * from rhisfruta_gastos where numalbar = " & DBSet(AlbAnt, "N")
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not Rs.EOF
        lGastos = Round2(Rs!Importe * Porcentaje / 100, 2)
        
        Sql2 = "insert into rhisfruta_gastos (numalbar,numlinea,codgasto,importe) values ("
        Sql2 = Sql2 & DBSet(numalbar, "N") & ","
        Sql2 = Sql2 & DBSet(Rs!numlinea, "N") & ","
        Sql2 = Sql2 & DBSet(Rs!Codgasto, "N") & ","
        Sql2 = Sql2 & DBSet(lGastos, "N", "S") & ")"
        
        conn.Execute Sql2
        
        Rs.MoveNext
    Wend
    
    Set Rs = Nothing
    
    RepartoGastos = True
    Exit Function
    
eRepartogastos:
    Mens = Mens & vbCrLf & "Reparto Gastos. " & Err.Description
End Function



Private Function GrabarIncidencias(AlbAnt As Long, numalbar As Long, Mens As String) As Boolean
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim Sql2 As String

Dim lGastos As Single

    On Error GoTo eGrabarIncidencias

    GrabarIncidencias = False

    Sql = "insert into rhisfruta_incidencia (numalbar,numnotac,codincid) "
    Sql = "select " & DBSet(numalbar, "N") & ",numnotac, codincid from rhisfruta_incidencia where numalbar = " & DBSet(AlbAnt, "N")
    conn.Execute Sql
        
    GrabarIncidencias = True
    Exit Function
    
eGrabarIncidencias:
    Mens = Mens & vbCrLf & "Grabar Incidencias. " & Err.Description
End Function



Private Function ActualizarEntradas(AlbAnt As Long, cadAlbaran As String, Mens As String) As Boolean
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim Sql2 As String

    On Error GoTo eActualizarEntradas

    ActualizarEntradas = False

    Sql = "select numnotac, sum(kilosbru) kilbru, sum(numcajon) numcaj, sum(kilosnet) kilnet, sum(imptrans) imptra, "
    Sql = Sql & " sum(impacarr) impaca, sum(imprecol) imprec, sum(imppenal) imppen "
    Sql = Sql & " from rhisfruta_entradas where numalbar in " & cadAlbaran
    Sql = Sql & " group by 1 order by 1"
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not Rs.EOF
        Sql2 = "update rhisfruta_entradas set kilosbru = kilosbru - " & DBSet(Rs!kilbru, "N")
        Sql2 = Sql2 & ", numcajon = numcajon - " & DBSet(Rs!numcaj, "N")
        Sql2 = Sql2 & ", kilosnet = kilosnet - " & DBSet(Rs!kilnet, "N")
        Sql2 = Sql2 & ", imptrans = imptrans - " & DBSet(Rs!imptra, "N")
        Sql2 = Sql2 & ", impacarr = impacarr - " & DBSet(Rs!impaca, "N")
        Sql2 = Sql2 & ", imprecol = imprecol - " & DBSet(Rs!ImpREC, "N")
        Sql2 = Sql2 & ", imppenal = imppenal - " & DBSet(Rs!imppen, "N")
        Sql2 = Sql2 & " where numalbar = " & DBSet(AlbAnt, "N")
        Sql2 = Sql2 & " and numnotac = " & DBSet(Rs!NumNotac, "N")
    
        conn.Execute Sql2
        Rs.MoveNext
    Wend
    
    Set Rs = Nothing
    
    ActualizarEntradas = True
    Exit Function
    
eActualizarEntradas:
    Mens = Mens & vbCrLf & "Actualizar Entradas. " & Err.Description
End Function



Private Function ActualizarClasificacion(AlbAnt As Long, cadAlbaran As String, Mens As String) As Boolean
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim Sql2 As String

Dim lKilosNet As Single
Dim tKilosNet As Single
Dim Calid As Long

    On Error GoTo eActualizarClasificacion

    ActualizarClasificacion = False

    Sql = "select codvarie, codcalid, sum(kilosnet) as kilosnet from rhisfruta_clasif where numalbar in " & cadAlbaran
    Sql = Sql & " group by 1,2 order by 1,2 "
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not Rs.EOF
        Sql2 = "update rhisfruta_clasif set kilosnet = kilosnet - " & DBSet(Rs!KilosNet, "N")
        Sql2 = Sql2 & " where numalbar = " & DBSet(AlbAnt, "N") & " and codvarie = " & DBSet(Rs!codvarie, "N")
        Sql2 = Sql2 & " and codcalid = " & DBSet(Rs!codcalid, "N")
        
        conn.Execute Sql2
        
        Rs.MoveNext
    Wend
    ' si no está correcto el totalkilonet el redondeo  va sobre la primera calidad con kilos
    Set Rs = Nothing
    
    ActualizarClasificacion = True
    Exit Function
    
eActualizarClasificacion:
    Mens = Mens & vbCrLf & "Actualizar Clasificación. " & Err.Description
End Function



Private Function ActualizarGastosAlbaranes(AlbAnt As Long, cadAlbaran As String, Mens As String) As Boolean
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim Sql2 As String

Dim lKilosNet As Single
Dim tKilosNet As Single
Dim Calid As Long

    On Error GoTo eActualizarGastosAlbaranes

    ActualizarGastosAlbaranes = False

    Sql = "select numlinea, sum(importe) as importe from rhisfruta_gastos where numalbar in " & cadAlbaran
    Sql = Sql & " group by 1 order by 1 "
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not Rs.EOF
        Sql2 = "update rhisfruta_gastos set importe = importe - " & DBSet(Rs!Importe, "N")
        Sql2 = Sql2 & " where numalbar = " & DBSet(AlbAnt, "N") & " and numlinea = " & DBSet(Rs!numlinea, "N")
        
        conn.Execute Sql2
        
        Rs.MoveNext
    Wend
    ' si no está correcto el totalkilonet el redondeo  va sobre la primera calidad con kilos
    Set Rs = Nothing
    
    ActualizarGastosAlbaranes = True
    Exit Function
    
eActualizarGastosAlbaranes:
    Mens = Mens & vbCrLf & "Actualizar Gastos Albaranes. " & Err.Description
End Function







