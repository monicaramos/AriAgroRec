VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmAlmzAsigPrecios 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informes"
   ClientHeight    =   6405
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   6630
   Icon            =   "frmAlmzAsigPrecios.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6405
   ScaleWidth      =   6630
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameEntradasCampo 
      Height          =   6330
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6615
      Begin VB.Frame Frame4 
         Caption         =   "Tipo Operación"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   825
         Left            =   330
         TabIndex        =   27
         Top             =   3075
         Width           =   5925
         Begin VB.OptionButton Option3 
            Caption         =   "Asignación de Precio"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   0
            Left            =   270
            TabIndex        =   29
            Top             =   360
            Width           =   2520
         End
         Begin VB.OptionButton Option3 
            Caption         =   "Aumento Precio"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   1
            Left            =   3045
            TabIndex        =   28
            Top             =   405
            Width           =   2160
         End
      End
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Height          =   1305
         Left            =   2850
         TabIndex        =   21
         Top             =   4125
         Width           =   3525
         Begin VB.Frame Frame2 
            Caption         =   "Tipo"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00972E0B&
            Height          =   870
            Left            =   45
            TabIndex        =   24
            Top             =   450
            Width           =   3405
            Begin VB.OptionButton Option1 
               Caption         =   "Porcentaje"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   1
               Left            =   1845
               TabIndex        =   26
               Top             =   405
               Width           =   1425
            End
            Begin VB.OptionButton Option1 
               Caption         =   "Suma Directa"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   0
               Left            =   135
               TabIndex        =   25
               Top             =   405
               Width           =   1725
            End
         End
         Begin VB.TextBox txtcodigo 
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
            Index           =   0
            Left            =   1680
            MaxLength       =   6
            TabIndex        =   22
            Tag             =   "Porcentaje|N|S|||rhisfruta|numalbar|##0.00|S|"
            Top             =   0
            Width           =   1230
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Aumento Precio"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00972E0B&
            Height          =   240
            Index           =   0
            Left            =   60
            TabIndex        =   23
            Top             =   30
            Width           =   1545
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   1455
         Left            =   450
         TabIndex        =   19
         Top             =   3975
         Width           =   2415
         Begin VB.TextBox txtcodigo 
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
            Index           =   2
            Left            =   1260
            MaxLength       =   10
            TabIndex        =   5
            Tag             =   "Porcentaje|N|S|||rhisfruta|prliquidalmz|##,##0.0000|S|"
            Top             =   150
            Width           =   1095
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Precio"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00972E0B&
            Height          =   240
            Index           =   5
            Left            =   60
            TabIndex        =   20
            Top             =   150
            Width           =   600
         End
      End
      Begin VB.CommandButton Command4 
         Height          =   440
         Left            =   7860
         Picture         =   "frmAlmzAsigPrecios.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   1440
         Visible         =   0   'False
         Width           =   380
      End
      Begin VB.CommandButton Command3 
         Height          =   440
         Left            =   7860
         Picture         =   "frmAlmzAsigPrecios.frx":0316
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   2215
         Visible         =   0   'False
         Width           =   380
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
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
         Height          =   360
         Index           =   15
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   9
         Text            =   "Text5"
         Top             =   2490
         Width           =   3555
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
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
         Height          =   360
         Index           =   14
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   8
         Text            =   "Text5"
         Top             =   2085
         Width           =   3555
      End
      Begin VB.TextBox txtcodigo 
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
         Index           =   15
         Left            =   1755
         MaxLength       =   6
         TabIndex        =   4
         Top             =   2490
         Width           =   1005
      End
      Begin VB.TextBox txtcodigo 
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
         Index           =   14
         Left            =   1755
         MaxLength       =   6
         TabIndex        =   3
         Top             =   2085
         Width           =   1005
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
         Left            =   4140
         TabIndex        =   6
         Top             =   5700
         Width           =   1035
      End
      Begin VB.CommandButton cmdCancel 
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
         Index           =   1
         Left            =   5265
         TabIndex        =   7
         Top             =   5700
         Width           =   1035
      End
      Begin VB.TextBox txtcodigo 
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
         Index           =   7
         Left            =   4425
         MaxLength       =   10
         TabIndex        =   2
         Top             =   1365
         Width           =   1350
      End
      Begin VB.TextBox txtcodigo 
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
         Index           =   6
         Left            =   1740
         MaxLength       =   10
         TabIndex        =   1
         Top             =   1350
         Width           =   1350
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   0
         Left            =   1455
         Picture         =   "frmAlmzAsigPrecios.frx":0620
         ToolTipText     =   "Buscar fecha"
         Top             =   1350
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   1
         Left            =   4110
         Picture         =   "frmAlmzAsigPrecios.frx":06AB
         ToolTipText     =   "Buscar fecha"
         Top             =   1365
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   15
         Left            =   1440
         MouseIcon       =   "frmAlmzAsigPrecios.frx":0736
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar variedad"
         Top             =   2520
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   14
         Left            =   1440
         MouseIcon       =   "frmAlmzAsigPrecios.frx":0888
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar variedad"
         Top             =   2085
         Width           =   240
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   26
         Left            =   735
         TabIndex        =   18
         Top             =   2520
         Width           =   645
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   25
         Left            =   735
         TabIndex        =   17
         Top             =   2130
         Width           =   690
      End
      Begin VB.Label Label3 
         Caption         =   "Asignación Masiva de Precios"
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
         TabIndex        =   16
         Top             =   420
         Width           =   5805
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Variedad"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   240
         Index           =   24
         Left            =   495
         TabIndex        =   15
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   21
         Left            =   3315
         TabIndex        =   14
         Top             =   1365
         Width           =   690
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   20
         Left            =   750
         TabIndex        =   13
         Top             =   1365
         Width           =   690
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   240
         Index           =   19
         Left            =   510
         TabIndex        =   12
         Top             =   1080
         Width           =   600
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
Attribute VB_Name = "frmAlmzAsigPrecios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' +-+-+-+-+-+-+-+-+-+-+-+
' +-+- Autor: MONICA +-+-
' +-+-+-+-+-+-+-+-+-+-+-+

'   LISTADO DE ENTRADAS DE BODEGA

Option Explicit

Public OpcionListado  As String
    ' 0 = Informe de Entradas de Bodega
    ' 1 = Extracto de entradas por Socio / Variedad

Public CadTag As String 'Cadena con el Tag del campo que se va a poner en D/H en los listados
                        'Se necesita si el tipo de codigo es texto

Public Event RectificarFactura(Cliente As String, Observaciones As String)

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
Private WithEvents frmCla As frmComercial 'Ayuda de Clases de comercial
Attribute frmCla.VB_VarHelpID = -1


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


    InicializarVbles
    
    
    If Not DatosOK Then Exit Sub
    
    
    
    'Añadir el parametro de Empresa
    CadParam = CadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    '======== FORMULA  ====================================
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
        
    ' seleccionamos solo almazara
    If Not AnyadirAFormula(cadFormula, "{grupopro.codgrupo} = 5") Then Exit Sub
    If Not AnyadirAFormula(cadSelect, "{grupopro.codgrupo} = 5") Then Exit Sub
    
    nTabla = "(rhisfruta INNER JOIN variedades ON rhisfruta.codvarie = variedades.codvarie) "
    nTabla = "(" & nTabla & ") INNER JOIN productos ON variedades.codprodu = productos.codprodu "
    nTabla = "(" & nTabla & ") INNER JOIN grupopro ON productos.codgrupo = grupopro.codgrupo "
        
    'Comprobar si hay registros a Mostrar antes de abrir el Informe
    If HayRegParaInforme(nTabla, cadSelect) Then
        
        If ModificarPrecioLiq(nTabla, cadSelect) Then
            MsgBox "Proceso realizado correctamente.", vbExclamation
            cmdCancel_Click (1)
        End If
    End If
    
End Sub

Private Sub cmdCancel_Click(Index As Integer)
    Unload Me
End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        
        PonerFoco txtcodigo(12)
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

    For H = 14 To 15
        Me.imgBuscar(H).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next H
    
    'Ocultar todos los Frames de Formulario
    FrameEntradasCampo.visible = False
    
    ' en principio esta activa la asignacion
    Option3(0).Value = True
    Option1(0).Value = True
    
    '###Descomentar
'    CommitConexion
    
    'LISTADOS DE MANTENIMIENTOS BASICOS
    '---------------------
    FrameEntradaBasculaVisible True, H, W
    indFrame = 1
    
    tabla = "rhisfruta"
    
    
    Frame1.visible = True
    Frame1.Enabled = True
    
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


Private Sub frmCla_DatoSeleccionado(CadenaSeleccion As String)
    txtcodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000") ' codigo de clase
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2) ' descripcion
End Sub

Private Sub frmMens_DatoSeleccionado(CadenaSeleccion As String)
Dim SQL As String
Dim Sql2 As String

    If CadenaSeleccion <> "" Then
        SQL = " {variedades.codvarie} in (" & CadenaSeleccion & ")"
        Sql2 = " {variedades.codvarie} in [" & CadenaSeleccion & "]"
    Else
        SQL = " {variedades.codvarie} = -1 "
    End If
    If Not AnyadirAFormula(cadSelect, SQL) Then Exit Sub
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
        Case 0, 1  ' Procedencia(poblacion)
            AbrirFrmPoblacion (Index)
        
        Case 6, 7  ' Clase
            AbrirFrmClase (Index)
        
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



Private Sub Option3_Click(Index As Integer)
    Frame1.Enabled = Option3(0).Value
    Frame3.Enabled = Not Option3(0).Value
    
    If Option3(0).Value Then txtcodigo(0).Text = ""
    If Not Option3(0).Value Then txtcodigo(2).Text = ""
    
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
            Case 6: KEYFecha KeyAscii, 0 'fecha entrada
            Case 7: KEYFecha KeyAscii, 1 'fecha entrada
            
            Case 14: KEYBusqueda KeyAscii, 14 'variedad desde
            Case 15: KEYBusqueda KeyAscii, 15 'variedad hasta
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
        Case 6, 7  'FECHAS
            If txtcodigo(Index).Text <> "" Then PonerFormatoFecha txtcodigo(Index)
            
            
        Case 14, 15 'VARIEDADES
            txtNombre(Index).Text = PonerNombreDeCod(txtcodigo(Index), "variedades", "nomvarie", "codvarie", "N")
            If txtcodigo(Index).Text <> "" Then txtcodigo(Index).Text = Format(txtcodigo(Index).Text, "000000")
        
        Case 2 ' Precio de liquidacion
            PonerFormatoDecimal txtcodigo(Index), 7
        
        Case 0 ' Aumento de precio
            PonerFormatoDecimal txtcodigo(Index), 8
        
    End Select
End Sub

Private Sub FrameEntradaBasculaVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Frame para el listado de socios por seccion
    Me.FrameEntradasCampo.visible = visible
    If visible = True Then
        Me.FrameEntradasCampo.Top = -90
        Me.FrameEntradasCampo.Left = 0
        Me.FrameEntradasCampo.Height = 5610
        Me.FrameEntradasCampo.Width = 6720
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

Private Sub LlamarImprimir()
    With frmImprimir
        .FormulaSeleccion = cadFormula
        .OtrosParametros = CadParam
        .NumeroParametros = numParam
        .SoloImprimir = False
        .EnvioEMail = False
        .Titulo = cadTitulo
        .NombreRPT = cadNombreRPT
        .ConSubInforme = ConSubInforme
        .Opcion = 0
        .Show vbModal
    End With
End Sub


Private Sub AbrirFrmPoblacion(Indice As Integer)
    indCodigo = Indice
    Set frmPob = New frmManPueblos
    frmPob.Caption = "Pueblos"
    frmPob.DatosADevolverBusqueda = "0|1|"
    frmPob.Show vbModal
    Set frmPob = Nothing
End Sub


Private Sub AbrirFrmSocios(Indice As Integer)
    indCodigo = Indice
    Set frmSoc = New frmManSocios
    frmSoc.DatosADevolverBusqueda = "0|1|"
    frmSoc.Show vbModal
    Set frmSoc = Nothing
End Sub

Private Sub AbrirFrmClase(Indice As Integer)
    indCodigo = Indice + 14
    
    Set frmCla = New frmComercial
    
    AyudaClasesCom frmCla, txtcodigo(Indice).Text
    
    Set frmCla = Nothing
End Sub

Private Sub AbrirFrmVariedad(Indice As Integer)
    indCodigo = Indice
    Set frmVar = New frmComVar
    frmVar.DatosADevolverBusqueda = "0|1|"
    frmVar.Show vbModal
    Set frmVar = Nothing
End Sub


Private Sub AbrirVisReport()
    Screen.MousePointer = vbHourglass
    CadenaDesdeOtroForm = ""
    With frmVisReport
        .FormulaSeleccion = cadFormula
'        .SoloImprimir = (Me.OptVisualizar(indFrame).Value = 1)
        .OtrosParametros = CadParam
        .NumeroParametros = numParam
        '##descomen
'        .MostrarTree = MostrarTree
'        .Informe = MIPATH & Nombre
'        .InfConta = InfConta
        '##
        
'        If NombreSubRptConta <> "" Then
'            .SubInformeConta = NombreSubRptConta
'        Else
'            .SubInformeConta = ""
'        End If
        '##descomen
'        .ConSubInforme = ConSubInforme
        '##
        .Opcion = 0
'        .ExportarPDF = (chkEMAIL.Value = 1)
        .Show vbModal
    End With
    
'    If Me.chkEMAIL.Value = 1 Then
'    '####Descomentar
'        If CadenaDesdeOtroForm <> "" Then frmEMail.Show vbModal
'    End If
    Unload Me
End Sub

Private Sub AbrirEMail()
    If CadenaDesdeOtroForm <> "" Then frmEMail.Show vbModal
End Sub


Private Function ConcatenarCampos(cTabla As String, cWhere As String) As String
Dim Rs As ADODB.Recordset
Dim SQL As String
Dim Sql1 As String

    ConcatenarCampos = ""

    cTabla = QuitarCaracterACadena(cTabla, "{")
    cTabla = QuitarCaracterACadena(cTabla, "}")
    SQL = "Select rcampos.codcampo FROM " & QuitarCaracterACadena(cTabla, "_1")
    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        SQL = SQL & " WHERE " & cWhere
    End If
    
    
    SQL = "select distinct rcampos.codcampo  from " & cTabla & " where " & cWhere
    Set Rs = New ADODB.Recordset
    
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
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
Dim SQL As String

    On Error GoTo eActualizarRegistros

    ActualizarRegistros = False
    cTabla = QuitarCaracterACadena(cTabla, "{")
    cTabla = QuitarCaracterACadena(cTabla, "}")
    SQL = "update " & QuitarCaracterACadena(cTabla, "_1") & " set impreso = 1 "
    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        SQL = SQL & " WHERE " & cWhere
    End If
    
    conn.Execute SQL
    
    ActualizarRegistros = True
    Exit Function
    
eActualizarRegistros:
    MuestraError Err.Number, "Actualizando registros", Err.Description
End Function


Private Function NombreCalidad(Var As String, Calid As String) As String
Dim SQL As String

    NombreCalidad = ""

    SQL = "select nomcalab from rcalidad where codvarie = " & DBSet(Var, "N")
    SQL = SQL & " and codcalid = " & DBSet(Calid, "N")
    
    NombreCalidad = DevuelveValor(SQL)
    
End Function




Private Function CopiarFichero() As Boolean
Dim nomFich As String

On Error GoTo ecopiarfichero

    CopiarFichero = False
    ' abrimos el commondialog para indicar donde guardarlo
'    Me.CommonDialog1.InitDir = App.path

    Me.cd1.DefaultExt = "txt"
    
    cd1.Filter = "Archivos txt|txt|"
    cd1.FilterIndex = 1
    
    ' copiamos el primer fichero
    cd1.FileName = "fichero.txt"
        
    Me.cd1.ShowSave
    
    If cd1.FileName <> "" Then
        FileCopy App.Path & "\fichero.txt", cd1.FileName
    End If
    
    CopiarFichero = True
    Exit Function

ecopiarfichero:
    If Err.Number <> 0 Then
        MuestraError Err.Number, Err.Description
    End If
    Err.Clear
End Function


Private Function ProductoCampo(campo As String) As String
Dim SQL As String

    ProductoCampo = ""
    
    SQL = "select variedades.codprodu from rcampos inner join variedades on rcampos.codvarie = variedades.codvarie "
    SQL = SQL & " where rcampos.codcampo = " & DBSet(campo, "N")
    
    ProductoCampo = DevuelveValor(SQL)

End Function



Private Function ModificarPrecioLiq(cTabla As String, cWhere As String) As Boolean
'Comprobar si hay registros a Mostrar antes de abrir el Informe
Dim SQL As String
Dim Sql2 As String
Dim cWhere2 As String
Dim Precio As Currency
Dim Rs As ADODB.Recordset

    On Error GoTo eModificarRdto

    ModificarPrecioLiq = False
 
    cTabla = QuitarCaracterACadena(cTabla, "{")
    cTabla = QuitarCaracterACadena(cTabla, "}")
    SQL = "Select numalbar FROM " & QuitarCaracterACadena(cTabla, "_1")
    If cWhere <> "" Then
        cWhere = QuitarCaracterACadena(cWhere, "{")
        cWhere = QuitarCaracterACadena(cWhere, "}")
        cWhere = QuitarCaracterACadena(cWhere, "_1")
        SQL = SQL & " WHERE " & cWhere
    End If
    
    cWhere2 = " numalbar in (" & SQL & ")"
    
    If Not BloqueaRegistro("rhisfruta", cWhere2) Then
        MsgBox "No se pueden actualizar Entradas. Hay registros bloqueados.", vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    Else
        conn.BeginTrans
    
        Set Rs = New ADODB.Recordset
        Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        While Not Rs.EOF
            
            If Option3(0).Value Then
            
                Precio = CCur(ImporteSinFormato(txtcodigo(2).Text))
                
                Sql2 = "update rhisfruta set prliquidalmz = " & DBSet(Precio, "N")
                Sql2 = Sql2 & " where numalbar = " & DBSet(Rs!numalbar, "N")
                
                conn.Execute Sql2
        
            Else
                If Option1(0).Value Then
                    Precio = CCur(ImporteSinFormato(txtcodigo(0).Text))
                    
                    Sql2 = "update rhisfruta set prliquidalmz = round(prliquidalmz + " & DBSet(Precio, "N") & ",4) "
                    Sql2 = Sql2 & " where numalbar = " & DBSet(Rs!numalbar, "N")
                Else
                    Precio = 1 + (CCur(ImporteSinFormato(txtcodigo(0).Text)) / 100)
                    
                    Sql2 = "update rhisfruta set prliquidalmz = round(prliquidalmz * " & DBSet(Precio, "N") & ",4) "
                    Sql2 = Sql2 & " where numalbar = " & DBSet(Rs!numalbar, "N")
                End If
                conn.Execute Sql2
            End If
        
        
            Rs.MoveNext
        Wend
        
        conn.CommitTrans
        ModificarPrecioLiq = True
        Exit Function
    End If
    
eModificarRdto:
    conn.RollbackTrans
    TerminaBloquear
    ModificarPrecioLiq = False
End Function



Private Function DatosOK() As Boolean
    
    DatosOK = False
    
    ' si estamos asignando precios
    If Option3(0).Value Then
    
        If txtcodigo(2).Text = "" Then
            MsgBox "Debe poner un valor en el campo precio. Revise.", vbExclamation
            PonerFoco txtcodigo(2)
            Exit Function
        End If

    Else
    ' si estamos aumentando precio
        If txtcodigo(0).Text = "" Then
            MsgBox "Debe poner un valor en el campo de aumento de precio. Revise.", vbExclamation
            PonerFoco txtcodigo(0)
            Exit Function
        End If
    End If

    DatosOK = True

End Function

