VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmAlmzListEntradas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informes"
   ClientHeight    =   7605
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   6690
   Icon            =   "frmAlmzListEntradas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7605
   ScaleWidth      =   6690
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cd1 
      Left            =   6960
      Top             =   5220
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame FrameEntradasCampo 
      Height          =   7515
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6615
      Begin VB.CheckBox Check6 
         Caption         =   "Sólo campos sin identificar"
         Height          =   285
         Left            =   3330
         TabIndex        =   50
         Top             =   6450
         Width           =   2445
      End
      Begin VB.CheckBox Check5 
         Caption         =   "Incluir datos del campo"
         Height          =   285
         Left            =   630
         TabIndex        =   49
         Top             =   6450
         Width           =   2445
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Ordenado por Fecha"
         Height          =   285
         Left            =   630
         TabIndex        =   46
         Top             =   6795
         Width           =   2445
      End
      Begin VB.CheckBox Check4 
         Caption         =   "Detallar variedad  "
         Height          =   285
         Left            =   630
         TabIndex        =   48
         Top             =   6810
         Width           =   2445
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Imprimir sólo Totales  "
         Height          =   285
         Left            =   630
         TabIndex        =   47
         Top             =   6495
         Width           =   2445
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Salta página por Socio"
         Height          =   285
         Left            =   630
         TabIndex        =   45
         Top             =   6180
         Width           =   2445
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   855
         Left            =   600
         TabIndex        =   39
         Top             =   5460
         Width           =   4935
         Begin VB.TextBox txtcodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   3
            Left            =   3630
            MaxLength       =   10
            TabIndex        =   41
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
            TabIndex        =   40
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
            TabIndex        =   44
            Top             =   90
            Width           =   540
         End
         Begin VB.Label Label2 
            Caption         =   "Desde"
            Height          =   195
            Index           =   4
            Left            =   390
            TabIndex        =   43
            Top             =   360
            Width           =   465
         End
         Begin VB.Label Label2 
            Caption         =   "Hasta"
            Height          =   195
            Index           =   3
            Left            =   2700
            TabIndex        =   42
            Top             =   360
            Width           =   420
         End
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   2745
         Locked          =   -1  'True
         TabIndex        =   35
         Text            =   "Text5"
         Top             =   4155
         Width           =   3375
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   2745
         Locked          =   -1  'True
         TabIndex        =   34
         Text            =   "Text5"
         Top             =   4545
         Width           =   3375
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   1920
         MaxLength       =   6
         TabIndex        =   7
         Top             =   4155
         Width           =   735
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   1
         Left            =   1920
         MaxLength       =   6
         TabIndex        =   8
         Top             =   4545
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   21
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   30
         Text            =   "Text5"
         Top             =   2580
         Width           =   3375
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   20
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   29
         Text            =   "Text5"
         Top             =   2220
         Width           =   3375
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   21
         Left            =   1935
         MaxLength       =   3
         TabIndex        =   4
         Top             =   2580
         Width           =   735
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   20
         Left            =   1935
         MaxLength       =   3
         TabIndex        =   3
         Top             =   2220
         Width           =   735
      End
      Begin VB.CommandButton Command4 
         Height          =   440
         Left            =   7860
         Picture         =   "frmAlmzListEntradas.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   1440
         Visible         =   0   'False
         Width           =   380
      End
      Begin VB.CommandButton Command3 
         Height          =   440
         Left            =   7860
         Picture         =   "frmAlmzListEntradas.frx":0316
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   2215
         Visible         =   0   'False
         Width           =   380
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   15
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   16
         Text            =   "Text5"
         Top             =   3570
         Width           =   3375
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   14
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   15
         Text            =   "Text5"
         Top             =   3210
         Width           =   3375
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   15
         Left            =   1935
         MaxLength       =   6
         TabIndex        =   6
         Top             =   3570
         Width           =   735
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   14
         Left            =   1935
         MaxLength       =   6
         TabIndex        =   5
         Top             =   3210
         Width           =   735
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   13
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   14
         Text            =   "Text5"
         Top             =   1635
         Width           =   3375
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   12
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   13
         Text            =   "Text5"
         Top             =   1275
         Width           =   3375
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   13
         Left            =   1935
         MaxLength       =   6
         TabIndex        =   2
         Top             =   1635
         Width           =   750
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   12
         Left            =   1935
         MaxLength       =   6
         TabIndex        =   1
         Top             =   1275
         Width           =   750
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   4080
         TabIndex        =   11
         Top             =   6855
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   1
         Left            =   5160
         TabIndex        =   12
         Top             =   6855
         Width           =   975
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   7
         Left            =   4245
         MaxLength       =   10
         TabIndex        =   10
         Top             =   5145
         Width           =   1095
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   6
         Left            =   1920
         MaxLength       =   10
         TabIndex        =   9
         Top             =   5130
         Width           =   1095
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   1620
         MouseIcon       =   "frmAlmzListEntradas.frx":0620
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar pueblo"
         Top             =   4170
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   1
         Left            =   1620
         MouseIcon       =   "frmAlmzListEntradas.frx":0772
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar pueblo"
         Top             =   4560
         Width           =   240
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   0
         Left            =   990
         TabIndex        =   38
         Top             =   4620
         Width           =   420
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   1
         Left            =   990
         TabIndex        =   37
         Top             =   4230
         Width           =   465
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Procedencia"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   2
         Left            =   660
         TabIndex        =   36
         Top             =   3990
         Width           =   900
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   7
         Left            =   1620
         MouseIcon       =   "frmAlmzListEntradas.frx":08C4
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar clase"
         Top             =   2610
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   6
         Left            =   1620
         MouseIcon       =   "frmAlmzListEntradas.frx":0A16
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar clase"
         Top             =   2220
         Width           =   240
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   28
         Left            =   1005
         TabIndex        =   33
         Top             =   2655
         Width           =   420
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   18
         Left            =   1005
         TabIndex        =   32
         Top             =   2265
         Width           =   465
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Clase"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   11
         Left            =   675
         TabIndex        =   31
         Top             =   2025
         Width           =   390
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   0
         Left            =   1590
         Picture         =   "frmAlmzListEntradas.frx":0B68
         ToolTipText     =   "Buscar fecha"
         Top             =   5130
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   1
         Left            =   3930
         Picture         =   "frmAlmzListEntradas.frx":0BF3
         ToolTipText     =   "Buscar fecha"
         Top             =   5145
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   15
         Left            =   1620
         MouseIcon       =   "frmAlmzListEntradas.frx":0C7E
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar variedad"
         Top             =   3600
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   14
         Left            =   1620
         MouseIcon       =   "frmAlmzListEntradas.frx":0DD0
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar variedad"
         Top             =   3210
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   13
         Left            =   1620
         MouseIcon       =   "frmAlmzListEntradas.frx":0F22
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1665
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   12
         Left            =   1620
         MouseIcon       =   "frmAlmzListEntradas.frx":1074
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1275
         Width           =   240
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Socio"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   27
         Left            =   675
         TabIndex        =   28
         Top             =   1080
         Width           =   405
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   26
         Left            =   1005
         TabIndex        =   27
         Top             =   3645
         Width           =   420
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   25
         Left            =   1005
         TabIndex        =   26
         Top             =   3255
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "Informe de Entradas"
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
         Left            =   630
         TabIndex        =   25
         Top             =   420
         Width           =   5805
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Variedad"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   24
         Left            =   675
         TabIndex        =   24
         Top             =   3015
         Width           =   630
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   23
         Left            =   960
         TabIndex        =   23
         Top             =   1680
         Width           =   420
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   22
         Left            =   960
         TabIndex        =   22
         Top             =   1320
         Width           =   465
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   21
         Left            =   3315
         TabIndex        =   21
         Top             =   5145
         Width           =   420
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   20
         Left            =   975
         TabIndex        =   20
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
         TabIndex        =   19
         Top             =   4950
         Width           =   450
      End
   End
End
Attribute VB_Name = "frmAlmzListEntradas"
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
    
    ' 2 = Impresion personalizada de los albaranes de almazara (Castelduc)

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
Dim Tabla As String
Dim Tabla1 As String
Dim Codigo As String 'Código para FormulaSelection de Crystal Report
Dim TipCod As String
Dim Orden1 As String 'Campo de Ordenacion (por codigo) para Cristal Report
Dim Orden2 As String 'Campo de Ordenacion (por nombre) para Cristal Report
Dim Tipo As String

Dim indice As Integer

Dim PrimeraVez As Boolean
Dim Contabilizada As Byte
Dim ConSubInforme As Boolean


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 0, cerrar
    If cerrar Then Unload Me

End Sub

Private Sub Check1_Click()
    Check3.Enabled = (Check1.Value = 0)
    Check4.Enabled = (Check1.Value = 0)
    If Not Check3.Enabled Then Check3.Value = 0
    If Not Check4.Enabled Then Check4.Value = 0
End Sub



Private Sub Check2_Click()
    Check6.Enabled = (Check2.Value = 0)
    If Not Check6.Enabled Then Check6.Value = 0
End Sub

Private Sub Check5_Click()
    Check2.Enabled = (Check5.Value = 0)
    If Not Check2.Enabled Then Check2.Value = 0
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
        Codigo = "{" & Tabla & ".codsocio}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHSocio=""") Then Exit Sub
    End If
    
    
    'D/H CLASE
    cDesde = Trim(txtcodigo(20).Text)
    cHasta = Trim(txtcodigo(21).Text)
    nDesde = txtNombre(20).Text
    nHasta = txtNombre(21).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{variedades.codclase}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHClase=""") Then Exit Sub
    End If
    
    
    vSQL = ""
    If txtcodigo(20).Text <> "" Then vSQL = vSQL & " and variedades.codclase >= " & DBSet(txtcodigo(20).Text, "N")
    If txtcodigo(21).Text <> "" Then vSQL = vSQL & " and variedades.codclase <= " & DBSet(txtcodigo(21).Text, "N")
    
    
    'D/H VARIEDAD
    cDesde = Trim(txtcodigo(14).Text)
    cHasta = Trim(txtcodigo(15).Text)
    nDesde = txtNombre(14).Text
    nHasta = txtNombre(15).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{" & Tabla & ".codvarie}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHVariedad=""") Then Exit Sub
    End If

    If txtcodigo(14).Text <> "" Then vSQL = vSQL & " and variedades.codvarie >= " & DBSet(txtcodigo(14).Text, "N")
    If txtcodigo(15).Text <> "" Then vSQL = vSQL & " and variedades.codvarie <= " & DBSet(txtcodigo(15).Text, "N")

    'D/H PROCEDENCIA(POBLACION)
    cDesde = Trim(txtcodigo(0).Text)
    cHasta = Trim(txtcodigo(1).Text)
    nDesde = txtNombre(0).Text
    nHasta = txtNombre(1).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{" & Tabla & ".codpobla}"
        TipCod = "T"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHPoblacion=""") Then Exit Sub
    End If
    
    'D/H fecha
    cDesde = Trim(txtcodigo(6).Text)
    cHasta = Trim(txtcodigo(7).Text)
    nDesde = ""
    nHasta = ""
    If Not (cDesde = "" And cHasta = "") Then
                'Cadena para seleccion Desde y Hasta
        Codigo = "{" & Tabla & ".fecalbar}"
        TipCod = "F"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHFecha=""") Then Exit Sub
    End If
        
        
    If OpcionListado = 0 Or OpcionListado = 2 Then
        'D/H nro de albaran
        cDesde = Trim(txtcodigo(2).Text)
        cHasta = Trim(txtcodigo(3).Text)
        nDesde = ""
        nHasta = ""
        If Not (cDesde = "" And cHasta = "") Then
                    'Cadena para seleccion Desde y Hasta
            Codigo = "{" & Tabla & ".numalbar}"
            TipCod = "N"
            If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHAlbaran=""") Then Exit Sub
        End If
    End If
        
    If Not AnyadirAFormula(cadFormula, "{grupopro.codgrupo} = 5") Then Exit Sub
    If Not AnyadirAFormula(cadSelect, "{grupopro.codgrupo} = 5") Then Exit Sub
    
    nTabla = "(rhisfruta INNER JOIN variedades ON rhisfruta.codvarie = variedades.codvarie) "
    nTabla = "(" & nTabla & ") INNER JOIN productos ON variedades.codprodu = productos.codprodu "
    nTabla = "(" & nTabla & ") INNER JOIN grupopro ON productos.codgrupo = grupopro.codgrupo "
        
        
    If OpcionListado = 0 Then
        If Check6.Value Then
            If Not AnyadirAFormula(cadFormula, "{rhisfruta.codcampo} = 0") Then Exit Sub
            If Not AnyadirAFormula(cadSelect, "{rhisfruta.codcampo} = 0") Then Exit Sub
        End If
    End If
        
    'Comprobar si hay registros a Mostrar antes de abrir el Informe
    If HayRegParaInforme(nTabla, cadSelect) Then
        
        ConSubInforme = False
        
        Select Case OpcionListado
            Case 0
                cadNombreRPT = "rAlmzInfEntradas.rpt" '"rInfEntradasClas.rpt"
                cadTitulo = "Informe de Entradas Almazara"
                
                indRPT = 72 ' informe de entradas
                
                If Not PonerParamRPT(indRPT, CadParam, numParam, nomDocu) Then Exit Sub
                
                cadNombreRPT = nomDocu ' rAlmzExtSocEntradas.rpt
                
                                
                If Check2.Value Then
                    cadNombreRPT = Replace(cadNombreRPT, "InfEntradas", "InfEntradasdiario") '"rAlmzInfEntradasDiario.rpt"
                End If
                '[Monica]05/11/2015: nuevo listado
                If Check5.Value Then
                    cadNombreRPT = "rAlmzInfEntradasCampo.rpt"
                End If
            Case 1
                If Check1.Value = 0 Then
                    ' no saltamos pagina por socio
                    CadParam = CadParam & "pSoloTotales=" & Check3.Value & "|"
                    numParam = numParam + 1
                    
                    CadParam = CadParam & "pDetalleVariedad=" & Check4.Value & "|"
                    numParam = numParam + 1
                    
                    cadNombreRPT = "rAlmzExtEntradas.rpt"
                    ConSubInforme = True
                Else
                    ' si saltamos pagina por socio y la cooperativa es Moixent: está personalizado
                    If vParamAplic.Cooperativa = 3 Then
                        indRPT = 35 ' extracto de entradas por socio Almazara
                        
                        If Not PonerParamRPT(indRPT, CadParam, numParam, nomDocu) Then Exit Sub
                        
                        cadNombreRPT = nomDocu ' rAlmzExtSocEntradas.rpt
                    Else
                        ' saltamos pagina por socio
                        cadNombreRPT = "rAlmzExtSocEntradas.rpt"
                    End If
                End If
                cadTitulo = "Extracto Entradas Almazara Socio/Variedad"
            
            Case 2 ' reimpresion de albaranes de almazara
                cadNombreRPT = "CasAlmzAlbaran.rpt" '"CasAlmzAlbaran.rpt"
                cadTitulo = "Impresión Albaranes de Almazara"
        End Select
        
        LlamarImprimir
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

    For H = 0 To 1
        Me.imgBuscar(H).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next H
    For H = 12 To 15
        Me.imgBuscar(H).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next H
    For H = 6 To 7
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
    
    Tabla = "rhisfruta"
    
    
    Select Case OpcionListado
        Case 0
            Label3.Caption = "Informe de Entradas Almazara"
        Case 1
            Label3.Caption = "Extracto Entradas por Socio/Variedad"
        Case 2
            Label3.Caption = "Reimpresión Albaranes Almazara"
    End Select
    
    Check1.Value = False
    Check1.visible = (OpcionListado = 1)
    Check1.Enabled = (OpcionListado = 1)
    
    ' solo puede ser informe diario si estamos en informe de entradas
    ' no en el extracto
    Check2.Value = False
    Check2.visible = (OpcionListado = 0)
    Check2.Enabled = (OpcionListado = 0)
    ' solo desde/hasta albaran si estamos en informe de entradas
    Frame1.visible = (OpcionListado = 0) Or (OpcionListado = 2)
    Frame1.Enabled = (OpcionListado = 0) Or (OpcionListado = 2)
    
    Check3.Value = False
    Check3.visible = (OpcionListado = 1)
    Check3.Enabled = (OpcionListado = 1)
    
    Check4.Value = False
    Check4.visible = (OpcionListado = 1)
    Check4.Enabled = (OpcionListado = 1)
        
    '[Monica]05/11/2015: en el caso de que quiera mostrar los datos del campo
    Check5.Value = False
    Check5.visible = (OpcionListado = 0)
    Check5.Enabled = (OpcionListado = 0)
    Check6.Value = False
    Check6.visible = (OpcionListado = 0)
    Check6.Enabled = (OpcionListado = 0)
    
    
    If OpcionListado = 2 Then
        txtcodigo(2).Text = CadTag
        txtcodigo(3).Text = CadTag
        Check1.visible = False
        Check1.Enabled = False
        Check2.visible = False
        Check2.Enabled = False
        Check3.visible = False
        Check3.Enabled = False
        Check4.visible = False
        Check4.Enabled = False
        Check5.visible = False
        Check5.Enabled = False
        Check6.visible = False
        Check6.Enabled = False
    End If
    
    
    'Esto se consigue poneinedo el cancel en el opcion k corresponda
'    Me.cmdCancel(indFrame).Cancel = True
    Me.Width = W + 70
    Me.Height = H + 350
End Sub

Private Sub frmC_Selec(vFecha As Date)
    ' *** repasar si el camp es txtAux o Text1 ***
    txtcodigo(indice).Text = Format(vFecha, "dd/mm/yyyy") '<===
    ' ********************************************
End Sub


Private Sub frmCla_DatoSeleccionado(CadenaSeleccion As String)
    txtcodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000") ' codigo de clase
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2) ' descripcion
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
            indice = Index + 6
    End Select


    imgFec(0).Tag = indice '<===
    ' *** repasar si el camp es txtAux o Text1 ***
    If txtcodigo(indice).Text <> "" Then frmC.NovaData = txtcodigo(indice).Text
    ' ********************************************

    frmC.Show vbModal
    Set frmC = Nothing
    ' *** repasar si el camp es txtAux o Text1 ***
    PonerFoco txtcodigo(indice) '<===
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

Private Sub KEYBusqueda(KeyAscii As Integer, indice As Integer)
    KeyAscii = 0
    imgBuscar_Click (indice)
End Sub

Private Sub KEYFecha(KeyAscii As Integer, indice As Integer)
    KeyAscii = 0
    imgFec_Click (indice)
End Sub


Private Sub txtCodigo_LostFocus(Index As Integer)
Dim Cad As String, cadTipo As String 'tipo cliente

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


Private Sub AbrirFrmPoblacion(indice As Integer)
    indCodigo = indice
    Set frmPob = New frmManPueblos
    frmPob.Caption = "Pueblos"
    frmPob.DatosADevolverBusqueda = "0|1|"
    frmPob.Show vbModal
    Set frmPob = Nothing
End Sub


Private Sub AbrirFrmSocios(indice As Integer)
    indCodigo = indice
    Set frmSoc = New frmManSocios
    frmSoc.DatosADevolverBusqueda = "0|1|"
    frmSoc.Show vbModal
    Set frmSoc = Nothing
End Sub

Private Sub AbrirFrmClase(indice As Integer)
    indCodigo = indice + 14
    
    Set frmCla = New frmComercial
    
    AyudaClasesCom frmCla, txtcodigo(indice).Text
    
    Set frmCla = Nothing
End Sub

Private Sub AbrirFrmVariedad(indice As Integer)
    indCodigo = indice
    Set frmVar = New frmComVar
    frmVar.DatosADevolverBusqueda = "0|1|"
    frmVar.Show vbModal
    Set frmVar = Nothing
End Sub

