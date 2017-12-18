VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmGeneraPrecios 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5340
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   7230
   Icon            =   "frmGeneraPrecios.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5340
   ScaleWidth      =   7230
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cd1 
      Left            =   6075
      Top             =   4770
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame FrameGeneraPreciosMasiva 
      Height          =   5310
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   7155
      Begin VB.ComboBox Combo1 
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
         Left            =   1980
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   1575
         Width           =   1575
      End
      Begin VB.TextBox txtCodigo 
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
         Left            =   1980
         MaxLength       =   10
         TabIndex        =   3
         Top             =   2475
         Width           =   1350
      End
      Begin VB.TextBox txtCodigo 
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
         Index           =   1
         Left            =   1980
         MaxLength       =   10
         TabIndex        =   2
         Top             =   2070
         Width           =   1350
      End
      Begin VB.CommandButton CmdCancel 
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
         Index           =   0
         Left            =   5550
         TabIndex        =   8
         Top             =   4545
         Width           =   1065
      End
      Begin VB.CommandButton CmdAcepGen 
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
         Index           =   0
         Left            =   4380
         TabIndex        =   7
         Top             =   4545
         Width           =   1065
      End
      Begin VB.TextBox txtCodigo 
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
         Index           =   3
         Left            =   1980
         MaxLength       =   7
         TabIndex        =   4
         Top             =   3060
         Width           =   1335
      End
      Begin VB.TextBox txtCodigo 
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
         Left            =   1980
         MaxLength       =   6
         TabIndex        =   0
         Top             =   1095
         Width           =   825
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
         Index           =   0
         Left            =   2835
         Locked          =   -1  'True
         TabIndex        =   12
         Text            =   "Text5"
         Top             =   1095
         Width           =   3825
      End
      Begin VB.CommandButton Command10 
         Height          =   440
         Left            =   7860
         Picture         =   "frmGeneraPrecios.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   2215
         Visible         =   0   'False
         Width           =   380
      End
      Begin VB.CommandButton Command9 
         Height          =   440
         Left            =   7860
         Picture         =   "frmGeneraPrecios.frx":0316
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   1440
         Visible         =   0   'False
         Width           =   380
      End
      Begin VB.TextBox txtCodigo 
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
         Left            =   1980
         MaxLength       =   7
         TabIndex        =   5
         Top             =   3465
         Width           =   1320
      End
      Begin VB.TextBox txtCodigo 
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
         Left            =   1980
         MaxLength       =   30
         TabIndex        =   6
         Top             =   3870
         Width           =   4650
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Fin"
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
         Index           =   39
         Left            =   450
         TabIndex        =   20
         Top             =   2520
         Width           =   960
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Factura"
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
         Index           =   38
         Left            =   450
         TabIndex        =   19
         Top             =   1620
         Width           =   1260
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Texto"
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
         Index           =   37
         Left            =   450
         TabIndex        =   18
         Top             =   3915
         Width           =   570
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Inicio"
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
         Index           =   50
         Left            =   450
         TabIndex        =   17
         Top             =   2115
         Width           =   1200
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
         Index           =   45
         Left            =   450
         TabIndex        =   16
         Top             =   1125
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Generacion de Precios Masiva"
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
         TabIndex        =   15
         Top             =   405
         Width           =   5025
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Precio Socio"
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
         Index           =   42
         Left            =   450
         TabIndex        =   14
         Top             =   3105
         Width           =   1215
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   1710
         MouseIcon       =   "frmGeneraPrecios.frx":0620
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar variedad"
         Top             =   1095
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   1
         Left            =   1710
         Picture         =   "frmGeneraPrecios.frx":0772
         ToolTipText     =   "Buscar fecha"
         Top             =   2475
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   0
         Left            =   1710
         Picture         =   "frmGeneraPrecios.frx":07FD
         ToolTipText     =   "Buscar fecha"
         Top             =   2070
         Width           =   240
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Precio Coop."
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
         Index           =   41
         Left            =   450
         TabIndex        =   13
         Top             =   3510
         Width           =   1245
      End
   End
   Begin VB.Frame FrameCopiaVariedad 
      Height          =   4680
      Left            =   0
      TabIndex        =   21
      Top             =   0
      Width           =   7200
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
         Index           =   9
         Left            =   3240
         Locked          =   -1  'True
         TabIndex        =   37
         Text            =   "Text5"
         Top             =   1590
         Width           =   3375
      End
      Begin VB.TextBox txtCodigo 
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
         Index           =   8
         Left            =   2445
         MaxLength       =   6
         TabIndex        =   30
         Top             =   1080
         Width           =   735
      End
      Begin VB.CommandButton Command3 
         Height          =   440
         Left            =   7860
         Picture         =   "frmGeneraPrecios.frx":0888
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   1440
         Visible         =   0   'False
         Width           =   380
      End
      Begin VB.CommandButton Command2 
         Height          =   440
         Left            =   7860
         Picture         =   "frmGeneraPrecios.frx":0B92
         Style           =   1  'Graphical
         TabIndex        =   23
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
         Index           =   8
         Left            =   3240
         Locked          =   -1  'True
         TabIndex        =   22
         Text            =   "Text5"
         Top             =   1095
         Width           =   3375
      End
      Begin VB.TextBox txtCodigo 
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
         Index           =   9
         Left            =   2445
         MaxLength       =   6
         TabIndex        =   31
         Top             =   1590
         Width           =   735
      End
      Begin VB.CommandButton CmdAcepGen 
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
         Index           =   1
         Left            =   4530
         TabIndex        =   35
         Top             =   3900
         Width           =   975
      End
      Begin VB.CommandButton CmdCancel 
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
         Left            =   5610
         TabIndex        =   36
         Top             =   3915
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
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
         Left            =   2445
         MaxLength       =   10
         TabIndex        =   34
         Top             =   3120
         Width           =   1350
      End
      Begin VB.TextBox txtCodigo 
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
         Left            =   2445
         MaxLength       =   10
         TabIndex        =   33
         Top             =   2700
         Width           =   1350
      End
      Begin VB.ComboBox Combo1 
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
         Left            =   2445
         Style           =   2  'Dropdown List
         TabIndex        =   32
         Top             =   2205
         Width           =   1800
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   2
         Left            =   2145
         MouseIcon       =   "frmGeneraPrecios.frx":0E9C
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar variedad"
         Top             =   1590
         Width           =   240
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Variedad Destino"
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
         Index           =   2
         Left            =   420
         TabIndex        =   38
         Top             =   1620
         Width           =   1665
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   3
         Left            =   2145
         Picture         =   "frmGeneraPrecios.frx":0FEE
         ToolTipText     =   "Buscar fecha"
         Top             =   3120
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   2
         Left            =   2145
         Picture         =   "frmGeneraPrecios.frx":1079
         ToolTipText     =   "Buscar fecha"
         Top             =   2700
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   1
         Left            =   2145
         MouseIcon       =   "frmGeneraPrecios.frx":1104
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar variedad"
         Top             =   1095
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Copia de Precios a otra Variedad"
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
         Left            =   390
         TabIndex        =   29
         Top             =   405
         Width           =   5025
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Variedad Origen"
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
         Index           =   4
         Left            =   420
         TabIndex        =   28
         Top             =   1125
         Width           =   1560
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Inicio"
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
         Index           =   3
         Left            =   450
         TabIndex        =   27
         Top             =   2745
         Width           =   1200
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Factura"
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
         Index           =   1
         Left            =   420
         TabIndex        =   26
         Top             =   2250
         Width           =   1260
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Fin"
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
         Left            =   450
         TabIndex        =   25
         Top             =   3150
         Width           =   960
      End
   End
End
Attribute VB_Name = "frmGeneraPrecios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' +-+-+-+-+-+-+-+-+-+-+-
' +-+- Autor: LAURA +-+-
' +-+-+-+-+-+-+-+-+-+-+-

Option Explicit

Public OpcionGenerar As Integer
Public NumCod As String 'Para indicar cod. Traspaso,Movimiento, etc. que llama
                        'Para indicar nº oferta a imprimir

Public CadTag As String 'Cadena con el Tag del campo que se va a poner en D/H en los listados
                        'Se necesita si el tipo de codigo es texto

Public Event RectificarFactura(Cliente As String, Observaciones As String)

Private Conexion As Byte
'1.- Conexión a BD Ariges  2.- Conexión a BD Conta

Private HaDevueltoDatos As Boolean


Private WithEvents frmVar As frmComVar 'Variedades de comercial
Attribute frmVar.VB_VarHelpID = -1
Private WithEvents frmCal As frmManCalidades 'Calidades
Attribute frmCal.VB_VarHelpID = -1
Private WithEvents frmB As frmBuscaGrid  'Busquedas
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmC As frmCal 'calendario fechas
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmSit As frmManSituCamp 'Situacion campos
Attribute frmSit.VB_VarHelpID = -1
Private WithEvents frmMens As frmMensajes 'Mensajes
Attribute frmMens.VB_VarHelpID = -1


'GENERALES PARA PASARLE A CRYSTAL REPORT
Private cadFormula As String 'Cadena con la FormulaSelection para Crystal Report
Private CadParam As String 'Cadena con los parametros para Crystal Report
Private numParam As Byte 'Numero de parametros que se pasan a Crystal Report
Private cadSelect As String 'Cadena para comprobar si hay datos antes de abrir Informe
Private cadTitulo As String 'Titulo para la ventana frmImprimir
Private cadNombreRPT As String 'Nombre del informe


Dim indCodigo As Integer 'indice para txtCodigo
Dim indFrame As Single 'nº de frame en el que estamos
 
'Se inicializan para cada Informe (tabla de BD a la que hace referencia
Dim tabla As String
Dim Codigo As String 'Código para FormulaSelection de Crystal Report
Dim TipCod As String
Dim Orden1 As String 'Campo de Ordenacion (por codigo) para Cristal Report
Dim Orden2 As String 'Campo de Ordenacion (por nombre) para Cristal Report
Dim Tipo As String


Dim PrimeraVez As Boolean
Dim Contabilizada As Byte

Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 0, cerrar
    If cerrar Then Unload Me

End Sub


Private Sub CmdAcepGen_Click(Index As Integer)
Dim cDesde As String
Dim cHasta As String
Dim nDesde As String
Dim nHasta As String

    Select Case Index
        Case 0
            cDesde = Trim(txtCodigo(1).Text)
            cHasta = Trim(txtCodigo(2).Text)
            nDesde = ""
            nHasta = ""
            If Not (cDesde = "" And cHasta = "") Then
        '        Codigo = "{" & Tabla & ".fechaent}"
                TipCod = "F"
                If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHFecha=""") Then Exit Sub
            End If
        
            If DatosOK Then
                If GeneraRegistros Then
                    MsgBox "Proceso realizado correctamente.", vbExclamation
                    cmdCancel_Click (Index)
                End If
            End If
        
        Case 1
            If DatosOK Then
                If CopiaRegistros Then
                End If
            End If
    End Select
End Sub

Private Sub cmdCancel_Click(Index As Integer)
    Unload Me
End Sub


Private Sub Combo1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        Combo1(0).ListIndex = 0
        Select Case OpcionGenerar
            Case 0
                PonerFoco txtCodigo(0)
            Case 1
                PonerFoco txtCodigo(8)
        End Select
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

    For H = 0 To imgBuscar.Count - 1
        Me.imgBuscar(H).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next H
    
    tabla = "rcalidad"
    
    CargaCombo
    
    Me.FrameGeneraPreciosMasiva.visible = False
    Me.FrameCopiaVariedad.visible = False
    
    Select Case OpcionGenerar
        Case 0
            H = FrameGeneraPreciosMasiva.Height
            W = FrameGeneraPreciosMasiva.Width
            PonerFrameVisible FrameGeneraPreciosMasiva, True, H, W
        Case 1
            H = FrameCopiaVariedad.Height
            W = FrameCopiaVariedad.Width
            PonerFrameVisible FrameCopiaVariedad, True, H, W
    End Select
    
    
    'Esto se consigue poneinedo el cancel en el opcion k corresponda
'    Me.cmdCancel(indFrame).Cancel = True
    Me.Width = W + 70
    Me.Height = H + 350
End Sub


Private Sub frmC_Selec(vFecha As Date)
    ' *** repasar si el camp es txtAux o Text1 ***
    txtCodigo(CByte(imgFec(0).Tag)).Text = Format(vFecha, "dd/mm/yyyy") '<===
    ' ********************************************
End Sub

Private Sub frmCal_DatoSeleccionado(CadenaSeleccion As String)
'Form de Consulta de calidades
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "00")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
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

Private Sub frmVar_DatoSeleccionado(CadenaSeleccion As String)
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub imgBuscar_Click(Index As Integer)
   Select Case Index
        Case 0 'VARIEDADES
            AbrirFrmVariedad (Index)
    
        Case 1, 2 'VARIEDADES
            AbrirFrmVariedad (Index + 7)
        
        
    
    End Select
    PonerFoco txtCodigo(indCodigo)
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
            indCodigo = Index + 1
        Case 2, 3
            indCodigo = Index + 4
    End Select
    


    imgFec(0).Tag = indCodigo 'Index '<===
    ' *** repasar si el camp es txtAux o Text1 ***
    If txtCodigo(indCodigo).Text <> "" Then frmC.NovaData = txtCodigo(indCodigo).Text
    ' ********************************************

    frmC.Show vbModal
    Set frmC = Nothing
    ' *** repasar si el camp es txtAux o Text1 ***
    PonerFoco txtCodigo(CByte(imgFec(0).Tag)) '<===
    ' ********************************************

End Sub

Private Sub txtCodigo_GotFocus(Index As Integer)
    ConseguirFoco txtCodigo(Index), 3
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
            Case 0: KEYBusqueda KeyAscii, 0 'variedad desde
        End Select
    Else
        KEYpress KeyAscii
    End If

End Sub

Private Sub KEYBusqueda(KeyAscii As Integer, Indice As Integer)
    KeyAscii = 0
    imgBuscar_Click (Indice)
End Sub


Private Sub txtCodigo_LostFocus(Index As Integer)
Dim cad As String, cadTipo As String 'tipo cliente

    'Quitar espacios en blanco por los lados
    txtCodigo(Index).Text = Trim(txtCodigo(Index).Text)
'    If txtCodigo(Index).Text = "" Then Exit Sub
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    Select Case Index
        Case 0, 8, 9 'VARIEDADES
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "variedades", "nomvarie", "codvarie", "N")
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "000000")
            
        Case 1, 2, 6, 7 'FECHAS
            If txtCodigo(Index).Text <> "" Then PonerFormatoFecha txtCodigo(Index)
        
        Case 3, 4 'PRECIOS
            PonerFormatoDecimal txtCodigo(Index), 8
        
    End Select
End Sub


Private Sub InicializarVbles()
    cadFormula = ""
    cadSelect = ""
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
        .Opcion = 0
        .Show vbModal
    End With
End Sub

Private Sub AbrirFrmVariedad(Indice As Integer)
    indCodigo = Indice
    Set frmVar = New frmComVar
    frmVar.DatosADevolverBusqueda = "0|1|"
'    frmCli.DeConsulta = True
'    frmCli.CodigoActual = txtCodigo(indCodigo)
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


Private Function DatosOK() As Boolean
Dim B As Boolean
Dim SQL As String
Dim Sql2 As String
Dim vClien As cSocio
' añadido
Dim Mens As String
Dim numfactu As String
Dim numser As String
Dim Fecha As Date
Dim cad As String

    B = True
    
    Select Case OpcionGenerar
        Case 0
            If txtCodigo(0).Text = "" Then
                MsgBox "Debe introducir la variedad", vbExclamation
                B = False
            Else
                SQL = DevuelveDesdeBDNew(cAgro, "variedades", "nomvarie", "codvarie", txtCodigo(0).Text, "N")
                If SQL = "" Then
                    MsgBox "No existe la variedad. Reintroduzca.", vbExclamation
                    B = False
                    PonerFoco txtCodigo(0)
                End If
            End If
            
            If B Then
                If (txtCodigo(1).Text = "" Or txtCodigo(2).Text = "") Then
                    MsgBox "El rango de fechas debe de tener un valor. Reintroduzca.", vbExclamation
                    B = False
                Else
                    B = ComprobacionRangoFechas(txtCodigo(0).Text, CStr(Combo1(0).ListIndex), "", txtCodigo(1).Text, txtCodigo(2).Text)
                    If Not B Then
                        MsgBox "Este rango de fechas se solapa con otro registro. Revise.", vbExclamation
                        B = False
                    End If
                End If
            End If
            
            If B Then
                If txtCodigo(3).Text = "" Or txtCodigo(4).Text = "" Then
                    cad = "El valor de los precios esta vacio." & vbCrLf & vbCrLf & "¿ Desea continuar ?" & vbCrLf
                    If MsgBox(cad, vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then
                        B = False
                    End If
                End If
            End If
        
        Case 1
            If txtCodigo(8).Text = "" Then
                MsgBox "Debe introducir la variedad a copiar", vbExclamation
                B = False
            Else
                SQL = DevuelveDesdeBDNew(cAgro, "variedades", "nomvarie", "codvarie", txtCodigo(8).Text, "N")
                If SQL = "" Then
                    MsgBox "No existe la variedad. Reintroduzca.", vbExclamation
                    B = False
                    PonerFoco txtCodigo(8)
                End If
            End If
            If txtCodigo(9).Text = "" Then
                MsgBox "Debe introducir la variedad destino", vbExclamation
                B = False
            Else
                SQL = DevuelveDesdeBDNew(cAgro, "variedades", "nomvarie", "codvarie", txtCodigo(9).Text, "N")
                If SQL = "" Then
                    MsgBox "No existe la variedad. Reintroduzca.", vbExclamation
                    B = False
                    PonerFoco txtCodigo(9)
                End If
            End If
            
            If B Then
                If (txtCodigo(6).Text = "" Or txtCodigo(7).Text = "") Then
                    MsgBox "El rango de fechas debe de tener un valor. Reintroduzca.", vbExclamation
                    B = False
                Else
'                    b = ComprobacionRangoFechas(txtCodigo(9).Text, CStr(Combo1(1).ListIndex), "", txtCodigo(6).Text, txtCodigo(7).Text)
'                    If Not b Then
'                        MsgBox "Este rango de fechas se solapa con otro registro. Revise.", vbExclamation
'                        b = False
'                    End If
                End If
            End If
    End Select
        
    DatosOK = B

End Function



' ********* si n'hi han combos a la capçalera ************
Private Sub CargaCombo()
Dim Ini As Integer
Dim Fin As Integer
Dim I As Integer

    ' *** neteje els combos, els pose valor i seleccione el valor per defecte ***
    For I = 0 To Combo1.Count - 1
        Combo1(I).Clear
    Next I
    
    'tipo de factura
    Combo1(0).AddItem "Anticipo"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 0
    Combo1(0).AddItem "Liquidacion"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 1
    
    'tipo de factura
    Combo1(1).AddItem "Anticipo"
    Combo1(1).ItemData(Combo1(1).NewIndex) = 0
    Combo1(1).AddItem "Liquidacion"
    Combo1(1).ItemData(Combo1(1).NewIndex) = 1
    
End Sub


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



Public Function GeneraRegistros() As Boolean
'Insertar en tablas cabecera/lineas de la Contabilidad la factura
' ariges.scafac --> conta.cabfact
' ariges.slifac --> conta.linfact
'Actualizar la tabla ariges.scafac.inconta=1 para indicar que ya esta contabilizada
Dim B As Boolean
Dim cadMen As String
Dim SQL As String
Dim NumF As Currency

    On Error GoTo EContab

    conn.BeginTrans
    
    'Obtenemos el siguiente contador para esta variedad tipo
    SQL = "select max(contador) from rprecios where codvarie = " & DBSet(txtCodigo(0).Text, "N")
    SQL = SQL & " and tipofact = " & DBSet(Combo1(0).ListIndex, "N")
    
    NumF = TotalRegistros(SQL) + 1
    
    'Insertar en rprecios
    B = InsertarCabecera(cadMen, CStr(NumF))
    cadMen = "Insertando Cabecera de Precios: " & cadMen
    
    If B Then
        'Insertar lineas rprecios_calidad
        B = InsertarLineas(cadMen, CStr(NumF))
        cadMen = "Insertando Lineas de Precios: " & cadMen

    End If
    
EContab:
    
    If Err.Number <> 0 Then
        B = False
        MuestraError Err.Number, "Generando Registros", Err.Description
    End If
    If B Then
        conn.CommitTrans
        GeneraRegistros = True
    Else
        conn.RollbackTrans
        GeneraRegistros = False
    End If
End Function

Public Function CopiaRegistros() As Boolean
'Insertar en tablas cabecera/lineas de la Contabilidad la factura
' ariges.scafac --> conta.cabfact
' ariges.slifac --> conta.linfact
'Actualizar la tabla ariges.scafac.inconta=1 para indicar que ya esta contabilizada
Dim B As Boolean
Dim cadMen As String
Dim SQL As String
Dim NumF As Currency
Dim NumFF As Currency


    On Error GoTo EContab

    conn.BeginTrans
    
    B = True
    
    SQL = "select max(contador) from rprecios where codvarie = " & DBSet(txtCodigo(8).Text, "N")
    SQL = SQL & " and tipofact = " & DBSet(Combo1(1).ListIndex, "N")
    SQL = SQL & " and fechaini = " & DBSet(txtCodigo(6).Text, "F")
    SQL = SQL & " and fechafin = " & DBSet(txtCodigo(7).Text, "F")
    
    NumFF = DevuelveValor(SQL)
    If NumFF = 0 Then
        MsgBox "No existe precio de la variedad origen. Revise.", vbExclamation
        conn.RollbackTrans
        Exit Function
    End If
    
    
    'Obtenemos el siguiente contador para esta variedad tipo
    SQL = "select max(contador) from rprecios where codvarie = " & DBSet(txtCodigo(9).Text, "N")
    SQL = SQL & " and tipofact = " & DBSet(Combo1(1).ListIndex, "N")
    SQL = SQL & " and fechaini = " & DBSet(txtCodigo(6).Text, "F")
    SQL = SQL & " and fechafin = " & DBSet(txtCodigo(7).Text, "F")
    
    NumF = DevuelveValor(SQL)
    
    If NumF = 0 Then
        ' comprobamos que no exista otro rango de fechas en el que se solape
        B = ComprobacionRangoFechas(txtCodigo(9).Text, CStr(Combo1(1).ListIndex), "", txtCodigo(6).Text, txtCodigo(7).Text)
        If Not B Then
            MsgBox "Este rango de fechas se solapa con otro registro. Revise.", vbExclamation
            conn.RollbackTrans
            Exit Function
        End If
        
        'Obtenemos el siguiente contador para esta variedad tipo
        SQL = "select max(contador) from rprecios where codvarie = " & DBSet(txtCodigo(9).Text, "N")
        SQL = SQL & " and tipofact = " & DBSet(Combo1(1).ListIndex, "N")
        
        NumF = DevuelveValor(SQL)
        
        'Insertar en rprecios
        SQL = "INSERT INTO rprecios (codvarie, tipofact, contador, fechaini, fechafin, textoper, precioindustria) "
        SQL = SQL & " select " & DBSet(txtCodigo(9).Text, "N") & ", tipofact, " & DBSet(NumF + 1, "N") & ",fechaini, fechafin, textoper, precioindustria "
        SQL = SQL & " from rprecios "
        SQL = SQL & " where codvarie = " & DBSet(txtCodigo(8).Text, "N") & " and tipofact = " & DBSet(Combo1(1).ListIndex, "N")
        SQL = SQL & " and contador = " & DBSet(NumFF, "N")
        
        conn.Execute SQL
        
        SQL = "insert into rprecios_calidad (codvarie, tipofact, contador, codcalid, precoop, presocio) "
        SQL = SQL & " select " & DBSet(txtCodigo(9).Text, "N") & ", tipofact, " & DBSet(NumF + 1, "N") & ", codcalid, precoop, presocio "
        SQL = SQL & " from rprecios_calidad "
        SQL = SQL & " where codvarie = " & DBSet(txtCodigo(8).Text, "N")
        SQL = SQL & " and tipofact = " & DBSet(Combo1(1).ListIndex, "N")
        SQL = SQL & " and contador = " & DBSet(NumFF, "N")
        
        conn.Execute SQL
    
    Else
        If NumF >= 1 Then
            SQL = "¿ Ya existe un registro para esta variedad en ese período, quiere actualizarlo ?"
            SQL = SQL & vbCrLf & "Si elije Sí, actualiza el último precio. "
            SQL = SQL & vbCrLf & "Si elije No, crea uno nuevo. "
            SQL = SQL & vbCrLf & "Si elije Cancelar, no hace nada. "
            
            Select Case MsgBox(SQL, vbQuestion + vbYesNoCancel)
                Case vbYes
                    
                    SQL = "delete from rprecios_calidad where codvarie = " & DBSet(txtCodigo(9).Text, "N")
                    SQL = SQL & " and tipofact = " & DBSet(Combo1(1).ListIndex, "N")
                    SQL = SQL & " and contador = " & DBSet(NumF, "N")
                    
                    conn.Execute SQL
                    
                    SQL = "insert into rprecios_calidad (codvarie, tipofact, contador, codcalid, precoop, presocio) "
                    SQL = SQL & " select " & DBSet(txtCodigo(9).Text, "N") & ", tipofact, " & DBSet(NumF, "N") & ", codcalid, precoop, presocio "
                    SQL = SQL & " from rprecios_calidad "
                    SQL = SQL & " where codvarie = " & DBSet(txtCodigo(8).Text, "N")
                    SQL = SQL & " and tipofact = " & DBSet(Combo1(1).ListIndex, "N")
                    SQL = SQL & " and contador = " & DBSet(NumFF, "N")
                    
                    conn.Execute SQL
                    
                Case vbNo
                    
                    'Insertar en rprecios
                    
                    'Obtenemos el siguiente contador para esta variedad tipo
                    SQL = "select max(contador) from rprecios where codvarie = " & DBSet(txtCodigo(9).Text, "N")
                    SQL = SQL & " and tipofact = " & DBSet(Combo1(1).ListIndex, "N")
                    
                    NumF = DevuelveValor(SQL)
                    
                    SQL = "INSERT INTO rprecios (codvarie, tipofact, contador, fechaini, fechafin, textoper, precioindustria) "
                    SQL = SQL & " select " & DBSet(txtCodigo(9).Text, "N") & ", tipofact, " & DBSet(NumF + 1, "N") & ",fechaini, fechafin, textoper, precioindustria "
                    SQL = SQL & " from rprecios "
                    SQL = SQL & " where codvarie = " & DBSet(txtCodigo(8).Text, "N") & " and tipofact = " & DBSet(Combo1(1).ListIndex, "N")
                    SQL = SQL & " and contador = " & DBSet(NumFF, "N")
                    
                    conn.Execute SQL
                    
                    SQL = "insert into rprecios_calidad (codvarie, tipofact, contador, codcalid, precoop, presocio) "
                    SQL = SQL & " select " & DBSet(txtCodigo(9).Text, "N") & ", tipofact, " & DBSet(NumF + 1, "N") & ", codcalid, precoop, presocio "
                    SQL = SQL & " from rprecios_calidad "
                    SQL = SQL & " where codvarie = " & DBSet(txtCodigo(8).Text, "N")
                    SQL = SQL & " and tipofact = " & DBSet(Combo1(1).ListIndex, "N")
                    SQL = SQL & " and contador = " & DBSet(NumFF, "N")
                    
                    conn.Execute SQL
                
                Case vbCancel
                    conn.RollbackTrans
                    Exit Function
            End Select
       End If
    End If
    
EContab:
    If Err.Number <> 0 Then
        B = False
        MuestraError Err.Number, "Generando Registros", Err.Description
    End If
    If B Then
        conn.CommitTrans
        CopiaRegistros = True
        MsgBox "Proceso realizado correctamente.", vbExclamation
        cmdCancel_Click (1)
    Else
        conn.RollbackTrans
        CopiaRegistros = False
    End If
End Function



Private Function InsertarCabecera(cadErr As String, Contador As String) As Boolean
'Insertando en tabla conta.cabfact
Dim SQL As String
Dim Rs As ADODB.Recordset
Dim cad As String
Dim NumF As Currency

    On Error GoTo EInsertar
    
    
    'Insertar en rprecios
    SQL = "INSERT INTO rprecios (codvarie, tipofact, contador, fechaini, fechafin, textoper) values ("
    SQL = SQL & DBSet(txtCodigo(0).Text, "N") & ","
    SQL = SQL & DBSet(Combo1(0).ListIndex, "N") & ","
    SQL = SQL & DBSet(Contador, "N") & ","
    SQL = SQL & DBSet(txtCodigo(1).Text, "F") & ","
    SQL = SQL & DBSet(txtCodigo(2).Text, "F") & ","
    SQL = SQL & DBSet(txtCodigo(5).Text, "T") & ")"
    
    conn.Execute SQL
    
EInsertar:
    If Err.Number <> 0 Then
        InsertarCabecera = False
        cadErr = Err.Description
    Else
        InsertarCabecera = True
    End If
End Function

Private Function InsertarLineas(cadErr As String, Contador As String) As Boolean
Dim SQL As String
Dim cad As String
Dim Cad1 As String
Dim Rs As ADODB.Recordset

    On Error GoTo EInLinea

    SQL = "select codcalid from rcalidad where codvarie = " & DBSet(txtCodigo(0).Text, "N")
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    cad = "insert into rprecios_calidad (codvarie, tipofact, contador, codcalid, precoop, presocio) values "
    
    Cad1 = ""
    
    While Not Rs.EOF
        Cad1 = Cad1 & "(" & DBSet(txtCodigo(0).Text, "N") & "," & DBSet(Combo1(0).ListIndex, "N") & ","
        Cad1 = Cad1 & DBSet(Contador, "N") & ","
        Cad1 = Cad1 & DBSet(Rs.Fields(0).Value, "N") & ","
        Cad1 = Cad1 & DBSet(txtCodigo(3).Text, "N") & ","
        Cad1 = Cad1 & DBSet(txtCodigo(4).Text, "N") & "),"
        
        Rs.MoveNext
    Wend
    Set Rs = Nothing
    
    If Cad1 <> "" Then
        ' quitamos la ultima coma
        Cad1 = Mid(Cad1, 1, Len(Cad1) - 1)
        ' concatenamos con el insert
        cad = cad & Cad1
        
        conn.Execute cad
    End If
    
EInLinea:
    If Err.Number <> 0 Then
        InsertarLineas = False
        cadErr = Err.Description
    Else
        InsertarLineas = True
    End If
    
End Function

