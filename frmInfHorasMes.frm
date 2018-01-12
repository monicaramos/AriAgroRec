VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmInfHorasMes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informes"
   ClientHeight    =   7485
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   6900
   Icon            =   "frmInfHorasMes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7485
   ScaleWidth      =   6900
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameHorasTrabajadas 
      Height          =   7410
      Left            =   45
      TabIndex        =   11
      Top             =   0
      Width           =   6795
      Begin VB.CheckBox Check1 
         Caption         =   "Detalle de Costes"
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
         Index           =   2
         Left            =   420
         TabIndex        =   34
         Top             =   5685
         Width           =   2130
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
         Left            =   1830
         MaxLength       =   3
         TabIndex        =   4
         Top             =   3120
         Width           =   930
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
         Left            =   1830
         MaxLength       =   3
         TabIndex        =   3
         Top             =   2730
         Width           =   930
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
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
         Left            =   2790
         Locked          =   -1  'True
         TabIndex        =   30
         Text            =   "Text5"
         Top             =   3120
         Width           =   3195
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
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
         Left            =   2790
         Locked          =   -1  'True
         TabIndex        =   29
         Text            =   "Text5"
         Top             =   2730
         Width           =   3195
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Sobre Horas Productivas"
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
         Left            =   420
         TabIndex        =   28
         Top             =   6285
         Width           =   2895
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Imprimir sólo Líquido"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Index           =   0
         Left            =   420
         TabIndex        =   27
         Top             =   5925
         Width           =   2310
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
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
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   23
         Text            =   "Text5"
         Top             =   4140
         Width           =   3195
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
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
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   22
         Text            =   "Text5"
         Top             =   3780
         Width           =   3195
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
         Left            =   1800
         MaxLength       =   3
         TabIndex        =   6
         Top             =   4140
         Width           =   930
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
         Left            =   1800
         MaxLength       =   3
         TabIndex        =   5
         Top             =   3780
         Width           =   930
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
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Tag             =   "Tipo|N|N|||straba|codsecci||N|"
         Top             =   1035
         Width           =   1710
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
         Left            =   5145
         TabIndex        =   10
         Top             =   6570
         Width           =   1065
      End
      Begin VB.CommandButton CmdAceptar 
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
         Left            =   3930
         TabIndex        =   9
         Top             =   6555
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
         Index           =   19
         Left            =   1845
         MaxLength       =   6
         TabIndex        =   2
         Top             =   2115
         Width           =   930
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
         Index           =   18
         Left            =   1845
         MaxLength       =   6
         TabIndex        =   1
         Top             =   1755
         Width           =   930
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
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
         Index           =   18
         Left            =   2805
         Locked          =   -1  'True
         TabIndex        =   13
         Text            =   "Text5"
         Top             =   1755
         Width           =   3195
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
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
         Index           =   19
         Left            =   2805
         Locked          =   -1  'True
         TabIndex        =   12
         Text            =   "Text5"
         Top             =   2130
         Width           =   3195
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
         Index           =   17
         Left            =   1785
         MaxLength       =   10
         TabIndex        =   8
         Top             =   5205
         Width           =   1305
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
         Index           =   16
         Left            =   1785
         MaxLength       =   10
         TabIndex        =   7
         Text            =   "00/00/0000"
         Top             =   4800
         Width           =   1305
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
         Index           =   5
         Left            =   735
         TabIndex        =   33
         Top             =   2760
         Width           =   735
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
         Index           =   4
         Left            =   735
         TabIndex        =   32
         Top             =   3120
         Width           =   690
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Almacén"
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
         TabIndex        =   31
         Top             =   2490
         Width           =   825
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   3
         Left            =   1515
         MouseIcon       =   "frmInfHorasMes.frx":000C
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar almacén"
         Top             =   3120
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   2
         Left            =   1515
         MouseIcon       =   "frmInfHorasMes.frx":015E
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar almacén"
         Top             =   2730
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   1
         Left            =   1485
         MouseIcon       =   "frmInfHorasMes.frx":02B0
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar forma de pago"
         Top             =   4140
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   1485
         MouseIcon       =   "frmInfHorasMes.frx":0402
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar forma de pago"
         Top             =   3780
         Width           =   240
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Forma de Pago"
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
         Left            =   450
         TabIndex        =   26
         Top             =   3510
         Width           =   1470
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
         Index           =   1
         Left            =   735
         TabIndex        =   25
         Top             =   4140
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
         Index           =   0
         Left            =   735
         TabIndex        =   24
         Top             =   3780
         Width           =   735
      End
      Begin VB.Label Label8 
         Caption         =   "Sección "
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
         Height          =   255
         Left            =   450
         TabIndex        =   21
         Top             =   1035
         Width           =   885
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
         Index           =   29
         Left            =   735
         TabIndex        =   20
         Top             =   1770
         Width           =   735
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
         Index           =   28
         Left            =   735
         TabIndex        =   19
         Top             =   2130
         Width           =   690
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Trabajador"
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
         Index           =   27
         Left            =   450
         TabIndex        =   18
         Top             =   1485
         Width           =   1065
      End
      Begin VB.Label Label7 
         Caption         =   "Informe de Horas Mensual"
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
         Left            =   405
         TabIndex        =   17
         Top             =   405
         Width           =   5925
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
         Index           =   26
         Left            =   735
         TabIndex        =   16
         Top             =   4815
         Width           =   735
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
         Index           =   25
         Left            =   735
         TabIndex        =   15
         Top             =   5175
         Width           =   690
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Recibo"
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
         Left            =   450
         TabIndex        =   14
         Top             =   4530
         Width           =   1320
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   15
         Left            =   1530
         MouseIcon       =   "frmInfHorasMes.frx":0554
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar trabajador"
         Top             =   2115
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   14
         Left            =   1530
         MouseIcon       =   "frmInfHorasMes.frx":06A6
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar trabajador"
         Top             =   1755
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   3
         Left            =   1470
         Picture         =   "frmInfHorasMes.frx":07F8
         Top             =   5205
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   2
         Left            =   1470
         Picture         =   "frmInfHorasMes.frx":0883
         Top             =   4800
         Width           =   240
      End
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   7680
      Top             =   5160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmInfHorasMes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' +-+-+-+-+-+-+-+-+-+-+-
' +-+- Autor: LAURA +-+-
' +-+-+-+-+-+-+-+-+-+-+-

Option Explicit

Public OpcionListado As Byte
    '==== Listados BASICOS ====
    '=============================
    ' 15 .- Listado de Horas trababajadas
    
Public NumCod As String 'Para indicar cod. Traspaso,Movimiento, etc. que llama
                        'Para indicar nº oferta a imprimir

Public CadTag As String 'Cadena con el Tag del campo que se va a poner en D/H en los listados
                        'Se necesita si el tipo de codigo es texto

Public Event RectificarFactura(Cliente As String, Observaciones As String)

Private Conexion As Byte
'1.- Conexión a BD Ariges  2.- Conexión a BD Conta

Private HaDevueltoDatos As Boolean


Private WithEvents frmTra As frmManTraba 'mantenimiento de trabajadores
Attribute frmTra.VB_VarHelpID = -1
Private WithEvents frmFPa As frmBasico2 'mantenimiento de formas de pago de comercial
Attribute frmFPa.VB_VarHelpID = -1
Private WithEvents frmAlm As frmBasico2 'mantenimiento de almacenes propios
Attribute frmAlm.VB_VarHelpID = -1

Private WithEvents frmB As frmBuscaGrid  'Busquedas
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmC As frmCal 'calendario fechas
Attribute frmC.VB_VarHelpID = -1

'GENERALES PARA PASARLE A CRYSTAL REPORT
Private cadFormula As String 'Cadena con la FormulaSelection para Crystal Report
Private CadParam As String 'Cadena con los parametros para Crystal Report
Private numParam As Byte 'Numero de parametros que se pasan a Crystal Report
Private cadSelect As String 'Cadena para comprobar si hay datos antes de abrir Informe
Private cadTitulo As String 'Titulo para la ventana frmImprimir
Private cadNombreRPT As String 'Nombre del informe
Private ConSubInforme As Boolean

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

Private Sub Check1_Click(Index As Integer)
Dim I As Integer
    Select Case Index
        Case 2
            For I = 0 To 1
                Check1(I).Enabled = (Check1(2).Value = 0)
            Next I
'[Monica]05/01/2012: SegSoc pasa a ser porcentaje daba error y lo he comentado
'            Check1(3).Enabled = (Check1(2).Value = 0)
        Case 3
            For I = 0 To 2
                Check1(I).Enabled = (Check1(3).Value = 0)
            Next I
    End Select
End Sub

Private Sub cmdAceptar_Click(Index As Integer)
Dim cOrden As String
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim numOp As Byte
    
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
    
    
    InicializarVbles
    
    'Añadir el parametro de Empresa
    CadParam = CadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    Select Case Index
        Case 0 ' impresion de recibos
            '======== FORMULA  ====================================
            'D/H TRABAJADOR
            cDesde = Trim(txtCodigo(18).Text)
            cHasta = Trim(txtCodigo(19).Text)
            nDesde = txtNombre(18).Text
            nHasta = txtNombre(19).Text
            If Not (cDesde = "" And cHasta = "") Then
                'Cadena para seleccion Desde y Hasta
                Codigo = "{horas.codtraba}"
                TipCod = "N"
                If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHTrabajador=""") Then Exit Sub
            End If
             
             'D/H ALMACEN
            cDesde = Trim(txtCodigo(2).Text)
            cHasta = Trim(txtCodigo(3).Text)
            nDesde = txtNombre(2).Text
            nHasta = txtNombre(3).Text
            If Not (cDesde = "" And cHasta = "") Then
                'Cadena para seleccion Desde y Hasta
                Codigo = "{horas.codalmac}"
                TipCod = "N"
                If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHAlmacen=""") Then Exit Sub
            End If
           
            'D/H fecha
            cDesde = Trim(txtCodigo(16).Text)
            cHasta = Trim(txtCodigo(17).Text)
            If Not (cDesde = "" And cHasta = "") Then
                'Cadena para seleccion Desde y Hasta
                Codigo = "{horas.fecharec}"
                TipCod = "F"
                If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHFecha=""") Then Exit Sub
            End If
            
            'D/H FORMA DE PAGO DEL TRABAJADOR
            cDesde = Trim(txtCodigo(0).Text)
            cHasta = Trim(txtCodigo(1).Text)
            nDesde = txtNombre(0).Text
            nHasta = txtNombre(1).Text
            If Not (cDesde = "" And cHasta = "") Then
                'Cadena para seleccion Desde y Hasta
                Codigo = "{straba.codforpa}"
                TipCod = "N"
                If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHFPago=""") Then Exit Sub
            End If
                       
            'Tipo de seccion
            AnyadirAFormula cadFormula, "{straba.codsecci} = " & Me.Combo1(1).ListIndex
            AnyadirAFormula cadSelect, "{straba.codsecci} = " & Me.Combo1(1).ListIndex
            
            ' La fecha de recibo no puede estar vacia
            AnyadirAFormula cadFormula, "not isnull({horas.fecharec})"
            AnyadirAFormula cadSelect, "not isnull({horas.fecharec})"
            
            CadParam = CadParam & "pSeccion=""" & Me.Combo1(1).Text & """|"
            numParam = numParam + 1
            '[Monica]29/02/2012: Natural era la cooperativa 0 junto con Catadau ahora es la 9
            If vParamAplic.Cooperativa = 0 Or vParamAplic.Cooperativa = 9 Or vParamAplic.Cooperativa = 2 Or vParamAplic.Cooperativa = 16 Then
                indRPT = 55 'Informe de horas mensual
                If Not PonerParamRPT(indRPT, CadParam, numParam, nomDocu) Then Exit Sub
                  
                'Nombre fichero .rpt a Imprimir
                cadNombreRPT = nomDocu
                
                frmImprimir.NombreRPT = cadNombreRPT '"rInfHorasMensNatural.rpt"
                tabla = "(rrecibosnomina INNER JOIN straba ON rrecibosnomina.codtraba = straba.codtraba) "
                
                cadSelect = Replace(cadSelect, "horas.fecharec", "rrecibosnomina.fechahora")
                cadFormula = Replace(cadFormula, "horas.fecharec", "rrecibosnomina.fechahora")
                
                cadSelect = Replace(cadSelect, "horas.codtraba", "rrecibosnomina.codtraba")
                cadFormula = Replace(cadFormula, "horas.codtraba", "rrecibosnomina.codtraba")
                
            Else
                If Check1(2).Value = 0 Then
                
                    ' sobre horas productivas
                    If Check1(1).Value Then ' se imprime el resumen
                        CadParam = CadParam & "pHProductivas=1|"
                    Else
                        CadParam = CadParam & "pHProductivas=0|"
                    End If
                    numParam = numParam + 1
                    
                    
                    tabla = "(horas INNER JOIN straba ON horas.codtraba = straba.codtraba) "
                               
                    If Me.Check1(0).Value = 0 Then
                        frmImprimir.NombreRPT = "rInfHorasMens.rpt"
                    Else
                        frmImprimir.NombreRPT = "rInfHorasMens1.rpt"
                    End If
                    
                    ConSubInforme = False
                    
                Else
                    ' detallamos los costes de los partes
                    tabla = "(horas INNER JOIN straba ON horas.codtraba = straba.codtraba) "
                    tabla = "(" & tabla & ") INNER JOIN rpartes_trabajador ON horas.nroparte = rpartes_trabajador.nroparte and rpartes_trabajador.codtraba = horas.codtraba "
                    
                    frmImprimir.NombreRPT = "rInfHorasMensDetalle.rpt"
                
                    ConSubInforme = True
                
                End If
            End If
                
            cadTitulo = "Informe Horas Mensual"
    
    End Select
    
    'Comprobar si hay registros a Mostrar antes de abrir el Informe
    If HayRegParaInforme(tabla, cadSelect) Then
        LlamarImprimir
    End If

End Sub


Private Sub cmdCancel_Click()
    Unload Me
End Sub


Private Sub Combo1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        Select Case OpcionListado
            Case 15 ' Informe de Horas Trabajadas
                PonerFoco txtCodigo(18)
        End Select
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
Dim H As Integer, W As Integer
Dim List As Collection
Dim I As Integer

    'Icono del formulario
    Me.Icon = frmPpal.Icon

    PrimeraVez = True
    limpiar Me

    'IMAGES para busqueda
    Set List = New Collection
    For H = 24 To 27
        List.Add H
    Next H
    For H = 1 To 10
        List.Add H
    Next H
    List.Add 12
    List.Add 13
    List.Add 14
    List.Add 15
    List.Add 18
    List.Add 19
    
    
'    For h = 1 To List.Count
'        Me.imgBuscar(List.item(h)).Picture = frmPpal.imgListImages16.ListImages(1).Picture
'    Next h
' ### [Monica] 09/11/2006    he sustituido el anterior
    For H = 14 To 15 'imgBuscar.Count - 1
        Me.imgBuscar(H).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next H
    For H = 0 To 3 'imgBuscar.Count - 1
        Me.imgBuscar(H).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next H
     
    
    Set List = Nothing

    'Ocultar todos los Frames de Formulario
    Me.FrameHorasTrabajadas.visible = False
    
    '###Descomentar
'    CommitConexion
    H = 7410
    W = 6795
    FrameHorasTrabajadasVisible True, H, W
    indFrame = 0
    tabla = "horas"
        
    CargaCombo
        
    For I = 0 To Check1.Count - 1
        '[Monica]29/02/2012: Natural era la cooperativa 0 junto con Catadau ahora es la 9
        Check1(I).Enabled = (vParamAplic.Cooperativa <> 0 And vParamAplic.Cooperativa <> 9)
        Check1(I).visible = (vParamAplic.Cooperativa <> 0 And vParamAplic.Cooperativa <> 9)
    Next I
        
        
    'Esto se consigue poneinedo el cancel en el opcion k corresponda
    Me.cmdCancel.Cancel = True
    Me.Width = W + 70
    Me.Height = H + 350
    
    Me.Combo1(1).ListIndex = 1
End Sub


Private Sub frmAlm_DatoSeleccionado(CadenaSeleccion As String)
'Form de Consulta de Almacenes
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmC_Selec(vFecha As Date)
    ' *** repasar si el camp es txtAux o Text1 ***
    txtCodigo(CByte(imgFecha(2).Tag) + 14).Text = Format(vFecha, "dd/mm/yyyy") '<===
    ' ********************************************
End Sub

Private Sub frmCal_DatoSeleccionado(CadenaSeleccion As String)
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub


Private Sub frmFPa_DatoSeleccionado(CadenaSeleccion As String)
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmTra_DatoSeleccionado(CadenaSeleccion As String)
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub


Private Sub imgBuscar_Click(Index As Integer)
   Select Case Index
        Case 0, 1 'Forma de pago
            AbrirFrmManFPago (Index)
   
        
        Case 2, 3 'Almacenes
            AbrirFrmManAlm (Index)
        
        
        Case 14, 15 'Horas trabajadas
            AbrirFrmManTraba (Index)
    
    End Select
    PonerFoco txtCodigo(indCodigo)
End Sub


Private Sub imgFecha_Click(Index As Integer)
    Dim esq As Long
    Dim dalt As Long
    Dim menu As Long
    Dim obj As Object

    Set frmC = New frmCal
    
    esq = imgFecha(Index).Left
    dalt = imgFecha(Index).Top
        
    Set obj = imgFecha(Index).Container
      
      While imgFecha(Index).Parent.Name <> obj.Name
            esq = esq + obj.Left
            dalt = dalt + obj.Top
            Set obj = obj.Container
      Wend
    
    menu = Me.Height - Me.ScaleHeight 'ací tinc el heigth del menú i de la toolbar

    frmC.Left = esq + imgFecha(Index).Parent.Left + 30
    frmC.Top = dalt + imgFecha(Index).Parent.Top + imgFecha(Index).Height + menu - 40

    imgFecha(2).Tag = Index '<===
    ' *** repasar si el camp es txtAux o Text1 ***
    If txtCodigo(Index + 14).Text <> "" Then frmC.NovaData = txtCodigo(Index + 14).Text
    ' ********************************************

    frmC.Show vbModal
    Set frmC = Nothing
    ' *** repasar si el camp es txtAux o Text1 ***
    PonerFoco txtCodigo(CByte(imgFecha(2).Tag) + 14) '<===
    ' ********************************************
End Sub

Private Sub txtCodigo_GotFocus(Index As Integer)
    ConseguirFoco txtCodigo(Index), 3
End Sub

Private Sub txtCodigo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub txtCodigo_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = teclaBuscar Then
        Select Case Index
            Case 0: KEYBusqueda KeyAscii, 0 'forma de pago desde
            Case 1: KEYBusqueda KeyAscii, 1 'forma de pago hasta
            Case 2: KEYBusqueda KeyAscii, 2 'almacen desde
            Case 3: KEYBusqueda KeyAscii, 3 'almacen hasta
            Case 16: KEYFecha KeyAscii, 14 'fecha desde
            Case 17: KEYFecha KeyAscii, 15 'fecha hasta
            Case 18: KEYBusqueda KeyAscii, 14 'trabajador desde
            Case 19: KEYBusqueda KeyAscii, 15 'trabajador hasta

        End Select
    Else
        KEYpress KeyAscii
    End If

End Sub

Private Sub KEYFecha(KeyAscii As Integer, Indice As Integer)
    KeyAscii = 0
    imgFecha_Click (Indice)
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
        Case 0, 1 'FORMAS DE PAGO
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "forpago", "nomforpa", "codforpa", "N")
            
        Case 2, 3 'almacen
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "salmpr", "nomalmac", "codalmac", "N")
        
        Case 16, 17 'FECHAS
            If txtCodigo(Index).Text <> "" Then PonerFormatoFecha txtCodigo(Index)
            
        Case 18, 19 'TRABAJADORES
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "straba", "nomtraba", "codtraba", "N")
    End Select
End Sub

Private Sub FrameHorasTrabajadasVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Frame para el listado de horas trabajadas
    Me.FrameHorasTrabajadas.visible = visible
    If visible = True Then
        Me.FrameHorasTrabajadas.Top = -90
        Me.FrameHorasTrabajadas.Left = 0
        Me.FrameHorasTrabajadas.Height = 7410 '6870
        Me.FrameHorasTrabajadas.Width = 6795
        W = Me.FrameHorasTrabajadas.Width
        H = Me.FrameHorasTrabajadas.Height
    End If
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
        .Titulo = cadTitulo
        .EnvioEMail = False
'        .NombreRPT = cadNombreRPT
        .ConSubInforme = ConSubInforme
        .Opcion = OpcionListado
        .Show vbModal
    End With
End Sub

Private Sub AbrirFrmManTraba(Indice As Integer)
    indCodigo = Indice + 4
    Set frmTra = New frmManTraba
    frmTra.DatosADevolverBusqueda = "0|2|"
    frmTra.Show vbModal
    Set frmTra = Nothing
End Sub

Private Sub AbrirFrmManFPago(Indice As Integer)

    Set frmFPa = New frmBasico2
    
    AyudaFPagoCom frmFPa, txtCodigo(Indice)
    
    Set frmFPa = Nothing
    
    PonerFoco txtCodigo(Indice)
End Sub

Private Sub AbrirFrmManAlm(Indice As Integer)
    indCodigo = Indice
    
    Set frmAlm = New frmBasico2
    
    AyudaAlmacenCom frmAlm, txtCodigo(Indice)
    
    Set frmAlm = Nothing
    
    PonerFoco txtCodigo(Indice)
End Sub


' ********* si n'hi han combos a la capçalera ************
Private Sub CargaCombo()
Dim Ini As Integer
Dim Fin As Integer
Dim I As Integer

    ' *** neteje els combos, els pose valor i seleccione el valor per defecte ***
'    For I = 0 To Combo1.Count - 1
'        Combo1(I).Clear
'    Next I

    Combo1(1).Clear
    
    Combo1(1).AddItem "Campo"
    Combo1(1).ItemData(Combo1(1).NewIndex) = 0
    Combo1(1).AddItem "Almacén"
    Combo1(1).ItemData(Combo1(1).NewIndex) = 1
    
End Sub

