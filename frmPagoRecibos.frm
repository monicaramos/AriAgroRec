VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmPagoRecibos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informes"
   ClientHeight    =   6990
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   6510
   Icon            =   "frmPagoRecibos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6990
   ScaleWidth      =   6510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cd1 
      Left            =   7680
      Top             =   5160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame FramePagoRecibosNatural 
      Height          =   6930
      Left            =   60
      TabIndex        =   17
      Top             =   0
      Width           =   6435
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
         Left            =   2040
         MaxLength       =   6
         TabIndex        =   35
         Top             =   5190
         Width           =   1005
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
         Index           =   5
         Left            =   2040
         MaxLength       =   6
         TabIndex        =   33
         Top             =   4770
         Width           =   1005
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
         Left            =   2040
         MaxLength       =   6
         TabIndex        =   31
         Top             =   4350
         Width           =   1005
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
         Left            =   2040
         MaxLength       =   6
         TabIndex        =   29
         Top             =   3960
         Width           =   1005
      End
      Begin VB.CommandButton CmdAceptarNat 
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
         Left            =   3735
         TabIndex        =   37
         Top             =   6270
         Width           =   1065
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
         Left            =   4920
         TabIndex        =   39
         Top             =   6270
         Width           =   1065
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
         Index           =   3
         Left            =   2040
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Tag             =   "Tipo|N|N|||straba|codsecci||N|"
         Top             =   1035
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
         Index           =   2
         Left            =   2040
         MaxLength       =   10
         TabIndex        =   23
         Top             =   2340
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
         Left            =   2040
         MaxLength       =   10
         TabIndex        =   21
         Top             =   1800
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
         Index           =   0
         Left            =   2040
         MaxLength       =   6
         TabIndex        =   25
         Top             =   2790
         Width           =   870
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
         Left            =   2925
         Locked          =   -1  'True
         TabIndex        =   18
         Text            =   "Text5"
         Top             =   2790
         Width           =   3150
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
         Index           =   2
         Left            =   2040
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Tag             =   "Tipo|N|N|||straba|codsecci||N|"
         Top             =   3465
         Width           =   1665
      End
      Begin MSComctlLib.ProgressBar Pb2 
         Height          =   240
         Left            =   480
         TabIndex        =   20
         Top             =   5940
         Width           =   5550
         _ExtentX        =   9790
         _ExtentY        =   423
         _Version        =   393216
         Appearance      =   1
      End
      Begin MSComDlg.CommonDialog CommonDialog2 
         Left            =   90
         Top             =   3915
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "% IRPF "
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
         Index           =   6
         Left            =   510
         TabIndex        =   40
         Top             =   5190
         Width           =   780
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "% Seg.Social 2"
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
         Left            =   510
         TabIndex        =   38
         Top             =   4800
         Width           =   1500
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "% Seg.Social 1"
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
         Left            =   510
         TabIndex        =   36
         Top             =   4380
         Width           =   1500
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "% Jornada"
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
         Left            =   510
         TabIndex        =   34
         Top             =   3990
         Width           =   1035
      End
      Begin VB.Label Label5 
         Caption         =   "Pago de Recibos"
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
         Left            =   480
         TabIndex        =   32
         Top             =   420
         Width           =   5775
      End
      Begin VB.Label Label4 
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
         Left            =   510
         TabIndex        =   30
         Top             =   1035
         Width           =   975
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   1
         Left            =   1740
         Picture         =   "frmPagoRecibos.frx":000C
         ToolTipText     =   "Buscar fecha"
         Top             =   2340
         Width           =   240
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Pago"
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
         Left            =   510
         TabIndex        =   28
         Top             =   2070
         Width           =   1155
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
         Index           =   1
         Left            =   510
         TabIndex        =   26
         Top             =   1530
         Width           =   1320
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   0
         Left            =   1740
         Picture         =   "frmPagoRecibos.frx":0097
         ToolTipText     =   "Buscar fecha"
         Top             =   1800
         Width           =   240
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Banco "
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
         Left            =   510
         TabIndex        =   24
         Top             =   2745
         Width           =   675
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   1740
         MouseIcon       =   "frmPagoRecibos.frx":0122
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar banco"
         Top             =   2790
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Concepto Transferencia "
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
         Left            =   510
         TabIndex        =   22
         Top             =   3195
         Width           =   2415
      End
   End
   Begin VB.Frame FrameHorasTrabajadas 
      Height          =   5130
      Left            =   45
      TabIndex        =   6
      Top             =   0
      Width           =   6435
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
         Left            =   540
         TabIndex        =   16
         Top             =   3870
         Width           =   2850
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
         Index           =   0
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Tag             =   "Tipo|N|N|||straba|codsecci||N|"
         Top             =   3420
         Width           =   1665
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
         Left            =   2745
         Locked          =   -1  'True
         TabIndex        =   12
         Text            =   "Text5"
         Top             =   2790
         Width           =   3285
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
         Left            =   1830
         MaxLength       =   6
         TabIndex        =   3
         Top             =   2790
         Width           =   870
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
         Left            =   1845
         MaxLength       =   10
         TabIndex        =   1
         Top             =   1800
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
         Index           =   20
         Left            =   1845
         MaxLength       =   10
         TabIndex        =   2
         Top             =   2340
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
         Left            =   1830
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Tag             =   "Tipo|N|N|||straba|codsecci||N|"
         Top             =   1035
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
         Left            =   4965
         TabIndex        =   5
         Top             =   4500
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
         Left            =   3750
         TabIndex        =   4
         Top             =   4485
         Width           =   1065
      End
      Begin MSComctlLib.ProgressBar Pb1 
         Height          =   240
         Left            =   480
         TabIndex        =   11
         Top             =   4170
         Width           =   5550
         _ExtentX        =   9790
         _ExtentY        =   423
         _Version        =   393216
         Appearance      =   1
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   90
         Top             =   3915
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label Label1 
         Caption         =   "Concepto Transferencia "
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
         Left            =   540
         TabIndex        =   15
         Top             =   3150
         Width           =   1875
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   14
         Left            =   1515
         MouseIcon       =   "frmPagoRecibos.frx":0274
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar banco"
         Top             =   2790
         Width           =   240
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Banco "
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
         Left            =   540
         TabIndex        =   13
         Top             =   2700
         Width           =   675
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   2
         Left            =   1530
         Picture         =   "frmPagoRecibos.frx":03C6
         ToolTipText     =   "Buscar fecha"
         Top             =   1800
         Width           =   240
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
         Left            =   540
         TabIndex        =   10
         Top             =   1530
         Width           =   1320
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Pago"
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
         Index           =   30
         Left            =   540
         TabIndex        =   9
         Top             =   2115
         Width           =   1155
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   6
         Left            =   1515
         Picture         =   "frmPagoRecibos.frx":0451
         ToolTipText     =   "Buscar fecha"
         Top             =   2340
         Width           =   240
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
         Left            =   540
         TabIndex        =   8
         Top             =   1035
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "Pago de Recibos"
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
         TabIndex        =   7
         Top             =   405
         Width           =   5925
      End
   End
End
Attribute VB_Name = "frmPagoRecibos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' +-+-+-+-+-+-+-+-+-+-+-
' +-+- Autor: MONICA +-+-
' +-+-+-+-+-+-+-+-+-+-+-

Option Explicit

Public OpcionListado As Byte
    '==== Listados BASICOS ====
    '=============================
    ' 1 .- Pago de Recibos de valsur y alzira
    ' 2 .- Pago de Recibos de natural de montaña
    
Public NumCod As String 'Para indicar cod. Traspaso,Movimiento, etc. que llama
                        'Para indicar nº oferta a imprimir

Public CadTag As String 'Cadena con el Tag del campo que se va a poner en D/H en los listados
                        'Se necesita si el tipo de codigo es texto

Public Event RectificarFactura(Cliente As String, Observaciones As String)

Private Conexion As Byte
'1.- Conexión a BD Ariges  2.- Conexión a BD Conta

Private HaDevueltoDatos As Boolean


Private WithEvents frmBan As frmBasico2 'Banco propio
Attribute frmBan.VB_VarHelpID = -1

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
Dim Tabla As String
Dim Codigo As String 'Código para FormulaSelection de Crystal Report
Dim TipCod As String
Dim Orden1 As String 'Campo de Ordenacion (por codigo) para Cristal Report
Dim Orden2 As String 'Campo de Ordenacion (por nombre) para Cristal Report
Dim Tipo As String
Dim Repetir As Boolean

Dim PrimeraVez As Boolean
Dim Contabilizada As Byte

Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 0, cerrar
    If cerrar Then Unload Me
End Sub

Private Sub cmdAceptar_Click(Index As Integer)
Dim cOrden As String
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim numOp As Byte
    
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
    
Dim cadSelect1 As String
Dim cadSelect2 As String
Dim cTabla As String
Dim SQL As String

    
    If Not DatosOk Then Exit Sub
    
    cadSelect = ""
    'Fecha de Recibo
    AnyadirAFormula cadSelect, "horas.fecharec = " & DBSet(txtCodigo(16).Text, "F")
              
    'Tipo de seccion
    AnyadirAFormula cadSelect, "straba.codsecci = " & Me.Combo1(1).ListIndex
    
    'La forma de pago tiene que ser de tipo Transferencia
    '[Monica]15/12/2016: en el caso de coopic no miro la fp
    If vParamAplic.Cooperativa <> 16 And vParamAplic.Cooperativa <> 0 Then
        AnyadirAFormula cadSelect, "forpago.tipoforp = 1"
    End If
    
    Tabla = "(horas INNER JOIN straba ON horas.codtraba = straba.codtraba) INNER JOIN forpago ON straba.codforpa = forpago.codforpa "
               
    cTabla = Tabla
    cadSelect1 = cadSelect
    cadSelect2 = cadSelect
    
    cTabla = QuitarCaracterACadena(cTabla, "{")
    cTabla = QuitarCaracterACadena(cTabla, "}")
    SQL = "Select count(*) FROM " & QuitarCaracterACadena(cTabla, "_1")
    If cadSelect1 <> "" Then
        cadSelect1 = QuitarCaracterACadena(cadSelect1, "{")
        cadSelect1 = QuitarCaracterACadena(cadSelect1, "}")
        cadSelect1 = QuitarCaracterACadena(cadSelect1, "_1")
        SQL = SQL & " WHERE " & cadSelect1
    End If
    
    If RegistrosAListar(SQL) = 0 Then
        MsgBox "No hay datos para mostrar en el Informe.", vbInformation
    Else
        AnyadirAFormula cadSelect, "horas.intconta = 0"

        'Comprobar si hay registros a Mostrar antes de abrir el Informe
        If HayRegParaInforme(Tabla, cadSelect) Then
            If vParamAplic.Cooperativa = 2 Or vParamAplic.Cooperativa = 16 Or vParamAplic.Cooperativa = 0 Or vParamAplic.Cooperativa = 18 Or vParamAplic.Cooperativa = 19 Then
                ProcesarCambiosPicassent (cadSelect)
            Else
                ProcesarCambios (cadSelect)
            End If
        Else
            Repetir = True
            If MsgBox("¿Desea repetir el proceso?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
                If vParamAplic.Cooperativa = 2 Or vParamAplic.Cooperativa = 16 Or vParamAplic.Cooperativa = 0 Or vParamAplic.Cooperativa = 18 Or vParamAplic.Cooperativa = 19 Then ' si la cooperativa es picassent repite norma como natural
                    RepetirNormaPicassent cadSelect
                Else
                    '[Monica]03/11/2010: anteriormente en Alzira no grababamos en rrecibosnomina, ahora sí
                    'ProcesarCambios (cadSelect2)
                    RepetirNormaPicassent cadSelect
                End If
            End If
        End If
    End If
    
    cmdCancel_Click (0)
    
End Sub

Private Sub ProcesarCambios(cadWHERE As String)
Dim SQL As String
Dim Sql2 As String
Dim Sql3 As String
Dim cad As String
Dim i As Integer
Dim HayReg As Integer
Dim B As Boolean
Dim Rs As ADODB.Recordset
Dim Rs2 As ADODB.Recordset
Dim Mens As String

Dim ImpHoras As Currency
Dim ImpHorasE As Currency
Dim ImpBruto As Currency
Dim IRPF As Currency
Dim SegSoc As Currency
Dim Neto As Currency
Dim Bruto As Currency
Dim CuentaPropia As String
Dim CodigoOrden34 As String

Dim Max As Long


On Error GoTo eProcesarCambios
    
    BorrarTMP
    CrearTMP

    conn.BeginTrans
    
    
    If cadWHERE <> "" Then
        cadWHERE = QuitarCaracterACadena(cadWHERE, "{")
        cadWHERE = QuitarCaracterACadena(cadWHERE, "}")
        cadWHERE = QuitarCaracterACadena(cadWHERE, "_1")
    End If
        
    SQL = "select count(distinct horas.codtraba) from (horas inner join straba on horas.codtraba = straba.codtraba) inner join forpago on straba.codforpa = forpago.codforpa where " & cadWHERE
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    Pb1.visible = True
    CargarProgres Pb1, Rs.Fields(0).Value
    
    Rs.Close
    
    Sql3 = "select max(idcontador) from rrecibosnomina"
    Max = DevuelveValor(Sql3) + 1
    
    
    If Check1(1).Value = 0 Then
        SQL = "select horas.codtraba, sum(horasdia), sum(compleme), sum(horasext) from (horas inner join straba on horas.codtraba = straba.codtraba) inner join forpago on straba.codforpa = forpago.codforpa where " & cadWHERE
    Else
        SQL = "select horas.codtraba, sum(horasproduc), sum(compleme), sum(horasext) from (horas inner join straba on horas.codtraba = straba.codtraba) inner join forpago on straba.codforpa = forpago.codforpa where " & cadWHERE
    End If
    SQL = SQL & " group by horas.codtraba "
    
'    BorrarTMP
'    CrearTMP
    
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not Rs.EOF
        IncrementarProgres Pb1, 1
        Mens = "Calculando Importes" & vbCrLf & vbCrLf & "Trabajador: " & Rs!CodTraba & vbCrLf
        
        Sql2 = "select salarios.impsalar, salarios.imphorae, straba.dtosirpf, straba.dtosegso, straba.porc_antig from salarios, straba where straba.codtraba = " & DBSet(Rs!CodTraba, "N")
        Sql2 = Sql2 & " and salarios.codcateg = straba.codcateg "
        
        Set Rs2 = New ADODB.Recordset
        Rs2.Open Sql2, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        ImpHoras = Round2(DBLet(Rs.Fields(1).Value, "N") * DBLet(Rs2!impsalar, "N"), 2)
        ImpHorasE = Round2(DBLet(Rs.Fields(3).Value, "N") * DBLet(Rs2!imphorae, "N"), 2)
        ImpBruto = Round2(ImpHoras + ImpHorasE + DBLet(Rs.Fields(2).Value, "N"), 2)
        
'        [Monica]23/03/2010: incrementamos el bruto el porcentaje de antigüedad si lo tiene, si no 0
        ImpBruto = ImpBruto + Round2(ImpBruto * DBLet(Rs2!porc_antig, "N") / 100, 2)
        
        IRPF = Round2(ImpBruto * DBLet(Rs2!dtosirpf, "N") / 100, 2)
        SegSoc = Round2(ImpBruto * DBLet(Rs2!dtosegso, "N") / 100, 2)
        
        Neto = Round2(ImpBruto - IRPF - SegSoc, 2)
        
        '[Monica]23/03/2016: si el importe es negativo no entra
        If Neto >= 0 Then
        
            Sql3 = "insert into tmpImpor (codtraba, importe) values ("
            Sql3 = Sql3 & DBSet(Rs.Fields(0).Value, "N") & "," & DBSet(ImporteSinFormato(CStr(Neto)), "N") & ")"
            
            conn.Execute Sql3
            
            Sql3 = "insert into rrecibosnomina (codtraba, fechahora, importe, base34, porcsegso1, porcsegso2, porcirpf, "
            Sql3 = Sql3 & "importesegso1, importesegso2, importeirpf, complemento, neto34, idcontador) values ("
            Sql3 = Sql3 & DBSet(Rs.Fields(0).Value, "N") & ","
            Sql3 = Sql3 & DBSet(txtCodigo(16).Text, "F") & "," & DBSet(ImporteSinFormato(CStr(ImpBruto)), "N") & ","
            Sql3 = Sql3 & DBSet(ImporteSinFormato(CStr(ImpBruto)), "N") & ","
            Sql3 = Sql3 & DBSet(Rs2!dtosegso, "N") & ","
            Sql3 = Sql3 & "0,"
            Sql3 = Sql3 & DBSet(Rs2!dtosirpf, "N") & ","
            Sql3 = Sql3 & DBSet(SegSoc, "N") & ",0," & DBSet(IRPF, "N") & ","
            Sql3 = Sql3 & DBSet(0, "N") & ","
            Sql3 = Sql3 & DBSet(Neto, "N") & ","
            Sql3 = Sql3 & DBSet(Max, "N")
            Sql3 = Sql3 & ")"
            
            conn.Execute Sql3
            
        End If
        
        Set Rs2 = Nothing
            
        Rs.MoveNext
    Wend
    Set Rs = Nothing
    
    '[Monica]22/11/2013: iban
    SQL = "select codbanco, codsucur, digcontr, cuentaba, codorden34, iban from banpropi where codbanpr = " & DBSet(txtCodigo(18).Text, "N")
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    CodigoOrden34 = ""
    
    If Rs.EOF Then
        cad = ""
    Else
        If IsNull(Rs!CodBanco) Then
            cad = ""
        Else
            '[Monica]22/11/2013: iban
            cad = Format(Rs!CodBanco, "0000") & "|" & Format(DBLet(Rs!CodSucur, "T"), "0000") & "|" & DBLet(Rs!digcontr, "T") & "|" & Format(DBLet(Rs!CuentaBa, "T"), "0000000000") & "|" & DBLet(Rs!Iban, "T") & "|"
        End If
        CodigoOrden34 = DBLet(Rs!codorden34, "T")
    End If
    
    Set Rs = Nothing
    
    CuentaPropia = cad
    
    '[Monica]22/11/2013: iban
    Dim vSeccion As CSeccion
    If vSeccion Is Nothing Then
        Set vSeccion = New CSeccion
        If vSeccion.LeerDatos(vParamAplic.Seccionhorto) Then
            vSeccion.AbrirConta
        End If
    End If
    If vEmpresa.AplicarNorma19_34Nueva = 1 Then
        If HayXML Then
            B = GeneraFicheroNorma34SEPA_XML(vParam.CifEmpresa, CDate(txtCodigo(20).Text), CuentaPropia, "", "Pago Nómina", Combo1(0).ListIndex, CodigoOrden34)
        Else
            B = GeneraFicheroNorma34SEPA(vParam.CifEmpresa, CDate(txtCodigo(20).Text), CuentaPropia, "", "Pago Nómina", Combo1(0).ListIndex, CodigoOrden34)
        End If
    Else
        B = GeneraFicheroNorma34New(vParam.CifEmpresa, CDate(txtCodigo(20).Text), CuentaPropia, 9, 0, "Pago Nómina", CodigoOrden34, Combo1(0).ListIndex)
    End If
    vSeccion.CerrarConta
    Set vSeccion = Nothing
    
    
    If B Then
        B = CopiarFichero
        If B Then
            If Not Repetir Then
                If MsgBox("¿Proceso realizado correctamente para actualizar?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
                    SQL = "update horas, straba, forpago set horas.intconta = 1 where horas.codtraba = straba.codtraba and straba.codforpa = forpago.codforpa and " & cadWHERE
                    conn.Execute SQL
                    
                Else
                    SQL = "delete from rrecibosnomina where fechahora = " & DBSet(txtCodigo(16).Text, "F")
                    SQL = SQL & " and idcontador = " & DBSet(Max, "N")
                    
                    conn.Execute SQL
                    
                End If
            End If
        End If
    End If

eProcesarCambios:
    If Err.Number <> 0 Then
        Mens = Err.Description
        B = False
    End If
    If B Then
        conn.CommitTrans
        MsgBox "Proceso realizado correctamente.", vbExclamation
        cmdCancel_Click (0)
    Else
        conn.RollbackTrans
        MsgBox "Error " & Mens, vbExclamation
    End If
End Sub



Private Sub ProcesarCambiosNatural(cadWHERE As String)
Dim SQL As String
Dim Sql2 As String
Dim Sql3 As String
Dim cad As String
Dim i As Integer
Dim HayReg As Integer
Dim B As Boolean
Dim Rs As ADODB.Recordset
Dim Rs2 As ADODB.Recordset
Dim Mens As String

Dim ImpHoras As Currency
Dim ImpHorasE As Currency
Dim ImpBruto As Currency
Dim IRPF As Currency
Dim SegSoc As Currency
Dim SegSoc1 As Currency
Dim Neto As Currency
Dim Bruto As Currency
Dim CuentaPropia As String
Dim CodigoOrden34 As String

Dim Neto34 As Currency
Dim Bruto34 As Currency
Dim Jornadas As Currency
Dim Diferencia As Currency
Dim BaseSegso As Currency
Dim Complemento As Currency
Dim TSegSoc As Currency
Dim TSegSoc1 As Currency
Dim Max As Long

Dim Sql5 As String
Dim RS5 As ADODB.Recordset

Dim Anticipo As Currency

On Error GoTo eProcesarCambios
    
    BorrarTMP
    CrearTMP

    conn.BeginTrans
    
    If cadWHERE <> "" Then
        cadWHERE = QuitarCaracterACadena(cadWHERE, "{")
        cadWHERE = QuitarCaracterACadena(cadWHERE, "}")
        cadWHERE = QuitarCaracterACadena(cadWHERE, "_1")
    End If
        
    SQL = "select count(distinct horas.codtraba) from (horas inner join straba on horas.codtraba = straba.codtraba) inner join forpago on straba.codforpa = forpago.codforpa where " & cadWHERE
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    Pb2.visible = True
    CargarProgres Pb2, Rs.Fields(0).Value
    
    Rs.Close
    
    SQL = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute SQL
    
    Sql3 = "select max(idcontador) from rrecibosnomina"
    Max = DevuelveValor(Sql3) + 1
    
    SQL = "select horas.codtraba, sum(horasdia), sum(compleme), sum(horasext) from (horas inner join straba on horas.codtraba = straba.codtraba) inner join forpago on straba.codforpa = forpago.codforpa where " & cadWHERE
    SQL = SQL & " group by horas.codtraba "
    
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not Rs.EOF
        IncrementarProgres Pb2, 1
        Mens = "Calculando Importes" & vbCrLf & vbCrLf & "Trabajador: " & Rs!CodTraba & vbCrLf
        
        Sql2 = "select salarios.*, straba.porc_antig, straba.dtosegso, straba.dtosirpf from salarios, straba where straba.codtraba = " & DBSet(Rs!CodTraba, "N")
        Sql2 = Sql2 & " and salarios.codcateg = straba.codcateg "
        
        Set Rs2 = New ADODB.Recordset
        Rs2.Open Sql2, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        ImpHoras = Round2(DBLet(Rs.Fields(1).Value, "N") * DBLet(Rs2!impsalar, "N"), 2)
        ImpHorasE = Round2(DBLet(Rs.Fields(3).Value, "N") * DBLet(Rs2!imphorae, "N"), 2)
        ImpBruto = Round2(ImpHoras + ImpHorasE + DBLet(Rs.Fields(2).Value, "N"), 2)
        
'        [Monica]23/03/2010: incrementamos el bruto el porcentaje de antigüedad si lo tiene, si no 0
        ImpBruto = ImpBruto + Round2(ImpBruto * DBLet(Rs2!porc_antig, "N") / 100, 2)
        
        ' natural de montaña
        Neto = Round2(ImpBruto, 2)
        
        
        '[Monica]18/09/2013: anticipos pendientes de descuento
        Anticipo = AnticiposPendientes(Rs!CodTraba)
        
        '[Monica] 24/08/2010 Añadido el tema de fijos en natural de Montaña
        If (ImpHoras + ImpHorasE) = 0 Then
            Bruto34 = Neto
            
            IRPF = 0
            SegSoc = 0
            SegSoc1 = 0
            
            Neto34 = 0
        
        
            IRPF = Round2(Bruto34 * DBLet(Rs2!dtosirpf, "N") / 100, 2)
            TSegSoc = Round2(Bruto34 * DBLet(Rs2!dtosegso, "N") / 100, 2)
            TSegSoc1 = 0
            Complemento = 0
        
            Neto34 = Round2(Bruto34 + Complemento - IRPF - TSegSoc - TSegSoc1, 2)
            
            '[Monica]18/09/2013: anticipos pendientes de descuento
            Neto34 = Neto34 - Anticipo
        
            '[Monica]23/03/2016: si el importe es negativo no entra
            If Neto34 >= 0 Then
                Sql3 = "insert into rrecibosnomina (codtraba, fechahora, importe, base34, porcsegso1, porcsegso2, porcirpf, "
                Sql3 = Sql3 & "importesegso1, importesegso2, importeirpf, complemento, neto34, anticipo, idcontador) values ("
                Sql3 = Sql3 & DBSet(Rs.Fields(0).Value, "N") & ","
                Sql3 = Sql3 & DBSet(txtCodigo(1).Text, "F") & "," & DBSet(ImporteSinFormato(CStr(Neto)), "N") & ","
                Sql3 = Sql3 & DBSet(ImporteSinFormato(CStr(Bruto34)), "N") & ","
                Sql3 = Sql3 & DBSet(Rs2!dtosegso, "N") & ","
                Sql3 = Sql3 & "0,"
                Sql3 = Sql3 & DBSet(Rs2!dtosirpf, "N") & ","
                Sql3 = Sql3 & DBSet(TSegSoc, "N") & "," & DBSet(TSegSoc1, "N") & "," & DBSet(IRPF, "N") & ","
                Sql3 = Sql3 & DBSet(Complemento, "N") & ","
                Sql3 = Sql3 & DBSet(Neto34, "N") & ","
                '[Monica]18/09/2013: anticipos pendientes de descuento
                Sql3 = Sql3 & DBSet(Anticipo, "N") & ","
                Sql3 = Sql3 & DBSet(Max, "N")
                Sql3 = Sql3 & ")"
                
                conn.Execute Sql3
            End If
        
        Else ' para el resto funciona como antes ( los que van por horas )
        
            Jornadas = Round2((Neto / vParamAplic.EurosTrabdiaNOMI) * ImporteSinFormato(txtCodigo(3).Text) / 100, 0)
            
            '[Monica]27/07/2010: si el nro maximo de jornadas es superior al maximo, se deja el maximo
            If vParamAplic.NroMaxJornadasNOMI <> 0 And Jornadas > vParamAplic.NroMaxJornadasNOMI Then
                Jornadas = vParamAplic.NroMaxJornadasNOMI
            End If
            
            Bruto34 = Round2(Jornadas * vParamAplic.EurosTrabdiaNOMI, 2)
            
            IRPF = 0
            SegSoc = 0
            SegSoc1 = 0
            
            Neto34 = 0
            
    '        BaseSegso = Round2(Bruto34 * ImporteSinFormato(txtCodigo(3).Text) / 100, 2)
            
            IRPF = Round2(ImpBruto * ImporteSinFormato(txtCodigo(6).Text) / 100, 2)
            TSegSoc = Round2(Bruto34 * ImporteSinFormato(txtCodigo(4).Text) / 100, 2)
            TSegSoc1 = Round2(Bruto34 * ImporteSinFormato(txtCodigo(5).Text) / 100, 2)
            
    '        SegSoc = Round2(BaseSegso * ImporteSinFormato(txtCodigo(4).Text) / 100, 2)
    '        SegSoc1 = Round2(BaseSegso * ImporteSinFormato(txtCodigo(5).Text) / 100, 2)
            
            Complemento = ImpBruto - Bruto34
            
            Neto34 = Round2(Bruto34 + Complemento - IRPF - TSegSoc - TSegSoc1, 2)
            
            '[Monica]18/09/2013: anticipos pendientes de descuento
            Neto34 = Neto34 - Anticipo
        
            '[Monica]23/03/2016: si el importe es negativo no entra
            If Neto34 >= 0 Then
        
                Sql3 = "insert into rrecibosnomina (codtraba, fechahora, importe, base34, porcsegso1, porcsegso2, porcirpf, "
                Sql3 = Sql3 & "importesegso1, importesegso2, importeirpf, complemento, neto34, anticipo, idcontador) values ("
                Sql3 = Sql3 & DBSet(Rs.Fields(0).Value, "N") & ","
                Sql3 = Sql3 & DBSet(txtCodigo(1).Text, "F") & "," & DBSet(ImporteSinFormato(CStr(Neto)), "N") & ","
                Sql3 = Sql3 & DBSet(ImporteSinFormato(CStr(Bruto34)), "N") & ","
                Sql3 = Sql3 & DBSet(ImporteSinFormato(txtCodigo(4).Text), "N") & ","
                Sql3 = Sql3 & DBSet(ImporteSinFormato(txtCodigo(5).Text), "N") & ","
                Sql3 = Sql3 & DBSet(ImporteSinFormato(txtCodigo(6).Text), "N") & ","
                Sql3 = Sql3 & DBSet(TSegSoc, "N") & "," & DBSet(TSegSoc1, "N") & "," & DBSet(IRPF, "N") & ","
                Sql3 = Sql3 & DBSet(Complemento, "N") & ","
                Sql3 = Sql3 & DBSet(Neto34, "N") & ","
                '[Monica]18/09/2013: anticipos pendientes de descuento
                Sql3 = Sql3 & DBSet(Anticipo, "N") & ","
                Sql3 = Sql3 & DBSet(Max, "N")
                Sql3 = Sql3 & ")"
                
                conn.Execute Sql3
            End If
        End If

        '[Monica]23/03/2016: si el importe es negativo no entra
        If Neto34 >= 0 Then
            '[Monica]18/09/2013: anticipos pendientes de descuento
            SQL = "update horasanticipos set descontado = 1, fechahora = " & DBSet(txtCodigo(1).Text, "F") & ", idcontador = " & DBSet(Max, "N")
            SQL = SQL & " where codtraba = " & DBSet(Rs.Fields(0).Value, "N") & " and descontado = 0 "
            conn.Execute SQL
                        
            
            Sql3 = "insert into tmpinformes (codusu, codigo1, fecha1) values (" & vUsu.Codigo & "," & DBSet(Rs.Fields(0).Value, "N") & ","
            Sql3 = Sql3 & DBSet(txtCodigo(1).Text, "F") & ")"
            
            conn.Execute Sql3
            
            Sql3 = "insert into tmpImpor (codtraba, importe) values ("
            Sql3 = Sql3 & DBSet(Rs.Fields(0).Value, "N") & "," & DBSet(ImporteSinFormato(CStr(Neto34)), "N") & ")"
            
            conn.Execute Sql3
        End If
        
        Set Rs2 = Nothing
            
        Rs.MoveNext
    Wend
    
    Set Rs = Nothing
    '[Monica]22/11/2013:iban
    SQL = "select codbanco, codsucur, digcontr, cuentaba, codorden34, iban from banpropi where codbanpr = " & DBSet(txtCodigo(0).Text, "N")
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    CodigoOrden34 = ""
    
    If Rs.EOF Then
        cad = ""
    Else
        If IsNull(Rs!CodBanco) Then
            cad = ""
        Else
            '[Monica]22/11/2013:iban
            cad = Format(Rs!CodBanco, "0000") & "|" & Format(DBLet(Rs!CodSucur, "T"), "0000") & "|" & DBLet(Rs!digcontr, "T") & "|" & Format(DBLet(Rs!CuentaBa, "T"), "0000000000") & "|" & DBLet(Rs!Iban, "T") & "|"
        End If
        CodigoOrden34 = DBLet(Rs!codorden34, "T")
    End If
    
    Set Rs = Nothing
    
    CuentaPropia = cad
    '[Monica]22/11/2013: iban
    Dim vSeccion As CSeccion
    If vSeccion Is Nothing Then
        Set vSeccion = New CSeccion
        If vSeccion.LeerDatos(vParamAplic.Seccionhorto) Then
            vSeccion.AbrirConta
        End If
    End If
    If vEmpresa.AplicarNorma19_34Nueva = 1 Then
        '[Monica]29/01/2014: hay que pasarle el CodigoOrden34
        If HayXML Then
            B = GeneraFicheroNorma34SEPA_XML(vParam.CifEmpresa, CDate(txtCodigo(2).Text), CuentaPropia, "", "Pago Nómina", Combo1(2).ListIndex, CodigoOrden34)
        Else
            B = GeneraFicheroNorma34SEPA(vParam.CifEmpresa, CDate(txtCodigo(2).Text), CuentaPropia, "", "Pago Nómina", Combo1(2).ListIndex, CodigoOrden34)
        End If
    Else
        B = GeneraFicheroNorma34New(vParam.CifEmpresa, CDate(txtCodigo(2).Text), CuentaPropia, 9, 0, "Pago Nómina", CodigoOrden34, Combo1(2).ListIndex)
    End If
    vSeccion.CerrarConta
    Set vSeccion = Nothing
    
'antes
'    b = GeneraFicheroNorma34New(vParam.CifEmpresa, CDate(txtcodigo(2).Text), CuentaPropia, 9, 0, "Pago Nómina", CodigoOrden34, Combo1(2).ListIndex)
    If B Then
        If CopiarFichero Then
'            If MsgBox("¿Desea realizar la impresión de cheques?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
'                    cadparam = "pEmpresa=""" & """|"
'                    numParam = 1
'                    cadFormula = "{tmpinformes.codusu}=" & vUsu.Codigo
'                    cadNombreRPT = "rChequeNomina.rpt"
'                    cadTitulo = "Impresion de Cheques de nómina"
'                    ConSubInforme = False
'
'                    LlamarImprimir
'            Else
'                    MsgBox "No se ha podido realizar la impresion de cheques.", vbExclamation
'            End If
            
            If Not Repetir Then
                If MsgBox("¿Proceso realizado correctamente para actualizar?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
                    SQL = "update horas, straba, forpago set horas.intconta = 1 where horas.codtraba = straba.codtraba and straba.codforpa = forpago.codforpa and " & cadWHERE
                    conn.Execute SQL
                    
                Else
                    '[Monica]18/09/2013: añado esto para que si no es correcto para actualizar lo deje como estaba
                    SQL = "update horasanticipos set descontado = 0, fechahora = null, idcontador = null "
                    SQL = SQL & " where descontado = 1 and fechahora = " & DBSet(txtCodigo(1).Text, "F") & " and  idcontador = " & DBSet(Max, "N")
                    conn.Execute SQL
                End If
            End If
        End If
    End If

eProcesarCambios:
    If Err.Number <> 0 Then
        Mens = Err.Description
        B = False
    End If
    If B Then
        conn.CommitTrans
        MsgBox "Proceso realizado correctamente.", vbExclamation
        cmdCancel_Click (1)
    Else
        conn.RollbackTrans
        MsgBox "Error " & Mens, vbExclamation
    End If
End Sub


Private Sub RepetirNormaNatural(cadWHERE As String)
Dim SQL As String
Dim Sql2 As String
Dim Sql3 As String
Dim cad As String
Dim i As Integer
Dim HayReg As Integer
Dim B As Boolean
Dim Rs As ADODB.Recordset
Dim Rs2 As ADODB.Recordset
Dim Mens As String

Dim ImpHoras As Currency
Dim ImpHorasE As Currency
Dim ImpBruto As Currency
Dim IRPF As Currency
Dim SegSoc As Currency
Dim SegSoc1 As Currency
Dim Neto As Currency
Dim Bruto As Currency
Dim CuentaPropia As String
Dim CodigoOrden34 As String

Dim Neto34 As Currency
Dim Bruto34 As Currency
Dim Jornadas As Currency
Dim Diferencia As Currency
Dim BaseSegso As Currency
Dim Complemento As Currency
Dim TSegSoc As Currency
Dim TSegSoc1 As Currency
Dim Max As Long

Dim Sql5 As String
Dim RS5 As ADODB.Recordset

Dim IdContador As Long

On Error GoTo eRepetirNormaNatural
    
    BorrarTMP
    CrearTMP

    
    If cadWHERE <> "" Then
        cadWHERE = QuitarCaracterACadena(cadWHERE, "{")
        cadWHERE = QuitarCaracterACadena(cadWHERE, "}")
        cadWHERE = QuitarCaracterACadena(cadWHERE, "_1")
    End If
        
    SQL = "select max(idcontador) from rrecibosnomina where fechahora = " & DBSet(txtCodigo(1).Text, "F")
    IdContador = DevuelveValor(SQL)
    
    SQL = "select count(*) from rrecibosnomina where idcontador = " & DBSet(IdContador, "N")
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    Pb2.visible = True
    CargarProgres Pb2, Rs.Fields(0).Value
    
    Rs.Close
    
    SQL = "select * from rrecibosnomina where idcontador = " & DBSet(IdContador, "N")
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not Rs.EOF
    
        IncrementarProgres Pb2, 1
        Mens = "Calculando Importes" & vbCrLf & vbCrLf & "Trabajador: " & Rs!CodTraba & vbCrLf
        
        Sql3 = "insert into tmpImpor (codtraba, importe) values ("
        Sql3 = Sql3 & DBSet(Rs.Fields(0).Value, "N") & "," & DBSet(ImporteSinFormato(CStr(Rs!Neto34)), "N") & ")"
        
        conn.Execute Sql3
        
            
        Rs.MoveNext
    Wend
    
    Set Rs = Nothing
    '[Monica]22/11/2013: iban
    SQL = "select codbanco, codsucur, digcontr, cuentaba, codorden34, iban from banpropi where codbanpr = " & DBSet(txtCodigo(0).Text, "N")
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    CodigoOrden34 = ""
    
    If Rs.EOF Then
        cad = ""
    Else
        If IsNull(Rs!CodBanco) Then
            cad = ""
        Else
            '[Monica]22/11/2013: iban
            cad = Format(Rs!CodBanco, "0000") & "|" & Format(DBLet(Rs!CodSucur, "T"), "0000") & "|" & DBLet(Rs!digcontr, "T") & "|" & Format(DBLet(Rs!CuentaBa, "T"), "0000000000") & "|" & DBLet(Rs!Iban, "T") & "|"
        End If
        CodigoOrden34 = DBLet(Rs!codorden34, "T")
    End If
    
    Set Rs = Nothing
    
    CuentaPropia = cad
    '[Monica]22/11/2013:iban
    Dim vSeccion As CSeccion
    If vSeccion Is Nothing Then
        Set vSeccion = New CSeccion
        If vSeccion.LeerDatos(vParamAplic.Seccionhorto) Then
            vSeccion.AbrirConta
        End If
    End If
    
    If vEmpresa.AplicarNorma19_34Nueva = 1 Then
        If HayXML Then
            B = GeneraFicheroNorma34SEPA_XML(vParam.CifEmpresa, CDate(txtCodigo(2).Text), CuentaPropia, "", "Pago Nómina", Combo1(2).ListIndex, CodigoOrden34)
        Else
            B = GeneraFicheroNorma34SEPA(vParam.CifEmpresa, CDate(txtCodigo(2).Text), CuentaPropia, "", "Pago Nómina", Combo1(2).ListIndex, CodigoOrden34)
        End If
    Else
        B = GeneraFicheroNorma34New(vParam.CifEmpresa, CDate(txtCodigo(2).Text), CuentaPropia, 9, 0, "Pago Nómina", CodigoOrden34, Combo1(2).ListIndex)
    End If
    vSeccion.CerrarConta
    Set vSeccion = Nothing
  'antes
'    b = GeneraFicheroNorma34New(vParam.CifEmpresa, CDate(txtcodigo(2).Text), CuentaPropia, 9, 0, "Pago Nómina", CodigoOrden34, Combo1(2).ListIndex)
    If B Then
        If CopiarFichero Then
'            If MsgBox("¿Desea realizar la impresión de cheques?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
'                    cadparam = "pEmpresa=""" & """|"
'                    numParam = 1
'                    cadFormula = "{tmpinformes.codusu}=" & vUsu.Codigo
'                    cadNombreRPT = "rChequeNomina.rpt"
'                    cadTitulo = "Impresion de Cheques de nómina"
'                    ConSubInforme = False
'
'                    LlamarImprimir
'            Else
'                    MsgBox "No se ha podido realizar la impresion de cheques.", vbExclamation
'            End If
        End If
    End If

eRepetirNormaNatural:
    If Err.Number <> 0 Then
        Mens = Err.Description
        B = False
    End If
    If B Then
        MsgBox "Proceso realizado correctamente.", vbExclamation
        cmdCancel_Click (1)
    Else
        MsgBox "Error " & Mens, vbExclamation
    End If
End Sub


Private Sub RepetirNormaPicassent(cadWHERE As String)
Dim SQL As String
Dim Sql2 As String
Dim Sql3 As String
Dim cad As String
Dim i As Integer
Dim HayReg As Integer
Dim B As Boolean
Dim Rs As ADODB.Recordset
Dim Rs2 As ADODB.Recordset
Dim Mens As String

Dim ImpHoras As Currency
Dim ImpHorasE As Currency
Dim ImpBruto As Currency
Dim IRPF As Currency
Dim SegSoc As Currency
Dim SegSoc1 As Currency
Dim Neto As Currency
Dim Bruto As Currency
Dim CuentaPropia As String
Dim CodigoOrden34 As String

Dim Neto34 As Currency
Dim Bruto34 As Currency
Dim Jornadas As Currency
Dim Diferencia As Currency
Dim BaseSegso As Currency
Dim Complemento As Currency
Dim TSegSoc As Currency
Dim TSegSoc1 As Currency
Dim Max As Long

Dim Sql5 As String
Dim RS5 As ADODB.Recordset

Dim IdContador As Long
Dim TieneEmbargo As String

On Error GoTo eRepetirNormaPicassent
    
    BorrarTMP
    CrearTMP

    
    If cadWHERE <> "" Then
        cadWHERE = QuitarCaracterACadena(cadWHERE, "{")
        cadWHERE = QuitarCaracterACadena(cadWHERE, "}")
        cadWHERE = QuitarCaracterACadena(cadWHERE, "_1")
    End If
        
    SQL = "select max(idcontador) from rrecibosnomina where fechahora = " & DBSet(txtCodigo(16).Text, "F")
    IdContador = DevuelveValor(SQL)
    
    SQL = "select count(*) from rrecibosnomina where idcontador = " & DBSet(IdContador, "N")
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    Pb2.visible = True
    CargarProgres Pb2, Rs.Fields(0).Value
    
    Rs.Close
    
    SQL = "select * from rrecibosnomina where idcontador = " & DBSet(IdContador, "N")
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not Rs.EOF
    
        IncrementarProgres Pb2, 1
        Mens = "Calculando Importes" & vbCrLf & vbCrLf & "Trabajador: " & Rs!CodTraba & vbCrLf
        
        TieneEmbargo = DevuelveValor("select hayembargo from straba where codtraba = " & DBSet(Rs!CodTraba, "N"))
        If TieneEmbargo = "0" Then
            Sql3 = "insert into tmpImpor (codtraba, importe) values ("
            Sql3 = Sql3 & DBSet(Rs.Fields(0).Value, "N") & "," & DBSet(Rs!Neto34, "N") & ")"
        
            conn.Execute Sql3
        End If
            
        Rs.MoveNext
    Wend
    
    Set Rs = Nothing
    '[Monica]22/11/2013: iban
    SQL = "select codbanco, codsucur, digcontr, cuentaba, codorden34, iban from banpropi where codbanpr = " & DBSet(txtCodigo(18).Text, "N")
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    CodigoOrden34 = ""
    
    If Rs.EOF Then
        cad = ""
    Else
        If IsNull(Rs!CodBanco) Then
            cad = ""
        Else
            '[Monica]22/11/2013: iban
            cad = Format(Rs!CodBanco, "0000") & "|" & Format(DBLet(Rs!CodSucur, "T"), "0000") & "|" & DBLet(Rs!digcontr, "T") & "|" & Format(DBLet(Rs!CuentaBa, "T"), "0000000000") & "|" & DBLet(Rs!Iban, "T") & "|"
        End If
        CodigoOrden34 = DBLet(Rs!codorden34, "T")
    End If
    
    Set Rs = Nothing
    
    CuentaPropia = cad
    '[Monica]22/11/2013: iban
    Dim vSeccion As CSeccion
    If vSeccion Is Nothing Then
        Set vSeccion = New CSeccion
        If vSeccion.LeerDatos(vParamAplic.Seccionhorto) Then
            vSeccion.AbrirConta
        End If
    End If
    
    If vEmpresa.AplicarNorma19_34Nueva = 1 Then
        If HayXML Then
            B = GeneraFicheroNorma34SEPA_XML(vParam.CifEmpresa, CDate(txtCodigo(20).Text), CuentaPropia, "", "Pago Nómina", Combo1(0).ListIndex, CodigoOrden34)
        Else
            B = GeneraFicheroNorma34SEPA(vParam.CifEmpresa, CDate(txtCodigo(20).Text), CuentaPropia, "", "Pago Nómina", Combo1(0).ListIndex, CodigoOrden34)
        End If
    Else
        B = GeneraFicheroNorma34New(vParam.CifEmpresa, CDate(txtCodigo(20).Text), CuentaPropia, 9, 0, "Pago Nómina", CodigoOrden34, Combo1(0).ListIndex)
    End If
    
    vSeccion.CerrarConta
    Set vSeccion = Nothing
    
    
'antes
'    b = GeneraFicheroNorma34New(vParam.CifEmpresa, CDate(txtcodigo(20).Text), CuentaPropia, 9, 0, "Pago Nómina", CodigoOrden34, Combo1(0).ListIndex)
    If B Then
        Mens = "Copiar Fichero"
        If CopiarFichero Then
'            If MsgBox("¿Desea realizar la impresión de cheques?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
'                    cadparam = "pEmpresa=""" & """|"
'                    numParam = 1
'                    cadFormula = "{tmpinformes.codusu}=" & vUsu.Codigo
'                    cadNombreRPT = "rChequeNomina.rpt"
'                    cadTitulo = "Impresion de Cheques de nómina"
'                    ConSubInforme = False
'
'                    LlamarImprimir
'            Else
'                    MsgBox "No se ha podido realizar la impresion de cheques.", vbExclamation
'            End If
        Else
            B = False
        End If
    End If

eRepetirNormaPicassent:
    If Err.Number <> 0 Then
        Mens = Err.Description
        B = False
    End If
    If B Then
        MsgBox "Proceso realizado correctamente.", vbExclamation
        cmdCancel_Click (1)
    Else
        MsgBox "Error " & Mens, vbExclamation
    End If
End Sub



Private Sub CmdAceptarNat_Click()
Dim cOrden As String
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim numOp As Byte
    
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
    
Dim cadSelect1 As String
Dim cadSelect2 As String
Dim cTabla As String
Dim SQL As String

    
    If Not DatosOk Then Exit Sub
    
    cadSelect = ""
    'Fecha de Recibo
    AnyadirAFormula cadSelect, "horas.fecharec = " & DBSet(txtCodigo(1).Text, "F")
              
    'Tipo de seccion
    AnyadirAFormula cadSelect, "straba.codsecci = " & Me.Combo1(3).ListIndex
    
    'La forma de pago tiene que ser de tipo Transferencia
    '[Monica]15/12/2016: solo en el caso de coopic me da igual la fp pq ellos no pagan
    If vParamAplic.Cooperativa <> 16 Then
        AnyadirAFormula cadSelect, "forpago.tipoforp = 1"
    End If
    
    Tabla = "(horas INNER JOIN straba ON horas.codtraba = straba.codtraba) INNER JOIN forpago ON straba.codforpa = forpago.codforpa "
               
    cTabla = Tabla
    cadSelect1 = cadSelect
    cadSelect2 = cadSelect
    
    cTabla = QuitarCaracterACadena(cTabla, "{")
    cTabla = QuitarCaracterACadena(cTabla, "}")
    SQL = "Select count(*) FROM " & QuitarCaracterACadena(cTabla, "_1")
    If cadSelect1 <> "" Then
        cadSelect1 = QuitarCaracterACadena(cadSelect1, "{")
        cadSelect1 = QuitarCaracterACadena(cadSelect1, "}")
        cadSelect1 = QuitarCaracterACadena(cadSelect1, "_1")
        SQL = SQL & " WHERE " & cadSelect1
    End If
    
    If RegistrosAListar(SQL) = 0 Then
        MsgBox "No hay datos para mostrar en el Informe.", vbInformation
    Else
        AnyadirAFormula cadSelect, "horas.intconta = 0"

        '[Monica]06/05/2015: comprobamos que todos los trabajadores tengan direccion
        If Not DireccionesOk(Tabla, cadSelect) Then Exit Sub

        'Comprobar si hay registros a Mostrar antes de abrir el Informe
        If HayRegParaInforme(Tabla, cadSelect) Then
            ProcesarCambiosNatural (cadSelect)
        Else
            Repetir = True
            If MsgBox("¿Desea repetir el proceso?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
                RepetirNormaNatural (cadSelect2)
            End If
        End If
    End If
    
    cmdCancel_Click (1)

End Sub

Private Function DireccionesOk(cTabla As String, cWhere As String) As Boolean
Dim SQL As String
Dim cadResult As String
Dim Rs As ADODB.Recordset

    On Error GoTo eDireccionesOk
    
    DireccionesOk = False

    SQL = "Select straba.* FROM " & cTabla & "  WHERE " & cWhere
    SQL = SQL & " and (domtraba is null or domtraba = '' or codpobla is null or codpobla = ''  or pobtraba is null or pobtraba is null or protraba is null or protraba = '') "
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    cadResult = ""
    While Not Rs.EOF
        cadResult = cadResult & DBLet(Rs!CodTraba) & ","
        
        Rs.MoveNext
    Wend
    Set Rs = Nothing
    
    If cadResult <> "" Then
        cadResult = Mid(cadResult, 1, Len(cadResult) - 1)
    
        MsgBox "Los siguientes trabajadores no tienen la dirección correcta: " & vbCrLf & vbCrLf & cadResult, vbExclamation
    
    End If
    
    
    DireccionesOk = (cadResult = "")
    Exit Function
eDireccionesOk:
    MuestraError Err.Number, "Direcciones Correctas", Err.Description
End Function


Private Sub cmdCancel_Click(Index As Integer)
    Unload Me
End Sub


Private Sub Combo1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        Select Case OpcionListado
            Case 1 ' Pago de recibos de valsur y alzira
                Combo1(0).ListIndex = 0
                Combo1(1).ListIndex = 0
                
                PonerFocoCmb Combo1(1)
            
            Case 2 ' Pago de recibos de natural de montaña
                txtCodigo(3).Text = Format(vParamAplic.PorcJornadaNOMI, "##0.00")
                txtCodigo(4).Text = Format(vParamAplic.PorcSegSo1NOMI, "##0.00")
                txtCodigo(5).Text = Format(vParamAplic.PorcSegSo2NOMI, "##0.00")
                txtCodigo(6).Text = Format(vParamAplic.PorcIRPFNOMI, "##0.00")
            
                Combo1(2).ListIndex = 0
                Combo1(3).ListIndex = 0
                
                PonerFocoCmb Combo1(3)
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
    For H = 14 To 14 'imgBuscar.Count - 1
        Me.imgBuscar(H).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next H
     
    For H = 0 To 0 'imgBuscar.Count - 1
        Me.imgBuscar(H).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next H
    
    Set List = Nothing

    'Ocultar todos los Frames de Formulario
    Me.FrameHorasTrabajadas.visible = False
    Me.FramePagoRecibosNatural.visible = False
    
    CargaCombo
    Combo1(0).ListIndex = 0
        
        
    '###Descomentar
'    CommitConexion
    Select Case OpcionListado
        Case 1
            H = 5055
            W = 6660
            FrameHorasTrabajadasVisible True, H, W
            indFrame = 0
            Me.cmdCancel(0).Cancel = True
        Case 2
            H = 7530
            W = 6660
            FramePagoRecibosNaturalVisible True, H, W
            indFrame = 0
            Me.cmdCancel(1).Cancel = True
    End Select
        
    Tabla = "horas"
        
    'Esto se consigue poneinedo el cancel en el opcion k corresponda
    Me.Width = W + 70
    Me.Height = H + 350
    
    Me.Combo1(1).ListIndex = 1
    
    If vParamAplic.Cooperativa = 16 Then
        Check1(1).Caption = "Es anticipo"
    End If
    
    
    Pb1.visible = False
    Pb2.visible = False
End Sub



Private Sub frmBan_DatoSeleccionado(CadenaSeleccion As String)
'Form de Consulta de banco propio
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "00")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmC_Selec(vFecha As Date)
    ' *** repasar si el camp es txtAux o Text1 ***
    txtCodigo(CByte(imgFecha(2).Tag)).Text = Format(vFecha, "dd/mm/yyyy") '<===
    ' ********************************************
End Sub

Private Sub imgBuscar_Click(Index As Integer)
   Select Case Index
        Case 14  'Banco propio
            indCodigo = Index + 4
            AbrirFrmManBanco (Index)
        Case 0
            indCodigo = 0
            AbrirFrmManBanco (Index)
    End Select
    PonerFoco txtCodigo(indCodigo)
End Sub

Private Sub imgFecha_Click(Index As Integer)
    Dim esq As Long
    Dim dalt As Long
    Dim menu As Long
    Dim obj As Object
    Dim Indice As Integer

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

    Select Case Index
        Case 2, 6
            Indice = Index + 14
        Case 0, 1
            Indice = Index + 1
    End Select

    imgFecha(2).Tag = Indice '<===
    ' *** repasar si el camp es txtAux o Text1 ***
    If txtCodigo(Indice).Text <> "" Then frmC.NovaData = txtCodigo(Indice).Text
    ' ********************************************

    frmC.Show vbModal
    Set frmC = Nothing
    ' *** repasar si el camp es txtAux o Text1 ***
    PonerFoco txtCodigo(CByte(imgFecha(2).Tag)) '<===
    ' ********************************************
End Sub

Private Sub ListView1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
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
            Case 1: KEYFecha KeyAscii, 0  'fecha recibo
            Case 2: KEYFecha KeyAscii, 1 'fecha pago
            Case 16: KEYFecha KeyAscii, 2  'fecha recibo
            Case 20: KEYFecha KeyAscii, 6 'fecha pago
            Case 18: KEYBusqueda KeyAscii, 14 'banco
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
    imgFecha_Click (Indice)
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
            
        Case 1, 2, 16, 20   'FECHAS
            If txtCodigo(Index).Text <> "" Then PonerFormatoFecha txtCodigo(Index)
            
        Case 0, 18 ' banco propio
            If txtCodigo(Index).Text <> "" Then PonerFormatoEntero txtCodigo(Index)
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "banpropi", "nombanpr", "codbanpr", "N")
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "000")
            
        Case 3, 4, 5, 6 'porcentajes
            PonerFormatoDecimal txtCodigo(Index), 4
        
            
    End Select
End Sub

Private Sub MandaBusquedaPrevia(CadB As String)
'Carga el formulario frmBuscaGrid con los valores correspondientes
Dim cad As String
Dim Tabla As String
Dim Titulo As String

    'Llamamos a al form
    cad = ""
    Conexion = cAgro    'Conexión a BD: Ariges
'    Select Case OpcionListado
'        Case 7 'Traspaso de Almacenes
'            cad = cad & "Nº Trasp|scatra|codtrasp|N|0000000|40·Almacen Origen|scatra|almaorig|N|000|20·Almacen Destino|scatra|almadest|N|000|20·Fecha|scatra|fechatra|F||20·"
'            Tabla = "scatra"
'            titulo = "Traspaso Almacenes"
'        Case 8 'Movimientos de Almacen
'            cad = cad & "Nº Movim.|scamov|codmovim|N|0000000|40·Almacen|scamov|codalmac|N|000|30·Fecha|scamov|fecmovim|F||30·"
'            Tabla = "scamov"
'            titulo = "Movimientos Almacen"
'        Case 9, 12, 13, 14, 15, 16, 17 '9: Movimientos Articulos
'                   '12: Inventario Articulos
'                   '14:Actualizar Diferencias de Stock Inventariado
'                   '16: Listado Valoracion stock inventariado
'            cad = cad & "Código|sartic|codartic|T||30·Denominacion|sartic|nomartic|T||70·"
'            Tabla = "sartic"
'            titulo = "Articulos"
'    End Select
          
    If cad <> "" Then
        Screen.MousePointer = vbHourglass
        Set frmB = New frmBuscaGrid
        frmB.vCampos = cad
        frmB.vtabla = Tabla
        frmB.vSQL = CadB
        HaDevueltoDatos = False
        '###A mano
        'frmB.vDevuelve = "0|1|"
        frmB.vDevuelve = "0|1|"
        frmB.vTitulo = Titulo
        frmB.vSelElem = 0
'        frmB.vConexionGrid = Conexion
'        frmB.vBuscaPrevia = 1
        '#
        frmB.Show vbModal
        Set frmB = Nothing
        'Si ha puesto valores y tenemos que es formulario de busqueda entonces
        'tendremos que cerrar el form lanzando el evento
'        If HaDevueltoDatos Then
''            If (Not Data1.Recordset.EOF) And DatosADevolverBusqueda <> "" Then _
''                cmdRegresar_Click
'        Else   'de ha devuelto datos, es decir NO ha devuelto datos
'            PonerFoco Text1(kCampo)
'        End If
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub FrameHorasTrabajadasVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Frame para el listado de horas trabajadas
    Me.FrameHorasTrabajadas.visible = visible
    If visible = True Then
        Me.FrameHorasTrabajadas.Top = -90
        Me.FrameHorasTrabajadas.Left = 0
        Me.FrameHorasTrabajadas.Height = H
        Me.FrameHorasTrabajadas.Width = W
        W = Me.FrameHorasTrabajadas.Width
        H = Me.FrameHorasTrabajadas.Height
    End If
End Sub


Private Sub FramePagoRecibosNaturalVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Frame para el listado de horas trabajadas
    Me.FramePagoRecibosNatural.visible = visible
    If visible = True Then
        Me.FramePagoRecibosNatural.Top = -90
        Me.FramePagoRecibosNatural.Left = 0
        Me.FramePagoRecibosNatural.Height = H
        Me.FramePagoRecibosNatural.Width = W
        W = Me.FramePagoRecibosNatural.Width
        H = Me.FramePagoRecibosNatural.Height
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
        .NombreRPT = cadNombreRPT
        .Opcion = 0
        .ConSubInforme = True
        .Show vbModal
    End With
End Sub

Private Function PonerGrupo(numGrupo As Byte, cadgrupo As String) As Byte
Dim campo As String
Dim nomCampo As String

    campo = "pGroup" & numGrupo & "="
    nomCampo = "pGroup" & numGrupo & "Name="
    PonerGrupo = 0

    Select Case cadgrupo
'        Case "Codigo"
'            cadParam = cadParam & campo & "{" & Tabla & ".codclien}" & "|"
'            cadParam = cadParam & nomcampo & " {" & "scoope" & ".nomcoope}" & "|"
'            cadParam = cadParam & "pTitulo1" & "=""Código""" & "|"
'            numParam = numParam + 3
'
'        Case "Alfabetico"
'            cadParam = cadParam & campo & "{" & Tabla & ".tipsocio}" & "|"
'            cadParam = cadParam & nomcampo & " {" & "tiposoci" & ".nomtipso}" & "|"
'            cadParam = cadParam & "pTitulo1" & "=""Colectivo""" & "|"
'            numParam = numParam + 3
'
        
        'Informe de variedades
        Case "Clase"
            CadParam = CadParam & campo & "{" & Tabla & ".codclase}" & "|"
            CadParam = CadParam & nomCampo & " {" & "clases" & ".nomclase}" & "|"
            CadParam = CadParam & "pTitulo1" & "=""Producto""" & "|"
            numParam = numParam + 3
            
        Case "Producto"
            CadParam = CadParam & campo & "{" & Tabla & ".codprodu}" & "|"
            CadParam = CadParam & nomCampo & " {" & "productos" & ".nomprodu}" & "|"
            CadParam = CadParam & "pTitulo1" & "=""Clase""" & "|"
            numParam = numParam + 3

        'Informe de calibres
        Case "Variedad"
            CadParam = CadParam & campo & "{" & Tabla & ".codvarie}" & "|"
            CadParam = CadParam & nomCampo & " {" & "variedades" & ".nomvarie}" & "|"
            CadParam = CadParam & "pTitulo1" & "=""Variedad""" & "|"
            numParam = numParam + 3
            
        Case "Calibre"
            CadParam = CadParam & campo & "{" & Tabla & ".codcalib}" & "|"
            CadParam = CadParam & nomCampo & " {" & "calibres" & ".nomcalib}" & "|"
            CadParam = CadParam & "pTitulo1" & "=""Calibre""" & "|"
            numParam = numParam + 3
            
'        'Informe de Horas Trabajadas
'        Case "Trabajador"
'            cadParam = cadParam & campo & "{" & Tabla & ".codtraba}" & "|"
'            cadParam = cadParam & nomcampo & " {" & "straba" & ".nomtraba}" & "|"
'            cadParam = cadParam & "pTitulo1" & "=""Fecha""" & "|"
'            numParam = numParam + 3
'
'        Case "Fecha"
'            cadParam = cadParam & "pGroup1=" & "{" & Tabla & ".fechahora}" & "|"
'            cadParam = cadParam & "pGroup1Name=" & " {" & "horas" & ".fechahora}" & "|"
'            cadParam = cadParam & "pTitulo1" & "=""Trabajadores""" & "|"
'            numParam = numParam + 3
        

End Select

End Function

Private Function PonerOrden(cadgrupo As String) As Byte
Dim campo As String
Dim nomCampo As String

    PonerOrden = 0

    Select Case cadgrupo
        Case "Codigo"
            CadParam = CadParam & "Orden" & "= {" & Tabla
            Select Case OpcionListado
                Case 10
                    CadParam = CadParam & ".codclien}|"
                Case 11
                    CadParam = CadParam & ".codprove}|"
            End Select
            Tipo = "Código"
        Case "Alfabético"
            CadParam = CadParam & "Orden" & "= {" & Tabla
            Select Case OpcionListado
                Case 10
                    CadParam = CadParam & ".nomclien}|"
                Case 11
                    CadParam = CadParam & ".nomprove}|"
            End Select
            Tipo = "Alfabético"
    End Select
    
    numParam = numParam + 1

End Function

Private Sub AbrirFrmManBanco(Indice As Integer)
    Set frmBan = New frmBasico2
    
    AyudaBancosCom frmBan, txtCodigo(indCodigo)
    
    Set frmBan = Nothing
    
    PonerFoco txtCodigo(indCodigo)
End Sub

Private Sub AbrirVisReport()
    Screen.MousePointer = vbHourglass
    CadenaDesdeOtroForm = ""
    With frmVisReport
        .FormulaSeleccion = cadFormula
        .OtrosParametros = CadParam
        .NumeroParametros = numParam
        .Opcion = OpcionListado
        .Show vbModal
    End With
    
    Unload Me
End Sub

Private Sub AbrirEMail()
    If CadenaDesdeOtroForm <> "" Then frmEMail.Show vbModal
End Sub


' ********* si n'hi han combos a la capçalera ************
Private Sub CargaCombo()
Dim Ini As Integer
Dim Fin As Integer
Dim i As Integer

' *** neteje els combos, els pose valor i seleccione el valor per defecte ***
'    For I = 0 To Combo1.Count - 1
'        Combo1(I).Clear
'    Next I

    Combo1(1).Clear
    
    Combo1(1).AddItem "Campo"
    Combo1(1).ItemData(Combo1(1).NewIndex) = 0
    Combo1(1).AddItem "Almacén"
    Combo1(1).ItemData(Combo1(1).NewIndex) = 1
    
    Combo1(0).Clear
    
    Combo1(0).AddItem "Nómina"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 0
    Combo1(0).AddItem "Pensión"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 1
    Combo1(0).AddItem "Otros Conceptos"
    Combo1(0).ItemData(Combo1(0).NewIndex) = 2
    
    
    Combo1(3).Clear
    
    Combo1(3).AddItem "Campo"
    Combo1(3).ItemData(Combo1(3).NewIndex) = 0
    Combo1(3).AddItem "Almacén"
    Combo1(3).ItemData(Combo1(3).NewIndex) = 1
    
    Combo1(2).Clear
    
    Combo1(2).AddItem "Nómina"
    Combo1(2).ItemData(Combo1(2).NewIndex) = 0
    Combo1(2).AddItem "Pensión"
    Combo1(2).ItemData(Combo1(2).NewIndex) = 1
    Combo1(2).AddItem "Otros Conceptos"
    Combo1(2).ItemData(Combo1(2).NewIndex) = 2
    
    
    
    
End Sub

Private Function DatosOk() As Boolean
Dim B As Boolean
Dim SQL As String
'Dim Datos As String

    On Error GoTo EDatosOK

    B = True

    
    Select Case OpcionListado
        Case 1
            If txtCodigo(16).Text = "" Then
                MsgBox "Debe introducir una Fecha de Recibo.", vbExclamation
                txtCodigo(16).Text = ""
                PonerFoco txtCodigo(16)
                B = False
            End If
            If B And txtCodigo(20).Text = "" Then
                MsgBox "Debe introducir una Fecha de Pago.", vbExclamation
                txtCodigo(20).Text = ""
                PonerFoco txtCodigo(20)
                B = False
            End If
    
        Case 2
            If txtCodigo(1).Text = "" Then
                MsgBox "Debe introducir una Fecha de Recibo.", vbExclamation
                txtCodigo(1).Text = ""
                PonerFoco txtCodigo(1)
                B = False
            End If
            If B And txtCodigo(2).Text = "" Then
                MsgBox "Debe introducir una Fecha de Pago.", vbExclamation
                txtCodigo(2).Text = ""
                PonerFoco txtCodigo(2)
                B = False
            End If
            If B And txtCodigo(4).Text = "" Then
                MsgBox "Debe introducir un porcentaje de Seguridad Social 1.", vbExclamation
                txtCodigo(4).Text = ""
                PonerFoco txtCodigo(4)
            End If
            If B And txtCodigo(5).Text = "" Then
                MsgBox "Debe introducir un porcentaje de Seguridad Social 2.", vbExclamation
                txtCodigo(5).Text = ""
                PonerFoco txtCodigo(5)
            End If
            If B And txtCodigo(6).Text = "" Then
                MsgBox "Debe introducir un porcentaje de IRPF.", vbExclamation
                txtCodigo(6).Text = ""
                PonerFoco txtCodigo(6)
            End If
    
    End Select
    
    
    DatosOk = B
    
EDatosOK:
    If Err.Number <> 0 Then MsgBox Err.Number & ": " & Err.Description, vbExclamation
End Function


Private Function ActualizarRegistros(Tabla As String, cWhere As String) As Boolean
Dim SQL As String
    On Error GoTo eActualizarRegistros
    
    ActualizarRegistros = False
    
    cWhere = QuitarCaracterACadena(cWhere, "{")
    cWhere = QuitarCaracterACadena(cWhere, "}")
    cWhere = QuitarCaracterACadena(cWhere, "_1")

    SQL = "update horas, straba set fecharec = " & DBSet(txtCodigo(20).Text, "F")
    SQL = SQL & " where " & cWhere
    SQL = SQL & " and horas.codtraba = straba.codtraba"
'    (codtraba, fechahora) in (select horas.codtraba, horas.fechahora from " & tabla & " where " & cWhere & ")"
    
    conn.Execute SQL
        
    ActualizarRegistros = True
    
    Exit Function

eActualizarRegistros:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Error en la actualizacion de Registros" & vbCrLf & Err.Description
    End If
End Function

Public Sub BorrarTMP()
On Error Resume Next

    conn.Execute " DROP TABLE IF EXISTS tmpImpor;"
    If Err.Number <> 0 Then Err.Clear
End Sub

Public Function CrearTMP() As Boolean
'Crea una temporal donde inserta la clave primaria de las
'facturas seleccionadas para facturar y trabaja siempre con ellas
Dim SQL As String
    
    On Error GoTo ECrear
    
    CrearTMP = False
    
    SQL = "CREATE TEMPORARY TABLE tmpImpor ( "
    SQL = SQL & "codtraba int(6) unsigned NOT NULL default '0',"
    SQL = SQL & "importe decimal(12,2)  NOT NULL default '0')"
    
    conn.Execute SQL
     
    CrearTMP = True
    
ECrear:
     If Err.Number <> 0 Then
        CrearTMP = False
        'Borrar la tabla temporal
        SQL = " DROP TABLE IF EXISTS tmpImpor;"
        conn.Execute SQL
    End If
End Function

Public Function CopiarFichero() As Boolean
Dim nomFich As String

On Error GoTo ecopiarfichero

    CopiarFichero = False
    ' abrimos el commondialog para indicar donde guardarlo
'    Me.CommonDialog1.InitDir = App.path

    Me.CommonDialog1.DefaultExt = "txt"
    
    CommonDialog1.Filter = "Archivos txt|txt|"
    CommonDialog1.FilterIndex = 1
    
    ' copiamos el primer fichero
    CommonDialog1.FileName = "norma34.txt"
    Me.CommonDialog1.ShowSave
    
    If CommonDialog1.FileName <> "" Then
        FileCopy App.Path & "\norma34.txt", CommonDialog1.FileName
    End If
    
    CopiarFichero = True
    Exit Function

ecopiarfichero:
    If Err.Number <> 0 Then
        MuestraError Err.Number, Err.Description
    End If
    Err.Clear
End Function

Public Function CopiarFicheroA3() As Boolean
Dim nomFich As String

On Error GoTo ecopiarfichero

    CopiarFicheroA3 = False
    ' abrimos el commondialog para indicar donde guardarlo
'    Me.CommonDialog1.InitDir = App.path

    Me.CommonDialog1.DefaultExt = "txt"
    
    CommonDialog1.Filter = "Archivos txt|txt|"
    CommonDialog1.FilterIndex = 1
    
    ' copiamos el primer fichero
    CommonDialog1.FileName = "anticipoA3.txt"
    Me.CommonDialog1.ShowSave
    
    If CommonDialog1.FileName <> "" Then
        FileCopy App.Path & "\anticipoA3.txt", CommonDialog1.FileName
    End If
    
    CopiarFicheroA3 = True
    Exit Function

ecopiarfichero:
    If Err.Number <> 0 Then
        MuestraError Err.Number, Err.Description
    End If
    Err.Clear
End Function





Private Sub ProcesarCambiosPicassent(cadWHERE As String)
Dim SQL As String
Dim Sql2 As String
Dim Sql3 As String
Dim cad As String
Dim i As Integer
Dim HayReg As Integer
Dim B As Boolean
Dim Rs As ADODB.Recordset
Dim Rs2 As ADODB.Recordset
Dim Mens As String

Dim ImpHoras As Currency
Dim ImpHorasE As Currency
Dim ImpBruto As Currency
Dim ImpBruto2 As Currency
Dim IRPF As Currency
Dim SegSoc As Currency
Dim SegSoc1 As Currency
Dim Neto As Currency
Dim Bruto As Currency
Dim Retencion As Currency
Dim CuentaPropia As String
Dim CodigoOrden34 As String

Dim Neto34 As Currency
Dim Bruto34 As Currency
Dim Jornadas As Currency
Dim Diferencia As Currency
Dim BaseSegso As Currency
Dim Complemento As Currency
Dim TSegSoc As Currency
Dim TSegSoc1 As Currency
Dim Max As Long

Dim Sql5 As String
Dim RS5 As ADODB.Recordset

Dim Dias As Long


On Error GoTo eProcesarCambiosPicassent
    
    BorrarTMP
    CrearTMP

    conn.BeginTrans
    
    If cadWHERE <> "" Then
        cadWHERE = QuitarCaracterACadena(cadWHERE, "{")
        cadWHERE = QuitarCaracterACadena(cadWHERE, "}")
        cadWHERE = QuitarCaracterACadena(cadWHERE, "_1")
    End If
        
    SQL = "select count(distinct horas.codtraba) from (horas inner join straba on horas.codtraba = straba.codtraba) inner join forpago on straba.codforpa = forpago.codforpa where " & cadWHERE
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    Pb2.visible = True
    CargarProgres Pb2, Rs.Fields(0).Value
    
    Rs.Close
    
    SQL = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute SQL
    
    Sql3 = "select max(idcontador) from rrecibosnomina"
    Max = DevuelveValor(Sql3) + 1
    
    SQL = "select horas.codtraba, horas.fechahora , sum(if(horasdia is null,0,horasdia)), sum(if(compleme is null,0,compleme)), sum(if(penaliza is null,0,penaliza)), sum(if(importe is null,0,importe)) from (horas inner join straba on horas.codtraba = straba.codtraba) inner join forpago on straba.codforpa = forpago.codforpa where " & cadWHERE
    SQL = SQL & " group by horas.codtraba, horas.fechahora "
    SQL = SQL & " order by 1, 2 "
        
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    Dim AntCodTraba As Long
    Dim ActCodTraba As Long
    Dim TIRPF As Currency
    Dim TImpbruto As Currency
    Dim TImpBruto2 As Currency
    Dim TRetencion As Currency
    Dim TNeto34 As Currency
    Dim TSegSo As Currency
    
    TIRPF = 0
    TImpbruto = 0
    TImpBruto2 = 0
    TRetencion = 0
    TNeto34 = 0
    TSegSo = 0
    
    If Not Rs.EOF Then
        AntCodTraba = DBLet(Rs!CodTraba, "N")
        ActCodTraba = AntCodTraba
        Sql2 = "select salarios.*, straba.dtoreten, straba.dtosegso, straba.dtosirpf, straba.pluscapataz, straba.hayembargo from salarios, straba where straba.codtraba = " & DBSet(Rs!CodTraba, "N")
        Sql2 = Sql2 & " and salarios.codcateg = straba.codcateg "
        
        Set Rs2 = New ADODB.Recordset
        Rs2.Open Sql2, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    End If
    
    HayReg = False
    
    While Not Rs.EOF
        HayReg = True
        
        ActCodTraba = DBLet(Rs!CodTraba, "N")
        
        If AntCodTraba <> ActCodTraba Then
            IncrementarProgres Pb2, 1
            Mens = "Calculando Importes" & vbCrLf & vbCrLf & "Trabajador: " & AntCodTraba & vbCrLf
            
            
            '[Monica]23/03/2016: si el importe es negativo no entra
            If TNeto34 >= 0 Then
        
                Sql3 = "insert into rrecibosnomina (codtraba, fechahora, importe, base34, porcsegso1, porcsegso2, porcirpf, "
                Sql3 = Sql3 & "importesegso1, importesegso2, importeirpf, complemento, neto34, idcontador, hayembargo) values ("
                Sql3 = Sql3 & DBSet(AntCodTraba, "N") & ","
                Sql3 = Sql3 & DBSet(txtCodigo(16).Text, "F") & "," & DBSet(ImporteSinFormato(CStr(TImpbruto)), "N") & ","
                Sql3 = Sql3 & DBSet(ImporteSinFormato(CStr(TImpBruto2)), "N") & ","
                '[Monica]05/01/2012: SegSoc pasa a ser porcentaje
                'Sql3 = Sql3 & DBSet(0, "N") & ","
                Sql3 = Sql3 & DBSet(Rs2!dtosegso, "N") & ","
                Sql3 = Sql3 & DBSet(Rs2!dtoreten, "N") & ","
                Sql3 = Sql3 & DBSet(Rs2!dtosirpf, "N") & ","
                Sql3 = Sql3 & DBSet(TSegSo, "N") & "," & DBSet(TRetencion, "N") & "," & DBSet(TIRPF, "N") & ","
                Sql3 = Sql3 & DBSet(0, "N") & ","
                Sql3 = Sql3 & DBSet(TNeto34, "N") & ","
                Sql3 = Sql3 & DBSet(Max, "N") & "," & DBSet(Rs2!HayEmbargo, "N") & ")"
                
                conn.Execute Sql3
        
                Sql3 = "insert into tmpinformes (codusu, codigo1, fecha1, importe1, importe2) values (" & vUsu.Codigo & "," & DBSet(AntCodTraba, "N") & ","
                Sql3 = Sql3 & DBSet(txtCodigo(16).Text, "F") & "," & DBSet(TNeto34, "N") & "," & DBSet(Rs2!HayEmbargo, "N") & ")"
                
                conn.Execute Sql3
        
        
        
                '[Monica]26/09/2016: si no hay embargo le pagamos
                If DBLet(Rs2!HayEmbargo) = 0 Then
                    
                    Sql3 = "insert into tmpImpor (codtraba, importe) values ("
                    Sql3 = Sql3 & DBSet(AntCodTraba, "N") & "," & DBSet(ImporteSinFormato(CStr(TNeto34)), "N") & ")"
                    
                    conn.Execute Sql3
                End If
            End If
            
            TIRPF = 0
            TImpbruto = 0
            TImpBruto2 = 0
            TRetencion = 0
            TNeto34 = 0
            TSegSo = 0
            
            AntCodTraba = ActCodTraba
            ActCodTraba = DBSet(Rs!CodTraba, "N")
        
            Set Rs2 = Nothing
            
            Sql2 = "select salarios.*, straba.dtoreten, straba.dtosegso, straba.dtosirpf, straba.pluscapataz, straba.hayembargo from salarios, straba where straba.codtraba = " & DBSet(ActCodTraba, "N")
            Sql2 = Sql2 & " and salarios.codcateg = straba.codcateg "
            
            Set Rs2 = New ADODB.Recordset
            Rs2.Open Sql2, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        End If
        
        
        ImpHoras = Round2(DBLet(Rs.Fields(2).Value, "N") * DBLet(Rs2!impsalar, "N"), 2)
                                    ' importe + pluscapataz + complemento - penalizacion
                                    
        If vParamAplic.Cooperativa = 2 Then
            ImpBruto = Round2(ImpHoras + DBLet(Rs.Fields(5).Value, "N") + DBLet(Rs2!PlusCapataz, "N") + DBLet(Rs.Fields(3).Value, "N") - DBLet(Rs.Fields(4).Value, "N"), 2)
        Else
            ' en coopic llevamos en el bruto el plus del capataz
            ' y no hay imphoras
            ImpBruto = Round2(DBLet(Rs.Fields(5).Value, "N") + DBLet(Rs.Fields(3).Value, "N") - DBLet(Rs.Fields(4).Value, "N"), 2)
        End If
        
        TImpbruto = TImpbruto + ImpBruto
        
        IRPF = Round2(ImpBruto * DBLet(Rs2!dtosirpf, "N") / 100, 2)
        TIRPF = TIRPF + IRPF

'[Monica]05/01/2012: SegSoc pasa a ser porcentaje
        SegSoc = Round2(ImpBruto * DBLet(Rs2!dtosegso, "N") / 100, 2)
        
'[Monica]05/01/2012: SegSoc pasa a ser porcentaje
'        ImpBruto2 = ImpBruto - DBLet(Rs2!dtosegso, "N")
        ImpBruto2 = ImpBruto - DBLet(SegSoc, "N")
        TImpBruto2 = TImpBruto2 + ImpBruto2
        
'[Monica]05/01/2012: SegSoc pasa a ser porcentaje
'        TSegSo = TSegSo + DBLet(Rs2!dtosegso, "N")
        TSegSo = TSegSo + SegSoc
        
        Retencion = Round2(ImpBruto2 * DBLet(Rs2!dtoreten, "N") / 100, 2)
        TRetencion = TRetencion + Retencion
        
        Neto34 = ImpBruto2 - IRPF - Retencion
        
        
        TNeto34 = TNeto34 + Neto34
        
        Rs.MoveNext
    Wend
    
    If HayReg Then
        IncrementarProgres Pb2, 1
        Mens = "Calculando Importes" & vbCrLf & vbCrLf & "Trabajador: " & AntCodTraba & vbCrLf
        
        '[Monica]23/03/2016: si el importe es negativo no entra
        If TNeto34 >= 0 Then
            Sql3 = "insert into rrecibosnomina (codtraba, fechahora, importe, base34, porcsegso1, porcsegso2, porcirpf, "
            Sql3 = Sql3 & "importesegso1, importesegso2, importeirpf, complemento, neto34, idcontador, hayembargo) values ("
            Sql3 = Sql3 & DBSet(AntCodTraba, "N") & ","
            Sql3 = Sql3 & DBSet(txtCodigo(16).Text, "F") & "," & DBSet(ImporteSinFormato(CStr(TImpbruto)), "N") & ","
            Sql3 = Sql3 & DBSet(ImporteSinFormato(CStr(TImpBruto2)), "N") & ","
            '[Monica]05/01/2012: SegSoc pasa a ser porcentaje
            'Sql3 = Sql3 & DBSet(0, "N") & ","
            Sql3 = Sql3 & DBSet(Rs2!dtosegso, "N") & ","
            Sql3 = Sql3 & DBSet(Rs2!dtoreten, "N") & ","
            Sql3 = Sql3 & DBSet(Rs2!dtosirpf, "N") & ","
            Sql3 = Sql3 & DBSet(TSegSo, "N") & "," & DBSet(TRetencion, "N") & "," & DBSet(TIRPF, "N") & ","
            Sql3 = Sql3 & DBSet(0, "N") & ","
            Sql3 = Sql3 & DBSet(TNeto34, "N") & ","
            Sql3 = Sql3 & DBSet(Max, "N") & "," & DBSet(Rs2!HayEmbargo, "N") & ")"
            
            conn.Execute Sql3
    
            Sql3 = "insert into tmpinformes (codusu, codigo1, fecha1, importe1, importe2) values (" & vUsu.Codigo & "," & DBSet(AntCodTraba, "N") & ","
            Sql3 = Sql3 & DBSet(txtCodigo(16).Text, "F") & "," & DBSet(TNeto34, "N") & "," & DBSet(Rs2!HayEmbargo, "N") & ")"
            
            conn.Execute Sql3
            
            
            '[Monica]26/09/2016: si no hay embargo le pagamos
            If DBLet(Rs2!HayEmbargo) = 0 Then
                
                Sql3 = "insert into tmpImpor (codtraba, importe) values ("
                Sql3 = Sql3 & DBSet(AntCodTraba, "N") & "," & DBSet(ImporteSinFormato(CStr(TNeto34)), "N") & ")"
                
                conn.Execute Sql3
            End If
        End If
        
        Set Rs2 = Nothing
    End If
    
    Set Rs = Nothing
    '[Monica]22/11/2013: iban
    SQL = "select codbanco, codsucur, digcontr, cuentaba, codorden34, iban from banpropi where codbanpr = " & DBSet(txtCodigo(18).Text, "N")
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    CodigoOrden34 = ""
    
    If Rs.EOF Then
        cad = ""
    Else
        If IsNull(Rs!CodBanco) Then
            cad = ""
        Else
            '[Monica]22/11/2013: iban
            cad = Format(Rs!CodBanco, "0000") & "|" & Format(DBLet(Rs!CodSucur, "T"), "0000") & "|" & DBLet(Rs!digcontr, "T") & "|" & Format(DBLet(Rs!CuentaBa, "T"), "0000000000") & "|" & DBLet(Rs!Iban, "T") & "|"
        End If
        CodigoOrden34 = DBLet(Rs!codorden34, "T")
    End If
    
    Set Rs = Nothing
    
    CuentaPropia = cad
    '[Monica]22/11/2013: iban
    Dim vSeccion As CSeccion
    If vSeccion Is Nothing Then
        Set vSeccion = New CSeccion
        If vSeccion.LeerDatos(vParamAplic.Seccionhorto) Then
            vSeccion.AbrirConta
        End If
    End If
    
    '[Monica]02/02/2018: Catadau ha de generar el fichero
    If vParamAplic.Cooperativa = 2 Or vParamAplic.Cooperativa = 0 Or vParamAplic.Cooperativa = 18 Or vParamAplic.Cooperativa = 19 Then
        If vEmpresa.AplicarNorma19_34Nueva = 1 Then
            If HayXML Then
                B = GeneraFicheroNorma34SEPA_XML(vParam.CifEmpresa, CDate(txtCodigo(20).Text), CuentaPropia, "", "Pago Nómina", Combo1(0).ListIndex, CodigoOrden34)
            Else
                B = GeneraFicheroNorma34SEPA(vParam.CifEmpresa, CDate(txtCodigo(20).Text), CuentaPropia, "", "Pago Nómina", Combo1(0).ListIndex, CodigoOrden34)
            End If
        Else
            B = GeneraFicheroNorma34New(vParam.CifEmpresa, CDate(txtCodigo(20).Text), CuentaPropia, 9, 0, "Pago Nómina", CodigoOrden34, Combo1(0).ListIndex)
        End If
    Else
        ' generamos el fichero plano del anticipo
        B = GeneraFicheroA3(Max, txtCodigo(16).Text)
    End If
    
    vSeccion.CerrarConta
    Set vSeccion = Nothing
    
'antes
'    b = GeneraFicheroNorma34New(vParam.CifEmpresa, CDate(txtcodigo(20).Text), CuentaPropia, 9, 0, "Pago Nómina", CodigoOrden34, Combo1(0).ListIndex)
    If B Then
        Mens = "Copiar fichero"
        '[Monica]02/02/2018: Catadau pasa a funcionar como Picassent
        If vParamAplic.Cooperativa = 2 Or vParamAplic.Cooperativa = 0 Or vParamAplic.Cooperativa = 18 Or vParamAplic.Cooperativa = 19 Then
            CopiarFichero
        Else
            CopiarFicheroA3
        End If
        
        
        If B Then
            CadParam = "|pEmpresa=""" & vEmpresa.nomempre & """|"
            CadParam = CadParam & "pFechaRecibo=""" & txtCodigo(16).Text & """|pFechaPago=""" & txtCodigo(20).Text & """|" & "pImpagados=0|"
            numParam = 4
            cadFormula = "{tmpinformes.codusu}=" & vUsu.Codigo & " and {tmpinformes.importe2} = 0"
            cadNombreRPT = "rListadoPagos.rpt"
            cadTitulo = "Impresion de Pagos"
            ConSubInforme = True

            LlamarImprimir
            
            '[Monica]17/10/2016: impresion de los impagados de Picassent
            SQL = "select count(*) from tmpinformes where codusu = " & vUsu.Codigo & " and importe2 = 1"
            If CInt(DevuelveValor(SQL)) <> 0 Then
                CadParam = "|pEmpresa=""" & vEmpresa.nomempre & """|"
                CadParam = CadParam & "pFechaRecibo=""" & txtCodigo(16).Text & """|pFechaPago=""" & txtCodigo(20).Text & """|" & "pImpagados=1|"
                numParam = 4
                cadFormula = "{tmpinformes.codusu}=" & vUsu.Codigo & " and {tmpinformes.importe2} = 1"
                cadNombreRPT = "rListadoPagos.rpt"
                cadTitulo = "Impresion de Impagos"
                ConSubInforme = True
    
                LlamarImprimir
            End If
            
            If Not Repetir Then
                If MsgBox("¿Proceso realizado correctamente para actualizar?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
                    SQL = "update horas, straba, forpago set horas.intconta = 1 where horas.codtraba = straba.codtraba and straba.codforpa = forpago.codforpa and " & cadWHERE
                    conn.Execute SQL
                Else
                    SQL = "delete from rrecibosnomina where fechahora = " & DBSet(txtCodigo(16).Text, "F")
                    SQL = SQL & " and idcontador = " & DBSet(Max, "N")
                    
                    conn.Execute SQL
                End If
            End If
        Else
            B = False
        End If
    End If

eProcesarCambiosPicassent:
    If Err.Number <> 0 Then
        Mens = Err.Description
        B = False
    End If
    If B Then
        conn.CommitTrans
        MsgBox "Proceso realizado correctamente.", vbExclamation
        cmdCancel_Click (1)
    Else
        conn.RollbackTrans
        MsgBox "Error " & Mens, vbExclamation
    End If
End Sub


Private Function AnticiposPendientes(CodTraba As String) As Currency
Dim SQL As String

    SQL = "select sum(importe) from horasanticipos where codtraba = " & DBSet(CodTraba, "N")
    SQL = SQL & " and descontado = 0 "
    
    AnticiposPendientes = DevuelveValor(SQL)
    
End Function
