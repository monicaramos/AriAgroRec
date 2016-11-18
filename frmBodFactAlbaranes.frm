VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmBodFactAlbaranes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informes"
   ClientHeight    =   6375
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   7530
   Icon            =   "frmBodFactAlbaranes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6375
   ScaleWidth      =   7530
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameFacturar 
      Height          =   6285
      Left            =   30
      TabIndex        =   23
      Top             =   -30
      Width           =   7395
      Begin VB.Frame FrameProgress 
         Height          =   1050
         Left            =   300
         TabIndex        =   49
         Top             =   4980
         Visible         =   0   'False
         Width           =   4695
         Begin MSComctlLib.ProgressBar ProgressBar1 
            Height          =   345
            Left            =   120
            TabIndex        =   50
            Top             =   600
            Width           =   4335
            _ExtentX        =   7646
            _ExtentY        =   609
            _Version        =   393216
            Appearance      =   1
         End
         Begin VB.Label lblProgess 
            Caption         =   "Iniciando el proceso ..."
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   52
            Top             =   350
            Width           =   4335
         End
         Begin VB.Label lblProgess 
            Caption         =   "Facturando:"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   51
            Top             =   135
            Width           =   4215
         End
      End
      Begin VB.Frame Frame4 
         Height          =   4065
         Left            =   300
         TabIndex        =   34
         Top             =   780
         Width           =   6855
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   34
            Left            =   2160
            MaxLength       =   10
            TabIndex        =   54
            Top             =   300
            Width           =   1215
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   42
            Left            =   2940
            Locked          =   -1  'True
            TabIndex        =   46
            Text            =   "Text5"
            Top             =   3210
            Width           =   3735
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   42
            Left            =   2160
            MaxLength       =   6
            TabIndex        =   31
            Tag             =   "Forma Pago|N|N|0|999|scaalb|codforpa|000||"
            Top             =   3210
            Width           =   735
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   41
            Left            =   2160
            MaxLength       =   6
            TabIndex        =   30
            Tag             =   "Cod. Cliente|N|N|0|999999|scaalb|codclien|000000||"
            Top             =   2730
            Width           =   735
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   41
            Left            =   2940
            Locked          =   -1  'True
            TabIndex        =   42
            Text            =   "Text5"
            Top             =   2730
            Width           =   3735
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   40
            Left            =   2160
            MaxLength       =   6
            TabIndex        =   29
            Tag             =   "Cod. Cliente|N|N|0|999999|scaalb|codclien|000000||"
            Top             =   2370
            Width           =   735
         End
         Begin VB.TextBox txtNombre 
            BackColor       =   &H80000018&
            Height          =   285
            Index           =   40
            Left            =   2940
            Locked          =   -1  'True
            TabIndex        =   41
            Text            =   "Text5"
            Top             =   2370
            Width           =   3735
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   38
            Left            =   2160
            MaxLength       =   10
            TabIndex        =   27
            Top             =   1650
            Width           =   1215
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   39
            Left            =   2160
            MaxLength       =   10
            TabIndex        =   28
            Top             =   1980
            Width           =   1215
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   36
            Left            =   2160
            MaxLength       =   10
            TabIndex        =   25
            Top             =   810
            Width           =   1215
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   37
            Left            =   2160
            MaxLength       =   10
            TabIndex        =   26
            Top             =   1170
            Width           =   1215
         End
         Begin VB.Image imgFecha 
            Height          =   240
            Index           =   14
            Left            =   1860
            Picture         =   "frmBodFactAlbaranes.frx":000C
            ToolTipText     =   "Buscar fecha"
            Top             =   300
            Width           =   240
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Factura"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00972E0B&
            Height          =   195
            Index           =   5
            Left            =   240
            TabIndex        =   55
            Top             =   270
            Width           =   1035
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   22
            Left            =   1860
            ToolTipText     =   "Buscar forma pago"
            Top             =   3210
            Width           =   240
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Forma pago"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00972E0B&
            Height          =   195
            Index           =   3
            Left            =   240
            TabIndex        =   47
            Top             =   3180
            Width           =   855
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   20
            Left            =   1860
            ToolTipText     =   "Buscar socio"
            Top             =   2370
            Width           =   240
         End
         Begin VB.Image imgBuscarOfer 
            Height          =   240
            Index           =   21
            Left            =   1860
            ToolTipText     =   "Buscar socio"
            Top             =   2730
            Width           =   240
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Hasta"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   50
            Left            =   1335
            TabIndex        =   45
            Top             =   2730
            Width           =   420
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Desde"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   51
            Left            =   1335
            TabIndex        =   44
            Top             =   2370
            Width           =   450
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Socio"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00972E0B&
            Height          =   195
            Index           =   2
            Left            =   240
            TabIndex        =   43
            Top             =   2220
            Width           =   375
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Hasta"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   37
            Left            =   1350
            TabIndex        =   40
            Top             =   1980
            Width           =   420
         End
         Begin VB.Image imgFecha 
            Height          =   240
            Index           =   12
            Left            =   1860
            Picture         =   "frmBodFactAlbaranes.frx":0097
            ToolTipText     =   "Buscar fecha"
            Top             =   1665
            Width           =   240
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Albarán"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00972E0B&
            Height          =   195
            Index           =   1
            Left            =   240
            TabIndex        =   39
            Top             =   1440
            Width           =   1035
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Desde"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   46
            Left            =   1350
            TabIndex        =   38
            Top             =   1650
            Width           =   450
         End
         Begin VB.Image imgFecha 
            Height          =   240
            Index           =   13
            Left            =   1860
            Picture         =   "frmBodFactAlbaranes.frx":0122
            ToolTipText     =   "Buscar fecha"
            Top             =   1995
            Width           =   240
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Hasta"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   36
            Left            =   1380
            TabIndex        =   37
            Top             =   1170
            Width           =   420
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Nº Albarán"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00972E0B&
            Height          =   195
            Index           =   4
            Left            =   240
            TabIndex        =   36
            Top             =   600
            Width           =   780
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Desde"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   45
            Left            =   1380
            TabIndex        =   35
            Top             =   810
            Width           =   450
         End
      End
      Begin VB.CommandButton cmdAceptarFac 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   5220
         TabIndex        =   32
         Top             =   5670
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   6
         Left            =   6300
         TabIndex        =   33
         Top             =   5670
         Width           =   975
      End
      Begin VB.Label Label10 
         Caption         =   "Facturación de Albaranes Retirada"
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
         Index           =   0
         Left            =   360
         TabIndex        =   24
         Top             =   240
         Width           =   6615
      End
      Begin VB.Label Label10 
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
         Index           =   10
         Left            =   120
         TabIndex        =   53
         Top             =   3360
         Width           =   6855
      End
   End
   Begin VB.Frame FramePreFacturar 
      Height          =   5775
      Left            =   30
      TabIndex        =   0
      Top             =   0
      Width           =   7035
      Begin VB.Frame Frame7 
         Caption         =   "Tipo Informe"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   735
         Left            =   420
         TabIndex        =   48
         Top             =   3930
         Width           =   3135
         Begin VB.OptionButton OptDetalle 
            Caption         =   "Resumen"
            Height          =   255
            Index           =   1
            Left            =   1680
            TabIndex        =   12
            Top             =   300
            Width           =   1335
         End
         Begin VB.OptionButton OptDetalle 
            Caption         =   "Detalle"
            Height          =   255
            Index           =   0
            Left            =   360
            TabIndex        =   10
            Top             =   300
            Width           =   1455
         End
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   26
         Left            =   1710
         MaxLength       =   10
         TabIndex        =   5
         Top             =   2190
         Width           =   1215
      End
      Begin VB.CommandButton cmdAceptarPreFac 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   4080
         TabIndex        =   14
         Top             =   5040
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   5
         Left            =   5160
         TabIndex        =   16
         Top             =   5040
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Index           =   27
         Left            =   3870
         MaxLength       =   10
         TabIndex        =   6
         Top             =   2190
         Width           =   1215
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   30
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   3
         Top             =   1200
         Width           =   945
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   31
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   4
         Top             =   1560
         Width           =   945
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   29
         Left            =   1710
         MaxLength       =   6
         TabIndex        =   8
         Top             =   3270
         Width           =   945
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   29
         Left            =   2700
         Locked          =   -1  'True
         TabIndex        =   2
         Text            =   "Text5"
         Top             =   3270
         Width           =   3615
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   28
         Left            =   1710
         MaxLength       =   6
         TabIndex        =   7
         Top             =   2910
         Width           =   945
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   28
         Left            =   2700
         Locked          =   -1  'True
         TabIndex        =   1
         Text            =   "Text5"
         Top             =   2910
         Width           =   3615
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   44
         Left            =   3090
         TabIndex        =   22
         Top             =   2190
         Width           =   420
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   10
         Left            =   1410
         Picture         =   "frmBodFactAlbaranes.frx":01AD
         Top             =   2190
         Width           =   240
      End
      Begin VB.Label Label9 
         Caption         =   "Prefacturación de Albaranes Retirada"
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
         Index           =   0
         Left            =   390
         TabIndex        =   21
         Top             =   360
         Width           =   6375
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Albaran"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   43
         Left            =   450
         TabIndex        =   20
         Top             =   1950
         Width           =   1035
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   42
         Left            =   885
         TabIndex        =   19
         Top             =   2190
         Width           =   450
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   11
         Left            =   3570
         Picture         =   "frmBodFactAlbaranes.frx":0238
         Top             =   2190
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Nº Albarán"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   41
         Left            =   450
         TabIndex        =   18
         Top             =   960
         Width           =   780
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   40
         Left            =   855
         TabIndex        =   17
         Top             =   1200
         Width           =   450
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   39
         Left            =   855
         TabIndex        =   15
         Top             =   1560
         Width           =   420
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   35
         Left            =   885
         TabIndex        =   13
         Top             =   3270
         Width           =   420
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   15
         Left            =   1410
         Top             =   3300
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   34
         Left            =   885
         TabIndex        =   11
         Top             =   2910
         Width           =   450
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Socio"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   33
         Left            =   450
         TabIndex        =   9
         Top             =   2670
         Width           =   375
      End
      Begin VB.Image imgBuscarOfer 
         Height          =   240
         Index           =   14
         Left            =   1410
         Top             =   2910
         Width           =   240
      End
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   7800
      Top             =   6480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmBodFactAlbaranes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Event DatoSeleccionado(CadenaSeleccion As String)

Public OpcionListado As Integer
'(ver opciones en frmListado)
      
      
      
'Alguna opcion mas
'                   1000.-  Es cuando paso pedido a albaran y este a factura en el mismo proceso
'                   1001.-  Facturar un unico albaran
      
      
Public Tipo As Byte 'Para indicar el tipo de lineas de albaranes que son
                    ' 0 = lineas de almazara
                    ' 1 = lineas de bodega

Public CodClien As String 'Para seleccionar inicialmente las ofertas del Proveedor


'Private HaDevueltoDatos As Boolean
Private NomTabla As String
Private NomTablaLin As String

'Private WithEvents frmMtoCartasOfe As frmFacCartasOferta
Private WithEvents frmSoc As frmManSocios
Attribute frmSoc.VB_VarHelpID = -1
Private WithEvents frmArt As frmADVArticulos
Attribute frmArt.VB_VarHelpID = -1
Private WithEvents frmFPago As frmForpaConta
Attribute frmFPago.VB_VarHelpID = -1
'Private WithEvents frmB As frmBuscaGrid  'Busquedas
Private WithEvents frmF As frmCal 'Calendario de Fechas
Attribute frmF.VB_VarHelpID = -1

'----- Variables para el INforme ----
Private cadFormula As String 'Cadena con la FormulaSelection para Crystal Report
Private CadParam As String
Private numParam As Byte
Private cadSelect As String 'Cadena para comprobar si hay datos antes de abrir Informe
Private Titulo As String 'Titulo informe que se pasa a frmImprimir
Private nomRPT As String 'nombre del fichero .rpt a imprimir
Private conSubRPT As Boolean 'si tiene subinformes para enlazarlos a las tablas correctas
'-------------------------------------

Dim TipCod As String
Dim indCodigo As Integer 'indice para txtCodigo

Dim Codigo As String 'Código para FormulaSelection de Crystal Report
Dim Orden1 As String 'Campo de Ordenacion (por codigo) para Cristal Report
Dim Orden2 As String 'Campo de Ordenacion (por nombre) para Cristal Report


Dim PrimeraVez As Boolean

Dim vSeccion As CSeccion

Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 2, cerrar
    If cerrar Then Unload Me
End Sub


Private Sub chkSoloFacturar_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub cmdAceptarFac_Click()
'Facturacion de Albaranes
Dim campo As String, Cad As String
Dim cadFrom As String
Dim cadSQL As String 'Para seleccionar los Albaranes del rango seleccion
                      'que no se van a facturar
Dim CambiamosConta As Boolean
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta


    
    InicializarVbles
    cadFrom = ""
    CambiamosConta = False
    '--- Comprobar q los campos tienen valor
    If Trim(txtcodigo(34).Text) = "" Then 'Fecha factura
        MsgBox "El campo Fecha Factura debe tener valor.", vbExclamation
        PonerFoco txtcodigo(34)
        Exit Sub
    End If
    If Trim(txtcodigo(42).Text) = "" Then 'la forma de pago debe tener un valor
        MsgBox "El campo Forma de Pago debe tener un valor.", vbExclamation
        PonerFoco txtcodigo(42)
        Exit Sub
    End If
   
    
    '--- Seleccinar los Albaranes que cumplen los criterios introducidos
    'Desde/Hasta Nº ALBARAN
    '-------------------------
    cDesde = Trim(txtcodigo(36).Text)
    cHasta = Trim(txtcodigo(37).Text)
    nDesde = ""
    nHasta = ""
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{rbodalbaran.numalbar}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHAlbaran= """) Then Exit Sub
    End If

    'Desde/Hasta FECHA del ALBARAN
    '--------------------------------------------
    cDesde = Trim(txtcodigo(38).Text)
    cHasta = Trim(txtcodigo(39).Text)
    nDesde = ""
    nHasta = ""
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{rbodalbaran.fechaalb}"
        TipCod = "F"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHFecha= """) Then Exit Sub
    End If
    

    'Cadena para seleccion D/H SOCIO
    '----------------------------------------
    cDesde = Trim(txtcodigo(40).Text)
    cHasta = Trim(txtcodigo(41).Text)
    nDesde = ""
    nHasta = ""
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{rbodalbaran.codsocio}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHSocio= """) Then Exit Sub
    End If
    
    Select Case Tipo
        Case 0 ' almazara
            If Not AnyadirAFormula(cadSelect, "{productos.codgrupo} = 5 ") Then Exit Sub
        Case 1 ' bodega
            If Not AnyadirAFormula(cadSelect, "{productos.codgrupo} = 6 ") Then Exit Sub
    End Select

    
    '--- Comprobar q se han Seleccionados registros de Albaran con esos criterios
    Cad = "Select count(*) " ' & NomTabla & " INNER JOIN " & nomTablaLin
    
    cadFrom = "(((rbodalbaran inner join rsocios on rbodalbaran.codsocio = rsocios.codsocio)"
    cadFrom = cadFrom & " INNER JOIN rbodalbaran_variedad ON rbodalbaran.numalbar = rbodalbaran_variedad.numalbar) "
    cadFrom = cadFrom & " INNER JOIN variedades ON rbodalbaran_variedad.codvarie = variedades.codvarie) "
    cadFrom = cadFrom & " INNER JOIN productos ON variedades.codprodu = productos.codprodu "
    
    
    Cad = Cad & " FROM " & cadFrom

    If Not HayRegParaInforme(cadFrom, cadSelect) Then Exit Sub
    
    
    Cad = Cad & " WHERE " & cadSelect
    'Pasar Albaranes a Facturas
    Cad = Replace(Cad, "count(*)", "*")
   

    Me.Height = Me.Height + 300
    Me.FrameFacturar.Height = Me.FrameFacturar.Height + 300
    Me.FrameProgress.visible = True
'--monica
'    Me.FrameProgress.Top = 6250
    Me.ProgressBar1.Left = 200
    Me.ProgressBar1.Value = 0
    Me.lblProgess(1).Caption = "Inicializando el proceso..."
        
    'proceso normal
    Screen.MousePointer = vbHourglass
     
    '------------------------------------------------------------------------------
    '  LOG de acciones.
    Set LOG = New cLOG
    LOG.Insertar 2, vUsu, ""
    Set LOG = Nothing
    '-----------------------------------------------------------------------------

    campo = "" ' txtCSB(0).Text & "|" & txtCSB(1).Text & "|" & txtCSB(2).Text & "|"
    TraspasoAlbaranesFacturas Cad, cadSelect, txtcodigo(34).Text, "", Me.ProgressBar1, Me.lblProgess(1), True, CodClien, campo, txtcodigo(42).Text, CByte(Tipo)

    Screen.MousePointer = vbDefault
    
    Me.Height = Me.Height - 300
    Me.FrameFacturar.Height = Me.FrameFacturar.Height - 300
    Me.FrameProgress.visible = False
End Sub



'#### Laura 14/11/2006 Recuperar facturas ALZIRA
Private Function ComprobarCliente_RecuperarFac(cadSelAlb As String, FecFac As String, numFac As String) As Boolean
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim codMacta1 As String 'cliente factura ariges
Dim codMacta2 As String 'cliente factura conta
Dim LEtra As String

    On Error GoTo ErrCompCliente
    ComprobarCliente_RecuperarFac = False
    
    'codmacta del cliente del albaran a facturar en Ariges
    Sql = "select scaalb.codclien,sclien.codmacta"
    Sql = Sql & " from scaalb inner join sclien on scaalb.codclien=sclien.codclien "
    Sql = Sql & " Where " & cadSelAlb
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs.EOF Then
        codMacta1 = DBLet(Rs!Codmacta, "T")
    
    End If
    Set Rs = Nothing
    
    
    'codmacta en la contabilidad
    LEtra = ObtenerLetraSerie("FAV")
    Sql = "SELECT codmacta FROM cabfact "
    Sql = Sql & " WHERE numserie=" & DBSet(LEtra, "T") & " AND codfaccl=" & numFac & " AND anofaccl=" & Year(FecFac)
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, ConnConta, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs.EOF Then
        codMacta2 = DBLet(Rs!Codmacta, "T")
    End If
    Set Rs = Nothing
    
    If codMacta1 <> "" And codMacta2 <> "" Then
        If codMacta1 = codMacta2 Then
            ComprobarCliente_RecuperarFac = True
        Else
            ComprobarCliente_RecuperarFac = False
            MsgBox "La cuenta contable en la factura de Contabilidad no coincide con la del cliente del Albaran", vbExclamation
        End If
    Else
        ComprobarCliente_RecuperarFac = False
        MsgBox "No se ha podido leer la cuenta contable del cliente", vbExclamation
    End If
    
    Exit Function
    
ErrCompCliente:
    ComprobarCliente_RecuperarFac = False
    MuestraError Err.Number, "Comprobar cliente", Err.Description
End Function
'#####

Private Sub cmdAceptarPreFac_Click()
'Prevision de Facturacion de Albaranes
Dim campo As String, Cad As String
Dim b As Boolean
Dim indice As Integer
Dim cTabla As String
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta


    InicializarVbles
    b = (OpcionListado = 50)
    
    'Pasar nombre de la Empresa como parametro
    CadParam = "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    '--- Seleccinar los Albaranes que cumplen los criterios introducidos
    'Desde/Hasta Nº ALBARAN
    '-------------------------
'    If txtCodigo(30).Text <> "" Or txtCodigo(31).Text <> "" Then
'        'Para Crystal Report
'        campo = "{rbodalbaran.numalbar}"
'        cad = "pDHAlbaran=""Albarán: "
'        If Not PonerDesdeHasta(campo, "N", 30, 31, cad) Then Exit Sub
'    End If
    cDesde = Trim(txtcodigo(30).Text)
    cHasta = Trim(txtcodigo(31).Text)
    nDesde = ""
    nHasta = ""
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{rbodalbaran.numalbar}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHAlbaran=""Albarán: ") Then Exit Sub
    End If
    
    'Desde/Hasta FECHA del ALBARAN
    '--------------------------------------------
'    If Trim(txtCodigo(26).Text) <> "" Or Trim(txtCodigo(27).Text) <> "" Then
'        'Para Crystal Report
'        campo = "{rbodalbaran.fechaalb}"
'        cad = "pDHFecha=""Fecha: "
'        If Not PonerDesdeHasta(campo, "F", 26, 27, cad) Then Exit Sub
'    End If
    cDesde = Trim(txtcodigo(26).Text)
    cHasta = Trim(txtcodigo(27).Text)
    nDesde = ""
    nHasta = ""
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{rbodalbaran.fechaalb}"
        TipCod = "F"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHFecha= ""Fecha: ") Then Exit Sub
    End If
    




    'Cadena para seleccion SOCIO
    '--------------------------------------------
'    If txtCodigo(28).Text <> "" Or txtCodigo(29).Text <> "" Then
'        campo = "{rbodalbaran.codsocio}"
'        cad = "pDHSocio=""Socio: "
'        If Not PonerDesdeHasta(campo, "N", 28, 29, cad) Then Exit Sub
'    End If
    cDesde = Trim(txtcodigo(28).Text)
    cHasta = Trim(txtcodigo(29).Text)
    nDesde = ""
    nHasta = ""
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{rbodalbaran.codsocio}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHSocio=""Socio: ") Then Exit Sub
    End If



    Select Case Tipo
        Case 0 ' almazara
            If Not AnyadirAFormula(cadSelect, "{productos.codgrupo} = 5 ") Then Exit Sub
            If Not AnyadirAFormula(cadFormula, "{productos.codgrupo} = 5 ") Then Exit Sub
        Case 1 ' bodega
            If Not AnyadirAFormula(cadSelect, "{productos.codgrupo} = 6 ") Then Exit Sub
            If Not AnyadirAFormula(cadFormula, "{productos.codgrupo} = 6 ") Then Exit Sub
    End Select
    
    
    cTabla = "(((rbodalbaran inner join rsocios on rbodalbaran.codsocio = rsocios.codsocio)"
    cTabla = cTabla & " INNER JOIN rbodalbaran_variedad ON rbodalbaran.numalbar = rbodalbaran_variedad.numalbar) "
    cTabla = cTabla & " INNER JOIN variedades ON rbodalbaran_variedad.codvarie = variedades.codvarie) "
    cTabla = cTabla & " INNER JOIN productos ON variedades.codprodu = productos.codprodu "
    
    
    'Comprobar si hay registros a Mostrar antes de abrir el Informe
    If Not HayRegParaInforme(cTabla, cadSelect) Then Exit Sub
    
    
    Titulo = "Previsión Facturación Retirada"
    '-- Si estan activos los servicios hay diferentes posibilidades y el título
    '   las refleja, la variabele 'indice' lleva la información del combo seleccionado y
    '   ha sido cargada un poco más arriba [SERVICIOS]
    conSubRPT = True
    If Me.OptDetalle(0).Value = True Then
        nomRPT = "rFacPrevFactRetDetalle.rpt"
    Else
        nomRPT = "rFacPrevFactRetResum.rpt"
    End If
    
    Cad = "pCodUsu=" & vUsu.Codigo & "|"
    CadParam = CadParam & Cad
    numParam = numParam + 1
    
    '-- Ahora el título depende de los tipos de albaranes seleccionados [SERVICIOS]
    Cad = "pTitulo=""" & Titulo & """|"
    CadParam = CadParam & Cad
    numParam = numParam + 1
    
    
    LlamarImprimir
    
EPreFact:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Informe Prefacturación", Err.Description
    End If
End Sub



Private Sub cmdCancel_Click(Index As Integer)
    Unload Me
     
End Sub


Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        Select Case OpcionListado
            Case 50 '50: Prevision de Facturacion Albaranes (NO IMPRIME LISTADO)
                PonerFoco txtcodigo(30)
            Case 52 '52: Facturacion de Albaranes
                PonerFoco txtcodigo(26)
        End Select
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub Form_Load()
Dim H As Integer, W As Integer
Dim i As Integer
Dim indFrame As Single


    PrimeraVez = True
    limpiar Me

    'Ocultar todos los Frames de Formulario
    Me.FramePreFacturar.visible = False
    Me.FrameFacturar.visible = False
    
    
    For i = 14 To 15
        Me.imgBuscarOfer(i).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next i
    For i = 20 To 22
        Me.imgBuscarOfer(i).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next i

    CommitConexion
    
    ' Necesitamos la conexion a la contabilidad de la seccion de adv
    ' para sacar los porcentajes de iva de los articulos y calcular
    ' los datos de la factura
    Select Case Tipo
        Case 0
            ConexionConta vParamAplic.SeccionAlmaz
        Case 1
            ConexionConta vParamAplic.SeccionBodega
    End Select
    
    NomTabla = "rbodalbaran"
    NomTablaLin = "rbodalbaran_variedad"
        
'    OpcionListado = 52
    
    Select Case OpcionListado
        'LISTADOS DE FACTURACION
        '-----------------------
        Case 50 '50: Prevision Facturacion de Albaranes de Retirada (NO IMPRIME LISTADO)
            PonerFramePreFacVisible True, H, W
            indFrame = 5 'solo para el boton cancelar
        Case 52 '52: Facturacion de Albaranes de Retirada
            PonerFrameFacVisible True, H, W
            txtcodigo(34).Text = Format(Now, "dd/mm/yyyy")
            txtcodigo(39).Text = Format(CDate(txtcodigo(34).Text) - 1, "dd/mm/yyyy")
            indFrame = 6

            NomTabla = "rbodalbaran"
            NomTablaLin = "rbodalbaran_variedad"

    End Select
    
    
    'Esto se consigue poneinedo el cancel en el opcion k corresponda
    Me.cmdCancel(indFrame).Cancel = True
    Me.Width = W + 70
    Me.Height = H + 350
        
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not vSeccion Is Nothing Then
        vSeccion.CerrarConta
        Set vSeccion = Nothing
    End If
End Sub

Private Sub frmF_Selec(vFecha As Date)
'Calendario de Fecha
    txtcodigo(indCodigo).Text = Format(vFecha, "dd/mm/yyyy")
End Sub



Private Sub frmArt_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Articulos
    txtcodigo(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1)
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub


Private Sub frmCli_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Clientes
    txtcodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmFPa_DatoSeleccionado(CadenaSeleccion As String)
'Form de Mantenimiento de Formas de Pabo
    txtcodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub imgBuscarOfer_Click(Index As Integer)
    Select Case Index
            
        Case 14, 15, 20, 21 'Cod. Socio
            Select Case Index
                Case 14, 15: indCodigo = Index + 14
                Case 20, 21: indCodigo = Index + 20
            End Select
            Set frmSoc = New frmManSocios
            frmSoc.DatosADevolverBusqueda = "0|2|"
            If Not IsNumeric(txtcodigo(indCodigo).Text) Then txtcodigo(indCodigo).Text = ""
            frmSoc.Show vbModal
            Set frmSoc = Nothing
            
        Case 22  'Forma de PAGO
            indCodigo = Index + 20
            AbrirFrmForpaConta indCodigo
            
    End Select
    PonerFoco txtcodigo(indCodigo)
End Sub


Private Sub imgFecha_Click(Index As Integer)
   
'++monica

   '++monica
    Dim esq As Long
    Dim dalt As Long
    Dim menu As Long
    Dim obj As Object

    Set frmF = New frmCal
    
    esq = imgFecha(Index).Left
    dalt = imgFecha(Index).Top
    
    Set obj = imgFecha(Index).Container

    While imgFecha(Index).Parent.Name <> obj.Name
        esq = esq + obj.Left
        dalt = dalt + obj.Top
        Set obj = obj.Container
    Wend
    
    menu = Me.Height - Me.ScaleHeight 'ací tinc el heigth del menú i de la toolbar

    frmF.Left = esq + imgFecha(Index).Parent.Left + 30
    frmF.Top = dalt + imgFecha(Index).Parent.Top + imgFecha(Index).Height + menu - 40
   
   frmF.NovaData = Now
   
   Select Case Index
        Case 10 'FramePreFacturar
            indCodigo = 26
        Case 11 'FramePreFacturar
            indCodigo = 27
        Case 12 'Frame Factura
            indCodigo = 38
        Case 13 'Frame Factura
            indCodigo = 39
        Case 14 'FrameFactura
            indCodigo = 34
   
   End Select
   
   PonerFormatoFecha txtcodigo(indCodigo)
   If txtcodigo(indCodigo).Text <> "" Then frmF.NovaData = CDate(txtcodigo(indCodigo).Text)

   Screen.MousePointer = vbDefault
   frmF.Show vbModal
   Set frmF = Nothing
   PonerFoco txtcodigo(indCodigo)

End Sub

Private Sub OptTipoInf_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub OptDetalle_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtCodigo_GotFocus(Index As Integer)
    ConseguirFoco txtcodigo(Index), 3
End Sub

Private Sub txtCodigo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub txtCodigo_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub txtCodigo_LostFocus(Index As Integer)
Dim devuelve As String
Dim codcampo As String, nomCampo As String
Dim Tabla As String
      
    Select Case Index
        Case 1 'Importe (Decimal(12,2))
            PonerFormatoDecimal txtcodigo(Index), 1
            
        
        'FECHA Desde Hasta
        Case 26, 27, 38, 39, 34
            If txtcodigo(Index).Text <> "" Then
                PonerFormatoFecha txtcodigo(Index)
            End If
           
        
        Case 30, 31, 36, 37 'Nº de albaran
            If PonerFormatoEntero(txtcodigo(Index)) Then
                txtcodigo(Index).Text = Format(txtcodigo(Index).Text, "0000000")
            End If
            

        Case 28, 29, 40, 41 'Cod. Socio
            If PonerFormatoEntero(txtcodigo(Index)) Then
                nomCampo = "nomsocio"
                Tabla = "rsocios"
                codcampo = "codsocio"
                txtNombre(Index).Text = PonerNombreDeCod(txtcodigo(Index), Tabla, nomCampo, codcampo, "N")
                If txtcodigo(Index).Text <> "" Then txtcodigo(Index).Text = Format(txtcodigo(Index).Text, "000000")
            Else
                txtNombre(Index).Text = ""
            End If
            
        Case 42 'Cod. Formas de PAGO de la contabilidad de almazara o de bodega
            If PonerFormatoEntero(txtcodigo(Index)) Then
                If vParamAplic.ContabilidadNueva Then
                    txtNombre(Index).Text = PonerNombreDeCod(txtcodigo(Index), "formapago", "nomforpa", "codforpa", "N", cConta)
                Else
                    txtNombre(Index).Text = PonerNombreDeCod(txtcodigo(Index), "sforpa", "nomforpa", "codforpa", "N", cConta)
                End If
                If txtcodigo(Index).Text <> "" Then txtcodigo(Index).Text = Format(txtcodigo(Index).Text, "000")
            Else
                txtNombre(Index).Text = ""
            End If
        
    End Select
End Sub



Private Sub PonerFramePreFacVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Pone el Frame del Prevision Facturacion Albaran Visible y Ajustado al Formulario, y visualiza los controles
Dim b As Boolean
Dim Cad As String

    H = 5600
    If OpcionListado = 51 Then 'Inf. Incum. plazos entrega
        H = 5300
        Me.cmdAceptarPreFac.Top = 4600
        Me.cmdCancel(5).Top = Me.cmdAceptarPreFac.Top
    End If
    W = 7040
    'Ajustar Tamaño del Frame para ajustar tamaño de Formulario al del Frame
    PonerFrameVisible Me.FramePreFacturar, visible, H, W
    If visible = True Then
        b = (OpcionListado = 50)
        Label4(41).visible = b
        Me.txtcodigo(30).visible = b
        Me.txtcodigo(31).visible = b
        'solo albaranes a facturar
        
        'Detalle o resumen
        Me.Frame7.visible = b
        Me.OptDetalle(0).Value = True
        Select Case Tipo
            Case 0 ' almazara
                Me.Label9(0).Caption = "Previsión facturación retirada Almazara"
            Case 1 ' bodega
                Me.Label9(0).Caption = "Previsión facturación retirada Bodega"
        End Select
    End If
End Sub


Private Sub PonerFrameFacVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Pone el Frame de Facturacion de Albaran Visible y Ajustado al Formulario, y visualiza los controles
Dim Cad As String

    H = 6285
    W = 7395
    
    If visible = True Then
         Select Case CodClien 'aqui guardamos el tipo de movimiento
            Case "PAR": Cad = "(ADV)"
                
        End Select
        
        Me.Label10(0).Caption = "Factura de Albaranes de Retirada " & Cad
        Me.Caption = "Facturación"
    End If
    
    PonerFrameVisible Me.FrameFacturar, visible, H, W
End Sub


'Private Function AnyadirParametroDH(cad As String, indD As Byte, indH As Byte) As String
'On Error Resume Next
'
'    If txtCodigo(indD).Text <> "" And txtCodigo(indH).Text <> "" Then
'        If txtCodigo(indD).Text = txtCodigo(indH).Text Then
'            cad = cad & txtCodigo(indD).Text
'            If txtNombre(indD).Text <> "" Then cad = cad & " - " & txtNombre(indD).Text
'            AnyadirParametroDH = cad
'            Exit Function
'        End If
'    End If
'
'    If txtCodigo(indD).Text <> "" Then
'        cad = cad & "desde " & txtCodigo(indD).Text
'        If txtNombre(indD).Text <> "" Then cad = cad & " - " & txtNombre(indD).Text
'    End If
'    If txtCodigo(indH).Text <> "" Then
'        cad = cad & "  hasta " & txtCodigo(indH).Text
'        If txtNombre(indH).Text <> "" Then cad = cad & " - " & txtNombre(indH).Text
'    End If
'    AnyadirParametroDH = cad
'End Function


'Private Function PonerDesdeHasta(campo As String, Tipo As String, indD As Byte, indH As Byte, param As String) As Boolean
'Dim devuelve As String
'
'    PonerDesdeHasta = False
'    devuelve = CadenaDesdeHasta(txtCodigo(indD).Text, txtCodigo(indH).Text, campo, Tipo)
'    If devuelve = "Error" Then Exit Function
'    If Not AnyadirAFormula(cadFormula, devuelve) Then Exit Function
'    If Tipo <> "F" Then
'        If Not AnyadirAFormula(cadSelect, devuelve) Then Exit Function
'    Else
'        devuelve2 = CadenaDesdeHastaBD(codD, codH, Codigo, TipCod)
'        If devuelve2 = "Error" Then Exit Function
'        If Not AnyadirAFormula(cadSelect, devuelve2) Then Exit Function
'    End If
'    If devuelve <> "" Then
'        If param <> "" Then
'            'Parametro Desde/Hasta
'            cadparam = cadparam & AnyadirParametroDH(param, indD, indH) & """|"
'            numParam = numParam + 1
'        End If
'        PonerDesdeHasta = True
'    End If
'End Function

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


Private Sub InicializarVbles()
    cadFormula = ""
    cadSelect = ""
    CadParam = ""
    numParam = 0
End Sub


Private Sub LlamarImprimir()
    With frmImprimir
        .FormulaSeleccion = cadFormula
        .OtrosParametros = CadParam
        .NumeroParametros = numParam
        .SoloImprimir = False
        .EnvioEMail = False
        .Opcion = OpcionListado
        .Titulo = Titulo
        .ConSubInforme = conSubRPT
        .NombreRPT = nomRPT  'nombre del informe
        .Show vbModal
    End With
End Sub

Private Sub txtCodigo_Validate(Index As Integer, Cancel As Boolean)
    Select Case Index
           Case 15, 16 'ARTICULO
            txtNombre(Index).Text = PonerNombreDeCod(txtcodigo(Index), "sartic", "nomartic", "codartic", "Articulo", "T")
            If txtNombre(Index).Text = "" And txtcodigo(Index) <> "" Then Cancel = True
    End Select
End Sub

Private Function ObtenerClientes(cadW As String, Importe As String) As String
Dim Sql As String
Dim Rs As ADODB.Recordset

    On Error GoTo EClientes
    
    cadW = Replace(cadW, "{", "")
    cadW = Replace(cadW, "}", "")
    
    Sql = "select codclien,nomclien,sum(baseimp1),sum(baseimp2),sum(baseimp3),sum(baseimp1)+ sum(if(isnull(baseimp2),0,baseimp2))+ sum(if(isnull(baseimp3),0,baseimp3)) as BaseImp"
    Sql = Sql & " From scafac "
    If cadW <> "" Then Sql = Sql & " where " & cadW
    Sql = Sql & " group by codclien "
    If Importe <> "" Then Sql = Sql & "having baseimp>" & Importe
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Sql = ""
    While Not Rs.EOF
'        If RS!BaseImp >= CCur(Importe) Then
            Sql = Sql & Rs!CodClien & ","
'        End If
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing
    If Sql <> "" Then
        Sql = Mid(Sql, 1, Len(Sql) - 1)
        Sql = "( {scafac.codclien} IN [" & Sql & "] )"
    End If
    ObtenerClientes = Sql
    
EClientes:
   If Err.Number <> 0 Then MuestraError Err.Number, , Err.Description
End Function

Private Sub AbrirFrmForpaConta(indice As Integer)
'    indCodigo = indice
    Set frmFPago = New frmForpaConta
    frmFPago.DatosADevolverBusqueda = "0|1|"
    frmFPago.CodigoActual = txtcodigo(indCodigo)
'    frmFpa.Conexion = cContaFacSoc
    frmFPago.Show vbModal
    Set frmFPago = Nothing
End Sub


Private Sub ConexionConta(Seccion As String)
    
    If vSeccion Is Nothing Then
        Set vSeccion = New CSeccion
        If vSeccion.LeerDatos(Seccion) Then
            vSeccion.AbrirConta
        End If
    Else
        ' si el objeto existia: cerramos la conexion y volvemos crearlo
        vSeccion.CerrarConta
        Set vSeccion = Nothing
        
        
        Set vSeccion = New CSeccion
        If vSeccion.LeerDatos(Seccion) Then
            vSeccion.AbrirConta
        End If
    End If
End Sub

