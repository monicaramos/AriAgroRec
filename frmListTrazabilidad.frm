VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmListTrazabilidad 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informes"
   ClientHeight    =   7365
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   7290
   Icon            =   "frmListTrazabilidad.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7365
   ScaleWidth      =   7290
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameOrigenPaletConf 
      Height          =   4170
      Left            =   0
      TabIndex        =   54
      Top             =   0
      Width           =   6645
      Begin VB.CheckBox Check5 
         Caption         =   "Resumen por variedad"
         Height          =   285
         Left            =   4170
         TabIndex        =   114
         Top             =   2580
         Width           =   1995
      End
      Begin VB.CheckBox Check3 
         Caption         =   "GlobalGap"
         Height          =   285
         Left            =   4170
         TabIndex        =   101
         Top             =   3030
         Width           =   1995
      End
      Begin VB.Frame Frame1 
         Caption         =   "Cálculo por"
         ForeColor       =   &H00972E0B&
         Height          =   705
         Left            =   300
         TabIndex        =   76
         Top             =   3060
         Width           =   2925
         Begin VB.OptionButton Option1 
            Caption         =   "Línea"
            Height          =   225
            Index           =   1
            Left            =   1560
            TabIndex        =   78
            Top             =   300
            Width           =   1095
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Variedad"
            Height          =   315
            Index           =   0
            Left            =   210
            TabIndex        =   77
            Top             =   240
            Width           =   1185
         End
      End
      Begin VB.CommandButton CmdAceptarOri 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   4140
         TabIndex        =   59
         Top             =   3615
         Width           =   975
      End
      Begin VB.CommandButton CmdCancelOri 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   5310
         TabIndex        =   60
         Top             =   3615
         Width           =   975
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   7
         Left            =   1755
         MaxLength       =   7
         TabIndex        =   57
         Tag             =   "Nº Factura|N|S|||facturas|numfactu|0000000|S|"
         Top             =   2115
         Width           =   1065
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   6
         Left            =   1770
         MaxLength       =   7
         TabIndex        =   58
         Tag             =   "Nº Factura|N|S|||facturas|numfactu|0000000|S|"
         Top             =   2670
         Width           =   1065
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   5
         Left            =   1740
         MaxLength       =   10
         TabIndex        =   56
         Top             =   1545
         Width           =   1095
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   4
         Left            =   1740
         MaxLength       =   10
         TabIndex        =   55
         Top             =   1170
         Width           =   1095
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Número de Pedido"
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
         Index           =   0
         Left            =   450
         TabIndex        =   66
         Top             =   2430
         Width           =   1305
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha"
         ForeColor       =   &H00972E0B&
         Height          =   255
         Index           =   3
         Left            =   435
         TabIndex        =   65
         Top             =   945
         Width           =   945
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Palet Confeccionado"
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
         Left            =   465
         TabIndex        =   64
         Top             =   1890
         Width           =   1470
      End
      Begin VB.Label Label3 
         Caption         =   "Listado Origen Palets Confeccionados"
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
         Left            =   495
         TabIndex        =   63
         Top             =   315
         Width           =   5940
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   1
         Left            =   1440
         Picture         =   "frmListTrazabilidad.frx":000C
         ToolTipText     =   "Buscar fecha"
         Top             =   1560
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   0
         Left            =   1440
         Picture         =   "frmListTrazabilidad.frx":0097
         ToolTipText     =   "Buscar fecha"
         Top             =   1200
         Width           =   240
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   11
         Left            =   840
         TabIndex        =   62
         Top             =   1605
         Width           =   420
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   10
         Left            =   840
         TabIndex        =   61
         Top             =   1260
         Width           =   465
      End
   End
   Begin VB.Frame FrameDestinoNotas 
      Height          =   4680
      Left            =   0
      TabIndex        =   67
      Top             =   0
      Width           =   6645
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   14
         Left            =   1770
         MaxLength       =   10
         TabIndex        =   69
         Top             =   2415
         Width           =   1095
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   15
         Left            =   1770
         MaxLength       =   10
         TabIndex        =   70
         Top             =   2820
         Width           =   1095
      End
      Begin VB.Frame Frame2 
         Height          =   705
         Left            =   480
         TabIndex        =   107
         Top             =   780
         Width           =   5475
         Begin VB.OptionButton Option2 
            Caption         =   "Campo"
            Height          =   255
            Index           =   1
            Left            =   3720
            TabIndex        =   109
            Top             =   270
            Width           =   855
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Nota de campo"
            Height          =   255
            Index           =   0
            Left            =   540
            TabIndex        =   108
            Top             =   270
            Width           =   1845
         End
      End
      Begin VB.CheckBox Check4 
         Caption         =   "GlobalGap"
         Height          =   285
         Left            =   420
         TabIndex        =   73
         Top             =   3450
         Width           =   1995
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   8
         Left            =   1785
         MaxLength       =   8
         TabIndex        =   68
         Tag             =   "Nº Factura|N|S|||facturas|numfactu|00000000|S|"
         Top             =   1620
         Width           =   1065
      End
      Begin VB.CommandButton CmdCancelDest 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   4860
         TabIndex        =   72
         Top             =   3765
         Width           =   975
      End
      Begin VB.CommandButton CmdAceptarDest 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3690
         TabIndex        =   71
         Top             =   3765
         Width           =   975
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   8
         Left            =   1470
         Picture         =   "frmListTrazabilidad.frx":0122
         ToolTipText     =   "Buscar fecha"
         Top             =   2400
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   9
         Left            =   1470
         Picture         =   "frmListTrazabilidad.frx":01AD
         ToolTipText     =   "Buscar fecha"
         Top             =   2820
         Width           =   240
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   19
         Left            =   840
         TabIndex        =   113
         Top             =   2835
         Width           =   420
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   18
         Left            =   840
         TabIndex        =   112
         Top             =   2430
         Width           =   465
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha inicio Palet Confeccionado"
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
         Index           =   17
         Left            =   480
         TabIndex        =   111
         Top             =   2040
         Width           =   2340
      End
      Begin VB.Label Label6 
         Caption         =   "Destino de Notas de Campo"
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
         Left            =   495
         TabIndex        =   75
         Top             =   315
         Width           =   5940
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Nota de Campo"
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
         Left            =   480
         TabIndex        =   74
         Top             =   1650
         Width           =   1110
      End
   End
   Begin VB.Frame FramePaletsEntrada 
      Height          =   3870
      Left            =   60
      TabIndex        =   39
      Top             =   30
      Width           =   5685
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   13
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   48
         Top             =   2310
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         Caption         =   "GlobalGap"
         Height          =   225
         Left            =   450
         TabIndex        =   100
         Top             =   3090
         Width           =   2025
      End
      Begin VB.CommandButton CmdCancelPal 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   4170
         TabIndex        =   50
         Top             =   3210
         Width           =   975
      End
      Begin VB.CommandButton CmdAceptarPal 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3090
         TabIndex        =   49
         Top             =   3210
         Width           =   975
      End
      Begin VB.CommandButton Command6 
         Height          =   440
         Left            =   7860
         Picture         =   "frmListTrazabilidad.frx":0238
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   2215
         Visible         =   0   'False
         Width           =   380
      End
      Begin VB.CommandButton Command5 
         Height          =   440
         Left            =   7860
         Picture         =   "frmListTrazabilidad.frx":0542
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   1440
         Visible         =   0   'False
         Width           =   380
      End
      Begin VB.Frame Frame7 
         BorderStyle     =   0  'None
         Height          =   1035
         Left            =   240
         TabIndex        =   40
         Top             =   1020
         Width           =   2565
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   31
            Left            =   1440
            MaxLength       =   10
            TabIndex        =   47
            Top             =   645
            Width           =   1095
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   30
            Left            =   1425
            MaxLength       =   10
            TabIndex        =   46
            Top             =   240
            Width           =   1095
         End
         Begin VB.Image imgFec 
            Height          =   240
            Index           =   5
            Left            =   1110
            Picture         =   "frmListTrazabilidad.frx":084C
            ToolTipText     =   "Buscar fecha"
            Top             =   660
            Width           =   240
         End
         Begin VB.Image imgFec 
            Height          =   240
            Index           =   4
            Left            =   1125
            Picture         =   "frmListTrazabilidad.frx":08D7
            ToolTipText     =   "Buscar fecha"
            Top             =   270
            Width           =   240
         End
         Begin VB.Label Label2 
            Caption         =   "Hasta"
            Height          =   195
            Index           =   2
            Left            =   510
            TabIndex        =   43
            Top             =   645
            Width           =   420
         End
         Begin VB.Label Label2 
            Caption         =   "Desde"
            Height          =   195
            Index           =   1
            Left            =   510
            TabIndex        =   42
            Top             =   300
            Width           =   465
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Fecha"
            ForeColor       =   &H00972E0B&
            Height          =   195
            Index           =   0
            Left            =   180
            TabIndex        =   41
            Top             =   60
            Width           =   450
         End
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Campo"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   16
         Left            =   450
         TabIndex        =   110
         Top             =   2160
         Width           =   495
      End
      Begin VB.Label Label5 
         Caption         =   "Informe de Palets en Entrada"
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
         Left            =   360
         TabIndex        =   51
         Top             =   330
         Width           =   5025
      End
   End
   Begin VB.Frame FrameListadoStocks 
      Height          =   4170
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6645
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   3
         Left            =   1740
         MaxLength       =   10
         TabIndex        =   4
         Top             =   2700
         Width           =   1095
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   2
         Left            =   1740
         MaxLength       =   10
         TabIndex        =   3
         Top             =   2355
         Width           =   1095
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   29
         Left            =   2610
         Locked          =   -1  'True
         TabIndex        =   13
         Text            =   "Text5"
         Top             =   1680
         Width           =   3675
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   28
         Left            =   2625
         Locked          =   -1  'True
         TabIndex        =   12
         Text            =   "Text5"
         Top             =   1290
         Width           =   3675
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   29
         Left            =   1725
         MaxLength       =   7
         TabIndex        =   2
         Tag             =   "Nº Factura|N|S|||facturas|numfactu|0000000|S|"
         Top             =   1665
         Width           =   830
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   28
         Left            =   1725
         MaxLength       =   7
         TabIndex        =   1
         Tag             =   "Nº Factura|N|S|||facturas|numfactu|0000000|S|"
         Top             =   1275
         Width           =   830
      End
      Begin VB.CommandButton CmdCancelStock 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   5310
         TabIndex        =   6
         Top             =   3345
         Width           =   975
      End
      Begin VB.CommandButton CmdAceptarStock 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   4140
         TabIndex        =   5
         Top             =   3345
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   4
         Left            =   840
         TabIndex        =   53
         Top             =   2400
         Width           =   465
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   3
         Left            =   840
         TabIndex        =   52
         Top             =   2745
         Width           =   420
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   3
         Left            =   1440
         Picture         =   "frmListTrazabilidad.frx":0962
         ToolTipText     =   "Buscar fecha"
         Top             =   2730
         Width           =   240
      End
      Begin VB.Image imgFec 
         Height          =   240
         Index           =   2
         Left            =   1440
         Picture         =   "frmListTrazabilidad.frx":09ED
         ToolTipText     =   "Buscar fecha"
         Top             =   2370
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   29
         Left            =   1440
         MouseIcon       =   "frmListTrazabilidad.frx":0A78
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar clase"
         Top             =   1680
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   28
         Left            =   1440
         MouseIcon       =   "frmListTrazabilidad.frx":0BCA
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar clase"
         Top             =   1290
         Width           =   240
      End
      Begin VB.Label Label8 
         Caption         =   "Listado de Stocks"
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
         Left            =   495
         TabIndex        =   11
         Top             =   315
         Width           =   5160
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   26
         Left            =   795
         TabIndex        =   10
         Top             =   1665
         Width           =   420
      End
      Begin VB.Label Label4 
         Caption         =   "Desde"
         Height          =   195
         Index           =   25
         Left            =   795
         TabIndex        =   9
         Top             =   1305
         Width           =   465
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Clase"
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
         Index           =   24
         Left            =   435
         TabIndex        =   8
         Top             =   1050
         Width           =   390
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha"
         ForeColor       =   &H00972E0B&
         Height          =   255
         Index           =   23
         Left            =   435
         TabIndex        =   7
         Top             =   2085
         Width           =   1815
      End
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   6030
      Top             =   5160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame FrameDesviacionAforos 
      Height          =   5220
      Left            =   60
      TabIndex        =   14
      Top             =   30
      Width           =   6285
      Begin VB.CheckBox Check2 
         Caption         =   "Salta página por Socio"
         Height          =   255
         Left            =   630
         TabIndex        =   21
         Top             =   3930
         Width           =   2265
      End
      Begin VB.Frame Frame5 
         Caption         =   "Tipo Hanegadas"
         ForeColor       =   &H00972E0B&
         Height          =   885
         Left            =   330
         TabIndex        =   29
         Top             =   2790
         Width           =   5475
         Begin VB.OptionButton Option4 
            Caption         =   "Cooperativa"
            Height          =   225
            Index           =   0
            Left            =   300
            TabIndex        =   20
            Top             =   390
            Width           =   1305
         End
         Begin VB.OptionButton Option4 
            Caption         =   "Sigpac"
            Height          =   225
            Index           =   1
            Left            =   2040
            TabIndex        =   31
            Top             =   390
            Width           =   1305
         End
         Begin VB.OptionButton Option4 
            Caption         =   "Catastro"
            Height          =   225
            Index           =   2
            Left            =   3630
            TabIndex        =   30
            Top             =   390
            Width           =   1035
         End
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   2430
         Locked          =   -1  'True
         TabIndex        =   28
         Text            =   "Text5"
         Top             =   2070
         Width           =   3375
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   2430
         Locked          =   -1  'True
         TabIndex        =   27
         Text            =   "Text5"
         Top             =   2430
         Width           =   3375
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   1635
         MaxLength       =   3
         TabIndex        =   18
         Top             =   2070
         Width           =   735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   1
         Left            =   1650
         MaxLength       =   3
         TabIndex        =   19
         Top             =   2430
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Height          =   440
         Left            =   7860
         Picture         =   "frmListTrazabilidad.frx":0D1C
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   1440
         Visible         =   0   'False
         Width           =   380
      End
      Begin VB.CommandButton Command2 
         Height          =   440
         Left            =   7860
         Picture         =   "frmListTrazabilidad.frx":1026
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   2215
         Visible         =   0   'False
         Width           =   380
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   9
         Left            =   2460
         Locked          =   -1  'True
         TabIndex        =   22
         Text            =   "Text5"
         Top             =   1110
         Width           =   3375
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   10
         Left            =   2460
         Locked          =   -1  'True
         TabIndex        =   16
         Text            =   "Text5"
         Top             =   1470
         Width           =   3375
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   9
         Left            =   1650
         MaxLength       =   6
         TabIndex        =   15
         Top             =   1110
         Width           =   750
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   10
         Left            =   1650
         MaxLength       =   6
         TabIndex        =   17
         Top             =   1470
         Width           =   750
      End
      Begin VB.CommandButton cmdAceptarDesv 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3780
         TabIndex        =   23
         Top             =   4605
         Width           =   975
      End
      Begin VB.CommandButton cmdCancelDesv 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   4860
         TabIndex        =   25
         Top             =   4605
         Width           =   975
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   1350
         MouseIcon       =   "frmListTrazabilidad.frx":1330
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar clase"
         Top             =   2070
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   1
         Left            =   1350
         MouseIcon       =   "frmListTrazabilidad.frx":1482
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar clase"
         Top             =   2430
         Width           =   240
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   5
         Left            =   705
         TabIndex        =   38
         Top             =   2505
         Width           =   420
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   6
         Left            =   705
         TabIndex        =   37
         Top             =   2115
         Width           =   465
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Clase"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   7
         Left            =   330
         TabIndex        =   36
         Top             =   1860
         Width           =   390
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   9
         Left            =   1320
         MouseIcon       =   "frmListTrazabilidad.frx":15D4
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1140
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   10
         Left            =   1320
         MouseIcon       =   "frmListTrazabilidad.frx":1726
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar socio"
         Top             =   1500
         Width           =   240
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Socio"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   8
         Left            =   330
         TabIndex        =   35
         Top             =   930
         Width           =   405
      End
      Begin VB.Label Label1 
         Caption         =   "Informe de Desviación de Aforos"
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
         Left            =   360
         TabIndex        =   34
         Top             =   330
         Width           =   5025
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   9
         Left            =   660
         TabIndex        =   33
         Top             =   1530
         Width           =   420
      End
      Begin VB.Label Label2 
         Caption         =   "Desde"
         Height          =   195
         Index           =   12
         Left            =   660
         TabIndex        =   32
         Top             =   1170
         Width           =   465
      End
   End
   Begin VB.Frame FrameCargasFecha 
      Height          =   4980
      Left            =   0
      TabIndex        =   79
      Top             =   60
      Width           =   6735
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   63
         Left            =   2580
         Locked          =   -1  'True
         TabIndex        =   106
         Text            =   "Text5"
         Top             =   3660
         Width           =   3735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   63
         Left            =   1650
         MaxLength       =   6
         TabIndex        =   88
         Top             =   3660
         Width           =   855
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   62
         Left            =   2580
         Locked          =   -1  'True
         TabIndex        =   105
         Text            =   "Text5"
         Top             =   3300
         Width           =   3735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   62
         Left            =   1650
         MaxLength       =   6
         TabIndex        =   87
         Top             =   3300
         Width           =   855
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   61
         Left            =   1650
         MaxLength       =   6
         TabIndex        =   86
         Top             =   2685
         Width           =   855
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   61
         Left            =   2580
         Locked          =   -1  'True
         TabIndex        =   96
         Text            =   "Text5"
         Top             =   2685
         Width           =   3735
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   60
         Left            =   1650
         MaxLength       =   6
         TabIndex        =   85
         Top             =   2340
         Width           =   855
      End
      Begin VB.TextBox txtNombre 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   60
         Left            =   2580
         Locked          =   -1  'True
         TabIndex        =   95
         Text            =   "Text5"
         Top             =   2340
         Width           =   3735
      End
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Height          =   1035
         Left            =   240
         TabIndex        =   82
         Top             =   1020
         Width           =   2565
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   12
            Left            =   1440
            MaxLength       =   10
            TabIndex        =   84
            Top             =   630
            Width           =   1095
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   11
            Left            =   1440
            MaxLength       =   10
            TabIndex        =   83
            Top             =   225
            Width           =   1095
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Fecha"
            ForeColor       =   &H00972E0B&
            Height          =   195
            Index           =   15
            Left            =   180
            TabIndex        =   93
            Top             =   60
            Width           =   450
         End
         Begin VB.Label Label2 
            Caption         =   "Desde"
            Height          =   195
            Index           =   14
            Left            =   510
            TabIndex        =   91
            Top             =   300
            Width           =   465
         End
         Begin VB.Label Label2 
            Caption         =   "Hasta"
            Height          =   195
            Index           =   13
            Left            =   510
            TabIndex        =   89
            Top             =   645
            Width           =   420
         End
         Begin VB.Image imgFec 
            Height          =   240
            Index           =   7
            Left            =   1125
            Picture         =   "frmListTrazabilidad.frx":1878
            ToolTipText     =   "Buscar fecha"
            Top             =   630
            Width           =   240
         End
         Begin VB.Image imgFec 
            Height          =   240
            Index           =   6
            Left            =   1110
            Picture         =   "frmListTrazabilidad.frx":1903
            ToolTipText     =   "Buscar fecha"
            Top             =   240
            Width           =   240
         End
      End
      Begin VB.CommandButton Command8 
         Height          =   440
         Left            =   7860
         Picture         =   "frmListTrazabilidad.frx":198E
         Style           =   1  'Graphical
         TabIndex        =   81
         Top             =   1440
         Visible         =   0   'False
         Width           =   380
      End
      Begin VB.CommandButton Command7 
         Height          =   440
         Left            =   7860
         Picture         =   "frmListTrazabilidad.frx":1C98
         Style           =   1  'Graphical
         TabIndex        =   80
         Top             =   2215
         Visible         =   0   'False
         Width           =   380
      End
      Begin VB.CommandButton CmdAcepCargasFecha 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   4110
         TabIndex        =   90
         Top             =   4110
         Width           =   975
      End
      Begin VB.CommandButton CmdCancelCarF 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   5190
         TabIndex        =   92
         Top             =   4110
         Width           =   975
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   5
         Left            =   1350
         ToolTipText     =   "Buscar variedad"
         Top             =   3660
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   4
         Left            =   1350
         ToolTipText     =   "Buscar variedad"
         Top             =   3300
         Width           =   240
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Variedad"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   2
         Left            =   420
         TabIndex        =   104
         Top             =   3090
         Width           =   630
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         Height          =   195
         Index           =   1
         Left            =   750
         TabIndex        =   103
         Top             =   3330
         Width           =   465
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
         Height          =   195
         Index           =   0
         Left            =   750
         TabIndex        =   102
         Top             =   3675
         Width           =   420
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
         Height          =   195
         Index           =   7
         Left            =   750
         TabIndex        =   99
         Top             =   2685
         Width           =   420
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   3
         Left            =   1350
         ToolTipText     =   "Buscar producto"
         Top             =   2685
         Width           =   240
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         Height          =   195
         Index           =   6
         Left            =   750
         TabIndex        =   98
         Top             =   2340
         Width           =   465
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   2
         Left            =   1350
         ToolTipText     =   "Buscar producto"
         Top             =   2340
         Width           =   240
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Producto"
         ForeColor       =   &H00972E0B&
         Height          =   195
         Index           =   5
         Left            =   420
         TabIndex        =   97
         Top             =   2100
         Width           =   645
      End
      Begin VB.Label Label7 
         Caption         =   "Informe Cargas por Fecha/Producto"
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
         Left            =   360
         TabIndex        =   94
         Top             =   330
         Width           =   5895
      End
   End
End
Attribute VB_Name = "frmListTrazabilidad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' +-+-+-+-+-+-+-+-+-+-+-
' +-+- Autor: LAURA +-+-
' +-+-+-+-+-+-+-+-+-+-+-

Option Explicit

Public OpcionListado As Byte
    '==== Listados / Procesos TOMA DE DATOS ====
    '=============================
    ' 1 .- Informe de Palets en Entrada
    ' 2 .- Informe Detalle de cargas en lineas de confeccion
    ' 3 .- Informe de origenes del palet confeccionado
    ' 4 .- Informe de Destino Albaranes de Venta
    ' 5 .- Informe de Listado de Stocks
    ' 6 .- Manejo de Cargas de Confeccion
    ' 7 .- Cargas en linea de confeccion por fecha
    
Public NumCod As String 'Para indicar cod. Traspaso,Movimiento, etc. que llama
                        'Para indicar nº oferta a imprimir

Public CadTag As String 'Cadena con el Tag del campo que se va a poner en D/H en los listados
                        'Se necesita si el tipo de codigo es texto

Public Event RectificarFactura(Cliente As String, Observaciones As String)

Private Conexion As Byte
'1.- Conexión a BD Ariges  2.- Conexión a BD Conta

Private HaDevueltoDatos As Boolean


Private WithEvents frmCar As frmTrzManCargas 'mantenimiento de manejo de cargas de confeccion
Attribute frmCar.VB_VarHelpID = -1

Private WithEvents frmSec As frmManSeccion 'Secciones
Attribute frmSec.VB_VarHelpID = -1
Private WithEvents frmSoc As frmManSocios 'Socios
Attribute frmSoc.VB_VarHelpID = -1
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
Private WithEvents frmCla As frmComercial 'Ayuda de Clases de comercial
Attribute frmCla.VB_VarHelpID = -1
Private WithEvents frmMens As frmMensajes 'Mensajes
Attribute frmMens.VB_VarHelpID = -1
Private WithEvents frmArea As frmTrzAreas 'Mensajes
Attribute frmArea.VB_VarHelpID = -1
Private WithEvents frmProd As frmComercial 'Productos
Attribute frmProd.VB_VarHelpID = -1


'GENERALES PARA PASARLE A CRYSTAL REPORT
Private cadFormula As String 'Cadena con la FormulaSelection para Crystal Report
Private cadParam As String 'Cadena con los parametros para Crystal Report
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
Dim tipo As String

Dim indice As Integer

Dim PrimeraVez As Boolean
Dim Contabilizada As Byte

Private Sub KEYpress(KeyAscii As Integer)
    If KeyAscii = 13 Then 'ENTER
        KeyAscii = 0
        SendKeys "{tab}"
    ElseIf KeyAscii = 27 Then Unload Me  'ESC
    End If
End Sub


Private Sub Check2_KeyDown(KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub Check2_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub CmdAcepCargasFecha_Click()
Dim cOrden As String
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim numOp As Byte

Dim cadDesde As Date
Dim cadhasta As Date
Dim cadAux As String

Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String
Dim nTabla As String

Dim nRegs As Long
Dim FecFac As Date

Dim b As Boolean
Dim TipoPrec As Byte

    InicializarVbles
    
    'Añadir el parametro de Empresa
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    If DatosOk Then
        '======== FORMULA  ====================================
        'D/H fecha
        cDesde = Trim(txtCodigo(11).Text)
        cHasta = Trim(txtCodigo(12).Text)
        nDesde = ""
        nHasta = ""
        If Not (cDesde = "" And cHasta = "") Then
            'Cadena para seleccion Desde y Hasta
            Codigo = "{trzlineas_cargas.fecha}"
            TipCod = "F"

            If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHFecha=""") Then Exit Sub
        End If
        
        
        'D/H producto
        cDesde = Trim(txtCodigo(60).Text)
        cHasta = Trim(txtCodigo(61).Text)
        nDesde = ""
        nHasta = ""
        If Not (cDesde = "" And cHasta = "") Then
            'Cadena para seleccion Desde y Hasta
            Codigo = "{productos.codprodu}"
            TipCod = "N"

            If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHProducto=""") Then Exit Sub
        End If
                
        'D/H variedades
        cDesde = Trim(txtCodigo(62).Text)
        cHasta = Trim(txtCodigo(63).Text)
        nDesde = ""
        nHasta = ""
        If Not (cDesde = "" And cHasta = "") Then
            'Cadena para seleccion Desde y Hasta
            Codigo = "{variedades.codvarie}"
            TipCod = "N"

            If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHVariedad=""") Then Exit Sub
        End If
        
        
        nTabla = "((trzlineas_cargas INNER JOIN trzpalets ON trzpalets.idpalet = trzlineas_cargas.idpalet)"
        nTabla = nTabla & " INNER JOIN variedades ON trzpalets.codvarie = variedades.codvarie) "
        nTabla = nTabla & " INNER JOIN productos ON variedades.codprodu = productos.codprodu "
       
        
        cadNombreRPT = "rTrzCargasFechaProd.rpt"
        cadTitulo = "Informe Cargas por Fecha / Producto"
        ConSubInforme = False
        
        
       'Comprobar si hay registros a Mostrar antes de abrir el Informe
        If HayRegParaInforme(nTabla, cadSelect) Then
            LlamarImprimir
        End If
   End If

End Sub


Private Sub CmdAceptarDest_Click()
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim cadTabla As String, cOrden As String
Dim i As Byte
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String
Dim Tipos As String
Dim nTabla As String

Dim vSQL As String

    InicializarVbles
    
    If Not DatosOk Then Exit Sub
    
    '========= PARAMETROS  =============================
    'Añadir el parametro de Empresa
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
     '======== FORMULA  ====================================
    'Seleccionar registros de la empresa conectada
'    Codigo = "{" & tabla & ".codempre}=" & vEmpresa.codEmpre
'    If Not AnyadirAFormula(cadFormula, Codigo) Then Exit Sub
'    If Not AnyadirAFormula(cadSelect, Codigo) Then Exit Sub
    
    
     'D/H fecha
     cDesde = Trim(txtCodigo(4).Text)
     cHasta = Trim(txtCodigo(5).Text)
     nDesde = ""
     nHasta = ""
     If Not (cDesde = "" And cHasta = "") Then
         'Cadena para seleccion Desde y Hasta
         Codigo = "{" & Tabla & ".fecha}"
         TipCod = "F"

         If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHFecha=""") Then Exit Sub
    End If
    
    If CargarTemporalDestinos() Then
        If HayRegistros("trztmp_palets_lineas_cargas", "codusu=" & vUsu.Codigo) Then
            'Nombre fichero .rpt a Imprimir
            cadNombreRPT = "rTrzDesAlbEnt.rpt"
            '[Monica]14/11/2011:globalgap
            If Check4.Value Then cadNombreRPT = "rTrzDesAlbEntGGap.rpt"
            
            '[Monica]05/02/2014: listado de destino por campo
            If Option2(0).Value Then
                cadTitulo = "Listado Destino de Notas de Campo"
                cadParam = cadParam & "pTitulo=""Destino de Albaranes de Entrada""|"
                cadParam = cadParam & "pTipo=0|"
                
            Else
                cadTitulo = "Listado Destino de Campos"
                cadParam = cadParam & "pTitulo=""Destino de Campos""|"
                cadParam = cadParam & "pTipo=1|"
            End If
            numParam = numParam + 2
            
            ConSubInforme = True
            cadFormula = "{trztmp_palets_lineas_cargas.codusu}=" & vUsu.Codigo
            LlamarImprimir
        End If
    End If


End Sub

Private Sub CmdAceptarOri_Click()
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim cadTabla As String, cOrden As String
Dim i As Byte
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String
Dim Tipos As String
Dim nTabla As String

Dim vSQL As String

    InicializarVbles
    
    If Not DatosOk Then Exit Sub
    
    '========= PARAMETROS  =============================
    'Añadir el parametro de Empresa
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
     '======== FORMULA  ====================================
    'Seleccionar registros de la empresa conectada
'    Codigo = "{" & tabla & ".codempre}=" & vEmpresa.codEmpre
'    If Not AnyadirAFormula(cadFormula, Codigo) Then Exit Sub
'    If Not AnyadirAFormula(cadSelect, Codigo) Then Exit Sub
    
    
     'D/H fecha
     cDesde = Trim(txtCodigo(4).Text)
     cHasta = Trim(txtCodigo(5).Text)
     nDesde = ""
     nHasta = ""
     If Not (cDesde = "" And cHasta = "") Then
         'Cadena para seleccion Desde y Hasta
         Codigo = "{" & Tabla & ".fecha}"
         TipCod = "F"

         If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHFecha=""") Then Exit Sub
     End If
     
     '[Monica]08/04/2015: nuevo informe solo para catadau donde dadas 2 fechas saca por variedades agrupadas los kilos y la fecha de abocamiento
     If vParamAplic.Cooperativa = 0 And Check5.Value Then
        If CargarTemporalAbocamiento Then
            If HayRegistros("trztmp_palets_lineas_cargas", "codusu=" & vUsu.Codigo) Then
                cadTitulo = "Resumen Origenes Palets Confeccionados"
                
                cadNombreRPT = "rTrzOrigenPaletConfResumen.rpt"
                  
                ConSubInforme = False
                cadFormula = "{trztmp_palets_lineas_cargas.codusu}=" & vUsu.Codigo
                LlamarImprimir
                Exit Sub
            End If
        End If
     
     End If
     
    
    If CargarTemporal(txtCodigo(7).Text, txtCodigo(6).Text) Then
        If HayRegistros("trztmp_palets_lineas_cargas", "codusu=" & vUsu.Codigo) Then
            'Nombre fichero .rpt a Imprimir
            '[Monica] 24/05/2010 si es por variedad
            If Option1(0).Value Then
                cadNombreRPT = "rTrzOrigenPaletConf.rpt"
            Else
                cadNombreRPT = "rTrzOrigenPaletConf1.rpt"
            End If
            
            '[Monica]14/11/2011: globalgap
            If Me.Check3.Value Then cadNombreRPT = Replace(cadNombreRPT, ".rpt", "GGap.rpt")
            
            cadTitulo = "Listado Origenes de Palets Confeccionados"
              
            ConSubInforme = False
            cadFormula = "{trztmp_palets_lineas_cargas.codusu}=" & vUsu.Codigo
            LlamarImprimir
        End If
    End If

End Sub

Private Sub CmdAceptarStock_Click()
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim cadTabla As String, cOrden As String
Dim i As Byte
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String
Dim Tipos As String
Dim nTabla As String

Dim vSQL As String

    InicializarVbles
    
    If Not DatosOk Then Exit Sub
    
    '========= PARAMETROS  =============================
    'Añadir el parametro de Empresa
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
     '======== FORMULA  ====================================
    'Seleccionar registros de la empresa conectada
'    Codigo = "{" & tabla & ".codempre}=" & vEmpresa.codEmpre
'    If Not AnyadirAFormula(cadFormula, Codigo) Then Exit Sub
'    If Not AnyadirAFormula(cadSelect, Codigo) Then Exit Sub
    
    
    'D/H Clase
    cDesde = Trim(txtCodigo(28).Text)
    cHasta = Trim(txtCodigo(29).Text)
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{variedades.codclase}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHClase= """) Then Exit Sub
    End If
    
    vSQL = ""
    If txtCodigo(28).Text <> "" Then vSQL = vSQL & " and variedades.codclase >= " & DBSet(txtCodigo(28).Text, "N")
    If txtCodigo(29).Text <> "" Then vSQL = vSQL & " and variedades.codclase <= " & DBSet(txtCodigo(29).Text, "N")
    
     'D/H fecha
     cDesde = Trim(txtCodigo(2).Text)
     cHasta = Trim(txtCodigo(3).Text)
     nDesde = ""
     nHasta = ""
     If Not (cDesde = "" And cHasta = "") Then
         'Cadena para seleccion Desde y Hasta
         Codigo = "{" & Tabla & ".fecha}"
         TipCod = "F"

         If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHFecha=""") Then Exit Sub
    End If
    
'    NTabla = "((trzpalets INNER JOIN variedades ON trzpalets.codvarie = variedades.codvarie) "
'    NTabla = NTabla & " INNER JOIN trzareas ON trzpalets.idarea = trzareas.codarea) "
'    NTabla = NTabla & " INNER JOIN rsocios ON trzpalets.codsocio = rsocios.codsocio "
    nTabla = "(trzpalets INNER JOIN variedades ON trzpalets.codvarie = variedades.codvarie) "
    
    Set frmMens = New frmMensajes
    
    frmMens.OpcionMensaje = 16
    frmMens.cadWHERE = vSQL
    frmMens.Show vbModal
    
    Set frmMens = Nothing
    
    ' seleccionamos solo los que tienen CRFID asignado
    If Not AnyadirAFormula(cadFormula, "not isnull({trzpalets.CRFID}) and {trzpalets.CRFID} <> ''") Then Exit Sub
    If Not AnyadirAFormula(cadSelect, "not trzpalets.CRFID is null and {trzpalets.CRFID} <> ''") Then Exit Sub
    
    If HayRegistros(nTabla, cadSelect) Then
        'Nombre fichero .rpt a Imprimir
        cadNombreRPT = "rTrzPaletsStock.rpt"
        cadTitulo = "Listado de Stocks"
          
        ConSubInforme = False
        
        LlamarImprimir
    End If


End Sub

Private Sub cmdAceptarDesv_Click()
Dim cOrden As String
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim numOp As Byte

Dim cadDesde As Date
Dim cadhasta As Date
Dim cadAux As String

Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String
Dim nTabla As String

Dim nRegs As Long
Dim FecFac As Date

Dim b As Boolean
Dim TipoPrec As Byte

Dim vSQL As String

    InicializarVbles
    
    'Añadir el parametro de Empresa
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    If DatosOk Then
        '======== FORMULA  ====================================
        'D/H SOCIO
        cDesde = Trim(txtCodigo(9).Text)
        cHasta = Trim(txtCodigo(10).Text)
        nDesde = txtNombre(9).Text
        nHasta = txtNombre(10).Text
        If Not (cDesde = "" And cHasta = "") Then
            'Cadena para seleccion Desde y Hasta
            Codigo = "{" & Tabla & ".codsocio}"
            TipCod = "N"
            If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHSocio=""") Then Exit Sub
        End If
        
        'D/H CLASE
        cDesde = Trim(txtCodigo(0).Text)
        cHasta = Trim(txtCodigo(1).Text)
        nDesde = txtNombre(0).Text
        nHasta = txtNombre(1).Text
        If Not (cDesde = "" And cHasta = "") Then
            'Cadena para seleccion Desde y Hasta
            Codigo = "{variedades.codclase}"
            TipCod = "N"
            If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHClase=""") Then Exit Sub
        End If
        
        vSQL = ""
        If txtCodigo(0).Text <> "" Then vSQL = vSQL & " and variedades.codclase >= " & DBSet(txtCodigo(0).Text, "N")
        If txtCodigo(1).Text <> "" Then vSQL = vSQL & " and variedades.codclase <= " & DBSet(txtCodigo(1).Text, "N")
        
        
        'SECCION
        If Not AnyadirAFormula(cadSelect, "{rsocios_seccion.codsecci} = " & vParamAplic.Seccionhorto) Then Exit Sub
        If Not AnyadirAFormula(cadFormula, "{rsocios_seccion.codsecci} = " & vParamAplic.Seccionhorto) Then Exit Sub
        
        'CAMPOS DADOS DE ALTA
        If Not AnyadirAFormula(cadSelect, "{rcampos.fecbajas} is null ") Then Exit Sub
        If Not AnyadirAFormula(cadFormula, "isnull({rcampos.fecbajas})") Then Exit Sub
        
        nTabla = "(rcampos INNER JOIN variedades ON rcampos.codvarie = variedades.codvarie) "
        nTabla = nTabla & " INNER JOIN rsocios_seccion ON rcampos.codsocio = rsocios_seccion.codsocio "

        cadNombreRPT = "rInfDesvAfo.rpt"
        cadTitulo = "Informe de Desviación de Aforos"
        
        'tipo de hanegada
        If Option4(0).Value Then cadParam = cadParam & "pTipoHa=0|"
        If Option4(1).Value Then cadParam = cadParam & "pTipoHa=1|"
        If Option4(2).Value Then cadParam = cadParam & "pTipoHa=2|"
        numParam = numParam + 1
             
        If Check2.Value Then
            cadParam = cadParam & "pSaltoSocio=1|"
        Else
            cadParam = cadParam & "pSaltoSocio=0|"
        End If
        numParam = numParam + 1
             
        Set frmMens = New frmMensajes
        
        frmMens.OpcionMensaje = 16
        frmMens.cadWHERE = vSQL
        frmMens.Show vbModal
        
        Set frmMens = Nothing
        
        'Comprobar si hay registros a Mostrar antes de abrir el Informe
        If HayRegParaInforme(nTabla, cadSelect) Then
            ConSubInforme = False
            LlamarImprimir
        End If
    End If


End Sub

Private Sub CmdAceptarPal_Click()
Dim cOrden As String
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim numOp As Byte

Dim cadDesde As Date
Dim cadhasta As Date
Dim cadAux As String

Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim devuelve As String
Dim nTabla As String

Dim nRegs As Long
Dim FecFac As Date

Dim b As Boolean
Dim TipoPrec As Byte

    InicializarVbles
    
    'Añadir el parametro de Empresa
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    If DatosOk Then
        '======== FORMULA  ====================================
        'D/H fecha
        cDesde = Trim(txtCodigo(30).Text)
        cHasta = Trim(txtCodigo(31).Text)
        nDesde = ""
        nHasta = ""
        If Not (cDesde = "" And cHasta = "") Then
            'Cadena para seleccion Desde y Hasta
            Codigo = "{" & Tabla & ".fecha}"
            TipCod = "F"

            If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHFecha=""") Then Exit Sub
       End If
        
       Select Case OpcionListado
            Case 1
                '[Monica]06/02/2014: insertamos para poder buscar por campo
                If txtCodigo(13).Text <> "" Then
                    If Not AnyadirAFormula(cadSelect, "{trzpalets.codcampo} = " & DBSet(txtCodigo(13).Text, "N")) Then Exit Sub
                    If Not AnyadirAFormula(cadFormula, "{trzpalets.codcampo} = " & DBSet(txtCodigo(13).Text, "N")) Then Exit Sub
                End If
            
                '[Monica]14/11/2011: globalgap
                If Me.Check1.Value Then
                    cadNombreRPT = "rTrzPaletsEntradosGGap.rpt"
                Else
                    cadNombreRPT = "rTrzPaletsEntrados.rpt"
                End If
                cadTitulo = "Informe de Palets Entrados"
                nTabla = "trzpalets"
                ConSubInforme = False
            Case 2
                '[Monica]14/11/2011: globalgap
                If Me.Check1.Value Then
                    cadNombreRPT = "rTrzCargasLineasGGap.rpt"
                Else
                    cadNombreRPT = "rTrzCargasLineas.rpt"
                End If
                cadTitulo = "Informe Detalle Cargas en Lineas Confección"
                nTabla = "trzlineas_cargas"
                ConSubInforme = True
                
            Case 6
                Set frmCar = New frmTrzManCargas
                
                frmCar.FechaCarga = txtCodigo(30).Text
                frmCar.Show vbModal
                Set frmCar = Nothing
                
                Set frmCar = Nothing
            
                Exit Sub
       End Select
        
        
       'Comprobar si hay registros a Mostrar antes de abrir el Informe
       If HayRegParaInforme(nTabla, cadSelect) Then
            LlamarImprimir
       End If
   End If

End Sub



Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub CmdCancelCarF_Click()
    Unload Me
End Sub

Private Sub CmdCancelDest_Click()
    Unload Me
End Sub

Private Sub cmdCancelDesv_Click()
    Unload Me
End Sub

Private Sub CmdCancelOri_Click()
    Unload Me
End Sub

Private Sub CmdCancelPal_Click()
    Unload Me
End Sub

Private Sub CmdCancelResul_Click()
    Unload Me
End Sub


Private Sub CmdCancelStock_Click()
    Unload Me
End Sub

Private Sub Command4_Click()

End Sub

Private Sub Command3_Click()

End Sub

Private Sub Form_Activate()
   If PrimeraVez Then
        PrimeraVez = False
        Select Case OpcionListado
            Case 1, 2 ' 1-Informe de Palets entrados
                      ' 2-Informe de detalle de cargas en lineas de confeccion
                txtCodigo(30).Text = Format(Now, "dd/mm/yyyy")
                txtCodigo(31).Text = Format(Now, "dd/mm/yyyy")
                
                PonerFoco txtCodigo(30)
                
            Case 3 ' 3-Informe de origen de palets confeccionados
                txtCodigo(4).Text = Format(Now, "dd/mm/yyyy")
                txtCodigo(5).Text = Format(Now, "dd/mm/yyyy")
                
                Option1(0).Value = True ' por variedad
                
                PonerFoco txtCodigo(4)
                
            Case 4 ' 4-Informe de destino de notas de campo
                PonerFoco txtCodigo(8)
                Option2(0).Value = True
                
                
            Case 5  ' 5-Listado de Stocks
                PonerFoco txtCodigo(28)
            
            Case 6  ' 6-manejo de Cargas de Confeccion
                txtCodigo(30).Text = Format(Now, "dd/mm/yyyy")
                
                PonerFoco txtCodigo(30)
                        
            Case 7 ' cargas por linea de confecccion por fecha/producto
                txtCodigo(11).Text = Format(Now, "dd/mm/yyyy")
                txtCodigo(12).Text = Format(Now, "dd/mm/yyyy")
                
                PonerFoco txtCodigo(11)
            
        End Select
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
Dim H As Integer, W As Integer
Dim List As Collection

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
    
    For H = 0 To 5
        Me.imgBuscar(H).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next H
    For H = 9 To 10
        Me.imgBuscar(H).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next H
    For H = 28 To 29
        Me.imgBuscar(H).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next H
    
    Set List = Nothing

    'Ocultar todos los Frames de Formulario
    FrameListadoStocks.visible = False
    FrameDesviacionAforos.visible = False
    FramePaletsEntrada.visible = False
    FrameOrigenPaletConf.visible = False
    FrameDestinoNotas.visible = False
    FrameCargasFecha.visible = False
    '###Descomentar
'    CommitConexion
    
    Select Case OpcionListado
    
    'LISTADOS DE MANTENIMIENTOS BASICOS
    '---------------------
    Case 1   '1- Informe de Palets de Entrada
        FramePaletsEntradaVisible True, H, W
        Tabla = "trzpalets"
        Me.Label5.Caption = "Informe de Palets de Entrada"
    
    Case 2   '2- Informe de detalle de cargas en lineas de confeccion
        FramePaletsEntradaVisible True, H, W
        Tabla = "trzlineas_cargas"
        Me.Label5.Caption = "Detalle Cargas en Línea Confección"
    
    Case 3   '3- Informe de origen de palets confeccionados
        FrameOrigenPaletsConfeccionadosVisible True, H, W
        Tabla = "trzlineas_cargas"
    
    Case 4   '4- Informe de destinos de notas de entrada
        FrameDestinoNotasVisible True, H, W
        Tabla = "trzlineas_cargas"
    
    
    Case 5   '5- Listado de stocks
        FrameListadoStocksVisible True, H, W
        Tabla = "trzpalets"
        Me.Label5.Caption = "Informe de Palets de Entrada"
        
        '[Monica]08/05/2015: solo para el caso de catadau quieren un listado diferente
        Me.Check5.visible = (vParamAplic.Cooperativa = 0)
        Me.Check5.Enabled = (vParamAplic.Cooperativa = 0)
        
    
    Case 6   '6- Manejo de Cargas de Confeccion
        FramePaletsEntradaVisible True, H, W
        Tabla = "trzlineas_cargas"
        Me.Label5.Caption = "Manejo de Cargas de Confección"
    
        Label2(1).visible = False
        Label2(2).visible = False
        imgFec(5).visible = False
        imgFec(5).Enabled = False
        txtCodigo(31).visible = False
        txtCodigo(31).Enabled = False
    
    Case 7   '2- Informe de detalle de cargas en lineas por fecha/producto
        FrameCargasFechaVisible True, H, W
        Tabla = "trzlineas_cargas"
    
    
    End Select
    'Esto se consigue poneinedo el cancel en el opcion k corresponda
'    Me.cmdCancel(indFrame).Cancel = True
    Me.Width = W + 70
    Me.Height = H + 350
End Sub

Private Sub frmC_Selec(vFecha As Date)
    ' *** repasar si el camp es txtAux o Text1 ***
    txtCodigo(indice).Text = Format(vFecha, "dd/mm/yyyy") '<===
    ' ********************************************
End Sub

Private Sub frmCal_DatoSeleccionado(CadenaSeleccion As String)
'Form de Consulta de calidades
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "00")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub


Private Sub frmCla_DatoSeleccionado(CadenaSeleccion As String)
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000") ' codigo de clase
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

Private Sub frmProd_DatoSeleccionado(CadenaSeleccion As String)
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmSec_DatoSeleccionado(CadenaSeleccion As String)
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmSit_DatoSeleccionado(CadenaSeleccion As String)
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "00")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmSoc_DatoSeleccionado(CadenaSeleccion As String)
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmVar_DatoSeleccionado(CadenaSeleccion As String)
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub



Private Sub imgBuscar_Click(Index As Integer)
   Select Case Index
        Case 0, 1, 20, 21, 28, 29  'Clases
            AbrirFrmClase (Index)
        
        Case 9, 10, 12, 13, 16, 17, 24, 25 'SOCIOS
            AbrirFrmSocios (Index)
            
        Case 18, 19 ' Variedades
            AbrirFrmVariedad (Index)
        
        Case 2, 3 ' productos
            AbrirFrmProducto (Index)
        
        Case 4, 5 'variedades
            AbrirFrmVariedad (Index + 58)
        
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
            indice = Index + 4
        Case 2, 3
            indice = Index
        Case 4, 5
            indice = Index + 26
        Case 6, 7
            indice = Index + 5
        Case 8, 9
            indice = Index + 6
    End Select

    ' *** repasar si el camp es txtAux o Text1 ***
    If txtCodigo(indice).Text <> "" Then frmC.NovaData = txtCodigo(indice).Text
    ' ********************************************

    frmC.Show vbModal
    Set frmC = Nothing
    ' *** repasar si el camp es txtAux o Text1 ***
    PonerFoco txtCodigo(indice) '<===
    ' ********************************************

End Sub



Private Sub Option2_Click(Index As Integer)
    If Option2(0).Value Then
        Label6.Caption = "Destino de Notas de campo"
        Label4(5).Caption = "Notas de Campo"
    Else
        Label6.Caption = "Destino de Campos"
        Label4(5).Caption = "Campo"
    End If
End Sub

Private Sub Option4_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KEYdown KeyCode
End Sub

Private Sub Option4_KeyPress(Index As Integer, KeyAscii As Integer)
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
            Case 9: KEYBusqueda KeyAscii, 9 'socio desde
            Case 10: KEYBusqueda KeyAscii, 10 'socio hasta
            Case 12: KEYBusqueda KeyAscii, 12 'socio desde
            Case 16: KEYBusqueda KeyAscii, 16 'socio desde
            Case 17: KEYBusqueda KeyAscii, 17 'socio hasta
            Case 24: KEYBusqueda KeyAscii, 24 'socio desde
            Case 25: KEYBusqueda KeyAscii, 25 'socio hasta
            Case 0: KEYBusqueda KeyAscii, 0 'clase desde
            Case 1: KEYBusqueda KeyAscii, 1 'clase hasta
            Case 18: KEYBusqueda KeyAscii, 18 'variedad desde
            Case 19: KEYBusqueda KeyAscii, 19 'variedad hasta
            Case 20: KEYBusqueda KeyAscii, 20 'clase desde
            Case 21: KEYBusqueda KeyAscii, 21 'clase hasta
            Case 28: KEYBusqueda KeyAscii, 28 'clase desde
            Case 29: KEYBusqueda KeyAscii, 29 'clase hasta
            Case 4: KEYFecha KeyAscii, 0 'fecha desde
            Case 5: KEYFecha KeyAscii, 1 'fecha hasta
            Case 2: KEYBusqueda KeyAscii, 3 'area desde
            Case 3: KEYBusqueda KeyAscii, 4 'area hasta
            Case 30: KEYFecha KeyAscii, 4 'fecha desde
            Case 31: KEYFecha KeyAscii, 5 'fecha hasta
            
            Case 11: KEYFecha KeyAscii, 6 'fecha
            Case 14: KEYFecha KeyAscii, 8 'fecha
            Case 15: KEYFecha KeyAscii, 9 'fecha
            
            Case 62: KEYBusqueda KeyAscii, 4 'variedad desde
            Case 63: KEYBusqueda KeyAscii, 5 'variedad hasta
        
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
Dim b As Boolean

    'Quitar espacios en blanco por los lados
    txtCodigo(Index).Text = Trim(txtCodigo(Index).Text)
'    If txtCodigo(Index).Text = "" Then Exit Sub
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    Select Case Index
        Case 6, 7, 8
            PonerFormatoEntero txtCodigo(Index)
        
        Case 60, 61 'productos
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "productos", "nomprodu", "codprodu", "N")
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "000000")
        
        Case 9, 10, 16, 17, 24, 25    'SOCIOS
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "rsocios", "nomsocio", "codsocio", "N")
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "000000")
            
        Case 2, 3, 4, 5, 30, 31, 11, 12, 14, 15 'FECHAS
            b = True
            If txtCodigo(Index).Text <> "" Then
                b = PonerFormatoFecha(txtCodigo(Index))
            End If
            
        Case 0, 1, 28, 29 'CLASES
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "clases", "nomclase", "codclase", "N")
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "000")
            
        Case 18, 19, 62, 63 ' variedades
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "variedades", "nomvarie", "codvarie", "N")
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "000000")
        
    End Select
End Sub

Private Sub FrameDesviacionAforosVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Frame para el listado de diferencias de produccion
    Me.FrameDesviacionAforos.visible = visible
    If visible = True Then
        Me.FrameDesviacionAforos.Top = -90
        Me.FrameDesviacionAforos.Left = 0
        Me.FrameDesviacionAforos.Height = 5220
        Me.FrameDesviacionAforos.Width = 6285
        W = Me.FrameDesviacionAforos.Width
        H = Me.FrameDesviacionAforos.Height
    End If
End Sub

Private Sub FramePaletsEntradaVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Frame para el listado de diferencias de produccion
    Me.FramePaletsEntrada.visible = visible
    If visible = True Then
        Me.FramePaletsEntrada.Top = -90
        Me.FramePaletsEntrada.Left = 0
        Me.FramePaletsEntrada.Height = 3870
        Me.FramePaletsEntrada.Width = 5685
        W = Me.FramePaletsEntrada.Width
        H = Me.FramePaletsEntrada.Height
    End If
End Sub


Private Sub FrameCargasFechaVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Frame para el listado de diferencias de produccion
    Me.FrameCargasFecha.visible = visible
    If visible = True Then
        Me.FrameCargasFecha.Top = -90
        Me.FrameCargasFecha.Left = 0
        Me.FrameCargasFecha.Height = 4980
        Me.FrameCargasFecha.Width = 6735
        W = Me.FrameCargasFecha.Width
        H = Me.FrameCargasFecha.Height
    End If
End Sub


Private Sub FrameOrigenPaletsConfeccionadosVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Frame para el listado de diferencias de produccion
    Me.FrameOrigenPaletConf.visible = visible
    If visible = True Then
        Me.FrameOrigenPaletConf.Top = -90
        Me.FrameOrigenPaletConf.Left = 0
        Me.FrameOrigenPaletConf.Height = 4170
        Me.FrameOrigenPaletConf.Width = 6645
        W = Me.FrameOrigenPaletConf.Width
        H = Me.FrameOrigenPaletConf.Height
    End If
End Sub

Private Sub FrameDestinoNotasVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Frame para el listado de diferencias de produccion
    Me.FrameDestinoNotas.visible = visible
    If visible = True Then
        Me.FrameDestinoNotas.Top = -90
        Me.FrameDestinoNotas.Left = 0
        Me.FrameDestinoNotas.Height = 4680 '3030
        Me.FrameDestinoNotas.Width = 6645
        W = Me.FrameDestinoNotas.Width
        H = Me.FrameDestinoNotas.Height
    End If
End Sub



Private Sub FrameListadoStocksVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Frame para el listado de stocks
    Me.FrameListadoStocks.visible = visible
    If visible = True Then
        Me.FrameListadoStocks.Top = -90
        Me.FrameListadoStocks.Left = 0
        Me.FrameListadoStocks.Height = 4170
        Me.FrameListadoStocks.Width = 6645
        W = Me.FrameListadoStocks.Width
        H = Me.FrameListadoStocks.Height
    End If
End Sub


Private Sub InicializarVbles()
    cadFormula = ""
    cadSelect = ""
    cadParam = ""
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
            cadParam = cadParam & AnyadirParametroDH(param, codD, codH, nomD, nomH)
            numParam = numParam + 1
        End If
        PonerDesdeHasta = True
    End If
End Function

Private Sub LlamarImprimir()
    With frmImprimir
        .FormulaSeleccion = cadFormula
        .OtrosParametros = cadParam
        .NumeroParametros = numParam
        .SoloImprimir = False
        .EnvioEMail = False
        .Titulo = cadTitulo
        .NombreRPT = cadNombreRPT
        .Opcion = OpcionListado
        .ConSubInforme = ConSubInforme
        .Show vbModal
    End With
End Sub

Private Sub AbrirFrmCalidad(indice As Integer)
    indCodigo = indice
    Set frmCal = New frmManCalidades
    frmCal.DatosADevolverBusqueda = "2|3|"
'    frmCli.DeConsulta = True
'    frmCli.CodigoActual = txtCodigo(indCodigo)
    frmCal.Show vbModal
    Set frmCal = Nothing
End Sub

Private Sub AbrirFrmArea(indice As Integer)
    indCodigo = indice
    Set frmArea = New frmTrzAreas
    frmArea.DatosADevolverBusqueda = "0|1|"
    frmArea.Show vbModal
    Set frmArea = Nothing
End Sub

Private Sub AbrirFrmSocios(indice As Integer)
    indCodigo = indice
    Set frmSoc = New frmManSocios
    frmSoc.DatosADevolverBusqueda = "0|1|"
'    frmCli.DeConsulta = True
'    frmCli.CodigoActual = txtCodigo(indCodigo)
    frmSoc.Show vbModal
    Set frmSoc = Nothing
End Sub

Private Sub AbrirFrmSituacion(indice As Integer)
    indCodigo = indice
    Set frmSit = New frmManSituCamp
    frmSit.DatosADevolverBusqueda = "0|1|"
    frmSit.Show vbModal
    Set frmSit = Nothing
End Sub

Private Sub AbrirFrmSocio(indice As Integer)
    indCodigo = indice
    Set frmSoc = New frmManSocios
    frmSoc.DatosADevolverBusqueda = "0|1|"
    frmSoc.Show vbModal
    Set frmSoc = Nothing
End Sub

Private Sub AbrirFrmClase(indice As Integer)
    indCodigo = indice
    Set frmCla = New frmComercial
    
    AyudaClasesCom frmCla, txtCodigo(indice).Text
        
    Set frmCla = Nothing
End Sub

Private Sub AbrirFrmVariedad(indice As Integer)
    indCodigo = indice
    Set frmVar = New frmComVar
    frmVar.DatosADevolverBusqueda = "0|1|"
    frmVar.Show vbModal
    Set frmVar = Nothing
End Sub

Private Sub AbrirFrmProducto(indice As Integer)
    
    indCodigo = indice + 58
    Set frmProd = New frmComercial
    
    AyudaProductosCom frmProd, txtCodigo(indCodigo).Text
    
    Set frmProd = Nothing
    
End Sub





Private Sub AbrirVisReport()
    Screen.MousePointer = vbHourglass
    CadenaDesdeOtroForm = ""
    With frmVisReport
        .FormulaSeleccion = cadFormula
'        .SoloImprimir = (Me.OptVisualizar(indFrame).Value = 1)
        .OtrosParametros = cadParam
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
        .Opcion = OpcionListado
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


Private Function DatosOk() As Boolean
Dim b As Boolean
Dim SQL As String
Dim Sql2 As String
Dim vClien As cSocio
' añadido
Dim Mens As String
Dim numfactu As String
Dim numser As String
Dim Fecha As Date
Dim vCont As CTiposMov
Dim tipoMov As String

Dim RS As ADODB.Recordset

    b = True
    
    Select Case OpcionListado
        Case 4
            ' listado de destino de notas de entrada
            If b And txtCodigo(8).Text = "" Then
                If Option2(0).Value Then
                    MsgBox "Debe introducir un número de nota.", vbExclamation
                Else
                    MsgBox "Debe introducir un número de campo.", vbExclamation
                End If
                PonerFoco txtCodigo(8)
                b = False
            End If
        
    End Select
    DatosOk = b

End Function


Private Function ConcatenarCampos(cTabla As String, cWhere As String) As String
Dim RS As ADODB.Recordset
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
    Set RS = New ADODB.Recordset
    
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    Sql1 = ""
    While Not RS.EOF
        Sql1 = Sql1 & DBLet(RS.Fields(0).Value, "N") & ","
        RS.MoveNext
    Wend
    Set RS = Nothing
    'quitamos el ultimo or
    ConcatenarCampos = Mid(Sql1, 1, Len(Sql1) - 1)
    
End Function

Private Function CargarTemporal(codpalet As String, codEnvio As String) As Boolean
' codpalet = palets.numpalet
' codenvio = palets.numpedid
Dim RS As ADODB.Recordset
Dim SQL As String
Dim Sql1 As String
Dim Sql2 As String
Dim Rs2 As Recordset
Dim DFecHoraPalet As Date
Dim HFecHoraPalet As Date

Dim Cad As String
Dim HayReg As Boolean

    On Error GoTo eCargarTemporal
    
    CargarTemporal = False

    '-- Primero borramos la información de la temporal
    Sql2 = "delete from trztmp_palets_lineas_cargas where codusu = " & vUsu.Codigo
    conn.Execute Sql2

    If codEnvio = "" Then
        If codpalet = "" Then
'12/06/2009
'            SQL = "select * from palets where fechaini >= " & DBSet(txtcodigo(4).Text, "F") & _
'                            " and fechaini <= " & DBSet(txtcodigo(5).Text, "F")
'12/06/2009: cambiado por la fecha de confeccion

'14/12/2009
'            SQL = "select * from palets where fechaconf >= " & DBSet(txtcodigo(4).Text, "F") & _
'                            " and fechaconf <= " & DBSet(txtcodigo(5).Text, "F")
'14/12/2009: cambiado pq ahora enlazamos por la variedad del palet
'            Sql = "select distinct palets.*, palets_variedad.codvarie from palets, palets_variedad where fechaconf >= " & DBSet(txtcodigo(4).Text, "F") & _
'                            " and fechaconf <= " & DBSet(txtcodigo(5).Text, "F") & _
'                            " and palets.numpalet = palets_variedad.numpalet "
'24/05/2010: ahora puede ser por variedad o por linea
            If Me.Option1(0).Value Then ' si por variedad
                SQL = "select distinct palets.*, palets_variedad.codvarie from palets, palets_variedad where fechaconf >= " & DBSet(txtCodigo(4).Text, "F") & _
                                " and fechaconf <= " & DBSet(txtCodigo(5).Text, "F") & _
                                " and palets.numpalet = palets_variedad.numpalet "
            Else ' si por linea
                SQL = "select * from palets where fechaconf >= " & DBSet(txtCodigo(4).Text, "F") & _
                                " and fechaconf <= " & DBSet(txtCodigo(5).Text, "F")
            
            End If

        Else
'14/12/2009
'            SQL = "select * from palets where numpalet = " & DBSet(CStr(codpalet), "N")
'14/12/2009: cambiado pq ahora enlazamos por la variedad del palet
'            Sql = "select distinct palets.*, palets_variedad.codvarie from palets, palets_variedad where palets.numpalet = " & DBSet(CStr(codpalet), "N") & _
'                            " and palets.numpalet = palets_variedad.numpalet "
'24/05/2010: ahora puede ser por variedad o por linea
            If Me.Option1(0).Value Then ' si por variedad
                SQL = "select distinct palets.*, palets_variedad.codvarie from palets, palets_variedad where palets.numpalet = " & DBSet(CStr(codpalet), "N") & _
                                " and palets.numpalet = palets_variedad.numpalet "
            Else
                SQL = "select * from palets where numpalet = " & DBSet(CStr(codpalet), "N")
            End If
        
        End If
    Else
'14/12/2009
'        SQL = "select * from palets where numpedid = " & DBSet(CStr(codEnvio), "N")
'14/12/2009: cambiado pq ahora enlazamos por la variedad del palet
'         Sql = "select distinct palets.*, palets_variedad.codvarie from palets, palets_variedad where palets.numpedid = " & DBSet(CStr(codEnvio), "N") & _
'                            " and palets.numpalet = palets_variedad.numpalet "
'24/05/2010: ahora puede ser por variedad o por linea
        If Me.Option1(0).Value Then ' si por variedad
            SQL = "select distinct palets.*, palets_variedad.codvarie from palets, palets_variedad where palets.numpedid = " & DBSet(CStr(codEnvio), "N") & _
                               " and palets.numpalet = palets_variedad.numpalet "
        Else
            SQL = "select * from palets where numpedid = " & DBSet(CStr(codEnvio), "N")
        End If
    End If
    
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not RS.EOF Then
        RS.MoveFirst
        While Not RS.EOF
            '-- 10 minutos antes de empezar y diez minutos antes de parar
'12/06/2009
'            DFecHoraPalet = DateAdd("n", -10, CDate(Format(Rs!FechaIni, "dd/mm/yyyy") & Format(Rs!horaini, " hh:mm:ss")))
'            HFecHoraPalet = DateAdd("n", -10, CDate(Format(Rs!FechaFin, "dd/mm/yyyy") & Format(Rs!HoraFin, " hh:mm:ss")))
'12/06/2009: cambiado por la fecha de confeccion
            DFecHoraPalet = DateAdd("n", -10, RS!horaiconf)
            HFecHoraPalet = DateAdd("n", -10, RS!horafconf)
            
            '-- Buscamos las cargas en ese periodo
'14/12/2009
'                    "where linea = " & CStr(Rs!linconfe)
'14/12/2009: cambiado por 1=1

'15/02/2010: cambiamos el from ahora enlazamos con trzpalets y le pasamos el codvarie
'            Sql = "select * from trzlineas_cargas, trzpalets " & _
'                    "where 1=1 " & _
'                    " and trzlineas_cargas.idpalet = trzpalets.idpalet " & _
'                    " and trzpalets.codvarie = " & DBSet(Rs!codvarie, "N") & _
'                    " and fechahora >= " & DBSet(DFecHoraPalet, "FH") & _
'                    " and fechahora <= " & DBSet(HFecHoraPalet, "FH")
'            Set Rs2 = New ADODB.Recordset
'            Rs2.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'24/05/2010:  ahora puede ser por variedad o por linea
            If Option1(0).Value Then
                SQL = "select * from trzlineas_cargas, trzpalets " & _
                        "where 1=1 " & _
                        " and trzlineas_cargas.idpalet = trzpalets.idpalet " & _
                        " and trzpalets.codvarie = " & DBSet(RS!codvarie, "N") & _
                        " and fechahora >= " & DBSet(DFecHoraPalet, "FH") & _
                        " and fechahora <= " & DBSet(HFecHoraPalet, "FH")
                Set Rs2 = New ADODB.Recordset
                Rs2.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            Else
                SQL = "select * from trzlineas_cargas " & _
                    "where linea = " & CStr(RS!linconfe) & _
                        " and fechahora >= " & DBSet(DFecHoraPalet, "FH") & _
                        " and fechahora <= " & DBSet(HFecHoraPalet, "FH")

                Set Rs2 = New ADODB.Recordset
                Rs2.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            End If
            
            If Not Rs2.EOF Then
                Rs2.MoveFirst
                While Not Rs2.EOF
                    SQL = "insert into trztmp_palets_lineas_cargas (codusu, numpalet, linea, palet, fechahora, fecha)"
                    SQL = SQL & " values("
                    SQL = SQL & DBSet(vUsu.Codigo, "N") & ","
                    SQL = SQL & CStr(RS!NumPalet) & ","
'14/12/2009
'                    SQL = SQL & CStr(Rs2!linea) & ","
'14/12/2009: no insertamos en la temporal la linea sino la variedad
'                    Sql = Sql & CStr(Rs!codvarie) & ","
'24/05/2010:  ahora puede ser por variedad o por linea
                    If Option1(0).Value Then ' si es por variedad
                        SQL = SQL & CStr(RS!codvarie) & ","
                    Else
                        SQL = SQL & CStr(Rs2!Linea) & ","
                    End If
                    
                    SQL = SQL & CStr(Rs2!IdPalet) & ","
                    SQL = SQL & DBSet(Rs2!FechaHora, "FH") & ","
                    SQL = SQL & DBSet(Rs2!Fecha, "F") & ")"
                    conn.Execute SQL
                    Rs2.MoveNext
                Wend
            End If
            RS.MoveNext
        Wend
    Else
        MsgBox "No se han encontrado palets confeccionados"
        CargarTemporal = False
        Exit Function
    End If
    
    CargarTemporal = True
    Exit Function
    
eCargarTemporal:
    CargarTemporal = False
    MuestraError "Cargando temporal Origen Palets Confeccionados", Err.Description
End Function


Private Function CargarTemporalDestinos() As Boolean
'-- Carga la base de datos temporal con la información que toca.
Dim DFecHoraPalet As Date
Dim HFecHoraPalet As Date
Dim FecHoraCarga As Date
Dim SQL As String
Dim RS As ADODB.Recordset
Dim Rs2 As ADODB.Recordset
Dim HoraPalet As String
Dim HoraInicio As String
Dim HoraFin As String
Dim NumNota As String

Dim Variedad As String

    
    CargarTemporalDestinos = False
    
    '-- Primero borramos lo que hubiera.
    SQL = "delete from trztmp_palets_lineas_cargas where codusu = " & DBSet(vUsu.Codigo, "N")
    conn.Execute SQL
    '-- Buscamos palets abocados con ese código de referencia
    NumNota = txtCodigo(8).Text
    SQL = "select * from trzlineas_cargas where idpalet in "
    
    If Option2(0).Value Then
        SQL = SQL & "(select IdPalet from trzpalets where numnotac = " & DBSet(txtCodigo(8).Text, "N") & ")" '& _
'               " or idpalet in (select a.IdPalet from trzpalet_palets as a, trzpalets as b" & _
'                " where b.numnotac = '5234252' and b.IdPalet = a.IdPalet2 )"

        '[Monica]04/06/2014: guardamos las variedad
        Variedad = DevuelveValor("select distinct codvarie from trzpalets where numnotac = " & DBSet(txtCodigo(8).Text, "N"))

    '[Monica]05/02/2014: nuevo listado de destinos por campo
    Else
        SQL = SQL & "(select IdPalet from trzpalets where codcampo = " & DBSet(txtCodigo(8).Text, "N") & ")"
        
        '[Monica]04/06/2014: guardamos las variedad
        Variedad = DevuelveValor("select distinct codvarie from trzpalets where numnotac = " & DBSet(txtCodigo(8).Text, "N"))
        
    End If
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not RS.EOF Then
        RS.MoveFirst
        While Not RS.EOF
            '-- 10 minutos antes de empezar y diez minutos antes de parar
            FecHoraCarga = DateAdd("n", 10, RS!FechaHora)
            HoraPalet = Format(FecHoraCarga, "hh:mm:ss")
            '-- Cogemos todos los palets confeccionados en la fecha porque la
            '   selección por horas no funciona
            SQL = "select * from palets where" & _
                        " fechaini = " & DBSet(FecHoraCarga, "F") & _
                        " and linconfe = " & CStr(RS!Linea)
            '[Monica]12/02/2014: introducimos el desde/hasta fecha de inicio de palet confeccionado
            If txtCodigo(14).Text <> "" Then SQL = SQL & " and fechaini >= " & DBSet(txtCodigo(14).Text, "F")
            If txtCodigo(15).Text <> "" Then SQL = SQL & " and fechaini <= " & DBSet(txtCodigo(15).Text, "F")
            
            '04/06/2014: miramos que sea la misma variedad
            If vParamAplic.Cooperativa = 12 Then
                SQL = SQL & " and numpalet in (select numpalet from palets_variedad where codvarie = " & DBSet(Variedad, "N") & ")"
            End If
                        
                        
            Set Rs2 = New ADODB.Recordset
            Rs2.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            If Not Rs2.EOF Then
                Rs2.MoveFirst
                While Not Rs2.EOF
                    HoraInicio = Format(Rs2!HoraIni, "hh:mm:ss")
                    HoraFin = Format(Rs2!HoraFin, "hh:mm:ss")
                    If (HoraInicio <= HoraPalet) And (HoraFin >= HoraPalet) And (Not YaEstaPalet(Rs2!NumPalet, RS!IdPalet)) Then
                        '-- este es un posible palet de confección
                        SQL = "insert into trztmp_palets_lineas_cargas (codusu, numpalet, linea, palet, codtipo, fechahora, fecha, numnotac)"
                        SQL = SQL & " values("
                        SQL = SQL & DBSet(vUsu.Codigo, "N") & ","
                        SQL = SQL & CStr(Rs2!NumPalet) & ","
                        SQL = SQL & CStr(RS!Linea) & ","
                        SQL = SQL & CStr(RS!IdPalet) & ","
                        SQL = SQL & CStr(RS!tipo) & ","
                        SQL = SQL & DBSet(RS!FechaHora, "FH") & ","
                        SQL = SQL & DBSet(RS!Fecha, "F") & ","
                        SQL = SQL & DBSet(NumNota, "N") & ")"
                        conn.Execute SQL
                    End If
                    Rs2.MoveNext
                Wend
                CargarTemporalDestinos = True
            Else
'[Monica]12/02/2014: al meter el desde/hasta fecha ya no tiene sentido que le digamos que no tiene referencia en los confeccionados
'                MsgBox "El palet abocado " & CStr(RS!IdPalet) & " no tiene referencia en los confeccionados" & vbCrLf & _
'                    "Seguramente el número de linea no fue bien introducida en el confeccionado"
            End If
            Set Rs2 = Nothing
            RS.MoveNext
        Wend
    Else
        MsgBox "No se han encontrado palets abocados a línea de confección con esta referencia"
        CargarTemporalDestinos = False
    End If
    
    Set RS = Nothing
    
End Function

Private Function YaEstaPalet(codpalet As Long, Palet As Long) As Boolean
Dim RS As ADODB.Recordset
Dim SQL As String
    
    SQL = "select * from trztmp_palets_lineas_cargas where numpalet = " & CStr(codpalet) & _
            " and palet = " & CStr(Palet) & _
            " and codusu = " & vUsu.Codigo '[Monica]25/05/2016:faltaba esta condicion
            
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    YaEstaPalet = Not RS.EOF

    Set RS = Nothing

End Function



Private Function CargarTemporalAbocamiento() As Boolean
Dim RS As ADODB.Recordset
Dim SQL As String
Dim Sql1 As String
Dim Sql2 As String
Dim Rs2 As Recordset
Dim DFecHoraPalet As Date
Dim HFecHoraPalet As Date

Dim Cad As String
Dim HayReg As Boolean

    On Error GoTo eCargarTemporal
    
    CargarTemporalAbocamiento = False

    '-- Primero borramos la información de la temporal
    Sql2 = "delete from tmpinformes where codusu = " & vUsu.Codigo
    conn.Execute Sql2

    Sql2 = "delete from trztmp_palets_lineas_cargas where codusu = " & vUsu.Codigo
    conn.Execute Sql2

    If Me.Option1(0).Value Then ' si por variedad
        SQL = "select distinct palets.*, palets_variedad.codvarie from palets, palets_variedad where fechaconf >= " & DBSet(txtCodigo(4).Text, "F") & _
                        " and fechaconf <= " & DBSet(txtCodigo(5).Text, "F") & _
                        " and palets.numpalet = palets_variedad.numpalet "
    Else ' si por linea
        SQL = "select * from palets where fechaconf >= " & DBSet(txtCodigo(4).Text, "F") & _
                        " and fechaconf <= " & DBSet(txtCodigo(5).Text, "F")
    
    End If
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not RS.EOF Then
        RS.MoveFirst
        While Not RS.EOF
            '-- 10 minutos antes de empezar y diez minutos antes de parar
            DFecHoraPalet = DateAdd("n", -10, RS!horaiconf)
            HFecHoraPalet = DateAdd("n", -10, RS!horafconf)
            
            '-- Buscamos las cargas en ese periodo
            If Option1(0).Value Then
                SQL = "select * from trzlineas_cargas, trzpalets " & _
                        "where 1=1 " & _
                        " and trzlineas_cargas.idpalet = trzpalets.idpalet " & _
                        " and trzpalets.codvarie = " & DBSet(RS!codvarie, "N") & _
                        " and fechahora >= " & DBSet(DFecHoraPalet, "FH") & _
                        " and fechahora <= " & DBSet(HFecHoraPalet, "FH")
                Set Rs2 = New ADODB.Recordset
                Rs2.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            Else
                SQL = "select * from trzlineas_cargas " & _
                    "where linea = " & CStr(RS!linconfe) & _
                        " and fechahora >= " & DBSet(DFecHoraPalet, "FH") & _
                        " and fechahora <= " & DBSet(HFecHoraPalet, "FH")

                Set Rs2 = New ADODB.Recordset
                Rs2.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            End If
            
            If Not Rs2.EOF Then
                Rs2.MoveFirst
                While Not Rs2.EOF
                    SQL = "insert into trztmp_palets_lineas_cargas (codusu, numpalet, linea, palet, fechahora, fecha)"
                    SQL = SQL & " values("
                    SQL = SQL & DBSet(vUsu.Codigo, "N") & ","
                    SQL = SQL & CStr(RS!NumPalet) & ","
                    
                    If Option1(0).Value Then ' si es por variedad
                        SQL = SQL & CStr(RS!codvarie) & ","
                    Else
                        SQL = SQL & CStr(Rs2!Linea) & ","
                    End If
                    
                    SQL = SQL & CStr(Rs2!IdPalet) & ","
                    SQL = SQL & DBSet(Rs2!FechaHora, "FH") & ","
                    SQL = SQL & DBSet(Rs2!Fecha, "F") & ")"
                    conn.Execute SQL
                    
                    Rs2.MoveNext
                Wend
            End If
            RS.MoveNext
        Wend
    End If

'    SQL = "insert into tmpinformes (codusu, codigo1, fecha1, importe1) "
'    SQL = SQL & "select " & vUsu.Codigo & ", trzpalets.codvarie, date(trzlineas_cargas.fechahora), sum(trzpalets.numkilos) from trzlineas_cargas, trzpalets "
'    SQL = SQL & " where date(trzlineas_cargas.fechahora) between " & DBSet(txtCodigo(4).Text, "F") & " and " & DBSet(txtCodigo(5).Text, "F")
'    SQL = SQL & " group by 1, 2"
'    SQL = SQL & " order by 1, 2"
'
'    conn.Execute SQL

    CargarTemporalAbocamiento = True
    Exit Function
    
eCargarTemporal:
    CargarTemporalAbocamiento = False
    MuestraError "Cargando temporal Abocamiento Palets Confeccionados", Err.Description
End Function


