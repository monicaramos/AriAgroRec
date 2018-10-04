VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmListadoComer 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informes"
   ClientHeight    =   8010
   ClientLeft      =   45
   ClientTop       =   2430
   ClientWidth     =   8745
   Icon            =   "frmListadoComer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8010
   ScaleWidth      =   8745
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cd1 
      Left            =   7680
      Top             =   5160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame FrameMovimientoEnvases 
      Height          =   7545
      Left            =   0
      TabIndex        =   76
      Top             =   0
      Width           =   7155
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
         Index           =   28
         Left            =   2700
         Locked          =   -1  'True
         TabIndex        =   145
         Text            =   "Text5"
         Top             =   4530
         Width           =   3780
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
         Index           =   28
         Left            =   1830
         MaxLength       =   6
         TabIndex        =   84
         Top             =   4530
         Width           =   875
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
         Index           =   26
         Left            =   2700
         Locked          =   -1  'True
         TabIndex        =   144
         Text            =   "Text5"
         Top             =   4140
         Width           =   3780
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
         Index           =   26
         Left            =   1830
         MaxLength       =   6
         TabIndex        =   83
         Top             =   4140
         Width           =   875
      End
      Begin VB.CheckBox Check5 
         Caption         =   "Sólo con saldo distinto de cero"
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
         Left            =   3645
         TabIndex        =   140
         Top             =   6165
         Width           =   3435
      End
      Begin VB.CheckBox Check4 
         Caption         =   "Sacar Saldo"
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
         Left            =   3345
         TabIndex        =   139
         Top             =   5865
         Width           =   2670
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Sacar Compras"
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
         Left            =   3345
         TabIndex        =   126
         Top             =   5535
         Width           =   2670
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Ordenado por Cliente"
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
         Left            =   3345
         TabIndex        =   125
         Top             =   5190
         Width           =   2670
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
         Index           =   23
         Left            =   1830
         MaxLength       =   6
         TabIndex        =   82
         Top             =   3615
         Width           =   875
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
         Index           =   22
         Left            =   1830
         MaxLength       =   6
         TabIndex        =   81
         Top             =   3210
         Width           =   875
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
         Index           =   23
         Left            =   2700
         Locked          =   -1  'True
         TabIndex        =   121
         Text            =   "Text5"
         Top             =   3615
         Width           =   3780
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
         Index           =   22
         Left            =   2700
         Locked          =   -1  'True
         TabIndex        =   120
         Text            =   "Text5"
         Top             =   3210
         Width           =   3780
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
         Index           =   21
         Left            =   1845
         MaxLength       =   16
         TabIndex        =   80
         Top             =   2655
         Width           =   1605
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
         MaxLength       =   16
         TabIndex        =   79
         Top             =   2265
         Width           =   1605
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
         Index           =   21
         Left            =   3465
         Locked          =   -1  'True
         TabIndex        =   116
         Text            =   "Text5"
         Top             =   2655
         Width           =   3015
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
         Index           =   20
         Left            =   3465
         Locked          =   -1  'True
         TabIndex        =   115
         Text            =   "Text5"
         Top             =   2265
         Width           =   3015
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
         Index           =   15
         Left            =   1845
         MaxLength       =   10
         TabIndex        =   86
         Top             =   5580
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
         Index           =   14
         Left            =   1845
         MaxLength       =   10
         TabIndex        =   85
         Top             =   5175
         Width           =   1350
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
         Index           =   13
         Left            =   2715
         Locked          =   -1  'True
         TabIndex        =   90
         Text            =   "Text5"
         Top             =   1680
         Width           =   3780
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
         Index           =   12
         Left            =   2715
         Locked          =   -1  'True
         TabIndex        =   89
         Text            =   "Text5"
         Top             =   1320
         Width           =   3780
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
         Index           =   13
         Left            =   1845
         MaxLength       =   2
         TabIndex        =   78
         Top             =   1680
         Width           =   875
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
         Index           =   12
         Left            =   1845
         MaxLength       =   2
         TabIndex        =   77
         Top             =   1320
         Width           =   875
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
         Index           =   2
         Left            =   4290
         TabIndex        =   87
         Top             =   6810
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
         Index           =   4
         Left            =   5415
         TabIndex        =   88
         Top             =   6810
         Width           =   1065
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   22
         Left            =   1470
         MouseIcon       =   "frmListadoComer.frx":000C
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar destino"
         Top             =   4575
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   21
         Left            =   1470
         MouseIcon       =   "frmListadoComer.frx":015E
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar destino"
         Top             =   4185
         Width           =   240
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Destino"
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
         Index           =   40
         Left            =   540
         TabIndex        =   143
         Top             =   3915
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
         Index           =   39
         Left            =   810
         TabIndex        =   142
         Top             =   4560
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
         Index           =   37
         Left            =   810
         TabIndex        =   141
         Top             =   4200
         Width           =   735
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
         Index           =   35
         Left            =   810
         TabIndex        =   124
         Top             =   3255
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
         Index           =   34
         Left            =   810
         TabIndex        =   123
         Top             =   3615
         Width           =   690
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Cliente"
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
         Index           =   33
         Left            =   540
         TabIndex        =   122
         Top             =   2970
         Width           =   675
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   19
         Left            =   1470
         MouseIcon       =   "frmListadoComer.frx":02B0
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar cliente"
         Top             =   3615
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   18
         Left            =   1470
         MouseIcon       =   "frmListadoComer.frx":0402
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar cliente"
         Top             =   3255
         Width           =   240
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
         Index           =   32
         Left            =   825
         TabIndex        =   119
         Top             =   2310
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
         Index           =   31
         Left            =   825
         TabIndex        =   118
         Top             =   2670
         Width           =   690
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Artículo"
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
         Left            =   555
         TabIndex        =   117
         Top             =   1980
         Width           =   750
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   17
         Left            =   1485
         MouseIcon       =   "frmListadoComer.frx":0554
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar artículo"
         Top             =   2670
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   16
         Left            =   1485
         MouseIcon       =   "frmListadoComer.frx":06A6
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar artículo"
         Top             =   2310
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   1
         Left            =   1485
         Picture         =   "frmListadoComer.frx":07F8
         Top             =   5580
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   0
         Left            =   1485
         Picture         =   "frmListadoComer.frx":0883
         Top             =   5175
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   13
         Left            =   1485
         MouseIcon       =   "frmListadoComer.frx":090E
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar tipo envase"
         Top             =   1725
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   12
         Left            =   1485
         MouseIcon       =   "frmListadoComer.frx":0A60
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar tipo envase"
         Top             =   1320
         Width           =   240
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
         Index           =   23
         Left            =   555
         TabIndex        =   97
         Top             =   4950
         Width           =   600
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
         Index           =   22
         Left            =   825
         TabIndex        =   96
         Top             =   5550
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
         Index           =   21
         Left            =   825
         TabIndex        =   95
         Top             =   5235
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "Informe de Movimientos de Envases"
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
         Left            =   585
         TabIndex        =   94
         Top             =   360
         Width           =   5430
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Envase"
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
         Index           =   20
         Left            =   555
         TabIndex        =   93
         Top             =   990
         Width           =   1200
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
         Index           =   19
         Left            =   825
         TabIndex        =   92
         Top             =   1680
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
         Index           =   18
         Left            =   825
         TabIndex        =   91
         Top             =   1320
         Width           =   735
      End
   End
   Begin VB.Frame FrameCalculoHorasProductivas 
      Height          =   3525
      Left            =   0
      TabIndex        =   127
      Top             =   0
      Width           =   5835
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
         Index           =   24
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   137
         Text            =   "Text5"
         Top             =   2190
         Width           =   3000
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
         Index           =   24
         Left            =   1620
         MaxLength       =   2
         TabIndex        =   130
         Top             =   2190
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
         Index           =   27
         Left            =   1620
         MaxLength       =   10
         TabIndex        =   128
         Top             =   1290
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
         Index           =   25
         Left            =   1620
         MaxLength       =   6
         TabIndex        =   129
         Top             =   1740
         Width           =   1035
      End
      Begin VB.CommandButton CmdAcepCalHProd 
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
         Left            =   3420
         TabIndex        =   131
         Top             =   2760
         Width           =   1065
      End
      Begin VB.CommandButton cmdCancelCalHProd 
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
         Left            =   4530
         TabIndex        =   132
         Top             =   2760
         Width           =   1065
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   20
         Left            =   1335
         MouseIcon       =   "frmListadoComer.frx":0BB2
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar almacén"
         Top             =   2220
         Width           =   240
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
         Index           =   36
         Left            =   435
         TabIndex        =   138
         Top             =   2250
         Width           =   825
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   13
         Left            =   1365
         Picture         =   "frmListadoComer.frx":0D04
         Top             =   1290
         Width           =   240
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
         Index           =   41
         Left            =   435
         TabIndex        =   135
         Top             =   1290
         Width           =   600
      End
      Begin VB.Label Label8 
         Caption         =   "Cálculo de Horas Productivas"
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
         TabIndex        =   134
         Top             =   480
         Width           =   4725
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
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
         ForeColor       =   &H00972E0B&
         Height          =   240
         Index           =   38
         Left            =   435
         TabIndex        =   133
         Top             =   1800
         Width           =   1065
      End
   End
   Begin VB.Frame FrameInfPaletsCamaras 
      Height          =   4455
      Left            =   0
      TabIndex        =   152
      Top             =   0
      Width           =   7020
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
         Index           =   7
         Left            =   5730
         TabIndex        =   169
         Top             =   3735
         Width           =   975
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
         Index           =   4
         Left            =   4650
         TabIndex        =   167
         Top             =   3720
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
         Index           =   33
         Left            =   1920
         MaxLength       =   6
         TabIndex        =   159
         Top             =   1680
         Width           =   885
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
         Index           =   32
         Left            =   1920
         MaxLength       =   6
         TabIndex        =   158
         Top             =   1275
         Width           =   885
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
         Index           =   33
         Left            =   2850
         Locked          =   -1  'True
         TabIndex        =   157
         Text            =   "Text5"
         Top             =   1680
         Width           =   3825
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
         Index           =   32
         Left            =   2850
         Locked          =   -1  'True
         TabIndex        =   156
         Text            =   "Text5"
         Top             =   1275
         Width           =   3825
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
         Index           =   31
         Left            =   1920
         MaxLength       =   10
         TabIndex        =   161
         Top             =   2760
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
         Index           =   29
         Left            =   1935
         MaxLength       =   10
         TabIndex        =   160
         Top             =   2355
         Width           =   1350
      End
      Begin VB.CommandButton Command8 
         Height          =   440
         Left            =   7860
         Picture         =   "frmListadoComer.frx":0D8F
         Style           =   1  'Graphical
         TabIndex        =   155
         Top             =   2215
         Visible         =   0   'False
         Width           =   380
      End
      Begin VB.CommandButton Command7 
         Height          =   440
         Left            =   7860
         Picture         =   "frmListadoComer.frx":1099
         Style           =   1  'Graphical
         TabIndex        =   154
         Top             =   1440
         Visible         =   0   'False
         Width           =   380
      End
      Begin VB.Frame Frame2 
         Caption         =   "Ordenar por"
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
         Height          =   975
         Left            =   495
         TabIndex        =   153
         Top             =   3195
         Width           =   2190
         Begin VB.OptionButton Opcion 
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
            Height          =   345
            Index           =   3
            Left            =   315
            TabIndex        =   165
            Top             =   540
            Width           =   1650
         End
         Begin VB.OptionButton Opcion 
            Caption         =   "Cámara"
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
            Left            =   315
            TabIndex        =   163
            Top             =   270
            Width           =   1335
         End
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   1215
         Index           =   5
         Left            =   6120
         TabIndex        =   162
         Top             =   1440
         Visible         =   0   'False
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   2143
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   6
         Left            =   1665
         Picture         =   "frmListadoComer.frx":13A3
         Top             =   2790
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   5
         Left            =   1665
         Picture         =   "frmListadoComer.frx":142E
         Top             =   2340
         Width           =   240
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
         Index           =   48
         Left            =   960
         TabIndex        =   173
         Top             =   1320
         Width           =   690
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
         Index           =   47
         Left            =   960
         TabIndex        =   172
         Top             =   1680
         Width           =   645
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Cámara"
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
         Index           =   46
         Left            =   600
         TabIndex        =   171
         Top             =   990
         Width           =   735
      End
      Begin VB.Label Label10 
         Caption         =   "Informe de Palets en Cámaras"
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
         Left            =   600
         TabIndex        =   170
         Top             =   360
         Width           =   6105
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
         Index           =   45
         Left            =   960
         TabIndex        =   168
         Top             =   2400
         Width           =   690
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
         Index           =   44
         Left            =   960
         TabIndex        =   166
         Top             =   2790
         Width           =   645
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
         Index           =   42
         Left            =   600
         TabIndex        =   164
         Top             =   2070
         Width           =   600
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   24
         Left            =   1650
         MouseIcon       =   "frmListadoComer.frx":14B9
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar cámara"
         Top             =   1680
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   23
         Left            =   1650
         MouseIcon       =   "frmListadoComer.frx":160B
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar cámara"
         Top             =   1320
         Width           =   240
      End
   End
   Begin VB.Frame FrameCalibres 
      Height          =   4455
      Left            =   90
      TabIndex        =   52
      Top             =   90
      Width           =   7020
      Begin VB.Frame FrameStockMaxMin 
         Caption         =   "Ordenar por"
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
         Height          =   975
         Left            =   495
         TabIndex        =   73
         Top             =   3195
         Width           =   2505
         Begin VB.OptionButton Opcion 
            Caption         =   "Calibre"
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
            Left            =   495
            TabIndex        =   75
            Top             =   585
            Width           =   975
         End
         Begin VB.OptionButton Opcion 
            Caption         =   "Variedad "
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   0
            Left            =   495
            TabIndex        =   74
            Top             =   225
            Width           =   1290
         End
      End
      Begin VB.CommandButton Command6 
         Height          =   440
         Left            =   7860
         Picture         =   "frmListadoComer.frx":175D
         Style           =   1  'Graphical
         TabIndex        =   64
         Top             =   1440
         Visible         =   0   'False
         Width           =   380
      End
      Begin VB.CommandButton Command5 
         Height          =   440
         Left            =   7860
         Picture         =   "frmListadoComer.frx":1A67
         Style           =   1  'Graphical
         TabIndex        =   63
         Top             =   2215
         Visible         =   0   'False
         Width           =   380
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
         Index           =   11
         Left            =   2805
         Locked          =   -1  'True
         TabIndex        =   62
         Text            =   "Text5"
         Top             =   2760
         Width           =   3555
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
         Index           =   10
         Left            =   2805
         Locked          =   -1  'True
         TabIndex        =   61
         Text            =   "Text5"
         Top             =   2400
         Width           =   3555
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
         Index           =   11
         Left            =   1935
         MaxLength       =   6
         TabIndex        =   60
         Top             =   2760
         Width           =   875
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
         Index           =   10
         Left            =   1920
         MaxLength       =   6
         TabIndex        =   59
         Top             =   2400
         Width           =   875
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
         Index           =   9
         Left            =   2805
         Locked          =   -1  'True
         TabIndex        =   58
         Text            =   "Text5"
         Top             =   1635
         Width           =   3555
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
         Index           =   8
         Left            =   2805
         Locked          =   -1  'True
         TabIndex        =   57
         Text            =   "Text5"
         Top             =   1275
         Width           =   3555
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
         Left            =   1920
         MaxLength       =   6
         TabIndex        =   56
         Top             =   1635
         Width           =   875
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
         Left            =   1920
         MaxLength       =   6
         TabIndex        =   55
         Top             =   1275
         Width           =   875
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
         Index           =   1
         Left            =   4380
         TabIndex        =   54
         Top             =   3720
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
         Index           =   3
         Left            =   5460
         TabIndex        =   53
         Top             =   3735
         Width           =   1065
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   1215
         Index           =   3
         Left            =   6120
         TabIndex        =   65
         Top             =   1440
         Visible         =   0   'False
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   2143
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   11
         Left            =   1650
         MouseIcon       =   "frmListadoComer.frx":1D71
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar artículo"
         Top             =   2790
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   10
         Left            =   1650
         MouseIcon       =   "frmListadoComer.frx":1EC3
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar articulo"
         Top             =   2400
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   9
         Left            =   1650
         MouseIcon       =   "frmListadoComer.frx":2015
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar familia"
         Top             =   1680
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   8
         Left            =   1650
         MouseIcon       =   "frmListadoComer.frx":2167
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar familia"
         Top             =   1275
         Width           =   240
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Calibre"
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
         Index           =   17
         Left            =   600
         TabIndex        =   72
         Top             =   2160
         Width           =   975
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
         Index           =   16
         Left            =   870
         TabIndex        =   71
         Top             =   2790
         Width           =   735
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
         Index           =   15
         Left            =   870
         TabIndex        =   70
         Top             =   2400
         Width           =   780
      End
      Begin VB.Label Label5 
         Caption         =   "Informe de Calibres"
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
         Left            =   600
         TabIndex        =   69
         Top             =   360
         Width           =   6735
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
         Index           =   14
         Left            =   600
         TabIndex        =   68
         Top             =   1080
         Width           =   1170
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
         Index           =   13
         Left            =   870
         TabIndex        =   67
         Top             =   1680
         Width           =   735
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
         Index           =   12
         Left            =   870
         TabIndex        =   66
         Top             =   1320
         Width           =   780
      End
   End
   Begin VB.Frame FrameCreacionPalets 
      Height          =   3525
      Left            =   0
      TabIndex        =   146
      Top             =   0
      Width           =   5835
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
         Index           =   30
         Left            =   1620
         MaxLength       =   10
         TabIndex        =   149
         Top             =   1290
         Width           =   1350
      End
      Begin VB.CommandButton CmdAcepCreacionPalet 
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
         Left            =   3150
         TabIndex        =   148
         Top             =   2805
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
         Index           =   6
         Left            =   4260
         TabIndex        =   147
         Top             =   2805
         Width           =   1065
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   4
         Left            =   1320
         Picture         =   "frmListadoComer.frx":22B9
         Top             =   1290
         Width           =   240
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
         Index           =   43
         Left            =   570
         TabIndex        =   151
         Top             =   1290
         Width           =   600
      End
      Begin VB.Label Label9 
         Caption         =   "Creación automática de Palets"
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
         TabIndex        =   150
         Top             =   480
         Width           =   4725
      End
   End
   Begin VB.Frame FrameHorasTrabajadas 
      Height          =   4455
      Left            =   45
      TabIndex        =   98
      Top             =   90
      Width           =   7425
      Begin VB.CheckBox Check3 
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
         Left            =   600
         TabIndex        =   136
         Top             =   3360
         Width           =   2850
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
         Index           =   5
         Left            =   5730
         TabIndex        =   108
         Top             =   3690
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
         Index           =   3
         Left            =   4545
         TabIndex        =   106
         Top             =   3690
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
         Left            =   1890
         MaxLength       =   6
         TabIndex        =   102
         Top             =   1665
         Width           =   840
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
         Left            =   1890
         MaxLength       =   6
         TabIndex        =   101
         Top             =   1305
         Width           =   840
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
         Left            =   2715
         Locked          =   -1  'True
         TabIndex        =   100
         Text            =   "Text5"
         Top             =   1305
         Width           =   4095
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
         Left            =   2715
         Locked          =   -1  'True
         TabIndex        =   99
         Text            =   "Text5"
         Top             =   1680
         Width           =   4095
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
         Left            =   1890
         MaxLength       =   10
         TabIndex        =   104
         Top             =   2745
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
         Index           =   16
         Left            =   1890
         MaxLength       =   10
         TabIndex        =   103
         Top             =   2340
         Width           =   1350
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   1215
         Index           =   4
         Left            =   5355
         TabIndex        =   114
         Top             =   2250
         Visible         =   0   'False
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   2143
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Width           =   2540
         EndProperty
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
         Left            =   915
         TabIndex        =   113
         Top             =   1320
         Width           =   690
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
         Left            =   915
         TabIndex        =   112
         Top             =   1680
         Width           =   645
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
         Left            =   510
         TabIndex        =   111
         Top             =   1035
         Width           =   1065
      End
      Begin VB.Label Label7 
         Caption         =   "Informe de Horas Trabajadas"
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
         TabIndex        =   110
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
         Left            =   915
         TabIndex        =   109
         Top             =   2400
         Width           =   690
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
         Left            =   915
         TabIndex        =   107
         Top             =   2715
         Width           =   645
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
         Index           =   24
         Left            =   510
         TabIndex        =   105
         Top             =   2115
         Width           =   600
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   15
         Left            =   1620
         MouseIcon       =   "frmListadoComer.frx":2344
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar trabajador"
         Top             =   1665
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   14
         Left            =   1620
         MouseIcon       =   "frmListadoComer.frx":2496
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar trabajador"
         Top             =   1305
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   3
         Left            =   1575
         Picture         =   "frmListadoComer.frx":25E8
         Top             =   2745
         Width           =   240
      End
      Begin VB.Image imgFecha 
         Height          =   240
         Index           =   2
         Left            =   1575
         Picture         =   "frmListadoComer.frx":2673
         Top             =   2340
         Width           =   240
      End
   End
   Begin VB.Frame FrameTraspasoAlbaran 
      Height          =   3465
      Left            =   0
      TabIndex        =   174
      Top             =   0
      Width           =   7380
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
         Index           =   35
         Left            =   2670
         Locked          =   -1  'True
         TabIndex        =   180
         Text            =   "Text5"
         Top             =   1185
         Width           =   3960
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
         Index           =   34
         Left            =   2670
         Locked          =   -1  'True
         TabIndex        =   179
         Text            =   "Text5"
         Top             =   1680
         Width           =   3960
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
         Index           =   35
         Left            =   1785
         MaxLength       =   6
         TabIndex        =   175
         Top             =   1185
         Width           =   840
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
         Index           =   34
         Left            =   1785
         MaxLength       =   3
         TabIndex        =   176
         Top             =   1680
         Width           =   840
      End
      Begin VB.CommandButton CmdAcepTraspaso 
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
         Left            =   4335
         TabIndex        =   177
         Top             =   2475
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
         Index           =   8
         Left            =   5505
         TabIndex        =   178
         Top             =   2475
         Width           =   1065
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   26
         Left            =   1485
         MouseIcon       =   "frmListadoComer.frx":26FE
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar cliente"
         Top             =   1215
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   25
         Left            =   1485
         MouseIcon       =   "frmListadoComer.frx":2850
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar destino"
         Top             =   1710
         Width           =   240
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Destino"
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
         Index           =   54
         Left            =   540
         TabIndex        =   183
         Top             =   1665
         Width           =   735
      End
      Begin VB.Label Label11 
         Caption         =   "Traspaso de Albarán de Venta"
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
         TabIndex        =   182
         Top             =   360
         Width           =   6735
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Cliente"
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
         Index           =   51
         Left            =   540
         TabIndex        =   181
         Top             =   1215
         Width           =   675
      End
   End
   Begin VB.Frame FrameVariedades 
      Height          =   4455
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   8595
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
         Left            =   5010
         TabIndex        =   27
         Top             =   3645
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
         Left            =   3840
         TabIndex        =   26
         Top             =   3645
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
         Index           =   7
         Left            =   1875
         MaxLength       =   4
         TabIndex        =   25
         Top             =   1680
         Width           =   840
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
         Left            =   1875
         MaxLength       =   4
         TabIndex        =   24
         Top             =   1275
         Width           =   840
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
         Index           =   7
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   23
         Text            =   "Text5"
         Top             =   1680
         Width           =   3375
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
         Index           =   6
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   22
         Text            =   "Text5"
         Top             =   1275
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
         Index           =   1
         Left            =   1875
         MaxLength       =   6
         TabIndex        =   21
         Top             =   2760
         Width           =   960
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
         Left            =   1875
         MaxLength       =   6
         TabIndex        =   20
         Top             =   2355
         Width           =   960
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
         Left            =   2850
         Locked          =   -1  'True
         TabIndex        =   19
         Text            =   "Text5"
         Top             =   2760
         Width           =   3285
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
         Left            =   2850
         Locked          =   -1  'True
         TabIndex        =   18
         Text            =   "Text5"
         Top             =   2355
         Width           =   3285
      End
      Begin VB.CommandButton Command2 
         Height          =   440
         Left            =   7860
         Picture         =   "frmListadoComer.frx":29A2
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   2215
         Visible         =   0   'False
         Width           =   380
      End
      Begin VB.CommandButton Command1 
         Height          =   440
         Left            =   7860
         Picture         =   "frmListadoComer.frx":2CAC
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   1440
         Visible         =   0   'False
         Width           =   380
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   1215
         Index           =   1
         Left            =   6360
         TabIndex        =   28
         Top             =   1440
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   2143
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Width           =   2540
         EndProperty
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
         Index           =   11
         Left            =   870
         TabIndex        =   36
         Top             =   1320
         Width           =   690
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
         Index           =   10
         Left            =   870
         TabIndex        =   35
         Top             =   1680
         Width           =   645
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Producto"
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
         Index           =   9
         Left            =   600
         TabIndex        =   34
         Top             =   990
         Width           =   930
      End
      Begin VB.Label Label1 
         Caption         =   "Informe de Variedades"
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
         Left            =   600
         TabIndex        =   33
         Top             =   360
         Width           =   6735
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
         Index           =   8
         Left            =   870
         TabIndex        =   32
         Top             =   2400
         Width           =   690
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
         Index           =   7
         Left            =   870
         TabIndex        =   31
         Top             =   2715
         Width           =   645
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
         Index           =   6
         Left            =   600
         TabIndex        =   30
         Top             =   2070
         Width           =   900
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Orden del Informe"
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
         Left            =   6360
         TabIndex        =   29
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   7
         Left            =   1560
         MouseIcon       =   "frmListadoComer.frx":2FB6
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar producto"
         Top             =   1680
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   6
         Left            =   1560
         MouseIcon       =   "frmListadoComer.frx":3108
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar producto"
         Top             =   1320
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   1
         Left            =   1560
         MouseIcon       =   "frmListadoComer.frx":325A
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar variedad"
         Top             =   2760
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   0
         Left            =   1560
         MouseIcon       =   "frmListadoComer.frx":33AC
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar variedad"
         Top             =   2400
         Width           =   240
      End
   End
   Begin VB.Frame FrameProveedores 
      Height          =   3420
      Left            =   45
      TabIndex        =   37
      Top             =   90
      Width           =   8595
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
         Left            =   5010
         TabIndex        =   45
         Top             =   2685
         Width           =   1065
      End
      Begin VB.CommandButton CmdAceptar3 
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
         Left            =   3840
         TabIndex        =   44
         Top             =   2685
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
         Left            =   1875
         MaxLength       =   6
         TabIndex        =   43
         Top             =   1680
         Width           =   875
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
         Left            =   1875
         MaxLength       =   6
         TabIndex        =   42
         Top             =   1320
         Width           =   875
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
         Index           =   3
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   41
         Text            =   "Text5"
         Top             =   1680
         Width           =   3330
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
         Index           =   2
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   40
         Text            =   "Text5"
         Top             =   1320
         Width           =   3330
      End
      Begin VB.CommandButton Command4 
         Height          =   440
         Left            =   7860
         Picture         =   "frmListadoComer.frx":34FE
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   2215
         Visible         =   0   'False
         Width           =   380
      End
      Begin VB.CommandButton Command3 
         Height          =   440
         Left            =   7860
         Picture         =   "frmListadoComer.frx":3808
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   1440
         Visible         =   0   'False
         Width           =   380
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   1215
         Index           =   2
         Left            =   6360
         TabIndex        =   46
         Top             =   1440
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   2143
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label3 
         Caption         =   "Informe de Proveedores"
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
         Left            =   600
         TabIndex        =   51
         Top             =   360
         Width           =   6735
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
         Index           =   2
         Left            =   870
         TabIndex        =   50
         Top             =   1365
         Width           =   645
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
         Left            =   870
         TabIndex        =   49
         Top             =   1680
         Width           =   600
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Proveedor"
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
         Left            =   600
         TabIndex        =   48
         Top             =   1035
         Width           =   990
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Orden del Informe"
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
         Left            =   6360
         TabIndex        =   47
         Top             =   1200
         Width           =   1770
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   3
         Left            =   1560
         MouseIcon       =   "frmListadoComer.frx":3B12
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar cliente"
         Top             =   1680
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   2
         Left            =   1560
         MouseIcon       =   "frmListadoComer.frx":3C64
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar cliente"
         Top             =   1365
         Width           =   240
      End
   End
   Begin VB.Frame FrameClientes 
      Height          =   3420
      Left            =   0
      TabIndex        =   0
      Top             =   45
      Width           =   8595
      Begin VB.CommandButton cmdSubir 
         Height          =   440
         Left            =   7860
         Picture         =   "frmListadoComer.frx":3DB6
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   1440
         Visible         =   0   'False
         Width           =   380
      End
      Begin VB.CommandButton cmdBajar 
         Height          =   440
         Left            =   7860
         Picture         =   "frmListadoComer.frx":40C0
         Style           =   1  'Graphical
         TabIndex        =   7
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
         Index           =   5
         Left            =   2805
         Locked          =   -1  'True
         TabIndex        =   6
         Text            =   "Text5"
         Top             =   1770
         Width           =   3285
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
         Index           =   4
         Left            =   2805
         Locked          =   -1  'True
         TabIndex        =   5
         Text            =   "Text5"
         Top             =   1365
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
         Index           =   5
         Left            =   1920
         MaxLength       =   6
         TabIndex        =   2
         Top             =   1770
         Width           =   875
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
         Left            =   1920
         MaxLength       =   6
         TabIndex        =   1
         Top             =   1365
         Width           =   875
      End
      Begin VB.CommandButton cmdAceptar2 
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
         Left            =   3825
         TabIndex        =   3
         Top             =   2700
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
         Index           =   2
         Left            =   5010
         TabIndex        =   4
         Top             =   2700
         Width           =   1065
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   1215
         Index           =   0
         Left            =   6360
         TabIndex        =   9
         Top             =   1440
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   2143
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   5
         Left            =   1560
         MouseIcon       =   "frmListadoComer.frx":43CA
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar cliente"
         Top             =   1815
         Width           =   240
      End
      Begin VB.Image imgBuscar 
         Height          =   240
         Index           =   4
         Left            =   1560
         MouseIcon       =   "frmListadoComer.frx":451C
         MousePointer    =   4  'Icon
         ToolTipText     =   "Buscar cliente"
         Top             =   1365
         Width           =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Orden del Informe"
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
         Index           =   46
         Left            =   6360
         TabIndex        =   14
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Cliente"
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
         Left            =   600
         TabIndex        =   13
         Top             =   1035
         Width           =   720
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
         Index           =   5
         Left            =   870
         TabIndex        =   12
         Top             =   1770
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
         Index           =   4
         Left            =   870
         TabIndex        =   11
         Top             =   1365
         Width           =   735
      End
      Begin VB.Label lbltitulo2 
         Caption         =   "Informe de Clientes"
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
         Left            =   600
         TabIndex        =   10
         Top             =   360
         Width           =   6735
      End
   End
End
Attribute VB_Name = "frmListadoComer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' +-+-+-+-+-+-+-+-+-+-+-
' +-+- Autor: LAURA +-+-
' +-+-+-+-+-+-+-+-+-+-+-

Option Explicit

Public Opcionlistado As Byte
    '==== Listados BASICOS ====
    '=============================
    ' 10 .- Listado de Clientes
    ' 11 .- Listado de Proveedores
    ' 12 .- Listado de Variedades
    ' 13 .- Listado de Calibres
    ' 15 .- Listado de Horas trababajadas
    ' 16 .- Calculo de Horas productivas
    
    
    ' 1 .- Creacion automatica de palets
    ' 2 .- Informe de palets en camaras
    
    ' 3 .- Traspaso de albaran a bddestino
    
    
    
Public NumCod As String 'Para indicar cod. Traspaso,Movimiento, etc. que llama
                        'Para indicar nº oferta a imprimir
    
Private Conexion As Byte
'1.- Conexión a BD Ariges  2.- Conexión a BD Conta

Private HaDevueltoDatos As Boolean


'Private WithEvents frmPro As frmManProve 'Proveedores
'Private WithEvents frmCli As frmClientes 'Clientes
Private WithEvents frmCli2 As frmBasico 'Basico
Attribute frmCli2.VB_VarHelpID = -1
'Private WithEvents frmProd As frmManProductos 'Productos
Private WithEvents frmVar As frmManVariedad 'Variedades
Attribute frmVar.VB_VarHelpID = -1
'Private WithEvents frmCal As frmManCalibres 'Calibres
'Private WithEvents frmTArt As frmManTipArtic 'tipos de articulos
'Private WithEvents frmTra As frmManTraba 'mantenimiento de trabajadores
'Private WithEvents frmArt As frmManArtic 'mantenimiento de articulos
'Private WithEvents frmAlm As frmManAlmProp 'mantenimiento de almacenes propios
'Private WithEvents frmDes As frmDestCli 'Destinos de Clientes
Private WithEvents frmDes2 As frmBasico 'Destinos de Clientes
Attribute frmDes2.VB_VarHelpID = -1
'Private WithEvents frmMensDestino As frmMensajes 'mensajes
'Private WithEvents frmCam As frmManCamara 'mantenimiento de camaras

Private WithEvents frmB As frmBuscaGrid  'Busquedas
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmC As frmCal 'calendario fechas
Attribute frmC.VB_VarHelpID = -1

'GENERALES PARA PASARLE A CRYSTAL REPORT
Private cadFormula As String 'Cadena con la FormulaSelection para Crystal Report
Private cadParam As String 'Cadena con los parametros para Crystal Report
Private numParam As Byte 'Numero de parametros que se pasan a Crystal Report
Private cadselect As String 'Cadena para comprobar si hay datos antes de abrir Informe
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
Dim ConSubInforme As Boolean

Dim albaranes As String


Private Sub KEYpress(KeyAscii As Integer)
Dim cerrar As Boolean

    KEYpressGnral KeyAscii, 0, cerrar
    If cerrar Then Unload Me
End Sub

Private Sub Check2_Click()
    Check4.Enabled = (Check2.Value = 0)
    Check5.Enabled = (Check2.Value = 0)
End Sub


Private Sub CmdAcepCalHProd_Click()
Dim Sql As String

    If txtCodigo(27).Text = "" Then
        MsgBox "Debe introducir obligatoriamente un valor en el campo Fecha para realizar el cálculo.", vbExclamation
        Exit Sub
    End If
    
    If txtCodigo(25).Text = "" Then
        MsgBox "Debe introducir un porcentaje para realizar el cálculo.", vbExclamation
        Exit Sub
    End If

    If txtCodigo(24).Text = "" Then
        MsgBox "Debe introducir el almacén para realizar el cálculo.", vbExclamation
        Exit Sub
    End If
    
    Sql = "select * from horas where fechahora = " & DBSet(txtCodigo(27).Text, "F")
    Sql = Sql & " and codalmac = " & DBSet(txtCodigo(24), "N")
    Sql = Sql & " and codtraba in (select codtraba from straba where codsecci = 1)"

    If TotalRegistros(Sql) = 0 Then
        MsgBox "No existen registros para esa fecha en el almacén introducido. Revise.", vbExclamation
        PonerFoco txtCodigo(27)
    Else
        If CalculoHorasProductivas Then
            MsgBox "Proceso realizado correctamente.", vbExclamation
           
            cmdCancelCalHProd_Click
        End If
    End If
End Sub

Private Sub CmdAcepCreacionPalet_Click()
Dim Sql As String


    Sql = "select * from trzlineas_cargas where fecha = " & DBSet(txtCodigo(30).Text, "F")
    Sql = Sql & " and idpalet not in (select idpalet from palets) "
    
    If TotalRegistros(Sql) = 0 Then
        MsgBox "No se ha realizado ningún volcado esa fecha.", vbExclamation
    Else
        If ProcesoCarga(Sql) Then
            MsgBox "Proceso realizado correctamente.", vbExclamation
        End If
    End If


End Sub

Private Function ProcesoCarga(vSQL As String) As Boolean
Dim vMens As String

    On Error GoTo eProcesoCarga
    
    conn.BeginTrans
    
    vMens = ""
    If CargarPaletsConfeccionados(vSQL, vMens) Then
        If RepartoAlbaranes(vMens) Then
            conn.CommitTrans
            Exit Function
        End If
    End If
    
eProcesoCarga:
    conn.RollbackTrans
    MuestraError Err.Number, vMens
End Function

Private Function RepartoAlbaranes(vMens As String) As Boolean
Dim Sql As String
Dim Sql2 As String
Dim Rs As ADODB.Recordset
Dim Rs2 As ADODB.Recordset
Dim Salir As Boolean
Dim KilosVar As Long
Dim NumLinea As Integer
Dim resto As Long
Dim vcodigo As Long

    On Error GoTo eRepartoAlbaranes

    RepartoAlbaranes = False

    ' para todos los albaranes que han salido repartimos
    Sql = "select albaran.numalbar, codvarie, sum(numcajas), sum(pesoneto) pesoneto from albaranes_variedad inner join albaran on albaran_variedad.numalbar = albaran.numalbar "
    Sql = Sql & " where albaran.fecalbar = " & DBSet(txtCodigo(30).Text, "F")
    Sql = Sql & " group by 1,2 "
    Sql = Sql & " order by 1,2 "
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not Rs.EOF
        Sql2 = "select sum(kilos) from trzmovim where numalbar = 0 and codvarie = " & DBSet(Rs!Codvarie, "N")
        
        KilosVar = DBLet(Rs!PesoNeto)
        If DevuelveValor(Sql2) < DBLet(Rs!PesoNeto) Then
            MsgBox "No hay suficiente existencias de la variedad " & DBLet(Rs!Codvarie), vbExclamation
            Exit Function
        Else
            Sql2 = "select * from trzmovim where numalbar = 0 and codvarie = " & DBSet(Rs!Codvarie, "N")
            Sql2 = Sql2 & " order by fecha desc "
            
            Set Rs2 = New ADODB.Recordset
            Rs2.Open Sql2, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            
            Salir = False
            
            NumLinea = DevuelveValor("select coalesce(numlinea, 0) + 1 from albaran_palets where numalbar = " & DBSet(Rs!NumAlbar, "N"))
            
            While Not Rs2.EOF And Not Salir
                Sql = "insert into albaran_palets (numalbar, numlinea, numpalet) values ("
                Sql = Sql & DBSet(Rs!NumAlbar, "N") & "," & DBSet(NumLinea, "N") & "," & DBSet(Rs2!NumPalet, "N") & ")"
                
                conn.Execute Sql
            
                If DBLet(Rs2!Kilos) <= KilosVar Then
                    
                    KilosVar = KilosVar - DBLet(Rs2!Kilos)
                    
                    Sql = "update trzmovim set numalbar = " & DBSet(Rs!NumAlbar, "N")
                    Sql = Sql & " where codigo = " & DBSet(Rs2!Codigo, "N")
                    
                    conn.Execute Sql
                Else
                    resto = DBLet(Rs2!Kilos) - KilosVar
                
                    Sql = "update trzmovim set numalbar = " & DBSet(Rs!NumAlbar, "N")
                    Sql = Sql & ", kilos =  " & DBSet(Rs!Kilos, "N")
                    Sql = Sql & " where codigo = " & DBSet(Rs2!Codigo, "N")
                
                    conn.Execute Sql
                    
                    ' insertamos una linea con la diferencia que nos queda
                    vcodigo = DevuelveValor("select max(coalesce(codigo,0)) from trzmovim")
                    vcodigo = vcodigo + 1
                    
                    Sql = "insert into trzmovim (codigo, numpalet, numalbar, fecha, codvarie, kilos) values "
                    Sql = Sql & "(" & DBSet(vcodigo, "N") & "," & DBSet(Rs2!NumPalet, "N") & ",0," & DBSet(Rs2!Fecha, "F") & "," & DBSet(Rs!Codvarie, "N") & ","
                    Sql = Sql & DBSet(resto, "N") & ")"
                    
                    conn.Execute Sql
                    
                    Salir = True
                End If
        
                Rs2.MoveNext
            Wend
            
            Set Rs2 = Nothing
            
        End If
        
    Wend
    Set Rs = Nothing
    
    RepartoAlbaranes = True
    Exit Function
    
eRepartoAlbaranes:
    vMens = "Reparto de Albaranes"
    
End Function


Private Function CargarPaletsConfeccionados(vSQL As String, vMens As String) As Boolean
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim Rs1 As ADODB.Recordset
Dim SQLinsert As String
Dim SqlInsert2 As String
Dim SqlInsert3 As String
Dim SqlValues As String
Dim NroPalet As Long
Dim Marca As Integer
Dim Forfait As String
Dim Calibre As Integer
Dim vcodigo As Long

    On Error GoTo eCargarPaletsConfeccionados

    CargarPaletsConfeccionados = False
    

    NroPalet = DevuelveValor("select max(numpalet) from palets")
    NroPalet = NroPalet + 1
    
    SQLinsert = "insert into palets (numpalet,fechaini,horaini,fechafin,horafin,codpalet,linconfe,tipmercan,"
    SQLinsert = SQLinsert & "fechaconf,horaiconf,horafconf,codlinconf,intorden,linentrada,linsalida,idpalet) values "
    
    SqlInsert2 = "insert into palets_variedad (numpalet,numlinea,codvarie,codvarco,codmarca,codforfait,pesobrut,pesoneto,numcajas) values "
    
    SqlInsert3 = "insert into palets_calibre (numpalet,numlinea,numline1,codvarie,codcalib,numcajas) values "
    
    Marca = DevuelveValor("select min(codmarca) from marcas")
    Forfait = DevuelveValor("select min(codforfait) from forfaits")
    vcodigo = DevuelveValor("select max(coalesce(codigo,0)) from trzmovim")
    
    Set Rs = New ADODB.Recordset
    Rs.Open vSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not Rs.EOF
    
        Calibre = DevuelveValor("select min(codcalib) from calibres where codvarie = " & DBSet(Rs!Codvarie, "N"))
    
        SqlValues = "(" & DBSet(NroPalet, "N") & "," & DBSet(txtCodigo(30).Text, "F") & "," & DBSet(txtCodigo(30).Text & " 00:00:00", "FH") & ","
        SqlValues = SqlValues & DBSet(txtCodigo(30).Text, "F") & "," & DBSet(txtCodigo(30).Text & " 00:00:00", "FH") & ",1,1,0,"
        SqlValues = SqlValues & DBSet(txtCodigo(30).Text, "F") & "," & DBSet(txtCodigo(30).Text & " 00:00:00", "FH") & ",1,1,1,1,"
        SqlValues = SqlValues & DBSet(Rs!IdPalet, "N") & ")"
    
        conn.Execute SQLinsert & SqlValues
    
        Sql = "select * from trzpalets where idpalet = " & DBSet(Rs!IdPalet, "N")
        
        Set Rs1 = New ADODB.Recordset
        Rs1.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        If Not Rs1.EOF Then
            'palets_variedad
            SqlValues = "(" & DBSet(NroPalet, "N") & ",1," & DBSet(Rs1!Codvarie, "N") & "," & DBSet(Rs1!Codvarie, "N") & "," & DBSet(Marca, "N") & ","
            SqlValues = SqlValues & DBSet(Forfait, "T") & "," & DBSet(Rs1!NumKilos, "N") & "," & DBSet(Rs1!NumKilos, "N") & "," & DBSet(Rs1!NumCajones, "N") & ")"
            
            conn.Execute SqlInsert2 & SqlValues
            
            'palets_calibre
            SqlValues = "(" & DBSet(NroPalet, "N") & ",1,1," & DBSet(Rs1!Codvarie, "N") & "," & DBSet(Calibre, "N") & "," & DBSet(Rs1!NumCajones, "N") & ")"
            
            conn.Execute SqlInsert2 & SqlValues
        End If
        
        ' metemos en la tabla de movimientos de traza
        vcodigo = vcodigo + 1
        
        Sql = "insert into trzmovim (codigo, numpalet, numalbar, fecha, codvarie, kilos) values "
        Sql = Sql & "(" & DBSet(vcodigo, "N") & "," & DBSet(NroPalet, "N") & ",0," & DBSet(txtCodigo(30).Text, "F") & "," & DBSet(Rs1!Codvarie, "N") & ","
        Sql = Sql & DBSet(Rs1!NumKilos, "N") & ")"
        
        conn.Execute Sql
        
        Set Rs1 = Nothing
        Rs.MoveNext
    Wend
    Set Rs = Nothing

    CargarPaletsConfeccionados = True
    
    Exit Function

eCargarPaletsConfeccionados:
    vMens = "Cargar Palets Confeccionados:" & vbCrLf & Err.Description
End Function


Private Sub cmdAceptar_Click(Index As Integer)
Dim cOrden As String
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim numOp As Byte
Dim indRPT As Byte 'Indica el tipo de Documento en la tabla "scryst"
Dim nomDocu As String 'Nombre de Informe rpt de crystal
Dim vsqlDestino As String

    
    InicializarVbles
    
    'Añadir el parametro de Empresa
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    Select Case Index
       Case 0 'Frame Informe de articulos
            '======== FORMULA  ====================================
            'D/H Producto
            cDesde = Trim(txtCodigo(6).Text)
            cHasta = Trim(txtCodigo(7).Text)
            nDesde = txtNombre(6).Text
            nHasta = txtNombre(7).Text
            If Not (cDesde = "" And cHasta = "") Then
                'Cadena para seleccion Desde y Hasta
                Codigo = "{" & tabla & ".codprodu}"
                TipCod = "N"
                If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHProducto= """) Then Exit Sub
            End If
            
            'D/H Variedad
            cDesde = Trim(txtCodigo(0).Text)
            cHasta = Trim(txtCodigo(1).Text)
            nDesde = txtNombre(0).Text
            nHasta = txtNombre(1).Text
            If Not (cDesde = "" And cHasta = "") Then
                'Cadena para seleccion Desde y Hasta
                Codigo = "{" & tabla & ".codvarie}"
                TipCod = "N"
                If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHVariedad= """) Then Exit Sub
            End If
            
            'Obtener el parametro con el ORDEN del Informe
            '---------------------------------------------
        '    numOp = PonerGrupo(1, ListView1.ListItems(1).Text)
        '    numOp = PonerGrupo(2, ListView1.ListItems(2).Text)
        ' ### [Monica] 10/11/2006    he sustituido las dos anteriores instrucciones por la siguiente
            If ListView1(1).SelectedItem = "Producto" Then
                numOp = PonerGrupo(1, ListView1(1).ListItems(1).Text)
                numOp = PonerGrupo(2, ListView1(1).ListItems(2).Text)
            Else
                numOp = PonerGrupo(1, ListView1(1).ListItems(2).Text)
                numOp = PonerGrupo(2, ListView1(1).ListItems(1).Text)
            End If
'            Debug.Print cadParam
            
            cadNombreRPT = "rManVarie.rpt"
            cadTitulo = "Listado de Variedades"
            ConSubInforme = False
            
       Case 1 'Frame Informe de calibres
            '======== FORMULA  ====================================
            'D/H Variedad
            cDesde = Trim(txtCodigo(8).Text)
            cHasta = Trim(txtCodigo(9).Text)
            nDesde = txtNombre(8).Text
            nHasta = txtNombre(9).Text
            If Not (cDesde = "" And cHasta = "") Then
                'Cadena para seleccion Desde y Hasta
                Codigo = "{" & tabla & ".codvarie}"
                TipCod = "N"
                If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHVariedad= """) Then Exit Sub
            End If
            
            'D/H Calibre
            cDesde = Trim(txtCodigo(10).Text)
            cHasta = Trim(txtCodigo(11).Text)
            nDesde = txtNombre(10).Text
            nHasta = txtNombre(11).Text
            If Not (cDesde = "" And cHasta = "") Then
                'Cadena para seleccion Desde y Hasta
                Codigo = "{" & tabla & ".codcalib}"
                TipCod = "N"
                If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHCalibre= """) Then Exit Sub
            End If
            
            'Obtener el parametro con el ORDEN del Informe
            '---------------------------------------------
        '    numOp = PonerGrupo(1, ListView1.ListItems(1).Text)
        '    numOp = PonerGrupo(2, ListView1.ListItems(2).Text)
        ' ### [Monica] 10/11/2006    he sustituido las dos anteriores instrucciones por la siguiente
            If Opcion(0).Value Then numOp = PonerGrupo(1, ListView1(3).ListItems(2).Text)
            If Opcion(1).Value Then numOp = PonerGrupo(1, ListView1(3).ListItems(1).Text)
            
            cadNombreRPT = "rManCalibres.rpt"
            cadTitulo = "Listado de Calibres"
            ConSubInforme = False
        
        Case 2 ' informe de movimiento de envases
            albaranes = ""
            '******************************************************
            ' SOLO SACAMOS LOS REGISTROS DE LA TABLA ALBARAN_ENVASE
            '******************************************************
            If Me.Check2.Value = 0 Then
                '======== FORMULA  ====================================
                'D/H TIPO DE ENVASE
                cDesde = Trim(txtCodigo(12).Text)
                cHasta = Trim(txtCodigo(13).Text)
                nDesde = txtNombre(12).Text
                nHasta = txtNombre(13).Text
                If Not (cDesde = "" And cHasta = "") Then
                    'Cadena para seleccion Desde y Hasta
                    Codigo = "{sartic.codtipar}"
                    TipCod = "T"
                    If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHTipo=""") Then Exit Sub
                End If
                
                'D/H ARTICULO
                cDesde = Trim(txtCodigo(20).Text)
                cHasta = Trim(txtCodigo(21).Text)
                nDesde = txtNombre(20).Text
                nHasta = txtNombre(21).Text
                If Not (cDesde = "" And cHasta = "") Then
                    'Cadena para seleccion Desde y Hasta
                    Codigo = "{albaran_envase.codartic}"
                    TipCod = "T"
                    If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHArticulo=""") Then Exit Sub
                End If
                
                'D/H CLIENTE
                cDesde = Trim(txtCodigo(22).Text)
                cHasta = Trim(txtCodigo(23).Text)
                nDesde = txtNombre(22).Text
                nHasta = txtNombre(23).Text
                If Not (cDesde = "" And cHasta = "") Then
                    'Cadena para seleccion Desde y Hasta
                    Codigo = "{albaran_envase.codclien}"
                    TipCod = "N"
                    If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHCliente=""") Then Exit Sub
                End If
                
                '[Monica]22/10/2012: añadido desde/hasta destino
                'D/H Destino
                vsqlDestino = ""
                If txtCodigo(26).Text <> "" Then vsqlDestino = vsqlDestino & " and destinos.coddesti >= " & DBSet(txtCodigo(26).Text, "N")
                If txtCodigo(28).Text <> "" Then vsqlDestino = vsqlDestino & " and destinos.coddesti <= " & DBSet(txtCodigo(28).Text, "N")
                
'                If vsqlDestino <> "" And txtCodigo(22).Text = txtCodigo(23).Text And txtCodigo(22).Text <> "" Then
'                    Set frmMensDestino = New frmMensajes
'
'                    frmMensDestino.OpcionMensaje = 21
'                    frmMensDestino.Label5 = "Destinos"
'                    frmMensDestino.cadWHERE = vsqlDestino & " and destinos.codclien = " & txtCodigo(22).Text
'                    frmMensDestino.Show vbModal
'
'                    Set frmMensDestino = Nothing
'                End If
                
                
                'D/H fecha movimiento
                cDesde = Trim(txtCodigo(14).Text)
                cHasta = Trim(txtCodigo(15).Text)
                If Not (cDesde = "" And cHasta = "") Then
                    'Cadena para seleccion Desde y Hasta
                    Codigo = "{albaran_envase.fechamov}"
                    TipCod = "F"
                    If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHFecha=""") Then Exit Sub
                End If
    
                cadParam = cadParam & "pOrden=" & Me.Check1.Value & "|"
                numParam = numParam + 1
                
                If Me.Check1.Value = 1 Then
                    cadParam = cadParam & "pGroup1={albaran_envase.codclien}|"
                    cadParam = cadParam & "pGroupName1={clientes.nomclien}|"
                Else
                    cadParam = cadParam & "pGroup1={albaran_envase.codartic}|"
                    cadParam = cadParam & "pGroupName1={sartic.nomartic}|"
                End If
                numParam = numParam + 2
    
    
                If Check4.Value = 0 Then
                    cadNombreRPT = "rMovEnvasesRet.rpt"
                    
                    indRPT = 64 'Personalizamos el informe para Picassent
                    If Not PonerParamRPT(indRPT, cadParam, numParam, nomDocu) Then Exit Sub
                    
                    cadNombreRPT = nomDocu
                    
                Else
                    cadParam = cadParam & "pSinSaldosCero=" & Check5.Value & "|"
                    numParam = numParam + 1
                    
                    cadNombreRPT = "rMovEnvasesRetSaldo.rpt"
                    
                    indRPT = 65 'Personalizamos el informe para Picassent
                    If Not PonerParamRPT(indRPT, cadParam, numParam, nomDocu) Then Exit Sub
                    
                    cadNombreRPT = nomDocu
                    
                    ' tenemos que insertar en tmpenvasesret los albaranes que sean: solo los que tienen saldo
                    ' o todos.
                    If CargarTablaTemporal2 Then
                        If HayRegParaInforme("tmpinformes", "codusu= " & vUsu.Codigo) Then
                            cadFormula = "{tmpinformes.codusu} = " & vUsu.Codigo
                            
                            cadParam = Replace(cadParam, "{albaran_envase.codclien}", "{tmpinformes.codigo1}")
                            cadParam = Replace(cadParam, "{albaran_envase.codartic}", "{tmpinformes.nombre1}")
                            
                            cadTitulo = "Informe de Movimientos de Envases"
                            ConSubInforme = True
                            
                            LlamarImprimir
                            
                            Exit Sub
                        End If
                    End If
                End If
                cadTitulo = "Informe de Movimientos de Envases"
                ConSubInforme = True
                
                tabla = "(albaran_envase INNER JOIN sartic on albaran_envase.codartic = sartic.codartic)"
                tabla = tabla & " INNER JOIN stipar on sartic.codtipar = stipar.codtipar "
            Else
            '******************************************************
            ' SACAMOS LOS REGISTROS DE LAS TABLAS: ALBARAN_ENVASE Y SMOVAL
            '******************************************************
                InicializarVbles
                cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
                numParam = numParam + 1
                 
                 'D/H TIPO DE ENVASE
                cDesde = Trim(txtCodigo(12).Text)
                cHasta = Trim(txtCodigo(13).Text)
                nDesde = txtNombre(12).Text
                nHasta = txtNombre(13).Text
                If Not (cDesde = "" And cHasta = "") Then
                    'Cadena para seleccion Desde y Hasta
                    Codigo = "{sartic.codtipar}"
                    TipCod = "T"
                    If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHTipo=""") Then Exit Sub
                End If
                
                'D/H ARTICULO
                cDesde = Trim(txtCodigo(20).Text)
                cHasta = Trim(txtCodigo(21).Text)
                nDesde = txtNombre(20).Text
                nHasta = txtNombre(21).Text
                If Not (cDesde = "" And cHasta = "") Then
                    'Cadena para seleccion Desde y Hasta
                    Codigo = "{albaran_envase.codartic}"
                    TipCod = "T"
                    If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHArticulo=""") Then Exit Sub
                End If
                
                'D/H CLIENTE
                cDesde = Trim(txtCodigo(22).Text)
                cHasta = Trim(txtCodigo(23).Text)
                nDesde = txtNombre(22).Text
                nHasta = txtNombre(23).Text
                If Not (cDesde = "" And cHasta = "") Then
                    'Cadena para seleccion Desde y Hasta
                    Codigo = "{albaran_envase.codclien}"
                    TipCod = "N"
                    If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHCliente=""") Then Exit Sub
                End If
                
                '[Monica]22/10/2012: añadido desde/hasta destino
                'D/H Destino
                vsqlDestino = ""
                If txtCodigo(26).Text <> "" Then vsqlDestino = vsqlDestino & " and destinos.coddesti >= " & DBSet(txtCodigo(26).Text, "N")
                If txtCodigo(28).Text <> "" Then vsqlDestino = vsqlDestino & " and destinos.coddesti <= " & DBSet(txtCodigo(28).Text, "N")
                
'                If vsqlDestino <> "" And txtCodigo(26).Text <> txtCodigo(28).Text And txtCodigo(22).Text = txtCodigo(23).Text And txtCodigo(22).Text <> "" Then
'                    Set frmMensDestino = New frmMensajes
'
'                    frmMensDestino.OpcionMensaje = 21
'                    frmMensDestino.Label5 = "Destinos"
'                    frmMensDestino.cadWHERE = vsqlDestino & " and destinos.codclien = " & txtCodigo(22).Text
'                    frmMensDestino.Show vbModal
'
'                    Set frmMensDestino = Nothing
'                End If
                
                'D/H fecha movimiento
                cDesde = Trim(txtCodigo(14).Text)
                cHasta = Trim(txtCodigo(15).Text)
                If Not (cDesde = "" And cHasta = "") Then
                    'Cadena para seleccion Desde y Hasta
                    Codigo = "{albaran_envase.fechamov}"
                    TipCod = "F"
                    If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHFecha=""") Then Exit Sub
                End If
                
                'Añadir el parametro de Empresa

                If CargarTablaTemporal Then
                    indRPT = 63 'Personalizamos el informe para Picassent
                    If Not PonerParamRPT(indRPT, cadParam, numParam, nomDocu) Then Exit Sub
                
                    ConSubInforme = True
                    
                    cadNombreRPT = nomDocu '  "rMovEnvasesRetCompras.rpt"
                    cadTitulo = "Informe de Movimientos de Envases"
                    
                    If HayRegParaInforme("tmpenvasesret", "codusu= " & vUsu.Codigo) Then
                        cadFormula = "{tmpenvasesret.codusu} = " & vUsu.Codigo
                        LlamarImprimir
                    End If
                End If
                Exit Sub
            End If
            
        Case 3 ' informe de horas trabajadas
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
            
            'D/H fecha
            cDesde = Trim(txtCodigo(16).Text)
            cHasta = Trim(txtCodigo(17).Text)
            If Not (cDesde = "" And cHasta = "") Then
                'Cadena para seleccion Desde y Hasta
                Codigo = "{horas.fechahora}"
                TipCod = "F"
                If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHFecha=""") Then Exit Sub
            End If

            cadParam = cadParam & "pHProductivas=" & Me.Check3.Value & "|"
            numParam = numParam + 1
            
            ConSubInforme = False
            cadNombreRPT = "rManHorasTrab.rpt"
            cadTitulo = "Informe de Horas Trabajadas"
    
    
        Case 4 ' informe de palets en camaras
            '======== FORMULA  ====================================
            'D/H camara
            cDesde = Trim(txtCodigo(32).Text)
            cHasta = Trim(txtCodigo(33).Text)
            nDesde = txtNombre(32).Text
            nHasta = txtNombre(33).Text
            If Not (cDesde = "" And cHasta = "") Then
                'Cadena para seleccion Desde y Hasta
                Codigo = "{palets.codcamara}"
                TipCod = "N"
                If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHCamara= """) Then Exit Sub
            End If
            
            'D/H fecha
            cDesde = Trim(txtCodigo(29).Text)
            cHasta = Trim(txtCodigo(31).Text)
            If Not (cDesde = "" And cHasta = "") Then
                'Cadena para seleccion Desde y Hasta
                Codigo = "{palets.fechaconf}"
                TipCod = "F"
                If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHFecha=""") Then Exit Sub
            End If
            
            'Obtener el parametro con el ORDEN del Informe
            '---------------------------------------------
        '    numOp = PonerGrupo(1, ListView1.ListItems(1).Text)
        '    numOp = PonerGrupo(2, ListView1.ListItems(2).Text)
        ' ### [Monica] 10/11/2006    he sustituido las dos anteriores instrucciones por la siguiente
            If Opcion(2).Value Then numOp = PonerGrupo(1, ListView1(5).ListItems(1).Text)
            If Opcion(3).Value Then numOp = PonerGrupo(1, ListView1(5).ListItems(2).Text)
            
            If Opcion(2).Value Then
                cadParam = cadParam & "pOrden=0|"
                cadParam = cadParam & "pOrden1={tmpinformes.fecha1}|"
            Else
                cadParam = cadParam & "pOrden=1|"
                cadParam = cadParam & "pOrden1={tmpinformes.importe2}|"
            End If
            numParam = numParam + 2
            
            
            If CargarTablaTemporalPalets Then
                cadNombreRPT = "rInfPaletsCamaras.rpt"
                cadTitulo = "Informe de Palets en Cámaras"
                ConSubInforme = True
                tabla = "tmpinformes"
                cadselect = "{tmpinformes.codusu} = " & vUsu.Codigo
'                cadParam = cadParam & "pUsu=" & vUsu.Codigo & "|"
'                numParam = numParam + 1
                cadFormula = "{tmpinformes.codusu} = " & vUsu.Codigo
            End If
    End Select
    
    'Comprobar si hay registros a Mostrar antes de abrir el Informe
    If HayRegParaInforme(tabla, cadselect) Then
        LlamarImprimir
    End If

End Sub

Private Function CargarTablaTemporalPalets() As Boolean
Dim Sql As String
Dim SQL1 As String
Dim Rs As ADODB.Recordset

    On Error GoTo eCargarTablaTemporal
    
    CargarTablaTemporalPalets = False

    Sql = "delete from tmpinformes where codusu = " & DBSet(vUsu.Codigo, "N")
    conn.Execute Sql

    Sql = "delete from tmpinformes2 where codusu = " & DBSet(vUsu.Codigo, "N")
    conn.Execute Sql

    Sql = "select " & vUsu.Codigo & ", palets.numpalet, palets.codcamara, palets.fechaconf, palets_variedad.codforfait, forfaits.nomconfe,  "
    Sql = Sql & " palets_variedad.pesoneto, palets_variedad.numcajas, palets.numpedid, camaras.nomcamara, 0, palets_variedad.numlinea, palets_variedad.codvarie  "
    Sql = Sql & " from (((palets LEFT JOIN camaras on palets.codcamara = camaras.codcamara) inner join palets_variedad on palets.numpalet = palets_variedad.numpalet) inner join forfaits on palets_variedad.codforfait = forfaits.codforfait) "
    Sql = Sql & " where (1=1) "
    
    If txtCodigo(32).Text <> "" Then Sql = Sql & " and palets.codcamara >= " & DBSet(txtCodigo(32).Text, "N")
    If txtCodigo(33).Text <> "" Then Sql = Sql & " and palets.codcamara <= " & DBSet(txtCodigo(33).Text, "N")
    
    If txtCodigo(29).Text <> "" Then Sql = Sql & " and palets.fechaconf >= " & DBSet(txtCodigo(29).Text, "F")
    If txtCodigo(31).Text <> "" Then Sql = Sql & " and palets.fechaconf <= " & DBSet(txtCodigo(31).Text, "F")
    
    
    SQL1 = "insert into tmpinformes (codusu, importe1, importe2, fecha1, nombre3, nombre1, importe3, importe4, importe5, nombre2, importeb1, importeb2, importeb3)  " & Sql
    conn.Execute SQL1
    
    ' marcamos que palets estan en pedidos y albaranes, es decir, han salido
    Sql = "update tmpinformes, pedidos  set importeb1 = 1 "
    Sql = Sql & " where tmpinformes.codusu = " & DBSet(vUsu.Codigo, "N")
    Sql = Sql & " and tmpinformes.importe5 = pedidos.numpedid "
    Sql = Sql & " and not pedidos.numalbar is null and pedidos.numalbar <> 0 "
    conn.Execute Sql
    
    SQL1 = "insert into tmpinformes2 (codusu, importe1, importe2, importe3, nombre1, importe4, importe5) select " & vUsu.Codigo & ", palets_calibre.numpalet, palets_calibre.numlinea, palets_calibre.codcalib, calibres.nomcalib, palets_calibre.numcajas, palets_calibre.codvarie "
    SQL1 = SQL1 & " from tmpinformes, palets_calibre, calibres where tmpinformes.codusu = " & vUsu.Codigo & " and tmpinformes.importe1 = palets_calibre.numpalet "
    SQL1 = SQL1 & " and tmpinformes.importeb2 = palets_calibre.numlinea and palets_calibre.codcalib = calibres.codcalib "
    SQL1 = SQL1 & " and tmpinformes.importeb3 = calibres.codvarie "
    conn.Execute SQL1
    
    CargarTablaTemporalPalets = True
    Exit Function
    
eCargarTablaTemporal:
    MuestraError Err.Number, "Carga Tabla Temporal de Palets"
End Function



'Frame Informe Clientes

Private Sub cmdAceptar2_Click()
Dim cOrden As String
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim numOp As Byte

    InicializarVbles
    
    'Añadir el parametro de Empresa
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    '======== FORMULA  ====================================
    
    'D/H Cliente
    cDesde = Trim(txtCodigo(4).Text)
    cHasta = Trim(txtCodigo(5).Text)
    nDesde = txtNombre(4).Text
    nHasta = txtNombre(5).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        If Opcionlistado = 10 Then
            Codigo = "{" & tabla & ".codclien}"
        ElseIf Opcionlistado = 14 Then
            Codigo = "{" & tabla & ".gruprove}"
        End If
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHCliente= """) Then Exit Sub
    End If
    
    'Obtener el parametro con el ORDEN del Informe
    '---------------------------------------------
'    numOp = PonerGrupo(1, ListView1.ListItems(1).Text)
'    numOp = PonerGrupo(2, ListView1.ListItems(2).Text)
' ### [Monica] 10/11/2006    he sustituido las dos anteriores instrucciones por la siguiente
    numOp = PonerOrden(ListView1(0).SelectedItem.Text)

    'Comprobar si hay registros a Mostrar antes de abrir el Informe
    If HayRegParaInforme(tabla, cadselect) Then
        cadNombreRPT = "rManClien.rpt"
        cadTitulo = "Listado de Clientes " & Tipo
        cadParam = cadParam & "pTipo= """ & Tipo & """|"
        numParam = numParam + 1
        ConSubInforme = False
        LlamarImprimir
    End If
End Sub


Private Sub CmdAceptar3_Click()
Dim cOrden As String
Dim cDesde As String, cHasta As String 'cadena codigo Desde/Hasta
Dim nDesde As String, nHasta As String 'cadena Descripcion Desde/Hasta
Dim numOp As Byte

    InicializarVbles
    
    'Añadir el parametro de Empresa
    cadParam = cadParam & "|pEmpresa=""" & vEmpresa.nomempre & """|"
    numParam = numParam + 1
    
    '======== FORMULA  ====================================
    
    'D/H Cliente
    cDesde = Trim(txtCodigo(2).Text)
    cHasta = Trim(txtCodigo(3).Text)
    nDesde = txtNombre(2).Text
    nHasta = txtNombre(3).Text
    If Not (cDesde = "" And cHasta = "") Then
        'Cadena para seleccion Desde y Hasta
        Codigo = "{" & tabla & ".codprove}"
        TipCod = "N"
        If Not PonerDesdeHasta(cDesde, cHasta, nDesde, nHasta, "pDHProveedor= """) Then Exit Sub
    End If
    
    'Obtener el parametro con el ORDEN del Informe
    '---------------------------------------------
'    numOp = PonerGrupo(1, ListView1.ListItems(1).Text)
'    numOp = PonerGrupo(2, ListView1.ListItems(2).Text)
' ### [Monica] 10/11/2006    he sustituido las dos anteriores instrucciones por la siguiente
    numOp = PonerOrden(ListView1(2).SelectedItem.Text)

    'Comprobar si hay registros a Mostrar antes de abrir el Informe
    If HayRegParaInforme(tabla, cadselect) Then
        cadNombreRPT = "rManProve.rpt"
        cadTitulo = "Listado de Proveedores " & Tipo
        cadParam = cadParam & "pTipo= """ & Tipo & """|"
        numParam = numParam + 1
        ConSubInforme = False
        LlamarImprimir
    End If

End Sub

'Private Sub CmdAcepTraspaso_Click()
'Dim Sql As String
'
'    Sql = ""
'    If txtCodigo(35).Text = "" Then
'        MsgBox "Debe introducir el cliente de la empresa destino.", vbExclamation
'        PonerFoco txtCodigo(35)
'    Else
'        'comprobamos que exista en la base de datos destino
'        Sql = DevuelveDesdeBDNew(cAgro, vParamAplic.BDDestino & ".clientes", "nomclien", "codclien", txtCodigo(35).Text, "N")
'        If Sql = "" Then
'            MsgBox "No existe cliente en la empresa destino. Reintroduzca.", vbExclamation
'            PonerFoco txtCodigo(35)
'        Else
'            txtNombre(35).Text = Sql
'            ' comprobacion del destino
'            If txtCodigo(34).Text = "" Then
'                MsgBox "Debe introducir el destino del cliente de la empresa destino.", vbExclamation
'                PonerFoco txtCodigo(34)
'            Else
'                Sql = DevuelveDesdeBDNew(cAgro, vParamAplic.BDDestino & ".destinos", "nomdesti", "codclien", txtCodigo(35).Text, "N", , "coddesti", txtCodigo(34).Text, "N")
'                If Sql = "" Then
'                    MsgBox "No existe el destino del cliente en la empresa destino. Reintroduzca.", vbExclamation
'                    PonerFoco txtCodigo(34)
'                Else
'                    txtNombre(34).Text = Sql
'                End If
'            End If
'        End If
'    End If
'
'    ' comprobacion de si existe el albaran en la bd destino
'    Dim ExisteAlb As String
'    ExisteAlb = DevuelveDesdeBDNew(cAgro, vParamAplic.BDDestino & ".albaran", "numalbar", "numalbar", NumCod, "N")
'    If ExisteAlb <> "" Then
'        If MsgBox("Albaran existente en la empresa destino. ¿ Desea modificarlo ?", vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then Exit Sub
'    End If
'
'    ' comprobacion de las claves refereciales
'    If ComprobarReferenciales(NumCod) Then
'        If TraspasoAlbaran(NumCod, ExisteAlb) Then
'            MsgBox "Proceso realizado correctamente.", vbExclamation
'            cmdCancel_Click (1)
'        End If
'    End If
'
'End Sub


'Private Function TraspasoAlbaran(Albaran As String, ExisteAlb As String) As Boolean
'Dim Sql As String
'
'    On Error GoTo eTraspasoAlbaran
'
'    TraspasoAlbaran = False
'
'    If ExisteAlb = "" Then
'        ' insertamos
'        'albaran
'        Sql = "insert into " & vParamAplic.BDDestino & ".albaran "
'        Sql = Sql & "(numalbar,fechaalb,codclien,coddesti,codtrans,matriveh,matrirem,refclien,codtimer,"
'        Sql = Sql & "totpalet,portespre,nrocontra,nroactas,numpedid,fechaped,observac,pasaridoc,"
'        Sql = Sql & "codalmac,portespag,paletspag,numerocmr,comisionespre,comisionespag,codcomis,"
'        Sql = Sql & "codsocio,airline,AWB,flight1,flight2,airorigin,airdestiny,ETD,ETA,precnodef,"
'        Sql = Sql & "estacomunicada,codtipom)"
'        Sql = Sql & " select numalbar, fechaalb, " & DBSet(txtCodigo(35).Text, "N") & "," & DBSet(txtCodigo(34).Text, "N") & ",codtrans,matriveh,matrirem,refclien,codtimer,"
'        Sql = Sql & "totpalet,portespre,nrocontra,nroactas,numpedid,fechaped,observac,pasaridoc,"
'        Sql = Sql & "codalmac,portespag,paletspag,numerocmr,comisionespre,comisionespag,codcomis,"
'        Sql = Sql & "codsocio,airline,AWB,flight1,flight2,airorigin,airdestiny,ETD,ETA,precnodef,"
'        Sql = Sql & "estacomunicada,codtipom "
'        Sql = Sql & " from albaran "
'        Sql = Sql & " where numalbar = " & DBSet(Albaran, "N")
'
'        conn.Execute Sql
'
'        'albaran_variedad
'        Sql = "insert into " & vParamAplic.BDDestino & ".albaran_variedad "
'        Sql = Sql & "(numalbar,numlinea,codvarie,codvarco,codforfait,codmarca,categori,totpalet,numcajas,"
'        Sql = Sql & "pesobrut,pesoneto,preciopro,preciodef,codincid,impcomis,observac,unidades,"
'        Sql = Sql & "referencia,codpalet,nrotraza,codtipo,sefactura,codcomis,nrotraza1,nrotraza2,nrotraza3,"
'        Sql = Sql & "nrotraza4,nrotraza5,nrotraza6,expediente) "
'        Sql = Sql & " select numalbar,numlinea,codvarie,codvarco,codforfait,codmarca,categori,totpalet,numcajas, "
'        Sql = Sql & "pesobrut,pesoneto,preciopro,preciodef,codincid,impcomis,observac,unidades,"
'        Sql = Sql & "referencia,codpalet,nrotraza,codtipo,sefactura,codcomis,nrotraza1,nrotraza2,nrotraza3,"
'        Sql = Sql & "nrotraza4,nrotraza5,nrotraza6,expediente "
'        Sql = Sql & " from albaran_variedad "
'        Sql = Sql & " where numalbar = " & DBSet(Albaran, "N")
'
'        conn.Execute Sql
'
'        'albaran_calibre
'        Sql = "insert into " & vParamAplic.BDDestino & ".albaran_calibre "
'        Sql = Sql & "(numalbar,numlinea,numline1,codvarie,codcalib,numcajas,pesobrut,pesoneto,unidades,preciopro) "
'        Sql = Sql & " select numalbar,numlinea,numline1,codvarie,codcalib,numcajas,pesobrut,pesoneto,unidades,preciopro "
'        Sql = Sql & " from albaran_calibre "
'        Sql = Sql & " where numalbar = " & DBSet(Albaran, "N")
'
'        conn.Execute Sql
'
'        'albaran_costes
'        Sql = "insert into " & vParamAplic.BDDestino & ".albaran_costes "
'        Sql = Sql & "(numalbar,numlinea,tipogasto,codcoste,impcoste,importes,unidades,codartic) "
'        Sql = Sql & " select numalbar,numlinea,tipogasto,codcoste,impcoste,importes,unidades,codartic "
'        Sql = Sql & " from albaran_costes "
'        Sql = Sql & " where numalbar = " & DBSet(Albaran, "N")
'
'        conn.Execute Sql
'
'        'albaran_envase
'        Sql = "insert into " & vParamAplic.BDDestino & ".albaran_envase "
'        Sql = Sql & "(numalbar,numlinea,fechamov,codartic,tipomovi,cantidad,codclien,impfianza,factura,fecfactu) "
'        Sql = Sql & " select numalbar,numlinea,fechamov,codartic,tipomovi,cantidad,codclien,impfianza,factura,fecfactu "
'        Sql = Sql & " from albaran_envase "
'        Sql = Sql & " where numalbar = " & DBSet(Albaran, "N")
'
'        conn.Execute Sql
'
'        'albaran_palets
'        Sql = "insert into " & vParamAplic.BDDestino & ".albaran_palets "
'        Sql = Sql & "(numalbar,numlinea,numpalet) "
'        Sql = Sql & " select numalbar,numlinea,numpalet "
'        Sql = Sql & " from albaran_palets "
'        Sql = Sql & " where numalbar = " & DBSet(Albaran, "N")
'
'        conn.Execute Sql
'
'        'albaran_costreal
'        Sql = "insert into " & vParamAplic.BDDestino & ".albaran_costreal "
'        Sql = Sql & "(numalbar,numlinea,codcoste,impcoste) "
'        Sql = Sql & " select numalbar,numlinea,codcoste,impcoste "
'        Sql = Sql & " from albaran_costreal "
'        Sql = Sql & " where numalbar = " & DBSet(Albaran, "N")
'
'        conn.Execute Sql
'
'    Else
'
'        'albaran
'        Sql = "update " & vParamAplic.BDDestino & ".albaran dd, albaran ff set "
'        Sql = Sql & "dd.numalbar = ff.numalbar,dd.fechaalb = ff.fechaalb, dd.codclien = ff.codclien,dd.coddesti = ff.coddesti,dd.codtrans=ff.codtrans,"
'        Sql = Sql & "dd.matriveh=ff.matriveh,dd.matrirem=ff.matrirem,dd.refclien=ff.refclien,dd.codtimer=dd.codtimer,"
'        Sql = Sql & "dd.totpalet=ff.totpalet,dd.portespre=ff.portespre,dd.nrocontra=ff.nrocontra,dd.nroactas=ff.nroactas,dd.numpedid=ff.numpedid,"
'        Sql = Sql & "dd.fechaped=ff.fechaped,dd.observac=ff.observac,dd.pasaridoc=ff.pasaridoc,"
'        Sql = Sql & "dd.codalmac=ff.codalmac,dd.portespag=ff.portespag,dd.paletspag=ff.paletspag,dd.numerocmr=ff.numerocmr,dd.comisionespre=ff.comisionespre,"
'        Sql = Sql & "dd.comisionespag=ff.comisionespag,dd.codcomis=ff.codcomis,"
'        Sql = Sql & "dd.codsocio=ff.codsocio,dd.airline=ff.airline,dd.AWB=ff.AWB,dd.flight1=ff.flight1,dd.flight2=ff.flight2,dd.airorigin=ff.airorigin,"
'        Sql = Sql & "dd.airdestiny=ff.airdestiny,dd.ETD=ff.ETD,dd.ETA=ff.ETA,dd.precnodef=ff.precnodef,"
'        Sql = Sql & "dd.estacomunicada=ff.estacomunicada,dd.codtipom=ff.codtipom "
'        Sql = Sql & " where dd.numalbar = " & DBSet(Albaran, "N")
'        Sql = Sql & " and dd.numalbar = ff.numalbar "
'
'        conn.Execute Sql
'
'        'albaran_variedad
'        Sql = "update " & vParamAplic.BDDestino & ".albaran_variedad dd, albaran_variedad ff set "
'        Sql = Sql & "dd.codvarie=ff.codvarie,dd.codvarco=ff.codvarco,dd.codforfait=ff.codforfait,"
'        Sql = Sql & "dd.codmarca=ff.codmarca,dd.categori=ff.categori,dd.totpalet=ff.totpalet,dd.numcajas=ff.numcajas,"
'        Sql = Sql & "dd.pesobrut=ff.pesobrut,dd.pesoneto=ff.pesoneto,dd.preciopro=ff.preciopro,dd.preciodef=ff.preciodef,dd.codincid=ff.codincid,"
'        Sql = Sql & "dd.impcomis=ff.impcomis,dd.observac=ff.observac,dd.unidades=ff.unidades,"
'        Sql = Sql & "dd.referencia=ff.referencia,dd.codpalet=ff.codpalet,dd.nrotraza=ff.nrotraza,dd.codtipo=ff.codtipo,dd.sefactura=ff.sefactura,"
'        Sql = Sql & "dd.codcomis=ff.codcomis,dd.nrotraza1=ff.nrotraza1,dd.nrotraza2=ff.nrotraza2,dd.nrotraza3=ff.nrotraza3,"
'        Sql = Sql & "dd.nrotraza4=ff.nrotraza4,dd.nrotraza5=ff.nrotraza6,dd.nrotraza6=ff.nrotraza6,dd.expediente=ff.expediente "
'        Sql = Sql & " where dd.numalbar = " & DBSet(Albaran, "N")
'        Sql = Sql & " and dd.numalbar = ff.numalbar and dd.numlinea = ff.numlinea "
'
'        conn.Execute Sql
'
'        'albaran_calibre
'        Sql = "update " & vParamAplic.BDDestino & ".albaran_calibre dd, albaran_calibre ff set "
'        Sql = Sql & "dd.codvarie = ff.codvarie, dd.codcalib = ff.codcalib,"
'        Sql = Sql & "dd.numcajas = ff.numcajas, dd.pesobrut = ff.pesobrut, dd.pesoneto = ff.pesoneto, dd.unidades = ff.unidades, dd.preciopro = ff.preciopro "
'        Sql = Sql & " where dd.numalbar = " & DBSet(Albaran, "N")
'        Sql = Sql & " and dd.numalbar = ff.numalbar and dd.numlinea = ff.numlinea and dd.numline1 = ff.numline1 "
'
'        conn.Execute Sql
'
'        'albaran_costes
'        Sql = "update " & vParamAplic.BDDestino & ".albaran_costes dd, albaran_costes ff set "
'        Sql = Sql & "dd.tipogasto=ff.tipogasto,dd.codcoste=ff.codcoste,dd.impcoste=ff.impcoste,"
'        Sql = Sql & "dd.importes = ff.importes,dd.unidades=ff.unidades,dd.codartic=ff.codartic "
'        Sql = Sql & " where dd.numalbar = " & DBSet(Albaran, "N")
'        Sql = Sql & " and dd.numalbar = ff.numalbar and dd.numlinea = ff.numlinea"
'
'        conn.Execute Sql
'
'        'albaran_envase
'        Sql = "update " & vParamAplic.BDDestino & ".albaran_envase dd, albaran_envase ff set "
'        Sql = Sql & "dd.fechamov=ff.fechamov,dd.codartic=ff.codartic,dd.tipomovi=ff.tipomovi,"
'        Sql = Sql & "dd.cantidad=ff.cantidad,dd.codclien=ff.codclien,dd.impfianza=ff.impfianza,dd.factura=ff.factura,dd.fecfactu=ff.fecfactu "
'        Sql = Sql & " where dd.numalbar = " & DBSet(Albaran, "N")
'        Sql = Sql & " and dd.numalbar = ff.numalbar and dd.numlinea=ff.numlinea"
'
'        conn.Execute Sql
'
'        'albaran_palets
'        Sql = "update " & vParamAplic.BDDestino & ".albaran_palets dd, albaran_palets ff set "
'        Sql = Sql & "dd.numpalet=ff.numpalet "
'        Sql = Sql & " where dd.numalbar = " & DBSet(Albaran, "N")
'        Sql = Sql & " and dd.numalbar = ff.numalbar and dd.numlinea=ff.numlinea"
'
'        conn.Execute Sql
'
'        'albaran_costreal
'        Sql = "update " & vParamAplic.BDDestino & ".albaran_costreal dd, albaran_costreal ff set "
'        Sql = Sql & "dd.codcoste=ff.codcoste,dd.impcoste=ff.impcoste "
'        Sql = Sql & " where dd.numalbar = " & DBSet(Albaran, "N")
'        Sql = Sql & " and dd.numalbar = ff.numalbar and dd.numlinea=ff.numlinea"
'
'        conn.Execute Sql
'
'
'    End If
'
'    TraspasoAlbaran = True
'    Exit Function
'
'eTraspasoAlbaran:
'    MuestraError Err.Number, "Traspaso Albaran", Err.Description
'End Function


'Private Function ComprobarReferenciales(Albaran As String) As Boolean
'Dim Sql As String
'Dim Rs As ADODB.Recordset
'
'    On Error GoTo eComprobarReferenciales
'
'    ComprobarReferenciales = False
'
'
'    'albaran
'    Sql = "select * from albaran where numalbar = " & DBSet(Albaran, "N")
'
'    Set Rs = New ADODB.Recordset
'    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'
'    If Not Rs.EOF Then
'
'        'agencias de transporte
'        Sql = DevuelveDesdeBDNew(cAgro, vParamAplic.BDDestino & ".agencias", "nomtrans", "codtrans", DBLet(Rs!codTrans, "N"), "N")
'        If Sql = "" Then
'            MsgBox "No existe la Agencia de Transporte " & DBLet(Rs!codTrans, "N") & ". Revise.", vbExclamation
'            Exit Function
'        End If
'
'        'tipo de mercado
'        Sql = DevuelveDesdeBDNew(cAgro, vParamAplic.BDDestino & ".tipomer", "nomtimer", "codtimer", DBLet(Rs!Codtimer, "N"), "N")
'        If Sql = "" Then
'            MsgBox "No existe el Tipo de Mercado " & DBLet(Rs!Codtimer, "N") & ". Revise.", vbExclamation
'            Exit Function
'        End If
'
'        'Almacen
'        Sql = DevuelveDesdeBDNew(cAgro, vParamAplic.BDDestino & ".salmpr", "nomalmac", "codalmac", DBLet(Rs!codAlmac, "N"), "N")
'        If Sql = "" Then
'            MsgBox "No existe el Almacén " & DBLet(Rs!codAlmac, "N") & ". Revise.", vbExclamation
'            Exit Function
'        End If
'
'        'Comisionista
'        If Not IsNull(Rs!codcomis) Then
'            Sql = DevuelveDesdeBDNew(cAgro, vParamAplic.BDDestino & ".agencias", "nomtrans", "codtrans", DBLet(Rs!codcomis, "N"), "N")
'            If Sql = "" Then
'                MsgBox "No existe el Comisionista " & DBLet(Rs!codcomis, "N") & ". Revise.", vbExclamation
'                Exit Function
'            End If
'        End If
'
'        'socio
'        If Not IsNull(Rs!Codsocio) Then
'            Sql = DevuelveDesdeBDNew(cAgro, vParamAplic.BDDestino & ".rsocios", "nomsocio", "codsocio", DBLet(Rs!Codsocio, "N"), "N")
'            If Sql = "" Then
'                MsgBox "No existe el Socio " & DBLet(Rs!Codsocio, "N") & ". Revise.", vbExclamation
'                Exit Function
'            End If
'        End If
'
'    End If
'    Set Rs = Nothing
'
'
'    'albaran_variedad
'    Sql = "select * from albaran_variedad where numalbar = " & DBSet(Albaran, "N")
'
'    Set Rs = New ADODB.Recordset
'    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'    While Not Rs.EOF
'
'        'variedad
'        Sql = DevuelveDesdeBDNew(cAgro, vParamAplic.BDDestino & ".variedades", "nomvarie", "codvarie", DBLet(Rs!Codvarie, "N"), "N")
'        If Sql = "" Then
'            MsgBox "No existe la Variedad " & DBLet(Rs!Codvarie, "N") & ". Revise.", vbExclamation
'            Exit Function
'        End If
'
'        'variedad comercial
'        Sql = DevuelveDesdeBDNew(cAgro, vParamAplic.BDDestino & ".variedades", "nomvarie", "codvarie", DBLet(Rs!codvarco, "N"), "N")
'        If Sql = "" Then
'            MsgBox "No existe la Variedad Comercial " & DBLet(Rs!codvarco, "N") & ". Revise.", vbExclamation
'            Exit Function
'        End If
'
'        'marca
'        Sql = DevuelveDesdeBDNew(cAgro, vParamAplic.BDDestino & ".marcas", "nommarca", "codmarca", DBLet(Rs!Codmarca, "N"), "N")
'        If Sql = "" Then
'            MsgBox "No existe la Marca " & DBLet(Rs!Codmarca, "N") & ". Revise.", vbExclamation
'            Exit Function
'        End If
'
'        'forfaits
'        Sql = DevuelveDesdeBDNew(cAgro, vParamAplic.BDDestino & ".forfaits", "nomconfe", "codforfait", DBLet(Rs!codforfait, "T"), "T")
'        If Sql = "" Then
'            MsgBox "No existe el Forfait " & DBLet(Rs!codforfait, "T") & ". Revise.", vbExclamation
'            Exit Function
'        End If
'
'        'incidencia
'        Sql = DevuelveDesdeBDNew(cAgro, vParamAplic.BDDestino & ".inciden", "nomincid", "codincid", DBLet(Rs!Codincid, "N"), "N")
'        If Sql = "" Then
'            MsgBox "No existe la Incidencia " & DBLet(Rs!Codincid, "N") & ". Revise.", vbExclamation
'            Exit Function
'        End If
'
'        'palet
'        If Not IsNull(Rs!codpalet) Then
'            Sql = DevuelveDesdeBDNew(cAgro, vParamAplic.BDDestino & ".confpale", "nompalet", "codpalet", DBLet(Rs!codpalet, "N"), "N")
'            If Sql = "" Then
'                MsgBox "No existe el Tipo de Palet " & DBLet(Rs!codpalet, "N") & ". Revise.", vbExclamation
'                Exit Function
'            End If
'        End If
'
'        'tipo de variedad
'        If Not IsNull(Rs!codtipo) Then
'            Sql = DevuelveDesdeBDNew(cAgro, vParamAplic.BDDestino & ".tipovarie", "nomtipo", "codtipo", DBLet(Rs!codtipo, "N"), "N")
'            If Sql = "" Then
'                MsgBox "No existe el Tipo de Variedad " & DBLet(Rs!codtipo, "N") & ". Revise.", vbExclamation
'                Exit Function
'            End If
'        End If
'
'        Rs.MoveNext
'
'    Wend
'
'    Set Rs = Nothing
'
'    'albaran_calibre
'    Sql = "select * from albaran_calibre where numalbar = " & DBSet(Albaran, "N")
'
'    Set Rs = New ADODB.Recordset
'    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'    While Not Rs.EOF
'        'calibres
'        Sql = DevuelveDesdeBDNew(cAgro, vParamAplic.BDDestino & ".calibres", "nomcalib", "codvarie", DBLet(Rs!Codvarie, "N"), "N", , "codcalib", DBLet(Rs!codcalib, "N"), "N")
'        If Sql = "" Then
'            MsgBox "No existe el calibre " & DBLet(Rs!Codvarie, "N") & "-" & DBLet(Rs!codcalib, "N") & ". Revise.", vbExclamation
'            Exit Function
'        End If
'        Rs.MoveNext
'    Wend
'
'    Set Rs = Nothing
'
'
'    'albaran_costes
'    Sql = "select * from albaran_costes where numalbar = " & DBSet(Albaran, "N")
'
'    Set Rs = New ADODB.Recordset
'    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'    While Not Rs.EOF
'        'nombcoste
'        If Not IsNull(Rs!codCoste) Then
'            Sql = DevuelveDesdeBDNew(cAgro, vParamAplic.BDDestino & ".nombcoste", "denominacion", "codcoste", DBLet(Rs!codCoste, "N"), "N")
'            If Sql = "" Then
'                MsgBox "No existe el coste " & DBLet(Rs!codCoste, "N") & ". Revise.", vbExclamation
'                Exit Function
'            End If
'        End If
'        'articulo
'        If Not IsNull(Rs!codArtic) Then
'            Sql = DevuelveDesdeBDNew(cAgro, vParamAplic.BDDestino & ".sartic", "nomartic", "codartic", DBLet(Rs!codArtic, "T"), "T")
'            If Sql = "" Then
'                MsgBox "No existe el artículo " & DBLet(Rs!codArtic, "T") & ". Revise.", vbExclamation
'                Exit Function
'            End If
'        End If
'        Rs.MoveNext
'    Wend
'
'    Set Rs = Nothing
'
'
'    'albaran_costreal
'    Sql = "select * from albaran_costreal where numalbar = " & DBSet(Albaran, "N")
'
'    Set Rs = New ADODB.Recordset
'    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'    While Not Rs.EOF
'        'nombcoste
'        If Not IsNull(Rs!codCoste) Then
'            Sql = DevuelveDesdeBDNew(cAgro, vParamAplic.BDDestino & ".nombcoste", "denominacion", "codcoste", DBLet(Rs!codCoste, "N"), "N")
'            If Sql = "" Then
'                MsgBox "No existe el coste " & DBLet(Rs!codCoste, "N") & ". Revise.", vbExclamation
'                Exit Function
'            End If
'        End If
'        Rs.MoveNext
'    Wend
'
'    Set Rs = Nothing
'
'
'    'albaran_envase
'    Sql = "select * from albaran_envase where numalbar = " & DBSet(Albaran, "N")
'
'    Set Rs = New ADODB.Recordset
'    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'    While Not Rs.EOF
'        'cliente
'        Sql = DevuelveDesdeBDNew(cAgro, vParamAplic.BDDestino & ".clientes", "nomclien", "codclien", DBLet(Rs!CodClien, "N"), "N")
'        If Sql = "" Then
'            MsgBox "No existe el cliente " & DBLet(Rs!CodClien, "N") & ". Revise.", vbExclamation
'            Exit Function
'        End If
'        'articulo
'        Sql = DevuelveDesdeBDNew(cAgro, vParamAplic.BDDestino & ".sartic", "nomartic", "codartic", DBLet(Rs!codArtic, "T"), "T")
'        If Sql = "" Then
'            MsgBox "No existe el artículo " & DBLet(Rs!codArtic, "T") & ". Revise.", vbExclamation
'            Exit Function
'        End If
'        Rs.MoveNext
'    Wend
'
'    Set Rs = Nothing
'
'
'    'albaran_palets
'    Sql = "select * from albaran_palets where numalbar = " & DBSet(Albaran, "N")
'
'    Set Rs = New ADODB.Recordset
'    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'    While Not Rs.EOF
'        'palets
'        Sql = DevuelveDesdeBDNew(cAgro, vParamAplic.BDDestino & ".palets", "numpalet", "numpalet", DBLet(Rs!NumPalet, "N"), "N")
'        If Sql = "" Then
'            MsgBox "No existe el palet " & DBLet(Rs!NumPalet, "N") & ". Revise.", vbExclamation
'            Exit Function
'        End If
'        Rs.MoveNext
'    Wend
'
'    Set Rs = Nothing
'
'
'    ComprobarReferenciales = True
'    Exit Function
'
'eComprobarReferenciales:
'    MuestraError Err.Number, "Comprobar Referenciales", Err.Description
'End Function
'


Private Sub cmdBajar_Click()
'Bajar el item seleccionado del listview2
    BajarItemList Me.ListView1
End Sub

Private Sub cmdCancel_Click(Index As Integer)
    Unload Me
End Sub

Private Sub cmdCancelCalHProd_Click()
    Unload Me
End Sub

Private Sub cmdSubir_Click()
    SubirItemList Me.ListView1
End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        Select Case Opcionlistado
            Case 2 ' informe de palets en camaras
                PonerFoco txtCodigo(32)
                
            Case 3 ' traspaso de albaran a bddestino
                PonerFoco txtCodigo(35)
                
            Case 10 ' Listado de Clientes
                PonerFoco txtCodigo(4)
                
            Case 11 ' Listado de Proveedores
                PonerFoco txtCodigo(2)
            
            Case 12 ' Listado de Variedades
                PonerFoco txtCodigo(6)
        
            Case 13 ' Listado de Calibres
                PonerFoco txtCodigo(8)
                
            Case 14 ' Imforme de Movimientos de calibres
                PonerFoco txtCodigo(12)
            
            Case 15 ' Informe de Horas Trabajadas
                PonerFoco txtCodigo(18)
                
            Case 16 ' calculo de horas productivas
                PonerFoco txtCodigo(27)
            
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
    
    
    For H = 0 To imgBuscar.Count - 1
        Me.imgBuscar(H).Picture = frmPpal.imgListImages16.ListImages(1).Picture
    Next H
     
    Set List = Nothing

    'Ocultar todos los Frames de Formulario
    FrameClientes.visible = False
    FrameVariedades.visible = False
    FrameProveedores.visible = False
    FrameCalibres.visible = False
    Me.FrameMovimientoEnvases.visible = False
    Me.FrameHorasTrabajadas.visible = False
    Me.FrameCalculoHorasProductivas.visible = False
    Me.FrameCreacionPalets.visible = False
    FrameInfPaletsCamaras.visible = False
    FrameTraspasoAlbaran.visible = False
    
    
    '###Descomentar
'    CommitConexion
    
    Select Case Opcionlistado
    Case 1 ' Creacion de palets de forma automática
        FrameCreacionPaletsVisible True, H, W
        indFrame = 0
        tabla = "albaran_envase"
    
    
    Case 2 ' Informe de palets en camaras
        FrameInfPaletsCamarasVisible True, H, W
        Opcion(2).Value = True
        CargarListViewOrden (5)
        indFrame = 0
        tabla = "albaran_envase"
    
    Case 3 ' Traspaso de albaran
        H = 3465
        W = 7380
        PonerFrameVisible FrameTraspasoAlbaran, True, H, W
        tabla = "albaran"
    
    
    'LISTADOS DE MANTENIMIENTOS BASICOS
    '---------------------
    Case 10 '10: Listado de Clientes
        FrameClienteVisible True, H, W
        CargarListViewOrden (0)
        Me.lbltitulo2.Caption = "Informe de Clientes"
        Me.Label2(3).Caption = "Cliente"
        indFrame = 2
        tabla = "clientes"
    
    Case 11 ' Listado de Proveedores
        FrameProveedoresVisible True, H, W
        CargarListViewOrden (2)
        Me.lbltitulo2.Caption = "Informe de Provedores"
        Me.Label2(3).Caption = "Proveedores"
        indFrame = 0
        tabla = "proveedor"
    
    Case 12 ' Listado de Variedades
        FrameVariedadesVisible True, H, W
        CargarListViewOrden (1)
        Me.lbltitulo2.Caption = "Informe de Variedades"
        Me.Label2(3).Caption = "Variedades"
        indFrame = 0
        tabla = "variedades"
    
    Case 13 ' Listado de Calibres
        FrameCalibresVisible True, H, W
        Opcion(0).Value = True
        CargarListViewOrden (3)
        Me.lbltitulo2.Caption = "Informe de Calibres"
        Me.Label2(3).Caption = "Calibres"
        indFrame = 0
        tabla = "calibres"
        
    Case 14 ' Informe de Movimientos de envases
        FrameMovimientosVisible True, H, W
        indFrame = 0
        tabla = "albaran_envase"
        
    Case 15 ' Informe de Horas Trabajadas
        FrameHorasTrabajadasVisible True, H, W
        CargarListViewOrden (4)
        indFrame = 0
        tabla = "horas"
        
    Case 16 ' Proceso de Calculo de Horas Productivas
        FrameCalculoHorasProductivasVisible True, H, W
        CargarListViewOrden (4)
        indFrame = 0
        tabla = "horas"
        
    End Select
    'Esto se consigue poneinedo el cancel en el opcion k corresponda
    Me.CmdCancel(indFrame).Cancel = True
    Me.Width = W + 70
    Me.Height = H + 350
End Sub


Private Sub frmAlm_DatoSeleccionado(CadenaSeleccion As String)
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "00")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmC_Selec(vFecha As Date)
    ' *** repasar si el camp es txtAux o Text1 ***
    txtCodigo(indCodigo).Text = Format(vFecha, "dd/mm/yyyy") '<===
    ' ********************************************
End Sub

Private Sub frmCal_DatoSeleccionado(CadenaSeleccion As String)
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmCam_DatoSeleccionado(CadenaSeleccion As String)
'Form de Consulta de Clientes
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmCli_DatoSeleccionado(CadenaSeleccion As String)
'Form de Consulta de Clientes
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmCol_DatoSeleccionado(CadenaSeleccion As String)
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "00")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmFam_DatoSeleccionado(CadenaSeleccion As String)
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmArt_DatoSeleccionado(CadenaSeleccion As String)
    txtCodigo(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1)
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmDes_DatoSeleccionado(CadenaSeleccion As String)
'Form de Consulta de Destinos
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmMensDestino_DatoSeleccionado(CadenaSeleccion As String)
Dim Sql As String
Dim Sql2 As String
Dim SqlAlb As String
Dim Rs As ADODB.Recordset



    If CadenaSeleccion <> "" Then

        SqlAlb = "select distinct numalbar from albaran where coddesti in (" & CadenaSeleccion & ") and codclien = " & DBSet(txtCodigo(22).Text, "N")
        SqlAlb = SqlAlb & " order by 1 "
        
        Set Rs = New ADODB.Recordset
        Rs.Open SqlAlb, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
        albaranes = ""
        While Not Rs.EOF
            albaranes = albaranes & DBSet(Rs!NumAlbar, "N") & ","
            
            Rs.MoveNext
        Wend
        Set Rs = Nothing
        
        If albaranes <> "" Then
            albaranes = Mid(albaranes, 1, Len(albaranes) - 1)
            Sql = " {albaran_envase.numalbar} in (" & albaranes & ")"
            Sql2 = " {albaran_envase.numalbar} in [" & albaranes & "]"
        Else
            Sql = " {albaran_envase.numalbar} = -1 "
        End If
        If Not AnyadirAFormula(cadselect, Sql) Then Exit Sub
        If Not AnyadirAFormula(cadFormula, Sql2) Then Exit Sub
   
   End If

End Sub


Private Sub frmPro_DatoSeleccionado(CadenaSeleccion As String)
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub


Private Sub frmTArt_DatoSeleccionado(CadenaSeleccion As String)
    txtCodigo(indCodigo).Text = RecuperaValor(CadenaSeleccion, 1)
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmVar_DatoSeleccionado(CadenaSeleccion As String)
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub frmTra_DatoSeleccionado(CadenaSeleccion As String)
    txtCodigo(indCodigo).Text = Format(RecuperaValor(CadenaSeleccion, 1), "000000")
    txtNombre(indCodigo).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub


Private Sub imgBuscar_Click(Index As Integer)
   Select Case Index
        Case 0, 1 'Para listados básicos
            ' productos
             AbrirFrmProductos (Index)
            
        Case 2, 3 'PROVEEDORES
            AbrirFrmProveedores (Index)
            
        Case 4, 5 'CLIENTES
            AbrirFrmClientes (Index)
            
        Case 6, 7 'Clientes / Proveedores
'            AbrirFrmFamilias (Index)

        Case 8, 9 'VARIEDADES
            AbrirFrmVariedades (Index)
        
        Case 10, 11 'CALIBRES
            AbrirFrmCalibres (Index)
        
        Case 12, 13 'TIPOS DE ENVASES
            AbrirFrmTipEnvases (Index)
    
        Case 14, 15 'Horas trabajadas
            AbrirFrmManTraba (Index)
    
        Case 16, 17 'Articulos
            AbrirFrmManArtic (Index)
        
        Case 18, 19 'Clientes
            AbrirFrmManClien (Index)
    
        Case 20
            AbrirFrmManAlmac (Index)
            
        Case 21 ' Destinos
            If txtCodigo(23).Text <> "" Then AbrirFrmDestinos (26)
        Case 22
            If txtCodigo(23).Text <> "" Then AbrirFrmDestinos (28)
            
        Case 23
            AbrirFrmManCamara (Index)
        Case 24
            AbrirFrmManCamara (Index)
            
        ' traspaso de albaranes
        Case 26
            AbrirFrmClientes2 (Index + 9)
        
        Case 25
            AbrirFrmDestinos2 Index + 9
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

    If Index = 4 Then
        indCodigo = 30
    ElseIf Index = 5 Then
        indCodigo = 29
    ElseIf Index = 6 Then
        indCodigo = 31
    Else
        indCodigo = Index + 14
    End If
    
    imgFecha(0).Tag = indCodigo '<===
    ' *** repasar si el camp es txtAux o Text1 ***
    If txtCodigo(indCodigo).Text <> "" Then frmC.NovaData = txtCodigo(indCodigo).Text
    ' ********************************************

    frmC.Show vbModal
    Set frmC = Nothing
    ' *** repasar si el camp es txtAux o Text1 ***
    PonerFoco txtCodigo(indCodigo) 'txtCodigo(CByte(imgFecha(0).Tag) + 14) '<===
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
    If KeyAscii = teclaBuscar Then
        Select Case Index
            Case 2: KEYBusqueda KeyAscii, 2 'proveedor desde
            Case 3: KEYBusqueda KeyAscii, 3 'proveedor hasta
            Case 4: KEYBusqueda KeyAscii, 4 'cliente desde
            Case 5: KEYBusqueda KeyAscii, 5 'cliente hasta
            Case 6: KEYBusqueda KeyAscii, 6 'producto desde
            Case 7: KEYBusqueda KeyAscii, 7 'producto hasta
            Case 0: KEYBusqueda KeyAscii, 0 'variedad desde
            Case 1: KEYBusqueda KeyAscii, 1 'variedad hasta
            Case 8: KEYBusqueda KeyAscii, 8 'cliente de la factura rectificativa
            Case 18: KEYBusqueda KeyAscii, 14 'trabajador desde
            Case 19: KEYBusqueda KeyAscii, 15 'trabajador hasta
            Case 20: KEYBusqueda KeyAscii, 16 'articulo desde
            Case 21: KEYBusqueda KeyAscii, 17 'articulo hasta
            Case 24: KEYBusqueda KeyAscii, 20 'almacen para el calculo de horas productivas
            Case 22: KEYBusqueda KeyAscii, 18 'cliente desde
            Case 23: KEYBusqueda KeyAscii, 19 'cliente hasta
            
            Case 32: KEYBusqueda KeyAscii, 23 'camara desde
            Case 33: KEYBusqueda KeyAscii, 24 'camara hasta
            Case 29: KEYFecha KeyAscii, 5 'fecha desde
            Case 31: KEYFecha KeyAscii, 6 'fecha hasta
            
            Case 30: KEYFecha KeyAscii, 16 'fecha de carga automatica de palets
        
            ' traspaso de albaran a bd destino
            Case 35: KEYBusqueda KeyAscii, 26 'cliente
            Case 34: KEYBusqueda KeyAscii, 25 'destino
        
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
    imgFecha_Click (indice)
End Sub

Private Sub txtCodigo_LostFocus(Index As Integer)
Dim Cad As String, cadTipo As String 'tipo cliente

    'Quitar espacios en blanco por los lados
    txtCodigo(Index).Text = Trim(txtCodigo(Index).Text)
'    If txtCodigo(Index).Text = "" Then Exit Sub
    
    'Si se ha abierto otro formulario, es que se ha pinchado en prismaticos y no
    'mostrar mensajes ni hacer nada
    If Screen.ActiveForm.Name <> Me.Name Then Exit Sub

    Select Case Index
        Case 0, 1 'VARIEDADES
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "variedades", "nomvarie", "codvarie", "N")
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "000000")
        
        Case 2, 3 'PROVEEDORES
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "proveedor", "nomprove", "codprove", "N")
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "000000")
        
        Case 4, 5, 22, 23 'CLIENTE
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "clientes", "nomclien", "codclien", "N")
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "000000")
            
            If Index = 8 Then ' en la factura rectificativa el nuevo cliente ha de existir
                If txtCodigo(8).Text <> "" And txtNombre(8).Text = "" Then
                    MsgBox "El cliente introducido no existe. Si introduce número de cliente éste debe existir.", vbExclamation
                    PonerFoco txtCodigo(8)
                End If
            End If
            
            If Index = 22 Or Index = 23 Then
                ' solo se puede introducir destino si cliente desde y hasta son iguales
                txtCodigo(26).Enabled = (txtCodigo(22).Text = txtCodigo(23).Text)
                imgBuscar(21).Enabled = txtCodigo(26).Enabled
                imgBuscar(22).Enabled = txtCodigo(26).Enabled
                If Not txtCodigo(26).Enabled Then
                    txtCodigo(26).Text = ""
                    txtNombre(26).Text = ""
                End If
                txtCodigo(28).Enabled = (txtCodigo(22).Text = txtCodigo(23).Text)
                If Not txtCodigo(28).Enabled Then
                    txtCodigo(28).Text = ""
                    txtNombre(28).Text = ""
                End If
                
                If Index = 23 Then
                    If txtCodigo(26).Enabled Then
                        PonerFoco txtCodigo(26)
                    Else
                        PonerFoco txtCodigo(14)
                    End If
                End If
            End If
            
        
        Case 6, 7 'PRODUCTOS
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "productos", "nomprodu", "codprodu", "N")
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "000")
        
        Case 8, 9 'VARIEDADES
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "variedades", "nomvarie", "codvarie", "N")
            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "000000")
            
        Case 10, 11 'CALIBRES
            If txtCodigo(8).Text = txtCodigo(9).Text And txtCodigo(8).Text <> "" Then
                txtNombre(Index) = DevuelveDesdeBDNew(cAgro, "calibres", "codcalib", "codvarie", txtCodigo(8).Text, "N", , "codcalib", txtCodigo(Index).Text, "N")
            End If
            
        Case 12, 13 'TIPOS DE ENVASES
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "stipar", "nomtipar", "codtipar", "N")
            
        Case 14, 15, 16, 17, 27, 29, 31 'FECHAS
            If txtCodigo(Index).Text <> "" Then PonerFormatoFecha txtCodigo(Index)
            
        Case 18, 19 'TRABAJADORES
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "straba", "nomtraba", "codtraba", "N")
            
        Case 20, 21 'ARTICULOS
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "sartic", "nomartic", "codartic", "T")
            
        Case 25 ' porcentaje
            If txtCodigo(Index).Text <> "" Then
                 PonerFormatoDecimal txtCodigo(Index), 9
            End If

        Case 24 'ALMACEN
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "salmpr", "nomalmac", "codalmac", "N")

        Case 26, 28  'DESTINO
            If txtCodigo(22).Text <> "" And txtCodigo(22).Text = txtCodigo(23).Text Then
                txtNombre(Index).Text = DevuelveDesdeBDNew(cAgro, "destinos", "nomdesti", "codclien", txtCodigo(22).Text, "N", , "coddesti", txtCodigo(Index).Text, "N")
                If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "000")
            End If
        
        
        Case 32, 33 'CAMARAS
            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), "camaras", "nomcamara", "codcamara", "N")
            
        '
        Case 35 ' CLIENTE
'            txtNombre(Index).Text = PonerNombreDeCod(txtCodigo(Index), vParamAplic.BDDestino & ".clientes", "nomclien", "codclien", "N")
'            If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "000000")
        
        Case 34 'DESTINO
'            If txtCodigo(35).Text <> "" Then
'                txtNombre(Index).Text = DevuelveDesdeBDNew(cAgro, vParamAplic.BDDestino & ".destinos", "nomdesti", "codclien", txtCodigo(35).Text, "N", , "coddesti", txtCodigo(Index).Text, "N")
'                If txtCodigo(Index).Text <> "" Then txtCodigo(Index).Text = Format(txtCodigo(Index).Text, "000")
'            End If
        
    
    
    
    End Select
End Sub


Private Sub FrameClienteVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Frame para el listado de clientes
    Me.FrameClientes.visible = visible
    If visible = True Then
        Me.FrameClientes.Top = -90
        Me.FrameClientes.Left = 0
        Me.FrameClientes.Height = 3420
        Me.FrameClientes.Width = 8600
        W = Me.FrameClientes.Width
        H = Me.FrameClientes.Height
    End If
End Sub

Private Sub FrameProveedoresVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
    Me.FrameProveedores.visible = visible
    If visible = True Then
        Me.FrameProveedores.Top = -90
        Me.FrameProveedores.Left = 0
        Me.FrameProveedores.Height = 3420
        Me.FrameProveedores.Width = 8600
        W = Me.FrameProveedores.Width
        H = Me.FrameProveedores.Height
    End If
End Sub

Private Sub FrameVariedadesVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
    Me.FrameVariedades.visible = visible
    If visible = True Then
        Me.FrameVariedades.Top = -90
        Me.FrameVariedades.Left = 0
        Me.FrameVariedades.Height = 4820
        Me.FrameVariedades.Width = 8600
        W = Me.FrameVariedades.Width
        H = Me.FrameVariedades.Height
    End If
End Sub

Private Sub FrameCalibresVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
    Me.FrameCalibres.visible = visible
    If visible = True Then
        Me.FrameCalibres.Top = -90
        Me.FrameCalibres.Left = 0
        Me.FrameCalibres.Height = 4820
        Me.FrameCalibres.Width = 6600
        W = Me.FrameCalibres.Width
        H = Me.FrameCalibres.Height
    End If
End Sub

Private Sub FrameMovimientosVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
    Me.FrameMovimientoEnvases.visible = visible
    If visible = True Then
        Me.FrameMovimientoEnvases.Top = -90
        Me.FrameMovimientoEnvases.Left = 0
        Me.FrameMovimientoEnvases.Height = 7545
        Me.FrameMovimientoEnvases.Width = 7155
        W = Me.FrameMovimientoEnvases.Width
        H = Me.FrameMovimientoEnvases.Height
    End If
End Sub

Private Sub FrameCreacionPaletsVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
    Me.FrameCreacionPalets.visible = visible
    If visible = True Then
        Me.FrameCreacionPalets.Top = -90
        Me.FrameCreacionPalets.Left = 0
        Me.FrameCreacionPalets.Height = 3525
        Me.FrameCreacionPalets.Width = 5835
        W = Me.FrameCreacionPalets.Width
        H = Me.FrameCreacionPalets.Height
    End If
End Sub


Private Sub FrameInfPaletsCamarasVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
    Me.FrameInfPaletsCamaras.visible = visible
    If visible = True Then
        Me.FrameInfPaletsCamaras.Top = -90
        Me.FrameInfPaletsCamaras.Left = 0
        Me.FrameInfPaletsCamaras.Height = 4455
        Me.FrameInfPaletsCamaras.Width = 7020
        W = Me.FrameInfPaletsCamaras.Width
        H = Me.FrameInfPaletsCamaras.Height
    End If
End Sub




Private Sub FrameHorasTrabajadasVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
    Me.FrameHorasTrabajadas.visible = visible
    If visible = True Then
        Me.FrameHorasTrabajadas.Top = -90
        Me.FrameHorasTrabajadas.Left = 0
        Me.FrameHorasTrabajadas.Height = 4455
        Me.FrameHorasTrabajadas.Width = 7425
        W = Me.FrameHorasTrabajadas.Width
        H = Me.FrameHorasTrabajadas.Height
    End If
End Sub

Private Sub FrameCalculoHorasProductivasVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
    Me.FrameCalculoHorasProductivas.visible = visible
    If visible = True Then
        Me.FrameCalculoHorasProductivas.Top = -90
        Me.FrameCalculoHorasProductivas.Left = 0
        Me.FrameCalculoHorasProductivas.Height = 3525
        Me.FrameCalculoHorasProductivas.Width = 5835
        W = Me.FrameCalculoHorasProductivas.Width
        H = Me.FrameCalculoHorasProductivas.Height
    End If
End Sub

Private Sub CargarListViewOrden(Index As Integer)
Dim ItmX As ListItem

    'Los encabezados
    ListView1(Index).ColumnHeaders.Clear
    ListView1(Index).ColumnHeaders.Add , , "Campo", 1390

    Select Case Index
        Case 0
            Set ItmX = ListView1(Index).ListItems.Add
            ItmX.Text = "Codigo"
            Set ItmX = ListView1(Index).ListItems.Add
            ItmX.Text = "Alfabético"
        Case 1
            Set ItmX = ListView1(Index).ListItems.Add
            ItmX.Text = "Clase"
            Set ItmX = ListView1(Index).ListItems.Add
            ItmX.Text = "Producto"
        Case 2
            Set ItmX = ListView1(Index).ListItems.Add
            ItmX.Text = "Codigo"
            Set ItmX = ListView1(Index).ListItems.Add
            ItmX.Text = "Alfabético"
        Case 3
            Set ItmX = ListView1(Index).ListItems.Add
            ItmX.Text = "Calibre"
            Set ItmX = ListView1(Index).ListItems.Add
            ItmX.Text = "Variedad"
        Case 4
            Set ItmX = ListView1(Index).ListItems.Add
            ItmX.Text = "Trabajador"
            Set ItmX = ListView1(Index).ListItems.Add
            ItmX.Text = "Fecha"
            
        '[Monica]04/12/2017: nuevo informe de palets en camaras
        Case 5
            Set ItmX = ListView1(Index).ListItems.Add
            ItmX.Text = "Cámara"
            Set ItmX = ListView1(Index).ListItems.Add
            ItmX.Text = "Fecha"
    
    End Select
        
End Sub

Private Sub InicializarVbles()
    cadFormula = ""
    cadselect = ""
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
        If Not AnyadirAFormula(cadselect, devuelve) Then Exit Function
    Else
        devuelve2 = CadenaDesdeHastaBD(codD, codH, Codigo, TipCod)
        If devuelve2 = "Error" Then Exit Function
        If Not AnyadirAFormula(cadselect, devuelve2) Then Exit Function
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
        .Titulo = cadTitulo
        .NombreRPT = cadNombreRPT
        .EnvioEMail = False
        .ConSubInforme = ConSubInforme
        .Opcion = Opcionlistado
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
        'Informe de variedades
        Case "Clase"
            cadParam = cadParam & campo & "{" & tabla & ".codclase}" & "|"
            cadParam = cadParam & nomCampo & " {" & "clases" & ".nomclase}" & "|"
            cadParam = cadParam & "pTitulo1" & "=""Producto""" & "|"
            numParam = numParam + 3
            
        Case "Producto"
            cadParam = cadParam & campo & "{" & tabla & ".codprodu}" & "|"
            cadParam = cadParam & nomCampo & " {" & "productos" & ".nomprodu}" & "|"
            cadParam = cadParam & "pTitulo1" & "=""Clase""" & "|"
            numParam = numParam + 3

        'Informe de calibres
        Case "Variedad"
            cadParam = cadParam & campo & "{" & tabla & ".codvarie}" & "|"
            cadParam = cadParam & nomCampo & " {" & "variedades" & ".nomvarie}" & "|"
            cadParam = cadParam & "pTitulo1" & "=""Variedad""" & "|"
            numParam = numParam + 3
            
        Case "Calibre"
            cadParam = cadParam & campo & "{" & tabla & ".codcalib}" & "|"
            cadParam = cadParam & nomCampo & " {" & "calibres" & ".nomcalib}" & "|"
            cadParam = cadParam & "pTitulo1" & "=""Calibre""" & "|"
            numParam = numParam + 3
            
        'Informe de palets en camaras
        Case "Cámara"
            cadParam = cadParam & campo & "{tmpinformes.importe2}" & "|"
            cadParam = cadParam & "pTitulo1" & "=""Camara""" & "|"
            numParam = numParam + 2
        
        Case "Fecha"
            cadParam = cadParam & campo & "{tmpinformes.fecha1}" & "|"
            cadParam = cadParam & "pTitulo1" & "=""Fecha""" & "|"
            numParam = numParam + 2
            
    End Select

End Function

Private Function PonerOrden(cadgrupo As String) As Byte
Dim campo As String
Dim nomCampo As String

    PonerOrden = 0

    Select Case cadgrupo
        Case "Codigo"
            cadParam = cadParam & "Orden" & "= {" & tabla
            Select Case Opcionlistado
                Case 10
                    cadParam = cadParam & ".codclien}|"
                Case 11
                    cadParam = cadParam & ".codprove}|"
            End Select
            Tipo = "Código"
        Case "Alfabético"
            cadParam = cadParam & "Orden" & "= {" & tabla
            Select Case Opcionlistado
                Case 10
                    cadParam = cadParam & ".nomclien}|"
                Case 11
                    cadParam = cadParam & ".nomprove}|"
            End Select
            Tipo = "Alfabético"
    End Select
    
    numParam = numParam + 1

End Function

Private Sub AbrirFrmDestinos(indice As Integer, Optional Cliente As String)
'    indCodigo = indice
'    Set frmDes = New frmDestCli
'    frmDes.DatosADevolverBusqueda = "0|1|"
''    frmDes.DeConsulta = True
'    If Cliente <> "" Then
'        frmDes.Cliente = Cliente
'    Else
'        frmDes.Cliente = txtCodigo(22).Text
'    End If
'    frmDes.CodigoActual = txtCodigo(indCodigo)
'    frmDes.Show vbModal
'    Set frmDes = Nothing
End Sub



Private Sub AbrirFrmProveedores(indice As Integer)
'    indCodigo = indice
'    Set frmPro = New frmManProve
'    frmPro.DatosADevolverBusqueda = "0|1|"
'    frmPro.DeConsulta = True
'    frmPro.CodigoActual = txtCodigo(indCodigo)
'    frmPro.Show vbModal
'    Set frmPro = Nothing
End Sub

Private Sub AbrirFrmProductos(indice As Integer)
'    indCodigo = indice
'    Set frmProd = New frmManProductos
'    frmProd.DatosADevolverBusqueda = "0|1|"
'    frmProd.DeConsulta = True
'    frmProd.CodigoActual = txtCodigo(indCodigo)
'    frmProd.Show vbModal
'    Set frmProd = Nothing
End Sub


Private Sub AbrirFrmClientes(indice As Integer)
'    indCodigo = indice
'    Set frmCli = New frmClientes
'    frmCli.DatosADevolverBusqueda = "0|2|"
'    frmCli.Show vbModal
'    Set frmCli = Nothing
End Sub

Private Sub AbrirFrmClientes2(indice As Integer)
    indCodigo = indice
    
    Set frmCli2 = New frmBasico
        
    AyudaClienteCom frmCli2, txtCodigo(35).Text
    
    Set frmCli2 = Nothing
End Sub

Private Sub AbrirFrmDestinos2(indice As Integer)
'    indCodigo = indice
'
'    Set frmDes2 = New frmBasico
'
'    AyudaDestinos frmDes2, txtCodigo(35).Text
'
'    Set frmDes2 = Nothing
End Sub




Private Sub AbrirFrmVariedades(indice As Integer)
    indCodigo = indice
    Set frmVar = New frmManVariedad
    frmVar.DatosADevolverBusqueda = "0|1|"
    frmVar.Show vbModal
    Set frmVar = Nothing
End Sub

Private Sub AbrirFrmCalibres(indice As Integer)
'    indCodigo = indice
'    Set frmCal = New frmManCalibres
'    frmCal.DatosADevolverBusqueda = "2|3|"
'    frmCal.Show vbModal
'    Set frmCal = Nothing
End Sub

Private Sub AbrirFrmTipEnvases(indice As Integer)
'    indCodigo = indice
'    Set frmTArt = New frmManTipArtic
'    frmTArt.DatosADevolverBusqueda = "0|1|"
'    frmTArt.Show vbModal
'    Set frmTArt = Nothing
End Sub

Private Sub AbrirFrmManTraba(indice As Integer)
'    indCodigo = indice + 4
'    Set frmTra = New frmManTraba
'    frmTra.DatosADevolverBusqueda = "0|2|"
'    frmTra.Show vbModal
'    Set frmTra = Nothing
End Sub

Private Sub AbrirFrmManArtic(indice As Integer)
'    indCodigo = indice + 4
'    Set frmArt = New frmManArtic
'    frmArt.DatosADevolverBusqueda = "0|1|"
'    frmArt.Show vbModal
'    Set frmArt = Nothing
End Sub

Private Sub AbrirFrmManClien(indice As Integer)
'    indCodigo = indice + 4
'    Set frmCli = New frmClientes
'    frmCli.DatosADevolverBusqueda = "0|2|"
'    frmCli.Show vbModal
'    Set frmCli = Nothing
End Sub

Private Sub AbrirFrmManAlmac(indice As Integer)
'    indCodigo = indice + 4
'    Set frmAlm = New frmManAlmProp
'    frmAlm.DatosADevolverBusqueda = "0|1|"
'    frmAlm.Show vbModal
'    Set frmAlm = Nothing
End Sub


Private Sub AbrirFrmManCamara(indice As Integer)
'    indCodigo = indice + 9
'    Set frmCam = New frmManCamara
'    frmCam.DatosADevolverBusqueda = "0|1|"
'    frmCam.Show vbModal
'    Set frmCam = Nothing
End Sub




Private Function CargarTablaTemporal() As Boolean
'Dim Sql As String
'Dim SQL1 As String
'Dim Rs As ADODB.Recordset
'
'    On Error GoTo eCargarTablaTemporal
'
'    CargarTablaTemporal = False
'
'    Sql = "delete from tmpenvasesret where codusu = " & DBSet(vUsu.Codigo, "N")
'    conn.Execute Sql
'
''select albaran_envase.codartic, albaran_envase.fechamov
''from (albaran_envase inner join sartic on albaran_envase.codartic = sartic.codartic) inner join stipar on sartic.codtipar = stipar.codtipar
''Where stipar.esretornable = 1
''Union
''select smoval.codartic, smoval.fechamov
''from (smoval inner join  sartic on smoval.codartic = sartic.codartic) inner join stipar on sartic.codtipar = stipar.codtipar
''Where stipar.esretornable = 1
'
'    '[Monica]11/06/2014: agrupamos la cantidad
'    Sql = "select " & vUsu.Codigo & ", albaran_envase.codartic, albaran_envase.fechamov, sum(albaran_envase.cantidad) cantidad, albaran_envase.tipomovi, albaran_envase.numalbar, "
'    Sql = Sql & " albaran_envase.codclien, clientes.nomclien, " & DBSet(vParamAplic.CodTipomAlb, "T") 'DBSet("ALV", "T")
'    Sql = Sql & " from ((albaran_envase inner join sartic on albaran_envase.codartic = sartic.codartic) inner join stipar on sartic.codtipar = stipar.codtipar) "
'    Sql = Sql & " inner join clientes on albaran_envase.codclien = clientes.codclien "
'    Sql = Sql & " where stipar.esretornable = 1 "
'
'    If txtCodigo(12).Text <> "" Then Sql = Sql & " and stipar.codtipar >= " & DBSet(txtCodigo(12).Text, "N")
'    If txtCodigo(13).Text <> "" Then Sql = Sql & " and stipar.codtipar <= " & DBSet(txtCodigo(13).Text, "N")
'
'    If txtCodigo(20).Text <> "" Then Sql = Sql & " and albaran_envase.codartic >= " & DBSet(txtCodigo(20).Text, "T")
'    If txtCodigo(21).Text <> "" Then Sql = Sql & " and albaran_envase.codartic <= " & DBSet(txtCodigo(21).Text, "T")
'
'    If txtCodigo(22).Text <> "" Then Sql = Sql & " and albaran_envase.codclien >= " & DBSet(txtCodigo(22).Text, "N")
'    If txtCodigo(23).Text <> "" Then Sql = Sql & " and albaran_envase.codclien <= " & DBSet(txtCodigo(23).Text, "N")
'
'    If txtCodigo(14).Text <> "" Then Sql = Sql & " and albaran_envase.fechamov >= " & DBSet(txtCodigo(14).Text, "F")
'    If txtCodigo(15).Text <> "" Then Sql = Sql & " and albaran_envase.fechamov <= " & DBSet(txtCodigo(15).Text, "F")
'
'    If albaranes <> "" Then Sql = Sql & " and albaran_envase.numalbar in (" & albaranes & ")"
'
'    '[Monica]11/06/2014: agrupamos pq sumamos las cantidades del mismo tipo, artículo y demas
'    Sql = Sql & " group by 1,2,3,5,6,7,8,9 "
'
'
'    Sql = Sql & " union "
'
'    Sql = Sql & "select " & vUsu.Codigo & ", smoval.codartic, smoval.fechamov, sum(smoval.cantidad) cantidad, smoval.tipomovi, smoval.document, "
'    Sql = Sql & " smoval.codigope, proveedor.nomprove, " & DBSet("ALC", "T")
'    Sql = Sql & " from ((smoval inner join sartic on smoval.codartic = sartic.codartic "
'    '[Monica]22/11/2010:faltaba añadir que sean las compras
'    Sql = Sql & " and smoval.detamovi = 'ALC'"
'    Sql = Sql & ") inner join stipar on sartic.codtipar = stipar.codtipar) "
'    Sql = Sql & " inner join proveedor on smoval.codigope = proveedor.codprove "
'    Sql = Sql & " where stipar.esretornable = 1 "
'
'    If txtCodigo(12).Text <> "" Then Sql = Sql & " and stipar.codtipar >= " & DBSet(txtCodigo(12).Text, "N")
'    If txtCodigo(13).Text <> "" Then Sql = Sql & " and stipar.codtipar <= " & DBSet(txtCodigo(13).Text, "N")
'
'    If txtCodigo(20).Text <> "" Then Sql = Sql & " and smoval.codartic >= " & DBSet(txtCodigo(20).Text, "T")
'    If txtCodigo(21).Text <> "" Then Sql = Sql & " and smoval.codartic <= " & DBSet(txtCodigo(21).Text, "T")
'
'    If txtCodigo(14).Text <> "" Then Sql = Sql & " and smoval.fechamov >= " & DBSet(txtCodigo(14).Text, "F")
'    If txtCodigo(15).Text <> "" Then Sql = Sql & " and smoval.fechamov <= " & DBSet(txtCodigo(15).Text, "F")
'
'    '[Monica]11/06/2014: agrupamos pq sumamos las cantidades del mismo tipo, artículo y demas
'    Sql = Sql & " group by 1,2,3,5,6,7,8,9 "
'
'    SQL1 = "insert into tmpenvasesret " & Sql
'    conn.Execute SQL1
'
'    CargarTablaTemporal = True
'    Exit Function
'
'eCargarTablaTemporal:
'    MuestraError Err.Number, "Carga Tabla Temporal"
End Function


' cargamos la tabla temporal para saber que albaranes tienen saldo 0 o distinto de cero segun el check5
' solo se carga en caso de que tengamos que sacar el informe con saldos iguales o distintos de cero

Private Function CargarTablaTemporal2() As Boolean
Dim Sql As String
Dim SQL1 As String
Dim Sql2 As String
Dim Rs As ADODB.Recordset
Dim Entradas As Long
Dim Salidas As Long
Dim Saldo As Long

    On Error GoTo eCargarTablaTemporal2
    
    CargarTablaTemporal2 = False

    Screen.MousePointer = vbHourglass
    
    Sql = "delete from tmpinformes where codusu = " & DBSet(vUsu.Codigo, "N")
    conn.Execute Sql

    Sql = "select  albaran_envase.codartic, albaran_envase.fechamov, albaran_envase.numalbar, albaran_envase.codclien  "
    Sql = Sql & " from ((albaran_envase inner join sartic on albaran_envase.codartic = sartic.codartic) inner join stipar on sartic.codtipar = stipar.codtipar) "
    Sql = Sql & " inner join clientes on albaran_envase.codclien = clientes.codclien "
    Sql = Sql & " where stipar.esretornable = 1 "
    
    If txtCodigo(12).Text <> "" Then Sql = Sql & " and stipar.codtipar >= " & DBSet(txtCodigo(12).Text, "N")
    If txtCodigo(13).Text <> "" Then Sql = Sql & " and stipar.codtipar <= " & DBSet(txtCodigo(13).Text, "N")
    
    If txtCodigo(20).Text <> "" Then Sql = Sql & " and albaran_envase.codartic >= " & DBSet(txtCodigo(20).Text, "T")
    If txtCodigo(21).Text <> "" Then Sql = Sql & " and albaran_envase.codartic <= " & DBSet(txtCodigo(21).Text, "T")
    
    If txtCodigo(22).Text <> "" Then Sql = Sql & " and albaran_envase.codclien >= " & DBSet(txtCodigo(22).Text, "N")
    If txtCodigo(23).Text <> "" Then Sql = Sql & " and albaran_envase.codclien <= " & DBSet(txtCodigo(23).Text, "N")
    
    If txtCodigo(14).Text <> "" Then Sql = Sql & " and albaran_envase.fechamov >= " & DBSet(txtCodigo(14).Text, "F")
    If txtCodigo(15).Text <> "" Then Sql = Sql & " and albaran_envase.fechamov <= " & DBSet(txtCodigo(15).Text, "F")
    
    If albaranes <> "" Then Sql = Sql & " and albaran_envase.numalbar in (" & albaranes & ")"
    
    
    Sql = Sql & " group by 1,2,3,4 "
    Sql = Sql & " order by 1,2,3,4 "
    
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    Sql2 = ""
    
    While Not Rs.EOF
        ' para cada registro que hay de entrada y que hay de salida : calculo el saldo
        Sql = "select sum(albaran_envase.cantidad) from albaran_envase where codartic = " & DBSet(Rs.Fields(0).Value, "T")
        Sql = Sql & " and fechamov = " & DBSet(Rs.Fields(1).Value, "F")
        Sql = Sql & " and numalbar = " & DBSet(Rs.Fields(2).Value, "N")
        Sql = Sql & " and codclien = " & DBSet(Rs.Fields(3).Value, "N")
        
        Entradas = DevuelveValor(Sql & " and tipomovi = 1 ")
        Salidas = DevuelveValor(Sql & " and tipomovi = 0 ")
        
        Saldo = Entradas - Salidas
        
        If Check5 = 0 Or (Check5 = 1 And Saldo <> 0) Then
            Sql2 = Sql2 & "(" & vUsu.Codigo & ", " & DBSet(Rs.Fields(0).Value, "T") & "," & DBSet(Rs.Fields(1).Value, "F") & ","
            Sql2 = Sql2 & DBSet(Rs.Fields(2).Value, "N") & ","
            Sql2 = Sql2 & DBSet(Rs.Fields(3).Value, "N") & ","
            Sql2 = Sql2 & DBSet(Entradas, "N") & "," & DBSet(Salidas, "N") & ","
            Sql2 = Sql2 & DBSet(Saldo, "N") & "),"
        End If
    
        Rs.MoveNext
    Wend
    
    Set Rs = Nothing
    
    'quitamos la ultima coma
    If Sql2 <> "" Then
        Sql2 = Mid(Sql2, 1, Len(Sql2) - 1)
                                               'articulo,fecha,  numalbar, codclien,entradas,  salidas, saldo
        SQL1 = "insert into tmpinformes (codusu,nombre1, fecha1, importe1, codigo1, importe2,  importe3, importe4) values " & Sql2
        
        conn.Execute SQL1
    End If
    
    CargarTablaTemporal2 = True
    Screen.MousePointer = vbDefault
    
    Exit Function
    
eCargarTablaTemporal2:
    Screen.MousePointer = vbDefault
    MuestraError Err.Number, "Carga Tabla Temporal"
End Function

Private Function CalculoHorasProductivas() As Boolean
Dim Rs As ADODB.Recordset
Dim Sql As String
Dim SQL1 As String

    On Error GoTo eCalculoHorasProductivas

    CalculoHorasProductivas = False

    Sql = "fechahora = " & DBSet(txtCodigo(27).Text, "F") & " and codalmac = " & DBSet(txtCodigo(24), "N")
    Sql = Sql & " and codtraba in (select codtraba from straba where codsecci = 1)"


    If BloqueaRegistro("horas", Sql) Then
        SQL1 = "update horas set horasproduc = round(horasdia * (1 + (" & DBSet(txtCodigo(25), "N") & "/ 100)),2) "
        SQL1 = SQL1 & " where fechahora = " & DBSet(txtCodigo(27).Text, "F")
        SQL1 = SQL1 & " and codalmac = " & DBSet(txtCodigo(24), "N")
        SQL1 = SQL1 & " and codtraba in (select codtraba from straba where codsecci = 1) "
        
        conn.Execute SQL1
    
        CalculoHorasProductivas = True
    End If

    TerminaBloquear
    Exit Function

eCalculoHorasProductivas:
    MuestraError Err.Number, "Calculo Horas Productivas", Err.Description
    TerminaBloquear
End Function

